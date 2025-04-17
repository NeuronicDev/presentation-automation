import xml.etree.ElementTree as ET
import logging, re, json
from typing import Tuple, Dict, Any


from config.llm_provider import gemini_flash_llm  
from config.config import LLM_API_KEY

from google import genai
client = genai.Client(api_key=LLM_API_KEY)


# Define namespaces used in PowerPoint XML
PPTX_NAMESPACES = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

def validate_xml_with_llm(xml_code: str, task: Dict[str, Any], slide_context: Dict[str, Any]) -> Tuple[bool, str]:
    
    XML_SEMANTIC_VALIDATION_PROMPT = """
    
    # ROLE: AI PowerPoint XML Semantic & Task Compliance Validator

    You are an expert AI assistant specializing in analyzing generated PowerPoint PresentationML XML code. 
    You are given the original task description and the generated XML snippet intended to fulfill that task. 
    Assume the provided XML snippet is already confirmed to be **syntactically well-formed**.

    Your goal is to determine if the **semantic changes** within the 'Generated Modified XML' accurately and completely implement the requirements described in the 'Intended Task'.

    **## Input Provided:**
    1.  **Intended Task:** (The specific task specification that generated the modification)
        ```json
        {task_spec_json}
        ```
        *(Focus on `action`, `task_description`, `params`, `target_element_hint`)*
    2.  **Generated Modified XML:** (The XML produced by the generation step - assume well-formed)
        ```xml
        {modified_xml}
        ```
        *(You don't have the original XML here, focus on whether this generated XML fulfills the task)*
    3.  **(Optional) Slide Image Context:** [Image data may be provided separately for visual reference]


    **## Your Validation Task:**
    1.  **Analyze Task Requirements:** Carefully read the `task_description`, `action`, `params`, and `target_element_hint` from the 'Intended Task'. What specific changes were requested? Which elements should be affected?
    2.  **Analyze Generated XML:** Examine the provided 'Generated Modified XML'.
        *   Does it contain the necessary XML elements and attributes to achieve the requested `action`? (e.g., for `change_font`, does it have `<a:latin typeface="...">` and `<a:sz val="..."/>` with correct values inside the relevant `<a:rPr>`?)
        *   Are the values used in the XML attributes consistent with the `params` from the task? (e.g., is `<a:sz val="1100"/>` present if `params` specified `font_size: 11`?)
        *   Does the XML structure suggest the change is applied to the elements implied by the `target_element_hint`? (This requires inferring structure, e.g., changes inside a specific `<p:sp>` likely corresponding to the hint).
        *   Does the XML seem complete for the requested task, or are parts missing?
        *   Are there any obvious *semantic* errors, even if syntactically valid (e.g., using `<a:srgbClr>` where `<a:schemeClr>` might be expected based on context, incorrect `algn` values)?


    **## Output Requirements:**
    Respond ONLY with a single, valid JSON object containing the following keys:

    *   `task_description`: (String) The `task_description` from the input task specification.
    *   `semantic_validation_status`: (String) Your assessment - one of: `"Valid"`, `"Valid_With_Warnings"`, `"Invalid"`.
        *   `"Valid"`: The XML appears to correctly and completely implement the task.
        *   `"Valid_With_Warnings"`: The XML likely implements the core task but has minor issues or potential inconsistencies (e.g., used RGB instead of theme color, slight deviation from params).
        *   `"Invalid"`: The XML clearly fails to implement the task correctly (e.g., missing required tags/attributes, wrong values, applied to wrong structure).
    *   `issues_or_warnings`: (List of Strings) Specific reasons for `"Invalid"` or `"Valid_With_Warnings"` status. Describe the semantic discrepancy found (e.g., "Missing <a:sz> tag for font size.", "Attribute 'val' for <a:sz> is '11', should be '1100'.", "Changes seem applied outside the element hinted as 'title'."). If status is `"Valid"`, return an empty list `[]`.
    *   `confidence_score`: (Float) Your confidence in this semantic validation (0.0 to 1.0).

    **Example Output JSON:**
    ```json
    {{
    "task_description": "Change font size to 11pt for the title.",
    "semantic_validation_status": "Invalid",
    "issues_or_warnings": [
        "Generated XML contains <a:latin typeface='Times New Roman'/> but is missing the required <a:sz val='1100'/> tag within the relevant <a:rPr>."
    ],
    "confidence_score": 0.95
    }}

    """

    task_desc = task.get("task_description", "Unknown Task")
    slide_image_bytes = slide_context.get("slide_image_bytes")
    main_prompt = XML_SEMANTIC_VALIDATION_PROMPT.format(task_spec_json=json.dumps(task, default=str, indent=2), modified_xml=xml_code)

    final_prompt = [main_prompt]
    slide_image_text_prompt = "The below is visual representation of the slide image. Use it to understand and visualize the current layout, structure, elements, spacing, overlaps, colors, and styles."
    final_prompt.append(slide_image_text_prompt)

    image = genai.types.Part.from_bytes(data=slide_image_bytes, mime_type="image/png") 
            
    try:
        response = client.models.generate_content(model="gemini-2.0-flash", contents=[final_prompt, image])
        json_match = re.search(r'```json\s*(\{.*?\})\s*```', response.text, re.DOTALL)
        if not json_match:
            raise ValueError("No valid JSON object found in LLM response")

        json_text = json_match.group(0)
        try:
            json_response = json.loads(json_text)
        except json.JSONDecodeError:
            logging.error(f"Failed to parse JSON from LLM response for task '{task_desc}'")
            return False, f"XML validation failed: Invalid JSON in LLM response"      
        
        semantic_validation_status = json_response.get("semantic_validation_status", "Unknown Status")
        issues_or_warnings = json_response.get("issues_or_warnings", [])    
        confidence_score = json_response.get("confidence_score", 0.0)
        
        if semantic_validation_status == "Valid":
            logging.info(f"XML validation passed for task '{task_desc}'")
            return True, f"XML validation passed for task '{task_desc}'"
        elif semantic_validation_status == "Valid_With_Warnings":
            logging.warning(f"XML validation with warnings for task '{task_desc}': {issues_or_warnings}")
            return False, f"XML validation with warnings for task '{task_desc}': {issues_or_warnings}"
        elif semantic_validation_status == "Invalid":
            logging.error(f"XML validation failed for task '{task_desc}': {issues_or_warnings}")
            return False, f"XML validation failed for task '{task_desc}': {issues_or_warnings}"
        else:
            logging.warning(f"Unexpected response from LLM for task '{task_desc}': {json_response}")
            return False, f"XML validation failed for task '{task_desc}': {json_response}"
    except (json.JSONDecodeError, ValueError) as e:
        logging.error(f"Failed to parse LLM response for task '{task_desc}': {e}")
        return False, f"XML validation failed for task '{task_desc}': {e}"
    except Exception as e:
        logging.error(f"XML validation failed for task '{task_desc}': {str(e)}")
        return False, f"XML validation failed for task '{task_desc}': {str(e)}"
    


def validate_xml_code(xml_content: str, task : Dict[str, Any], slide_context: Dict[str, Any]) -> Tuple[bool, str]:
    task_desc = task.get("task_description", "Unknown Task")
    if not xml_content or not isinstance(xml_content, str) or not xml_content.strip().startswith('<'):
        msg = f"Validation failed for task '{task_desc}': Input is not a non-empty XML string."
        logging.warning(msg)
        return False, msg
    try:
        root = ET.fromstring(xml_content)
        
        # Check root element (less robust namespace handling)
        if root.tag != f'{{{PPTX_NAMESPACES["p"]}}}sld':
                msg = f"Validation failed for task '{task_desc}': Missing or incorrect root <p:sld> element. Found: {root.tag}"
                logging.warning(msg)
                return False, msg

        # Check structure (less robust namespace handling)
        cSld = root.find('p:cSld', PPTX_NAMESPACES)
        if cSld is None:
            msg = f"Validation failed for task '{task_desc}': Missing <p:cSld> element."
            logging.warning(msg)
            return False, msg
        
        spTree = cSld.find('p:spTree', PPTX_NAMESPACES)
        if spTree is None:
            msg = f"Validation failed for task '{task_desc}': Missing <p:spTree> element."
            logging.warning(msg)
            return False, msg
        
        logging.info("Basic XML structural validation passed. Proceeding to LLM validation...")
        
        is_contextually_valid, context_msg = validate_xml_with_llm(xml_content, task, slide_context)
        if not is_contextually_valid:
            return False, f"LLM validation failed for task '{task_desc}': {context_msg}"
        
        return True, "XML is valid for the specified task"
    
    except ET.ParseError as e:
        msg = f"Validation failed for task '{task_desc}': XML parsing error: {str(e)}"
        logging.error(msg)
        logging.error(f"Invalid XML snippet (first 500 chars):\n{xml_content[:500]}")
        return False, msg
    except Exception as e:
        msg = f"Validation failed for task '{task_desc}': Unexpected error: {str(e)}"
        logging.error(msg, exc_info=True)
        return False, msg


def correct_xml_code_with_llm(xml_code, original_task, validation_error):
    XML_CORRECTION_PROMPT = """
    
    # ROLE: AI PowerPoint XML Debugger & Corrector

    You are an expert AI assistant specializing in debugging and correcting invalid or incomplete PowerPoint PresentationML XML code. 
    You will receive a snippet of potentially invalid XML generated for a specific task, along with the validation error message encountered.

    **## Input Provided:**
    1.  **Original Task Description:** `{original_task_description}` (What the XML was supposed to achieve)
    2.  **Invalid XML Snippet:**
        ```xml
        {invalid_xml_code}
        ```
    3.  **Validation Error Message:** `{validation_error}` (The error reported by the XML parser or validator)


    **## Your Task:**
    1.  **Analyze the Error:** Understand the `Validation Error Message` in the context of the `Invalid XML Snippet`.
    2.  **Identify the Problem:** Pinpoint the exact cause of the error (e.g., syntax error, missing tag, incorrect namespace, invalid attribute value, incomplete structure).
    3.  **Correct the XML:** Modify the `Invalid XML Snippet` to fix ONLY the identified error(s), ensuring the corrected XML is well-formed and structurally sound according to PresentationML schema basics.
    4.  **Preserve Intent:** Ensure the corrected XML still attempts to fulfill the `Original Task Description`. Do not remove unrelated valid parts of the XML unless necessary to fix the core error.
    5.  **Preserve Structure:** Maintain the overall XML structure, namespaces, and existing valid elements as much as possible.


    **## CRITICAL Output Requirements:**
    *   Return ONLY the **complete, corrected XML content** for the slide (`<p:sld>...</p:sld>`).
    *   Ensure the output is well-formed XML.
    *   DO NOT include explanations, apologies, markdown formatting, or any text other than the corrected XML code itself.
    *   If the error cannot be reliably fixed or the input XML is too corrupted, return the original `Invalid XML Snippet` unmodified.

    """

    try:
        prompt = XML_CORRECTION_PROMPT.format(original_task_description=original_task, invalid_xml_code=xml_code, validation_error=validation_error)
        response = client.models.generate_content(model="gemini-2.0-flash", contents=[prompt])
        corrected_xml_match = re.search(r'```xml\s*(.*?)\s*```', response.text, re.DOTALL)
        if corrected_xml_match:
            corrected_xml = corrected_xml_match.group(1).strip()
            return corrected_xml
        else:
            logging.error("Failed to extract corrected XML from LLM response")
            return None
    except Exception as e:
        logging.error(f"LLM correction failed: {str(e)}")
        return None