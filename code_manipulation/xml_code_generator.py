import json, logging, re
from typing import Dict, Any, Optional

from langchain.prompts import PromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_core.messages import HumanMessage

from config.llm_provider import gemini_flash_llm  
from config.config import LLM_API_KEY

from google import genai
client = genai.Client(api_key=LLM_API_KEY)

XML_MODIFICATION_PROMPT = """
    # PowerPoint XML Modification Expert
    You are a **State-of-the-Art expert in PowerPoint XML structure** with deep expertise in modifying PPTX XML to implement presentation changes accurately. 
    Your task is to modify the provided XML content to implement specific changes based on the provided feedback instruction, task details and slide context. while maintaining the overall structure and validity of the XML.

    ## Instructions:
    1. Analyze the provided XML structure carefully to understand the current state of the slide.
    2. Identify the exact XML elements that need to be modified based on the task description and target element hint.
    3. Make precise, targeted changes to implement the requested modifications.
    4. Ensure your modified XML maintains valid PPTX XML structure.
    5. Only modify what's necessary - preserve all other content and structure.

    ## Input Context:
        - Agent: {agent_name}
        - Slide Index: {slide_index}
        - Original Instruction: {original_instruction}
        - Task Description: {task_description}
        - Action: {action}
        - Target Element Hint: {target_element_hint}
        - Parameters: {params}
        - Slide XML Structure: {slide_xml_structure}


    ## IMPORTANT: PowerPoint XML Guidance:

    ### Common XML Namespaces in PPTX:
    - a: http://schemas.openxmlformats.org/drawingml/2006/main
    - p: http://schemas.openxmlformats.org/presentationml/2006/main
    - r: http://schemas.openxmlformats.org/officeDocument/2006/relationships

    ### Key XML Elements:
    - `<p:sp>`: Shape element (includes text boxes)
    - `<p:txBody>`: Text content container
    - `<a:p>`: Paragraph within text
    - `<a:r>`: Text run (contains text with consistent formatting)
    - `<a:t>`: Actual text content
    - `<a:solidFill>`: Solid color fill
    - `<a:srgbClr val="RRGGBB">`: RGB color definition
    - `<a:tbl>`: Table element
    - `<a:tr>`: Table row
    - `<a:tc>`: Table cell
    - `<p:graphicFrame>`: Container for tables, charts
    - `<p:pic>`: Image element

    ### Formatting Properties:
    - Font properties: `<a:latin typeface="Arial"/>`, `<a:sz val="2400"/>` (1/100 point)
    - Text alignment: `<a:pPr algn="ctr"/>` (values: l, ctr, r, just, dist)
    - Bold text: `<a:b val="1"/>`
    - Italic text: `<a:i val="1"/>`
    - Line spacing: `<a:lnSpc><a:spcPts val="1200"/></a:lnSpc>` (1/100 point)

    ### Positioning and Size:
    - Shape position: `<p:spPr><a:xfrm><a:off x="1234" y="5678"/></a:xfrm></p:spPr>` (EMUs)
    - Shape size: `<p:spPr><a:xfrm><a:ext cx="1234" cy="5678"/></a:xfrm></p:spPr>` (EMUs)
    - One inch = 914400 EMUs

    ### Color Values:
    - RGB Hex: `<a:srgbClr val="FF0000"/>` (red)
    - RGB Percentage: `<a:scrgbClr r="1.0" g="0.0" b="0.0"/>` (red)
    - Theme Color: `<a:schemeClr val="accent1"/>`


    
    ## OUTPUT REQUIREMENTS:
    1. Return ONLY the modified XML content - no explanations or markdown formatting.
    2. Ensure the XML is well-formed and valid.
    3. Do not introduce any new namespaces or non-standard elements.
    4. Preserve the XML declaration and all namespace declarations.

    """


def generate_modified_xml_code(original_xml: str, agent_task_specification: Dict[str, Any], slide_context: Dict[str, Any]) -> Optional[str]:
    try:
        logging.info(f"Generating modified XML code...: {agent_task_specification}")
        agent_name = agent_task_specification.get("agent_name")
        slide_number = agent_task_specification.get("slide_number")
        original_instruction = agent_task_specification.get("original_instruction")
        task_description = agent_task_specification.get("task_description")
        action = agent_task_specification.get("action")
        target_element_hint = agent_task_specification.get("target_element_hint")
        params = agent_task_specification.get("params", {})
        slide_xml_structure = slide_context.get("slide_xml_structure")
        slide_image_base64 = slide_context.get("slide_image_base64") 
        slide_image_bytes = slide_context.get("slide_image_bytes")  

        main_prompt = XML_MODIFICATION_PROMPT.format(
            agent_name=agent_name,
            slide_index=slide_number - 1,
            original_instruction=original_instruction,
            task_description=task_description,
            action=action,
            target_element_hint=target_element_hint,
            params=json.dumps(params, indent=2),
            slide_xml_structure=original_xml
        )
        
        
        final_prompt = [main_prompt]
        slide_image_text_prompt = "The below is visual representation of the slide image. Use it to understand and visualize the current layout, structure, elements, spacing, overlaps, colors, and styles."
        final_prompt.append(slide_image_text_prompt)

        image = genai.types.Part.from_bytes(data=slide_image_bytes, mime_type="image/png") 
                
        try:
            response = client.models.generate_content(model="gemini-2.0-flash", contents=[final_prompt, image])
            modified_xml_str = response.text.strip()
            
            if not modified_xml_str.startswith('<?xml'):
                xml_match = re.search(r'```xml\s+(.*?)\s+```', modified_xml_str, re.DOTALL)
                if xml_match:
                    modified_xml_str = xml_match.group(1)
                else:
                    logging.warning("Modified XML doesn't start with XML declaration. This might not be valid XML.")
                    
            logging.info(f"Successfully generated modified XML for task: {task_description}")
            return modified_xml_str
        
        except Exception as e:
            logging.info(f"Error generated XML: {e}")
            raise 
    
    except Exception as e:
        logging.error(f"Error generating modified XML: {e}", exc_info=True)
        return None