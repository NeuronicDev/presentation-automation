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

    # State-of-the-Art PowerPoint XML Modification Expert 

    You are a **world-leading AI expert in PowerPoint Open XML Presentation Language (PPTX XML)**, possessing unparalleled expertise in directly manipulating PPTX XML to achieve complex and varied presentation modifications.
    Your task is to modify the provided XML slide content to implement specific changes based on the provided feedback instruction, task details and slide context while maintaining the overall structure and validity of the XML.
    You will replace the *entire* original slide XML with your generated XML output.
    
    ## INPUT CONTEXT:
        - Agent: {agent_name}
    - Slide Index (0-based): {slide_index}
    - Original User Instruction: "{original_instruction}"
    - Task Description: "{task_description}"  
        - Action: {action}
        - Target Element Hint: {target_element_hint}
        - Parameters: {params}
    - Slide XML Structure: "{slide_xml_structure}" <- XML content of the slide (critical for analysis and modification).

    ## INSTRUCTIONS:
    1.  **Deep XML Analysis:** Thoroughly analyze the provided current state of the `slide_xml_structure`. Understand the XML hierarchy, namespaces, elements, and attributes relevant to the `task_description` and `target_element_hint`. Use the XML structure to precisely locate the elements you need to modify.
    2.  **PPTX XML Schema Mastery:** Leverage your expert knowledge of the PowerPoint Open XML Presentation Language schema. Understand how different PPTX features (text, shapes, tables, charts, images, layouts, themes, etc.) are represented in XML. Refer to the "IMPORTANT: PowerPoint XML Guidance" section below for key elements and attributes.
    3.  **Preserve Unmodified Parts:** **Critically important:** You must **preserve all parts of the original XML structure and content that are *not* directly related to the task.**  Your generated XML should be as close to the original XML as possible, with only the necessary modifications applied. Do not introduce unintended changes or deletions that could unintentionally alter other aspects of the slide.
    4.  **Prioritize Attribute Modification:** Whenever possible, prefer modifying XML attributes over more complex element creation/deletion. Attribute modifications are generally safer and less likely to corrupt the PPTX structure.
    5.  **Maintain XML Validity:** **ABSOLUTELY ENSURE that the generated XML is valid and well-formed PPTX slide XML.**  It must conform to the PowerPoint Open XML Presentation Language schema. Invalid XML will corrupt the PPTX file. Pay extreme attention to XML syntax, namespaces, element hierarchy, and attribute values.
    6.  **Reference Original XML Structure:** Use the `Original Slide XML Structure` as a template and starting point.  Do not generate XML from scratch. Instead, **modify the provided XML** to incorporate the changes. This is crucial for preserving the overall slide structure and minimizing errors.
    7.  **Handle Namespaces Correctly:**  PPTX XML uses namespaces extensively (e.g., `a:`, `p:`, `r:`). When generating XML instructions, **always be mindful of namespaces and use them correctly in XPath queries and element/attribute modifications.** Refer to the "Common XML Namespaces" section below.

    ## IMPORTANT: PowerPoint XML Guidance for XML Modification:
    ### Common XML Namespaces in PPTX:
    - `a`: http://schemas.openxmlformats.org/drawingml/2006/main (DrawingML - shapes, text, graphics)
    - `p`: http://schemas.openxmlformats.org/presentationml/2006/main (PresentationML - slides, slide structure)
    - `r`: http://schemas.openxmlformats.org/officeDocument/2006/relationships (Relationships)
    - `c`: http://schemas.openxmlformats.org/drawingml/2006/chart (ChartML - charts)
    - `dgm`: http://schemas.openxmlformats.org/drawingml/2006/diagram (DiagramML - SmartArt)
    - `pic`: http://schemas.openxmlformats.org/drawingml/2006/picture (PictureML - Images)
    - `w`: http://schemas.openxmlformats.org/wordprocessingml/2006/main (WordprocessingML - sometimes used in text content)

    ### Key XML Elements & Attributes for Feature Implementation:
    
    **Formatting Enhancements:**
    - **Text Elements:** `<a:t>` (text content), `<a:rPr>` (text run properties - font, size, color, bold, italic), `<a:pPr>` (paragraph properties - alignment, spacing, bullet style), `<a:txBody>` (text body container).
        - Attributes: `a:latin/@typeface` (font name), `a:sz/@val` (font size in EMU), `a:algn/@val` (text alignment - 'l', 'ctr', 'r', 'just', 'dist'), `a:b/@val` (bold - '1' or '0'), `a:i/@val` (italic - '1' or '0'), `a:spcPts/@val` (line spacing in 1/100 points), `<a:solidFill>/<a:srgbClr/@val` (RGB color).
    - **Table Elements:** `<a:tbl>` (table), `<a:tr>` (row), `<a:tc>` (cell), `<a:tcPr>` (cell properties - fill, borders), `<a:gridCol>` (column widths).
        - Attributes: `<a:gridCol w="EMU_VALUE"/>` (column width), `<a:tr h="EMU_VALUE"/>` (row height), `<a:tc>/<a:txBody>` (cell text content), `<a:tcPr>/<a:lnL>`, `<a:tcPr>/<a:lnR>`, `<a:tcPr>/<a:lnT>`, `<a:tcPr>/<a:lnB>` (cell border lines - use `<a:solidFill>` and `<a:prstDash>` within these for styling).
    - **Shape Elements:** `<p:sp>` (shape), `<p:spPr>` (shape properties - position, size, rotation, fill, line), `<a:xfrm>/<a:off x="EMU_VALUE" y="EMU_VALUE"/>` (position), `<a:xfrm>/<a:ext cx="EMU_VALUE" cy="EMU_VALUE"/>` (size), `<a:solidFill>/<a:srgbClr/@val` (fill color), `<a:ln>/<a:solidFill>/<a:srgbClr/@val` (line color), `<a:ln w="EMU_VALUE"/>` (line width), `<a:ln prstDash="STYLE_ENUM"/>` (line dash style - use `MSO_LINE_DASH_STYLE` enums).
    - **Slide Layout/Template:** `<p:sldLayout>` (slide layout), `<p:sldMaster>` (slide master), `<p:theme>` (theme - color schemes, fonts, effects).
        - Instructions for template application might involve replacing the entire `<p:sldLayout>` or modifying theme-related XML parts. This is COMPLEX and requires advanced PPTX XML knowledge.

    **Cleanup Enhancements:**
    - **Element Deletion:** Instructions might involve deleting specific shapes (`<p:sp>`), text boxes, or table rows/columns. Use XPath to target elements for deletion, but be VERY CAUTIOUS.
    - **Spacing & Alignment:** Adjust shape positions (`<a:xfrm>/<a:off>`), paragraph spacing (`<a:pPr>/<a:spcBef>`, `<a:pPr>/<a:spcAft>`), and text box margins (`<a:txBody>/<a:bodyPr lIns="EMU_VALUE" rIns="EMU_VALUE" tIns="EMU_VALUE" bIns="EMU_VALUE"/>`).
    - **Bullet Point Formatting:** Modify bullet point styles within `<a:pPr>/<a:bu...>` elements (bullet type, size, color).

    **Visual Enhancements:**
    - **Color & Opacity:** Use `<a:solidFill>/<a:srgbClr/@val` (RGB color), `<a:alpha val="PERCENTAGE"/>` (opacity/transparency).
    - **Backgrounds:** Modify slide background fill (`<p:cSld>/<p:bg>/<p:bgPr>`).
    - **Icons & Images:** `<p:pic>` (image element), `<p:blipFill>/<a:blip r:embed="RELATIONSHIP_ID"/>` (image embedding), `<p:sp>/<p:pic>` (icons often embedded as pictures within shapes). Instructions might involve replacing image embeddings or adjusting image positions/sizes.
    - **Timelines & Diagrams (SmartArt):** `<dgm:spTree>` (SmartArt diagrams). SmartArt XML is COMPLEX. Instructions for timeline/diagram redesign might be very challenging to implement via direct XML manipulation and might be better handled programmatically using `python-pptx` object model if possible.
    - **Effects (Shadows, 3D):** `<p:spPr>/<a:effectLst>` (effect list), `<a:outerShdw>` (outer shadow effect), `<a:innerShdw>` (inner shadow effect), `<a:prstShdw type="SHADOW_TYPE_ENUM"/>` (preset shadow types).

    **General XML Manipulation Guidelines:**
    *   **XPath is Your Friend:** Use XPath expressions to precisely target XML elements for modification. Learn XPath syntax for navigating PPTX XML.
    *   **Namespaces are Critical:** Always use namespaces correctly when querying or modifying XML elements (e.g., `namespaces={{'p': '...', 'a': '...'}}`).
    *   **EMU Units:** Remember that PowerPoint XML uses English Metric Units (EMU) for positioning and sizing. Convert units accordingly (1 inch = 914400 EMU, 1 point = 100 EMU).
    *   **Attribute Values as Strings:** Attribute values in XML are strings. Ensure you are setting string values (e.g., `"1100"` for font size, not integer `1100`) (1/100 point, e.g., 11pt = 1100)
    *   **Test Incrementally:**  For complex tasks, break down XML modifications into smaller steps and test incrementally to avoid introducing errors.
    *   **Validation is Key:** After any XML modification, always validate the PPTX file to ensure it is still valid and opens correctly in PowerPoint.
    
    "CRITICAL: You MUST include ALL original, unmodified elements from the input <p:spTree>...</p:spTree> unless the task explicitly requires deleting them. Only modify or add elements relevant to the task. DO NOT return only the modified parts; return the complete, modified <p:sld>...</p:sld> structure."
    
    ## OUTPUT REQUIREMENTS:
    Provide **only** the complete, modified XML code for the entire slide as text, enclosed in a ```xml code block. Ensure the XML is well-formed and valid.
    Do not include any preamble, explanations, step-by-step instructions, or markdown formatting outside the XML code block. 
    The output should be a self-contained, valid XML document representing the *entire* modified slide.
    
    **Begin! Generate the complete, modified XML code for the entire slide, incorporating the requested changes and preserving the rest of the slide structure:**

"""



def generate_modified_xml_code(original_xml: str, agent_task_specification: Dict[str, Any], slide_context: Dict[str, Any]) -> Optional[str]:
    try:
        logging.info(f"Generating modified XML code...")
        logging.info(f"Original XML: {original_xml[:100]}")
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
                    
            # logging.info(f"Successfully generated modified XML for task: {modified_xml_str}")
            return modified_xml_str
        
        except Exception as e:
            logging.info(f"Error generated XML: {e}")
            raise 
    
    except Exception as e:
        logging.error(f"Error generating modified XML: {e}", exc_info=True)
        return None