import json, logging, re
from typing import Dict, Any, Optional

from langchain.prompts import PromptTemplate
from langchain_core.output_parsers import StrOutputParser
from langchain_core.messages import HumanMessage

from config.llm_provider import gemini_flash_llm  


CODE_GENERATION_PROMPT = """
    You are an expert AI code generator specialized in creating Python code using the `python-pptx` library to manipulate PowerPoint presentations.
    Your task is to generate precise, executable Python code snippets to modify a PowerPoint slide based on the provided task details and slide context.
    Ensure the generated code accurately performs the specified action on the target element(s) while preserving the original slide content.

    **Input Context:**
    - Agent: {agent_name}
    - Slide Index: {slide_index}
    - Original Instruction: {original_instruction}
    - Task Description: {task_description}
    - Action: {action}
    - Target Element Hint: {target_element_hint}
    - Parameters: {params}
    - Slide XML Structure: {slide_xml_structure}


    **Instructions & Assumptions:**
    - Understand the task description and original instruction provided by the agent.
    - Analyze the slide's current state using the provided image and XML structure and determine the changes to be made.
    - Identify the target element(s) based on the `target_element_hint` and the slide's context.
    - Assume you are working within a Python environment where the `python-pptx` library is already imported and available as `pptx`.
    - Assume you have access to a `slide` object representing the target slide (already obtained from `prs.slides[slide_index]`).
    - You have access to `pptx.util` for units like `Pt`, `Inches`, etc and use them where appropriate.
    - You have access to `pptx.enum.text` and other relevant `python-pptx` enums if needed.
    - Handle potential errors gracefully within the snippet if possible (e.g., check `shape.has_text_frame` before accessing `text_frame`).  
    - Ensure the original text and data remain intact after applying the modifications. Do not create new content.  
    - Ensure the code is executable and modifies the slide accurately based on the specified action and parameters.


    **Output:** 
    - Provide **only** the executable Python code snippet as text.  Do not include any preamble, explanations, comments or markdown formatting. JUST the code.
    - Ensure the code snippet is formatted correctly and ready for execution within a Python environment.
    
    """


def generate_python_code(agent_task_specification: Dict[str, Any], slide_context: Dict[str, Any]) -> Optional[str]:
    try:
        agent_name = agent_task_specification.get("agent_name")
        slide_number = agent_task_specification.get("slide_number")
        original_instruction = agent_task_specification.get("original_instruction")
        task_description = agent_task_specification.get("task_description")
        action = agent_task_specification.get("action")
        target_element_hint = agent_task_specification.get("target_element_hint")
        params = agent_task_specification.get("params", {})
        slide_xml_structure = slide_context.get("slide_xml_structure")
        slide_image_base64 = slide_context.get("slide_image_base64")   

        main_prompt = CODE_GENERATION_PROMPT.format(
            agent_name=agent_name,
            slide_index=slide_number - 1,
            original_instruction=original_instruction,
            task_description=task_description,
            action=action,
            target_element_hint=target_element_hint,
            params=params,
            slide_xml_structure=slide_xml_structure,
        )       
        
        final_prompt = [
            {
                "type": "text",
                "text": "The below is the image of the slide. Use it to understand and visualize the current layout and elements.",
            },
            {
                "type": "image_url",
                "image_url": {"url": f"data:image/jpeg;base64,{slide_image_base64}"},
            },
            {
                "type": "text",
                "text": main_prompt,
            }
        ]
        
        response = gemini_flash_llm.invoke([HumanMessage(content=final_prompt)])
        
        try:
            generated_code_str = response.strip()
            logging.info(f"Generated code: {generated_code_str}")
            code_block = re.search(r'```python\n(.*?)\n```', generated_code_str, re.DOTALL)
            
            if code_block:
                extracted_code = code_block.group(1)
                logging.info(f"Extracted code: {extracted_code}")
                return extracted_code
            else:
                logging.info(f"No markdown code block found. Using entire string.")
                return generated_code_str
        except Exception as e:
            logging.info(f"Error generated code: {e}")
            raise 
    
    except Exception as e:
        logging.error(f"Error generating Python code for task: {agent_task_specification.get('action')}. {e}", exc_info=True)
        return None

