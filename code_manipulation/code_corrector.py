import logging, json, re
from typing import Dict, Any, Optional

from langchain_core.messages import HumanMessage
from langchain.prompts import PromptTemplate

from config.llm_provider import gemini_flash_llm  

CODE_CORRECTION_PROMPT = """
    You are an expert AI code correction agent specialized in debugging and fixing Python code snippets for PowerPoint manipulation using the `python-pptx` library.

    **Goal:**
    Analyze the provided Python code snippet and the error message it produced during execution within a Docker container. 
    Identify the root cause of the error and generate a corrected version of the code snippet that fixes the error while still achieving the original task.

    **Input Context:**
    
    - **Original action and Task Description:** 
        "{action}", "{task_description}" 
        
    - **Failing Python Code Snippet:**
        ```python
        {failing_code}
        ```
        
    - **Error Message:** 
        "{error_message}"

    **VERIFIED IMPORTS:**
    import pptx
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu, Cm
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, PP_PARAGRAPH_ALIGNMENT, MSO_TEXT_UNDERLINE_TYPE
    from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_CONNECTOR_TYPE, MSO_AUTO_SHAPE_TYPE, PP_PLACEHOLDER_TYPE, PP_MEDIA_TYPE
    from pptx.enum.dml import MSO_FILL_TYPE, MSO_LINE_DASH_STYLE, MSO_COLOR_TYPE, MSO_PATTERN_TYPE, MSO_THEME_COLOR_INDEX
    from pptx.chart.data import CategoryChartData, ChartData, XyChartData, BubbleChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_MARK, XL_TICK_LABEL_POSITION, XL_MARKER_STYLE, XL_DATA_LABEL_POSITION
    from pptx.enum.action import PP_ACTION_TYPE
    from pptx.table import Table, _Cell 
    

    **Instructions:**
    1.  **Understand the Task:** Review the task description to understand the original goal of the code.
    2.  **Analyze the Error Message:** Carefully examine the `error_message` to pinpoint the type of error (e.g., `AttributeError`, `TypeError`, `IndexError`, `ImportError`, syntax error) and the line of code where it occurred.
    3.  **Inspect the Failing Code:** Analyze the `failing_code` snippet, paying close attention to the line indicated in the error message and the surrounding code.
    4.  **Identify the Root Cause:** Determine the reason for the error. Is it a typo, incorrect `python-pptx` library usage, missing import, logical error in the code, or an assumption about the PPT structure that is incorrect?
    5.  **Generate Corrected Code:**  Modify the `failing_code` to fix the identified error and ensure it still performs the intended task as described in the `task_description`.
    6.  **Maintain Original Intent:** Ensure that the corrected code still aims to achieve the *original* task described in `task_description`. Do not change the overall goal, just fix the error.
    7.  **Output:** Provide **only** the corrected Python code snippet as text. Do not include any preamble, explanations, comments (unless essential for very complex fixes), or markdown formatting. JUST the corrected code.

    """


def generate_code_with_retry(failing_code: str, error_message: str, agent_task_specification: Dict[str, Any]) -> Optional[str]:
    logging.info(f"Attempting CODE CORRECTION for task: {agent_task_specification.get('action')} with error: {error_message}...")
    try:
        task_description = agent_task_specification.get("task_description")
        action = agent_task_specification.get("action")
        target_element_hint = agent_task_specification.get("target_element_hint")
        params = agent_task_specification.get("params", {})


        prompt_template = PromptTemplate.from_template(CODE_CORRECTION_PROMPT)
        prompt_input = {
            "task_description": task_description, 
            "action": action,
            "failing_code": failing_code,
            "error_message": error_message
        }

        chain = prompt_template | gemini_flash_llm

        try:
            response = chain.invoke(prompt_input)
            generated_code_str = response.strip()
            # logging.info(f"Generated code:\n {generated_code_str}")
            code_block = re.search(r'```python\n(.*?)\n```', generated_code_str, re.DOTALL)
            
            if code_block:
                extracted_code = code_block.group(1)
                logging.info(f"GENERATED CORRECTED CODE:\n{extracted_code}")
                return extracted_code
            else:
                logging.info(f"No markdown code block found. Using entire string.")
                return generated_code_str
        except Exception as e:
            logging.info(f"Error generated code: {e}")
            raise 
    
    except Exception as e:
        logging.error(f"Error during LLM-based code correction for task: {agent_task_specification.get('action')}. {e}", exc_info=True)
        return None
