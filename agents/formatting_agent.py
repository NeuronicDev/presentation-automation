import logging, json
from typing import Dict, Any, List
from langchain_core.messages import HumanMessage

from config.llm_provider import gemini_flash_llm

FORMATTING_TASK_DESCRIPTION_PROMPT  = """
    You are an expert AI assistant acting as a bridge between a parsed user request and a PowerPoint code generator. 
    Your task is to generate a clear, detailed natural language description for formatting changes based on the given instruction and slide context that needs to be performed on a PowerPoint slide. 
    This description will guide the subsequent code generation step.
    
    **Input:**
    Original User instruction: {original_instruction}
    Slide Number: {slide_number}
    action: {action}
    target_element_hint: {target_element_hint}  
    params: {params}
    Slide XML Structure: {slide_xml_structure}
    
    
    **You are provided with:**
    1.  **Original User instruction:** The high-level feedback instruction provided by the user.
    2.  **Slide Number:** The target slide for the modification.
    3.  **Specific Sub-Task Details:**
        *   `action`: The programmatic action to perform (e.g., 'change_font', 'align_elements').
        *   `target_element_hint`: A text hint describing the target element(s) (e.g., 'title', 'the chart on left').
        *   `params`: Specific parameters for the action (e.g., {{'font_name': 'Arial', 'size': 12}}).
    4.  **Slide Context :**
        *   `slide_xml_structure`: A representation of the slide's current XML structure.
        *   `slide_image_base64`: A base64 encoded image representing the slide's current visual appearance.


    **Your Goal:**
    Generate a detailed, unambiguous natural language description of the formatting change. Explain precisely:
    *   **EXISTS** in the current state of the slide.
    *   **WHAT** change needs to be made.
    *   **WHERE** it applies.
    *   **HOW and ON WHICH ELEMENTS** it should be done.
    *   Relate it back to the **Original User Request** for context.

    **Output Format:**
    Provide **only** the natural language description text. Do not include any preamble, JSON formatting, or markdown. JUST the description.

    """

def formatting_agent(classified_instruction: Dict[str, Any], slide_context: Dict[str, Any]) -> list[Dict[str, Any]]:
    processed_subtasks = []
    slide_number = classified_instruction.get("slide_number")
    original_instruction = classified_instruction.get("original_instruction", "")
    sub_tasks = classified_instruction.get("tasks", [])
    
    if not isinstance(sub_tasks, list) or not sub_tasks:
        logging.warning(f"Formatting agent received task with no valid sub-tasks: {classified_instruction}")
        return []

    for sub_task in sub_tasks:
        action = sub_task.get("action")
        target_hint = sub_task.get("target_element_hint")
        params = sub_task.get("params", {})
        
        if not action:
            logging.warning(f"Skipping sub-task with no action: {sub_task} in instruction: '{original_instruction}'")
            continue

        slide_xml = slide_context.get("slide_xml_structure", "")
        slide_image_base64 = slide_context.get("slide_image_base64", "")

        final_prompt = []
        
        slide_image_text_prompt = {
            "type": "text",  
            "text": "The below is the image of the slide. Please also use this as a reference to generate the description.",
        }
        final_prompt.append(slide_image_text_prompt)
        
        slide_image_base64_prompt = {
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{slide_image_base64}"},
        }
        final_prompt.append(slide_image_base64_prompt)
        
        main_prompt = FORMATTING_TASK_DESCRIPTION_PROMPT.format(
            original_instruction=original_instruction,
            slide_number=slide_number,
            action=action,
            target_element_hint=target_hint,
            params=json.dumps(params),
            slide_xml_structure=slide_xml,
        )
        final_prompt.append(main_prompt) 
            
        try:
            response = gemini_flash_llm.invoke([HumanMessage(content=final_prompt)])
            description = response.strip()
        except Exception as e:
            logging.error(f"Error generating description for formatting task: {e}")
            description = "Failed to generate description."
        
        flattened_task = {
            "agent_name": "formatting",
            "slide_number": slide_number,
            "original_instruction": original_instruction,
            "task_description": description, 
            "action": action,
            "target_element_hint": target_hint,
            "params": params
        }
        processed_subtasks.append(flattened_task)
        logging.debug(f"Formatting agent processed sub-task: {action} for slide {slide_number}")

    return processed_subtasks

