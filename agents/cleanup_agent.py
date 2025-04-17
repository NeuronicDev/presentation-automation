import logging, json, re
from typing import Dict, Any, List
from langchain_core.messages import HumanMessage

from config.llm_provider import gemini_flash_llm
from config.config import LLM_API_KEY

from google import genai
client = genai.Client(api_key=LLM_API_KEY)

CLEANUP_TASK_DESCRIPTION_PROMPT  = """
    You are an expert AI assistant acting as a bridge between a parsed user request and a PowerPoint code generator. 
    Your task is to generate a clear, detailed natural language description for cleanup changes and tasks based on the given instruction and slide context that needs to be performed on a PowerPoint slide. 
    This description will guide the subsequent code generation step.
    
    **Input:**
    Original User instruction: {original_instruction}
    Slide Number: {slide_number}
    action: {action}
    target_element_hint: {target_element_hint}  
    params: {params}
    Slide XML Structure: {slide_xml_structure}
    
    
    **## Your Task (Conditional):**

    1.  **Analyze Input:** Examine the provided `action`, `target_element_hint`, and `params`.
    2.  **Determine Specificity:**
        *   **IF** the `action` is specific (e.g., 'remove_elements', 'delete_shape', 'consolidate_text', 'reduce_content') and the `target_element_hint` clearly defines a single modification:
            *   Generate a detailed natural language description explaining WHAT cleanup change to make, WHERE it applies (using hints/context), and HOW (using params), relating it to the original request.
            *   Output ONLY this description in the specified JSON format below (Output A).
        *   **ELSE IF** the `action` is general/vague (e.g., 'cleanup_slide', 'remove_clutter', 'simplify') OR the original instruction is broad ("Clean this up", "Remove unnecessary elements", "Make slide cleaner"):
            *   Analyze the slide's current state using the provided Image and XML context.
            *   Identify specific cleanup improvements needed based on standard design principles (e.g., removing redundant elements, consolidating text, simplifying complex graphics, removing distracting elements).
            *   Generate a list of concrete, actionable sub-tasks required to clean up the slide according to these principles. Use standard action verbs (e.g., 'remove_elements', 'delete_shape', 'consolidate_text', 'reduce_content', 'remove_duplicates').
            *   Output ONLY this list of sub-tasks in the specified JSON format below (Output B).

    **You are provided with:**
    1.  **Original User instruction:** The high-level feedback instruction provided by the user.
    2.  **Slide Number:** The target slide for the modification.
    3.  **Specific Sub-Task Details:**
        *   `action`: The programmatic action to perform (e.g., 'remove_elements', 'adjust_spacing').
        *   `target_element_hint`: A text hint describing the target element(s) (e.g., 'bullet points', 'large table').
        *   `params`: Specific parameters for the action (e.g., {{"consistency": "uniform"}}).
    4.  **Slide Context :**
        *   `slide_xml_structure`: A representation of the slide's current XML structure.
        *   `slide_image_base64`: A base64 encoded image representing the slide's current visual appearance.


    **Your Goal:**
    Generate a detailed, unambiguous natural language description of the cleanup changes. Explain precisely:
    *   **EXISTS** in the current state of the slide.
    *   **WHAT** change needs to be made.
    *   **WHERE** it applies.
    *   **HOW and ON WHICH ELEMENTS** it should the changes be done such that origial user instruction is met.
    *   Relate it back to the **Original User instruction and Slide Image** for context.


    **## Output Requirements:**
    Respond ONLY with a single, valid JSON object in ONE of the following formats:
    CRITICAL: Choose ONLY ONE output format based on whether the input task was specific or vague. Do not include explanations or markdown.

    **Output A (For Specific Tasks):**
    {{
    "task_description": "Detailed natural language description of the single specific cleanup action..."
    }}

    **Output B (For General vague Tasks):**
    {{
    "expanded_tasks": [
        {{
        "action": "specific_action_1",
        "task_description": "Detailed natural language description of the specific cleanup action...",
        "target_element_hint": "hint_for_action_1",
        "params": {{ ...params_for_action_1... }}
        }},
        {{
        "action": "specific_action_2",
        "task_description": "Detailed natural language description of the specific cleanup action...",
        "target_element_hint": "hint_for_action_2",
        "params": {{ ...params_for_action_2... }}
        }}
        // ... more specific sub-tasks identified ...
    ]
    }}

    """

def cleanup_agent(classified_instruction: Dict[str, Any], slide_context: Dict[str, Any]) -> list[Dict[str, Any]]:
    processed_subtasks = []
    slide_number = classified_instruction.get("slide_number")
    original_instruction = classified_instruction.get("original_instruction", "")
    sub_tasks = classified_instruction.get("tasks", [])
    
    if not isinstance(sub_tasks, list) or not sub_tasks:
        logging.warning(f"cleanup_agent received task with no valid sub-tasks: {classified_instruction}")
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
        slide_image_bytes = slide_context.get("slide_image_bytes", "")

        final_prompt = []

        main_prompt = CLEANUP_TASK_DESCRIPTION_PROMPT.format(
            original_instruction=original_instruction,
            slide_number=slide_number,
            action=action,
            target_element_hint=target_hint,
            params=json.dumps(params),
            slide_xml_structure=slide_xml,
        )
        final_prompt.append(main_prompt) 

        slide_image_text_prompt ="The below is the image of the slide. Please also use this as a reference to generate the description. Analyse what text, images, shapes, other elements, structure and layout are currently present on the slide"
        
        final_prompt.append(slide_image_text_prompt)
        image = genai.types.Part.from_bytes(data=slide_image_bytes, mime_type="image/png") 

        try:
            response = client.models.generate_content(model="gemini-2.0-flash", contents=[final_prompt, image])
            logging.info(f"LLM cleanup_agent response: {response.text}")
            
            json_match = re.search(r'(\{[\s\S]*\})', response.text)
            
            if json_match:
                json_str = json_match.group(0)
                
                try:
                    mapping = json.loads(json_str)
                    
                    if "task_description" in mapping:
                        # Format A: Single task description
                        flattened_task = {
                            "agent_name": "cleanup",
                            "slide_number": slide_number,
                            "original_instruction": original_instruction,
                            "task_description": mapping["task_description"], 
                            "action": action,
                            "target_element_hint": target_hint,
                            "params": params
                        }
                        processed_subtasks.append(flattened_task)

                        
                    elif "expanded_tasks" in mapping and isinstance(mapping["expanded_tasks"], list):
                        # Format B: Multiple expanded tasks 
                        for expanded_task in mapping["expanded_tasks"]:
                            if not isinstance(expanded_task, dict):
                                logging.warning(f"Skipping invalid expanded task (not a dict): {expanded_task}")
                                continue
                                
                            flattened_task = {
                                "agent_name": "cleanup",
                                "slide_number": slide_number,
                                "original_instruction": original_instruction,
                                "task_description": expanded_task.get("task_description", "Missing description"),
                                "action": expanded_task.get("action", "unknown_action"),
                                "target_element_hint": expanded_task.get("target_element_hint", ""),
                                "params": expanded_task.get("params", {})
                            }
                            
                            # Ensure params is a dictionary
                            if not isinstance(flattened_task["params"], dict):
                                flattened_task["params"] = {}
                                
                            processed_subtasks.append(flattened_task)
                    else:
                        # Fallback if JSON doesn't match expected format
                        logging.warning(f"JSON response doesn't match expected format: {mapping}")
                        flattened_task = {
                            "agent_name": "cleanup",
                            "slide_number": slide_number,
                            "original_instruction": original_instruction,
                            "task_description": f"Parsing error: Unexpected JSON format. Raw response: {response.text[:100]}...", 
                            "action": action,
                            "target_element_hint": target_hint,
                            "params": params
                        }
                        processed_subtasks.append(flattened_task)
                        
                except json.JSONDecodeError as je:
                    logging.error(f"JSON parsing error: {je} for string: {json_str[:100]}...")
                    flattened_task = {
                        "agent_name": "cleanup",
                        "slide_number": slide_number,
                        "original_instruction": original_instruction,
                        "task_description": f"Failed to parse JSON response: {str(je)}", 
                        "action": action,
                        "target_element_hint": target_hint,
                        "params": params
                    }
                    processed_subtasks.append(flattened_task)
            else:
                logging.warning(f"No JSON found in LLM response: {response.text[:100]}...")
                flattened_task = {
                    "agent_name": "cleanup",
                    "slide_number": slide_number,
                    "original_instruction": original_instruction,
                    "task_description": "Failed to extract JSON from LLM response.", 
                    "action": action,
                    "target_element_hint": target_hint,
                    "params": params
                }
                processed_subtasks.append(flattened_task)
                
        except Exception as e:
            logging.error(f"Error in cleanup agent: {e}")
            flattened_task = {
                "agent_name": "cleanup",
                "slide_number": slide_number,
                "original_instruction": original_instruction,
                "task_description": f"Error processing cleanup task: {str(e)}", 
                "action": action,
                "target_element_hint": target_hint,
                "params": params
            }
            processed_subtasks.append(flattened_task)

    logging.info(f"Processed subtasks for cleanup agent final: {processed_subtasks}")
    return processed_subtasks
