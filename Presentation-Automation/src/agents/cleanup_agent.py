import logging, json, re
from typing import Dict, Any, List
from langchain_core.messages import HumanMessage
from config.llmProvider import gemini_flash_llm
from config.config import LLM_API_KEY
from google import genai
client = genai.Client(api_key=LLM_API_KEY)

CLEANUP_TASK_DESCRIPTION_PROMPT  = """
    You are an expert AI assistant acting as a bridge between a parsed user request and a PowerPoint code generator.
    Your task is to analyze a user's cleanup request and the current slide's visual state, then generate a clear, detailed natural language description of the specific cleanup changes required.
    This description will guide the subsequent code generation step, focusing on improving layout, alignment, spacing, consistency, and ensuring all elements fit within the slide boundaries without drastic repositioning or deletion.
    
    **Input:**
    Original User instruction: {original_instruction}
    Slide Number: {slide_number}
    action: {action} # The initial classification of the user's intent
    target_element_hint: {target_element_hint} # Hint for target elements, if any
    params: {params} # Specific parameters from initial parsing
    Slide Context: slide_xml_structure: {slide_xml_structure} # XML structure for element identification
        
    **## Your Task (Conditional Analysis and Description Generation):**

    1.  **Analyze Input Task:** Examine the initial classified task details (`action`, `target_element_hint`, `params`) derived from the `original_instruction`. Understand the user's primary cleanup goal as classified. Note if the request is general (e.g., "cleanup", "fit to page") or specific.
    2.  **Analyze Slide Context (Visual and Structural):**
        *   **Always** examine the `slide_image_base64` and relevant parts of the `slide_xml_structure`.
        *   **Identify Specific Layout Issues:**
            *   **Overflow:** Identify any element(s) extending beyond the standard slide boundaries (approx. 960x540 units). Pay special attention if the `original_instruction` mentioned "fit to page" or "fit screen".
            *   **Misalignment:** Check for elements within logical rows, columns, or groups that are not properly aligned (e.g., tops, centers, left edges not matching when they visually should). Look for deviations > 5px.
            *   **Uneven Spacing (Distribution):** Check for inconsistent gaps between elements in sequences (horizontally or vertically). Look for uneven distribution within logical groups.
            *   **Inconsistent Sizing:** Identify visually similar elements (e.g., boxes in a row, icons in a list) that have significantly different widths or heights (>5px difference) without apparent reason.
            *   **Poor Grouping/Overlap:** Note elements that seem logically related but are positioned disjointedly, or elements that overlap awkwardly.
            *   **Text Fit:** Check if text clearly overflows its container shape.
        *   **Identify Logical Groups:** Recognize sets of elements that function as a unit (e.g., a row of process steps, a column of list items, items within a container shape) as cleanup actions often apply collectively.
        *   **Assess Overall Composition:** Consider the balance and visual hierarchy.

    3.  **Generate Output Task Description(s):**
        *   **IF** the input task (`action`, `target_element_hint`) was already highly specific (e.g., "align top edge of shape A and B", "resize title text"):
            *   Generate a single, detailed `task_description`. Explain WHAT specific change to make, WHERE (identifying target elements clearly using hints/context/XML), and HOW (using params if provided, or inferring details like alignment type). Confirm this action aligns with the visual context.
            *   Output format: **Output A**.
        *   **ELSE IF** the input task was general (`cleanup_slide`, `adjust_layout`, `fit_to_page`) OR the `original_instruction` was broad ("Clean this up", "Make it fit"):
            *   Generate a **list** of concrete sub-tasks (`expanded_tasks`) based on the issues identified in Step 2.
            *   Each sub-task should address **one primary issue** (e.g., fix overflow for group X, align row Y, standardize sizes of Zs, distribute column A).
            *   Assign specific `action` verbs (e.g., `resize_shape`, `move_shape`, `align_elements`, `distribute_elements`, `standardize_dimensions`, `bring_within_bounds`).
            *   Write a clear `task_description` for each sub-task:
                *   **WHAT:** Describe the specific action (e.g., "Resize", "Align left edges", "Distribute vertically").
                *   **WHICH:** Identify the target element(s) or logical group clearly (e.g., "the blue chevron shapes in the top row", "all rounded rectangles in the left list", "the three image placeholders in the right column"). Use hints from XML if helpful.
                *   **HOW:** Specify parameters (e.g., "to have a consistent width of 80px", "so they are evenly spaced", "relative to each other", "to fit within the slide margins").
                *   **CONTEXT/WHY:** Briefly explain the reason, linking back to the overall goal (e.g., "to ensure all elements fit on the slide", "for improved visual consistency", "to create a cleaner vertical flow").
            *   **Prioritization:** Address `bring_within_bounds` / overflow issues first. Then tackle major alignment, spacing, and sizing inconsistencies.
            *   **Constraint Handling:**
                *   **Fit to Page/Screen:** If requested, tasks should likely involve *resizing* multiple elements smaller and *redistributing* them more compactly. The description must reflect this combination.
                *   **Minimal Movement:** When generating `move_shape` or `resize_shape` tasks, aim to preserve the element's general location unless moving is essential for alignment, distribution, or fitting. The description should implicitly favor minimal necessary adjustments.
                *   **Group Consideration:** When adjusting one element in a group, consider the knock-on effects. Descriptions might need to specify adjusting related elements simultaneously (e.g., "Resize the container box and vertically distribute the logos within it.").
            *   **CRITICAL RULE:** **You MUST NOT suggest deleting or removing any shape.** Focus only on resizing, repositioning, aligning, distributing, and standardizing existing elements. If text is problematic beyond fitting, flag it but do not delete shapes.
            *   Output format: **Output B**.
        
    **Your Goal:**
    Generate a detailed, unambiguous natural language description of the cleanup changes. Explain precisely:
    *   **EXISTS** in the current state of the slide.
    *   **WHAT** change needs to be made.
    *   **WHERE** it applies.
    *   **HOW and ON WHICH ELEMENTS** it should the changes be done such that origial user instruction is met.
    *   Relate it back to the **Original User instruction and Slide Image** for context.

    **## Output Requirements:**
    Respond ONLY with a single, valid JSON object in ONE of the following formats. Do NOT include explanations or markdown formatting outside the JSON structure.

    **Output A (For Specific, Pre-defined Tasks):**
    {{
    "task_description": "Detailed natural language description of the single specific cleanup action, clearly identifying the target(s), the precise change, and referencing the visual context or original specific request."
    }}

    **Output B (For General Cleanup / Context-Derived Tasks):**
    {{
    "expanded_tasks": [
        {{
        "action": "specific_action_1", // e.g., bring_within_bounds, resize_shape, align_elements
        "task_description": "Detailed natural language description for sub-task 1: WHAT action, WHICH element(s)/group, HOW (specific parameters/targets), and CONTEXT/WHY (e.g., 'Resize the five blue vertical shapes on the left significantly smaller and align their top edges to ensure they fit within the slide height.')",
        "target_element_hint": "hint_for_action_1", // e.g., "Blue vertical shapes on left", "Checkmarks in middle column"
        "params": {{ ...params_for_action_1... }} // e.g., {{"target_dimension": "height", "value": 100}}, {{"alignment": "top"}}
        }},
        {{
        "action": "specific_action_2", // e.g., distribute_elements, standardize_dimensions
        "task_description": "Detailed natural language description for sub-task 2: WHAT, WHICH, HOW, CONTEXT/WHY (e.g., 'Vertically distribute the rows containing the blue shapes and grey boxes evenly within the available space below the header.')",
        "target_element_hint": "hint_for_action_2", // e.g., "Rows containing content", "Logos in middle column, second group"
        "params": {{ ...params_for_action_2... }} // e.g., {{"direction": "vertical"}}
        }}
        // ... potentially more specific sub-tasks identified ...
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

    return processed_subtasks