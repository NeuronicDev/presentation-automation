# refiner_agent.py
import logging, json, re, asyncio, os, copy, aiofiles
from config.config import LLM_API_KEY
from google import genai
from typing import Dict, Any, List, Optional, Tuple
import google.api_core.exceptions

client = genai.Client(api_key=LLM_API_KEY)

logger = logging.getLogger(__name__)
REFINER_OUTPUT_DIR = "uploaded_pptx/slide_images/presentation" 
METADATA_DIR = "uploaded_pptx/slide_images/metadata"
TASKS_DIR = "uploaded_pptx/slide_images/presentation"

def _load_nl_subtasks(slide_number: int) -> List[str]:
    file_path = os.path.join(TASKS_DIR, f"slide{slide_number}_tasks.json")
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Task file not found for slide {slide_number} at {file_path}")
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            tasks_data = json.load(f)
        if not isinstance(tasks_data, list):
            raise ValueError(f"Invalid format in task file {file_path}: Expected a list.")
        instruction_list = [
            task.get("task_description", str(task)) for task in tasks_data
            if isinstance(task.get("task_description"), str)
        ]
        if not instruction_list:
            logger.warning(f"No valid 'task_description' strings found in {file_path}")
        return instruction_list
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON in task file {file_path}: {e}") from e
    except Exception as e:
        logger.error(f"Error loading NL tasks from {file_path}: {e}", exc_info=True)
        raise
    
def _load_and_copy_metadata(slide_number: int) -> Tuple[Optional[List[Dict[str, Any]]], Optional[List[Dict[str, Any]]]]:
    metadata_path = os.path.join(METADATA_DIR, f"metadata_{slide_number}.json")
    if not os.path.exists(metadata_path):
        logger.error(f"Metadata file not found for slide {slide_number} at {metadata_path}")
        return None, None
    try:
        with open(metadata_path, "r", encoding="utf-8") as f:
            shapes = json.load(f)
        if not isinstance(shapes, list):
            raise ValueError("Metadata file does not contain a valid list.")
        simulated_metadata = copy.deepcopy(shapes)
        logger.info(f"Loaded and copied metadata for slide {slide_number}")
        return shapes, simulated_metadata
    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON in metadata file {metadata_path}: {e}")
        return None, None
    except Exception as e:
        logger.error(f"Error loading metadata from {metadata_path}: {e}", exc_info=True)
        return None, None


def _update_simulated_metadata(simulated_metadata: List[Dict[str, Any]], changes: Dict[str, Any]) -> bool:
    if not changes or not simulated_metadata:
        return False

    updated_count = 0
    target_ids = changes.get('ids', [])
    prop = changes.get('property')
    value = changes.get('value')

    if not target_ids or not prop or value is None:
        logger.warning(f"Invalid changes data for metadata update: {changes}")
        return False

    shape_lookup = {str(shape.get("id", "")): shape for shape in simulated_metadata}

    for target_id in target_ids:
        shape = shape_lookup.get(target_id)
        if shape:
            if prop in shape:
                try:
                    current_value = shape[prop]
                    new_value = float(value)  
                    shape[prop] = new_value
                    logger.debug(f"Simulated update: ID {target_id}, Property '{prop}', Value '{current_value}' -> '{new_value}'")
                    updated_count += 1
                except (ValueError, TypeError) as e:
                    logger.error(f"Failed to apply update: Invalid value type '{value}' or conversion error for property '{prop}' on ID {target_id}: {e}")
                except Exception as e:
                    logger.error(f"Failed to apply update for ID {target_id}, property '{prop}': {e}")
            else:
                logger.warning(f"Property '{prop}' not found for shape ID {target_id} in simulated metadata.")
        else:
            logger.warning(f"Target shape ID {target_id} not found in simulated metadata lookup.")

    if updated_count < len(target_ids):
        logger.warning(f"Metadata update incomplete: Expected to update {len(target_ids)} shapes, updated {updated_count}.")
    return updated_count > 0

# --- Refiner Prompt Template ---
REFINER_PROMPT_TEMPLATE = f"""
You are an expert AI assistant acting as a meticulous layout refiner and translator. Your task is to convert **ONE** natural language (NL) PowerPoint modification instruction into an **explicit, executable, context-aware command string with precise calculations**. You MUST rigorously verify targets, understand the slide's intended logical structure and layout dependencies (based on visuals and current simulated metadata), calculate accurately, and ensure the final layout preserves the original design structure while fitting boundaries.

**Goal:** For the **SINGLE** 'Input Instruction' below:
1.  **CRITICAL - Identify & Verify Targets:** Accurately map the NL hint to the **correct Shape IDs** using the `Original Slide Visual Image` AND the `Current Simulated Shape Metadata`. **Verify visual description match between image hint and metadata details.** If mismatched, output an error comment. Identify reference shapes/groups. Use visual context to infer structure.
2.  **Interpret Layout Intent & Analyze Dependencies:** Understand the true layout goal. Based on the inferred structure and visual adjacency (from image), identify potential dependencies: How will modifying the target shape(s) affect adjacent elements?
3.  **Calculate Precise Parameters:** Use the provided formulas and context. Calculate exact coordinates, dimensions, spacing **using values from the `Current Simulated Shape Metadata` as the source of truth for current state**. State the calculation basis. Consider aspect ratio.
4.  **Apply Minimal Change & Constraints:** Modify only essential properties. Respect slide boundaries and margins (~10px safe area: `~10 < left`, `~10 < top`, `left+width < ~914`, `top+height < ~500`). Handle overflow by resizing offenders first.
5.  **Generate Refined Command String(s) for Structural Integrity:** Output ONE precise instruction string (or multiple sequential strings ONLY if structure preservation demands it, e.g., resize + move adjacent). 

**Context:**
Input Instruction (Process ONLY this one):
{{instruction}}

Current Simulated Shape Metadata (JSON - Reflects previous simulated changes. Use THIS for current coordinates/dimensions):
{{current_metadata_state_json}}

Original Slide Visual Image (base64-encoded - Use for initial ID mapping, structure, relationships):
Provided as image input.

**CRITICAL Processing Principles:**
*   **Target Verification:** Mandatory check. Mismatch -> `// Refinement Error: ID mismatch...`
*   **Layout Structure Preservation:** Ensure adjacent elements maintain visual organization/spacing observed in original image, using updated coordinates. May require generating dependent move instructions. Note in justification.
*   **Dependency Awareness:** Consider cause-and-effect.
*   **Minimal Change:** Modify ONLY essential properties + necessary structure preservation adjustments.
*   **Fit to Slide & Margins:** Enforce safe area (~10-914px H, ~10-500px V).
*   **Aspect Ratio:** Maintain generally when fitting.
*   **Calculation Basis:** Justification MUST state how value was derived (e.g., "based on simulated min left", "matching ref id: Z based on current state").

**Detailed Calculation Formulas & Refinement Logic:**
*   **Alignment (Group):** Use min/max/average on `Current Simulated Metadata`. Output absolute target coordinate(s).
*   **Alignment (Pairwise):** Identify pairs using Original Image structure. Read reference value from `Current Simulated Metadata`. Apply to target.
*   **Alignment (Slide):** Calculate based on slide center/middle (480px/270px) and shape size from `Current Simulated Metadata`.
*   **Distribution:** Calculate spacing based on available space (respecting margins) and element dimensions from `Current Simulated Metadata`. Alert if target is complex ('rows'/'columns').
*   **Resizing/Standardization:** Calculate dimension based on average/reference from `Current Simulated Metadata` or available space (`< 914`/`< 500`). Maintain aspect ratio if fitting. Check dependencies using Original Image structure + Current Metadata positions. Generate move instruction for adjacent elements if needed.

**Construct Output String(s):** Format:
`Instruction: Action targeting specific Shape IDs (verified), using calculated parameters (with 'px' units), justified by calculation basis/goal and noting structure preservation.` 
*   Start `"Instruction: "`. Use `(id: ID)`, `(ids: [IDs])`. Use `px` units.
*   Justification examples: ", based on simulated min left value.", ", matching ref shape (id: Z)'s current simulated top.", ", to fit safe width <914px.", ", shifting adjacent column (ids: [...]) left to preserve layout structure." , ", maintaining aspect ratio for image."

**Output Format:**
Return ONLY a JSON object containing the single refined instruction string (or list if multiple steps needed for one NL instruction) or an error/alert comment. Format:
```json
{{{{
    "refined_instruction_output": [
        "Instruction: Set width for shape (id: 123) to 200.0px..."
        // Or potentially multiple if needed:
        // "Instruction: Set width for shape (id: 123) to 200.0px...",
        // "Instruction: Set left for shape (id: 124) to 250.0px..."
        // Or an error/alert comment:
        // "// Refinement Error: ID mismatch..."
    ]
}}}}
```
**CRITICAL Constraints & Fallback:**
*   Mandatory Target Verification. Preserve Inferred Layout Structure. Minimal Change. Fit/Margins. Calculation Basis.
*   If impossible to resolve reliably: Output `// Refinement Error/Alert: [Reason]`.
*   Output ONLY the JSON. Use {{{{ and }}}} for literal braces defining the JSON structure. Use ```json for the output block fence. Use standard {{instruction}} or {{current_metadata_state_json}} for variables to be formatted by Python.
"""

def _parse_refined_instruction(instruction: str) -> Optional[Dict[str, Any]]:
    logger.debug(f"Parsing refined instruction: {instruction}")
    if not instruction or instruction.strip().startswith("//"):
        return None

    patterns = [
        re.compile(r'Set\s+(left|top)\s+coordinate\s+for\s+shapes?\s+\(ids?:\s*\[?([\d\s,]+)\]?\)\s+to\s+([\d\.]+)\s*px', re.IGNORECASE),
        re.compile(r'Set\s+(width|height)\s+for\s+shapes?\s+\(ids?:\s*\[?([\d\s,]+)\]?\)\s+to\s+([\d\.]+)\s*px', re.IGNORECASE),
        re.compile(r'Set\s+(left|top)\s+coordinate\s+for\s+shape\s+\(id:\s*(\d+)\)\s+to\s+match.*?at\s+([\d\.]+)\s*px', re.IGNORECASE),
    ]

    for pattern in patterns:
        match = pattern.search(instruction)
        if match:
            try:
                groups = match.groups()
                if pattern.pattern.startswith(r'Set\s+(left|top)\s+coordinate\s+for\s+shapes?\s+\(ids?'): # Group align / Set coordinate for multiple
                    prop = groups[0].lower()
                    id_str = groups[1].strip()
                    ids = [s_id.strip() for s_id in id_str.split(',') if s_id.strip().isdigit()]
                    value = float(groups[2])
                    logger.debug(f"Parsed Group Align/Set: ids={ids}, property={prop}, value={value}")
                    return {'ids': ids, 'property': prop, 'value': value}
                elif pattern.pattern.startswith(r'Set\s+(width|height)\s+for\s+shapes?\s+\(ids?'): # Set dimension for multiple
                    prop = groups[0].lower()
                    id_str = groups[1].strip()
                    ids = [s_id.strip() for s_id in id_str.split(',') if s_id.strip().isdigit()]
                    value = float(groups[2])
                    logger.debug(f"Parsed Set Dimension: ids={ids}, property={prop}, value={value}")
                    return {'ids': ids, 'property': prop, 'value': value}
                elif pattern.pattern.startswith(r'Set\s+(left|top)\s+coordinate\s+for\s+shape\s+\(id:'): # Pairwise align / Set coordinate for single
                    prop = groups[0].lower()
                    ids = [groups[1].strip()]
                    value = float(groups[2])
                    logger.debug(f"Parsed Pairwise/Single Align: ids={ids}, property={prop}, value={value}")
                    return {'ids': ids, 'property': prop, 'value': value}

            except Exception as e:
                logger.warning(f"Parsing failed for matched pattern '{pattern.pattern}' on instruction '{instruction}': {e}")
                return None 

    logger.warning(f"Could not parse known refined instruction structure: {instruction}")
    return None

async def refiner_agent(slide_number: int, slide_context: Dict[str, Any]) -> Dict[str, Any]:
    global client    
    client = genai.Client(api_key=LLM_API_KEY)
    logger.info(f"--- Starting Iterative Refiner Agent for Slide {slide_number} ---")
    final_refined_instructions = []
    all_errors_or_alerts = []

    try:
        detailed_nl_instructions = _load_nl_subtasks(slide_number)
        original_metadata, simulated_metadata = _load_and_copy_metadata(slide_number)
        slide_image_base64 = slide_context.get("slide_image_base64")

        if not detailed_nl_instructions:
            return {"refined_instructions": [], "message": f"No NL sub-tasks found to refine for slide {slide_number}."}
        if not original_metadata or not simulated_metadata:
            raise FileNotFoundError("Failed to load or copy metadata.")
        if not slide_image_base64:
            logger.warning(f"No slide image context provided for slide {slide_number}. Proceeding without image.")

    except (FileNotFoundError, ValueError) as e:
        return {"error": f"Failed during initial load: {str(e)}"}
    except Exception as e:
        return {"error": f"Unexpected initial load error: {str(e)}"}

    # --- Iterative Refinement ---
    for idx, nl_instruction in enumerate(detailed_nl_instructions):
        iteration_log_prefix = f"Slide {slide_number}, Iteration {idx + 1}/{len(detailed_nl_instructions)}"
        logger.debug(f"{iteration_log_prefix}: Refining NL: \"{nl_instruction}\"")

        try:
            current_metadata_state_json = json.dumps(simulated_metadata, indent=2)
        except TypeError as json_err:
            all_errors_or_alerts.append(f"Iter {idx+1}: Metadata serialization error.")
            continue

        try:
            prompt = REFINER_PROMPT_TEMPLATE.format(instruction=nl_instruction, current_metadata_state_json=current_metadata_state_json)
            contents = [prompt]
            if slide_image_base64:
                contents.append({"inline_data": {"mime_type": "image/png", "data": slide_image_base64}})

            response = await asyncio.to_thread(client.models.generate_content, model="gemini-2.0-flash", contents=contents)
            raw_response_text = response.text.strip() if hasattr(response, 'text') else ""

            if not raw_response_text:
                logger.warning(f"{iteration_log_prefix}: Empty LLM response.")
                all_errors_or_alerts.append(f"Iter {idx+1}: Empty LLM response.")
                continue

            parsed_json_output = None
            json_match = re.search(r'```json\s*(\{[\s\S]*?\})\s*```', raw_response_text, re.DOTALL)
            if not json_match:
                json_match = re.search(r'(\{[\s\S]*?\})', raw_response_text, re.DOTALL)

            if json_match:
                try:
                    parsed_json_output = json.loads(json_match.group(1))
                except json.JSONDecodeError as json_err:
                    all_errors_or_alerts.append(f"Iter {idx+1}: JSON decode error.")
                    continue
            else:
                if raw_response_text.strip().startswith("//"):
                    all_errors_or_alerts.append(f"Iter {idx+1}: LLM Comment.")
                    final_refined_instructions.append(raw_response_text.strip())
                    continue
                else:
                    all_errors_or_alerts.append(f"Iter {idx+1}: No JSON found.")
                    continue

            # --- Validate and Update ---
            if parsed_json_output and isinstance(parsed_json_output, dict) and "refined_instruction_output" in parsed_json_output:
                refined_output = parsed_json_output["refined_instruction_output"]
                if isinstance(refined_output, list) and refined_output:
                    simulation_success_for_this_step = True
                    for refined_inst_str in refined_output:
                        if isinstance(refined_inst_str, str):
                            refined_inst_str = refined_inst_str.strip()
                            if refined_inst_str.startswith("//"):
                                logger.warning(f"{iteration_log_prefix}: Refiner Alert/Error: {refined_inst_str}")
                                all_errors_or_alerts.append(f"Iter {idx+1}: {refined_inst_str}")
                                if "Error:" in refined_inst_str:
                                    simulation_success_for_this_step = False
                                final_refined_instructions.append(refined_inst_str)
                            elif refined_inst_str:
                                logger.info(f"{iteration_log_prefix}: Successfully refined: \"{refined_inst_str}\"")
                                final_refined_instructions.append(refined_inst_str)

                                if simulation_success_for_this_step:
                                    changes_to_apply = _parse_refined_instruction(refined_inst_str)
                                    if changes_to_apply and not _update_simulated_metadata(simulated_metadata, changes_to_apply):
                                        logger.error(f"{iteration_log_prefix}: Failed to simulate state update.")
                                        all_errors_or_alerts.append(f"Iter {idx+1}: State update failed.")
                                        simulation_success_for_this_step = False
                                    else:
                                        logger.warning(f"{iteration_log_prefix}: Refined instruction parsing failed.")
                                        all_errors_or_alerts.append(f"Iter {idx+1}: Refined instruction parsing failed.")
                                        simulation_success_for_this_step = False
                            else:
                                logger.warning(f"{iteration_log_prefix}: Encountered empty instruction.")
                else:
                    all_errors_or_alerts.append(f"Iter {idx+1}: Empty refined output list.")
            else:
                all_errors_or_alerts.append(f"Iter {idx+1}: Invalid LLM JSON structure.")
                continue

        except google.api_core.exceptions.ResourceExhausted as e:
            logger.error(f"{iteration_log_prefix}: Google API Quota Exhausted: {e}")
            all_errors_or_alerts.append(f"Iter {idx+1}: API Quota Error.")
            break
        except google.api_core.exceptions.GoogleAPIError as e:
            logger.error(f"{iteration_log_prefix}: Google API Error: {e}")
            all_errors_or_alerts.append(f"Iter {idx+1}: API Error.")
            break
        except Exception as e:
            logger.error(f"{iteration_log_prefix}: Unexpected error: {e}", exc_info=True)
            all_errors_or_alerts.append(f"Iter {idx+1}: Unexpected error.")
            break

    # --- Save Final Instructions ---
    output_filename = f"slide{slide_number}_refined_tasks.json"
    output_path = os.path.join(REFINER_OUTPUT_DIR, output_filename)
    try:
        os.makedirs(REFINER_OUTPUT_DIR, exist_ok=True)
        output_data = {"refined_instructions": final_refined_instructions}
        async with aiofiles.open(output_path, "w", encoding="utf-8") as f:
            await f.write(json.dumps(output_data, indent=2))
        logger.info(f"Saved final refined instructions for slide {slide_number} to: {output_path}")
    except Exception as save_err:
        all_errors_or_alerts.append(f"Failed to save output file: {save_err}")

    logger.info(f"--- Finished Refiner Agent for Slide {slide_number}. Errors/Alerts encountered: {len(all_errors_or_alerts)} ---")
    if all_errors_or_alerts:
        return {
            "status": "partial_success" if final_refined_instructions else "failure",
            "error": "Some refinement steps failed or generated alerts.",
            "details": all_errors_or_alerts,
            "refined_instructions": final_refined_instructions
        }
    else:
        return {"status": "success", "refined_instructions": final_refined_instructions}
    


# # refiner_agent.py
# import logging, json, re, base64, math, asyncio, os
# from typing import Dict, Any, List, Optional, Set

# import aiofiles
# from config.config import LLM_API_KEY
# from google import genai

# client = genai.Client(api_key=LLM_API_KEY)

# logger = logging.getLogger(__name__)
# INSTRUCTIONS_PER_CHUNK = 4 
# REFINER_OUTPUT_DIR = "uploaded_pptx/slide_images/presentation" 
# METADATA_DIR = "uploaded_pptx/slide_images/metadata"
# TASKS_DIR = "uploaded_pptx/slide_images/presentation"

# def _load_tasks_from_file(slide_number: int):
#     file_path = os.path.join(TASKS_DIR, f"slide{slide_number}_tasks.json")
#     if os.path.exists(file_path):
#         with open(file_path, "r", encoding="utf-8") as f:
#             tasks_data = json.load(f)
#             if isinstance(tasks_data, list):
#                 instruction_list = [
#                     task.get("task_description", str(task)) for task in tasks_data
#                     if task.get("task_description")
#                 ]
#                 if not instruction_list:
#                      logger.warning(f"Task file {file_path} contained list but no 'task_description' fields.")
#                 return instruction_list
#             else:
#                 raise ValueError(f"Invalid format in task file for slide {slide_number}")
#     else:
#         raise FileNotFoundError(f"Task file not found for slide {slide_number}")

# def _load_shape_metadata_for_slide(slide_number: int):
#     metadata_path = f"uploaded_pptx/slide_images/metadata/metadata_{slide_number}.json"
#     if not os.path.exists(metadata_path):
#         raise FileNotFoundError(f"Shape metadata file not found for slide {slide_number}")
#     with open(metadata_path, "r", encoding="utf-8") as f:
#         shapes = json.load(f)
#     return shapes

# async def refiner_agent(slide_number: int, slide_context: Dict[str, Any]) -> Dict[str, Any]:
#     global client
#     logger.info(f"--- Starting Refiner Agent for Slide {slide_number} ---")
#     final_refined_instructions = []
#     all_errors = []

#     try:
#         # Load required inputs for this slide
#         detailed_nl_instructions = _load_tasks_from_file(slide_number)
#         full_slide_metadata = _load_shape_metadata_for_slide(slide_number)
#         slide_image_base64 = slide_context.get("slide_image_base64")

#         if not detailed_nl_instructions:
#             return {"refined_instructions": [], "message": "No instructions to refine."}
#         if not slide_image_base64:
#             logger.warning(f"No slide image context provided for slide {slide_number}. Proceeding without image.")

#         try:
#             shapes_json_full = json.dumps(full_slide_metadata)
#         except TypeError as json_err:
#              return {"error": f"Metadata serialization error: {json_err}"}

#         # Calculate chunks
#         num_chunks = math.ceil(len(detailed_nl_instructions) / INSTRUCTIONS_PER_CHUNK)
#         logger.info(f"Slide {slide_number}: Processing {len(detailed_nl_instructions)} instructions in {num_chunks} chunk(s)...")

#         # Loop through chunks
#         for i in range(num_chunks):
#             start_idx = i * INSTRUCTIONS_PER_CHUNK
#             end_idx = start_idx + INSTRUCTIONS_PER_CHUNK
#             instruction_chunk = detailed_nl_instructions[start_idx:end_idx]
#             chunk_log_prefix = f"Slide {slide_number}, Chunk {i + 1}/{num_chunks}"
#             logger.debug(f"{chunk_log_prefix}: Processing instructions {start_idx + 1}-{min(end_idx, len(detailed_nl_instructions))}")
            
#             try:
#                 tasks_json_chunk = json.dumps(instruction_chunk)

#                 CHUNKED_REFINER_PROMPT = f"""
#                 You are an AI assistant refining a *chunk* of PowerPoint modification instructions into explicit, executable command formats.

#                 **Goal:** Convert the 'Input Instructions Chunk' below into precise, actionable command strings using the provided context. Reference the Full Shape Metadata for details about shapes mentioned.

#                 **Input Instructions Chunk:**
#                 {tasks_json_chunk}

#                 **Context:**
#                 Full Slide Shape Metadata (JSON): An array of shape objects, each containing `id`, `type`, `top`, `left`, `width`, `height`, `text`, `font`, etc. {shapes_json_full}
#                 Full Slide Visual Image(base64-encoded, included below as image reference): Provided as image input.

#                 **Analysis & Output Instructions:**
#                 - Analyze EACH instruction in the input chunk.
#                 - Use the Full Slide Shape Metadata and image to resolve ambiguity (e.g., determine exact coordinates, sizes, fonts, alignment targets based on IDs or hints in the instructions).
#                 - Generate a refined instruction string for EACH input instruction in the chunk.
#                 - The refined instruction MUST include specific shape IDs `(id: 123)` or `(ids: [123, 456])` derived from the original instruction or metadata lookup.
#                 - The refined instruction MUST include exact parameters (e.g., `top at 150px`, `width to 200px`, `size 12pt`, `font 'Arial'`).
#                 - If an input instruction is already sufficiently precise, output it as is.

#                 **Output Format:**
#                 Return ONLY a JSON object containing a list of the refined instruction strings generated *for this chunk*. Format:
#                 ```json
#                 {{
#                 "refined_instructions_chunk": [
#                     "Refined instruction 1 for this chunk...",
#                     "Refined instruction 2 for this chunk...",
#                     ...
#                 ]
#                 }}
#                 ```
#                 DO NOT include explanations, only the JSON object.
#                 """
#                 contents = [CHUNKED_REFINER_PROMPT]
#                 if slide_image_base64:
#                     try:
#                          image_part = {"inline_data": {"mime_type":"image/png", "data":slide_image_base64}}
#                          contents.append(image_part) 
#                          logger.debug(f"{chunk_log_prefix}: Included image in prompt.")
#                     except Exception as img_err:
#                          logger.error(f"{chunk_log_prefix}: Failed to prepare image for LLM: {img_err}")
#                 else:
#                     logger.warning(f"{chunk_log_prefix}: No image provided, proceeding without visual context.")

#                 response = client.models.generate_content(
#                     model="gemini-2.0-flash", 
#                     contents=contents
#                 )
#                 raw_response_text = response.text.strip()
#                 logging.info(f"{chunk_log_prefix}: LLM Raw Response: {raw_response_text}...")

#                 # Extract JSON
#                 parsed_json_chunk = None
#                 json_match = re.search(r'\{.*\}', raw_response_text, re.DOTALL)
#                 if json_match:
#                     try:
#                         parsed_json_chunk = json.loads(json_match.group(0))
#                     except json.JSONDecodeError as json_err:
#                         logger.error(f"{chunk_log_prefix}: Failed to decode JSON: {json_err}. Raw: {raw_response_text}")
#                         all_errors.append(f"Chunk {i+1} JSON decode error")
#                         continue # Skip this chunk
#                 else:
#                     logger.error(f"{chunk_log_prefix}: No JSON found in response: {raw_response_text}")
#                     all_errors.append(f"Chunk {i+1} No JSON found")
#                     continue

#                 # Validate and aggregate
#                 if "refined_instructions_chunk" in parsed_json_chunk and isinstance(parsed_json_chunk["refined_instructions_chunk"], list):
#                     refined_chunk = parsed_json_chunk["refined_instructions_chunk"]
#                     final_refined_instructions.extend(refined_chunk)
#                     logger.info(f"{chunk_log_prefix}: Successfully refined {len(refined_chunk)} instructions.")
#                 else:
#                     logger.error(f"{chunk_log_prefix}: Invalid JSON structure: {parsed_json_chunk}")
#                     all_errors.append(f"Chunk {i+1} Invalid JSON structure")
#                     continue

#             except Exception as chunk_err:
#                 logger.error(f"{chunk_log_prefix}: Error processing chunk: {chunk_err}", exc_info=True)
#                 all_errors.append(f"Chunk {i+1} processing error: {chunk_err}")

#         # Save aggregated refined instructions to new file
#         output_filename = f"slide{slide_number}_refined_tasks.json"
#         output_path = os.path.join(REFINER_OUTPUT_DIR, output_filename)
#         try:
#             os.makedirs(REFINER_OUTPUT_DIR, exist_ok=True) 
#             output_data = {"refined_instructions": final_refined_instructions}
#             async with aiofiles.open(output_path, "w", encoding="utf-8") as f:
#                 await f.write(json.dumps(output_data, indent=2))
#             logger.info(f"Saved aggregated refined instructions for slide {slide_number} to: {output_path}")
#         except Exception as save_err:
#              logger.error(f"Failed to save refined tasks JSON for slide {slide_number}: {save_err}", exc_info=True)
#              all_errors.append(f"Failed to save output file: {save_err}") 

#         # Return result
#         logger.info(f"--- Finished Refiner Agent for Slide {slide_number}. Total refined: {len(final_refined_instructions)}. Errors: {len(all_errors)} ---")
#         if all_errors:
#             return {
#                 "error": "Some refinement steps failed.",
#                 "details": all_errors,
#                 "refined_instructions": final_refined_instructions 
#             }
#         else:
#             return {"refined_instructions": final_refined_instructions}

#     except FileNotFoundError as e:
#         logger.error(f"Refiner Agent failed for slide {slide_number}: Required input file not found - {e}")
#         return {"error": str(e)}
#     except ValueError as e: 
#          logger.error(f"Refiner Agent failed for slide {slide_number}: Invalid format in input file - {e}")
#          return {"error": str(e)}
#     except Exception as e:
#         logger.error(f"Unexpected error during instruction refinement for slide {slide_number}: {e}", exc_info=True)
#         return {"error": f"Unexpected refinement error: {str(e)}"}