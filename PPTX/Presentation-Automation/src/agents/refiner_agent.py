
# refiner_agent.py
import logging, json, re, base64, math, asyncio, os
from typing import Dict, Any, List, Optional, Set

import aiofiles
from config.config import LLM_API_KEY
from google import genai

client = genai.Client(api_key=LLM_API_KEY)

logger = logging.getLogger(__name__)
INSTRUCTIONS_PER_CHUNK = 4 
REFINER_OUTPUT_DIR = "uploaded_pptx/slide_images/presentation" 
METADATA_DIR = "uploaded_pptx/slide_images/metadata"
TASKS_DIR = "uploaded_pptx/slide_images/presentation"

def _load_tasks_from_file(slide_number: int):
    file_path = os.path.join(TASKS_DIR, f"slide{slide_number}_tasks.json")
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            tasks_data = json.load(f)
            if isinstance(tasks_data, list):
                instruction_list = [
                    task.get("task_description", str(task)) for task in tasks_data
                    if task.get("task_description")
                ]
                if not instruction_list:
                     logger.warning(f"Task file {file_path} contained list but no 'task_description' fields.")
                return instruction_list
            else:
                raise ValueError(f"Invalid format in task file for slide {slide_number}")
    else:
        raise FileNotFoundError(f"Task file not found for slide {slide_number}")

def _load_shape_metadata_for_slide(slide_number: int):
    metadata_path = f"uploaded_pptx/slide_images/metadata/metadata_{slide_number}.json"
    if not os.path.exists(metadata_path):
        raise FileNotFoundError(f"Shape metadata file not found for slide {slide_number}")
    with open(metadata_path, "r", encoding="utf-8") as f:
        shapes = json.load(f)
    return shapes

async def refiner_agent(slide_number: int, slide_context: Dict[str, Any]) -> Dict[str, Any]:
    global client
    logger.info(f"--- Starting Refiner Agent for Slide {slide_number} ---")
    final_refined_instructions = []
    all_errors = []

    try:
        # Load required inputs for this slide
        detailed_nl_instructions = _load_tasks_from_file(slide_number)
        full_slide_metadata = _load_shape_metadata_for_slide(slide_number)
        slide_image_base64 = slide_context.get("slide_image_base64")

        if not detailed_nl_instructions:
            return {"refined_instructions": [], "message": "No instructions to refine."}
        if not slide_image_base64:
            logger.warning(f"No slide image context provided for slide {slide_number}. Proceeding without image.")

        try:
            shapes_json_full = json.dumps(full_slide_metadata)
        except TypeError as json_err:
             return {"error": f"Metadata serialization error: {json_err}"}

        # Calculate chunks
        num_chunks = math.ceil(len(detailed_nl_instructions) / INSTRUCTIONS_PER_CHUNK)
        logger.info(f"Slide {slide_number}: Processing {len(detailed_nl_instructions)} instructions in {num_chunks} chunk(s)...")

        # Loop through chunks
        for i in range(num_chunks):
            start_idx = i * INSTRUCTIONS_PER_CHUNK
            end_idx = start_idx + INSTRUCTIONS_PER_CHUNK
            instruction_chunk = detailed_nl_instructions[start_idx:end_idx]
            chunk_log_prefix = f"Slide {slide_number}, Chunk {i + 1}/{num_chunks}"
            logger.debug(f"{chunk_log_prefix}: Processing instructions {start_idx + 1}-{min(end_idx, len(detailed_nl_instructions))}")
            
            try:
                tasks_json_chunk = json.dumps(instruction_chunk)

                CHUNKED_REFINER_PROMPT = f"""
                You are an AI assistant refining a *chunk* of PowerPoint modification instructions into explicit, executable command formats.

                **Goal:** Convert the 'Input Instructions Chunk' below into precise, actionable command strings using the provided context. Reference the Full Shape Metadata for details about shapes mentioned.

                **Input Instructions Chunk:**
                {tasks_json_chunk}

                **Context:**
                Full Slide Shape Metadata (JSON): An array of shape objects, each containing `id`, `type`, `top`, `left`, `width`, `height`, `text`, `font`, etc. {shapes_json_full}
                Full Slide Visual Image(base64-encoded, included below as image reference): Provided as image input.

                **Analysis & Output Instructions:**
                - Analyze EACH instruction in the input chunk.
                - Use the Full Slide Shape Metadata and image to resolve ambiguity (e.g., determine exact coordinates, sizes, fonts, alignment targets based on IDs or hints in the instructions).
                - Generate a refined instruction string for EACH input instruction in the chunk.
                - The refined instruction MUST include specific shape IDs `(id: 123)` or `(ids: [123, 456])` derived from the original instruction or metadata lookup.
                - The refined instruction MUST include exact parameters (e.g., `top at 150px`, `width to 200px`, `size 12pt`, `font 'Arial'`).
                - If an input instruction is already sufficiently precise, output it as is.

                **Output Format:**
                Return ONLY a JSON object containing a list of the refined instruction strings generated *for this chunk*. Format:
                ```json
                {{
                "refined_instructions_chunk": [
                    "Refined instruction 1 for this chunk...",
                    "Refined instruction 2 for this chunk...",
                    ...
                ]
                }}
                ```
                DO NOT include explanations, only the JSON object.
                """
                contents = [CHUNKED_REFINER_PROMPT]
                if slide_image_base64:
                    try:
                         image_part = {"inline_data": {"mime_type":"image/png", "data":slide_image_base64}}
                         contents.append(image_part) 
                         logger.debug(f"{chunk_log_prefix}: Included image in prompt.")
                    except Exception as img_err:
                         logger.error(f"{chunk_log_prefix}: Failed to prepare image for LLM: {img_err}")
                else:
                    logger.warning(f"{chunk_log_prefix}: No image provided, proceeding without visual context.")

                response = client.models.generate_content(
                    model="gemini-2.0-flash", 
                    contents=contents
                )
                raw_response_text = response.text.strip()
                logging.info(f"{chunk_log_prefix}: LLM Raw Response: {raw_response_text}...")

                # Extract JSON
                parsed_json_chunk = None
                json_match = re.search(r'\{.*\}', raw_response_text, re.DOTALL)
                if json_match:
                    try:
                        parsed_json_chunk = json.loads(json_match.group(0))
                    except json.JSONDecodeError as json_err:
                        logger.error(f"{chunk_log_prefix}: Failed to decode JSON: {json_err}. Raw: {raw_response_text}")
                        all_errors.append(f"Chunk {i+1} JSON decode error")
                        continue # Skip this chunk
                else:
                    logger.error(f"{chunk_log_prefix}: No JSON found in response: {raw_response_text}")
                    all_errors.append(f"Chunk {i+1} No JSON found")
                    continue

                # Validate and aggregate
                if "refined_instructions_chunk" in parsed_json_chunk and isinstance(parsed_json_chunk["refined_instructions_chunk"], list):
                    refined_chunk = parsed_json_chunk["refined_instructions_chunk"]
                    final_refined_instructions.extend(refined_chunk)
                    logger.info(f"{chunk_log_prefix}: Successfully refined {len(refined_chunk)} instructions.")
                else:
                    logger.error(f"{chunk_log_prefix}: Invalid JSON structure: {parsed_json_chunk}")
                    all_errors.append(f"Chunk {i+1} Invalid JSON structure")
                    continue

            except Exception as chunk_err:
                logger.error(f"{chunk_log_prefix}: Error processing chunk: {chunk_err}", exc_info=True)
                all_errors.append(f"Chunk {i+1} processing error: {chunk_err}")

        # Save aggregated refined instructions to new file
        output_filename = f"slide{slide_number}_refined_tasks.json"
        output_path = os.path.join(REFINER_OUTPUT_DIR, output_filename)
        try:
            os.makedirs(REFINER_OUTPUT_DIR, exist_ok=True) 
            output_data = {"refined_instructions": final_refined_instructions}
            async with aiofiles.open(output_path, "w", encoding="utf-8") as f:
                await f.write(json.dumps(output_data, indent=2))
            logger.info(f"Saved aggregated refined instructions for slide {slide_number} to: {output_path}")
        except Exception as save_err:
             logger.error(f"Failed to save refined tasks JSON for slide {slide_number}: {save_err}", exc_info=True)
             all_errors.append(f"Failed to save output file: {save_err}") 

        # Return result
        logger.info(f"--- Finished Refiner Agent for Slide {slide_number}. Total refined: {len(final_refined_instructions)}. Errors: {len(all_errors)} ---")
        if all_errors:
            return {
                "error": "Some refinement steps failed.",
                "details": all_errors,
                "refined_instructions": final_refined_instructions 
            }
        else:
            return {"refined_instructions": final_refined_instructions}

    except FileNotFoundError as e:
        logger.error(f"Refiner Agent failed for slide {slide_number}: Required input file not found - {e}")
        return {"error": str(e)}
    except ValueError as e: 
         logger.error(f"Refiner Agent failed for slide {slide_number}: Invalid format in input file - {e}")
         return {"error": str(e)}
    except Exception as e:
        logger.error(f"Unexpected error during instruction refinement for slide {slide_number}: {e}", exc_info=True)
        return {"error": f"Unexpected refinement error: {str(e)}"}