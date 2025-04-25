# main.py
import asyncio, json, logging, os, uvicorn
from typing import Dict, List
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# === Router Imports ===
from routes.metadata_handler import router as metadata_router
from routes.pptx_handler import router as pptx_router
from feedback_parsing.feedback_classifier import classify_feedback_instructions, parse_feedback_instruction

# === Agent & Context Imports ===
from agents.cleanup_agent import cleanup_agent
from agents.formatting_agent import formatting_agent
from agents.refiner_agent import refiner_agent
from agents.code_generation_agent import generate_code
from agents.visual_enhancement_agent import visual_enhancement_agent

from utils.load_files import load_slide_contexts

# --- Logging Configuration ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- FastAPI App Initialization ---
app = FastAPI(title="Slide Enhancement API", version="1.0.0")

slide_context_cache: dict[int, dict] = {}

# --- CORS Middleware ---
origins = ["*"]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- Pydantic Models for Request Bodies ---
class InstructionRequest(BaseModel):
    instruction: str
    slide_index: int
    total_slides: int

# === Route Registration ===
try:
    app.include_router(metadata_router, prefix="/upload-metadata", tags=["Metadata"])
    app.include_router(pptx_router, prefix="/upload-pptx", tags=["PPTX Upload"])
    logger.info("All routers included successfully.")
except NameError:
    logger.error("Failed to include routers. Check if router objects are defined correctly.")

@app.post("/process_instruction")
async def process_instruction(request: InstructionRequest):
    logger.info(f"Received instruction: {request.instruction}")
    logger.info(f"Target slide index: {request.slide_index}")
    logger.info(f"Total slide: {request.total_slides}")

    slide_images_dir = "uploaded_pptx/slide_images/presentation"
    slide_number = request.slide_index

    slide_image_path = os.path.join(slide_images_dir, f"slide{slide_number}_image.txt")
    slide_xml_path = os.path.join(slide_images_dir, f"slide{slide_number}.xml")
    slide_png_path = os.path.join(slide_images_dir, f"slide{slide_number}.png")

    if not (os.path.exists(slide_image_path) and os.path.exists(slide_xml_path)):
        logger.error(f"Slide context files for slide {slide_number} do not exist!")
        return {"status": "error", "message": f"Slide context for slide {slide_number} not found."}

    # --- Step 1: Classify Instruction ---
    feedback_item = {
        "instruction": request.instruction,
        "slide_number": slide_number,
        "total_slides": request.total_slides,
        "source": "user_input"
    }
    categorized_tasks = classify_feedback_instructions([feedback_item])
    logger.info(f"Categorized Tasks: {len(categorized_tasks)}")

    if not categorized_tasks:
        return {"status": "no_tasks", "message": "No actionable feedback classified."}

    task = categorized_tasks[0]
    category = task["category"]
    target_scope = task.get("instruction_scope", "current_slide")
    target_slides = task.get("target_slide_indices", [slide_number])

    logger.info(f"Task Category: {category}")
    logger.info(f"Scope: {target_scope}, Target Slides: {target_slides}")

    # --- Step 2: Load Context for Target Slides ---
    uncached = [i for i in target_slides if i not in slide_context_cache]
    if uncached:
        logger.info(f"Loading context for uncached slides: {uncached}")
        loaded = await load_slide_contexts(uncached)
        slide_context_cache.update(loaded)

    context_loaded = {i: slide_context_cache[i] for i in target_slides if i in slide_context_cache}
    logger.info(f"Context loaded for {len(context_loaded)} slides.")

    # --- Step 2: Run Agents ---
    logger.info("Processing Tasks with Agents & Context...")
    all_task_specifications = []

    for slide_id in target_slides:
        slide_context = slide_context_cache.get(slide_id)
        if not slide_context:
            logger.warning(f"Missing context for slide {slide_id}")
            continue

        for task in categorized_tasks:
            category = task["category"]
            task["slide_number"] = slide_id

            try:
                if category == "formatting":
                    result = formatting_agent(task, slide_context)
                elif category == "cleanup":
                    result = cleanup_agent(task, slide_context)
                elif category == "visual_enhancement":
                    result = visual_enhancement_agent(task, slide_context)
                else:
                    logger.warning(f"Unknown category: {category}")
                    continue

                if result:
                    # Tag each task with slide number for traceability
                    for r in result:
                        r["slide_number"] = slide_id
                        r["agent_name"] = category
                        r["original_instruction"] = task["original_instruction"]

                    # Save individual result for this slide
                    slide_path = os.path.join(slide_images_dir, f"slide{slide_id}_tasks.json")
                    with open(slide_path, "w", encoding="utf-8") as f:
                        json.dump(result, f, indent=4)

                    # logger.info(f"Saved processed subtasks to: {slide_path}")
                    all_task_specifications.extend(result)

            except Exception as e:
                logger.error(f"Error processing slide {slide_id}: {e}")

    #  --- Save task_specifications as JSON ---
    try:
        json_output_path = os.path.join(slide_images_dir, f"slide{slide_number}_tasks.json")
        with open(json_output_path, "w", encoding="utf-8") as f:
            json.dump(all_task_specifications, f, indent=4)
        logger.info(f"Saved processed subtasks to: {json_output_path}")
    except Exception as e:
        logger.error(f"Failed to save subtasks JSON: {e}")

    # --- Step 3: Refiner Agent (Run slides in parallel) ---
    logger.info("Refining Instructions:")
    refiner_tasks = []
    for slide_id in target_slides:
        slide_context = context_loaded.get(slide_id)

        if slide_context: 
            refiner_tasks.append(
                refiner_agent(slide_number=slide_id, slide_context=slide_context)
            )
        else:
            logger.warning(f"Skipping refinement task creation for slide {slide_id} due to missing context.")

    all_refiner_results = await asyncio.gather(*refiner_tasks, return_exceptions=True)

    # Process refinement results
    all_refined_instructions_dict: Dict[int, List[str]] = {}
    any_refiner_errors = False
    original_indices_for_results = [sid for sid in target_slides if sid in context_loaded] 

    for i, result in enumerate(all_refiner_results):
        slide_id = original_indices_for_results[i] 
        if isinstance(result, Exception):
             logger.error(f"Refiner task for slide {slide_id} failed: {result}")
             any_refiner_errors = True
        elif isinstance(result, dict):
            if "error" in result:
                 logger.error(f"Refiner agent reported error for slide {slide_id}: {result.get('details', result['error'])}")
                 any_refiner_errors = True
                 if "refined_instructions" in result:
                      all_refined_instructions_dict[slide_id] = result["refined_instructions"]
            elif "refined_instructions" in result:
                 all_refined_instructions_dict[slide_id] = result["refined_instructions"]
                 logger.info(f"Refined {len(result['refined_instructions'])} instructions for slide {slide_id}.")
            else:
                 logger.error(f"Refiner output for slide {slide_id} has unexpected structure: {result}")
                 any_refiner_errors = True
        else:
            logger.error(f"Refiner task for slide {slide_id} returned unexpected type: {type(result)}")
            any_refiner_errors = True





    # --- Step 5: Generate Code (Per Slide - Parallel) ---
    logger.info("Generating Office.js code per target slide...")
    generated_code_by_slide: Dict[int, str] = {}
    code_gen_success = True
    code_gen_tasks = []

    # Create tasks for concurrent execution
    for slide_id in target_slides:
        # Check if refinement succeeded for this slide before generating code
        if slide_id in all_refined_instructions_dict and all_refined_instructions_dict[slide_id]:
             code_gen_tasks.append(
                 generate_code(target_slide_index=slide_id) # Call async function
             )
        else:
            logger.warning(f"Skipping code generation for slide {slide_id} due to missing/empty refined instructions.")

    # Run code generation concurrently
    all_codegen_results = []
    if code_gen_tasks:
         logger.info(f"Awaiting {len(code_gen_tasks)} code generation calls...")
         all_codegen_results = await asyncio.gather(*code_gen_tasks, return_exceptions=True)
         logger.info("Code generation calls completed.")
    else:
         logger.info("No code generation tasks to run.")

    # Process code gen results
    processed_indices = [sid for sid in target_slides if sid in all_refined_instructions_dict and all_refined_instructions_dict[sid]] # Indices for which tasks were created

    for i, result in enumerate(all_codegen_results):
         slide_id = processed_indices[i] # Get corresponding slide ID
         if isinstance(result, Exception):
              logger.error(f"Code generation task for slide {slide_id} raised exception: {result}")
              code_gen_success = False
         elif isinstance(result, dict):
              if "error" in result:
                   logger.error(f"Code generation for slide {slide_id} failed: {result['error']}")
                   code_gen_success = False
              elif "code" in result:
                   generated_code_by_slide[slide_id] = result["code"]
                   logger.info(f"Code successfully generated for slide {slide_id}.")
              else:
                   logger.error(f"Code generation for slide {slide_id} returned unexpected dict: {result}")
                   code_gen_success = False
         else:
             logger.error(f"Code generation task for slide {slide_id} returned unexpected type: {type(result)}")
             code_gen_success = False




    # === Step 5: Return Final Response ===
    logger.info("Process instruction endpoint finished.")
    return {
        "status": "success",
        "message": "Process completed (Placeholders used).",
        "refined_instructions_by_slide": all_refined_instructions_dict,
        "generated_code": generated_code_by_slide
    }    


@app.get("/get-slide-context")
async def get_slide_context(
    target_slides: List[int] = Query(..., description="Target slide indices")
):
    # Determine which slides need to be fetched
    uncached = [i for i in target_slides if i not in slide_context_cache]
    if uncached:
        loaded = await load_slide_contexts(uncached)
        slide_context_cache.update(loaded)
    # Return context from cache
    result = {i: slide_context_cache[i] for i in target_slides if i in slide_context_cache}
    return {
        "message": "Context loaded",
        "count": len(result),
        "data": result
    }

# --- Main Execution ---
if __name__ == "__main__":
    logger.info("Starting Uvicorn server for development...")
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)