# import logging
# import uvicorn
# from fastapi import FastAPI
# from fastapi.middleware.cors import CORSMiddleware
# from pydantic import BaseModel

# # === Router Imports ===
# from routes.metadata_handler import router as metadata_router
# from routes.pptx_handler import router as pptx_router
# from routes.slide_info_handler import router as slide_info_router 

# # --- Logging Configuration ---
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# logger = logging.getLogger(__name__)

# # --- FastAPI App Initialization ---
# app = FastAPI(title="Slide Enhancement API", version="1.0.0")

# # --- CORS Middleware ---
# origins = ["*"]
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=origins,
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # --- Pydantic Models for Request Bodies ---
# class InstructionRequest(BaseModel):
#     instruction: str

# # === Route Registration ===
# try:
#     app.include_router(metadata_router, prefix="/upload-metadata", tags=["Metadata"])
#     app.include_router(pptx_router, prefix="/upload-pptx", tags=["PPTX Upload"])
#     app.include_router(slide_info_router, prefix="/upload-slide-info", tags=["Slide Info"])
#     logger.info("All routers included successfully.")
# except NameError:
#     logger.error("Failed to include routers. Check if router objects are defined correctly.")

# @app.post("/process_instruction")
# async def process_instruction(request: InstructionRequest):
#     print(f"Received instruction: {request.instruction}")
#     return {"status": "success", "instruction": request.instruction}

# # --- Main Execution ---
# if __name__ == "__main__":
#     logger.info("Starting Uvicorn server for development...")
#     uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)


import json
import logging
import os
import uvicorn
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# === Router Imports ===
from routes.metadata_handler import router as metadata_router
from routes.pptx_handler import router as pptx_router
from feedback_parsing.feedback_classifier import classify_feedback_instructions

# === Agent & Context Imports ===
from agents.cleanup_agent import cleanup_agent
from agents.formatting_agent import formatting_agent
from agents.refiner_agent import refiner_agent
from agents.code_generation_agent import generate_code
from agents.visual_enhancement_agent import visual_enhancement_agent

# --- Logging Configuration ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- FastAPI App Initialization ---
app = FastAPI(title="Slide Enhancement API", version="1.0.0")

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

# === Route Registration ===
try:
    app.include_router(metadata_router, prefix="/upload-metadata", tags=["Metadata"])
    app.include_router(pptx_router, prefix="/upload-pptx", tags=["PPTX Upload"])
    logger.info("All routers included successfully.")
except NameError:
    logger.error("Failed to include routers. Check if router objects are defined correctly.")

# # --- Instruction Endpoint ---
# @app.post("/process_instruction")
# async def process_instruction(request: InstructionRequest):
#     logger.info(f"Received instruction: {request.instruction}")
#     logger.info(f"Target slide index: {request.slide_index}")
    
#     # --- Slide Context Directory ---
#     slide_images_dir = "uploaded_pptx/slide_images/presentation"

#     slide_number = request.slide_index  # Human-friendly (1-based)
    
#     # Flat file paths, not subfolders
#     slide_image_path = os.path.join(slide_images_dir, f"slide{slide_number}_image.txt")
#     slide_xml_path = os.path.join(slide_images_dir, f"slide{slide_number}.xml")

#     # Check if required files exist
#     if not (os.path.exists(slide_image_path) and os.path.exists(slide_xml_path)):
#         logger.error(f"Slide context files for slide {slide_number} do not exist!")
#         return {"status": "error", "message": f"Slide context for slide {slide_number} not found."}

#     # Load context
#     slide_context = {}
#     try:
#         with open(slide_image_path, "r", encoding="utf-8") as img_file:
#             slide_context["slide_image_base64"] = img_file.read()

#         with open(slide_xml_path, "r", encoding="utf-8") as xml_file:
#             slide_context["slide_xml_structure"] = xml_file.read()

#         slide_png_path = os.path.join(slide_images_dir, f"slide{slide_number}.png")
#         if os.path.exists(slide_png_path):
#             with open(slide_png_path, "rb") as img_bin_file:
#                 slide_context["slide_image_bytes"] = img_bin_file.read()
#         else:
#             logger.warning(f"Slide PNG file not found at: {slide_png_path}")
#             slide_context["slide_image_bytes"] = b""

#         logger.info(f"Slide context loaded for slide {slide_number}")
#     except Exception as e:
#         logger.error(f"Error loading context for slide {slide_number}: {e}")
#         return {"status": "error", "message": f"Error loading context for slide {slide_number}"}

#     # Step 1: Instruction Interpretation
#     feedback_item = {
#         "instruction": request.instruction,
#         "slide_number": slide_number,
#         "source": "user_input"
#     }
#     categorized_tasks = classify_feedback_instructions([feedback_item])
#     logger.info(f"Categorized Tasks: {len(categorized_tasks)}")

#     if not categorized_tasks:
#         return {"status": "no_tasks", "message": "No actionable feedback classified."}

#     # Step 2: Delegate Tasks to Agents
#     logging.info("Processing Tasks with Agents & Context...")
#     task_specifications = []

#     for task in categorized_tasks:
#         category = task["category"]
#         instruction = task["original_instruction"]

#         try:
#             # Dispatch to the appropriate agent
#             if category == "formatting":
#                 result = formatting_agent(task, slide_context)
#             elif category == "cleanup":
#                 result = cleanup_agent(task, slide_context)
#             elif category == "visual_enhancement":
#                 result = visual_enhancement_agent(task, slide_context)
#             else:
#                 logger.warning(f"Unknown category: {category} for instruction: '{instruction}'")
#                 continue

#             if result:
#                 task_specifications.extend(result)
#         except Exception as e:
#             logger.error(f"Error processing task for slide {slide_number}: {e}")

#     if not task_specifications:
#         logger.warning("No tasks generated from the feedback. Nothing to process.")
#         return {"status": "no_tasks", "message": "No tasks to process"}

#     predefined_execution_report = {
#         "status": "success",
#         "processed_count": len(task_specifications),
#         "success_count": 0,
#         "errors": []
#     }

#     return {
#         "status": "success",
#         "parsed_output": task_specifications,
#         "execution_report": predefined_execution_report
#     }

@app.post("/process_instruction")
async def process_instruction(request: InstructionRequest):
    logger.info(f"Received instruction: {request.instruction}")
    logger.info(f"Target slide index: {request.slide_index}")

    slide_images_dir = "uploaded_pptx/slide_images/presentation"
    slide_number = request.slide_index

    slide_image_path = os.path.join(slide_images_dir, f"slide{slide_number}_image.txt")
    slide_xml_path = os.path.join(slide_images_dir, f"slide{slide_number}.xml")
    slide_png_path = os.path.join(slide_images_dir, f"slide{slide_number}.png")

    if not (os.path.exists(slide_image_path) and os.path.exists(slide_xml_path)):
        logger.error(f"Slide context files for slide {slide_number} do not exist!")
        return {"status": "error", "message": f"Slide context for slide {slide_number} not found."}

    # --- Load Slide Context ---
    slide_context = {}
    try:
        with open(slide_image_path, "r", encoding="utf-8") as img_file:
            slide_context["slide_image_base64"] = img_file.read()

        with open(slide_xml_path, "r", encoding="utf-8") as xml_file:
            slide_context["slide_xml_structure"] = xml_file.read()

        if os.path.exists(slide_png_path):
            with open(slide_png_path, "rb") as img_bin_file:
                slide_context["slide_image_bytes"] = img_bin_file.read()
        else:
            logger.warning(f"Slide PNG file not found at: {slide_png_path}")
            slide_context["slide_image_bytes"] = b""

        logger.info(f"Slide context loaded for slide {slide_number}")
    except Exception as e:
        logger.error(f"Error loading context for slide {slide_number}: {e}")
        return {"status": "error", "message": f"Error loading context for slide {slide_number}"}

    # --- Step 1: Classify Instruction ---
    feedback_item = {
        "instruction": request.instruction,
        "slide_number": slide_number,
        "source": "user_input"
    }
    categorized_tasks = classify_feedback_instructions([feedback_item])
    logger.info(f"Categorized Tasks: {len(categorized_tasks)}")

    if not categorized_tasks:
        return {"status": "no_tasks", "message": "No actionable feedback classified."}

    # --- Step 2: Run Agents ---
    logger.info("Processing Tasks with Agents & Context...")
    task_specifications = []

    for task in categorized_tasks:
        category = task["category"]
        instruction = task["original_instruction"]

        try:
            if category == "formatting":
                result = formatting_agent(task, slide_context)
            elif category == "cleanup":
                result = cleanup_agent(task, slide_context)
            elif category == "visual_enhancement":
                result = visual_enhancement_agent(task, slide_context)
            else:
                logger.warning(f"Unknown category: {category} for instruction: '{instruction}'")
                continue

            if result:
                task_specifications.extend(result)
        except Exception as e:
            logger.error(f"Error processing task for slide {slide_number}: {e}")

    if not task_specifications:
        logger.warning("No tasks generated from the feedback. Nothing to process.")
        return {"status": "no_tasks", "message": "No tasks to process"}

    #  --- Save task_specifications as JSON ---
    try:
        json_output_path = os.path.join(slide_images_dir, f"slide{slide_number}_tasks.json")
        with open(json_output_path, "w", encoding="utf-8") as f:
            json.dump(task_specifications, f, indent=4)
        logger.info(f"Saved processed subtasks to: {json_output_path}")
    except Exception as e:
        logger.error(f"Failed to save subtasks JSON: {e}")

    # --- Step 3: Refiner Agent (loads task JSON internally) ---
    logger.info("Refining Instructions with Refiner Agent...")
    refined_instructions = refiner_agent(
        slide_number=slide_number,
        slide_context=slide_context
    )

    # --- Step 4: Generate Office.js Code from Refined Instructions ---
    if not isinstance(refined_instructions, dict) or "refined_instructions" not in refined_instructions:
        logger.error(f"Refiner agent failed or returned unexpected result: {refined_instructions}")
        return {"status": "error", "message": "Refiner agent failed or returned invalid structure."}

    logger.info("Refined instructions received from refiner agent:")
    logger.info(f"--- BEGIN REFINED INSTRUCTIONS ---\n{refined_instructions['refined_instructions']}\n--- END REFINED INSTRUCTIONS ---")

    # Join the list of refined instructions into a single string
    refined_instruction_str = "\n".join(refined_instructions['refined_instructions'])

    # Now pass the concatenated string to generate_code
    logger.info("Generating Office.js code from refined instructions...")
    code_generation_result = generate_code(refined_instruction_str)


    if "code" in code_generation_result:
        logger.info("Code generation successful.")
        return {
            "status": "success",
            "message": "Office.js code generated successfully.",
            "code": code_generation_result["code"]  # <-- Return this!
        }
    else:
        logger.error("Code generation failed.")
        return {
            "status": "error",
            "message": code_generation_result.get("error", "Unknown error during code generation.")
        }

# --- Main Execution ---
if __name__ == "__main__":
    logger.info("Starting Uvicorn server for development...")
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)


# # --- Instruction Endpoint ---
# @app.post("/process_instruction")
# async def process_instruction(request: InstructionRequest):
#     logger.info(f"Received instruction: {request.instruction}")
#     logger.info(f"Target slide index: {request.slide_index}")
    
#     # --- Slide Context Directory ---
#     slide_images_dir = "uploaded_pptx/slide_images/presentation"

#     slide_number = request.slide_index 
    
#     # Flat file paths, not subfolders
#     slide_image_path = os.path.join(slide_images_dir, f"slide{slide_number}_image.txt")
#     slide_xml_path = os.path.join(slide_images_dir, f"slide{slide_number}.xml")

#     # Check if required files exist
#     if not (os.path.exists(slide_image_path) and os.path.exists(slide_xml_path)):
#         logger.error(f"Slide context files for slide {slide_number} do not exist!")
#         return {"status": "error", "message": f"Slide context for slide {slide_number} not found."}

#     # Load context
#     slide_context = {}
#     try:
#         with open(slide_image_path, "r", encoding="utf-8") as img_file:
#             slide_context["slide_image_base64"] = img_file.read()

#         with open(slide_xml_path, "r", encoding="utf-8") as xml_file:
#             slide_context["slide_xml_structure"] = xml_file.read()

#         slide_png_path = os.path.join(slide_images_dir, f"slide{slide_number}.png")
#         if os.path.exists(slide_png_path):
#             with open(slide_png_path, "rb") as img_bin_file:
#                 slide_context["slide_image_bytes"] = img_bin_file.read()
#         else:
#             logger.warning(f"Slide PNG file not found at: {slide_png_path}")
#             slide_context["slide_image_bytes"] = b""

#         logger.info(f"Slide context loaded for slide {slide_number}")
#     except Exception as e:
#         logger.error(f"Error loading context for slide {slide_number}: {e}")
#         return {"status": "error", "message": f"Error loading context for slide {slide_number}"}

#     # Step 1: Instruction Interpretation
#     feedback_item = {
#         "instruction": request.instruction,
#         "slide_number": slide_number,
#         "source": "user_input"
#     }
#     categorized_tasks = classify_feedback_instructions([feedback_item])
#     logger.info(f"Categorized Tasks: {len(categorized_tasks)}")

#     if not categorized_tasks:
#         return {"status": "no_tasks", "message": "No actionable feedback classified."}

#     # Step 2: Delegate Tasks to Agents
#     logging.info("Processing Tasks with Agents & Context...")
#     task_specifications = []

#     for task in categorized_tasks:
#         category = task["category"]
#         instruction = task["original_instruction"]

#         try:
#             # Dispatch to the appropriate agent
#             if category == "formatting":
#                 result = formatting_agent(task, slide_context)
#             elif category == "cleanup":
#                 result = cleanup_agent(task, slide_context)
#             elif category == "visual_enhancement":
#                 result = visual_enhancement_agent(task, slide_context)
#             else:
#                 logger.warning(f"Unknown category: {category} for instruction: '{instruction}'")
#                 continue

#             if result:
#                 task_specifications.extend(result)
#         except Exception as e:
#             logger.error(f"Error processing task for slide {slide_number}: {e}")

#     if not task_specifications:
#         logger.warning("No tasks generated from the feedback. Nothing to process.")
#         return {"status": "no_tasks", "message": "No tasks to process"}

#     # Step 3: Refine Instructions using Refiner Agent
#     logger.info("Refining Instructions with Refiner Agent...")
#     refined_instructions = refiner_agent(
#         instruction_output=task_specifications,
#         slide_context=slide_context,
#         slide_number=slide_number,
#         shape_metadata= "uploaded_pptx/slide_images/metadata.json"
#     )

#     if not refined_instructions:
#         logger.warning("No refined instructions were produced.")
#         return {"status": "no_refined_instructions", "message": "No refined instructions were generated."}

#     # Report the tasks and refined instructions
#     predefined_execution_report = {
#         "status": "success",
#         "processed_count": len(task_specifications),
#         "success_count": len(refined_instructions),
#         "errors": []
#     }

#     return {
#         "status": "success",
#         "parsed_output": refined_instructions,
#         "execution_report": predefined_execution_report
#     }


# # --- Main Execution ---
# if __name__ == "__main__":
#     logger.info("Starting Uvicorn server for development...")
#     uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)