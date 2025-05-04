
import logging
import uvicorn
from fastapi import Body, FastAPI, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from agents.grid_analyzer import analyze_grid_structure_and_save 
from feedback_parsing.feedback_classifier import classify_instruction
from agents.cleanup_agent import analyze_slide
from code_manipulation.code_generator import generate_code
from routes.metadata_handler import router as metadata_router
from routes.pptx_handler import router as pptx_router

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
class ClassifyInstructionPayload(BaseModel):
    instruction: str
    slide_number: int | str | None = None 

class GenerateCodePayload(BaseModel):
    instructions: str
    
# === Route Registration ===
try:
    app.include_router(metadata_router, prefix="/upload-metadata", tags=["Metadata"])
    app.include_router(pptx_router, prefix="/upload-pptx", tags=["PPTX Upload"])
    logger.info("Metadata and PPTX routers included successfully.")
except NameError:
    logger.error("Failed to include routers. Check if router objects are defined correctly.")

# === API Endpoints ===
@app.post("/agent/classify-instruction", tags=["Agent Actions"])
async def classify_instruction_handler(payload: ClassifyInstructionPayload = Body(...)):
    """Classifies a natural language instruction for slide modification."""
    logger.info(f"Received request to classify instruction for slide {payload.slide_number}")
    try:
        classification_result = classify_instruction(payload.instruction, payload.slide_number)
        logger.info("Instruction classified successfully.")
        return {
            "status": "success",
            "result": classification_result
        }
    except Exception as e:
        logger.error(f"Error during instruction classification: {e}", exc_info=True)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to classify instruction: {e}"
        )

@app.post("/agent/cleanup")
def cleanup_handler():
    # === Step 1: Run Grid Analysis ===
    print("Starting grid analysis...")
    analysis_result_data = analyze_grid_structure_and_save()
    if analysis_result_data is None:
        print("\nGrid analysis failed. Aborting cleanup.")
        raise HTTPException(status_code=500, detail="Grid analysis step failed. Cannot proceed.")
    else:
        print("\nGrid analysis complete. Result saved.")

    print("Generating cleanup instructions...")
    result = analyze_slide()

    if "analysis" in result:
        print("Cleanup instructions received and forwarded to frontend.")
        code_result = generate_code(result["analysis"])
        if "code" in code_result:
            print("Generated Code:\n", code_result["code"])
            return {
                "status": "success",
                "instructions": result["analysis"],
                "code": code_result["code"]
            }
        else:
            return {
                "status": "partial_success",
                "instructions": result["analysis"],
                "message": code_result.get("error")
            }
    else:
        print("Cleanup failed:", result.get("error"))
        return {
            "status": "error",
            "message": result.get("error", "Unknown error")
        }

@app.post("/agent/generate_code", tags=["Agent Actions"])
async def generate_code_handler(payload: GenerateCodePayload = Body(...)):
    logger.info("Received request to generate code directly from instructions.")
    try:
        code_result = generate_code(payload.instructions)
        if "error" in code_result:
            error_msg = code_result.get("error", "Unknown error from code generation")
            logger.error(f"Direct code generation failed: {error_msg}")
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Code generation failed: {error_msg}"
            )
        if "code" not in code_result or not code_result["code"]:
             logger.error("Direct code generation returned empty code.")
             raise HTTPException(
                 status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                 detail="Code generation resulted in empty code."
             )
        logger.info("Direct code generation successful.")
        return {
            "status": "success",
            "code": code_result["code"]
        }
    except Exception as e:
        logger.error(f"Unexpected error during direct code generation call: {e}", exc_info=True)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"An unexpected error occurred during code generation: {e}"
        )

# --- Main Execution ---
if __name__ == "__main__":
    logger.info("Starting Uvicorn server for development...")
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)