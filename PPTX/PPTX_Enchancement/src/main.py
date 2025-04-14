from fastapi import Body, FastAPI
from fastapi.middleware.cors import CORSMiddleware
from feedback_parsing.feedback_classifier import classify_instruction
from agents.cleanup_agent import analyze_slide
from code_manipulation.code_generator import generate_code
import uvicorn
from routes.metadata_handler import router as metadata_router
from routes.pptx_handler import router as pptx_router

app = FastAPI()

# Enable CORS (allow all for development)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# === Route registration ===
app.include_router(metadata_router, prefix="/upload-metadata", tags=["Metadata"])
app.include_router(pptx_router, prefix="/upload-pptx", tags=["PPTX Upload"])

@app.post("/agent/classify-instruction")
def classify_instruction_handler(payload: dict = Body(...)):
    instruction = payload.get("instruction")
    slide_number = payload.get("slide_number")

    if not instruction:
        return {"status": "error", "message": "Instruction is required"}

    classification_result = classify_instruction(instruction, slide_number)

    return {
        "status": "success",
        "result": classification_result
    }

@app.post("/agent/cleanup")
def cleanup_handler():
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

@app.post("/agent/generate_code")
def generate_code_handler(data: dict):
    instructions = data.get("instructions", "")
    return generate_code(instructions)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True) 