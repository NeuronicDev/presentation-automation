import base64
import logging, json, re
from typing import Dict, Any
from config.config import LLM_API_KEY
from google import genai
import os

client = genai.Client(api_key=LLM_API_KEY)
logger = logging.getLogger(__name__)

METADATA_PATH = "uploaded_pptx/slide_images/metadata.json"

def load_tasks_from_file(slide_number: int):
    file_path = f"uploaded_pptx/slide_images/presentation/slide{slide_number}_tasks.json"
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        raise FileNotFoundError(f"Task file not found for slide {slide_number}")

def load_shape_metadata_for_slide(slide_number: int):
    if not os.path.exists(METADATA_PATH):
        raise FileNotFoundError("Shape metadata file not found")
    with open(METADATA_PATH, "r", encoding="utf-8") as f:
        all_shapes = json.load(f)
    return [shape for shape in all_shapes if shape.get("slideIndex") == slide_number]

def refiner_agent(slide_number: int, slide_context: Dict[str, str]) -> Dict[str, Any]:
    try:
        tasks = load_tasks_from_file(slide_number)
        shapes = load_shape_metadata_for_slide(slide_number)
        slide_image_base64 = slide_context.get("slide_image_base64", "")
        
        # Refiner Agent Prompt Template
        INSTRUCTION_REFINER_PROMPT = f"""
        You are an AI assistant that transforms abstract or natural language layout modification tasks into **explicit, executable PowerPoint (Office.js-compatible)** commands.  
        Your task is to Convert vague or general cleanup instructions into **precise, technically actionable commands** using the provided slide image, shape metadata, and original natural language instructions. Your response will help a code generator apply these exact changes.

        **Input:**
        Raw Natural Language Instructions: {json.dumps(tasks)}
        **Slide Shape Metadata (JSON):** An array of shape objects, each containing `id`, `type`, `top`, `left`, `width`, `height`, `text`, `font`, etc. {json.dumps(shapes)}
        **Slide Visual Image** (base64-encoded, included below as image reference):
        {slide_image_base64}

        **Analysis Instructions:**
        - **Use the shape metadata and image context** to resolve any ambiguities in the task descriptions.
        - **Match instructions to specific shape(s)** using shape ID, text, position, or overlaps.
        - Use standard layout understanding to resolve alignment, font, spacing, or overflow changes.
        - When IDs are involved, **always include them** in the output instruction.

        **Output Format:**
        Return ONLY a list of explicit Office.js-compatible instruction strings in this format:

        ```json
        {{ 
        "refined_instructions": [
            "Change font for shape (id: 101) to 'Arial' size 32pt.",
            "Reduce the width of shape (id: 31) to 180px (currently overflowing) to fit within the slide width (960px).",
            "Align shapes (ids: [102, 103, 104]) to the top at 245px and distribute them horizontally with 40px spacing, based on their row alignment in the grid."
        ]
        }}
        ```
        DO NOT include explanations, only the JSON output. 
        """

        # Correct API call to Gemini API (this will vary)
        response = client.models.generate_content(
            model="gemini-2.0-flash", 
            contents=[INSTRUCTION_REFINER_PROMPT, slide_image_base64]
        )
        logging.info(f"LLM refiner_agent response: {response.text}")

        # Use regex to find the JSON block in the response text
        json_match = re.search(r'(\{[\s\S]*\})', response.text)
        if json_match:
            refined_instructions = json.loads(json_match.group(0))
            return refined_instructions
        else:
            logging.error("No valid JSON found in response.")
            return {"error": "No valid JSON found in response."}

    except Exception as e:
        logger.error(f"Error during instruction refinement: {e}")
        return {"error": str(e)}