import json
import os
import base64
import requests
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

# CORS setup (optional for frontend integration)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

LLM_API_KEY = "AIzaSyAv0sTw83EOKcJtoSyT9ug4cnzwGagkMJY"
GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={LLM_API_KEY}"

SLIDE_LAYOUT_OPTIMIZATION_PROMPT_TEMPLATE = """You are an expert assistant for slide layout refinement. Your role is to analyze the visual layout of a presentation slide using a screenshot and shape metadata, then generate clear and actionable natural language formatting instructions to improve the visual consistency, alignment, proportions, and overall structure of the shapes on the slide.

---

**SLIDE DIMENSIONS:** 960px width x 540px height (16:9 aspect ratio)

---

**INPUT:**

1.  **Slide Image:** A screenshot of the slide (for visual reference).
2.  **Shape Metadata (JSON):** An array of shape objects, each containing:
    *   id: Unique identifier for the shape
    *   name: Shape name or label (if any)
    *   type: Type of shape (e.g., 'rectangle', 'chevron', 'text box')
    *   top, left, width, height: Position and dimensions of the shape (in pixels)
    *   zIndex: The stacking order of the shape
    *   text: The visible text in the shape
    *   isGrouped: Boolean indicating if shape is part of a group
    *   overlapsWith: List of shape ids this shape overlaps with
    *   font: Font information (e.g., name, size, color)

---

**YOUR TASK:**

Analyze the layout structure using the visual and metadata input, and then:

1.  Detect inconsistencies in alignment (top, bottom, left, right, center), size (width, height), spacing, especially among visually related elements.
2.  Identify implicit table- or grid-like structures, paying close attention to rows or columns of similar shapes (e.g., rectangular content blocks, chevrons, process steps).
3.  Determine whether the shapes can be cleanly arranged into a grid layout without causing visual or spatial overflow.
4.  Propose minimal, proportional layout improvements to enhance structure and visual appeal:
    *   Prioritize **uniform dimensions (width and/or height)** for shapes intended to be visually consistent within a row or grid (e.g., all steps in a process flow).
    *   Suggest **adjustments to height or width** to create visual balance between adjacent rows or sections, or to align elements more effectively.
    *   Recommend **grid-based alignments** (top-left anchoring, center alignment within columns, or consistent spacing between rows).
    *   Detect and suggest solutions to **overflow risks** (e.g., if the total width of a row exceeds the slide width, suggest reducing shape size or spacing).
    *   Adjust spacing between elements for consistency and visual clarity.
    *   Correct inconsistencies in text formatting for similar types of labels or elements.
    *   Limit positional shifts to what's necessary for alignment, spacing, or avoiding overlap — preserve the overall design intent.

---

**OUTPUT:**

For each detected improvement, output a **clear, natural language instruction** in the following format. Each instruction must target a specific, actionable change.

*   **Instruction:** "Align shapes (ids: [102, 103, 104]) into a single row with top alignment at 245px and equal horizontal spacing of 40px between them."
*   **Instruction:** "Resize all chevron shapes in the top row (ids: [11, 12, 13, 14]) to a uniform width of 150px and uniform height of 55px to create visual consistency."
*   **Instruction:** "Increase the height of the rectangular content blocks (ids: [21, 22, 23, 24]) to 120px to visually balance them with the top row and provide more content space."
*   **Instruction:** "Reduce the width of grid-aligned shapes (ids: [31, 32, 33, 34]) from 200px to 170px to prevent horizontal overflow and maintain consistent spacing."

Each instruction must:
*   Reference shape `ids` clearly.
*   Specify the property to modify (e.g., top, left, width, height, font color).
*   Provide the **exact new value** (e.g., `150px`, `green`, `#38761d`) or describe the alignment/distribution target.
*   Briefly justify the change in terms of visual consistency, alignment, uniformity, proportionality, or overflow prevention if not immediately obvious.

---

**IMPORTANT CONSTRAINTS:**

*   Focus solely on **formatting and layout adjustments**. Do not suggest content changes.
*   Prioritize minimal changes that yield significant visual improvement. Avoid large or disruptive redesigns.
*   Respect existing groupings and element relationships.
*   Be mindful of surrounding elements to prevent new overlaps or visual imbalance.
*   Output only the **formatted instruction lines** – no summaries, explanations, or conversational text.
"""

@app.get("/analyze-slide")
def analyze_slide():
    image_path = "slide_images/images/presentation/slide_1.png"
    metadata_path = "slide_images/metadata.json"

    if not os.path.exists(image_path):
        return {"error": "Slide image not found."}
    if not os.path.exists(metadata_path):
        return {"error": "Metadata JSON not found."}

    try:
        with open(image_path, "rb") as img_file:
            base64_img = base64.b64encode(img_file.read()).decode("utf-8")
        with open(metadata_path, "r", encoding="utf-8") as json_file:
            metadata_content = json.load(json_file)
            metadata_str = json.dumps(metadata_content, indent=2)
    except Exception as read_err:
        return {"error": f"File read error: {read_err}"}

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": SLIDE_LAYOUT_OPTIMIZATION_PROMPT_TEMPLATE},
                    {
                        "inline_data": {
                            "mime_type": "image/png",
                            "data": base64_img
                        }
                    },
                    {"text": f"Here is the shape metadata (JSON):\n\n{metadata_str}"}
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.2,
            "topK": 40,
            "topP": 1.0,
            "maxOutputTokens": 2048
        }
    }

    response = requests.post(GEMINI_URL, json=payload)
    if response.status_code != 200:
        return {"error": f"Gemini API Error: {response.text}"}

    try:
        output = response.json()["candidates"][0]["content"]["parts"][0]["text"]
        return {"analysis": output}
    except Exception as e:
        return {"error": f"Invalid response from Gemini: {e}"}
