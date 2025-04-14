import requests
import base64
import os

GEMINI_API_KEY = "AIzaSyAv0sTw83EOKcJtoSyT9ug4cnzwGagkMJY"
GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={GEMINI_API_KEY}"

CODE_GEN_PROMPT_TEMPLATE = """
You are a senior PowerPoint automation developer using Office.js (JavaScript API for PowerPoint Add-ins).
You are given natural language instructions that describe how to clean or modify a PowerPoint slide.
Your task is to generate **fully-executed, production-ready JavaScript code** using only supported Office.js APIs.
Your response **must include the logic to implement every single instruction provided**, even if there are multiple shapes or steps involved. **Do not skip or summarize**. Each action (alignment, resizing, repositioning, etc.) must be directly reflected in code.
You must generate code that implements **every single instruction** accurately and completely, following the strict structure below...

The code must follow this exact structure:
- Use `const shapes = context.presentation.slides.getItemAt(0).shapes;`
- Use `shapes.load("items/name, items/left, items/top, items/width, items/height");`
- `await context.sync();`
- For each shape to be modified:
  - Use `const shapeXYZ = shapes.items.find(s => s.name === "Google Shape;ID;pX");`
    - Replace `ID` with the actual shape ID (e.g., `474`)
    - Replace `pX` with the slide index (e.g., `p1`, `p2`, etc.)
    - Variable name should be based on the shape ID (e.g., `shape474`, `shape480`)
  - Always check `if (shapeXYZ)` before modifying or accessing any properties
  - If reading properties (e.g., `.width`, `.height`), they should have been loaded in the initial `shapes.load()`.
  - If reading or comparing shape properties like `.left`, `.top`, `.width`, etc., ensure the initial `load()` and `await context.sync()` have been performed.
  - Apply updates (e.g., `shapeXYZ.top = 100`)
- Use `await context.sync();` after all updates
- Use `else {{ console.error("Shape not found"); }}` if a shape is missing

Ensure:
- **All shape names follow** the format: `"Google Shape;ID;pX"`
- Do **not** convert to any default shape name like "Rectangle", "Title", etc.

Do NOT include:
- `PowerPoint.run(...)`
- Wrappers like `async function` or `cleanSlide(...)`
- Logs or `console.log` (except inside `else`)
- Comments or explanations outside the code

Instruction: **Based on the slide analysis, implement the following adjustments:**
{}

Return only the code block that follows the exact layout and conventions above.
"""


def generate_code(analysis_output: str):
    instructions = f"Based on the slide analysis, implement the following adjustments:\n{analysis_output}"
    prompt = CODE_GEN_PROMPT_TEMPLATE.format(instructions.strip())

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt}
                ]
            }
        ]
    }

    response = requests.post(GEMINI_URL, json=payload)
    if response.status_code != 200:
        return {"error": f"Gemini API Error: {response.status_code} - {response.text}"}

    try:
        res_json = response.json()
        candidates = res_json.get("candidates", [])
        if not candidates:
            return {"error": "No candidates in Gemini response."}

        parts = candidates[0].get("content", {}).get("parts", [])
        if not parts or "text" not in parts[0]:
            return {"error": "Malformed Gemini response structure."}

        code = parts[0]["text"]
        return {"code": code}

    except Exception as e:
        return {"error": f"Exception in processing Gemini response: {e}"}