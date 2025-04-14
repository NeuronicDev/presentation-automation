import base64
import json
import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import requests

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

SLIDE_GRID_ANALYSIS_PROMPT_TEMPLATE = """
You are an expert assistant specializing in refining presentation slide layouts based on a pre-analyzed grid structure, visual appearance, and shape metadata. Your task is to identify layout inconsistencies (primarily dimensional uniformity within grid rows/columns) and propose specific, actionable formatting instructions to fix them, while strictly adhering to the slide boundaries and maintaining the overall grid integrity.

---

**INPUT:**

1.  **Slide Image:** A screenshot of the slide (for visual reference).
2.  **Shape Metadata (JSON):** An array of shape objects with `id`, `type`, `top`, `left`, `width`, `height`, `text`, `font`, etc.
3.  **Grid Analysis Results (JSON):** The output from the grid analysis step, containing:
    *   `is_grid_structure`: boolean
    *   `grid_size`: { "rows": number, "columns": number }
    *   `grid_structure`: { "rows": { "row_N": [ids...] }, "columns": { "col_N": [ids...] } }
    *   `reasoning`: text
4.  **Slide Dimensions:**
    *   `width`: 960 (pixels)
    *   `height`: 540 (pixels)

---

**YOUR TASK:**

**Pre-condition:** Only proceed if `is_grid_structure` in the Grid Analysis Results is `true`. If `false`, output no instructions.

1.  **Analyze Grid Content:** Using the `grid_structure` and `Shape Metadata`:
    *   **Identify Inconsistent Dimensions:** For each `row` in `grid_structure.rows`, examine the shapes listed. Identify groups of shapes within that row that have the **same `type`** (e.g., all 'chevron' shapes, all 'rectangle' content blocks) but **different `width` or `height`**. These are candidates for uniform sizing. Do a similar check for shapes of the same type within each `column`.
    *   **Identify Text Color Inconsistencies:** Check for text elements (`text box` type or text within shapes) that serve similar purposes (e.g., labels on the left, text within main content blocks) but have inconsistent font colors.
    *   **(Optional) Check Alignment/Spacing:** Within rows, check for inconsistent vertical alignment (top/middle/bottom) of related elements. Within columns, check for inconsistent horizontal alignment. Check for uneven spacing between elements in rows or columns.

2.  **Propose Uniform Dimension Adjustments (Primary Focus):**
    *   For each group of same-type shapes within a row/column identified as having inconsistent dimensions:
        *   **Determine Target Size:** Propose a uniform `width` and/or `height`. A good target might be the average size, or the size of the most visually prominent element, *initially*.
        *   **Validate Against Constraints:**
            *   **Calculate Total Space:** Calculate the total width required for all elements in the row (or height for columns) if they adopt the target size, including existing or desired spacing.
            *   **Check Slide Boundaries:** Ensure the calculated total width/height does not exceed the slide dimensions (`960px` width, `540px` height) and that individual elements remain within bounds (`left >= 0`, `top >= 0`, `left + new_width <= 960`, `top + new_height <= 540`).
            *   **Check Overlaps:** Ensure the proposed changes don't cause significant new overlaps between elements that shouldn't overlap.
        *   **Adjust if Necessary (Conflict Resolution):**
            *   **If Overflow/Overlap:** Can the *target uniform size* be slightly **reduced** for *all* shapes in the group to make them fit within boundaries and avoid overlaps? If yes, propose this adjusted uniform size.
            *   **If Reduction Fails:** Can the **spacing** between the elements be slightly adjusted (likely reduced) to accommodate the desired uniform size? If yes, propose both the uniform size and the new spacing. (Be cautious not to make spacing too tight).
            *   **If Still Fails:** If no reasonable adjustment allows for uniform sizing without violating constraints, **do not propose** the uniform sizing change for that group. Prioritize fitting within boundaries over achieving perfect uniformity.
    *   **Generate Instruction:** If a valid uniform size (potentially adjusted) is found, generate a clear instruction specifying the shape `ids`, the property (`width`, `height`), and the final calculated `new value`. Justify it based on consistency and mention if adjustments were made for fitting.

3.  **Propose Other Adjustments (Secondary):**
    *   **Text Color:** If inconsistent text colors were found for related elements, propose changing them to a consistent color (e.g., pick the most common color or a standard one like black/green as appropriate). Generate an instruction.
    *   **Alignment/Spacing:** If significant alignment or spacing inconsistencies *remain* after dimension adjustments or were identified independently, propose fixes (e.g., "Align top edges of shapes [ids] at Y=...", "Set horizontal spacing between shapes [ids] to X px"). Validate these against boundaries too.

---

**OUTPUT FORMAT:**

Output **only** the formatted instruction lines for the changes you propose. If no changes are needed or possible according to the constraints, output nothing or a comment like `// No layout adjustments needed or possible within constraints.`.

*   **Instruction:** "Resize shapes (ids: [11, 12, 13, 14]) of type 'chevron' in row_1 to uniform width 145px and uniform height 55px for consistency (adjusted width from 150px to fit within slide bounds)."
*   **Instruction:** "Increase height of shapes (ids: [21, 22, 23, 24]) of type 'rectangle' in row_2 to 120px to visually balance with row_1." (Assuming validation passed).
*   **Instruction:** "Update text color for shapes (ids: [91, 92, 93, 94]) to green (hex: #38761d) for consistent label styling."
*   **Instruction:** "Align top edges of shapes (ids: [21, 22, 23, 24]) in row_2 at 180px."
*   **Instruction:** "Set horizontal spacing between shapes (ids: [11, 12, 13, 14]) in row_1 to 15px."

Each instruction must:
*   Reference shape `ids` clearly.
*   Specify the property to modify (e.g., width, height, font color, top, spacing).
*   Provide the **exact new value** or describe the alignment/distribution target.
*   Briefly justify the change (consistency, balance, alignment) and **note if values were adjusted due to constraints.**

---

**IMPORTANT CONSTRAINTS:**

*   **Strict Boundary Adherence:** No proposed change should result in any part of any shape exceeding the 960x540 slide dimensions.
*   **Grid Integrity:** Changes should respect the identified row/column structure. Don't move elements out of their logical grid position unless adjusting spacing *within* that row/column.
*   **Minimal Changes:** Prioritize fixing major inconsistencies (like dimensions) first. Avoid unnecessary tweaks.
*   **No Overlaps:** Avoid creating new, unintended overlaps between elements.
*   **Focus on Layout:** Do not suggest content changes.
*   **Output Only Instructions:** Do not include summaries, explanations, or conversational text outside the specific `Instruction:` lines.

"""

# Load image and encode as base64
def load_image_base64(image_path):
    with open(image_path, "rb") as img:
        return base64.b64encode(img.read()).decode("utf-8")

# Load JSON metadata
def load_metadata(metadata_path):
    with open(metadata_path, "r") as f:
        return json.load(f)

# Prepare and send the request
def analyze_grid_structure(image_path, metadata_path):
    image_base64 = load_image_base64(image_path)
    metadata = load_metadata(metadata_path)

    request_body = {
        "contents": [
            {
                "parts": [
                    {"text": SLIDE_GRID_ANALYSIS_PROMPT_TEMPLATE},
                    {
                        "inlineData": {
                            "mimeType": "image/png",
                            "data": image_base64,
                        }
                    },
                    {
                        "text": f"\n\nHere is the shape metadata:\n~~~json\n{json.dumps(metadata, indent=2)}\n~~~"
                    }
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

    response = requests.post(GEMINI_URL, json=request_body)
    if response.status_code == 200:
        content = response.json()
        try:
            reply = content["candidates"][0]["content"]["parts"][0]["text"]
            print("\n--- Gemini Analysis Output ---\n")
            print(reply)
        except Exception as e:
            print("Error parsing Gemini response:", e)
            print("Full response:", content)
    else:
        print("Request failed with status code:", response.status_code)
        print(response.text)

# Run analysis
if __name__ == "__main__":
    image_path = "images/presentation/slide_1.png"

    metadata_path = "metadata.json"
    analyze_grid_structure(image_path, metadata_path)