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

LAYOUT_ANALYSIS_PROMPT = """You are an expert AI assistant analyzing a PowerPoint slide to detect and resolve layout and design issues with minimal disruption to the existing layout. Your goal is to identify and resolve overflow (horizontal and vertical), misalignment, overlap, and inconsistent sizing among shapes, while being mindful of the element's surroundings and avoiding significant positional changes. Additionally, detect table-like layouts and suggest grid-based alignment where appropriate.

**Input:**
- Slide Number: {slide_number}
- Slide Image: A screenshot of the slide (for visual reference).
- Shape Metadata (JSON): An array of shape objects containing:
  - id, name, type, top, left, width, height, zIndex, text, isGrouped, overlapsWith

---

**1. Overflow Detection & Strict Resolution (No Overflow Allowed)**

**Slide boundaries:** width = `960px`, height = `540px`

**Horizontal Overflow (Right Edge):**
- For each shape, if `left + width >= 960`:
  - Identify its **row**: other shapes with `top` within ±10px.
  - Compute:
    - `row_min_left = min([s['left'] for s in row_shapes])`
    - `row_max_right = max([s['left'] + s['width'] for s in row_shapes])`
    - `row_width = row_max_right - row_min_left`
    - `horizontal_overflow = row_max_right - 960`
    - `width_scale_factor = (row_width - horizontal_overflow) / row_width`
      - If `width_scale_factor >= 0.7`, proceed with **proportional resizing**:
        - **Instruction (Slide {slide_number})**: "For shapes (ids: {row_ids}), proportionally resize to fit within slide by setting `width` to `original_width * {width_scale_factor:.2f}`, `height` to `original_height * {width_scale_factor:.2f}`, and `left` to `{row_min_left} + ((original_left - {row_min_left}) * {width_scale_factor:.2f})` to resolve right overflow."
      - Else:
        - `uniform_width_reduction = horizontal_overflow / len(row_shapes)`
        - **Instruction (Slide {slide_number})**: "Reduce `width` of shapes (ids: {row_ids}) uniformly by `{uniform_width_reduction:.2f}`px and adjust `left` by `original_left - (original_left - {row_min_left}) * ({uniform_width_reduction:.2f} / {row_width})` to preserve relative spacing and resolve right overflow."
      - **Crucially:** If even uniform reduction doesn't fully resolve overflow without extreme shrinkage (scale_factor < 0.6), consider a very minor left shift (<= 10px) of the entire row in conjunction with resizing. Document this carefully.

**Vertical Overflow (Bottom Edge):**
- If `top + height >= 540`:
  - Identify **column**: shapes with `left` within ±10px.
  - Apply similar logic as horizontal overflow, adjusting `top` and `height`.
  - **Instruction (Slide {slide_number})**: "Resize and adjust shapes (ids: {column_ids}) using vertical scaling to prevent bottom overflow."

---

**2. Layout Consistency Fixes (Aware of Surroundings & Limited Movement)**

**A. Misalignment in Rows/Columns:**
- Group into **rows** and **columns**.
- For each group:
  - Identify a well-positioned "anchor" shape (e.g., the first or largest).
  - For other shapes with `top` or `left` differing by > 5px from the anchor:
    - If adjustment to match anchor is within 10px of original:
      - **Instruction (Slide {slide_number})**: "Align shape '{misaligned_id}' by setting `top` to {anchor_top}px (or `left` to {anchor_left}px) to align with '{anchor_id}'."
    - Else, suggest the *minimal* adjustment needed to reduce the difference, not exceeding 10px.

**B. Overlapping Shapes (Not Grouped):**
- For each shape with `overlapsWith`:
  - If not `isGrouped`, and collision can be fixed with ≤10px shift:
    - Determine the direction of minimal shift (horizontal or vertical).
    - **Instruction (Slide {slide_number})**: "Shift shape '{overlapping_id}' by {shift_amount:.2f}px on `left` (or `top`) to avoid overlap with '{other_id}'."
  - If a small shift isn't enough, consider very minor proportional resizing (scale factor >= 0.95) of the overlapping shapes.

**C. Inconsistent Size (Same Type/Role, Aware of Row/Column):**
- For shapes in the same row/column with similar `name` or `type`:
  - Calculate the median `width` and `height`.
  - For shapes deviating by > 10% from the median:
    - If adjustment to median is small (<= 10% change in dimension):
      - **Instruction (Slide {slide_number})**: "Adjust `width` (or `height`) of shape '{id}' to {median_value:.2f}px to improve size consistency within the {row/column}."

---

**3. Table Layout Detection & Grid Alignment**

- **Detect Table-like Structure:** Identify groups of shapes that:
  - Are mostly rectangular or text placeholders.
  - Are arranged in clear rows and columns (consistent `top` and `left` alignments with regular spacing).
  - Have similar `height` within rows and similar `width` within columns.
- If a table-like structure is detected:
  - Identify the boundaries of the "grid" (min/max `left`, min/max `top`, number of rows/columns, average spacing).
  - For shapes within the identified table:
    - **Instruction (Slide {slide_number})**: "Align shapes (ids: [...]) to the detected table grid with column widths {col_widths} and row heights {row_heights}, adjusting `left` and `top` accordingly for precise alignment."

---

**4. Text Alignment Consistency (Metadata-limited)**
- For text boxes in a row/column with similar content:
  - If `textAlign` differs:
    - Suggest aligning to the most frequent `textAlign` with a minor adjustment if needed.
    - **Instruction (Slide {slide_number})**: "Normalize text alignment of shapes (ids: [...]) to '{common_alignment}'."

---

**5. General Principles & Output Format:**
- **Strictly prevent any overflow.**
- Prioritize **proportional resizing** for overflow, with minor shifting as a last resort.
- Maintain **relative positions** within groups.
- Avoid **significant movement** from original positions (limit to ~10px for minor adjustments).
- Be **aware of the surrounding elements** when suggesting changes to avoid creating new issues.
- **Detect and utilize grid alignment** for table-like layouts.
- For every fix, provide clear **Office.js-compatible instructions** including what to change, why, and how (values, scale factors, shift amounts, grid information).
- Format each fix as a clear **natural language instruction** that **includes the Slide Number**.
"""


@app.get("/analyze-slide")
def analyze_slide():
    # image_path = "slide_images/image.png"
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
                    {"text": LAYOUT_ANALYSIS_PROMPT},
                    {
                        "inline_data": {
                            "mime_type": "image/png",
                            "data": base64_img
                        }
                    },
                    {"text": f"Here is the shape metadata (JSON):\n\n{metadata_str}"}
                ]
            }
        ]
    }

    response = requests.post(GEMINI_URL, json=payload)
    if response.status_code != 200:
        return {"error": f"Gemini API Error: {response.text}"}

    try:
        output = response.json()["candidates"][0]["content"]["parts"][0]["text"]
        return {"analysis": output}
    except Exception as e:
        return {"error": f"Invalid response from Gemini: {e}"}


