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

# SLIDE_LAYOUT_OPTIMIZATION_PROMPT_TEMPLATE = """You are an expert assistant for slide layout refinement. Your role is to analyze the visual layout of a presentation slide using a screenshot and shape metadata, then generate clear and actionable natural language formatting instructions to improve the visual consistency, alignment, proportions, and overall structure of the shapes on the slide.

# ---

# **SLIDE DIMENSIONS:** 960px width x 540px height (16:9 aspect ratio)

# ---

# **INPUT:**

# 1.  **Slide Image:** A screenshot of the slide (for visual reference).
# 2.  **Shape Metadata (JSON):** An array of shape objects, each containing:
#     *   id: Unique identifier for the shape
#     *   name: Shape name or label (if any)
#     *   type: Type of shape (e.g., 'rectangle', 'chevron', 'text box')
#     *   top, left, width, height: Position and dimensions of the shape (in pixels)
#     *   zIndex: The stacking order of the shape
#     *   text: The visible text in the shape
#     *   isGrouped: Boolean indicating if shape is part of a group
#     *   overlapsWith: List of shape ids this shape overlaps with
#     *   font: Font information (e.g., name, size, color)

# ---

# **YOUR TASK:**

# Analyze the layout structure using the visual and metadata input, and then:

# 1.  Detect inconsistencies in alignment (top, bottom, left, right, center), size (width, height), spacing, especially among visually related elements.
# 2.  Identify implicit table- or grid-like structures, paying close attention to rows or columns of similar shapes (e.g., rectangular content blocks).
# 3.  Propose minimal, proportional layout improvements to enhance structure and visual appeal:
#     *   Prioritize **uniform dimensions (width and/or height)** for shapes intended to be visually consistent within a row or series (e.g., all steps in a process flow).
#     *   Suggest **adjustments to height or width** to create better visual balance between adjacent rows or sections, or to align elements more effectively.
#     *   Recommend alignment adjustments (e.g., top-aligning items in a row, center-aligning headers over content). Use precise pixel values or reference alignment to other key elements.
#     *   Adjust spacing between elements for consistency.
#     *   Correct inconsistencies in text formatting for similar types of labels or elements.
#     *   Limit positional shifts to what's necessary for alignment or spacing, preserving the overall layout structure.
# 4.  Preserve the original design intent and avoid introducing new layout conflicts or overlaps.

# ---

# **OUTPUT:**

# For each detected improvement, output a **clear, natural language instruction** in the following format. Each instruction must target a specific, actionable change.

# *   **Instruction:** "Align shapes (ids: [102, 103, 104]) into a single row with top alignment at 245px and equal horizontal spacing of 40px between them."
# *   **Instruction:** "Resize all chevron shapes in the top row (ids: [11, 12, 13, 14]) to a uniform width of 150px and uniform height of 55px to create visual consistency."
# *   **Instruction:** "Increase the height of the rectangular content blocks (ids: [21, 22, 23, 24]) to 120px to visually balance them with the top row and provide more content space."

# Each instruction must:
# *   Reference shape `ids` clearly.
# *   Specify the property to modify (e.g., top, left, width, height, font color).
# *   Provide the **exact new value** (e.g., `150px`, `green`, `#38761d`) or describe the alignment/distribution target.
# *   Briefly justify the change in terms of visual consistency, alignment, uniformity, or proportionality if not immediately obvious.

# ---

# **IMPORTANT CONSTRAINTS:**

# *   Focus solely on **formatting and layout adjustments**. Do not suggest content changes.
# *   Prioritize minimal changes that yield significant visual improvement. Avoid large or disruptive redesigns.
# *   Respect existing groupings and element relationships.
# *   Be mindful of surrounding elements to prevent new overlaps or visual imbalance.
# *   Output only the **formatted instruction lines** – no summaries, explanations, or conversational text.

# """


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




# SLIDE_LAYOUT_OPTIMIZATION_PROMPT_TEMPLATE = """
# You are an expert AI assistant for slide layout refinement. Your role is to analyze the visual layout of a presentation slide using a screenshot and shape metadata, then generate clear and actionable natural language formatting instructions to improve the visual consistency, alignment, proportions, and overall structure of the shapes on the slide, **strictly adhering to slide boundaries**.

# ---

# **SLIDE DIMENSIONS:** 960px width x 540px height (16:9 aspect ratio)

# ---

# **INPUT:**

# 1.  **Slide Image:** A screenshot of the slide (for visual reference).
# 2.  **Shape Metadata (JSON):** An array of shape objects, each containing:
#     *   id: Unique identifier for the shape
#     *   name: Shape name or label (if any)
#     *   type: Type of shape (e.g., 'rectangle', 'chevron', 'text box')
#     *   top, left, width, height: Position and dimensions of the shape (in pixels)
#     *   zIndex: The stacking order of the shape
#     *   text: The visible text in the shape
#     *   isGrouped: Boolean indicating if shape is part of a group
#     *   overlapsWith: List of shape ids this shape overlaps with
#     *   font: Font information (e.g., name, size, color)
#     *   *(Include other relevant metadata fields as available, like fill/line colors)*

# ---

# **YOUR TASK:**

# Analyze the layout structure using both visual and metadata input. Then:

# 1.  **Detect Inconsistencies:** Identify misalignments (top, bottom, left, right, center), variations in size (width, height), and uneven spacing, especially among visually related elements (e.g., items in a list, steps in a process, blocks in a row/column).
# 2.  **Identify Implicit Structures:** Recognize table-like or grid-like arrangements, focusing on rows or columns of similar shapes. Note the number of elements and their current total span.
# 3.  **Propose Layout Improvements:** Generate instructions for minimal, proportional adjustments to enhance structure, alignment, and visual appeal, **while rigorously applying the Overflow Prevention Logic below.**

# ---

# **CORE LAYOUT & OVERFLOW PREVENTION LOGIC:**

# When proposing alignment or uniform sizing for **rows** or **columns** of related shapes:

# 1.  **Identify Group & Spacing:** Determine the shapes forming the row/column (get their `ids`) and estimate the desired *consistent* spacing between them (e.g., use the average current spacing if somewhat consistent, or a standard small gap like 10-20px if inconsistent).
# 2.  **Calculate Initial Target Size:** If suggesting uniform size, determine the ideal common width (for rows) or height (for columns) based on visual consistency needs or matching the largest element.
# 3.  **Check for Overflow:**
#     *   **For Rows:** Calculate `Total Required Width = (Number of Shapes * Target Uniform Width) + (Number of Gaps * Desired Spacing)`.
#     *   **For Columns:** Calculate `Total Required Height = (Number of Shapes * Target Uniform Height) + (Number of Gaps * Desired Spacing)`.
#     *   Compare `Total Required Width` against `Slide Width (960px)`.
#     *   Compare `Total Required Height` against `Slide Height (540px)`.
# 4.  **Apply Proportional Resizing (If Overflow Detected):**
#     *   **If `Total Required Width > 960px`:**
#         *   Calculate `Available Width for Shapes = 960px - (Number of Gaps * Desired Spacing)`.
#         *   Calculate `New Uniform Width = Available Width for Shapes / Number of Shapes`. Round reasonably (e.g., nearest pixel or 0.5px).
#         *   Use this `New Uniform Width` in your instruction.
#     *   **If `Total Required Height > 540px`:**
#         *   Calculate `Available Height for Shapes = 540px - (Number of Gaps * Desired Spacing)`.
#         *   Calculate `New Uniform Height = Available Height for Shapes / Number of Shapes`. Round reasonably.
#         *   Use this `New Uniform Height` in your instruction.
# 5.  **Final Instruction:** Formulate the instruction using the *final calculated dimensions* (either the initial target or the reduced proportional size) and the desired alignment/spacing. Explicitly mention *why* a size was reduced if overflow prevention was applied.

# ---

# **STRICT CONSTRAINTS:**

# *   **Boundary Adherence:** The final proposed layout for *any* set of elements **must fit** within the 960x540 pixel slide dimensions. **No instruction should cause overflow.**
# *   **Proportionality:** When reducing size due to overflow, the reduction must be applied **proportionally** (i.e., making all elements in the group the *same smaller size*) to maintain consistency.
# *   **Minimal Change:** Only suggest adjustments necessary for alignment, consistent sizing/spacing, and boundary adherence. Avoid drastic repositioning or resizing unrelated to these goals.
# *   **Preserve Relationships:** Maintain original visual groupings, sequence, and stacking order unless the specific goal is to fix these.
# *   **No Content Changes:** Do not add, remove, or alter shape content (text, core meaning). Focus solely on layout properties (position, size, alignment, spacing).

# ---

# **OUTPUT FORMAT:**

# For each distinct layout improvement, output a **clear, actionable natural language instruction** following these examples:

# *   **Instruction:** "Align shapes (ids: [102, 103, 104]) top edges at 245px and distribute them horizontally with equal spacing of 20px between them."
# *   **Instruction:** "Resize chevron shapes in the top row (ids: [11, 12, 13, 14]) to a uniform width of 180px and uniform height of 55px for visual consistency."
# *   **Instruction:** "Resize step shapes (ids: [21, 22, 23, 24]) to a uniform width of 210px each and set horizontal spacing to 15px; this size was calculated to fit all shapes within the 960px slide width."
# *   **Instruction:** "Adjust shapes in the vertical column (ids: [51, 52, 53]) to a uniform height of 150px each and set vertical spacing to 10px; height adjusted proportionally to prevent exceeding the 540px slide height."
# *   **Instruction:** "Center-align the text horizontally within shapes (ids: [31, 32, 33])."

# Each instruction must:
# -   Reference shape `ids` clearly.
# -   Specify the **property** to modify (e.g., top, left, width, height, horizontalAlignment, spacing).
# -   Provide an **exact value, alignment target, or distribution method**.
# -   **Include justification** if the sizing was adjusted due to overflow prevention logic.

# ---

# **IMPORTANT:**
# Execute the **Overflow Prevention Logic** calculations mentally *before* finalizing any instruction involving uniform sizing of rows/columns. Ensure every instruction respects all constraints.
# Do **not** change content, add/remove shapes, or modify grouping logic.
# Keep all improvements minimal, proportional, and visually justifiable.
# Ensure no instruction causes layout conflicts, visual misalignment, or slide boundary overflows.
# Your output **must be plain English natural language instructions only**. 
# Each instruction must be a full English sentence, clearly referencing shape IDs and layout actions.

# """




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
