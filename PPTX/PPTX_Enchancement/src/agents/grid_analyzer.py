# grid_analyzer.py
import base64
import json
import os
import requests

# Define constants directly
LLM_API_KEY = "AIzaSyAv0sTw83EOKcJtoSyT9ug4cnzwGagkMJY" 
GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={LLM_API_KEY}"
IMAGE_PATH = "slide_images/images/presentation/slide_1.png"
METADATA_PATH = "slide_images/metadata.json"
OUTPUT_ANALYSIS_PATH = "slide_images/grid_analysis.json"


SLIDE_GRID_ANALYSIS_PROMPT_TEMPLATE = """
You are an expert assistant specializing in analyzing presentation slide layouts based on visual appearance and shape metadata. Your task is to first determine if the arrangement of shapes forms a recognizable grid structure and identify all participating shapes per row/column. Then, using this grid context and slide dimensions, you must detect specific layout issues: element overflow and inconsistent dimensions among similar shapes within grid rows/columns.
---

**INPUT:**

1.  **Slide Image:** A screenshot of the slide (for visual reference).
2.  **Shape Metadata (JSON):** An array of shape objects, each containing `id`, `type`, `top`, `left`, `width`, `height`, `text`, etc.
3.  **Slide Dimensions:** The overall canvas size (`width` and `height` in pixels). For this task: **{ "width": 960, "height": 540 }**

---

**YOUR TASK:**

**Part 1: Grid Analysis (As before)**

1.  **Analyze Grid Structure:** Examine the `Slide Image` and `Shape Metadata` meticulously to identify patterns of alignment and spacing that form rows and columns.
    *   **Rows:** Look for multiple shapes aligned horizontally (similar `top` or `center Y` coordinates). Consider distinct horizontal bands of elements as potential rows (e.g., a header/timeline row, main content rows, a footer row). **A row exists even if its internal elements don't align perfectly with every column defined above/below, especially if an element spans multiple columns.**
    *   **Columns:** Look for multiple shapes aligned vertically (similar `left` or `center X` coordinates). Consider distinct vertical sections, including potential label columns on the sides. The columns define the overall vertical structure.
    *   **Spanning Elements:** Explicitly look for shapes whose `width` suggests they span the horizontal space of multiple columns defined by other rows, or whose `height` suggests they span multiple rows. **These spanning elements ARE part of the grid.**
    *   **Include ALL Shapes:** Identify **every** shape that participates in the grid alignment for each row and column. This includes:
        *   Container shapes (rectangles, chevrons, etc.).
        *   Text boxes (even if they are *inside* or overlapping other shapes).
        *   Shapes used as icons or markers (like numbered circles, bullet points).
        *   Lines or connectors if they clearly delineate or belong to specific rows/columns.
        *   Labels or annotations aligned with rows or columns (e.g., vertical labels on the left).
        *   **Spanning Shapes:** A shape spanning multiple columns should be included in its corresponding row list and in the list for **each column** it visually occupies. A shape spanning multiple rows should be included in its corresponding column list and in the list for **each row** it visually occupies.
    *   **Determine Grid Size:** Based on the identified rows and columns containing participating shapes, calculate the overall grid dimensions (total number of distinct rows and total number of distinct columns observed across the *entire* structure).

2.  **Determine Grid Presence:** Based on the analysis, conclude if a significant portion of the slide's core content is organized into a recognizable grid structure. The presence of elements spanning multiple rows/columns or imperfect alignment does not automatically disqualify it, as long as a clear row/column pattern exists for major content blocks.

3.  **Extract Grid Structure:** If `is_grid_structure` is true, populate the `grid_structure` field. Create keys like `row_1`, `row_2`, ... and `col_1`, `col_2`, ... corresponding to the identified rows and columns. List the `id`s of all participating shapes under the appropriate keys. A shape can belong to both a row and a column list if applicable (especially corner elements).

**Part 2: Layout Issue Detection**

**Perform these checks ONLY if `is_grid_structure` from Part 1 is true, OR even if false, check all shapes for overflow.**

4.  **Detect Overflow:** Iterate through **all** shapes in the `Shape Metadata`. Using the `Slide Dimensions` (960x540), identify any shape where `left + width > 960` or `top + height > 540`. Record the `id` of each overflowing shape in the `overflowing_shape_ids` list.

5.  **Detect Inconsistent Dimensions within Grid:**
    *   **Perform ONLY if `is_grid_structure` from Part 1 is true.**
    *   Iterate through each identified **row** (e.g., `row_1`, `row_2`, ... from `grid_structure.rows`):
        *   Within that row's list of shapes, identify groups of shapes that share the **same `type`** (e.g., all 'rectangle' shapes in that row, all 'textbox' shapes in that row) AND appear visually intended to be uniform based on their role/positioning within the row.        
        *   For each such group with 2 or more shapes:
            *   Check if their `width` values show significant inconsistency (e.g., > 5px difference). If yes, record this specific inconsistency (dimension: "width").
            *   Check if their `height` values show significant inconsistency (e.g., > 5px difference). If yes, record this specific inconsistency (dimension: "height").

    *   Iterate through each identified **column** (e.g., `col_1`, `col_2`, ... from `grid_structure.columns`):
        *   Perform similar checks: Group shapes by `type` and apparent similar function/role within the column. Check `width` and `height` for significant inconsistencies within groups of 2 or more shapes. Record specifics if found.
    *   **Recording Inconsistencies:** For EACH distinct inconsistency found (e.g., inconsistent width in row 2, inconsistent height in row 2 could be two separate entries), add an object to the `inconsistent_dimensions` list. This object must include:
        *   `location_type`: "row" or "column".
        *   `location_key`: The specific key (e.g., "row_2", "col_1").
        *   `dimension`: The specific dimension found inconsistent ("width" or "height").
        *   `inconsistent_shape_ids`: A list containing the `id`s of **only** the shapes within that specific group exhibiting the inconsistency for that dimension.

**Part 3: Final Review (Crucial Step)**

6.  **Verify Analysis:** Before generating the final output, mentally review your entire analysis (Grid determination, row/column assignments, overflow checks, inconsistency checks). Ensure:
    *   The `is_grid_structure` conclusion is well-supported by the visual and metadata evidence.
    *   The `grid_structure` lists correctly include participating shapes for each row/column identified.
    *   The `overflowing_shape_ids` list is accurate based on coordinates and slide dimensions.
    *   The `inconsistent_dimensions` list correctly identifies the location, specific dimension (`width` or `height`), and the relevant shape IDs for each inconsistency found, based *only* on comparing similar-typed shapes within the defined groups.
    *   The final output strictly adheres to the specified JSON format.

---

**OUTPUT FORMAT:**

Provide your response strictly in JSON format with the following structure:

```json
{
  "grid_analysis": {
    "is_grid_structure": boolean,
    "grid_size": {
      "rows": number,
      "columns": number
    },
    "grid_structure": { // Populated only if is_grid_structure is true
      "rows": {
        "row_1": [id, ...],
        // ...
      },
      "columns": {
        "col_1": [id, ...],
        // ...
      }
    },
    "reasoning": "Brief explanation of grid detection."
  },
  "layout_issues": {
    "overflowing_shape_ids": [
      // List of shape IDs that overflow the slide boundaries (960x540)
      id_overflow1,
      id_overflow2,
      // ... (empty list if none)
    ],
    "inconsistent_dimensions": [
      // List of objects describing inconsistent dimensions within grid rows/columns
      // Populated only if is_grid_structure is true and inconsistencies found
      {
        "location_type": "row", // "row" or "column"
        "location_key": "row_2", // e.g., "row_2" or "col_1"
        "dimension": "height", // "width" or "height"
        "inconsistent_shape_ids": [id1, id2, id3] // IDs of similar shapes in this location with inconsistent dimension
      },
      {
        "location_type": "column",
        "location_key": "col_1",
        "dimension": "width",
        "inconsistent_shape_ids": [id_a, id_b]
      }
      // ... (empty list if none found or not a grid)
    ]
  }
}
```
"""

# Load image and encode as base64
def load_image_base64(image_path):
    """Loads an image file and returns its base64 encoded string."""
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"Image file not found: {image_path}")
    with open(image_path, "rb") as img:
        return base64.b64encode(img.read()).decode("utf-8")

# Load JSON metadata
def load_metadata(metadata_path):
    """Loads JSON data from a file."""
    if not os.path.exists(metadata_path):
        raise FileNotFoundError(f"Metadata file not found: {metadata_path}")
    with open(metadata_path, "r") as f:
        return json.load(f)

# Save JSON data to a file
def save_json_output(data, output_path):
    try:
        # Ensure directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        with open(output_path, "w") as f:
            json.dump(data, f, indent=2)
        print(f"Analysis successfully saved to {output_path}")
        return True
    except Exception as e:
        print(f"Error saving analysis to {output_path}: {e}")
        return False

# Prepare and send the request
def analyze_grid_structure_and_save():
    """
    Analyzes grid structure using hardcoded paths, saves the analysis JSON,
    and returns the parsed analysis data.
    """
    image_base64 = load_image_base64(IMAGE_PATH)
    metadata = load_metadata(METADATA_PATH)

    if image_base64 is None or metadata is None:
        print("Aborting analysis due to missing input files.")
        return None # Indicate failure

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
                        "text": f"\n\nHere is the shape metadata:\n```json\n{json.dumps(metadata, indent=2)}\n```"
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

    print("Sending request to Gemini API for grid analysis...")
    try:
        response = requests.post(GEMINI_URL, json=request_body, timeout=120) # Added timeout
        response.raise_for_status() # Raise HTTPError for bad responses (4xx or 5xx)

        content = response.json()
        try:
            # Extract the text part, remove potential markdown fences
            reply_text = content["candidates"][0]["content"]["parts"][0]["text"]
            reply_text = reply_text.strip().lstrip('```json').rstrip('```').strip() # Clean up common markdown fences

            print("\n--- Gemini Analysis Output Received ---")
            # print(reply_text) # Optionally print raw reply

            # Attempt to parse the cleaned text as JSON
            analysis_data = json.loads(reply_text)
            print("--- Gemini Analysis Parsed Successfully ---")

            # Save the parsed JSON data
            if save_json_output(analysis_data, OUTPUT_ANALYSIS_PATH):
                 return analysis_data # Return parsed data on success
            else:
                 return None # Indicate failure if saving failed

        except (KeyError, IndexError, TypeError) as e:
            print(f"Error: Could not extract valid reply part from Gemini response: {e}")
            print("Full response content:", content)
            return None
        except json.JSONDecodeError as e:
            print(f"Error: Could not parse Gemini response as JSON: {e}")
            print("Raw Gemini reply text was:")
            print(reply_text)
            return None
        except Exception as e:
            print(f"An unexpected error occurred during response processing: {e}")
            print("Full response content:", content)
            return None

    except requests.exceptions.RequestException as e:
        print(f"Error: Request to Gemini API failed: {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred during API request: {e}")
        return None