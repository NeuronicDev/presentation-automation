# code_generation_agent.py
import asyncio
import logging, re, aiofiles, os, json
from typing import List 
from config.config import LLM_API_KEY
from google import genai
import google.api_core.exceptions

client = genai.Client(api_key=LLM_API_KEY)
log = logging.getLogger(__name__)

REFINED_TASKS_DIR = "uploaded_pptx/slide_images/presentation"

# --- Helper to load refined instructions from file ---
async def _load_refined_tasks_from_file(slide_number: int) -> List[str]:
    file_path = os.path.join(REFINED_TASKS_DIR, f"slide{slide_number}_refined_tasks.json")
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Refined tasks file not found for slide {slide_number}")
    try:
        async with aiofiles.open(file_path, "r", encoding="utf-8") as f:
            data = json.loads(await f.read())
        if isinstance(data, dict) and "refined_instructions" in data and isinstance(data["refined_instructions"], list):
            instructions = data["refined_instructions"]
            if all(isinstance(item, str) for item in instructions):
                if not instructions:
                    log.warning(f"Refined tasks file {file_path} contained an empty instruction list.")
                return instructions
            else:
                raise ValueError("Invalid format: refined_instructions must be a list of strings.")
        else:
            raise ValueError("Invalid format in refined tasks file.")
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON in refined tasks file for slide {slide_number}") from e
    except Exception as e:
        log.error(f"Error reading refined tasks file {file_path}: {e}", exc_info=True)
        raise 

# --- Prompt Template ---
CODE_GEN_PROMPT_TEMPLATE = """
**Role:** You are an expert-level Software Engineer generating precise Office.js code snippets. Your primary goals are ACCURACY and ADHERENCE to instructions.
**Objective:** Generate an Office.js code snippet to execute ONLY the provided Input Instructions on the specified `target_slide_index`. The code runs inside `PowerPoint.run(async (context) => {{ ... }});`. Output ONLY the inner code block.

**CRITICAL RULES FOR SHAPE IDs:**
1.  You **MUST** identify shapes **ONLY** using the exact IDs provided within `(id: <ID>)` or `(ids: [<ID1>, <ID2>...])` markers in the **Input Instructions** section below.
2.  You **MUST NOT** use, substitute, invent, or infer any shape ID that is not explicitly present in those markers in the Input Instructions.
3.  Every shape reference in your generated code (e.g., `shape123`, `s.id === "123"`) **MUST** correspond directly to an ID found in the Input Instructions.

---
**Target Slide Index (0-based):** {target_slide_index}
---
**Input Instructions:**
(Instructions specify actions and target shapes using `(id: ...)` or `(ids: ...)` markers. Adhere strictly to these.)
---
{instructions}
---

**Code Generation Steps & Requirements:**

1.  **Analyze Input Instructions:** Carefully read ALL Input Instructions to determine the **complete set of properties** that need to be loaded.
2.  **Get Slide & Shapes:**
    ```javascript
    const slide = context.presentation.slides.getItemAt({target_slide_index});
    const shapes = slide.shapes;
    ```
3.  **Generate `shapes.load()` Call:** Create the `shapes.load(...)` string.
    ```javascript
    
    shapes.load("items/id, items/left, items/top, items/width, items/height, items/type");
    await context.sync(); // Sync 1

    shapes.items.forEach((shape) => {{
    if (shape.type === "ShapeTypeAutoShape" || shape.type === "ShapeTypeTextBox") {{
        shape.textFrame.load("textRange/font/name, textRange/font/size");
    }}
    }});

    ```
4.  **First Sync:** Add `await context.sync();` // Sync 1
5.  **Declare & Validate ALL Shape Variables (AFTER potential Sync 2):**
    *   Use the `uniqueShapeIds` Set identified earlier.
    *   **Generate a single block of code where ALL required shape variables are declared consecutively using `const`.**
    *   For **each unique ID** in `uniqueShapeIds`:
        *   Declare `const shape<ID> = shapes.items.find(s => s.id === "<ID>");` (ID as **string**). Use the results loaded from Sync 1 (or Sync 2 if it happened).
        *   **Immediately** add `if (!shape<ID>) {{ console.error(...) }}` check.
    *   Example Block Structure:
        ```javascript
        // --- Declare ALL required shape variables ---
        const shape463 = shapes.items.find(s => s.id === "463");
        if (!shape463) {{ console.error(`Critical: Shape with ID '463' not found.`); }}
        const shape468 = shapes.items.find(s => s.id === "468");
        if (!shape468) {{ console.error(`Critical: Shape with ID '468' not found.`); }}
        // ... continue for ALL unique IDs mentioned in instructions ...
        ```
6.  **Implement ALL Instructions (Operation Block):**
    *   Iterate through **every single** Input Instruction provided above.
    *   Generate the corresponding Office.js action.
    *   Use the **exact** `shape<ID>` variables declared in Step 5.
    *   Wrap actions in `if (shape<ID>) {{ ... }}` guards. Use defensive checks like `if (shape123 && shape123.textFrame)` before accessing text properties.
    *   Use correct property access and string literals for alignment (`"Center"`, `"Left"`).
7.  **Final Sync:** Add `await context.sync();` at the very end. //third sync
8.  **MANDATORY Self-Verification:** Before finishing, perform these checks:
    *   **Instruction Coverage:** Logic exists for EVERY Input Instruction?
    *   **ID Accuracy:** EVERY ID used (`shape<ID>`, `s.id === "<ID>"`) EXACTLY matches an ID from Input Instructions?
    *   **No Extraneous IDs:** NO IDs used that were NOT in Input Instructions?
    *   **Structure:** Correct slide index used? Correct properties loaded based on Step 1 analysis? Declarations AFTER syncs? Correct number of syncs (2 or 3 depending on `needsTextProps`)? `if` guards used?
    *   **If ANY check fails, output ONLY:** `// Error: Code generation verification failed.`
9.  **Output Formatting:** If verification passes, output **only** the generated JavaScript code block. No fences (````javascript`), no explanations, no comments (except `console.error` and the verification error).

---
**Input Instructions (for slide {target_slide_index}):**
---
{{instructions}}
---
**Generated Office.js Code Snippet (for slide {target_slide_index}):**
```javascript
// Office.js compliant code goes here
```

"""

async def generate_code(target_slide_index: int):
    """
    Loads refined instructions from file and generates Office.js code
    for a specific slide index. Now an async function.
    """
    if client is None:
        return {"error": "LLM client not available."}

    log.info(f"--- Generating code for slide index: {target_slide_index} ---")

    # Load refined instructions
    refined_instructions: List[str] = []
    try:
        refined_instructions = await _load_refined_tasks_from_file(target_slide_index)
        if not refined_instructions:
            log.warning(f"No refined instructions loaded for slide {target_slide_index}. Returning empty executable code.")
            return {"code": f"// No instructions to execute for slide {target_slide_index}."}

    except (FileNotFoundError, ValueError) as e:
        log.error(f"Failed to load refined instructions for slide {target_slide_index}: {e}")
        return {"error": f"Failed to load necessary instructions: {e}"}
    except Exception as e:
        log.error(f"Unexpected error loading instructions for slide {target_slide_index}: {e}", exc_info=True)
        return {"error": f"Unexpected error loading instructions: {e}"}

    # Prepare prompt input
    instructions_string = "\n".join(refined_instructions)
    log.debug(f"Formatted Instructions string for code gen prompt (slide {target_slide_index}):\n{instructions_string}")

    try:
        # Format the prompt
        prompt = CODE_GEN_PROMPT_TEMPLATE.format(
            instructions=instructions_string,
            target_slide_index=target_slide_index
        )
    except KeyError as e:
        log.error(f"Failed to format code gen prompt template for slide {target_slide_index}: {e}")
        return {"error": f"Internal error: Prompt template key missing ({e})."}

    try:
        def sync_llm_call(model_name, contents_list):
            try:
                return client.models.generate_content(model=model_name, contents=contents_list) 
            except Exception as llm_e:
                log.error(f"Sync LLM call failed: {llm_e}")
                raise llm_e

        log.info(f"Calling LLM for code generation (slide {target_slide_index})...")
        response = await asyncio.to_thread(sync_llm_call, "gemini-2.0-flash", [prompt]) 

        generated_code_str = response.text.strip() if response.text else ""
        log.info(f"Code Gen LLM Raw Response (slide {target_slide_index}): {generated_code_str}...")

        # Process response
        if not generated_code_str:
            return {"error": "No code generated by the LLM."}

        # Check for explicit verification failure message from LLM
        if generated_code_str.strip().startswith("// Error: Code generation verification failed"):
            log.error(f"LLM verification failed for slide {target_slide_index}: {generated_code_str}")
            return {"error": generated_code_str} 
        code_match = re.search(r'```javascript\s*([\s\S]*?)\s*```', generated_code_str, re.IGNORECASE)
        if code_match:
            clean_code = code_match.group(1).strip()
        else:
            code_match_alt = re.search(r'```\s*([\s\S]*?)\s*```', generated_code_str)
            if code_match_alt:
                clean_code = code_match_alt.group(1).strip()
            else:
                clean_code = generated_code_str.replace("// Office.js compliant code goes here", "").strip()
                if not ("context.sync()" in clean_code and ("shapes.load(" in clean_code or "slide.shapes" in clean_code)):
                    log.warning(f"Extracted code for slide {target_slide_index} appears invalid. Raw response used instead. Raw: {generated_code_str}")
                    return {"error": f"Failed to extract valid code structure. Raw response: {generated_code_str}"}

        if not clean_code:
            return {"error": "Failed to extract valid code from LLM response."}

        log.info(f"Successfully generated and cleaned Office.js code for slide {target_slide_index}.")
        return {"code": clean_code}

    # --- Error Handling for LLM call ---
    except google.api_core.exceptions.ResourceExhausted as e:
        return {"error": f"API Quota Error: {str(e)}"}
    except google.api_core.exceptions.GoogleAPIError as e:
        return {"error": f"API Error: {str(e)}"}
    except Exception as e:
        return {"error": f"Unexpected Code Gen Exception: {str(e)}"}
    