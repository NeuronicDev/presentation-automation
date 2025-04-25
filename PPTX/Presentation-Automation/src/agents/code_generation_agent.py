# code_generation_agent.py
import asyncio
import logging, re, aiofiles, os, json
from typing import List, Dict, Any 
from config.llmProvider import gemini_flash_llm, initialize_gemini_llm
from config.config import LLM_API_KEY, GEMINI_FLASH_2_0_MODEL
from google import genai
import google.api_core.exceptions

client = genai.Client(api_key=LLM_API_KEY)
log = logging.getLogger(__name__)

REFINED_TASKS_DIR = "uploaded_pptx/slide_images/presentation"

# --- Revised Prompt Template ---
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
2.  **Get Slide & Shapes:** Start with:
    ```javascript
    const slide = context.presentation.slides.getItemAt({target_slide_index});
    const shapes = slide.shapes;
    ```
3.  **Generate `shapes.load()` Call:** Create the `shapes.load(...)` string.
    *   **ALWAYS** include: `"items/id, items/name, items/left, items/top, items/width, items/height"`.
    *   **CONDITIONALLY ADD based on analysis of Input Instructions:**
        *   If *any* instruction mentions 'font', 'size', 'bold', or 'italic': ADD `, items/textFrame/textRange/font`.
        *   If *any* instruction mentions 'align' (for text), 'center', 'justify', 'distributed': ADD `, items/textFrame/horizontalAlignment`.
        *   If *any* instruction mentions changing text content itself: ADD `, items/textFrame/textRange/text`.
    *   Construct the final, correct load string. Example: `shapes.load("items/id, items/name, items/left, items/top, items/width, items/height, items/textFrame/textRange/font");`
4.  **First Sync:** Add `await context.sync();`
5.  **Declare Shape Variables:**
    *   Identify **all unique shape IDs** mentioned within `(id: ...)` or `(ids: ...)` in the Input Instructions.
    *   For each unique ID, declare **one** `const shape<ID>` variable using `shapes.items.find(s => s.id === "<ID>");` (ID as **string**).
    *   Immediately add the `if (!shape<ID>) {{ console.error(...) }}` check after each declaration.
6.  **Implement ALL Instructions:**
    *   Iterate through **every single** Input Instruction.
    *   Generate the corresponding Office.js action.
    *   Use the **exact** `shape<ID>` variables declared previously.
    *   Wrap actions in `if (shape<ID>) {{ ... }}` guards.
    *   Use correct property access (e.g., `shape123.left`, `shape123.textFrame.textRange.font.size`).
    *   Use exact string literals for alignment (`"Center"`, `"Left"`, etc.).
7.  **Final Sync:** Add `await context.sync();` at the end.

8.  **MANDATORY Self-Verification:** Before finishing, perform these checks:
    *   **Instruction Coverage:** Does my generated code contain logic implementing **EVERY** instruction listed in the "Input Instructions" section?
    *   **ID Accuracy:** Does **EVERY** shape ID used in my code (`shape<ID>`, `s.id === "<ID>"`) **exactly match** an ID that was present in the original "Input Instructions"?
    *   **No Extraneous IDs:** Did I avoid using **ANY** shape ID that was not explicitly in the "Input Instructions"?
    *   **Structure:** Does my code follow the exact structure (slide/shapes retrieval, dynamic load, sync, declarations+checks, operations+checks, final sync)?
    *   **If ANY verification check fails, you MUST output ONLY the following error message:**
        `// Error: Code generation verification failed. Review instruction coverage and ID usage.`

9.  **Output Formatting:** If verification passes, output **only** the generated JavaScript code block. No ```javascript` fences, no explanations, no comments (except `console.error`).

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

# === Global Initialization ===
_gemini_instance = None

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

async def generate_code(target_slide_index: int):
    """
    Loads refined instructions from file and generates Office.js code
    for a specific slide index. Now an async function.
    """
    if client is None:
        log.error("Code Gen: LLM Client not initialized.")
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
            log.warning(f"LLM returned empty code response for slide {target_slide_index}.")
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
            log.warning(f"Code extraction resulted in empty string for slide {target_slide_index}. Raw: {generated_code_str}")
            return {"error": "Failed to extract valid code from LLM response."}

        log.info(f"Successfully generated and cleaned Office.js code for slide {target_slide_index}.")
        return {"code": clean_code}

    # --- Error Handling for LLM call ---
    except google.api_core.exceptions.ResourceExhausted as e:
        log.error(f"Google API Quota Exhausted during code generation for slide {target_slide_index}: {e}")
        return {"error": f"API Quota Error: {str(e)}"}
    except google.api_core.exceptions.GoogleAPIError as e:
        log.error(f"Google API Error during code generation for slide {target_slide_index}: {e}")
        return {"error": f"API Error: {str(e)}"}
    except Exception as e:
        log.exception(f"An unexpected error occurred during code generation call/processing for slide {target_slide_index}.")
        return {"error": f"Unexpected Code Gen Exception: {str(e)}"}
    