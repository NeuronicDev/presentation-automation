
import logging, re
from config.llmProvider import gemini_flash_llm, initialize_gemini_llm
from config.config import GEMINI_FLASH_2_0_MODEL, LLM_API_KEY
from google import genai
import google.api_core.exceptions
log = logging.getLogger(__name__)

# === Prompt Template ===
CODE_GEN_PROMPT_TEMPLATE = """
**Role:** You are an expert-level Software Engineer specializing in Office Add-in development using the Office.js JavaScript API for PowerPoint.

**Objective:** Generate a production-quality, robust, and efficient Office.js code snippet to programmatically modify shapes on a PowerPoint slide based on a provided set of natural language instructions. The generated code must be directly usable within a `PowerPoint.run(async (context) => {{ ... }});` block.

**Input:** A list of natural language instructions detailing specific formatting, alignment, resizing, or positioning adjustments for shapes on a slide.

---

**Strict Requirements & Constraints:**

1.  **Office.js API Usage:**
    * Use only the asynchronous Office.js API (`context.sync`, `load`, etc.).
    * Target the correct slide using:  
      `const shapes = context.presentation.slides.getItemAt({{slideIndex}}).shapes;`  
      *(Replace `{{slideIndex}}` with `0` unless `pX` specifies otherwise.)*

2.  **Shape Identification:**
    *   Identify shapes strictly by their `name` property only.
    *   Names follow the pattern `"Google Shape;ID;pX"` where:
        - `ID` is a unique numeric identifier.
        - `pX` is the one-based slide index (e.g., `p1`, `p2`).
    *   If `pX` is not explicitly mentioned, **assume `p1`**.     
    *   **Never** rely on default shape names (e.g., "Rectangle 5").  

3.  **Code Structure (Strict Format):**
    **Initialization Block:**
    ```javascript
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    shapes.load("items/name, items/left, items/top, items/width, items/height");
    await context.sync();
    ```
    **Shape Declarations — declare first, before any operations:**
    * For every unique shape ID (extracted from `Google Shape;ID;pX`), declare once:
    ```javascript
    const shape<ID> = shapes.items.find(s => s.name === "Google Shape;<ID>;<pX>");
    if (!shape<ID>) {{
      console.error("Critical: Shape with ID <ID> not found on slide <pX>. Skipping related operations.");
    }}
    ```
    * Declare all shapes **before** using them in any conditionals or updates.

    **Operation Block:**
    * For each shape operation:
    ```javascript
    if (shape<ID>) {{
      // Apply operation (e.g., shape<ID>.top = 100)
    }}
    ```
    * For cross-shape operations:
    ```javascript
    if (shape<ID1> && shape<ID2>) {{
      // e.g., align shape<ID1> to shape<ID2>
    }}
    ```

    **Final Synchronization:**
    ```javascript
    await context.sync();
    ```

4.  **Execution Rules:**
    * Only use `await context.sync();` twice: once after load, once after all modifications.
    * Never use undeclared shape variables.
    * Do not redeclare variables.

5.  **Error Handling:**
    * Use `console.error(...)` for missing shapes — do not halt execution.
    * Never throw or halt execution on missing elements.

6.  **Output Formatting:**
    * Output must contain **only** code inside the `PowerPoint.run` inner block.
    * No wrapper, function headers, or extra explanations.
    * All instructions must be implemented exactly — no omissions or summaries.
    * Ensure proper ordering of declarations **before** usage.
---

**Input Instructions:**
---
{instructions}
---

**Generated Office.js Code Snippet:**
```javascript
// Office.js compliant code goes here
```

"""
# === Global Initialization ===
_gemini_instance = None

# Function to generate the code using the configured LangChain client
def generate_code(analysis_output: str):
    """
    Generates Office.js code based on natural language instructions using the configured Gemini LLM.

    Args:
        analysis_output: A string containing the natural language instructions.

    Returns:
        A dictionary containing either {"code": generated_code_string} on success,
        or {"error": error_message_string} on failure.
    """
    if not analysis_output or not isinstance(analysis_output, str):
        log.warning("generate_code received invalid or empty instructions.")
        return {"error": "Invalid or empty instructions provided."}

    instructions = f"Based on the slide analysis, implement the following adjustments:\n{analysis_output.strip()}"

    try:
        prompt = CODE_GEN_PROMPT_TEMPLATE.format(instructions=instructions)
    except KeyError as e:
        log.error(f"Failed to format prompt template. Missing key: {e}")
        return {"error": f"Internal error: Prompt template key missing ({e})."}
    log.info("Generating Office.js code...")
    try:
        # Check if the _gemini_instance is initialized, otherwise initialize it
        global _gemini_instance
        if _gemini_instance is None:
            _gemini_instance = initialize_gemini_llm(GEMINI_FLASH_2_0_MODEL)

        response = gemini_flash_llm.invoke(prompt)
        raw_generated_code = response if isinstance(response, str) else getattr(response, 'content', '')
        if not raw_generated_code:
            log.warning("LLM generated empty code response.")
            return {"error": "No code generated by the LLM."}

        # Extract the generated text from the response
        generated_code_str = (
            response if isinstance(response, str)
            else getattr(response, "text", getattr(response, "content", "")).strip()
        )
        log.info("Successfully generated Office.js code.")
        return {"code": generated_code_str}
    
    except google.api_core.exceptions.ResourceExhausted as e:
        log.error(f"Google API Quota Exhausted during code generation: {e}")
        return {"error": f"API Quota Error: {str(e)}"}
    except google.api_core.exceptions.GoogleAPIError as e:
        log.error(f"Google API Error during code generation: {e}")
        return {"error": f"API Error: {str(e)}"}
    except Exception as e:
        log.exception("An unexpected error occurred during code generation.") 
        return {"error": f"Unexpected Exception: {str(e)}"}