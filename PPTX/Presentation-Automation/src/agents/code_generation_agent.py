import logging
from config.llmProvider import gemini_flash_llm, initialize_gemini_llm
from config.config import LLM_API_KEY
from google import genai
from config.config import GEMINI_FLASH_2_0_MODEL
import google.api_core.exceptions

client = genai.Client(api_key=LLM_API_KEY)
log = logging.getLogger(__name__)

# ` const slide = context.presentation.slides.getItemAt({slide_index} - 1).shapes;`
        # shapes.load("items/id, items/left, items/top, items/width, items/height, items/textFrame/textRange/font/name, items/textFrame/textRange/font/size, items/textFrame/textRange/font/bold, items/textFrame/textRange/font/italic, items/textFrame/horizontalAlignment");

# === Prompt Template ===
CODE_GEN_PROMPT_TEMPLATE = """
**Role:** You are an expert-level Software Engineer specializing in Office Add-in development using the Office.js JavaScript API for PowerPoint.
**Objective:** Generate a production-quality, robust, and efficient Office.js code snippet to programmatically modify shapes on a PowerPoint slide based on a provided set of natural language instructions. The generated code must be directly usable within a `PowerPoint.run(async (context) => {{ ... }});` block.
**Input:** A list of natural language instructions detailing specific formatting, alignment, resizing, or positioning adjustments for shapes on a slide. Instructions will reference shapes using `(id: <ID>)`.

---

**Strict Requirements & Constraints:**

1. **Office.js API Usage:**
    * Use only the asynchronous Office.js API (`context.sync`, `load`, etc.).
    * Target the first slide (index 0) unless explicitly specified otherwise. `const shapes = context.presentation.slides.getItemAt(0).shapes;`

2. **Shape Identification:**
    * Shape IDs are explicitly denoted using `[ID=<ID>]`. Extract and use the exact ID value.
    * ❗️**STRICT RULE:** Use only the shape IDs provided in the input instructions. Do **not** modify, substitute, or invent any shape IDs. If the instruction references `(id: 463)`, then only use `"463"` exactly — do not change to a different value.
    * Never make assumptions about IDs based on position, context, or inferred relationships.
    * Example: For an instruction like `Adjust width of shape (id: 463)`, the shape must be referenced in code as `shapes.items.find(s => s.id === "463");`.

3. **Code Structure (Strict Format):**

    * **Initialization Block:**
        * Get the shapes collection: `const shapes = context.presentation.slides.getItemAt(0).shapes;`
        * **Load Required Properties:** Load **only** `id`, `name`, `left`, `top`, `width`, and `height` for positioning and resizing tasks.
        * **STRICT RULE:** Do **not** load additional properties such as `textFrame`, `font`, or `alignment` unless explicitly mentioned in the natural language instruction.
        ```javascript
        shapes.load("items/id, items/name, items/left, items/top, items/width, items/height");
        ```
        * Synchronize context: `await context.sync();` // Makes loaded properties available.

    * **Shape Variable Declaration & Validation (AFTER first sync):**
        * Identify **all unique** shape IDs mentioned across all the instructions.
        * For **each unique** ID (e.g., "463", "468", "479"):
            * Declare **one** `const` variable using `shapes.items.find(s => s.id === "<ID>");` (Note: Compare ID as a **string**).
            * Variable names **must** follow the pattern `shape<ID>` (e.g., `shape463`, `shape468`, `shape479`).
            * **Immediately** after each `find`, perform a null/undefined check and log an error if not found.
            ```javascript
            // Example for shape ID "463":
            const shape463 = shapes.items.find(s => s.id === "463");
            if (!shape463) {{
              console.error(`Critical: Shape with ID '463' not found. Skipping related operations.`);
            }}
            // Example for shape ID "468":
            const shape468 = shapes.items.find(s => s.id === "468");
            if (!shape468) {{
              console.error(`Critical: Shape with ID '468' not found. Skipping related operations.`);
            }}
            // Example for shape ID "479":
            const shape479 = shapes.items.find(s => s.id === "479");
            if (!shape479) {{
              console.error(`Critical: Shape with ID '479' not found. Skipping related operations.`);
            }}
            ```

    * **Operation Block (Implement Instructions):**
        * Go through the natural language instructions one by one.
        * For each instruction targeting `(id: <ID>)`:
            * **Wrap the modification logic inside the `if (shape<ID>) {{ ... }}` block** established during validation.
            * Access properties using the declared variable (e.g., `shape463`, `shape468`, `shape479`).
            * **Font Changes:** Use `shape<ID>.textFrame.textRange.font.name = '...';`, `shape<ID>.textFrame.textRange.font.size = ...;`, etc. Ensure the properties were loaded.
            * **Position Changes:** Use `shape<ID>.left = ...;`, `shape<ID>.top = ...;`.
            * **Text Alignment Changes:** Set the `horizontalAlignment` property using **string literals**, mapping precisely:
            * Instruction mentioning "center" -> `"Center"`
            * Instruction mentioning "left" -> `"Left"`
            * Instruction mentioning "right" -> `"Right"`
            * Instruction mentioning "justify" -> `"Justify"`
            * Instruction mentioning "distributed" -> `"Distributed"`
            * **Example:** `shape<ID>.textFrame.horizontalAlignment = "Center";` // *** USE STRING LITERAL ***

    * **Final Synchronization:**
        * `await context.sync();` once at the end.

4. **Execution Rules:**
    * Only use `await context.sync();` twice: once after initial load, once after all modifications.
    * Identify shapes using `find(s => s.id === "<ID>")` *after* the first sync.
    * Never use undeclared shape variables. Do not redeclare variables.

5. **Error Handling:**
    * Use `console.error(...)` for missing shapes — do not halt execution. Operations on missing shapes must be skipped via the `if (shape<ID>)` checks.

6. **Output Formatting:**
    * Output must contain **only** the JavaScript code intended to run inside the `PowerPoint.run` inner block.
    * No wrapper functions (`async function ...`), comments, or extra explanations.
    * Implement **all** input instructions accurately.

---
**Input Instructions:**
---
{{instructions}}
---
**Generated Office.js Code Snippet:**
```javascript
// Office.js compliant code goes here
```

"""

# === Global Initialization ===
_gemini_instance = None

# # Function to generate the code using the configured LangChain client
# def generate_code(refined_instruction: str):
#     if not refined_instruction or not isinstance(refined_instruction, str) or not refined_instruction.strip():
#         log.warning("generate_code received invalid or empty instructions.")
#         return {"error": "Invalid or empty instructions provided."}

#     log.info("generate_code received input from refiner agent:")
#     log.info(f"--- BEGIN INSTRUCTION ---\n{refined_instruction.strip()}\n--- END INSTRUCTION ---")

#     # Prepare the formatted prompt
#     try:
#         prompt = CODE_GEN_PROMPT_TEMPLATE.format(instructions=refined_instruction.strip())
#     except KeyError as e:
#         log.error(f"Failed to format prompt template. Missing key: {e}")
#         return {"error": f"Internal error: Prompt template key missing ({e})."}
#     log.info("Generating Office.js code...")

#     try:
#         global _gemini_instance
#         if _gemini_instance is None:
#             _gemini_instance = initialize_gemini_llm(GEMINI_FLASH_2_0_MODEL)

#         # Send prompt to LLM
#         response = gemini_flash_llm.invoke(prompt)
#         generated_code_str = (
#             response if isinstance(response, str)
#             else getattr(response, "text", getattr(response, "content", "")).strip()
#         )
#         if not generated_code_str:
#             log.warning("Gemini LLM returned empty code response.")
#             return {"error": "No code generated by the LLM."}

#         log.info("Successfully generated Office.js code.")
#         return {"code": generated_code_str}
    
#     except google.api_core.exceptions.ResourceExhausted as e:
#         log.error(f"Google API Quota Exhausted during code generation: {e}")
#         return {"error": f"API Quota Error: {str(e)}"}
#     except google.api_core.exceptions.GoogleAPIError as e:
#         log.error(f"Google API Error during code generation: {e}")
#         return {"error": f"API Error: {str(e)}"}
#     except Exception as e:
#         log.exception("An unexpected error occurred during code generation.") 
#         return {"error": f"Unexpected Exception: {str(e)}"}
    

# Function to generate the code using the configured LangChain client
def generate_code(refined_instruction: str):
    if not refined_instruction or not isinstance(refined_instruction, str) or not refined_instruction.strip():
        log.warning("generate_code received invalid or empty instructions.")
        return {"error": "Invalid or empty instructions provided."}

    log.info("generate_code received input from refiner agent:")
    log.info(f"--- BEGIN INSTRUCTION ---\n{refined_instruction.strip()}\n--- END INSTRUCTION ---")

    # Prepare the formatted prompt
    try:
        prompt = CODE_GEN_PROMPT_TEMPLATE.format(instructions=refined_instruction.strip())
    except KeyError as e:
        log.error(f"Failed to format prompt template. Missing key: {e}")
        return {"error": f"Internal error: Prompt template key missing ({e})."}

    log.info("Generating Office.js code...")

    try:
        global _gemini_instance
        if _gemini_instance is None:
            _gemini_instance = initialize_gemini_llm(GEMINI_FLASH_2_0_MODEL)

        # Use the client.models.generate_content for LLM invocation
        response = client.models.generate_content(
            model="gemini-2.0-flash", 
            contents=[prompt]  # Assuming the prompt is structured as a list of contents
        )
        log.info(f"LLM response: {response.text}")

        # Ensure we extract the correct code from the response
        generated_code_str = response.text.strip() if response.text else ""

        if not generated_code_str:
            log.warning("Gemini LLM returned empty code response.")
            return {"error": "No code generated by the LLM."}

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
