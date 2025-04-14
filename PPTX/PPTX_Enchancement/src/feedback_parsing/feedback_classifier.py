import json
from langchain.prompts import PromptTemplate
from config.llmProvider import gemini_flash_llm

instruction_classification_prompt = PromptTemplate(
    input_variables=["instruction", "slide_index"],
    template="""
You are an advanced AI assistant and an expert specializing in analyzing user instructions for PowerPoint slide editing.

Given a user's instruction, classify and extract structured information.

Respond in JSON format as:
{{
  "task": "add_shape" | "change_font_size" | "cleanup_slide" | "unknown",
  "shape": "circle" | "rectangle" | null,
  "position": "center" | "top-left" | "bottom-right" | null,
  "insert_title": true | false | null,
  "font_size": integer or null,
  "slide_number": integer or null,
  "original_instruction": "{instruction}"
}}

Context:
- If the user says "clean up the ppt", apply to all slides.
- If they say "clean up the slide", apply to current slide (slide number: {slide_index}).
- If slide number is mentioned explicitly, respect that.

Now process the following instruction:
"{instruction}"
"""
)

classification_chain = instruction_classification_prompt | gemini_flash_llm

def classify_instruction(instruction: str, slide_number: int = None) -> dict:
    try:
        # Use `slide_number` as slide_index input
        result = classification_chain.invoke({
            "instruction": instruction,
            "slide_index": slide_number or 1 
        })
        parsed = json.loads(result)

        if not parsed.get("slide_number") and slide_number is not None:
            parsed["slide_number"] = slide_number

        return parsed

    except Exception as e:
        return {
            "task": "unknown",
            "error": str(e),
            "original_instruction": instruction
        }
