# feedback_classifier.py
import json, logging, re
from typing import List, Dict, Any, Optional
from langchain.prompts import PromptTemplate
from config.llmProvider import gemini_flash_llm

FEEDBACK_CLASSIFICATION_PROMPT = """
    You are an advanced AI assistant and an expert specializing in analyzing feedback for PowerPoint presentations.
    Your task is to interpret a user's feedback instruction, classifying it accurately into one of three categories:
    - **formatting**: Tasks involving text adjustments, table modifications, alignment, or branding (e.g., font changes, table resizing, slide merging).
    - **cleanup**: Tasks improving structural clarity and consistency (e.g., spacing adjustments, bullet point formatting, splitting tables across slides, remove unneccessary text and elements).
    - **visual enhancement**: Tasks enhancing design aesthetics and visual communication beyond basic formatting(e.g., adding colors, icons, effects, timelines, or backgrounds).

    **Input Provided:**
    - Current Slide Number (0-based): {slide_number} (Context for 'this slide' references)
    - Total Slides in Presentation: {total_slides}
    - Source: {source} (Where the instruction came from, e.g., 'taskpane')
    - Instruction: {instruction_text} (The user's raw request)

    **Category Definitions & Examples:**

    1.  **formatting:** Focuses on applying consistent styles, adjusting appearance of existing elements, text properties, colors, basic alignment/positioning, table/chart formatting, or applying branding/templates. It modifies *how* existing things look.
        *   Keywords: *font, size, color, bold, italic, underline, align, distribute, spacing, position, order, resize, style, theme, template, layout (apply existing), table format, chart format, borders, bullets (style/indent), margins, merge slides.*
        *   Examples:
            *   "Change title font to Arial 32pt and make it blue." -> category: formatting, tasks: [{{ "action": "change_font", "target_element_hint": "title", "params": {{"font_name": "Arial", "size": 32, "color_hint": "blue"}} }}]
            *   "Align the top edges of the three header boxes." -> category: formatting, tasks: [{{ "action": "align_elements", "target_element_hint": "three header boxes", "params": {{"alignment": "top"}} }}]
            *   "Distribute the icons evenly horizontally." -> category: formatting, tasks: [{{ "action": "distribute_elements", "target_element_hint": "icons", "params": {{"axis": "horizontal", "spacing": "even"}} }}]
            *   "Apply the standard company template." -> category: formatting, tasks: [{{ "action": "apply_template", "target_element_hint": null, "params": {{"template_name": "standard_company"}} }}]
            *   "Reformat the bullet points to use checkmarks." -> category: formatting, tasks: [{{ "action": "change_bullet_style", "target_element_hint": "bullet points", "params": {{"style": "checkmark"}} }}]
            *   "Increase spacing between paragraphs." -> category: formatting, tasks: [{{ "action": "adjust_paragraph_spacing", "target_element_hint": "paragraphs", "params": {{"amount": "increase"}} }}]
            *   "Merge slides 5 and 6 for side-by-side comparison." -> category: formatting, tasks: [{{ "action": "merge_slides", "target_element_hint": null, "params": {{"source_slides": [5, 6], "mode": "comparison"}} }}]

    2.  **cleanup:** Focuses on improving logical structure, clarity, consistency across similar elements, removing redundancy, or fixing layout problems like overlaps or awkward positioning. It reorganizes or standardizes content.
        *   Keywords: *consistent, uniform, standardise, fix overlap, fix spacing, remove, delete, redundant, unnecessary, structure, organize, group, ungroup, split, divide, center, tidy up, clean up, consolidate, adjust layout.*
        *   Examples:
            *   "Ensure consistent font sizing across all body text." -> category: cleanup, tasks: [{{ "action": "standardize_font_size", "target_element_hint": "all body text", "params": {{"consistency_scope": "presentation"}} }}]
            *   "Remove the extra bullet points on slide 3." -> category: cleanup, tasks: [{{ "action": "remove_elements", "target_element_hint": "extra bullet points", "params": {{}} }}]
            *   "Fix the overlapping text boxes." -> category: cleanup, tasks: [{{ "action": "resolve_overlaps", "target_element_hint": "overlapping text boxes", "params": {{"strategy": "adjust_position"}} }}]
            *   "Resize the table to fit the content." -> category: cleanup, tasks: [{{ "action": "resize_table", "target_element_hint": "table", "params": {{"fit": "content"}} }}]
            *   "Split the large table on slide 8 across two slides." -> category: cleanup, tasks: [{{ "action": "split_table_across_slides", "target_element_hint": "large table", "params": {{"lock_header": true}} }}]
            *   "Fix the spacing between the icons and text." -> category: cleanup, tasks: [{{ "action": "adjust_spacing", "target_element_hint": "icons and text", "params": {{"consistency": "uniform"}} }}]
            *   "Center the text box within the blue background shape." -> category: cleanup, tasks: [{{ "action": "center_element_relative", "target_element_hint": "text box", "params": {{"relative_to_hint": "blue background shape"}} }}]
            
    3.  **visual_enhancement:** Focuses on significantly improving aesthetic appeal, adding new graphical elements, changing layouts substantially, applying visual effects, or making content more engaging beyond basic formatting/cleanup. It changes the *design* or adds *new* visual components.
        *   Keywords: *visuals, icon, image (add/replace), background, color scheme, contrast, effect, shadow, gradient, reflection, timeline (create/redesign), diagram (create/redesign), infographic, layout options, make engaging/appealing/professional, improve look, add graphic, data visualization.*
        *   Examples:
            *   "Add appropriate icons for each service listed." -> category: visual_enhancement, tasks: [{{ "action": "insert_icons_contextual", "target_element_hint": "each service listed", "params": {{}} }}]
            *   "Change the slide background to a subtle gradient." -> category: visual_enhancement, tasks: [{{ "action": "change_background", "target_element_hint": null, "params": {{"type": "gradient", "style_hint": "subtle"}} }}]
            *   "Create a process timeline based on the bullet points." -> category: visual_enhancement, tasks: [{{ "action": "create_timeline", "target_element_hint": "bullet points", "params": {{}} }}]
            *   "Convert the list into a SmartArt diagram." -> category: visual_enhancement, tasks: [{{ "action": "convert_to_smartart", "target_element_hint": "list", "params": {{"diagram_type_hint": "process"}} }}]
            *   "Make the photos look more impactful with a frame effect." -> category: visual_enhancement, tasks: [{{ "action": "apply_effect", "target_element_hint": "photos", "params": {{"effect_type": "frame"}} }}]
            *   "Improve the visual hierarchy of this slide." -> category: visual_enhancement, tasks: [{{ "action": "improve_visual_hierarchy", "target_element_hint": null, "params": {{}} }}]
            *   "Suggest 2 alternative layouts for this slide." -> category: visual_enhancement, tasks: [{{ "action": "generate_alternative_layouts", "target_element_hint": null, "params": {{"count": 2}} }}]

    **Handling Vague Instructions:**
    - If the instruction is very general like "cleanup this slide", "format the presentation", "make this look better", "fix slide":
        - Classify it into the most appropriate primary category (cleanup, formatting, visual_enhancement).
        - Determine the correct scope (current_slide, entire_presentation) based on keywords like "this slide" vs "presentation" or "all". Default to "current_slide" if ambiguous.
        - Generate a SINGLE task in the `tasks` list using a generic action name reflecting the category and scope.
        - Example for "cleanup this slide": category: cleanup, instruction_scope: current_slide, tasks: [{{ "action": "general_slide_cleanup", "target_element_hint": null, "params": {{}} }}]
        - Example for "format the whole presentation": category: formatting, instruction_scope: entire_presentation, tasks: [{{ "action": "general_presentation_formatting", "target_element_hint": null, "params": {{}} }}]
        - Example for "improve visuals on this slide": category: visual_enhancement, instruction_scope: current_slide, tasks: [{{ "action": "general_visual_enhancement", "target_element_hint": null, "params": {{}} }}]
    - **Do NOT** attempt to guess specific sub-tasks (like align shapes, change fonts) for these vague instructions at *this classification stage*. Output only the single general task. The next agent in the pipeline will handle expanding these general tasks.

    **Important Rules:**
    - **Primary Category:** Choose ONLY ONE category based on the main goal.
    - **Task Breakdown:** For SPECIFIC instructions, break into multiple atomic tasks if needed (e.g., change font AND resize). For VAGUE instructions, create only ONE general task.
    - **Parameter Extraction:** Extract literal parameters ("Arial", 12). For non-literal parameters, use hints from the text ("professional", "fit_slide"). Use `null` if no target is specified.
    - **Scope Determination:** Identify scope ("current_slide", "specific_slides", "entire_presentation"). Default to "current_slide" if ambiguous. Use keywords like "all", "whole", "entire", "consistent across" for "entire_presentation". Use "slide 3", "slides 2 and 4" for "specific_slides".
    - **Target Indices (0-based):** Determine `target_slide_indices` based on scope:
        - IF scope IS "current_slide": Output list containing ONLY the input `{slide_number}` (0-based). Example: Input 1 -> Output [1].
        - IF scope IS "entire_presentation": Output list containing indices from 0 to `{total_slides}`-1. Example: total_slides 3 -> Output [0, 1, 2].
        - IF scope IS "specific_slides": Parse numbers, convert to 0-based indices. Example: "slides 3 and 5" -> Output [2, 4].
    - **No Invention:** Do not add details or actions not implied by the instruction.
    - **Valid JSON Output:** Ensure the final output strictly adheres to the required JSON format.

    **Output Requirements:**
    Produce ONLY a valid JSON object containing the following fields:
    - "category": (String) "formatting", "cleanup", or "visual_enhancement".
    - "slide_number": (Integer) The *original* 0-based slide index input context value (`{slide_number}`).
    - "original_instruction": (String) The `{instruction_text}`.
    - "instruction_scope": (String) "current_slide", "specific_slides", or "entire_presentation".
    - "target_slide_indices": (List of Integers) List of 0-based indices. MUST follow rules above.
    - "tasks": (List of Objects) List of actions. Each object requires:
        - "action": (String) Concise snake_case action (use examples provided or create similar).
        - "target_element_hint": (String or Null) Hint from instruction text.
        - "params": (Object) Dictionary of parameters. (e.g., {{"font_name": "Arial", "size": 12}}, {{"alignment": "top"}}). Note: JSON requires double quotes.
"""

classification_prompt = PromptTemplate(
    template=FEEDBACK_CLASSIFICATION_PROMPT,
    input_variables=["slide_number", "source", "instruction_text", "total_slides"] 
)

classification_chain = classification_prompt | gemini_flash_llm

# === Feedback Parser ===
def parse_feedback_instruction(instruction_item: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    slide_num = instruction_item.get("slide_number")
    source = instruction_item.get("source", "unknown")
    instruction_text = instruction_item.get("instruction", "")
    total_slides = instruction_item.get("total_slides")

    logging.debug(f"Received instruction: {instruction_text}")
    # Validate total_slides
    if total_slides is None or not isinstance(total_slides, int) or total_slides < 0:
         logging.error(f"Invalid or missing 'total_slides' count ({total_slides}) received for instruction: '{instruction_text}'")
         return None 
    logging.debug(f"Classifying instruction: '{instruction_text}' for slide context: {slide_num}, total slides: {total_slides}")

    if not instruction_text:
        logging.warning("Skipping empty instruction.")
        return None
    
    try:
        response = classification_chain.invoke({
            "slide_number": slide_num,
            "source": source,
            "instruction_text": instruction_text,
            "total_slides": total_slides # Pass the count
        })
        raw_response_text = response.strip()
        logging.info(f"LLM Raw Response for Classification: {raw_response_text}")

        parsed_json = None
        try:
            start_index = raw_response_text.find('{')
            end_index = raw_response_text.rfind('}')
            if start_index != -1 and end_index != -1 and end_index > start_index:
                json_string = raw_response_text[start_index : end_index + 1]
                parsed_json = json.loads(json_string)
            else:
                 match = re.search(r'\{.*\}', raw_response_text, re.DOTALL) 
                 if match:
                      json_string = match.group(0)
                      parsed_json = json.loads(json_string)
        except json.JSONDecodeError as json_err:
            logging.error(f"Failed to decode JSON from LLM response. Error: {json_err}. Raw response: {raw_response_text}")
            return None
        
        if not parsed_json:
            logging.error(f"No valid JSON object could be extracted from LLM response for instruction: '{instruction_text}'")
            return None
        
        # Validate and return parsed JSON
        required_keys = ["category", "slide_number", "original_instruction", "instruction_scope", "target_slide_indices", "tasks"]
        if not all(k in parsed_json for k in required_keys):
            logging.error(f"Parsed JSON missing required keys...")
            return None
    
        if not isinstance(parsed_json.get("tasks"), list):
            logging.error(f"'tasks' field is not a list in parsed JSON for instruction: '{instruction_text}'. Parsed: {parsed_json}")
            return None
        
        if parsed_json.get("instruction_scope") not in ["current_slide", "specific_slides", "entire_presentation"]:
             logging.warning(f"Unexpected instruction_scope value: {parsed_json.get('instruction_scope')}")

        logging.info(f"Successfully parsed: Context Slide {slide_num}, Category: {parsed_json.get('category')}, Scope: {parsed_json.get('instruction_scope')}, Tasks: {len(parsed_json.get('tasks', []))}")
        return parsed_json

    except Exception as e:
        logging.error(f"Error processing instruction '{instruction_text}' in parse_feedback_instruction: {e}", exc_info=True)
        return None
    
# === Batch Classifier ===
def classify_feedback_instructions(feedback_list: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    categorized_tasks = []
    for feedback in feedback_list:
        result = parse_feedback_instruction(feedback)
        if result:
            categorized_tasks.append(result)
    logging.info(f"Batch classification finished. Parsed {len(categorized_tasks)} of {len(feedback_list)} feedback items.")
    return categorized_tasks


################## Modified feedback_classifier ################################## 


# # feedback_classifier.py
# import json, logging, re
# from typing import List, Dict, Any, Optional
# from langchain.prompts import PromptTemplate
# from config.llmProvider import gemini_flash_llm

# FEEDBACK_CLASSIFICATION_PROMPT = """
#     You are an advanced AI assistant expert in analyzing and decomposing feedback for PowerPoint presentations.
#     Your primary task is to interpret a user's instruction, identify the distinct actions requested, determine the scope and target slides for each action, and classify each action into a category.

#     **Input Provided:**
#     - Current Slide Number (0-based): {slide_number} (Context for 'this slide', 'current slide' references)
#     - Total Slides in Presentation: {total_slides}
#     - Source: {source} (Where the instruction came from, e.g., 'taskpane')
#     - Instruction: {instruction_text} (The user's raw request)

#     **Output Goal:** Produce a JSON object containing a list of "operations". Each operation represents a distinct high-level goal derived from the instruction, specifying its category, scope, target slides, and decomposed atomic tasks.

#     **Category Definitions & Examples:**

#     1.  **formatting:** Focuses on applying consistent styles, adjusting appearance of existing elements, text properties, colors, basic alignment/positioning, table/chart formatting, or applying branding/templates. It modifies *how* existing things look.
#         *   Keywords: *font, size, color, bold, italic, underline, align, distribute, spacing, position, order, resize, style, theme, template, layout (apply existing), table format, chart format, borders, bullets (style/indent), margins, merge slides.*
#         *   Examples:
#             *   "Change title font to Arial 32pt and make it blue." -> category: formatting, tasks: [{{ "action": "change_font", "target_element_hint": "title", "params": {{"font_name": "Arial", "size": 32, "color_hint": "blue"}} }}]
#             *   "Align the top edges of the three header boxes." -> category: formatting, tasks: [{{ "action": "align_elements", "target_element_hint": "three header boxes", "params": {{"alignment": "top"}} }}]
#             *   "Distribute the icons evenly horizontally." -> category: formatting, tasks: [{{ "action": "distribute_elements", "target_element_hint": "icons", "params": {{"axis": "horizontal", "spacing": "even"}} }}]
#             *   "Apply the standard company template." -> category: formatting, tasks: [{{ "action": "apply_template", "target_element_hint": null, "params": {{"template_name": "standard_company"}} }}]
#             *   "Reformat the bullet points to use checkmarks." -> category: formatting, tasks: [{{ "action": "change_bullet_style", "target_element_hint": "bullet points", "params": {{"style": "checkmark"}} }}]
#             *   "Increase spacing between paragraphs." -> category: formatting, tasks: [{{ "action": "adjust_paragraph_spacing", "target_element_hint": "paragraphs", "params": {{"amount": "increase"}} }}]
#             *   "Merge slides 5 and 6 for side-by-side comparison." -> category: formatting, tasks: [{{ "action": "merge_slides", "target_element_hint": null, "params": {{"source_slides": [5, 6], "mode": "comparison"}} }}]

#     2.  **cleanup:** Focuses on improving logical structure, clarity, consistency across similar elements, removing redundancy, or fixing layout problems like overlaps or awkward positioning. It reorganizes or standardizes content.
#         *   Keywords: *consistent, uniform, standardise, fix overlap, fix spacing, remove, delete, redundant, unnecessary, structure, organize, group, ungroup, split, divide, center, tidy up, clean up, consolidate, adjust layout.*
#         *   Examples:
#             *   "Ensure consistent font sizing across all body text." -> category: cleanup, tasks: [{{ "action": "standardize_font_size", "target_element_hint": "all body text", "params": {{"consistency_scope": "presentation"}} }}]
#             *   "Remove the extra bullet points on slide 3." -> category: cleanup, tasks: [{{ "action": "remove_elements", "target_element_hint": "extra bullet points", "params": {{}} }}]
#             *   "Fix the overlapping text boxes." -> category: cleanup, tasks: [{{ "action": "resolve_overlaps", "target_element_hint": "overlapping text boxes", "params": {{"strategy": "adjust_position"}} }}]
#             *   "Resize the table to fit the content." -> category: cleanup, tasks: [{{ "action": "resize_table", "target_element_hint": "table", "params": {{"fit": "content"}} }}]
#             *   "Split the large table on slide 8 across two slides." -> category: cleanup, tasks: [{{ "action": "split_table_across_slides", "target_element_hint": "large table", "params": {{"lock_header": true}} }}]
#             *   "Fix the spacing between the icons and text." -> category: cleanup, tasks: [{{ "action": "adjust_spacing", "target_element_hint": "icons and text", "params": {{"consistency": "uniform"}} }}]
#             *   "Center the text box within the blue background shape." -> category: cleanup, tasks: [{{ "action": "center_element_relative", "target_element_hint": "text box", "params": {{"relative_to_hint": "blue background shape"}} }}]
            
#     3.  **visual_enhancement:** Focuses on significantly improving aesthetic appeal, adding new graphical elements, changing layouts substantially, applying visual effects, or making content more engaging beyond basic formatting/cleanup. It changes the *design* or adds *new* visual components.
#         *   Keywords: *visuals, icon, image (add/replace), background, color scheme, contrast, effect, shadow, gradient, reflection, timeline (create/redesign), diagram (create/redesign), infographic, layout options, make engaging/appealing/professional, improve look, add graphic, data visualization.*
#         *   Examples:
#             *   "Add appropriate icons for each service listed." -> category: visual_enhancement, tasks: [{{ "action": "insert_icons_contextual", "target_element_hint": "each service listed", "params": {{}} }}]
#             *   "Change the slide background to a subtle gradient." -> category: visual_enhancement, tasks: [{{ "action": "change_background", "target_element_hint": null, "params": {{"type": "gradient", "style_hint": "subtle"}} }}]
#             *   "Create a process timeline based on the bullet points." -> category: visual_enhancement, tasks: [{{ "action": "create_timeline", "target_element_hint": "bullet points", "params": {{}} }}]
#             *   "Convert the list into a SmartArt diagram." -> category: visual_enhancement, tasks: [{{ "action": "convert_to_smartart", "target_element_hint": "list", "params": {{"diagram_type_hint": "process"}} }}]
#             *   "Make the photos look more impactful with a frame effect." -> category: visual_enhancement, tasks: [{{ "action": "apply_effect", "target_element_hint": "photos", "params": {{"effect_type": "frame"}} }}]
#             *   "Improve the visual hierarchy of this slide." -> category: visual_enhancement, tasks: [{{ "action": "improve_visual_hierarchy", "target_element_hint": null, "params": {{}} }}]
#             *   "Suggest 2 alternative layouts for this slide." -> category: visual_enhancement, tasks: [{{ "action": "generate_alternative_layouts", "target_element_hint": null, "params": {{"count": 2}} }}]

#     **Scope Determination:**
#     - **current_slide**: Refers *only* to the slide index provided in the input (`{{slide_number}}`). Keywords: "this slide", "current slide", or implied by context (e.g., "fix the title" usually means the current slide's title). **This is the default scope if not otherwise specified.**
#     - **specific_slides**: Refers to explicitly mentioned slide numbers (e.g., "slide 3", "slides 2 and 5").
#     - **entire_presentation**: Refers to all slides. Keywords: "all slides", "whole presentation", "entire presentation", "throughout", "consistent across slides", "apply theme/template".

#     **Instruction Analysis & Decomposition:**

#     1.  **Identify Distinct Operations:** Analyze `{instruction_text}`. Look for conjunctions ("and", "also", "then") connecting potentially different actions, categories, or scopes.
#         *   Example: "cleanup this slide **and** format slide 3" -> TWO operations.
#         *   Example: "Align the boxes **and** change their color" -> ONE formatting operation (multiple tasks possible within it).
#         *   Example: "cleanup the presentation" -> ONE cleanup operation.
#     2.  **For EACH Distinct Operation Identified:**
#         *   **Determine Category:** Assign the primary category (formatting, cleanup, visual_enhancement) for this operation.
#         *   **Determine Scope:** Identify the scope (current_slide, specific_slides, entire_presentation) for *this specific operation*. Default to "current_slide" if ambiguous for this part.
#         *   **Determine Target Indices (0-based):** Based on the operation's scope:
#             *   If scope is "current_slide", use the input `{{slide_number}}`. Target Indices: `[{{slide_number}}]`.
#             *   If scope is "entire_presentation", use all indices from 0 to `{{total_slides}}`-1. Target Indices: `[0, 1, ..., {{total_slides}}-1]`.
#             *   If scope is "specific_slides", parse mentioned numbers (assume 1-based if just numbers like "slide 3", convert to 0-based) and validate against `{{total_slides}}`. Target Indices: `[parsed_indices]`.
#         *   **Decompose into Atomic Tasks:**
#             *   If the operation is **vague** (e.g., "cleanup this slide", "make it look better", "format the presentation"), generate **ONLY ONE** general task reflecting the category and scope. Examples: `{{ "action": "general_slide_cleanup", ... }}`, `{{ "action": "general_presentation_formatting", ... }}`, `{{ "action": "general_visual_enhancement", ... }}`. Do NOT guess sub-tasks for vague instructions here.
#             *   If the operation is **vague** (e.g., "cleanup this slide"), generate **ONLY ONE** general task reflecting the category and scope (e.g., `{{ "action": "general_slide_cleanup", ... }}`). Do NOT guess sub-tasks here.

#     **Important Rules:**
#     - Create a separate entry in the "operations" list for each distinct operation identified.
#     - Use 0-based indices for `target_slide_indices`.
#     - For `target_element_hint`, capture the user's descriptive phrase (e.g., "title", "blue boxes", "icons near text"). Do NOT resolve to IDs here. Use `null` if no specific element is mentioned for the operation.
#     - Ensure valid JSON output. No explanations outside the JSON.

#     **Output Requirements:**
#     Produce ONLY a valid JSON object with a single key "operations", which is a list. Each object in the "operations" list represents one distinct operation and MUST contain:
#     - "category": (String) "formatting", "cleanup", or "visual_enhancement".
#     - "instruction_scope": (String) "current_slide", "specific_slides", or "entire_presentation".
#     - "target_slide_indices": (List of Integers) List of 0-based indices this operation applies to.
#     - "tasks": (List of Objects) List of atomic actions for this operation. Each task requires:
#         - "action": (String) Concise snake_case action (use examples or create similar).
#         - "target_element_hint": (String or Null) The user's phrase describing the target (e.g., "title", "header boxes", "chart", null if applies generally to slide/presentation). **This hint is crucial input for later agents.**
#         - "params": (Object) Dictionary of parameters (e.g., {{"font_name": "Arial", "size": 12}}, {{"alignment": "top"}}). JSON requires double quotes.

#     **Example Output Structure:**
#     ```json
#     {{
#     "operations": [
#         {{
#         "category": "cleanup",
#         "instruction_scope": "current_slide",
#         "target_slide_indices": [1],
#         "tasks": [
#             {{
#             "action": "general_slide_cleanup",
#             "target_element_hint": "this slide",
#             "params": {{}}
#             }}
#         ]
#         }},
#         {{
#         "category": "formatting",
#         "instruction_scope": "specific_slides",
#         "target_slide_indices": [4],
#         "tasks": [
#             {{
#             "action": "change_font",
#             "target_element_hint": "title",
#             "params": {{"font_name": "Arial", "size": 24}}
#             }}
#         ]
#         }}
#     ]
#     }}
#     ```
#     **Input for this run:**
#     Instruction: {instruction_text}
#     Current Slide Number (0-based): {slide_number}
#     Total Slides in Presentation: {total_slides}
#     Source: {source}

#     **Generate the JSON output:**
# """

# classification_prompt = PromptTemplate(
#     template=FEEDBACK_CLASSIFICATION_PROMPT,
#     input_variables=["slide_number", "source", "instruction_text", "total_slides"] 
# )

# classification_chain = classification_prompt | gemini_flash_llm

# # === Feedback Parser (Multi-operation Support) ===
# def parse_feedback_instruction(instruction_item: Dict[str, Any]) -> Optional[List[Dict[str, Any]]]:
#     if classification_chain is None:
#         logging.error("Classification chain not initialized. Cannot parse feedback.")
#         return None

#     slide_num = instruction_item.get("slide_number")
#     source = instruction_item.get("source", "unknown")
#     instruction_text = instruction_item.get("instruction", "")
#     total_slides = instruction_item.get("total_slides")

#     logging.debug(f"Received instruction: {instruction_text}")
    
#     if total_slides is None or not isinstance(total_slides, int) or total_slides < 0:
#         logging.error(f"Invalid or missing 'total_slides' count: {total_slides}")
#         return None
    
#     if not instruction_text:
#         logging.warning("Skipping empty instruction.")
#         return None

#     try:
#         response = classification_chain.invoke({
#             "slide_number": slide_num,
#             "source": source,
#             "instruction_text": instruction_text,
#             "total_slides": total_slides
#         })
#         raw_response_text = response.strip()
#         logging.info(f"LLM Raw Response for Classification: {raw_response_text}")

#         # Extract valid JSON
#         parsed_json = None
#         try:
#             start_index = raw_response_text.find('{')
#             end_index = raw_response_text.rfind('}')
#             if start_index != -1 and end_index != -1 and end_index > start_index:
#                 json_string = raw_response_text[start_index:end_index + 1]
#                 parsed_json = json.loads(json_string)
#             else:
#                 match = re.search(r'\{.*\}', raw_response_text, re.DOTALL)
#                 if match:
#                     json_string = match.group(0)
#                     parsed_json = json.loads(json_string)
#         except json.JSONDecodeError as json_err:
#             logging.error(f"Failed to decode JSON: {json_err}. Raw: {raw_response_text}")
#             return None

#         if not parsed_json:
#             logging.error("No valid JSON object extracted from response.")
#             return None

#         operations = parsed_json.get("operations")
#         if not isinstance(operations, list):
#             logging.error("Expected a list of 'operations' in parsed JSON.")
#             return None

#         valid_operations = []
#         for op in operations:
#             if not all(k in op for k in ["category", "slide_number", "original_instruction", "instruction_scope", "target_slide_indices", "tasks"]):
#                 logging.warning(f"Incomplete operation data: {op}")
#                 continue

#             if not isinstance(op["tasks"], list):
#                 logging.warning(f"Tasks field is not a list: {op}")
#                 continue

#             if op["instruction_scope"] not in ["current_slide", "specific_slides", "entire_presentation"]:
#                 logging.warning(f"Unexpected instruction_scope: {op['instruction_scope']}")

#             valid_operations.append(op)

#         if not valid_operations:
#             logging.warning("No valid operations found in parsed response.")
#             return None

#         logging.info(f"Successfully parsed {len(valid_operations)} operations from instruction.")
#         return valid_operations

#     except Exception as e:
#         logging.error(f"Exception while processing instruction '{instruction_text}': {e}", exc_info=True)
#         return None

# # === Batch Classifier (Supports Compound Outputs) ===
# def classify_feedback_instructions(feedback_list: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
#     categorized_tasks = []
#     for feedback in feedback_list:
#         result = parse_feedback_instruction(feedback)
#         if result:
#             categorized_tasks.extend(result) 
#     logging.info(f"Classification complete. Parsed {len(categorized_tasks)} operations from {len(feedback_list)} feedback items.")
#     return categorized_tasks