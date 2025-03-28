import json, logging
from typing import List, Dict, Any, Optional

from langchain.prompts import PromptTemplate

from config.llm_provider import gemini_flash_llm 

FEEDBACK_CLASSIFICATION_PROMPT = """

    You are an advanced AI assistant and an expert specializing in analyzing feedback for PowerPoint presentations.
    Your task is to interpret a user's feedback instruction, classifying it accurately into one of three categories:
    - **formatting**: Tasks involving text adjustments, table modifications, alignment, or branding (e.g., font changes, table resizing, slide merging).
    - **cleanup**: Tasks improving structural clarity and consistency (e.g., spacing adjustments, bullet point formatting, splitting tables across slides, remove unneccessary text and elements).
    - **visual enhancement**: Tasks enhancing design aesthetics and visual communication beyond basic formatting(e.g., adding colors, icons, effects, timelines, or backgrounds).

    **Input Instruction:**
    Slide Number: {slide_number}
    Source: {source}
    Instruction: {instruction_text}

    **Category Definitions & Examples:**

    1.  **formatting:** Changes related to the appearance and layout of existing elements, text styles, colors, alignment, and basic structure within a slide or across slides. Often involves applying brand guidelines.
        *   Keywords: font, size, color, bold, align, spacing, position, merge, resize, style, theme, template, layout, table format, graph format, borders, bullets.
        *   Examples:
            *   "Change title font to Arial 12pt." -> category: formatting, tasks: [{{"action": "change_font", "target_element_hint": "title", "params": {{"font_name": "Arial", "size": 12}}}}]
            *   "Align the three boxes at the top." -> category: formatting, tasks: [{{"action": "align_elements", "target_element_hint": "three boxes", "params": {{"alignment": "top"}}}}]
            *   "Apply the standard company template." -> category: formatting, tasks: [{{"action": "apply_template", "target_element_hint": null, "params": {{"template_name": "standard_company"}}}}]
            *   "Format the table to fit the slide and highlight alternate rows." -> category: formatting, tasks: [{{"action": "resize_table", "target_element_hint": "table", "params": {{"constraint": "fit_slide"}}}}, {{"action": "highlight_table_rows", "target_element_hint": "table", "params": {{"style": "alternate"}}}}]
            *   "Merge slides 5 and 6 for side-by-side comparison." -> category: formatting, tasks: [{{"action": "merge_slides", "target_element_hint": null, "params": {{"source_slides": [5, 6], "mode": "comparison"}}}}]

    2.  **cleanup:** Actions focused on improving structure, consistency, clarity, and removing redundancy. Often involves fixing inconsistencies or reorganizing content logically.
        *   Keywords: consistent, uniform, spacing, remove, delete, redundant, structure, organize, split, divide, standardise, fix overlap, center, clean up, adjust layout.
        *   Examples:
            *   "Ensure consistent font sizing across all body text." -> category: cleanup, tasks: [{{"action": "standardize_font_size", "target_element_hint": "all body text", "params": {{"consistency_scope": "presentation"}}}}]
            *   "Remove the extra bullet points on slide 3." -> category: cleanup, tasks: [{{"action": "remove_elements", "target_element_hint": "extra bullet points", "params": {{}}}}]
            *   "Split the large table on slide 8 across two slides." -> category: cleanup, tasks: [{{"action": "split_table_across_slides", "target_element_hint": "large table", "params": {{"lock_header": true}}}}]
            *   "Fix the spacing between the icons and text." -> category: cleanup, tasks: [{{"action": "adjust_spacing", "target_element_hint": "icons and text", "params": {{"consistency": "uniform"}}}}]
            *   "Center the text box within the blue background shape." -> category: cleanup, tasks: [{{"action": "center_element_relative", "target_element_hint": "text box", "params": {{"relative_to_hint": "blue background shape"}}}}]

    3.  **visual_enhancement:** Actions focused on improving the aesthetic appeal, engagement, and visual communication beyond basic formatting. Often involves adding graphical elements, effects, or redesigning components.
        *   Keywords: visual, icon, image, background, color scheme, contrast, effect, shadow, gradient, timeline, diagram, layout options, make engaging, improve look, add graphic.
        *   Examples:
            *   "Add relevant icons next to each bullet point." -> category: visual_enhancement, tasks: [{{"action": "insert_icons_contextual", "target_element_hint": "each bullet point", "params": {{}}}}]
            *   "Change the background to something more professional." -> category: visual_enhancement, tasks: [{{"action": "replace_background", "target_element_hint": null, "params": {{"style_hint": "professional"}}}}]
            *   "Make this timeline curved instead of straight." -> category: visual_enhancement, tasks: [{{"action": "redesign_timeline", "target_element_hint": "timeline", "params": {{"style": "curved"}}}}]
            *   "Apply a subtle shadow effect to the main boxes." -> category: visual_enhancement, tasks: [{{"action": "apply_effect", "target_element_hint": "main boxes", "params": {{"effect_type": "shadow", "intensity": "subtle"}}}}]
            *   "Suggest 2 alternative layouts for this slide." -> category: visual_enhancement, tasks: [{{"action": "generate_alternative_layouts", "target_element_hint": null, "params": {{"count": 2}}}}]


    **Important Rules:**
    - Categorize into ONLY ONE primary category, even if the instruction touches on multiple aspects. Choose the most dominant theme.
    - Break down complex instructions into multiple atomic tasks within the `tasks` list.
    - Extract parameters directly mentioned (e.g., "Arial", 12). If not mentioned, use the instruction text itself as a hint (e.g., "professional", "fit_slide").
    - Do NOT invent details not present in the instruction. Use `null` or descriptive hints based on the text.
    - Ensure the output is valid JSON.

    **Output Requirements:**
    Produce a JSON object containing the following fields:
    - "category": (String) ONE of the following predefined categories: "formatting", "cleanup", "visual_enhancement".
    - "slide_number": (Integer or Null) The slide number the instruction applies to (if specified in the input). Use the provided slide number. If none provided or applicable to the whole presentation, use null.
    - "original_instruction": (String) The original user instruction text.
    - "tasks": (List of Objects) A list of specific, atomic actions needed to fulfill the instruction. Each object in the list should have:
        - "action": (String) A concise verb phrase describing the specific programmatic action (e.g., "change_font", "resize_table", "align_elements", "apply_color_scheme", "insert_icon", "split_table", "add_shadow_effect", "reorder_rows"). Use snake_case.
        - "target_element_hint": (String or Null) A hint describing the element(s) the action applies to, based *only* on the instruction text (e.g., "title", "chart on the left", "all text boxes", "table", "logo", "the diagram"). If not specified or applies broadly, use null. *Do not invent element IDs.*
        - "params": (Object) A dictionary of parameters required for the action (e.g., {{"font_name": "Arial", "size": 12}}, {{"width_inches": 5, "height_inches": 3}}, {{"alignment": "top"}}, {{"color_rgb": [255, 0, 0]}}, {{"icon_name": "checkmark"}}, {{"target_slides": [6, 7]}}, {{"effect": "subtle_shadow"}}).

"""

def parse_feedback_instruction(instruction_item: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    slide_num = instruction_item.get("slide_number")
    source = instruction_item.get("source", "unknown")
    instruction_text = instruction_item.get("instruction", "")

    if not instruction_text:
        logging.warning("Skipping empty instruction.")
        return None

    prompt = PromptTemplate(template=FEEDBACK_CLASSIFICATION_PROMPT, input_variables=["slide_number", "source", "instruction_text"])
    try:
        chain = prompt | gemini_flash_llm
        response = chain.invoke({
            "slide_number": slide_num,
            "source": source,
            "instruction_text": instruction_text
        })
        
        raw_response_text = response.strip()
        logging.debug(f"LLM Raw Response: {raw_response_text}")

        json_match = None
        try:
            start_index = raw_response_text.find('{')
            end_index = raw_response_text.rfind('}')
            if start_index != -1 and end_index != -1 and end_index > start_index:
                json_string = raw_response_text[start_index : end_index + 1]
                parsed_json = json.loads(json_string)
                json_match = parsed_json 
            else:
                logging.warning(f"Could not find valid JSON markers '{{' and '}}' in response for instruction: '{instruction_text}'")
        except json.JSONDecodeError as json_err:
            logging.error(f"Failed to decode JSON response from LLM for instruction: '{instruction_text}'. Error: {json_err}. Response text: {raw_response_text}")
            return None 

        if not json_match:
            logging.error(f"No valid JSON object could be extracted from LLM response for instruction: '{instruction_text}'")
            return None

        if not all(k in json_match for k in ["category", "slide_number", "original_instruction", "tasks"]):
            logging.error(f"Parsed JSON missing required keys for instruction: '{instruction_text}'. Parsed: {json_match}")
            return None
        if not isinstance(json_match["tasks"], list):
            logging.error(f"'tasks' field is not a list in parsed JSON for instruction: '{instruction_text}'. Parsed: {json_match}")
            return None

        logging.info(f"Successfully parsed instruction: '{instruction_text[:50]}...' -> Category: {json_match.get('category')}, Tasks: {len(json_match.get('tasks', []))}")
        return json_match

    except Exception as e:
        logging.error(f"Error processing parse_feedback_instruction for instruction '{instruction_text}': {e}", exc_info=True)
        return None



def classify_feedback_instructions(feedback_list: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    categorized_tasks = []
    for feedback in feedback_list:
        classified_task = parse_feedback_instruction(feedback)
        if classified_task:
            categorized_tasks.append(classified_task)
    logging.info(f"Finished feedback instruction classification. Successfully categorized {len(categorized_tasks)} instructions.")
    return categorized_tasks

