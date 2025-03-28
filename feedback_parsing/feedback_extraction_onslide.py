import logging
from typing import List, Dict, Optional, Union, Tuple
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL_TYPE
from pptx.dml.color import RGBColor
from pptx.exc import PackageNotFoundError

DEFAULT_ANNOTATION_KEYWORDS = ["feedback:", "note:", "todo:", "comment:"]
DEFAULT_ANNOTATION_COLORS = [
    RGBColor(255, 255, 0),  # Yellow
    RGBColor(255, 192, 0), # Amber/Orange-Yellow
    RGBColor(255, 0, 0),    # Red
]


def _is_annotation_color(shape_fill, target_colors: List[RGBColor]) -> bool:
    if shape_fill.type == MSO_FILL_TYPE.SOLID:
        try:
            if hasattr(shape_fill, 'fore_color') and hasattr(shape_fill.fore_color, 'rgb'):
                if shape_fill.fore_color.rgb is None:
                    return False
                return shape_fill.fore_color.rgb in target_colors
        except AttributeError:
            logging.debug("AttributeError checking shape fill color.")
            return False
    return False


def extract_feedback_from_ppt_onslide(pptx_filepath: str, annotation_keywords: Optional[List[str]] = None, annotation_colors: Optional[List[RGBColor]] = None) -> List[Dict[str, Union[str, int, None, Dict]]]:
    feedback_list: List[Dict[str, Union[str, int, None, Dict]]] = []
    # keywords = annotation_keywords if annotation_keywords is not None else DEFAULT_ANNOTATION_KEYWORDS
    # colors = annotation_colors if annotation_colors is not None else DEFAULT_ANNOTATION_COLORS

    # logging.info(f"Starting on-slide feedback extraction from: {pptx_filepath}")
    # logging.info(f"Using keywords: {keywords}")
    # logging.info(f"Using colors: {[str(c) for c in colors]}") # Log color values

    # try:
    #     prs = Presentation(pptx_filepath)
    #     logging.info(f"Successfully opened presentation. Found {len(prs.slides)} slides.")

    #     for i, slide in enumerate(prs.slides):
    #         slide_number = i + 1
    #         logging.debug(f"Processing slide {slide_number}...")
    #         for shape in slide.shapes:
    #             is_feedback = False
    #             instruction_text = ""

    #             # 1. Check if shape has text and if it starts with a keyword
    #             if shape.has_text_frame and shape.text_frame.text:
    #                 shape_text_lower = shape.text_frame.text.strip().lower()
    #                 for keyword in keywords:
    #                     if shape_text_lower.startswith(keyword):
    #                         instruction_text = shape.text_frame.text.strip()
    #                         is_feedback = True
    #                         logging.debug(f"Slide {slide_number}: Found feedback shape by keyword '{keyword}' (Shape ID: {shape.shape_id})")
    #                         break 
                        
    #             # 2. If not found by keyword, check if shape's fill color matches
    #             if not is_feedback and hasattr(shape, 'fill'):
    #                 if _is_annotation_color(shape.fill, colors):
    #                     if shape.has_text_frame and shape.text_frame.text:
    #                         instruction_text = shape.text_frame.text.strip()
    #                         if instruction_text:
    #                             is_feedback = True
    #                             logging.debug(f"Slide {slide_number}: Found feedback shape by color '{shape.fill.fore_color.rgb}' (Shape ID: {shape.shape_id})")
    #                     else:
    #                         logging.debug(f"Slide {slide_number}: Shape matched color but had no text (Shape ID: {shape.shape_id})")


    #             # 3. If identified as feedback (and has text), add to list
    #             if is_feedback and instruction_text:
    #                 try:
    #                     pos_details = {
    #                         "left": shape.left,
    #                         "top": shape.top,
    #                         "width": shape.width,
    #                         "height": shape.height
    #                     }
    #                 except AttributeError:
    #                     pos_details = {} 

    #                 feedback_item = {
    #                     "source": "on_slide",
    #                     "slide_number": slide_number,
    #                     "instruction": instruction_text,
    #                     "element_details": {
    #                         "shape_id": shape.shape_id,
    #                         "shape_name": shape.name, 
    #                         "text": instruction_text,
    #                         "position": pos_details
    #                     }
    #                 }
    #                 feedback_list.append(feedback_item)

    #         logging.debug(f"Finished processing slide {slide_number}.")

    #     logging.info(f"Finished on-slide extraction. Found {len(feedback_list)} feedback items.")
    # except PackageNotFoundError:
    #     logging.error(f"Error: File not found or not a valid PPTX file: {pptx_filepath}")
    # except Exception as e:
    #     logging.error(f"An unexpected error occurred during on-slide extraction: {e}", exc_info=True)

    return feedback_list
