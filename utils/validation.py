import os, logging, json, time, datetime, asyncio, sys, pathlib, subprocess, tempfile, shutil, re
import logging.handlers
from io import BytesIO
from pptx import Presentation
from typing import Dict, Any, List
from pdf2image import convert_from_path

from utils.utils import convert_pptx_to_pdf
from config.config import LLM_API_KEY

from google import genai
client = genai.Client(api_key=LLM_API_KEY)


def pptx_to_images(pptx_path, output_dir):
    pdf_path = convert_pptx_to_pdf(pptx_path, output_dir)
    images = convert_from_path(pdf_path, dpi=300)
    if not images:
        raise RuntimeError(f"No images generated from PDF {pdf_path}")
    logging.info(f"Successfully converted {len(images)} slides to images")
    return images

VALIDATION_PROMPT = """
    # ROLE: AI PowerPoint Quality Assurance Analyst (Visual & Task Validation)

    You are an expert state-of-the-art AI assistant specializing in Quality Assurance for automated PowerPoint slide modifications. 
    Your task is to meticulously compare a 'Before' and 'After' image of a specific PowerPoint slide and validate whether the intended modifications ('Tasks Applied') were executed correctly, while also identifying any unintended visual regressions.

    ## Input Provided:
        **Tasks Applied:** A list of specific modification tasks intended for the original input slide
        *(Focus on the `action`, `task_description`, and `params` for each task to understand the desired change.)*

        {tasks_json}


    ## Your Validation Objectives:
    1.  **Task Verification:** For EACH task listed in 'Tasks Applied':
        *   Visually locate the relevant element(s) in both 'Before' and 'After' images based on the original instruction, task's description and target hint.
        *   Compare the appearance (e.g., font, size, color, position, alignment, content, style) of the target element(s) between the images.
        *   Determine if the specific modification described by the task (`action`, `params`) was visually applied correctly in the 'After' image. Consider edge cases like text overflow or elements moving off-slide.
        *   Assign a status: `Success`, `Partial Success` (if change is incomplete or has minor issues), or `Failure` (if change is wrong or not applied). Provide concise reasoning.
    2.  **Visual Regression Detection:** Carefully examine the 'After' image compared to the 'Before' image, looking for ANY unintended negative changes introduced during modification, such as:
        *   New element overlaps.
        *   Misalignments not related to the intended tasks.
        *   Unexpected color or style changes in unrelated elements.
        *   Distorted shapes or images.
        *   Missing elements (that shouldn't have been deleted).
        *   Text becoming unreadable or cut off (unless resizing was the goal).
        *   Significant deviation from standard layout principles (if not intended).
    3.  **Overall Assessment:** Based on the success of individual tasks and the presence/absence of regressions, provide an overall validation status for the slide modification.  


    ## Output Requirements:
    Respond ONLY with a single, valid JSON object containing the following keys:
    *   `success`: (boolean string) Your final assessment for this slide - one of: `"True"`, `"False"`.
    *   `task_assessments`: (List of Objects) One object for each task in the input `tasks_json`. Each object must have:
        *   `task_description`: (String) The `task_description` from the input task.
        *   `assessment`: (String) Your assessment for this specific task - one of: `"Success"`, `"Partial Success"`, `"Failure"`.
        *   `reasoning`: (String) Concise explanation for your assessment (max 2 sentences).
    *   `issues_found`: (List of any other issues detected) A list describing any specific issues detected. If none, return an empty list `[]`.
    *   `validation_success_percentage`: (int) Your confidence in this validation assessment as a percentage (0-100).

"""

    
def validate_presentation(original_pptx_path, modified_pptx_path, task_specifications):
    logging.info(f"Starting validation of modified presentation {modified_pptx_path}...")
    validation_report = {"issues_found": [], "success": True, "task_assessments": [], "validation_success_percentage": 0 }
    
    try:
        # Convert both PPTX files to images
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        original_images = pptx_to_images(original_pptx_path, output_dir=os.path.join(base_dir, "input_ppts", "converted_pdfs"))
        modified_images = pptx_to_images(modified_pptx_path, output_dir=os.path.join(base_dir, "output_ppts", "converted_pdfs"))
        
        if len(original_images) != len(modified_images):
            logging.warning("Mismatch in slide count between original and modified presentations")
            validation_report["issues_found"].append("Slide count mismatch")
            validation_report["success"] = False
            validation_report["validation_success_percentage"] = 0
            return validation_report, False, 0

        for slide_idx, (original_img, modified_img) in enumerate(zip(original_images, modified_images)):
            slide_tasks = [task for task in task_specifications if task["slide_number"] == slide_idx + 1]
            if not slide_tasks:
                continue
            
            original_img_bytes = BytesIO()
            original_img.save(original_img_bytes, format="PNG")
            original_img_bytes = original_img_bytes.getvalue()

            modified_img_bytes = BytesIO()
            modified_img.save(modified_img_bytes, format="PNG")
            modified_img_bytes = modified_img_bytes.getvalue()
            
            # Prepare image parts for VLM
            original_part = genai.types.Part.from_bytes(data=original_img_bytes, mime_type="image/png")
            modified_part = genai.types.Part.from_bytes(data=modified_img_bytes, mime_type="image/png")
            tasks_json = json.dumps(slide_tasks)

            prompt = VALIDATION_PROMPT.format(tasks_json=tasks_json)

            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=[prompt, original_part, modified_part]
            )
            
            try:
                json_match = re.search(r'\{.*\}', response.text, re.DOTALL)
                if not json_match:
                    raise ValueError("No valid JSON object found in LLM response")
                
                validation_result = json.loads(json_match.group(0))
                
                required_keys = {"success", "task_assessments", "issues_found", "validation_success_percentage"}
                if not all(key in validation_result for key in required_keys):
                    raise ValueError(f"Missing required keys in LLM response. Expected {required_keys}, got {validation_result.keys()}")

                if validation_result["success"] == "False":
                    validation_report["success"] = False
                
                validation_report["task_assessments"].extend(validation_result["task_assessments"])
                validation_report["issues_found"].extend([f"Slide {slide_idx+1}: {issue}" for issue in validation_result["issues_found"]])
                validation_report["validation_success_percentage"] = max(
                    validation_report["validation_success_percentage"],
                    validation_result["validation_success_percentage"]
                )

                if validation_result["issues_found"]:
                    logging.warning(f"Validation issues for slide {slide_idx}: {validation_result['issues_found']}")

            except (json.JSONDecodeError, ValueError) as e:
                logging.error(f"Failed to parse LLM response for slide {slide_idx}: {e}")
                validation_report["issues_found"].append(f"Slide {slide_idx}: Failed to parse validation response - {str(e)}")
                validation_report["success"] = False
                validation_report["validation_success_percentage"] = 0

        # Finalize success based on all slides
        validation_report["success"] = all(task["assessment"] == "Success" for task in validation_report["task_assessments"]) if validation_report["task_assessments"] else validation_report["success"]
        return validation_report, validation_report["success"], validation_report["validation_success_percentage"]

    except Exception as e:
        logging.error(f"Validation process failed: {e}", exc_info=True)
        validation_report["issues_found"].append(f"Validation error: {str(e)}")
        validation_report["success"] = False
        validation_report["validation_success_percentage"] = 0
        return validation_report, False, 0