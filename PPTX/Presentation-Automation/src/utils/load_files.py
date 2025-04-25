# utils/load_files.py
import os
import aiofiles
from typing import Dict, Any
from fastapi import HTTPException, status
import logging

logger = logging.getLogger(__name__)

CONTEXT_BASE_DIR = "uploaded_pptx/slide_images/presentation" 

async def load_slide_contexts(target_slides: list[int]) -> Dict[int, Dict[str, Any]]:
    loaded_context_dict: Dict[int, Dict[str, Any]] = {}

    if not target_slides:
        logger.warning("No target slide indices provided. Skipping context load.")
        return loaded_context_dict

    try:
        all_files_in_dir = os.listdir(CONTEXT_BASE_DIR)
        all_files_set = set(all_files_in_dir)

        for index in target_slides:

            slide_context: Dict[str, Any] = {}

            xml_path = os.path.join(CONTEXT_BASE_DIR, f"slide{index}.xml")
            img_txt_path = os.path.join(CONTEXT_BASE_DIR, f"slide{index}_image.txt")
            img_png_path = os.path.join(CONTEXT_BASE_DIR, f"slide{index}.png")

            if f"slide{index}.xml" not in all_files_set:
                raise HTTPException(
                    status_code=status.HTTP_404_NOT_FOUND,
                    detail=f"Missing XML context for slide {index}."
                )
            async with aiofiles.open(xml_path, "r", encoding="utf-8") as f:
                slide_context["slide_xml_structure"] = await f.read()

            if f"slide{index}_image.txt" not in all_files_set:
                raise HTTPException(
                    status_code=status.HTTP_404_NOT_FOUND,
                    detail=f"Missing base64 image TXT for slide {index}."
                )
            async with aiofiles.open(img_txt_path, "r", encoding="utf-8") as f:
                slide_context["slide_image_base64"] = await f.read()

            if f"slide{index}.png" in all_files_set:
                async with aiofiles.open(img_png_path, "rb") as f:
                    slide_context["slide_image_bytes"] = await f.read()
            else:
                slide_context["slide_image_bytes"] = None

            loaded_context_dict[index] = slide_context

        return loaded_context_dict

    except Exception as e:
        logger.exception("Error loading slide context.")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to load slide context: {str(e)}"
        )