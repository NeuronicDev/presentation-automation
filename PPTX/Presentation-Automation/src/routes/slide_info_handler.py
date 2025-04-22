# from fastapi import APIRouter, Body, status, HTTPException
# from pydantic import BaseModel, Field
# from datetime import datetime
# import logging
# import os
# import json
# import aiofiles

# router = APIRouter()

# # --- Logging Setup ---
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# logger = logging.getLogger(__name__)

# # --- Constants ---
# BASE_SAVE_PATH = "./uploaded_pptx/slide_images"
# DEFAULT_FILENAME = "slide_metadata.json"

# # --- Slide Info Model ---
# class SlideInfo(BaseModel):
#     slideIndex: int = Field(..., description="0-based index of the current slide")
#     slideNumber: int = Field(..., description="1-based number of the current slide")
#     capturedAt: datetime = Field(..., description="Timestamp of capture")

# @router.post("", status_code=status.HTTP_200_OK, response_model=dict)
# async def receive_slide_info(info: SlideInfo = Body(...)):
#     """
#     Receives current slide index/number from Office.js, logs it,
#     and stores it as metadata JSON inside ./slide_images.
#     """
#     logger.info(f"Received slide info: Slide #{info.slideNumber} (Index: {info.slideIndex}) at {info.capturedAt}")

#     try:
#         os.makedirs(BASE_SAVE_PATH, exist_ok=True)
#         metadata_path = os.path.join(BASE_SAVE_PATH, DEFAULT_FILENAME)

#         # Convert the model to a dict
#         metadata = info.dict()
#         metadata["timestamp"] = metadata.pop("capturedAt").isoformat()

#         # Write to file asynchronously
#         async with aiofiles.open(metadata_path, mode="w", encoding="utf-8") as f:
#             await f.write(json.dumps(metadata, indent=2))

#         logger.info(f"Slide metadata saved to {metadata_path}")
#         return {
#             "status": "success",
#             "message": f"Slide info saved to {DEFAULT_FILENAME}",
#             "filename": DEFAULT_FILENAME,
#             "path": metadata_path
#         }

#     except Exception as e:
#         logger.error(f"Error saving slide metadata: {e}", exc_info=True)
#         raise HTTPException(
#             status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
#             detail=f"Failed to save slide metadata: {str(e)}"
#         )