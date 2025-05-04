# metadata_handler.py
from fastapi import APIRouter, HTTPException, Body, status
from pydantic import BaseModel
import os
import json
import aiofiles
import logging

router = APIRouter()

# Define a Pydantic model for the request body
class MetadataPayload(BaseModel):
    data: dict | list
    filename: str = "metadata.json"

BASE_SAVE_PATH = "./slide_images"

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- Metadata Upload Handler ---
@router.post(
    "",
    status_code=status.HTTP_200_OK, 
    response_model=dict
)
async def upload_metadata(payload: MetadataPayload = Body(...)): 
    """
    Receives slide metadata and saves it to a JSON file, overwriting
    any existing file with the same name in the predefined directory.
    """
    try:
        os.makedirs(BASE_SAVE_PATH, exist_ok=True)
        safe_filename = os.path.basename(payload.filename)
        save_path = os.path.join(BASE_SAVE_PATH, safe_filename)

        # --- Async File Writing ---
        async with aiofiles.open(save_path, "w", encoding="utf-8") as f:
            # json.dumps creates the string first, then write is async
            json_string = json.dumps(payload.data, indent=2)
            await f.write(json_string)

        logger.info(f"Metadata successfully saved to: {save_path}")  
        return {"message": "Metadata saved successfully.", "path": save_path}

    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON data provided: {e}") 
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Invalid JSON data provided: {e}"
        )
    except IOError as e:
        logger.error(f"Failed to write metadata file to {save_path}: {e}")  
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to save metadata file due to IO error: {e}"
        )
    except Exception as e:
        logger.error(f"Unexpected error saving metadata: {e}")  
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"An unexpected error occurred: {e}"
        )