# metadata_handler.py
from fastapi import APIRouter, HTTPException, Body, status
from pydantic import BaseModel
import os, shutil, json, logging, aiofiles

router = APIRouter()

# Define a Pydantic model for the request body
class MetadataPayload(BaseModel):
    data: dict | list
    filename: str = "metadata.json"

BASE_SAVE_PATH = "./uploaded_pptx/slide_images/metadata"

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- Helper Functions ---
def clear_directory_contents(directory_path: str):
    """Removes all files and subdirectories within the specified directory."""
    if not os.path.isdir(directory_path):
        logger.info(f"Directory {directory_path} does not exist, nothing to clear.")
        return
    logger.info(f"Clearing contents of directory: {directory_path}")
    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
                logger.debug(f"Deleted file: {file_path}")
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
                logger.debug(f"Deleted directory: {file_path}")
        except Exception as e:
            logger.error(f"Failed to delete {file_path}. Reason: {e}")

# Flag to clear the metadata directory only once per session
cleared_once = False

# --- Metadata Upload Handler ---
@router.post(
    "",
    status_code=status.HTTP_200_OK, 
    response_model=dict
)
async def upload_metadata(payload: MetadataPayload = Body(...)):
    """
    Receives slide metadata and saves it to a JSON file.
    On first call, clears the metadata directory to remove old files.
    """
    global cleared_once
    try:
        os.makedirs(BASE_SAVE_PATH, exist_ok=True)
        # Clear directory contents only once per session
        if not cleared_once:
            clear_directory_contents(BASE_SAVE_PATH)
            cleared_once = True

        safe_filename = os.path.basename(payload.filename)
        save_path = os.path.join(BASE_SAVE_PATH, safe_filename)

        async with aiofiles.open(save_path, "w", encoding="utf-8") as f:
            json_string = json.dumps(payload.data, indent=2)
            await f.write(json_string)

        logger.info(f"[UPLOAD] Metadata saved: {save_path}")
        return {"message": "Metadata saved successfully.", "saved_file": safe_filename, "path": save_path}

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