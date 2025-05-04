# pptx_handler.py 
from fastapi import APIRouter, HTTPException, Body, status
from pydantic import BaseModel, Field
import os
import base64
import subprocess
import logging
import shutil 
import tempfile
import aiofiles 
from config.config import LIBREOFFICE_PATH, POPPLER_PATH
from pdf2image import convert_from_path
from pdf2image.exceptions import PDFInfoNotInstalledError, PDFPageCountError, PDFSyntaxError

router = APIRouter()

# --- Configuration ---
BASE_SLIDE_DIR = "./slide_images"
PPTX_SAVE_DIR = BASE_SLIDE_DIR
IMAGE_OUTPUT_BASE_DIR = os.path.join(BASE_SLIDE_DIR, "images")
TARGET_IMAGE_SUBDIR = "presentation"

# --- Pydantic Model for Request Body ---
class PPTXPayload(BaseModel):
    base64: str = Field(..., description="Base64 encoded content of the PPTX file")
    filename: str = "presentation.pptx"

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- Helper Functions ---
def clear_directory_contents(directory_path):
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

def convert_pptx_to_pdf_temp(pptx_path, temp_dir):
    """Converts PPTX to PDF in a specified temporary directory."""
    pptx_abspath = os.path.abspath(pptx_path)
    if not os.path.exists(pptx_abspath):
        raise FileNotFoundError(f"Input PPTX file not found: {pptx_abspath}")
    base_filename = os.path.splitext(os.path.basename(pptx_path))[0]
    temp_pdf_path = os.path.join(temp_dir, f"{base_filename}.pdf")
    logger.info(f"Attempting to convert {pptx_abspath} to temporary PDF: {temp_pdf_path}")
    try:
        process = subprocess.run(
            [
                LIBREOFFICE_PATH,
                "--headless",
                "--convert-to", "pdf",
                pptx_abspath,
                "--outdir", temp_dir 
            ],
            check=True, # Raises CalledProcessError on failure
            capture_output=True, # Capture stdout/stderr
            text=True, # Decode stdout/stderr as text
            timeout=120 # Add a timeout (e.g., 2 minutes)
        )
        logger.info(f"LibreOffice conversion stdout:\n{process.stdout}")
        if process.stderr:
             logger.warning(f"LibreOffice conversion stderr:\n{process.stderr}")

        # Verify the output file was created
        if not os.path.exists(temp_pdf_path):
             raise FileNotFoundError(f"LibreOffice conversion finished but output PDF not found: {temp_pdf_path}")

        logger.info(f"Successfully converted to temporary PDF: {temp_pdf_path}")
        return temp_pdf_path 

    except FileNotFoundError as e:
        # Handles case where LIBREOFFICE_PATH is wrong or input file disappears
         logger.error(f"File not found during conversion: {e}")
         raise
    except subprocess.CalledProcessError as e:
        logger.error(f"LibreOffice conversion failed with return code {e.returncode}.")
        logger.error(f"LibreOffice stdout:\n{e.stdout}")
        logger.error(f"LibreOffice stderr:\n{e.stderr}")
        raise RuntimeError(f"LibreOffice conversion failed (code {e.returncode}). Check logs.") from e
    except subprocess.TimeoutExpired as e:
        logger.error(f"LibreOffice conversion timed out after {e.timeout} seconds.")
        raise TimeoutError("PPTX to PDF conversion timed out.") from e
    except Exception as e:
        logger.error(f"An unexpected error occurred during PPTX to PDF conversion: {e}", exc_info=True)
        raise RuntimeError(f"PPTX to PDF conversion failed: {e}") from e

def convert_pdf_to_images(pdf_path, output_image_dir):
    """Converts PDF pages to PNG images in the specified output directory."""
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Input PDF file not found for image conversion: {pdf_path}")

    # Ensure the final *output* directory for images exists
    os.makedirs(output_image_dir, exist_ok=True)
    logger.info(f"Converting PDF {pdf_path} to images in {output_image_dir}")

    try:
        # Convert PDF pages to Pillow image objects
        images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH, fmt='png', thread_count=4) 
        if not images:
            logger.warning(f"No images were generated from PDF: {pdf_path}")
            return []

        image_paths = []
        # Save each image with a sequential name
        for idx, image in enumerate(images):
            # Use 1-based indexing for slide numbers, consistent naming
            image_filename = f"slide_{idx + 1}.png"
            image_path = os.path.join(output_image_dir, image_filename)
            image.save(image_path, "PNG") # Explicitly save as PNG
            image_paths.append(image_path)
            logger.debug(f"Saved image: {image_path}")

        logger.info(f"Successfully converted PDF to {len(image_paths)} images.")
        return image_paths

    # Catch specific pdf2image/Poppler errors
    except (PDFInfoNotInstalledError, PDFPageCountError, PDFSyntaxError) as e:
         logger.error(f"Poppler/pdf2image error during conversion of {pdf_path}: {e}", exc_info=True)
         raise RuntimeError(f"PDF to image conversion failed due to Poppler/PDF error: {e}") from e
    except Exception as e:
        logger.error(f"An unexpected error occurred during PDF to image conversion: {e}", exc_info=True)
        raise RuntimeError(f"PDF to image conversion failed: {e}") from e

# --- PPTX Upload and Conversion Endpoint ---
@router.post(
    "",
    status_code=status.HTTP_200_OK,
    response_model=dict
)
async def upload_pptx(payload: PPTXPayload = Body(...)):
    """
    Receives a Base64 encoded PPTX file, saves it, converts it to slide
    images (PNG format), cleaning the target image directory first.
    The intermediate PDF is not saved permanently.
    """
    pptx_path = None
    temp_dir = None
    temp_pdf_path = None

    try:
        # --- 1. Save Incoming PPTX File ---
        os.makedirs(PPTX_SAVE_DIR, exist_ok=True)
        # Sanitize filename to prevent path traversal issues
        safe_filename = os.path.basename(payload.filename)
        pptx_path = os.path.join(PPTX_SAVE_DIR, safe_filename)

        logger.info(f"Received request to process: {safe_filename}")
        logger.info(f"Saving PPTX to: {pptx_path}")
        try:
            # Decode Base64
            pptx_bytes = base64.b64decode(payload.base64)
            # Asynchronously write the PPTX file
            async with aiofiles.open(pptx_path, "wb") as f:
                await f.write(pptx_bytes)
            logger.info(f"PPTX file saved successfully: {pptx_path}")
        except (base64.binascii.Error, TypeError) as decode_err:
             logger.error(f"Base64 decoding failed for {safe_filename}: {decode_err}")
             raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Invalid Base64 data provided.")
        except IOError as io_err:
            logger.error(f"Failed to save PPTX file to {pptx_path}: {io_err}")
            raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail="Failed to save uploaded PPTX file.")

        # --- 2. Create Temporary Directory for PDF ---
        # This context manager handles cleanup automatically
        with tempfile.TemporaryDirectory(prefix="pptx_conv_") as temp_dir:
            logger.info(f"Created temporary directory: {temp_dir}")

            # --- 3. Convert PPTX to PDF (in Temp Dir) ---
            try:
                temp_pdf_path = convert_pptx_to_pdf_temp(pptx_path, temp_dir)
            except (FileNotFoundError, RuntimeError, TimeoutError, Exception) as conv_err:
                # Catch errors from the conversion function
                logger.error(f"PPTX to PDF conversion step failed: {conv_err}")
                # Raise specific HTTP exceptions based on error type if needed
                if isinstance(conv_err, FileNotFoundError):
                     raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=f"Conversion dependency error: {conv_err}")
                else:
                     raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=f"PPTX to PDF conversion failed: {conv_err}")

            # --- 4. Prepare Image Output Directory & Clear Old Images ---
            final_image_dir = os.path.join(IMAGE_OUTPUT_BASE_DIR, TARGET_IMAGE_SUBDIR)
            try:
                clear_directory_contents(final_image_dir) 
                os.makedirs(final_image_dir, exist_ok=True)
            except (IOError, OSError) as dir_err:
                 logger.error(f"Failed to prepare image output directory {final_image_dir}: {dir_err}")
                 raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail="Failed to prepare image output directory.")

            # --- 5. Convert PDF (from Temp Dir) to Images (in Final Dir) ---
            try:
                image_paths = convert_pdf_to_images(temp_pdf_path, final_image_dir)
            except (FileNotFoundError, RuntimeError, Exception) as img_conv_err:
                 logger.error(f"PDF to Image conversion step failed: {img_conv_err}")
                 raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=f"PDF to Image conversion failed: {img_conv_err}")

            # --- 6. Success Response ---
            logger.info(f"Process completed successfully for {safe_filename}.")
            return {
                "message": "PPTX processed and slide images generated successfully.",
                "original_filename": safe_filename,
                "saved_pptx_path": pptx_path,
                "slide_images_directory": final_image_dir,
                "slide_image_paths": image_paths 
            }

    except HTTPException as http_exc:
         # Re-raise HTTPExceptions directly
         raise http_exc
    except Exception as e:
        # Catch any other unexpected errors during the process
        logger.error(f"An unexpected error occurred during PPTX processing: {e}", exc_info=True)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"An unexpected server error occurred during processing: {e}"
        )