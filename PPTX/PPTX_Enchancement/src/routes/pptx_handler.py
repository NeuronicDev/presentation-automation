from fastapi import APIRouter, Request
import os, json, base64, subprocess, logging
from pdf2image import convert_from_path
from config.config import LIBREOFFICE_PATH, POPPLER_PATH
from pydantic import BaseModel

router = APIRouter()

# --- Upload schema ---
class PPTXData(BaseModel):
    filename: str
    filetype: str
    createdAt: str
    base64: str

def convert_pptx_to_pdf(pptx_path, output_dir=None):
    pptx_path = os.path.abspath(pptx_path)
    filename = os.path.basename(pptx_path)
    filename_without_ext = os.path.splitext(filename)[0]

    if output_dir is None:
        output_dir = os.path.abspath("./input_ppts/converted_pdfs")
    os.makedirs(output_dir, exist_ok=True)

    output_pdf = os.path.join(output_dir, f"{filename_without_ext}.pdf")

    try:
        subprocess.run(
            [LIBREOFFICE_PATH, "--headless", "--convert-to", "pdf", pptx_path, "--outdir", output_dir],
            check=True
        )
        return output_pdf
    except subprocess.CalledProcessError as e:
        logging.error(f"LibreOffice failed: {e}")
        raise
    except Exception as e:
        logging.error(f"General conversion error: {e}")
        raise

def convert_pdf_to_images(pdf_path, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    try:
        images = convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
        image_paths = []
        for idx, image in enumerate(images):
            image_path = os.path.join(output_dir, f"slide_{idx + 1}.png")
            image.save(image_path, format="PNG")
            image_paths.append(image_path)
        return image_paths
    except Exception as e:
        logging.error(f"PDF to image failed: {e}")
        raise

@router.post("")
async def upload_pptx(request: Request):
    try:
        raw_body = await request.body()
        json_data = json.loads(raw_body.decode("utf-8"))

        filename = json_data.get("filename", "presentation.pptx")
        base64_str = json_data.get("base64", "")

        # Save PPTX file
        slide_dir = "slide_images"
        os.makedirs(slide_dir, exist_ok=True)
        pptx_path = os.path.join(slide_dir, filename)

        with open(pptx_path, "wb") as f:
            f.write(base64.b64decode(base64_str))

        # Convert PPTX -> PDF -> images
        pdf_output_dir = os.path.join(slide_dir, "pdfs")
        pdf_path = convert_pptx_to_pdf(pptx_path, pdf_output_dir)

        image_output_dir = os.path.join(slide_dir, "images", os.path.splitext(filename)[0])
        image_paths = convert_pdf_to_images(pdf_path, image_output_dir)
        return {
            "message": "Conversion successful",
            "pdf_path": pdf_path,
            "slide_images": image_paths
        }
    except Exception as e:
        logging.error("Upload error", exc_info=True)
        return {"error": str(e)}