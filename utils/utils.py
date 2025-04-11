import logging, zipfile, pptx
from lxml import etree
import subprocess
from pdf2image import convert_from_path
import base64
from io import BytesIO
import os
import xml.dom.minidom
import re
from config.config import LIBREOFFICE_PATH


def convert_pptx_to_pdf(pptx_path, output_dir):
    
    pptx_path = os.path.abspath(pptx_path)
    filename = os.path.basename(pptx_path)
    filename_without_ext = os.path.splitext(filename)[0]
    
    os.makedirs(output_dir, exist_ok=True)
    
    output_pdf = os.path.join(output_dir, f"{filename_without_ext}.pdf")
    
    try:
        logging.info(f"Converting {pptx_path} to PDF at {output_pdf}...")
        try:
            subprocess.run(
                [LIBREOFFICE_PATH, "--headless", "--convert-to", "pdf", pptx_path, "--outdir", output_dir],
                check=True
            )
            logging.info(f"Successfully converted {pptx_path} to {output_pdf}")
            return output_pdf
        except subprocess.CalledProcessError as e:
            logging.error(f"Failed to convert PPTX to PDF: {e}")
            raise
    except Exception as e:
        logging.error(f"Error in conversion process: {e}")
        raise


# def extract_slide_xml(pptx_path, slide_index=0):
#     def extract_slide_xml_from_zip():
#         with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
#             slide_files = [f for f in zip_ref.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
            
#             if slide_index < len(slide_files):
#                 with zip_ref.open(slide_files[slide_index]) as slide_file:
#                     return slide_file.read().decode('utf-8')
#             else:
#                 raise ValueError(f"Slide index {slide_index} is out of range")

#     def extract_slide_xml_from_pptx():
#         prs = pptx.Presentation(pptx_path)
#         slide = prs.slides[slide_index]
#         slide_element = slide._element
#         return etree.tostring(slide_element, encoding='unicode')

#     try:
#         # Primary method: Zipfile extraction
#         xml_content = extract_slide_xml_from_zip()
#     except Exception as zip_error:
#         try:
#             # Fallback method: python-pptx extraction
#             xml_content = extract_slide_xml_from_pptx()
#         except Exception as pptx_error:
#             raise ValueError(f"Failed to extract slide XML: {zip_error}, {pptx_error}")
#     return xml_content


def extract_slide_xml_from_ppt(pptx_path, slide_number):
    try:
        slide_filename = f"ppt/slides/slide{slide_number}.xml"
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:

            if slide_filename not in zip_ref.namelist():
                logging.error(f"Could not find slide {slide_number} XML in {pptx_path}")
                return None
            
            xml_content = zip_ref.read(slide_filename)
            
            dom = xml.dom.minidom.parseString(xml_content)
            pretty_xml = dom.toprettyxml()
            
            logging.info(f"Successfully extracted XML for slide {slide_number}")
            return pretty_xml
    
    except Exception as e:
        logging.error(f"Error extracting slide XML: {e}")
        return None
    
    
def extract_slide_xml(prs, slide_index):
    try:
        slide = prs.slides[slide_index]
        slide_element = slide._element
        xml = etree.tostring(slide_element, encoding='unicode', pretty_print=True, method='xml')
        logging.info(f"Successfully extracted XML for slide index {slide_index}: {xml}")
        return xml
    except Exception as e:
        logging.error(f"Error extracting XML for slide index {slide_index}: {e}")
        raise

def generate_slide_image(pdf_path, slide_index):
    try:
        images = convert_from_path(pdf_path, first_page=slide_index + 1, last_page=slide_index + 1)
        image = images[0]
        buffered = BytesIO()
        image.save(buffered, format="PNG")
        img_bytes = buffered.getvalue()
        base64_image = base64.b64encode(buffered.getvalue()).decode('utf-8')
        return base64_image, img_bytes
    except Exception as e:
        logging.error(f"Failed to generate image for slide index {slide_index}: {e}")
        return "", None 

def generate_slide_context(prs, slide_number, pdf_path, image_cache):
    slide_index = slide_number - 1 
    
    base64_image = ""
    img_bytes = None

    if slide_number not in image_cache:
        generated_base64, generated_bytes = generate_slide_image(pdf_path, slide_index)
        image_cache[slide_number] = (generated_base64, generated_bytes)
        base64_image = generated_base64
        img_bytes = generated_bytes
        if generated_bytes is None:
             logging.warning(f"Image generation failed for slide {slide_number}. Context will lack image bytes.")
        else:
             logging.info(f"Successfully generated and cached image for slide {slide_number}.")
    else:
        logging.info(f"Cache hit for slide {slide_number}. Retrieving image data.")
        base64_image, img_bytes = image_cache[slide_number]

    try:
        slide_xml = extract_slide_xml(prs, slide_index)
    except Exception as e:
        logging.error(f"Failed to extract XML for slide index {slide_index}: {e}", exc_info=True)
        slide_xml = None

    context = {
        "slide_xml_structure": slide_xml,
        "slide_image_base64": base64_image,
        "slide_image_bytes": img_bytes 
    }
    return context

def update_slide_context_cache(pptx_path, pdf_path, slide_context_cache, image_cache):
    try:
        # Convert the modified PPTX to PDF
        output_dir = os.path.dirname(pdf_path)
        new_pdf_path = convert_pptx_to_pdf(pptx_path, output_dir=output_dir)
        
        # Load the modified presentation
        prs = pptx.Presentation(pptx_path)
        
        # Clear existing caches
        image_cache.clear()
        
        # Regenerate context for all existing slides
        for slide_number in list(slide_context_cache.keys()):
            if slide_number <= len(prs.slides):
                slide_context_cache[slide_number] = generate_slide_context(
                    prs, slide_number, new_pdf_path, image_cache
                )
                logging.info(f"Updated context for slide {slide_number}")
            else:
                logging.warning(f"Slide {slide_number} no longer exists, removing from cache")
                del slide_context_cache[slide_number]
        return slide_context_cache, new_pdf_path
    except Exception as e:
        logging.error(f"Error updating slide context cache: {e}")
        return slide_context_cache, pdf_path

