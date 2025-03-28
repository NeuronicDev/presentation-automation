import logging, zipfile, pptx
from lxml import etree
import subprocess
from pdf2image import convert_from_path
import base64
from io import BytesIO
import os

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


    
def extract_slide_xml(prs, slide_index):
    try:
        slide = prs.slides[slide_index]
        slide_element = slide._element
        xml = etree.tostring(slide_element, encoding='unicode', pretty_print=True, method='xml')
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
        base64_image = base64.b64encode(buffered.getvalue()).decode('utf-8')
        return base64_image
    except Exception as e:
        logging.error(f"Failed to generate image for slide index {slide_index}: {e}")
        return "" 
    

def generate_slide_context(prs, slide_number, pdf_path, image_cache):
    slide_index = slide_number - 1 
    if slide_number not in image_cache:
        image_cache[slide_number] = generate_slide_image(pdf_path, slide_index)
    slide_xml = extract_slide_xml(prs, slide_index)
    
    return {
        "slide_xml_structure": slide_xml,
        "slide_image_base64": image_cache[slide_number]
    }