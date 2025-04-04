import zipfile
import os
import logging
import tempfile
import shutil

def inject_xml_into_ppt(pptx_path, slide_number, modified_xml):

    try:
        temp_dir = tempfile.mkdtemp(prefix="pptx_xml_inject_")
        
        # The path to the slide XML in the PPTX archive
        slide_path = f"ppt/slides/slide{slide_number}.xml"
        
        # Extract the PPTX (which is a ZIP file)
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Replace the slide XML with our modified version
        full_slide_path = os.path.join(temp_dir, slide_path)
        
        # Check if the directory exists
        os.makedirs(os.path.dirname(full_slide_path), exist_ok=True)
        
        # Write the modified XML
        with open(full_slide_path, 'w', encoding='utf-8') as f:
            f.write(modified_xml)
        
        # Create a new ZIP file with the modified content
        temp_zip_path = os.path.join(temp_dir, "temp.pptx")
        
        with zipfile.ZipFile(temp_zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zipf:
            # Add all files from the temp directory
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    if file == "temp.pptx":
                        continue  # Skip the zip file itself
                    
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname)
        
        # Replace the original file with our new one
        shutil.move(temp_zip_path, pptx_path)
        
        # Clean up
        shutil.rmtree(temp_dir)
        
        logging.info(f"Successfully injected modified XML for slide {slide_number}")
        return True
        
    except Exception as e:
        logging.error(f"Error injecting modified XML: {e}")
        return False