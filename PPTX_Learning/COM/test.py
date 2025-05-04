import os
import csv
import base64
import tempfile
import win32com.client
from openai import OpenAI
from PIL import Image
from io import BytesIO
import logging

# --- Set up OpenAI client for Gemini API ---
client = OpenAI(
    api_key="AIzaSyAv0sTw83EOKcJtoSyT9ug4cnzwGagkMJY",
    base_url="https://generativelanguage.googleapis.com/v1beta/openai/"
)

# --- Configure Logging ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class PowerPointProcessor:
    def __init__(self, pptx_file, slide_number):
        self.pptx_file = pptx_file
        self.slide_number = slide_number
        self.temp_dir = tempfile.mkdtemp()
        self.output_csv = "output.csv"
        
        # --- Open PowerPoint ---
        self.app = win32com.client.Dispatch("PowerPoint.Application")
        self.app.Visible = 1  
        self.presentation = self.app.Presentations.Open(os.path.abspath(self.pptx_file))
        self.slide = self.presentation.Slides(self.slide_number)
    
    def extract_text_from_slide(self):
        """
        Extracts structured text elements from the slide while preserving empty cells.
        Ensures the first column’s first row remains empty if no header exists.
        """
        structured_data = []
        first_row = True 

        try:
            for shape in self.slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text = shape.TextFrame.TextRange.Text.strip()
                else:
                    text = ""  # Preserve empty spaces

                if first_row:
                    # If it's the first column in the first row and text is empty, keep it empty
                    structured_data.append([text if text else ""])  
                    first_row = False  # Move to next row processing
                else:
                    structured_data[-1].append(text)  # Append text to the last row

            if structured_data:
                self.save_csv(structured_data)
            else:
                logger.warning("No structured text found on the slide.")

        except Exception as e:
            logger.error(f"Error extracting text: {e}")

    @staticmethod
    def encode_image(image_bytes):
        """
        Encodes an image in Base64 format for API processing.
        """
        try:
            return base64.b64encode(image_bytes).decode("utf-8")
        except Exception as e:
            logger.error(f"Error encoding image: {e}")
            return None

    def process_image_with_gemini(self, image_bytes):
        """
        Sends the image to the Gemini API to extract table data in CSV format.
        """
        try:
            base64_image = self.encode_image(image_bytes)
            if not base64_image:
                return None
            
            prompt = """
                You are an AI assistant that extracts structured tables from PowerPoint slide images.
                - Perform OCR to detect and extract table data.
                - Return the extracted table in CSV format (comma-separated values).
                - Ensure the table structure is correctly maintained.
                - Do NOT include any markdown syntax like `|` or `---`.
                - Identify headers, values, and preserve row/column relationships.
                - Apply formatting rules:
                  - Alternate row highlighting
                  - Center-aligned headers
                  - Left-aligned values
                  - Bullet points where necessary
                  - Dotted lines between rows
                  - Merge rows/columns where needed
                  - Set appropriate width/height for cells
                  - Source information in the bottom-left corner
            """

            response = client.chat.completions.create(
                model="gemini-2.0-flash",
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}},
                        ],
                    }
                ],
            )

            extracted_text = response.choices[0].message.content.strip()
            if extracted_text:
                clean_rows = [row.strip().split(",") for row in extracted_text.split("\n") if row.strip()]
                self.save_csv(clean_rows)
                return clean_rows
            else:
                logger.warning("No text extracted from the image.")
                return None
        
        except Exception as e:
            logger.error(f"Error processing image with Gemini: {e}")
            return None

    def save_csv(self, data):
        try:
            with open(self.output_csv, "w", newline="", encoding="utf-8") as csvfile:
                writer = csv.writer(csvfile)
                writer.writerows(data)
            logger.info(f"Structured data saved to {self.output_csv}")
        except Exception as e:
            logger.error(f"Error saving CSV: {e}")

    def process_slide(self, image_bytes):
        try:
            self.extract_text_from_slide()
            self.process_image_with_gemini(image_bytes)
        except Exception as e:
            logger.error(f"Error processing slide: {e}")

    def close_ppt(self):
        """
        Closes the PowerPoint application properly.
        """
        try:
            self.presentation.Close()
            self.app.Quit()
            logger.info("PowerPoint closed successfully.")
        except Exception as e:
            logger.error(f"Error closing PowerPoint: {e}")

if __name__ == "__main__":
    pptx_file = "Source.pptx"  
    slide_number = 1  
    with open("image2.png", "rb") as image_file:
        image_bytes = image_file.read()

    processor = PowerPointProcessor(pptx_file, slide_number)
    processor.process_slide(image_bytes)
    processor.close_ppt()