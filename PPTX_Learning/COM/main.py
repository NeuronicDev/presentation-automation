# import os
# import base64
# import tempfile
# import win32com.client
# import logging
# from openai import OpenAI

# # Initialize logging
# logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# logger = logging.getLogger(__name__)

# # Initialize OpenAI Gemini API Client
# client = OpenAI(
#     api_key="AIzaSyAv0sTw83EOKcJtoSyT9ug4cnzwGagkMJY",
#     base_url="https://generativelanguage.googleapis.com/v1beta/openai/"
# )

# class PowerPointProcessor:
#     def __init__(self, pptx_file):
#         self.pptx_file = pptx_file
#         self.inserted_slides_count = 0

#         # Open PowerPoint via COM
#         self.app = win32com.client.Dispatch("PowerPoint.Application")
#         self.app.Visible = True  # Keep PowerPoint visible
#         self.presentation = self.app.Presentations.Open(os.path.abspath(pptx_file))

#     def extract_text_elements(self):
#         """Extract text elements from each slide."""
#         slide_data = {}
#         try:
#             for slide_num, slide in enumerate(self.presentation.Slides, start=1):
#                 text_content = []
#                 for shape in slide.Shapes:
#                     if shape.HasTextFrame and shape.TextFrame.HasText:
#                         text = shape.TextFrame.TextRange.Text.strip()
#                         if text:
#                             text_content.append(text)
                
#                 if text_content:
#                     slide_data[slide_num] = "\n".join(text_content)

#             logger.info(f"Extracted text from {len(slide_data)} slides.")
#             return slide_data
#         except Exception as e:
#             logger.error(f"Error extracting text from slides: {e}")
#             raise


#     # def analyze_slide_with_gemini(self, slide_text):
#     #     """Analyze slide text using Gemini and convert it into a structured table."""
#     #     try:
#     #         prompt = f"""
#     #             **Task:** Convert the provided slide content into a structured table, formatted correctly as CSV.

#     #             ### **Rules for Extraction**
#     #             1. Identify table-like structures from text boxes and align them **properly**.
#     #             2. Preserve the **exact** text content (e.g., `"Xxxx"`) without modification.
#     #             3. Ensure **each column aligns correctly** based on the slide’s layout.

#     #             ### **Column-Specific Guidelines**
#     #             - The **first column’s first row** may or may not have a header; if missing, **do NOT add new text**.
#     #             - If a column contains multiple **sub-rows** within a single row, **split them using visible spacing**.
#     #             - If **sub-rows are separated by a full-row space**, treat them as separate rows.
#     #             - If a **1×1 text box is empty**, it should be represented as an **empty cell** (`""` or `[EMPTY]`).

#     #             ### **Handling Empty Spaces**
#     #             - **Do NOT remove** empty spaces. Instead, keep them **exactly as they appear in the original slide**.
#     #             - If a cell is empty **due to spacing**, keep it empty **rather than shifting data**.
#     #             - Use explicit `"[EMPTY]"` placeholders in the output if necessary.

#     #             ### **Output Formatting**
#     #             - The first row should always contain **column headers** (even if inferred).
#     #             - Each row must have **the same number of columns**.
#     #             - Use the **vertical bar ("|")** as the delimiter instead of commas to avoid misalignment.
#     #             - Do **not** add extra explanations—return **only** the CSV content.

#     #             ### **Example Input & Expected Output**
#     #             #### **Input Slide Layout**


#     #             ---
#     #             **Example Input Layout:**  

#     #             ```
#     #                     Xxxx          Xxxx    

#     #         Xxxx                      Xxxx    
#     #                                   Xxxx  

#     #         Xxxx        Xxxx          Xxxx  
#     #                                   Xxxx  

#     #         Xxxx        Xxxx          Xxxx    
                
#     #             ```

#     #             **Expected CSV Output:**  
#     #             ```
#     #             [EMPTY] | Xxxx  | Xxxx  
#     #             Xxxx    | [EMPTY] | Xxxx  
#     #             [EMPTY] | [EMPTY] | Xxxx  
#     #             Xxxx    | Xxxx  | Xxxx  
#     #             [EMPTY] | [EMPTY] | Xxxx  
#     #             Xxxx    | Xxxx  | Xxxx  
#     #             ```

#     #             **Slide Content:**  
#     #             {slide_text}
#     #         """

#     #         response = client.chat.completions.create(
#     #             model="gemini-2.0-flash",
#     #             messages=[{"role": "user", "content": [{"type": "text", "text": prompt}]}],
#     #         )

#     #         extracted_text = response.choices[0].message.content.strip()
#     #         if not extracted_text:
#     #             logger.warning("No structured data extracted.")
#     #             return None

#     #         # Convert extracted text to structured table format
#     #         table_data = [row.strip().split("|") for row in extracted_text.split("\n") if row.strip()]
#     #         return table_data

#     #     except Exception as e:
#     #         logger.error(f"Error processing slide text with Gemini: {e}")
#     #         raise



#     def analyze_slide_with_gemini(self, slide_text):
#         """Analyze slide text using Gemini and convert it into a structured table."""
#         prompt = f"""
#         **Task:** Extract a structured table from the slide content and format it correctly as CSV.

#         ### **Rules**
#         - Maintain exact text content (e.g., `"Xxxx"`) without modification.
#         - **Align columns correctly** based on the slide’s layout.
#         - Empty spaces should **NOT be removed**.

#         ### **Handling Empty Spaces**
#         - If a cell is empty, it **must remain empty** in the CSV.
#         - Use `"[EMPTY]"` as a placeholder when needed.

#         ### **Output Formatting**
#         - Use `|` as the delimiter instead of commas.
#         - Ensure all rows have **the same number of columns**.
#         - Return **only the CSV** (no additional explanations).

#         **Slide Content:**  
#         {slide_text}
#         """

#         try:
#             response = client.chat.completions.create(
#                 model="gemini-2.0-flash",
#                 messages=[{"role": "user", "content": [{"type": "text", "text": prompt}]}],
#             )
#             extracted_text = response.choices[0].message.content.strip()
#             if not extracted_text:
#                 logger.warning("No structured data extracted.")
#                 return None

#             return [row.strip().split("|") for row in extracted_text.split("\n") if row.strip()]
#         except Exception as e:
#             logger.error(f"Error processing slide text with Gemini: {e}")
#             return None


#     def insert_table_into_ppt(self, table_data, slide_num):
#         """Insert a formatted table into a new slide."""
#         try:
#             adjusted_slide_num = slide_num + self.inserted_slides_count
#             slides = self.presentation.Slides
#             new_slide = slides.Add(adjusted_slide_num + 1, 5)  # 5: Title & Content Layout

#             # Formatting table
#             rows, cols = len(table_data), len(table_data[0])
#             table = new_slide.Shapes.AddTable(rows, cols, 50, 100, 600, 300).Table

#             for r, row in enumerate(table_data):
#                 for c, cell in enumerate(row):
#                     table.Cell(r + 1, c + 1).Shape.TextFrame.TextRange.Text = cell

#                     # Apply formatting: Center align headers, left-align values
#                     if r == 0:
#                         table.Cell(r + 1, c + 1).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # Center align
#                         table.Cell(r + 1, c + 1).Shape.TextFrame.TextRange.Font.Bold = True
#                     else:
#                         table.Cell(r + 1, c + 1).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 1  # Left align

#             # Adjusting column widths dynamically
#             for col in range(1, cols + 1):
#                 table.Columns(col).Width = 600 / cols  # Equal width for all columns

#             logger.info(f"Formatted table inserted after slide {slide_num}.")
#             self.inserted_slides_count += 1

#         except Exception as e:
#             logger.error(f"Error inserting table into PPT: {e}")
#             raise

#     def process_pptx(self):
#         """Main processing function."""
#         try:
#             slide_texts = self.extract_text_elements()
#             for slide_num, slide_text in slide_texts.items():
#                 table_data = self.analyze_slide_with_gemini(slide_text)
#                 if table_data:
#                     self.insert_table_into_ppt(table_data, slide_num)

#             self.save_pptx("output.pptx")
#         except Exception as e:
#             logger.error(f"Error during PPTX processing: {e}")
#             raise

#     def save_pptx(self, output_pptx):
#         """Save the modified PowerPoint file."""
#         try:
#             self.presentation.SaveAs(os.path.abspath(output_pptx))
#             self.presentation.Close()
#             self.app.Quit()
#             logger.info(f"Final Updated PPTX saved as {output_pptx}")
#         except Exception as e:
#             logger.error(f"Error saving the updated PPTX: {e}")
#             raise

# if __name__ == "__main__":
#     pptx_processor = PowerPointProcessor("Source.pptx")
#     pptx_processor.process_pptx()



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