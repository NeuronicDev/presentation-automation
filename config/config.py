import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

LLM_API_KEY = os.getenv("LLM_API_KEY")



GEMINI_FLASH_2_0_MODEL = "gemini-2.0-flash-001"
GEMINI_EMBEDDINGS_MODEL = "models/text-embedding-004"
GEMINI_FLASH_2_0_MODEL_LITE = "gemini-2.0-flash-lite-001"

LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"

OUTPUT_DIR = "output_ppts"
os.makedirs(OUTPUT_DIR, exist_ok=True)

DOCKER_IMAGE_NAME = "ppt-automation-executor:latest"


