# config.py
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# === API Keys ===
LLM_API_KEY = os.getenv("LLM_API_KEY")

# === Gemini Model Configs ===
GEMINI_FLASH_2_0_MODEL = "gemini-2.0-flash-001"
GEMINI_FLASH_2_0_MODEL_LITE = "gemini-2.0-flash-lite-001"
GEMINI_EMBEDDINGS_MODEL = "models/text-embedding-004"

# === Paths ===
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
POPPLER_PATH = r"C:\poppler-24.08.0\Library\bin"

DOCKER_IMAGE_NAME = "pptx-automation-api:latest"