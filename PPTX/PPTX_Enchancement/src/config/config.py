import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# === API Keys ===
LLM_API_KEY = os.getenv("LLM_API_KEY")

# === Gemini Model Configs ===
GEMINI_FLASH_2_0_MODEL = "gemini-2.0-flash"
GEMINI_FLASH_2_0_MODEL_LITE = "gemini-2.0-flash-lite-001"
GEMINI_EMBEDDINGS_MODEL = "models/text-embedding-004"

# === Paths ===
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"
POPPLER_PATH = r"C:\poppler-24.08.0\Library\bin"

# === Output Directory ===
OUTPUT_DIR = "output_ppts"
os.makedirs(OUTPUT_DIR, exist_ok=True)