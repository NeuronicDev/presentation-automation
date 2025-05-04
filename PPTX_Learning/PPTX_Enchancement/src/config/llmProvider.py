# llmProvider.py
import logging
from langchain_google_genai import GoogleGenerativeAI, GoogleGenerativeAIEmbeddings, HarmBlockThreshold, HarmCategory
from config.config import LLM_API_KEY, GEMINI_FLASH_2_0_MODEL, GEMINI_EMBEDDINGS_MODEL, GEMINI_FLASH_2_0_MODEL_LITE
import google.api_core.exceptions

# Logging
logging.basicConfig(level=logging.INFO)

# === Gemini Safety Settings ===
safety_settings = {
    HarmCategory.HARM_CATEGORY_UNSPECIFIED: HarmBlockThreshold.BLOCK_ONLY_HIGH,
    HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
    HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_ONLY_HIGH,
    HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
    HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
}

# === Model Initialization Functions ===
def initialize_gemini_llm(model_name):
    try:
        return GoogleGenerativeAI(
            model=model_name,
            google_api_key=LLM_API_KEY,
            temperature=0.1,
            max_tokens=8192,
            verbose=True,
            timeout=None,
            max_retries=2,
            safety_settings=safety_settings
        )
    except Exception as e:
        if "429 Resource has been exhausted" in str(e):
            logging.warning("API quota exhausted.")
        logging.error(f"Failed to initialize LLM {model_name}: {str(e)}")
        raise
    
def initialize_gemini_embeddings():
    try:
        return GoogleGenerativeAIEmbeddings(model=GEMINI_EMBEDDINGS_MODEL, google_api_key=LLM_API_KEY)
    except google.api_core.exceptions.ResourceExhausted:
        logging.warning("API key limit exhausted for Gemini Embeddings. Please check your quota or use a different API key.")
        raise 

    except Exception as e:
        logging.error(f"Failed to initialize embeddings: {str(e)}")
        raise 

# === Client Instances ===
gemini_flash_llm = initialize_gemini_llm(GEMINI_FLASH_2_0_MODEL)
gemini_flash_llm_lite = initialize_gemini_llm(GEMINI_FLASH_2_0_MODEL_LITE)
gemini_embeddings = initialize_gemini_embeddings()