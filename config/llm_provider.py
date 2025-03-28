import os,logging
from langchain_google_genai import GoogleGenerativeAI, ChatGoogleGenerativeAI, GoogleGenerativeAIEmbeddings, HarmBlockThreshold, HarmCategory
from config.config import LLM_API_KEY, GEMINI_FLASH_2_0_MODEL, GEMINI_EMBEDDINGS_MODEL, GEMINI_FLASH_2_0_MODEL_LITE
import google.api_core.exceptions

# Safety settings for the models
safety_settings = {
    HarmCategory.HARM_CATEGORY_UNSPECIFIED: HarmBlockThreshold.BLOCK_ONLY_HIGH,
    HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
    HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_ONLY_HIGH,
    HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
    HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_ONLY_HIGH,
}

def initialize_gemini_flash_llm():
    try:
        return GoogleGenerativeAI(
            model=GEMINI_FLASH_2_0_MODEL,
            google_api_key=LLM_API_KEY,
            temperature=0.1,
            max_tokens=8192,
            verbose=True,
            timeout=None,
            max_retries=2,
            safety_settings=safety_settings
        )
    
    except Exception as e:
        if "429 Resource has been exhausted" in e :
            logging.warning("API key limit exhausted for Gemini Flash LLM. Please check your quota or use a different key.")
            raise Exception("API key limit exhausted for Gemini Flash LLM. Please check your quota or use a different key.")
        raise Exception(f"Failed to initialize Gemini Pro LLM: {str(e)}")
    
    # except google.api_core.exceptions.ResourceExhausted:
    #     logging.warning("API key limit exhausted for Gemini Flash LLM. Please check your quota or use a different API key.")
    #     raise Exception("API key limit exhausted for Gemini Flash LLM. Please check your quota or use a different API key.")
    # except Exception as e:
    #     raise Exception(f"Failed to initialize Gemini Flash LLM: {str(e)}")


def initialize_gemini_flash_lite_llm():
    try:
        return GoogleGenerativeAI(
            model=GEMINI_FLASH_2_0_MODEL_LITE,
            google_api_key=LLM_API_KEY,
            temperature=0.1,
            max_tokens=8192,
            verbose=True,
            timeout=None,
            max_retries=2,
            safety_settings=safety_settings
        )
    # except google.api_core.exceptions.ResourceExhausted:
    #     logging.warning("API key limit exhausted for Gemini Pro LLM. Please check your quota or use a different API key.")
    #     raise Exception("API key limit exhausted for Gemini Pro LLM. Please check your quota or use a different API key.")
    
    except Exception as e:
        if "429 Resource has been exhausted" in e :
            logging.warning("API key limit exhausted for Gemini Pro LLM. Please check your quota or use a different key.")
            raise Exception("API key limit exhausted for Gemini Pro LLM. Please check your quota or use a different key.")
        raise Exception(f"Failed to initialize Gemini Pro LLM: {str(e)}")


def initialize_gemini_embeddings():
    try:
        return GoogleGenerativeAIEmbeddings(model=GEMINI_EMBEDDINGS_MODEL, google_api_key=LLM_API_KEY)
    except google.api_core.exceptions.ResourceExhausted:
        logging.warning("API key limit exhausted for Gemini Embeddings. Please check your quota or use a different API key.")
        raise Exception("API key limit exhausted for Gemini Embeddings. Please check your quota or use a different API key.")
    except Exception as e:
        raise Exception(f"Failed to initialize Gemini Embeddings: {str(e)}")



# Initialize the clients
gemini_flash_llm = initialize_gemini_flash_llm()
gemini_flash_llm_lite = initialize_gemini_flash_lite_llm()
gemini_embeddings = initialize_gemini_embeddings()