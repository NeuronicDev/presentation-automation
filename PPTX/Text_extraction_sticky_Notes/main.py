import zipfile
import xml.etree.ElementTree as ET
import requests
import json

# Groq API Configuration
API_KEY = "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"
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
HEADERS = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}

def extract_pptx_xml(pptx_path):
    """Extracts the raw XML content from a PPTX file."""
    with zipfile.ZipFile(pptx_path, "r") as pptx_zip:
        slide_files = [f for f in pptx_zip.namelist() if f.startswith("ppt/slides/slide") and f.endswith(".xml")]
        return {file: pptx_zip.read(file) for file in slide_files}

def extract_text_boxes(xml_data):
    """Extracts text boxes and their properties from slide XML."""
    text_boxes = []
    root = ET.fromstring(xml_data)
    for shape in root.findall(".//p:sp", namespaces={'p': "http://schemas.openxmlformats.org/presentationml/2006/main"}):
        text = "".join(t.text for t in shape.findall(".//a:t", namespaces={'a': "http://schemas.openxmlformats.org/drawingml/2006/main"}) if t.text)
        if text.strip():
            text_boxes.append(text.strip())
    return text_boxes

def validate_sticky_notes(text_boxes):
    """Uses LLaMA3-70B to extract only instruction-like text from sticky notes."""
    prompt = (
        "Extract and return only the instructional content from the following text. "
        "Instructional content includes tasks, action points, to-do lists, reminders, or directives. "
        "Do not return any explanations or additional context—only the extracted instructions as plain text."
        f"\n\nText: {text_boxes}\n\nExtracted Instructions:"
    )

    payload = {
        "model": "llama3-70b-8192",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.2,
    }

    response = requests.post(GROQ_URL, headers=HEADERS, json=payload)
    return response.json()["choices"][0]["message"]["content"] if response.status_code == 200 else "Validation Failed"

def process_pptx(pptx_path):
    """Extracts, validates, and processes sticky notes from a PPTX file."""
    slides_xml = extract_pptx_xml(pptx_path)
    extracted_notes = {}
    for slide, xml_data in slides_xml.items():
        text_boxes = extract_text_boxes(xml_data)
        if text_boxes:
            extracted_notes[slide] = validate_sticky_notes(text_boxes)
    return extracted_notes

pptx_file = "Source.pptx"
sticky_notes = process_pptx(pptx_file)
print(json.dumps(sticky_notes, indent=2))