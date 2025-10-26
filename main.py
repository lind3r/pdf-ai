import os
import base64
import json
import requests
from pathlib import Path
import PyPDF2
from docx import Document
import argparse
from pdf import generate_pdf_report

# === Configuration ===
API_URL = "http://localhost:11434/api/generate"
MODEL = "gemma3:12b"
REPORT_JSON = "report.json"
SUMMARY_REPORT_JSON = "summary_report.json"

def extract_text_from_txt(file_path):
    with open(file_path, "r", encoding="utf-8") as f:
        return f.read().strip()

def extract_text_from_pdf(file_path):
    text = ""
    with open(file_path, "rb") as f:
        reader = PyPDF2.PdfReader(f)
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text.strip()

def extract_text_from_docx(file_path):
    doc = Document(file_path)
    return "\n".join([p.text for p in doc.paragraphs]).strip()

def encode_image_to_base64(file_path):
    with open(file_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

def call_ollama_api(payload):
    """Send a request to Ollama API and return the response text."""
    try:
        response = requests.post(API_URL, json=payload)
        if response.status_code == 200:
            result = response.json()
            return result.get("response", "").strip()
        else:
            print(f"API request failed: {response.status_code} {response.text}")
            return None
    except Exception as e:
        print(f"Error sending API request: {e}")
        return None

def process_folder(folder_path):
    report = []

    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            file_path = os.path.join(root, filename)
            ext = Path(filename).suffix.lower()

            # === Handle text-based documents (.pdf, .docx, .txt) ===
            if ext in [".pdf", ".docx", ".txt"]:
                if ext == ".pdf":
                    text = extract_text_from_pdf(file_path)
                elif ext == ".docx":
                    text = extract_text_from_docx(file_path)
                else:  # .txt
                    text = extract_text_from_txt(file_path)

                payload = {
                    "model": MODEL,
                    "system": "Du är en strikt summariserande AI. Returnera bara ren text, inga förklaringar, inga hälsningar, inga kommentarer.",
                    "prompt": f"Beskriv ingående vad detta dokument handlar om och sammanfatta innehållet med högst 1000 ord:\n\n{text}",
                    "stream": False,
                }
                file_type = "text"

            # === Handle image files ===
            elif ext in [".png", ".jpg", ".jpeg"]:
                img_base64 = encode_image_to_base64(file_path)
                payload = {
                    "model": MODEL,
                    "system": "Du är en strikt bildbeskrivande AI. Returnera endast text på svenska, inga kommentarer eller frågor.",
                    "prompt": "Beskriv ingående vad som syns på denna bild, inklusive eventuell text, med högst 1000 ord.",
                    "images": [img_base64],
                    "stream": False,
                }
                file_type = "bild"

            else:
                print(f"Skipping unsupported file type: {filename}")
                continue

            # === Process file ===
            print(f"Processing {filename} ...")
            summary = call_ollama_api(payload)
            report.append({
                "filnamn": filename,
                "typ": file_type,
                "sammanfattning": summary
            })

    # Save report
    with open(REPORT_JSON, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, ensure_ascii=False)

    print(f"Report saved to {REPORT_JSON}")

def order_and_summarize_report(input_json_path, output_json_path):
    """Send the completed report to the AI for logical ordering and overall summary."""
    with open(input_json_path, "r", encoding="utf-8") as f:
        report_data = json.load(f)

    # The schema we expect Ollama to enforce
    json_schema = {
        "type": "object",
        "properties": {
            "ordered_files": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "filnamn": {"type": "string"},
                        "typ": {"type": "string"},
                        "sammanfattning": {"type": "string"},
                    },
                    "required": ["filnamn", "typ", "sammanfattning", "order"]
                }
            },
            "overall_summary": {"type": "string"}
        },
        "required": ["ordered_files", "overall_summary"]
    }

    prompt = f"""
Analysera följande lista av filer och deras sammanfattningar.
Ordna dem i en logisk sekvens där liknande ämnen följer varandra.
Skapa också en "overall_summary" som sammanfattar allt innehåll.

Här är rapporten:
{json.dumps(report_data, ensure_ascii=False, indent=2)}
"""

    payload = {
        "model": MODEL,
        "system": "Du är en strikt analytisk AI. Returnera endast giltig JSON enligt formatet.",
        "prompt": prompt.strip(),
        "stream": False,
        "format": json_schema
    }

    print("Analyzing and ordering summaries ...")
    response = call_ollama_api(payload)

    if not response:
        print("Failed to get response for ordering.")
        return

    try:
        result = json.loads(response)
    except json.JSONDecodeError:
        print("Invalid JSON returned by model, saving raw output.")
        result = {"error": "Invalid JSON", "raw_output": response}

    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2, ensure_ascii=False)

    print(f"Ordered report saved to {output_json_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Summarize documents and images using Ollama API.")
    parser.add_argument("folder", nargs="?", default="./contents",  help="Path to the folder containing files to process (default: ./contents)")
    args = parser.parse_args()

    process_folder(args.folder)
    order_and_summarize_report(REPORT_JSON, SUMMARY_REPORT_JSON)
    generate_pdf_report(SUMMARY_REPORT_JSON, "output.pdf", args.folder)
