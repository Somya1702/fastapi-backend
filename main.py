from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from fastapi.responses import FileResponse
import os
import fitz  # PyMuPDF for extracting text from PDFs
import openai

app = FastAPI(docs_url=None, redoc_url=None)  # Disable Swagger UI

# Serve the frontend at the root URL
@app.get("/")
def serve_frontend():
    return FileResponse("index.html")  # Ensure index.html is in the same folder as main.py

# ✅ Fix CORS Issue
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows requests from any domain
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods (GET, POST, etc.)
    allow_headers=["*"],  # Allows all headers
)


OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

def extract_text_from_pdf(pdf_path):
    """Extracts text from a PDF file"""
    text = ""
    pdf_document = fitz.open(pdf_path)
    for page in pdf_document:
        text += page.get_text("text") + "\n"

    # Trim text to first 2000 characters (GPT handles small inputs better)
    cleaned_text = text[:2000]  
    print("📄 Extracted Text (Trimmed):", cleaned_text)  # Debugging
    return cleaned_text

def get_company_name(text):
    """Send extracted text to OpenAI to get the company name."""
    client = openai.OpenAI(api_key=OPENAI_API_KEY)

    response = client.chat.completions.create(
        model="gpt-4-turbo",  # Use GPT-4 Turbo
        messages=[
            {"role": "system", "content": 
                "Extract only the official company name from this document. "
                "Do not include any extra words, addresses, or descriptions."},
            {"role": "user", "content": text}
        ]
    )

    print("🏢 Extracted Company Name:", response.choices[0].message.content.strip())  # Debugging
    return response.choices[0].message.content.strip()

def create_word_file(company_name):
    """Create a Word file with extracted company name."""
    doc = Document()
    doc.add_heading("Extracted Information", level=1)
    doc.add_paragraph(f"Company Name: {company_name}")

    file_path = "output.docx"

    # Ensure no previous file is interfering
    if os.path.exists(file_path):
        os.remove(file_path)

    doc.save(file_path)
    return file_path

@app.post("/upload/")
async def upload_pdf(file: UploadFile = File(...)):
    try:
        # Save uploaded file locally
        pdf_path = "latest_uploaded.pdf"
        with open(pdf_path, "wb") as f:
            f.write(await file.read())

        print("✅ New PDF uploaded:", pdf_path)

        # Extract text from PDF
        extracted_text = extract_text_from_pdf(pdf_path)

        # Send extracted text to OpenAI
        company_name = get_company_name(extracted_text)

        # Generate Word file
        word_path = create_word_file(company_name)
        print("📄 Generated Word File at:", word_path)

        return {"message": "Success", "download_url": "http://127.0.0.1:8000/download/"}
    except Exception as e:
        print("❌ Error:", str(e))
        return {"error": str(e)}

@app.get("/download/")
async def download_file():
    file_path = "output.docx"
    if not os.path.exists(file_path):
        return {"error": "File not found. Upload a PDF first."}
    
    return FileResponse(
        path=file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="extracted_info.docx"
    )
