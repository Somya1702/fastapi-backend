from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from fastapi.responses import FileResponse, JSONResponse
import os
import fitz  # PyMuPDF for extracting text from PDFs
import openai

app = FastAPI(
    title="PDF Extraction API",
    description="Upload a PDF, provide a custom GPT prompt, and generate a Word file.",
    version="1.1.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

# ✅ Fix CORS Issue
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], 
    allow_credentials=True,
    allow_methods=["*"], 
    allow_headers=["*"]
)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

def extract_text_from_pdf(pdf_path):
    """Extracts text from a PDF file"""
    text = ""
    pdf_document = fitz.open(pdf_path)
    for page in pdf_document:
        text += page.get_text("text") + "\n"
    return text[:3000]  # Trim text to first 3000 characters

def query_gpt(prompt, extracted_text):
    """Send the extracted text along with a user-defined prompt to GPT"""
    client = openai.OpenAI(api_key=OPENAI_API_KEY)

    response = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[
            {"role": "system", "content": "You are a helpful assistant that extracts structured information from PDFs."},
            {"role": "user", "content": f"{prompt}\n\nText from PDF:\n{extracted_text}"}
        ]
    )

    return response.choices[0].message.content.strip()

def create_word_file(extracted_info):
    """Create a Word file with extracted details."""
    doc = Document()
    doc.add_heading("Extracted Information", level=1)
    doc.add_paragraph(extracted_info)

    file_path = "output.docx"
    if os.path.exists(file_path):
        os.remove(file_path)

    doc.save(file_path)
    return file_path

@app.post("/upload/")
async def upload_pdf(file: UploadFile = File(...), prompt: str = Form(...)):
    try:
        # Save uploaded file locally
        pdf_path = "latest_uploaded.pdf"
        with open(pdf_path, "wb") as f:
            f.write(await file.read())

        print("✅ New PDF uploaded:", pdf_path)

        # Extract text from PDF
        extracted_text = extract_text_from_pdf(pdf_path)

        # Send extracted text to GPT with the custom prompt
        extracted_info = query_gpt(prompt, extracted_text)

        # Generate Word file
        word_path = create_word_file(extracted_info)
        print("📄 Generated Word File at:", word_path)

        return JSONResponse(content={
            "message": "Success",
            "extracted_info": extracted_info,
            "download_url": "https://fastapi-backend-f2mt.onrender.com/download/"
        })
    except Exception as e:
        print("❌ Error:", str(e))
        return JSONResponse(content={"error": str(e)}, status_code=500)

@app.get("/download/")
async def download_file():
    file_path = "output.docx"
    if not os.path.exists(file_path):
        return JSONResponse(content={"error": "File not found. Upload a PDF first."}, status_code=404)
    
    return FileResponse(
        path=file_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename="extracted_info.docx"
    )
