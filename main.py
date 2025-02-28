from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from docx.shared import Pt, RGBColor
from fastapi.responses import FileResponse, JSONResponse
import os
import fitz  # PyMuPDF for extracting text from PDFs
import openai
import re  # For extracting GSTIN

app = FastAPI(  # Enable Swagger UI
    title="PDF to Word API",
    description="Upload a PDF, enter a custom prompt, and generate a Word file.",
    version="1.0.2",
    docs_url="/docs",
    redoc_url="/redoc"
)

# Serve the frontend when visiting the root URL
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

    # Trim text to first 3000 characters for better GPT processing
    cleaned_text = text[:3000]  
    print("📄 Extracted Text (Trimmed):", cleaned_text[:500])  # Debugging
    return cleaned_text

def extract_gstin(text):
    """Extract GSTIN from the text using regex"""
    gstin_pattern = r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}\d{1}[Z]{1}[A-Z\d]{1}\b"
    gstin_matches = re.findall(gstin_pattern, text)

    if gstin_matches:
        print("✅ GSTIN Found:", gstin_matches[0])
        return gstin_matches[0]  # Return first valid GSTIN found

    print("❌ No GSTIN Found")
    return "Not Found"

def get_gpt_response(text, user_prompt):
    """Send extracted text and user prompt to OpenAI"""
    client = openai.OpenAI(api_key=OPENAI_API_KEY)

    response = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[
            {"role": "system", "content": user_prompt},  # User-defined prompt
            {"role": "user", "content": text}
        ]
    )

    extracted_data = response.choices[0].message.content.strip()
    print("🔍 Extracted Data:", extracted_data)  # Debugging

    return extracted_data

def create_word_file(response_text):
    """Create a styled Word file with the extracted response."""
    doc = Document()
    
    # Set document title
    title = doc.add_heading("Extracted Information", level=1)
    title.runs[0].font.name = "Roboto"
    title.runs[0].font.size = Pt(14)
    title.runs[0].font.color.rgb = RGBColor(0, 0, 0)  # Black color
    
    # Add extracted content
    paragraph = doc.add_paragraph(response_text)
    run = paragraph.runs[0]
    run.font.name = "Roboto"
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0, 0, 0)  # Black color

    file_path = "output.docx"

    # Ensure no previous file is interfering
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

        # Extract GSTIN
        gstin = extract_gstin(extracted_text)

        # Send extracted text and user-defined prompt to OpenAI
        gpt_response = get_gpt_response(extracted_text, prompt)
        
        # Generate Word file with the extracted response
        word_path = create_word_file(gpt_response)
        print("📄 Generated Word File at:", word_path)

        return JSONResponse(content={
            "message": "Success",
            "gstin": gstin,
            "extracted_text": gpt_response,
            "download_url": "http://127.0.0.1:8000/download/"
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
