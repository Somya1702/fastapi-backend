from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from fastapi.responses import FileResponse, JSONResponse
import os
import fitz  # PyMuPDF for extracting text from PDFs
import openai
import re  # For extracting GSTIN

app = FastAPI(  # Enable Swagger UI
    title="PDF to Word API",
    description="Upload a PDF, extract company details, and generate a Word file.",
    version="1.0.1",
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

    # Trim text to first 3000 characters (adjusted for better GPT processing)
    cleaned_text = text[:3000]  
    print("📄 Extracted Text (Trimmed):", cleaned_text[:500])  # Debugging
    return cleaned_text

def extract_gstin(text):
    """Extract GSTIN from the text using regex, ensuring correct format."""
    gstin_pattern = r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}\d{1}[Z]{1}[A-Z\d]{1}\b"  # GSTIN regex pattern
    gstin_matches = re.findall(gstin_pattern, text)

    if gstin_matches:
        print("✅ GSTIN Found:", gstin_matches[0])
        return gstin_matches[0]  # Return the first valid GSTIN found

    print("❌ No GSTIN Found")
    return "Not Found"

def get_company_details(text):
    """Send extracted text to OpenAI to get the correct company name and address."""
    client = openai.OpenAI(api_key=OPENAI_API_KEY)

    response = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[
            {"role": "system", "content": 
                "Extract only the correct official company name and address from this document. "
                "Do not include extra words, descriptions, or placeholders. "
                "Provide the response in JSON format with 'company_name' and 'address' fields. "
                "If the company name or address is not found, return 'Not Found'."},
            {"role": "user", "content": text}
        ]
    )

    try:
        extracted_data = response.choices[0].message.content.strip()
        print("🏢 Extracted Data (Raw):", extracted_data)  # Debugging

        # Ensure response is in JSON format
        extracted_data = eval(extracted_data) if "{" in extracted_data else {"company_name": "Not Found", "address": "Not Found"}
    except Exception as e:
        print("❌ Error parsing GPT response:", str(e))
        extracted_data = {"company_name": "Not Found", "address": "Not Found"}

    return extracted_data

def create_word_file(company_name, gstin, company_address):
    """Create a Word file with extracted details."""
    doc = Document()
    doc.add_heading("Extracted Information", level=1)
    doc.add_paragraph(f"Company Name: {company_name}")
    doc.add_paragraph(f"GSTIN: {gstin if gstin else 'Not Found'}")
    doc.add_paragraph(f"Address: {company_address if company_address else 'Not Found'}")

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

        # Extract GSTIN
        gstin = extract_gstin(extracted_text)

        # Send extracted text to OpenAI for Name & Address
        company_details = get_company_details(extracted_text)
        
        # Parse extracted details (Assuming JSON format from GPT)
        company_name = company_details.get("company_name", "Not Found")
        company_address = company_details.get("address", "Not Found")

        # Generate Word file
        word_path = create_word_file(company_name, gstin, company_address)
        print("📄 Generated Word File at:", word_path)

        return JSONResponse(content={
            "message": "Success",
            "company_name": company_name,
            "gstin": gstin,
            "company_address": company_address,
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
