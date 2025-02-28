from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime
from fastapi.responses import FileResponse, JSONResponse
import os
import fitz  
import openai
import re  

app = FastAPI(
    title="PDF to Word API",
    description="Upload a PDF and a Word file with GPT prompt to generate a formatted Word file.",
    version="1.0.4",
    docs_url="/docs",
    redoc_url="/redoc"
)

@app.get("/")
def serve_frontend():
    return FileResponse("index.html")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF document."""
    text = ""
    pdf_document = fitz.open(pdf_path)
    for page in pdf_document:
        text += page.get_text("text") + "\n"
    return text[:3000]  

def extract_gstin(text):
    """Extract GSTIN using regex pattern matching."""
    gstin_pattern = r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}\d{1}[Z]{1}[A-Z\d]{1}\b"
    matches = re.findall(gstin_pattern, text)
    return matches[0] if matches else "Not Found"

def extract_text_from_word(word_path):
    """Extract text from the uploaded Word file."""
    doc = Document(word_path)
    extracted_text = "\n".join([para.text for para in doc.paragraphs])
    return extracted_text.strip()

def get_gpt_response(text, user_prompt):
    """Send extracted text and user prompt to OpenAI for processing."""
    client = openai.OpenAI(api_key=OPENAI_API_KEY)
    response = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[{"role": "system", "content": user_prompt}, {"role": "user", "content": text}]
    )
    return response.choices[0].message.content.strip() or "Not Found"

def create_word_file(response_text):
    """Generate a structured Word file with correct formatting."""
    doc = Document()

    # ✅ Add Letterhead (Centered, Bold, Italic, Yellow Highlight)
    letterhead = doc.add_paragraph()
    letterhead_run = letterhead.add_run("<LETTERHEAD>")
    letterhead_run.bold = True
    letterhead_run.italic = True
    letterhead_run.font.size = Pt(14)
    letterhead_run.font.color.rgb = RGBColor(0, 0, 0)
    letterhead.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # ✅ Add "To," on Left and "Date:" on Right
    to_date_para = doc.add_paragraph()
    to_run = to_date_para.add_run("To,")
    to_run.bold = True
    to_date_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    date_run = to_date_para.add_run("\t\t\t\t\tDate: " + datetime.today().strftime("%d-%m-%Y"))
    date_run.bold = True

    # ✅ Add Commissioner Address (3 lines, left-aligned)
    commissioner_para = doc.add_paragraph("Commissioner of GST & Central Excise,\nCity Name,\nState - Pin Code.")
    commissioner_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # ✅ Add Assessee (Taxpayer) Details (Centered, 3 lines)
    doc.add_paragraph()  # Empty line
    assessee_para = doc.add_paragraph()
    assessee_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    assessee_para.add_run(f"Assessee Name\nAssessee Address\nGSTIN: {extract_gstin(response_text)}").bold = True

    # ✅ Add Subject Line (Bold)
    doc.add_paragraph("\nSubject: Reply to Show Cause Notice", style="Heading 2")

    # ✅ Add "Sir," Left Aligned
    doc.add_paragraph("\nSir,")

    # ✅ Add "BRIEF FACTS OF THE CASE" in Center with Bold and Underline
    facts_heading = doc.add_paragraph()
    facts_run = facts_heading.add_run("BRIEF FACTS OF THE CASE")
    facts_run.bold = True
    facts_run.underline = True
    facts_heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # ✅ Add Case Facts in Justified Paragraph with Numbering
    case_facts_para = doc.add_paragraph()
    case_facts_para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    case_facts_list = response_text.split("\n")
    for idx, fact in enumerate(case_facts_list, start=1):
        case_facts_para.add_run(f"{idx}. {fact.strip()}\n")

    # ✅ Save the Word File
    file_path = "output.docx"
    if os.path.exists(file_path):
        os.remove(file_path)
    doc.save(file_path)
    return file_path

@app.post("/upload/")
async def upload_pdf(
    pdf_file: UploadFile = File(...), 
    word_file: UploadFile = File(...)
):
    """Upload and process a PDF and a Word file to extract structured data."""
    pdf_path = "latest_uploaded.pdf"
    word_path = "uploaded_prompt.docx"

    # ✅ Save uploaded files
    with open(pdf_path, "wb") as f:
        f.write(await pdf_file.read())

    with open(word_path, "wb") as f:
        f.write(await word_file.read())

    extracted_text = extract_text_from_pdf(pdf_path)

    # ✅ Extract prompt from Word file
    extracted_prompt = extract_text_from_word(word_path)

    # ✅ Ask GPT to extract case facts using the extracted prompt
    gpt_response = get_gpt_response(extracted_text, extracted_prompt)
    word_path = create_word_file(gpt_response)

    return JSONResponse(content={
        "message": "Success",
        "extracted_text": gpt_response,
        "download_url": "/download/"
    })

@app.get("/download/")
async def download_file():
    """Download the generated Word file."""
    file_path = "output.docx"
    return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="extracted_info.docx")
