from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from datetime import datetime
from fastapi.responses import FileResponse, JSONResponse
import os
import fitz  
import openai
import re  

app = FastAPI(
    title="PDF to Word API",
    description="Upload a PDF, enter a custom prompt, and generate a Word file.",
    version="1.0.3",
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
    text = ""
    pdf_document = fitz.open(pdf_path)
    for page in pdf_document:
        text += page.get_text("text") + "\n"
    return text[:3000]  

def extract_gstin(text):
    gstin_pattern = r"\b\d{2}[A-Z]{5}\d{4}[A-Z]{1}\d{1}[Z]{1}[A-Z\d]{1}\b"
    matches = re.findall(gstin_pattern, text)
    return matches[0] if matches else "Not Found"

def get_gpt_response(text, user_prompt):
    client = openai.OpenAI(api_key=OPENAI_API_KEY)
    response = client.chat.completions.create(
        model="gpt-4-turbo",
        messages=[{"role": "system", "content": user_prompt}, {"role": "user", "content": text}]
    )
    return response.choices[0].message.content.strip() or "Not Found"

def create_word_file(assessee_name, assessee_address, gstin, case_facts):
    """Create a structured Word file based on provided formatting guidelines."""
    doc = Document()

    # ✅ Add Letterhead in center with bold, italics, and yellow highlight
    letterhead = doc.add_paragraph()
    letterhead_run = letterhead.add_run("<LETTERHEAD>")
    letterhead_run.bold = True
    letterhead_run.italic = True
    letterhead_run.font.size = Pt(14)
    letterhead_run.font.color.rgb = RGBColor(0, 0, 0)
    
    # Apply yellow highlight
    highlight = OxmlElement("w:highlight")
    highlight.set("w:val", "yellow")
    letterhead_run._r.get_or_add_rPr().append(highlight)
    
    letterhead.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # ✅ Add "To," on the left and "Date:" on the right
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
    assessee_para.add_run(f"{assessee_name}\n{assessee_address}\nGSTIN: {gstin}").bold = True

    # ✅ Add Subject Line (Bold)
    doc.add_paragraph("\nSubject: Reply to Show Cause Notice", style="Heading 2")

    # ✅ Add "Sir," Left Aligned
    doc.add_paragraph("\nSir,")

    # ✅ Add "BRIEF FACTS OF THE CASE" in center with bold and underline
    facts_heading = doc.add_paragraph()
    facts_run = facts_heading.add_run("BRIEF FACTS OF THE CASE")
    facts_run.bold = True
    facts_run.underline = True
    facts_heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # ✅ Add Case Facts in Justified Paragraph with Numbering
    case_facts_para = doc.add_paragraph()
    case_facts_para.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    case_facts_list = case_facts.split("\n")
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
    file: UploadFile = File(...), 
    prompt: str = Form(...)
):
    pdf_path = "latest_uploaded.pdf"
    with open(pdf_path, "wb") as f:
        f.write(await file.read())

    extracted_text = extract_text_from_pdf(pdf_path)
    gstin = extract_gstin(extracted_text)

    # ✅ Ask GPT to extract assessee details & case facts
    assessee_details = get_gpt_response(extracted_text, "Extract the taxpayer's name and address from the document.")
    case_facts = get_gpt_response(extracted_text, "Extract the facts of the case in a structured paragraph format.")

    # ✅ Parse Assessee Details
    assessee_name, assessee_address = assessee_details.split("\n")[0], "\n".join(assessee_details.split("\n")[1:])

    word_path = create_word_file(assessee_name, assessee_address, gstin, case_facts)

    return JSONResponse(content={
        "message": "Success",
        "gstin": gstin,
        "assessee_name": assessee_name,
        "assessee_address": assessee_address,
        "extracted_text": case_facts,
        "download_url": "/download/"
    })

@app.get("/download/")
async def download_file():
    file_path = "output.docx"
    return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="extracted_info.docx")
