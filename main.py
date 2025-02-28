from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from docx import Document
from docx.shared import Pt, RGBColor
from fastapi.responses import FileResponse, JSONResponse
import os
import fitz  
import openai
import re  

app = FastAPI(
    title="PDF to Word API",
    description="Upload a PDF, enter a custom prompt, and generate a Word file.",
    version="1.0.2",
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

def create_word_file(response_text):
    doc = Document()
    title = doc.add_heading("Extracted Information", level=1)
    title.runs[0].font.name = "Roboto"
    title.runs[0].font.size = Pt(14)
    title.runs[0].font.color.rgb = RGBColor(0, 0, 0)

    paragraph = doc.add_paragraph(response_text)
    run = paragraph.runs[0]
    run.font.name = "Roboto"
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0, 0, 0)

    file_path = "output.docx"
    if os.path.exists(file_path):
        os.remove(file_path)
    doc.save(file_path)
    return file_path

@app.post("/upload/")
async def upload_pdf(file: UploadFile = File(...), prompt: str = Form(...)):
    pdf_path = "latest_uploaded.pdf"
    with open(pdf_path, "wb") as f:
        f.write(await file.read())

    extracted_text = extract_text_from_pdf(pdf_path)
    gstin = extract_gstin(extracted_text)
    
    gpt_response = get_gpt_response(extracted_text, prompt)
    word_path = create_word_file(gpt_response)

    return JSONResponse(content={
        "message": "Success",
        "gstin": gstin,
        "extracted_text": gpt_response,
        "download_url": "/download/"
    })

@app.get("/download/")
async def download_file():
    file_path = "output.docx"
    return FileResponse(file_path, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="extracted_info.docx")
