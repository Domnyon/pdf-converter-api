from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pdf2docx import Converter
import os
import shutil
import uuid

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
def home():
    return {"status": "High-Precision PDF Converter API is Running!"}

@app.post("/convert")
async def convert_pdf_to_word(file: UploadFile = File(...)):
    unique_id = str(uuid.uuid4())
    pdf_path = f"/tmp/{unique_id}_{file.filename}"
    docx_path = f"/tmp/{unique_id}_converted.docx"

    with open(pdf_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        cv = Converter(pdf_path)

        # ضبط المحرك المتقدم للحفاظ على المربعات، الأشكال، الرموز، والجداول
        cv.convert(
            docx_path,
            start=0,
            end=None,
            multi_processing=True,
            connected_components=True,     # اكتشاف ودمج الأشكال والمربعات الرسومية
            extract_stream=True,           # استخراج الرموز والشعارات الرسومية Vector
            parse_lattice_tables=True,    # قراءة الجداول ذات الحدود والمربعات بدقة
            parse_stream_tables=True      # قراءة الجداول والقوائم غير المحددة بإطار
        )
        cv.close()

        return FileResponse(
            docx_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=file.filename.replace(".pdf", ".docx")
        )
    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
