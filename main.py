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
    return {"status": "Official Forms Precision Engine Live!"}

@app.post("/convert")
async def convert_pdf_to_word(file: UploadFile = File(...)):
    unique_id = str(uuid.uuid4())
    pdf_path = f"/tmp/{unique_id}_{file.filename}"
    docx_path = f"/tmp/{unique_id}_converted.docx"

    with open(pdf_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        cv = Converter(pdf_path)

        # إعدادات خاصة بالنماذج وحقول الإدخال ومربعات الاختيار
        cv.convert(
            docx_path,
            start=0,
            end=None,
            multi_processing=False,        # إيقاف التوازي لتفادي لخبطة ترتيب الحقول والنصوص المتجاورة
            keep_shapes=True,              # الحفاظ على المربعات وحقول التحديد والأشكال الهندسية
            connected_components=True,     # ربط الحقول والنصوص المحيطة بها
            parse_lattice_tables=True,     # قراءة الجداول وإطارات النماذج
            parse_stream_tables=True,      # كشف الحقول المصفوفة أفقياً وعمودياً
            line_break_free=False,         # الحفاظ على فواصل الأسطر لمنع دمج العناوين مع التواقيع
            margin_tolerance=0.01          # أعلى حساسية للمسافات وهوامش الحقول
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
