from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pdf2docx import Converter
import arabic_reshaper
from bidi.algorithm import get_display
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
    return {"status": "Arabic BiDi & Layout Enhanced PDF Engine Live!"}

@app.post("/convert")
async def convert_pdf_to_word(file: UploadFile = File(...)):
    unique_id = str(uuid.uuid4())
    pdf_path = f"/tmp/{unique_id}_{file.filename}"
    docx_path = f"/tmp/{unique_id}_converted.docx"

    with open(pdf_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        cv = Converter(pdf_path)

        # خيارات متقدمة لإجبار إنشاء مربعات نصية وحفظ الحدود والتنسيق الهندسي
        cv.convert(
            docx_path,
            start=0,
            end=None,
            multi_processing=True,
            connected_components=True,   # دمج ورسم المربعات النصية وحقول الإدخال
            extract_stream=True,         # الحفاظ على الأشكال الهندسية والرموز
            parse_lattice_tables=True,  # الحفاظ على إطارات الجداول والمربعات المغلقة
            parse_stream_tables=True,   # كشف القوائم والترتيب غير المحدد بإطار
            line_break_free=False,      # منع تدمير الأسطر لتفادي دمج نصوص المربعات المستقلة
            margin_tolerance=0.05,      # دقة متناهية في التقاط المسافات وهوامش المربعات
            text_direction_rtl=True     # ضبط اتجاه الكتابة من اليمين لليسار العربي
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
