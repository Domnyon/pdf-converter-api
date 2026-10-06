from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import os
import shutil
import uuid

# استيراد محركاتك المخصصة من الملفات التي أنشأتها
from arabic_engine import extract_and_sort_arabic_lines, is_arabic_text
from form_detector import extract_form_checkboxes_and_labels

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def set_cell_rtl(cell):
    """ضبط اتجاه خلية الجدول من اليمين لليسار"""
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(r'<w:tcMar %s><w:right w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/></w:tcMar>' % nsdecls('w'))
    tcPr.append(tcMar)

@app.get("/")
def home():
    return {"status": "Engine v3 (Arabic BiDi + Native Checkboxes) Live!"}

@app.post("/convert")
async def convert_pdf_to_word(file: UploadFile = File(...)):
    unique_id = str(uuid.uuid4())
    pdf_path = f"/tmp/{unique_id}_{file.filename}"
    docx_path = f"/tmp/{unique_id}_converted.docx"

    with open(pdf_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        pdf_doc = fitz.open(pdf_path)
        doc = Document()

        # ضبط اتجاه وهوامش صفحة الوورد
        for section in doc.sections:
            section.top_margin = Inches(0.6)
            section.bottom_margin = Inches(0.6)
            section.left_margin = Inches(0.6)
            section.right_margin = Inches(0.6)

        for page_idx, page in enumerate(pdf_doc):
            if page_idx > 0:
                doc.add_page_break()

            # 1. استخراج حقول النماذج والمربعات عبر form_detector
            checkbox_fields = extract_form_checkboxes_and_labels(page)

            # 2. استخراج الأسطر المرتبة عربياً عبر arabic_engine
            arabic_lines = extract_and_sort_arabic_lines(page)

            # إنشاء جدول مخفي لتنظيم حقول الخيارات إذا وُجدت مربعات اختيار
            if checkbox_fields:
                table = doc.add_table(rows=1, cols=len(checkbox_fields))
                table.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                row_cells = table.rows[0].cells

                for idx, field in enumerate(checkbox_fields):
                    cell = row_cells[idx]
                    set_cell_rtl(cell)
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                    # إضافة رمز المربع التفاعلي متبوعاً بالتسمية الصحيحة
                    run_box = p.add_run("☐ ")
                    run_box.font.name = "Arial"
                    run_box.font.size = Pt(13)
                    run_box.font.bold = True

                    run_label = p.add_run(field["label"])
                    run_label.font.name = "Arial"
                    run_label.font.size = Pt(11)

                doc.add_paragraph()  # سطر فاصل

            # 3. كتابة وتنسيق الأسطر مع الحفاظ على الترتيب واللغة العربية
            for line in arabic_lines:
                if not line.strip():
                    continue

                # تخطي تكرار نصوص المربعات التي وُضعت في الجدول بالأعلى
                if any(field["label"] in line and field["label"] != "بدون وصف" for field in checkbox_fields):
                    continue

                p = doc.add_paragraph()
                
                # فحص الاتجاه: إذا كان عربي، نجعل المحاذاة لليمين
                if is_arabic_text(line):
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    run = p.add_run(line)
                    run.font.name = "Arial"
                    run.font.size = Pt(11.5)
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = p.add_run(line)
                    run.font.name = "Calibri"
                    run.font.size = Pt(11.5)

        pdf_doc.close()
        doc.save(docx_path)

        return FileResponse(
            docx_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=file.filename.replace(".pdf", ".docx")
        )
    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
