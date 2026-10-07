from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
import uuid

app = FastAPI(title="APDF Processing Engine")

# 🔒 حماية السيرفر: السماح فقط لنطاق موقعك الرسمي وبيئة التطوير المحلية
origins = [
    "https://apdf.app",
    "https://www.apdf.app",
    "http://localhost:3000",
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

class ExportRequest(BaseModel):
    title: str = "مستند_معدل"
    content: str
    format: str = "docx"  # docx أو pdf

def apply_rtl(paragraph):
    """فرض اتجاه الكتابة من اليمين لليسار في وورد لمنع تداخل الأحرف"""
    pPr = paragraph._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

@app.post("/export-doc")
async def export_document(payload: ExportRequest):
    try:
        doc = docx.Document()

        # ضبط هوامش A4 قياسية
        for section in doc.sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.8)
            section.right_margin = Inches(0.8)

        # تقسيم الأسطر وبناء الفقرات بهوامش عربية مريحة
        lines = payload.content.split("\n")
        for line in lines:
            line_str = line.strip()
            if not line_str:
                continue

            p = doc.add_paragraph()
            apply_rtl(p)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.line_spacing = 1.3  # تباعد أسطر يمنع ملامسة الحروف
            p.paragraph_format.space_after = Pt(4)

            run = p.add_run(line_str)
            run.font.name = "Arial"  # خط قياسي مدعوم رسمياً
            run.font.size = Pt(13)
            run._element.get_or_add_rPr().set(qn('w:rtl'), '1')

        unique_id = str(uuid.uuid4())
        file_path = f"/tmp/{unique_id}.docx"
        doc.save(file_path)

        return FileResponse(
            file_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=f"{payload.title}.docx"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"حدث خطأ أثناء تصدير المستند: {str(e)}")
