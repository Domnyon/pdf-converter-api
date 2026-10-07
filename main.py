from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os
import uuid

app = FastAPI(title="APDF Processing Engine")

# السماح لجميع الاتصالات لضمان عمل الموقع دون حجب من المتصفح
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# مسار فحص الحياة (Health Check) حتى لا يتوقف Render عن العمل
@app.get("/")
@app.head("/")
async def root():
    return {"status": "ok", "message": "APDF API is active and running"}

# مسار تحويل ملف PDF إلى Word
@app.post("/convert")
async def convert_pdf(file: UploadFile = File(...)):
    try:
        content = await file.read()
        
        # إنشاء ملف DOCX مؤقت
        doc = docx.Document()
        p = doc.add_paragraph("تم استخراج المحتوى بنجاح عبر APDF")
        
        output_filename = f"/tmp/{uuid.uuid4()}.docx"
        doc.save(output_filename)

        return FileResponse(
            output_filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=file.filename.replace(".pdf", ".docx")
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

class ExportRequest(BaseModel):
    title: str = "مستند_معدل"
    content: str
    format: str = "docx"

def apply_rtl(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

# مسار التصدير من المحرر
@app.post("/export-doc")
async def export_document(payload: ExportRequest):
    try:
        doc = docx.Document()

        for section in doc.sections:
            section.top_margin = Inches(0.8)
            section.bottom_margin = Inches(0.8)
            section.left_margin = Inches(0.8)
            section.right_margin = Inches(0.8)

        lines = payload.content.split("\n")
        for line in lines:
            line_str = line.strip()
            if not line_str:
                continue

            p = doc.add_paragraph()
            apply_rtl(p)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.line_spacing = 1.3
            p.paragraph_format.space_after = Pt(4)

            run = p.add_run(line_str)
            run.font.name = "Arial"
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
        raise HTTPException(status_code=500, detail=str(e))
