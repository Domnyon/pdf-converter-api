import os
import uuid
import traceback
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

# استيراد محرك Adobe بالطريقة الرسمية والمضمونة
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat

# مكتبة المعالجة المحلية كخطة بديلة
from pdf2docx import Converter

import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = FastAPI(title="APDF Processing Engine")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
@app.head("/")
async def root():
    return {"status": "ok", "message": "APDF API is active"}

def get_adobe_services():
    client_id = os.environ.get("PDF_SERVICES_CLIENT_ID")
    client_secret = os.environ.get("PDF_SERVICES_CLIENT_SECRET")
    
    if not client_id or not client_secret:
        raise ValueError("Missing Adobe credentials in Environment Variables")
        
    credentials = ServicePrincipalCredentials(
        client_id=client_id.strip(),
        client_secret=client_secret.strip()
    )
    return PDFServices(credentials=credentials)

@app.post("/convert")
async def convert_pdf(file: UploadFile = File(...)):
    unique_id = str(uuid.uuid4())
    pdf_path = f"/tmp/{unique_id}.pdf"
    docx_path = f"/tmp/{unique_id}.docx"

    try:
        content = await file.read()
        with open(pdf_path, "wb") as f:
            f.write(content)

        adobe_success = False

        # 1. محاولة التحويل عبر Adobe
        try:
            print("==> Trying Adobe PDF Services...")
            pdf_services = get_adobe_services()
            
            with open(pdf_path, "rb") as input_file:
                input_asset = pdf_services.upload(
                    input_stream=input_file,
                    mime_type=PDFServicesMediaType.PDF
                )

            export_params = ExportPDFParams(target_format=ExportPDFTargetFormat.DOCX)
            export_job = ExportPDFJob(input_asset=input_asset, export_pdf_params=export_params)
            
            location = pdf_services.submit(export_job)
            # استخراج النتيجة دون الحاجة لكلاسات فرعية
            pdf_services_response = pdf_services.get_job_result(location, None)
            result_asset = pdf_services_response.get_result().get_asset()
            stream_asset = pdf_services.get_content(result_asset)

            with open(docx_path, "wb") as output_file:
                output_file.write(stream_asset.get_data_bytes())
                
            print("==> Adobe conversion succeeded!")
            adobe_success = True
        except Exception as adobe_err:
            print(f"⚠️ Adobe failed: {adobe_err}")
            traceback.print_exc()

        # 2. في حال فشل Adobe لأي سبب، يتم التحويل فوراً بالمحرك البديل
        if not adobe_success or not os.path.exists(docx_path):
            print("==> Falling back to native converter...")
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
            print("==> Native converter completed.")

        return FileResponse(
            docx_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=file.filename.replace(".pdf", ".docx")
        )

    except Exception as e:
        print(f"❌ Error during conversion: {str(e)}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except:
                pass

class ExportRequest(BaseModel):
    title: str = "مستند_معدل"
    content: str
    format: str = "docx"

def apply_rtl(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

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
