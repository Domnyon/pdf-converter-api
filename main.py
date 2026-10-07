import os
import uuid
import time
import requests
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from pdf2docx import Converter

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

def get_adobe_access_token(client_id: str, client_secret: str) -> str:
    """الحصول على توكن المصادقة من Adobe"""
    url = "https://ims-na1.adobelogin.com/ims/token/v3"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "client_credentials",
        "scope": "openid,AdobeID,read_organizations"
    }
    res = requests.post(url, headers=headers, data=data, timeout=15)
    res.raise_for_status()
    return res.json().get("access_token")

def convert_with_adobe_rest(pdf_bytes: bytes, client_id: str, client_secret: str) -> bytes:
    """تحويل مستند PDF إلى Word عبر Adobe REST API الرسمي المباشر"""
    token = get_adobe_access_token(client_id, client_secret)

    # 1. طلب رابط رفع من Adobe
    upload_url_req = "https://pdf-services.adobe.io/assets"
    headers = {
        "X-API-Key": client_id,
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    upload_res = requests.post(upload_url_req, headers=headers, json={"mediaType": "application/pdf"}, timeout=15)
    upload_res.raise_for_status()
    upload_data = upload_res.json()
    upload_uri = upload_data["uploadUri"]
    asset_id = upload_data["assetID"]

    # 2. رفع ملف الـ PDF
    put_res = requests.put(upload_uri, headers={"Content-Type": "application/pdf"}, data=pdf_bytes, timeout=30)
    put_res.raise_for_status()

    # 3. إرسال أمر التحويل إلى DOCX
    job_url = "https://pdf-services.adobe.io/operation/exportpdf"
    job_payload = {
        "assetID": asset_id,
        "targetFormat": "docx"
    }
    job_res = requests.post(job_url, headers=headers, json=job_payload, timeout=15)
    job_res.raise_for_status()
    poll_location = job_res.headers.get("Location")

    # 4. انتظار انتهاء التحويل وتنزيل النتيجة
    for _ in range(30):
        time.sleep(2)
        poll_res = requests.get(poll_location, headers=headers, timeout=15)
        if poll_res.status_code == 200:
            status = poll_res.json().get("status")
            if status == "done":
                download_uri = poll_res.json()["asset"]["downloadUri"]
                final_res = requests.get(download_uri, timeout=30)
                final_res.raise_for_status()
                return final_res.content
            elif status == "failed":
                raise Exception("Adobe job returned failed status")

    raise Exception("Adobe conversion timed out")

@app.post("/convert")
async def convert_pdf(file: UploadFile = File(...)):
    unique_id = str(uuid.uuid4())
    pdf_path = f"/tmp/{unique_id}.pdf"
    docx_path = f"/tmp/{unique_id}.docx"

    try:
        content = await file.read()
        converted = False

        client_id = os.environ.get("PDF_SERVICES_CLIENT_ID")
        client_secret = os.environ.get("PDF_SERVICES_CLIENT_SECRET")

        # المحاولة عبر محرك Adobe الرسمي
        if client_id and client_secret:
            try:
                print("==> Processing via Adobe REST API...")
                docx_bytes = convert_with_adobe_rest(content, client_id.strip(), client_secret.strip())
                with open(docx_path, "wb") as f:
                    f.write(docx_bytes)
                converted = True
                print("==> Adobe conversion successful!")
            except Exception as ad_err:
                print(f"⚠️ Adobe REST failed: {ad_err}, switching to fallback...")

        # الخطة البديلة في حال تعذر Adobe
        if not converted:
            print("==> Processing via native fallback...")
            with open(pdf_path, "wb") as f:
                f.write(content)
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
            print("==> Fallback completed.")

        return FileResponse(
            docx_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=file.filename.replace(".pdf", ".docx")
        )

    except Exception as e:
        print(f"❌ Conversion failed completely: {str(e)}")
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
