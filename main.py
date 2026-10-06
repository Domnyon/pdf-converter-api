from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import requests
import time
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

# بيانات الاعتماد الرسمية من أدوبي
CLIENT_ID = "84c6437aa8a346a086a2513dd702e45f"
CLIENT_SECRET = "p8e-ZTIVq_Nj4O6vVkhu3CTGvrU_l6-Uaaem"

def get_adobe_access_token():
    """الحصول على توكن الوصول المباشر من أدوبي"""
    url = "https://pdf-services.adobe.io/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET
    }
    response = requests.post(url, headers=headers, data=data)
    if response.status_code != 200:
        raise Exception(f"Failed to authenticate with Adobe: {response.text}")
    return response.json()["access_token"]

@app.get("/")
def home():
    return {"status": "Official Adobe Direct REST Engine is Live!"}

@app.post("/convert")
async def convert_pdf_to_word(file: UploadFile = File(...)):
    unique_id = str(uuid.uuid4())
    pdf_path = f"/tmp/{unique_id}_{file.filename}"
    docx_path = f"/tmp/{unique_id}_converted.docx"

    with open(pdf_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        token = get_adobe_access_token()
        headers = {
            "x-api-key": CLIENT_ID,
            "Authorization": f"Bearer {token}"
        }

        # 1. طلب رابط رفع الملف من أدوبي
        upload_init_url = "https://pdf-services.adobe.io/assets"
        upload_init_res = requests.post(
            upload_init_url,
            headers={**headers, "Content-Type": "application/json"},
            json={"mediaType": "application/pdf"}
        )
        if upload_init_res.status_code not in (200, 201):
            raise Exception(f"Asset creation failed: {upload_init_res.text}")

        asset_data = upload_init_res.json()
        upload_uri = asset_data["uploadUri"]
        asset_id = asset_data["assetID"]

        # 2. رفع ملف الـ PDF فعلياً إلى خوادم أدوبي
        with open(pdf_path, "rb") as f:
            upload_file_res = requests.put(
                upload_uri,
                headers={"Content-Type": "application/pdf"},
                data=f
            )
        if upload_file_res.status_code not in (200, 201):
            raise Exception("Failed to upload document to Adobe storage.")

        # 3. بدء وظيفة التحويل إلى DOCX
        export_job_url = "https://pdf-services.adobe.io/operation/exportpdf"
        job_payload = {
            "assetID": asset_id,
            "targetFormat": "docx"
        }
        job_res = requests.post(
            export_job_url,
            headers={**headers, "Content-Type": "application/json"},
            json=job_payload
        )
        if job_res.status_code != 201:
            raise Exception(f"Failed to trigger export job: {job_res.text}")

        # رابط متابعة حالة العملية (Polling URL)
        status_url = job_res.headers.get("location")

        # 4. انتظار انتهاء أدوبي من التحويل
        download_uri = None
        for _ in range(60):  # محاولة كل ثانيتين حتى دقيقتين كحد أقصى
            time.sleep(2)
            check_res = requests.get(status_url, headers=headers)
            check_data = check_res.json()
            status = check_data.get("status")

            if status == "done":
                download_uri = check_data["asset"]["downloadUri"]
                break
            elif status == "failed":
                raise Exception(f"Adobe conversion failed: {check_data}")

        if not download_uri:
            raise Exception("Adobe conversion timed out.")

        # 5. تنزيل مستند الـ Word المحول
        doc_res = requests.get(download_uri)
        with open(docx_path, "wb") as f:
            f.write(doc_res.content)

        return FileResponse(
            docx_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=file.filename.replace(".pdf", ".docx")
        )

    except Exception as e:
        print(f"Error: {e}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
