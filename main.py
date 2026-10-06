from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import os
import shutil
import uuid

# استيراد حزمة Adobe الرسمية للبايثون
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfops.export_pdf_operation import ExportPDFOperation
from adobe.pdfservices.operation.pdfops.options.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfops.options.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.io.stream_asset import StreamAsset

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

@app.get("/")
def home():
    return {"status": "Official Adobe PDF Engine is Running!"}

@app.post("/convert")
async def convert_pdf_to_word(file: UploadFile = File(...)):
    unique_id = str(uuid.uuid4())
    pdf_path = f"/tmp/{unique_id}_{file.filename}"
    docx_path = f"/tmp/{unique_id}_converted.docx"

    with open(pdf_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    try:
        # 1. إعداد المصادقة عبر بيانات اعتماد حساب أدوبي
        credentials = ServicePrincipalCredentials(
            client_id=CLIENT_ID,
            client_secret=CLIENT_SECRET
        )

        pdf_services = PDFServices(credentials=credentials)

        # 2. رفع الملف إلى خدمة Adobe
        with open(pdf_path, "rb") as input_file_stream:
            input_asset = pdf_services.upload(
                input_stream=input_file_stream,
                mime_type=PDFServicesMediaType.PDF
            )

            # 3. إعداد عملية التصدير إلى DOCX
            export_params = ExportPDFParams(
                target_format=ExportPDFTargetFormat.DOCX
            )

            export_operation = ExportPDFOperation(
                input_asset=input_asset,
                export_pdf_params=export_params
            )

            # 4. تنفيذ التحويل واستلام النتيجة الرسمية
            job_id = pdf_services.submit(export_operation)
            response = pdf_services.get_job_result(job_id=job_id, result_type=ExportPDFOperation)
            
            result_asset = response.get_result()
            stream_asset = pdf_services.get_content(result_asset)

            # 5. حفظ المستند الناتج
            with open(docx_path, "wb") as output_file_stream:
                output_file_stream.write(stream_asset.get_input_stream())

        return FileResponse(
            docx_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=file.filename.replace(".pdf", ".docx")
        )

    except Exception as e:
        print(f"Error during Adobe conversion: {e}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
