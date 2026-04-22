from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
import openpyxl
import io
from datetime import datetime

app = FastAPI()

# ✅ CORS 설정 (핵심)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://daesung-so.github.io"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/process-xlsm")
async def process_xlsm(file: UploadFile = File(...)):
    try:
        contents = await file.read()

        # 매크로 유지해서 엑셀 열기
        wb = openpyxl.load_workbook(
            filename=io.BytesIO(contents),
            keep_vba=True
        )

        # 테스트용 시트 생성
        ws = wb.create_sheet("DATA")

        now = datetime.now()
        ws["A1"] = "서버처리시간"
        ws["B1"] = now.strftime("%Y-%m-%d %H:%M")

        # 메모리에 다시 저장
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        # 파일명 생성
        filename = f"일일점검사항_v8__{now.strftime('%Y%m%d_%H시')}.xlsm"

        return StreamingResponse(
            output,
            media_type="application/vnd.ms-excel",
            headers={
                "Content-Disposition": f"attachment; filename={filename}"
            }
        )

    except Exception as e:
        return JSONResponse(
            {"error": str(e)},
            status_code=500
        )