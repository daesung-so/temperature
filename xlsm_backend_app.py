from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
import io
import os
import re
from datetime import datetime

import openpyxl
import requests


app = FastAPI()

# github.io -> render 요청 허용
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://daesung-so.github.io"],
    allow_methods=["*"],
    allow_headers=["*"],
)

WORKERS_API_URL = os.getenv(
    "WORKERS_API_URL",
    "https://so-schedule.sungsu.workers.dev/api/temperature"
)

FLOOR_SHEET_CANDIDATES = {
    "2층": ["2층", "2F", "2"],
    "3층": ["3층", "3F", "3"],
    "5층": ["5층", "5F", "5"],
    "6층": ["6층", "6F", "6"],
    "8층": ["8층", "8F", "8"],
}


@app.get("/")
def root():
    return {"ok": True, "message": "backend alive"}


@app.get("/health")
def health():
    return {"ok": True}


def normalize_text(value: str) -> str:
    if value is None:
        return ""
    return str(value).strip().replace(" ", "")


def find_sheet_name(workbook, floor_name: str):
    candidates = FLOOR_SHEET_CANDIDATES.get(floor_name, [floor_name])
    normalized_map = {normalize_text(ws.title): ws.title for ws in workbook.worksheets}

    for cand in candidates:
        key = normalize_text(cand)
        if key in normalized_map:
            return normalized_map[key]

    # 부분 일치 fallback
    for ws in workbook.worksheets:
        title_norm = normalize_text(ws.title)
        for cand in candidates:
            if normalize_text(cand) in title_norm:
                return ws.title
    return None


def get_latest_record():
    resp = requests.get(WORKERS_API_URL, timeout=20)
    resp.raise_for_status()
    data = resp.json()

    if not isinstance(data, list) or not data:
        raise ValueError("온습도 저장 데이터가 비어있습니다.")

    # createdAt / updatedAt / date+time 기준으로 최신값 찾기
    def sort_key(item):
        created = item.get("createdAt") or item.get("updatedAt") or ""
        date_str = item.get("date") or ""
        time_str = item.get("time") or ""
        return f"{created}|{date_str}|{time_str}"

    latest = sorted(data, key=sort_key)[-1]
    return latest


def parse_datetime_from_record(record: dict):
    date_str = str(record.get("date") or "").strip()
    time_str = str(record.get("time") or "").strip()

    if not date_str:
        now = datetime.now()
        return now

    # 예: 2026-04-22 / 2026.04.22 / 20260422
    date_str_clean = date_str.replace(".", "-").replace("/", "-").strip()

    dt = None
    candidates = []

    if time_str:
        candidates.extend([
            f"{date_str_clean} {time_str}",
        ])

    # 지원 포맷들
    formats = [
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H시",
        "%Y-%m-%d %H",
        "%Y%m%d %H:%M",
        "%Y%m%d %H시",
        "%Y%m%d %H",
        "%Y-%m-%d",
        "%Y%m%d",
    ]

    for candidate in candidates + [date_str_clean]:
        for fmt in formats:
            try:
                dt = datetime.strptime(candidate, fmt)
                return dt
            except Exception:
                pass

    # 마지막 fallback
    return datetime.now()


def make_output_filename(record: dict):
    dt = parse_datetime_from_record(record)
    return f"일일점검사항_v8__{dt.strftime('%Y%m%d_%H시')}.xlsm"


def copy_row_style_and_formulas(ws, src_row: int, dst_row: int):
    # 행 높이
    if ws.row_dimensions[src_row].height is not None:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height

    for col in range(1, ws.max_column + 1):
        src = ws.cell(src_row, col)
        dst = ws.cell(dst_row, col)

        if src.has_style:
            dst._style = src._style

        if src.number_format:
            dst.number_format = src.number_format

        if src.font:
            dst.font = src.font.copy()

        if src.fill:
            dst.fill = src.fill.copy()

        if src.border:
            dst.border = src.border.copy()

        if src.alignment:
            dst.alignment = src.alignment.copy()

        if src.protection:
            dst.protection = src.protection.copy()

        # 바로 위 행 수식을 한 줄 아래로 자동 보정 복사
        if isinstance(src.value, str) and src.value.startswith("="):
            # 아주 단순한 행번호 치환
            formula = src.value
            formula = re.sub(
                rf"(?<![A-Z]){src_row}(?!\d)",
                str(dst_row),
                formula
            )
            dst.value = formula


def find_target_row(ws, date_str: str, time_str: str):
    # A열=날짜, B열=시간 가정
    # 같은 날짜/시간 있으면 그 행 반환, 없으면 새 행
    last_data_row = max(ws.max_row, 2)

    for row in range(2, last_data_row + 1):
        a = ws.cell(row, 1).value
        b = ws.cell(row, 2).value
        if str(a).strip() == date_str and str(b).strip() == time_str:
            return row, False

    return last_data_row + 1, True


def floor_values_from_record(record: dict):
    # 최대한 유연하게 읽음
    # 예: record["floors"]["2층"] 또는 record["2층"]
    result = {}

    floors = record.get("floors")
    if isinstance(floors, dict):
        for floor_name in FLOOR_SHEET_CANDIDATES.keys():
            if floor_name in floors and isinstance(floors[floor_name], dict):
                result[floor_name] = floors[floor_name]

    for floor_name in FLOOR_SHEET_CANDIDATES.keys():
        if floor_name not in result and isinstance(record.get(floor_name), dict):
            result[floor_name] = record[floor_name]

    return result


def write_floor_row(ws, row: int, date_str: str, time_str: str, floor_data: dict):
    # 기본 매핑
    # A 날짜 / B 시간 / C 온도 / D 습도
    ws.cell(row, 1).value = date_str
    ws.cell(row, 2).value = time_str
    ws.cell(row, 3).value = floor_data.get("temperature")
    ws.cell(row, 4).value = floor_data.get("humidity")

    # 선택적으로 메모/비고 넣을 수 있으면 E
    if "note" in floor_data:
        ws.cell(row, 5).value = floor_data.get("note")


@app.post("/process-xlsm")
async def process_xlsm(file: UploadFile = File(...)):
    try:
        raw = await file.read()
        if not raw:
            return JSONResponse({"error": "업로드된 파일이 비어있습니다."}, status_code=400)

        latest = get_latest_record()
        dt = parse_datetime_from_record(latest)

        date_str = latest.get("date") or dt.strftime("%Y-%m-%d")
        time_str = latest.get("time") or dt.strftime("%H시")

        wb = openpyxl.load_workbook(
            filename=io.BytesIO(raw),
            keep_vba=True,
            data_only=False
        )

        floors_map = floor_values_from_record(latest)
        updated = []

        if not floors_map:
            # 테스트용 fallback: 시트 하나 추가해서 연결 확인 가능하게 함
            if "DATA" in wb.sheetnames:
                ws = wb["DATA"]
            else:
                ws = wb.create_sheet("DATA")

            ws["A1"] = "date"
            ws["B1"] = "time"
            ws["A2"] = date_str
            ws["B2"] = time_str
            ws["C1"] = "message"
            ws["C2"] = "온습도 데이터 구조를 서버가 아직 매칭하지 못했습니다."
            updated.append("DATA")
        else:
            for floor_name, floor_data in floors_map.items():
                sheet_name = find_sheet_name(wb, floor_name)
                if not sheet_name:
                    continue

                ws = wb[sheet_name]
                target_row, is_new = find_target_row(ws, str(date_str), str(time_str))

                if is_new and target_row > 2:
                    copy_row_style_and_formulas(ws, target_row - 1, target_row)

                write_floor_row(
                    ws=ws,
                    row=target_row,
                    date_str=str(date_str),
                    time_str=str(time_str),
                    floor_data=floor_data
                )
                updated.append(sheet_name)

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)

        output_name = make_output_filename(latest)

        headers = {
    "Content-Disposition": f"attachment; filename*=UTF-8''{requests.utils.quote(output_name)}"
}

return StreamingResponse(
    out,
    media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
    headers=headers
)

    except requests.RequestException as e:
        return JSONResponse(
            {"error": f"온습도 API 호출 실패: {str(e)}"},
            status_code=500
        )
    except Exception as e:
        return JSONResponse(
            {"error": f"엑셀 처리 실패: {str(e)}"},
            status_code=500
        )