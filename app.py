from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime, time
import io
import os
import json

app = Flask(__name__)
CORS(app)

# 층별 컬럼 구조 (엑셀 기준)
FLOOR_SCHEMA = {
    '2층': { 'units': [0,1,2,3,4,5,6,7], 'start_col': 9 },
    '3층': { 'units': [0,1,2,3],         'start_col': 9 },
    '5층': { 'units': [0,1,2,3],         'start_col': 9 },
    '6층': { 'units': [0,1,2,4,5,6,7],   'start_col': 9 },
    '8층': { 'units': [0,1,2,3],         'start_col': 9 },
}

LIM = { 't_min': 21, 't_max': 27, 'h_min': 30, 'h_max': 70 }

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'ok': True})

@app.route('/update-xlsm', methods=['POST'])
def update_xlsm():
    try:
        # xlsm 파일 + 점검 데이터 수신
        if 'file' not in request.files:
            return jsonify({'error': 'No file'}), 400
        if 'data' not in request.form:
            return jsonify({'error': 'No data'}), 400

        xlsm_file = request.files['file']
        record = json.loads(request.form['data'])

        # openpyxl로 xlsm 로드 (매크로 유지)
        buf = io.BytesIO(xlsm_file.read())
        wb = openpyxl.load_workbook(buf, keep_vba=True)

        date_str = record.get('date', '')       # "2026-04-21"
        day_str  = record.get('day', '')         # "월"
        time_val = record.get('timeVal', '00')  # "07"
        shift    = record.get('shift', '주간')
        inspector = record.get('inspector', '')
        floors_data = record.get('floors', {})

        # 날짜 파싱
        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            mon = str(date_obj.month)
            day = str(date_obj.day).zfill(2)
        except:
            mon = '0'; day = '00'

        key = mon + day + day_str + shift
        hh = int(time_val)
        time_obj = time(hh, 0)

        for floor_name, schema in FLOOR_SCHEMA.items():
            if floor_name not in wb.sheetnames:
                continue
            ws = wb[floor_name]
            floor_data = floors_data.get(floor_name, {})

            # 마지막 NO 찾기
            last_no = 0
            last_row = ws.max_row
            for r in range(ws.max_row, 3, -1):
                val = ws.cell(row=r, column=2).value
                if isinstance(val, (int, float)):
                    last_no = int(val)
                    last_row = r
                    break
            new_row = last_row + 1
            new_no  = last_no + 1

            # 기본 정보 입력
            ws.cell(row=new_row, column=1).value = key
            ws.cell(row=new_row, column=2).value = new_no
            ws.cell(row=new_row, column=3).value = mon
            ws.cell(row=new_row, column=4).value = day
            ws.cell(row=new_row, column=5).value = day_str
            ws.cell(row=new_row, column=6).value = shift
            ws.cell(row=new_row, column=7).value = time_obj
            ws.cell(row=new_row, column=7).number_format = 'h:mm'
            ws.cell(row=new_row, column=8).value = inspector

            # 호기별 온도/습도
            units = schema['units']
            sc    = schema['start_col']
            temps = []
            hums  = []
            for i, u in enumerate(units):
                d = floor_data.get(str(u), {})
                t_val = d.get('temp', '')
                h_val = d.get('hum', '')
                t = float(t_val) if t_val != '' else None
                h = float(h_val) if h_val != '' else None
                ws.cell(row=new_row, column=sc + i*2).value = t
                ws.cell(row=new_row, column=sc + i*2 + 1).value = h
                if t is not None: temps.append(t)
                if h is not None: hums.append(h)

            # 이전 행에서 스타일 복사
            prev_row = last_row
            for col in range(1, ws.max_column + 1):
                src = ws.cell(row=prev_row, column=col)
                dst = ws.cell(row=new_row, column=col)
                if src.has_style:
                    from copy import copy
                    dst.font       = copy(src.font)
                    dst.border     = copy(src.border)
                    dst.fill       = copy(src.fill)
                    dst.alignment  = copy(src.alignment)
                    dst.number_format = src.number_format

        # 파일명 생성
        date_part = date_str.replace('-', '')[2:]  # 260421
        filename = f'일일점검사항_v8__{date_part}_{time_val}시.xlsm'

        # 출력
        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)

        return send_file(
            out_buf,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.ms-excel.sheet.macroEnabled.12'
        )

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
