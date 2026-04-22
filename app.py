from flask import Flask, request, jsonify, send_file, make_response
from flask_cors import CORS
import openpyxl
from datetime import datetime, time as time_cls
import io
import os
import json
import traceback

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=False)

# 모든 응답에 CORS 헤더 강제 추가
@app.after_request
def add_cors_headers(response):
    response.headers['Access-Control-Allow-Origin']  = '*'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, PUT, OPTIONS'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'
    return response


FLOOR_SCHEMA = {
    '2층': { 'units': [0,1,2,3,4,5,6,7], 'start_col': 9 },
    '3층': { 'units': [0,1,2,3],         'start_col': 9 },
    '5층': { 'units': [0,1,2,3],         'start_col': 9 },
    '6층': { 'units': [0,1,2,4,5,6,7],   'start_col': 9 },
    '8층': { 'units': [0,1,2,3],         'start_col': 9 },
}


@app.route('/', methods=['GET'])
def index():
    return jsonify({'service': 'temperature-server', 'status': 'ok'})


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'ok': True})


@app.route('/update-xlsm', methods=['POST', 'OPTIONS'])
def update_xlsm():
    if request.method == 'OPTIONS':
        return make_response('', 204)

    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        if 'data' not in request.form:
            return jsonify({'error': 'No data provided'}), 400

        xlsm_file = request.files['file']
        record = json.loads(request.form['data'])

        buf = io.BytesIO(xlsm_file.read())
        wb = openpyxl.load_workbook(buf, keep_vba=True, keep_links=False, data_only=False)

        date_str  = record.get('date', '')
        day_str   = record.get('day', '')
        time_val  = record.get('timeVal', '00')
        shift     = record.get('shift', '주간')
        inspector = record.get('inspector', '')
        floors_data = record.get('floors', {})

        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            mon = str(date_obj.month)
            day = str(date_obj.day).zfill(2)
        except:
            return jsonify({'error': f'Invalid date: {date_str}'}), 400

        key = mon + day + day_str + shift
        hh = int(time_val)
        time_obj = time_cls(hh, 0)

        for floor_name, schema in FLOOR_SCHEMA.items():
            if floor_name not in wb.sheetnames:
                continue
            ws = wb[floor_name]
            floor_data = floors_data.get(floor_name, {})

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
            
            # 수식이 있는 템플릿 행 찾기 (요약 컬럼에 수식 있는 마지막 행)
            summary_col = schema['start_col'] + len(schema['units']) * 2  # Y, Q 등
            template_row = last_row
            for r in range(last_row, 3, -1):
                val = ws.cell(row=r, column=summary_col).value
                if isinstance(val, str) and val.startswith('='):
                    template_row = r
                    break

            ws.cell(row=new_row, column=1).value = key
            ws.cell(row=new_row, column=2).value = new_no
            ws.cell(row=new_row, column=3).value = mon
            ws.cell(row=new_row, column=4).value = day
            ws.cell(row=new_row, column=5).value = day_str
            ws.cell(row=new_row, column=6).value = shift
            ws.cell(row=new_row, column=7).value = time_obj
            ws.cell(row=new_row, column=7).number_format = 'h:mm'
            ws.cell(row=new_row, column=8).value = inspector

            units = schema['units']
            sc = schema['start_col']
            for i, u in enumerate(units):
                d = floor_data.get(str(u), {})
                t_val = d.get('temp', '')
                h_val = d.get('hum', '')
                t = float(t_val) if t_val != '' else None
                h = float(h_val) if h_val != '' else None
                ws.cell(row=new_row, column=sc + i*2).value     = t
                ws.cell(row=new_row, column=sc + i*2 + 1).value = h

            # 스타일 + 수식 복사 (템플릿 행 기준)
            from copy import copy
            import re
            prev_row = template_row
            data_end_col = schema['start_col'] + len(schema['units']) * 2 - 1  # 온도/습도 마지막 컬럼
            
            for col in range(1, ws.max_column + 1):
                src = ws.cell(row=prev_row, column=col)
                dst = ws.cell(row=new_row, column=col)
                
                # 스타일 복사
                if src.has_style:
                    dst._style = copy(src._style)
                    dst.number_format = src.number_format
                
                # 수식 복사 (데이터 컬럼 이후의 수식 컬럼만)
                # 온도/습도/기본정보 컬럼은 이미 위에서 값을 넣었으므로 건드리지 않음
                if col > data_end_col and isinstance(src.value, str) and src.value.startswith('='):
                    # 수식 내 행 번호 치환: A123 → A124 (행 번호만 변경)
                    formula = src.value
                    # 현재 행 번호(prev_row)를 새 행 번호(new_row)로 변경
                    # 단, 절대 참조($1)는 유지
                    # 행 번호만 치환 ($로 시작하는 절대 참조는 유지)
                    # 예: A1749 → A1750, $A1749 → $A1750, A$1749 → A$1749 (절대참조 유지)
                    old_r = str(prev_row)
                    new_r = str(new_row)
                    # 패턴: (컬럼문자)(행번호) - 행에 $가 없을 때만
                    new_formula = re.sub(
                        r'([A-Z]+)' + old_r + r'\b',
                        lambda m: m.group(1) + new_r,
                        formula
                    )
                    dst.value = new_formula

        date_part = date_str.replace('-', '')[2:]
        filename = f'일일점검사항_v8__{date_part}_{time_val}시.xlsm'

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
        print('ERROR:', str(e))
        print(traceback.format_exc())
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
