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
        wb = openpyxl.load_workbook(buf, keep_vba=True)

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

            # 스타일 복사
            from copy import copy
            prev_row = last_row
            for col in range(1, ws.max_column + 1):
                src = ws.cell(row=prev_row, column=col)
                dst = ws.cell(row=new_row, column=col)
                if src.has_style:
                    dst.font       = copy(src.font)
                    dst.border     = copy(src.border)
                    dst.fill       = copy(src.fill)
                    dst.alignment  = copy(src.alignment)
                    dst.number_format = src.number_format

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
