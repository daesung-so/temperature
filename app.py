from flask import Flask, request, jsonify, send_file, make_response
from flask_cors import CORS
import openpyxl
from datetime import datetime, time as time_cls
from copy import copy
import io
import os
import json
import re
import zipfile
import traceback

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=False)

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

DATA_SHEETS = {'2층', '3층', '5층', '6층', '8층'}


def restore_drawings(processed_buf, original_bytes):
    """openpyxl이 제거한 보조 파일들을 원본에서 복원
    - sharedStrings.xml, calcChain.xml, drawings/, ctrlProps/, printerSettings/, _rels/
    - 데이터 미입력 시트(종합그래프, 최고온도최저습도): 원본 sheet xml 그대로 복원
    - 데이터 입력 시트: 처리본 sheet xml에 drawing/legacyDrawing/oleObjects 태그 주입
    """
    processed_buf.seek(0)

    # 원본 sheet 매핑 분석
    orig_sheet_map = {}  # sheet_name → sheet_xml_path
    restore_full = {}     # 원본 그대로 복원할 파일들
    drawing_tags = {}     # 처리본 sheet_path → 주입할 drawing 태그 리스트

    with zipfile.ZipFile(io.BytesIO(original_bytes), 'r') as orig_zip:
        wb_xml  = orig_zip.read('xl/workbook.xml').decode('utf-8')
        wb_rels = orig_zip.read('xl/_rels/workbook.xml.rels').decode('utf-8')

        # 시트 이름 → rId
        sheet_rid = {}
        for m in re.finditer(r'<sheet\s+([^>]+?)/?>', wb_xml):
            attrs = m.group(1)
            n = re.search(r'name="([^"]+)"', attrs)
            r_ = re.search(r'r:id="(rId\d+)"', attrs)
            if n and r_:
                sheet_rid[n.group(1)] = r_.group(1)

        # rId → sheet 파일 경로
        rid_target = {}
        for m in re.finditer(r'<Relationship\s+([^>]+?)/>', wb_rels):
            attrs = m.group(1)
            id_m  = re.search(r'Id="(rId\d+)"', attrs)
            tgt_m = re.search(r'Target="(worksheets/sheet\d+\.xml)"', attrs)
            if id_m and tgt_m:
                rid_target[id_m.group(1)] = tgt_m.group(1)

        for sheet_name, rid in sheet_rid.items():
            if rid in rid_target:
                orig_sheet_map[sheet_name] = 'xl/' + rid_target[rid]

        # 복원 대상 보조 파일들 모두 읽기
        for name in orig_zip.namelist():
            if (name.startswith('xl/drawings/')
                or name.startswith('xl/ctrlProps/')
                or name.startswith('xl/printerSettings/')
                or name.startswith('xl/worksheets/_rels/')
                or name == 'xl/sharedStrings.xml'
                or name == 'xl/calcChain.xml'):
                restore_full[name] = orig_zip.read(name)

        # 데이터 미입력 시트는 원본 sheet xml 그대로 복원
        for sheet_name, sheet_path in orig_sheet_map.items():
            if sheet_name not in DATA_SHEETS:
                restore_full[sheet_path] = orig_zip.read(sheet_path)

        # 데이터 입력 시트는 drawing 관련 태그 추출
        for sheet_name in DATA_SHEETS:
            if sheet_name not in orig_sheet_map:
                continue
            sheet_path = orig_sheet_map[sheet_name]
            sheet_xml = orig_zip.read(sheet_path).decode('utf-8')
            tags = []
            for tag in ['drawing', 'legacyDrawing', 'oleObjects', 'controls']:
                # self-closing 또는 open-close 둘 다
                for m in re.finditer(
                    rf'<{tag}\s[^/>]*/>|<{tag}\s[^>]*>.*?</{tag}>',
                    sheet_xml, re.DOTALL
                ):
                    tags.append(m.group(0))
            if tags:
                drawing_tags[sheet_path] = tags

        original_ct = orig_zip.read('[Content_Types].xml').decode('utf-8')

    # 새 zip 작성
    new_buf = io.BytesIO()
    with zipfile.ZipFile(processed_buf, 'r') as proc_zip:
        with zipfile.ZipFile(new_buf, 'w', zipfile.ZIP_DEFLATED) as new_zip:
            written = set()

            for name in proc_zip.namelist():
                if name in restore_full:
                    continue  # 원본에서 복원할 것
                if name == '[Content_Types].xml':
                    new_zip.writestr(name, original_ct)
                elif name in drawing_tags:
                    # 데이터 시트: drawing 태그 주입
                    sheet_xml = proc_zip.read(name).decode('utf-8')
                    insert_str = ''.join(drawing_tags[name])
                    sheet_xml = sheet_xml.replace(
                        '</worksheet>', insert_str + '</worksheet>'
                    )
                    new_zip.writestr(name, sheet_xml)
                else:
                    new_zip.writestr(name, proc_zip.read(name))
                written.add(name)

            # 원본에서 복원
            for name, data in restore_full.items():
                if name not in written:
                    new_zip.writestr(name, data)

    new_buf.seek(0)
    return new_buf


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

        original_bytes = xlsm_file.read()
        buf = io.BytesIO(original_bytes)
        wb = openpyxl.load_workbook(
            buf, keep_vba=True, keep_links=False, data_only=False
        )

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

            # 마지막 NO/행 찾기
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

            # 수식 있는 템플릿 행 찾기
            summary_col = schema['start_col'] + len(schema['units']) * 2
            template_row = last_row
            for r in range(last_row, 3, -1):
                val = ws.cell(row=r, column=summary_col).value
                if isinstance(val, str) and val.startswith('='):
                    template_row = r
                    break

            # 기본 정보
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
            sc = schema['start_col']
            for i, u in enumerate(units):
                d = floor_data.get(str(u), {})
                t_val = d.get('temp', '')
                h_val = d.get('hum', '')
                t = float(t_val) if t_val != '' else None
                h = float(h_val) if h_val != '' else None
                ws.cell(row=new_row, column=sc + i*2).value     = t
                ws.cell(row=new_row, column=sc + i*2 + 1).value = h

            # 스타일 + 수식 복사 (template_row 기준)
            data_end_col = schema['start_col'] + len(schema['units']) * 2 - 1
            old_r = str(template_row)
            new_r = str(new_row)

            for col in range(1, ws.max_column + 1):
                src = ws.cell(row=template_row, column=col)
                dst = ws.cell(row=new_row, column=col)

                # 스타일
                if src.has_style:
                    dst._style = copy(src._style)
                    dst.number_format = src.number_format

                # 수식 복사 (요약 컬럼 이후)
                if col > data_end_col and isinstance(src.value, str) and src.value.startswith('='):
                    new_formula = re.sub(
                        r'([A-Z]+)' + old_r + r'\b',
                        lambda m: m.group(1) + new_r,
                        src.value
                    )
                    dst.value = new_formula

        # 파일명 (4자리 연도)
        date_part = date_str.replace('-', '')
        filename = f'일일점검사항_v8__{date_part}_{time_val}시.xlsm'

        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)

        # ZIP 조작으로 보조 파일 복원
        out_buf = restore_drawings(out_buf, original_bytes)

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
