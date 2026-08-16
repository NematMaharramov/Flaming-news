import io
import json
from datetime import date

from flask import Flask, request, render_template, send_file, jsonify

import parsers
import rules
import excel_writer
import data_store

app = Flask(__name__)
app.secret_key = 'flaming-news-secret'  # change in production / set via env var


def _today_iso():
    q = request.args.get('date') or (request.json.get('date') if request.is_json else None)
    return q or date.today().isoformat()


@app.route('/', methods=['GET'])
def index():
    iso_date = request.args.get('date') or date.today().isoformat()
    data, carried, source_date = data_store.load_today_or_carry_forward(iso_date)
    return render_template(
        'index.html',
        iso_date=iso_date,
        data_json=json.dumps(data, ensure_ascii=False),
        carried=carried,
        source_date=source_date,
    )


@app.route('/api/save', methods=['POST'])
def api_save():
    payload = request.get_json(force=True)
    iso_date = payload.get('iso_date') or date.today().isoformat()
    saved = data_store.save_day(iso_date, payload)
    return jsonify({'status': 'ok', 'saved_at': saved.get('_saved_at')})


@app.route('/api/upload/inhouse', methods=['POST'])
def api_upload_inhouse():
    return _upload_vip_list('inhouse_pdf', parsers.parse_in_house, 'vip_inhouse',
                             extra=lambda r: r.setdefault('departure_day', r.get('dep_date', '')))


@app.route('/api/upload/departures', methods=['POST'])
def api_upload_departures():
    return _upload_vip_list('departures_pdf', parsers.parse_departures, 'vip_departures',
                             extra=lambda r: r.setdefault('departure_day', r.get('dep_date', '')))


@app.route('/api/upload/arrivals', methods=['POST'])
def api_upload_arrivals():
    return _upload_vip_list('arrivals_pdf', parsers.parse_arrivals, 'vip_arrivals')


def _upload_vip_list(file_key, parse_fn, target_key, extra=None):
    f = request.files.get(file_key)
    if not f or not f.filename:
        return jsonify({'status': 'error', 'message': 'No file uploaded'}), 400
    try:
        raw_records = parse_fn(f.stream)
        for rec in raw_records:
            rec['company'] = parsers.clean_company(rec.get('company', ''))
        enriched = rules.apply_business_rules(raw_records)
        result = []
        for r in enriched:
            if extra:
                extra(r)
            result.append({
                'guest': r.get('name', ''),
                'code': r.get('vip_code', ''),
                'room': r.get('room', ''),
                'eta': r.get('eta', ''),
                'departure_day': r.get('departure_day', r.get('dep_date', '')),
                'departure_time': r.get('dep_time', ''),
                'remarks': r.get('remark', ''),
                'company': r.get('company', ''),
            })
        return jsonify({'status': 'ok', 'target': target_key, 'records': result})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 400


@app.route('/api/upload/fb', methods=['POST'])
def api_upload_fb():
    f = request.files.get('fb_report')
    if not f or not f.filename:
        return jsonify({'status': 'error', 'message': 'No file uploaded'}), 400
    try:
        records = parsers.parse_fb_report(f.stream)
        return jsonify({'status': 'ok', 'target': 'fb_performance', 'records': records})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 400


@app.route('/api/export', methods=['GET'])
def api_export():
    iso_date = request.args.get('date') or date.today().isoformat()
    data = data_store.load_day(iso_date)
    if data is None:
        return jsonify({'status': 'error', 'message': 'No saved data for this date yet — save first.'}), 400
    out = excel_writer.build_workbook(data)
    filename = f'Flaming_News_{iso_date}.xlsx'
    return send_file(
        out,
        as_attachment=True,
        download_name=filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@app.route('/health')
def health():
    return {'status': 'ok'}


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
