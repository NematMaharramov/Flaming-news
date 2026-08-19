import io
import os
import json
import uuid
from datetime import date

from flask import Flask, request, render_template, send_file, send_from_directory, jsonify

import parsers
import rules
import excel_writer
import data_store
import report_config
from admin_api import admin_bp

app = Flask(__name__)
app.secret_key = 'flaming-news-secret'  # change in production / set via env var
app.register_blueprint(admin_bp)

PHOTOS_DIR = os.path.join(data_store.DATA_DIR, 'photos')
os.makedirs(PHOTOS_DIR, exist_ok=True)


def _today_iso():
    q = request.args.get('date') or (request.json.get('date') if request.is_json else None)
    return q or date.today().isoformat()


@app.route('/admin', methods=['GET'])
def admin_page():
    return render_template('admin.html')


@app.route('/admin/editor', methods=['GET'])
def admin_editor_page():
    return render_template('admin_editor.html')


@app.route('/', methods=['GET'])
def index():
    iso_date = request.args.get('date') or date.today().isoformat()
    data, carried, source_date = data_store.load_today_or_carry_forward(iso_date)
    config = report_config.load_published()
    return render_template(
        'index.html',
        iso_date=iso_date,
        data_json=json.dumps(data, ensure_ascii=False),
        config_json=json.dumps(config, ensure_ascii=False),
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


@app.route('/api/upload/forecast', methods=['POST'])
def api_upload_forecast():
    f = request.files.get('forecast_pdf')
    if not f or not f.filename:
        return jsonify({'status': 'error', 'message': 'No file uploaded'}), 400
    try:
        forecast = parsers.parse_forecast(f.stream)
        if not forecast['dates']:
            return jsonify({'status': 'error', 'message': 'No forecast rows recognised in this PDF'}), 400
        return jsonify({'status': 'ok', 'target': 'forecast', 'forecast': forecast})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 400


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
                'photo': '',
                'remarks': r.get('remark', ''),
                'remarks_needs_review': r.get('remarks_needs_review', False),
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
        records = parsers.parse_fb_report(f.stream, f.filename or '')
        return jsonify({'status': 'ok', 'target': 'fb_performance', 'records': records})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 400


@app.route('/api/upload/photo', methods=['POST'])
def api_upload_photo():
    """
    Accepts an already-cropped square image (client crops before upload —
    see the crop modal in index.html) and stores it under
    data/photos/<iso_date>/<uuid>.jpg. Returns a URL the front-end can show
    immediately and that gets saved into the guest record's 'photo' field.
    """
    f = request.files.get('photo')
    iso_date = request.form.get('iso_date') or date.today().isoformat()
    if not f or not f.filename:
        return jsonify({'status': 'error', 'message': 'No photo uploaded'}), 400
    day_dir = os.path.join(PHOTOS_DIR, iso_date)
    os.makedirs(day_dir, exist_ok=True)
    filename = f'{uuid.uuid4().hex}.jpg'
    f.save(os.path.join(day_dir, filename))
    return jsonify({'status': 'ok', 'url': f'/photos/{iso_date}/{filename}'})


@app.route('/photos/<path:subpath>')
def serve_photo(subpath):
    return send_from_directory(PHOTOS_DIR, subpath)


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


@app.route('/api/export-pdf', methods=['GET'])
def api_export_pdf():
    """
    Renders the same saved data through templates/print.html — a static,
    JS-free template mirroring the Excel design exactly — so the PDF always
    looks identical regardless of who opens it or what's in their browser.
    """
    iso_date = request.args.get('date') or date.today().isoformat()
    data = data_store.load_day(iso_date)
    if data is None:
        return jsonify({'status': 'error', 'message': 'No saved data for this date yet — save first.'}), 400

    from datetime import datetime as _dt
    try:
        title_date = _dt.strptime(iso_date, '%Y-%m-%d').strftime('%d %B %Y')
    except Exception:
        title_date = iso_date

    # Resolve each VIP guest's photo to an absolute filesystem path the
    # print template can reference directly (file:// URLs), same mapping
    # excel_writer uses for the Excel embed.
    for list_key in ('vip_arrivals', 'vip_inhouse', 'vip_departures'):
        for g in data.get(list_key, []):
            g['photo_fs_path'] = excel_writer._photo_fs_path(g.get('photo'))

    # Occupancy % gradient — same rule/colours excel_writer uses, computed
    # once here so print.html (which can't do colour math itself) just
    # applies a precomputed hex string per day.
    config = report_config.load_published()
    occ_rule = config.get('occupancy_color_rule', {})
    occupancy_colors = [
        report_config.occupancy_color(report_config.parse_pct_string(v), occ_rule)
        for v in data.get('forecast', {}).get('occupancy_pct', [])
    ]

    html_str = render_template('print.html', data=data, title_date=title_date, occupancy_colors=occupancy_colors)

    import weasyprint
    pdf_bytes = weasyprint.HTML(string=html_str, base_url=request.url_root).write_pdf()

    filename = f'Flaming_News_{iso_date}.pdf'
    return send_file(
        io.BytesIO(pdf_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype='application/pdf',
    )


@app.route('/health')
def health():
    return {'status': 'ok'}


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
