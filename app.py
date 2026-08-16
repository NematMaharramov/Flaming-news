import io
from datetime import date
from flask import Flask, request, render_template, send_file, flash, redirect, url_for

import parsers
import rules
import excel_writer

app = Flask(__name__)
app.secret_key = 'flaming-news-secret'  # change in production / set via env var


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html', today=date.today().strftime('%d.%m.%Y'))


@app.route('/generate', methods=['POST'])
def generate():
    try:
        report_date = request.form.get('report_date') or date.today().strftime('%d.%m.%Y')

        forecast_file = request.files.get('forecast_pdf')
        arrivals_file = request.files.get('arrivals_pdf')
        inhouse_file = request.files.get('inhouse_pdf')
        departures_file = request.files.get('departures_pdf')
        photos_file = request.files.get('photos_zip')

        forecast = parsers.parse_forecast(forecast_file.stream) if forecast_file and forecast_file.filename else []
        arrivals = parsers.parse_arrivals(arrivals_file.stream) if arrivals_file and arrivals_file.filename else []
        in_house = parsers.parse_in_house(inhouse_file.stream) if inhouse_file and inhouse_file.filename else []
        departures = parsers.parse_departures(departures_file.stream) if departures_file and departures_file.filename else []

        # Clean company prefixes
        for group in (arrivals, in_house, departures):
            for rec in group:
                rec['company'] = parsers.clean_company(rec.get('company', ''))

        # Apply VIP business rules (remarks + colors), per section independently
        rules.apply_business_rules(arrivals)
        rules.apply_business_rules(in_house)
        rules.apply_business_rules(departures)

        photos_bytes = photos_file.read() if photos_file and photos_file.filename else None

        out = excel_writer.build_workbook(
            report_date, forecast, arrivals, in_house, departures, photos_bytes
        )

        filename = f'Flaming_News_{report_date.replace(".", "-")}.xlsx'
        return send_file(
            out,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
    except Exception as e:
        flash(f'Xəta baş verdi: {e}')
        return redirect(url_for('index'))


@app.route('/health')
def health():
    return {'status': 'ok'}


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
