from flask import Flask, request, send_file, render_template, redirect, url_for, session, make_response, flash
import re
from werkzeug.utils import secure_filename
import os
import pandas as pd
from io import StringIO
import logging
from dataclasses import dataclass, field
from typing import Optional, Dict
from datetime import datetime
from gevent.pywsgi import WSGIServer
from processor import ExcelDataLoader, DataFrameProcessor

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}
app.secret_key = 'super secret key'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part!', 'error')
            return redirect(url_for('upload_file'))
        file = request.files['file']
        if file.filename == '':
            flash('No file selected!', 'error')
            return redirect(url_for('upload_file'))
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            possible_sheets = ['Antalis', 'Pick-up order ATK', 'ATK', 'Sheet1']
            
            try:
                loader = ExcelDataLoader(file_path, possible_sheets)
                df = loader.load_data()
                processor = DataFrameProcessor(df)
                processor.normalize_headers()
                processor.rename_and_select_columns({
                    ("Reference", "Referencja"): "Zewn. nr zamówienia",
                    ("Description",): "Produkt",
                    ("Ordered in Std Pack",): "Ordered in Std Pack",
                    ("Unit",): "Sztuk",
                    ("Loop Size",): "Loop Size"
                })
                processor.add_additional_columns()
                processor.filter_data()
                processor.remove_empty_rows()
                processor.remove_weird_rows()
                processor.finalize_columns_order()

                processed_data = processor.df.to_csv(index=False, sep=';')
                session['data'] = processed_data
                return render_template('upload.html', data=processor.df.to_html(classes='data', index=False, na_rep=''), has_data=True)
            except Exception as e:
                logging.error(f"An error occurred while processing the Excel file: {e}")
                flash(f'Failed to process file: {str(e)}', 'error')
                return redirect(url_for('upload_file'))
        else:
            flash('Invalid file type', 'error')
            return redirect(url_for('upload_file'))
    else:
        return render_template('upload.html', data=None, has_data=False)

@app.route('/pobierz', methods=['GET'])
def download_file():
    if 'data' in session:
        csv_data = pd.read_csv(StringIO(session['data']), sep=';')
        client_name = csv_data['Klient'].iloc[0] if not csv_data['Klient'].empty else 'default'

        today = datetime.now().strftime('%Y-%m-%d')

        filename = f"{client_name}_{today}.csv"

        response = make_response(session['data'])
        response.headers["Content-Disposition"] = f"attachment; filename={filename}"
        response.headers["Content-Type"] = "text/csv"
        return response
    else:
        return redirect(url_for('upload_file'))

@app.route('/edycja', methods=['POST'])
def adjust_data():
    if 'data' in session:
        client_name = request.form.get('client_name', None)
        expected_date = request.form.get('expected_date', None)
        confirmed_date = request.form.get('confirmed_date', None)

        csv_data = pd.read_csv(StringIO(session['data']), sep=';')

        if client_name is not None and client_name.strip():
            csv_data['Klient'] = client_name.strip()

        if expected_date is not None and expected_date.strip():
            expected_date = datetime.strptime(expected_date.strip(), '%d/%m/%Y').strftime('%Y-%m-%d')
            csv_data['Oczekiwany termin realizacji'] = expected_date

        if confirmed_date is not None and confirmed_date.strip():
            confirmed_date = datetime.strptime(confirmed_date.strip(), '%d/%m/%Y').strftime('%Y-%m-%d')
            csv_data['Termin potwierdzony'] = confirmed_date

        updated_csv = csv_data.to_csv(index=False, sep=';', na_rep='')
        session['data'] = updated_csv

        return render_template('upload.html', data=csv_data.to_html(classes='data', index=False, na_rep=''), has_data=True)
    else:
        return redirect(url_for('upload_file'))

@app.route('/adjust_client_name', methods=['POST'])
def adjust_client_name():
    if 'data' in session:
        client_name = request.form.get('client_name', None)
        
        csv_data = pd.read_csv(StringIO(session['data']))

        if client_name is not None and client_name.strip():
            csv_data['Klient'] = client_name.strip()

        updated_csv = csv_data.to_csv(index=False, na_rep='')
        session['data'] = updated_csv

        return render_template('upload.html', data=csv_data.to_html(classes='data', index=False, na_rep=''), has_data=True)
    else:
        return redirect(url_for('upload_file'))

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

if __name__ == '__main__':
    # http_server = WSGIServer(('', 3000), app)
    # http_server.serve_forever()
    app.run(host="0.0.0.0", port=3000)