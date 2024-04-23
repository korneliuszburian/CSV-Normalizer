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

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads/'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls'}
app.secret_key = 'super secret key'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def load_product_mapping():
    product_mapping_df = pd.read_csv('products_data.csv')
    product_mapping_df['Dodatkowe oznaczenia'] = product_mapping_df['Dodatkowe oznaczenia'].fillna('').astype(str).str.strip()
    product_mapping_df['Nazwa'] = product_mapping_df['Nazwa'].str.strip()
    product_mapping = {k: v for k, v in product_mapping_df.set_index('Nazwa')['Dodatkowe oznaczenia'].items() if v}
    return product_mapping

@dataclass
class ExcelProcessorConfig:
    client_keyword: str = 'Valeo Electric'
    date_keywords: Dict[str, str] = field(default_factory=lambda: {
        'pick_up': 'Pick-up',
        'receiving': 'Receiving'
    })
    columns: list = field(default_factory=lambda: [
        'Planner', 'Reference', 'Description', 'Location', 'Unit', 'Details'
    ])

    column_mapping: Dict[str, str] = field(default_factory=lambda: {
        'Reference': ['Reference', 'REFERENCJA'],
        'Description': ['Description', 'OPIS'],
    })
    
    exclusion_keyword: str = 'Transportation mode'
    final_columns: list = field(default_factory=lambda: [
        'Klient', 'Oczekiwany termin realizacji', 'Termin potwierdzony', 'Reference', 'Produkt', 'Sztuk',
        'Uwagi dla wszystkich', 'Uwagi niewidoczne dla produkcji', 'Atrybut 1 (opcjonalnie)', 
        'Atrybut 2 (opcjonalnie)', 'Atrybut 3 (opcjonalnie)'
    ])

class ExcelProcessor:
    def __init__(self, config: ExcelProcessorConfig):
        self.config = config

    def extract_client_name(self, data: pd.DataFrame) -> Optional[str]:
        try:
            row = data[data.iloc[:, 3].str.contains(self.config.client_keyword, na=False)].iloc[0, 3]
            return row.split('\n')[0].split(',')[0]
        except IndexError as e:
            logging.error(f"Failed to extract client name: {e}")
            return None
        
    def find_relevant_sheet(self, excel_data: pd.ExcelFile) -> Optional[pd.DataFrame]:
        # Priority sheets
        priority_sheets = ['Antalis', 'Pick-up order ATK']
        for sheet_name in priority_sheets:
            if sheet_name in excel_data.sheet_names:
                return pd.read_excel(excel_data, sheet_name=sheet_name)

        # Fallback to search for the client keyword in all sheets if priority sheets are not found
        for sheet_name in excel_data.sheet_names:
            df = pd.read_excel(excel_data, sheet_name=sheet_name)
            if self.config.client_keyword in df.to_string():
                return df
        return None

    def extract_date(self, data: pd.DataFrame, keyword: str) -> str:
        try:
            logging.debug(f"Searching in DataFrame column: {data.iloc[:, 0].tolist()}")

            pattern = re.compile(r'\b' + re.escape(keyword) + r'\b', re.IGNORECASE)
            
            for index, row in data.iterrows():
                for col_index, cell_value in enumerate(row):
                    if pattern.search(str(cell_value)):
                        desired_data = None
                        for col in range(col_index + 1, len(row)):
                            if pd.notna(data.iloc[index, col]):
                                desired_data = str(data.iloc[index, col])
                                break
                        if desired_data is not None:
                            return desired_data
                        else:
                            return "Unknown date"
            
            logging.warning(f"No matches found for keyword '{keyword}'. Data might not contain the expected keywords.")
            return "Unknown date"
        
        except Exception as e:
            logging.error(f"An error occurred while extracting the date: {e}")
            return "Unknown date"

    def prepare_order_data(self, data: pd.DataFrame) -> pd.DataFrame:
        try:
            start_row = data[data.iloc[:, 0].str.contains('DB', na=False)].index[0]
            order_data = data.iloc[start_row:, :6]
            order_data.columns = self.config.columns
            return order_data.dropna(subset=['Reference']).query(f"Reference != '{self.config.exclusion_keyword}'")
        except IndexError as e:
            logging.error(f"Error processing order data: {e}")
            return pd.DataFrame()
        
    def process_file(self, file_path: str) -> pd.DataFrame:
        try:
            excel_data = pd.ExcelFile(file_path)
            antalis_data = self.find_relevant_sheet(excel_data)

            client_name = 'Valeo'
            pick_up_date = '19.08.2024'
            receiving_date = self.extract_date(antalis_data, self.config.date_keywords['receiving'])

            order_data = self.prepare_order_data(antalis_data)
            if not order_data.empty:
                order_data['Klient'] = client_name
                order_data['Oczekiwany termin realizacji'] = pick_up_date
                order_data['Termin potwierdzony'] = ""
                order_data['Produkt'] = order_data['Description'] + ' ' + order_data['Location'].fillna('')

                # product_mapping = load_product_mapping()
                # product_mapping = {k.strip(): v.strip() for k, v in product_mapping.items()}
                # order_data['Produkt'] = order_data['Produkt'].str.strip()
                # print("Data Before Mapping:", order_data['Produkt'].tolist())
                # order_data['Produkt'] = order_data['Produkt'].map(product_mapping).fillna(order_data['Produkt'])
                # print("Data After Mapping:", order_data['Produkt'].tolist())
                order_data['Sztuk'] = order_data.apply(lambda row: row['Details'] if pd.notna(row['Details']) and row['Details'] != 'NA' else '', axis=1)
                order_data = order_data[order_data['Sztuk'] != '']

                for column in self.config.final_columns:
                    if column not in order_data.columns:
                        order_data[column] = ''
                final_data = order_data[self.config.final_columns]
                final_data.columns = ['Klient', 'Oczekiwany termin realizacji', 'Termin potwierdzony', 'Zewn. nr zamówienia', 'Produkt', 'Sztuk', 'Uwagi dla wszystkich', 'Uwagi niewidoczne dla produkcji', 'Atrybut 1 (opcjonalnie)', 'Atrybut 2 (opcjonalnie)', 'Atrybut 3 (opcjonalnie)']

                return final_data
            else:
                return pd.DataFrame()
        except Exception as e:
            logging.error(f"An error occurred while processing the Excel file: {e}")
        return pd.DataFrame()

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('Brak pliku!', 'error')
            return redirect(url_for('upload_file'))
        file = request.files['file']
        if file.filename == '':
            flash('Nie wybrałeś pliku!', 'error')
            return redirect(url_for('upload_file'))
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            processed_data = processor.process_file(file_path)
            if not processed_data.empty:
                session['data'] = processed_data.to_csv(index=False, sep=';')
                return render_template('upload.html', data=processed_data.to_html(classes='data', index=False), has_data=True)
            else:
                flash('Brak informacji, bądź nie można ich znaleźć.', 'error')
                return redirect(url_for('upload_file'))
        else:
            flash('Nieprawidłowy typ pliku', 'error')
            return redirect(url_for('upload_file'))
    else:
        return render_template('upload.html', data=None, has_data=False)

@app.route('/normalizer', methods=['GET'])
def display_vr():
    if request.method == 'GET':
        return render_template('normalizer.html');  

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
            expected_date = datetime.strptime(expected_date.strip(), '%d.%m.%Y').strftime('%Y-%m-%d')
            csv_data['Oczekiwany termin realizacji'] = expected_date

        if confirmed_date is not None and confirmed_date.strip():
            confirmed_date = datetime.strptime(confirmed_date.strip(), '%d.%m.%Y').strftime('%Y-%m-%d')
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
    config = ExcelProcessorConfig()
    processor = ExcelProcessor(config)
    http_server = WSGIServer(('', 3000), app)
    http_server.serve_forever()