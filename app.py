"""
PDF to XLSX Extractor App
Extrae datos estructurados de PDFs de inventario de archivo de trámite y concentración
para generar plantillas PERGUB (Anexo 30 y Anexo 31).
"""

import io
import json
import os
import re
import tempfile
import threading
import uuid
from flask import Flask, request, jsonify, send_file, render_template, redirect, url_for
from flask_cors import CORS
import pdfplumber
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
CORS(app, expose_headers=['X-Records-Count', 'Content-Disposition'])

# Configuración
UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'pdf'}
COUNTER_FILE = os.path.join(os.path.dirname(__file__), 'counter.json')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB máximo

# Bloqueo para escritura segura en counter.json bajo concurrencia
_counter_lock = threading.Lock()

# Mapas de normalización compartidos
CLASIFICACION_MAP = {
    'PUBLICO': 'Pública',
    'PÚBLICO': 'Pública',
    'CONFIDENCIAL': 'Confidencial',
    'RESERVADO': 'Reservada',
}

CARACTER_MAP = {
    'ADMINISTRATIVO': 'Administrativo',
    'LEGAL': 'Legal',
    'FISCAL': 'Fiscal',
    'CONTABLE': 'Contable',
}

# Porcentajes de recorte vertical para excluir encabezado y pie de página
HEADER_PERCENT = 0.28
FOOTER_PERCENT = 0.14


def get_counter():
    """Obtiene el contador actual del archivo JSON de forma segura."""
    with _counter_lock:
        try:
            if os.path.exists(COUNTER_FILE):
                with open(COUNTER_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    data.setdefault('total_records', 0)
                    data.setdefault('total_files', 0)
                    return data
        except (OSError, json.JSONDecodeError):
            pass
    return {'total_records': 0, 'total_files': 0}


def update_counter(records_count):
    """Actualiza el contador de forma atómica y segura para múltiples hilos."""
    with _counter_lock:
        counter = {'total_records': 0, 'total_files': 0}
        try:
            if os.path.exists(COUNTER_FILE):
                with open(COUNTER_FILE, 'r', encoding='utf-8') as f:
                    counter = json.load(f)
        except (OSError, json.JSONDecodeError):
            counter = {'total_records': 0, 'total_files': 0}

        counter['total_records'] = counter.get('total_records', 0) + records_count
        counter['total_files'] = counter.get('total_files', 0) + 1

        tmp_path = f"{COUNTER_FILE}.{uuid.uuid4().hex}.tmp"
        with open(tmp_path, 'w', encoding='utf-8') as f:
            json.dump(counter, f, ensure_ascii=False)
        os.replace(tmp_path, COUNTER_FILE)
        return counter


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def clean_cell_text(cell):
    """Limpia saltos de línea múltiples y espacios sobrantes en una celda."""
    if cell is None:
        return ''
    text = str(cell).replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
    return re.sub(r'\s+', ' ', text).strip()


def normalize_records(records):
    """Limpia espacios, quita espacios en clave y normaliza clasificación y carácter."""
    for record in records:
        for key, value in record.items():
            if isinstance(value, str):
                record[key] = re.sub(r'\s+', ' ', value).strip()

        if isinstance(record.get('clave'), str):
            record['clave'] = record['clave'].replace(' ', '')

        clasif = record.get('clasificacion', '').strip().upper()
        if clasif in CLASIFICACION_MAP:
            record['clasificacion'] = CLASIFICACION_MAP[clasif]

        carac = record.get('caracter', '').strip().upper()
        if carac in CARACTER_MAP:
            record['caracter'] = CARACTER_MAP[carac]

    return records


# ============== MÓDULO DE TRÁMITE (ANEXO 30) ==============

def extract_table_data(pdf_path):
    """Extrae datos de tabla del PDF de Trámite usando recorte por coordenadas Y."""
    all_records = []
    current_record = None

    column_header_indicators = [
        'NO. PROGRESIVO', 'PROGRESIVO', 'NO. PROG',
        'CLAVE DEL EXPEDIENTE',
        'NOMBRE O DESCRIPCION', 'DESCRIPCIÓN DEL EXPEDIENTE',
        'MES O AÑO', 'MES O ANO',
        'LOCALIZACION', 'LOCALIZACIÓN',
        'CLASIFICACION', 'CLASIFICACIÓN',
        'CARACTER', 'CARÁCTER',
        'OBSERVACION', 'OBSERVACIÓN',
    ]

    def is_column_header_row(row_text):
        text_upper = row_text.upper()
        matches = sum(1 for kw in column_header_indicators if kw in text_upper)
        return matches >= 3

    def append_continuation(record, cleaned_row):
        """Concatena filas de continuación (incluyendo observaciones con múltiples saltos de línea)."""
        for i, cell in enumerate(cleaned_row):
            if not cell:
                continue
            if i == 1:
                record['clave'] += ' ' + cell
            elif i == 2:
                record['nombre'] += ' ' + cell
            elif i == 3:
                if not record['mes_ano']:
                    record['mes_ano'] = cell
            elif i == 4:
                record['localizacion'] += ' ' + cell
            elif i == 5:
                if not record['clasificacion']:
                    record['clasificacion'] = cell
            elif i == 6:
                if not record['caracter']:
                    record['caracter'] = cell
            elif i == 7:
                record['observacion'] += ' ' + cell
            elif i == 0:
                record['nombre'] += ' ' + cell

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            top_crop = page.height * HEADER_PERCENT
            bottom_crop = page.height * (1 - FOOTER_PERCENT)
            cropped_page = page.crop((0, top_crop, page.width, bottom_crop))

            tables = cropped_page.extract_tables({
                'vertical_strategy': 'lines',
                'horizontal_strategy': 'lines',
                'snap_tolerance': 3,
                'join_tolerance': 3,
            })

            if not tables:
                continue

            for table in tables:
                for row in table:
                    if not row or all(cell is None or str(cell).strip() == '' for cell in row):
                        continue

                    cleaned_row = [clean_cell_text(cell) for cell in row]
                    full_row_text = ' '.join(cleaned_row)

                    if is_column_header_row(full_row_text):
                        continue

                    first_cell = cleaned_row[0] if cleaned_row else ''

                    if first_cell:
                        try:
                            prog_num = int(first_cell)
                            if current_record:
                                all_records.append(current_record)

                            while len(cleaned_row) < 8:
                                cleaned_row.append('')

                            current_record = {
                                'no_progresivo': prog_num,
                                'clave': cleaned_row[1],
                                'nombre': cleaned_row[2],
                                'mes_ano': cleaned_row[3],
                                'localizacion': cleaned_row[4],
                                'clasificacion': cleaned_row[5],
                                'caracter': cleaned_row[6],
                                'observacion': cleaned_row[7],
                            }
                        except (ValueError, TypeError):
                            if current_record:
                                append_continuation(current_record, cleaned_row)
                    elif current_record:
                        # Continuación sin número progresivo (incluye observaciones multilínea)
                        append_continuation(current_record, cleaned_row)

    if current_record:
        all_records.append(current_record)

    return normalize_records(all_records)


def create_xlsx_bytes(records):
    """Genera el archivo XLSX de Trámite en memoria usando la plantilla anexo_30."""
    template_path = os.path.join(os.path.dirname(__file__), 'anexo_30_template.xlsx')

    if os.path.exists(template_path):
        wb = load_workbook(template_path)
        ws = wb.active
        start_row = 5
        for merged_range in list(ws.merged_cells.ranges):
            if merged_range.min_row >= start_row:
                ws.unmerge_cells(str(merged_range))
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = 'Inventario'
        start_row = 2
        headers = [
            'No.', 'Clasificación', 'Nombre del Expediente', 'Descripción',
            'Valor documental', 'Soporte Documental', 'Ubicación', 'Ejercicio', 'Observaciones'
        ]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)

    column_mapping = {
        'no_progresivo': 1,  # A: No.
        'clasificacion': 2,  # B: Clasificación
        'clave': 3,          # C: Nombre del Expediente
        'nombre': 4,         # D: Descripción
        'caracter': 5,       # E: Valor documental
        # F (6): Soporte Documental (se deja vacío)
        'localizacion': 7,   # G: Ubicación
        'mes_ano': 8,        # H: Ejercicio
        'observacion': 9,    # I: Observaciones
    }

    for row_num, record in enumerate(records, start_row):
        for field, col in column_mapping.items():
            ws.cell(row=row_num, column=col, value=record.get(field, ''))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============== MÓDULO DE CONCENTRACIÓN (ANEXO 31) ==============

def extract_table_data_concentracion(pdf_path):
    """Extrae datos de tabla del PDF de Concentración usando recorte por coordenadas Y."""
    all_records = []
    current_record = None

    column_header_indicators = [
        'NO. PROGRESIVO', 'PROGRESIVO', 'NO. PROG',
        'CLAVE DEL EXPEDIENTE',
        'NOMBRE O DESCRIPCION', 'DESCRIPCIÓN DEL EXPEDIENTE',
        'AÑO DEL EXPED', 'ANO DEL EXPED', 'AÑO',
        'NO. DE HOJAS', 'HOJAS',
        'UBICACION TIPOGRAFICA', 'UBICACIÓN TIPOGRÁFICA', 'UBICACION',
        'CLASIFICACION', 'CLASIFICACIÓN',
        'CARACTER', 'CARÁCTER',
        'OBSERVACION', 'OBSERVACIÓN',
    ]

    def is_column_header_row(row_text):
        text_upper = row_text.upper()
        matches = sum(1 for kw in column_header_indicators if kw in text_upper)
        return matches >= 3

    def append_continuation(record, cleaned_row):
        for i, cell in enumerate(cleaned_row):
            if not cell:
                continue
            if i == 2:
                record['clave'] += ' ' + cell
            elif i == 3:
                record['nombre'] += ' ' + cell
            elif i == 4:
                if not record['ano']:
                    record['ano'] = cell
            elif i == 8:
                if not record['no_hojas']:
                    record['no_hojas'] = cell
            elif i == 10:
                record['ubicacion'] += ' ' + cell
            elif i == 11:
                if not record['clasificacion']:
                    record['clasificacion'] = cell
            elif i == 12:
                if not record['caracter']:
                    record['caracter'] = cell
            elif i == 13:
                record['observacion'] += ' ' + cell

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            top_crop = page.height * HEADER_PERCENT
            bottom_crop = page.height * (1 - FOOTER_PERCENT)
            cropped_page = page.crop((0, top_crop, page.width, bottom_crop))

            tables = cropped_page.extract_tables({
                'vertical_strategy': 'lines',
                'horizontal_strategy': 'lines',
                'snap_tolerance': 3,
                'join_tolerance': 3,
            })

            if not tables:
                continue

            for table in tables:
                for row in table:
                    if not row or all(cell is None or str(cell).strip() == '' for cell in row):
                        continue

                    cleaned_row = [clean_cell_text(cell) for cell in row]
                    full_row_text = ' '.join(cleaned_row)

                    if is_column_header_row(full_row_text):
                        continue

                    first_cell = cleaned_row[0] if cleaned_row else ''

                    if first_cell:
                        try:
                            prog_num = int(first_cell)
                            if current_record:
                                all_records.append(current_record)

                            while len(cleaned_row) < 14:
                                cleaned_row.append('')

                            current_record = {
                                'no_progresivo': prog_num,
                                'clave': cleaned_row[2],
                                'nombre': cleaned_row[3],
                                'ano': cleaned_row[4],
                                'no_hojas': cleaned_row[8],
                                'ubicacion': cleaned_row[10],
                                'clasificacion': cleaned_row[11],
                                'caracter': cleaned_row[12],
                                'observacion': cleaned_row[13],
                            }
                        except (ValueError, TypeError):
                            if current_record:
                                append_continuation(current_record, cleaned_row)
                    elif current_record:
                        append_continuation(current_record, cleaned_row)

    if current_record:
        all_records.append(current_record)

    return normalize_records(all_records)


def create_xlsx_concentracion_bytes(records):
    """Genera el archivo XLSX de Concentración en memoria usando la plantilla anexo_31."""
    template_path = os.path.join(os.path.dirname(__file__), 'anexo_31_template.xlsx')

    if os.path.exists(template_path):
        wb = load_workbook(template_path)
        ws = wb.active
        start_row = 5
        for merged_range in list(ws.merged_cells.ranges):
            if merged_range.min_row >= start_row:
                ws.unmerge_cells(str(merged_range))
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = 'Concentración'
        start_row = 2
        headers = [
            'No.', 'Nombre del expediente', 'Sección', 'Código', 'Serie',
            'No. Cajas', 'No. Expedientes', 'Del', 'Al', 'Localización',
            'Ejercicio', 'Clasificación', 'Valor documental', 'Observaciones'
        ]
        for col, header in enumerate(headers, 1):
            ws.cell(row=1, column=col, value=header)

    column_mapping = {
        'no_progresivo': 1,  # A: No.
        'nombre': 2,         # B: Nombre del expediente
        'clave': 4,          # D: Código
        'no_hojas': 9,       # I: Al
        'ubicacion': 10,     # J: Localización
        'ano': 11,           # K: Ejercicio
        'clasificacion': 12, # L: Clasificación
        'caracter': 13,      # M: Valor documental
        'observacion': 14,   # N: Observaciones
    }

    for row_num, record in enumerate(records, start_row):
        for field, col in column_mapping.items():
            ws.cell(row=row_num, column=col, value=record.get(field, ''))

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============== PROCESAMIENTO COMPARTIDO DE SUBIDAS ==============

def process_pdf_upload(extract_fn, create_xlsx_fn, download_name):
    """Valida, guarda con nombre único, extrae y retorna el XLSX en memoria."""
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'No se encontró archivo'}), 400

    file = request.files['file']

    if not file.filename:
        return jsonify({'status': 'error', 'message': 'No se seleccionó archivo'}), 400

    if not allowed_file(file.filename):
        return jsonify({'status': 'error', 'message': 'Tipo de archivo no permitido. Solo PDF.'}), 400

    pdf_path = os.path.join(UPLOAD_FOLDER, f'sentre_{uuid.uuid4().hex}.pdf')
    try:
        file.save(pdf_path)
        records = extract_fn(pdf_path)

        if not records:
            return jsonify({
                'status': 'error',
                'message': 'No se encontraron registros en el PDF. Verifica que corresponda al formato del módulo seleccionado.'
            }), 400

        xlsx_stream = create_xlsx_fn(records)
        update_counter(len(records))

        response = send_file(
            xlsx_stream,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )
        response.headers['X-Records-Count'] = str(len(records))
        return response

    except Exception as e:
        return jsonify({
            'status': 'error',
            'message': f'Error al procesar el documento: {str(e)}'
        }), 500
    finally:
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except OSError:
                pass


# ============== RUTAS ==============

@app.errorhandler(413)
def file_too_large(_error):
    return jsonify({
        'status': 'error',
        'message': 'El archivo excede el tamaño máximo permitido (50 MB).'
    }), 413


@app.route('/')
def home():
    return render_template('home.html')


@app.route('/departamento/<slug>')
def departamento_redirect(slug):
    """Compatibilidad con enlaces previos: redirige directamente a la selección de módulo."""
    return redirect(url_for('home'))


@app.route('/tramite')
def tramite():
    return render_template('index.html')


@app.route('/concentracion')
def concentracion():
    return render_template('concentracion.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    return process_pdf_upload(
        extract_table_data,
        create_xlsx_bytes,
        'inventario_tramite_anexo30.xlsx',
    )


@app.route('/upload_concentracion', methods=['POST'])
def upload_file_concentracion():
    return process_pdf_upload(
        extract_table_data_concentracion,
        create_xlsx_concentracion_bytes,
        'inventario_concentracion_anexo31.xlsx',
    )


@app.route('/status', methods=['GET'])
def status():
    return jsonify({'status': 'ready', 'message': 'Servidor listo'})


@app.route('/counter', methods=['GET'])
def get_counter_endpoint():
    """Devuelve el contador global actual de registros y archivos procesados."""
    return jsonify(get_counter())


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
