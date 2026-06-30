import json
import re
import sys
import io
import time
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

from flask import Flask, jsonify, request, send_file, Response, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

CORS(
    app,
    resources={r"/*": {"origins": ["http://localhost:5173", "http://127.0.0.1:5173"]}},
    supports_credentials=True,
    allow_headers=["Content-Type", "Authorization"],
    methods=["GET", "POST", "OPTIONS"],
)

ROOT = Path(__file__).resolve().parent
RULES_PATH = ROOT / 'rules.json'

# Global cache for validated files (In-memory storage)
VALIDATED_FILES_CACHE = {}

@app.errorhandler(Exception)
def handle_unexpected_error(error: Exception):
    print(f"[ERROR] Unhandled exception: {error}")
    return jsonify({"success": False, "message": "Excel processing failed"}), 500

@app.before_request
def log_request_context():
    if request.path in {"/analyze", "/validate"}:
        print(f"[REQUEST] {request.method} {request.path}")

def load_rules() -> Dict[str, Any]:
    with RULES_PATH.open('r', encoding='utf-8') as fh:
        return json.load(fh)

RULES = load_rules()

SUPPORTED_EXTENSIONS = {'.xlsx', '.xls'}
XLSX_MIME_TYPES = {
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
}
XLS_MIME_TYPES = {'application/vnd.ms-excel', 'application/xls'}


def detect_excel_format(filename: Optional[str], mime_type: Optional[str], file_bytes: Optional[bytes]) -> Dict[str, Any]:
    if not file_bytes:
        return {'kind': 'unsupported', 'reason': 'Uploaded file is empty'}

    if not isinstance(file_bytes, (bytes, bytearray, memoryview)):
        file_bytes = bytes(file_bytes)

    filename = (filename or '').lower()
    extension = Path(filename).suffix.lower()
    mime = (mime_type or '').lower()
    header = file_bytes[:8]
    has_zip_signature = header.startswith(b'PK\x03\x04')
    has_ole_signature = header.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1')

    unsupported_extensions = {'.csv', '.xlsm', '.xltx', '.xltm', '.txt', '.invalid'}
    if extension in unsupported_extensions:
        return {'kind': 'unsupported', 'reason': 'Unsupported Excel format'}

    if extension == '.xlsx':
        if mime in XLSX_MIME_TYPES or mime in {'application/zip', 'application/octet-stream', ''} or has_zip_signature:
            return {'kind': 'xlsx', 'reason': None}
        return {'kind': 'unsupported', 'reason': 'Unsupported Excel format'}

    if extension == '.xls':
        if mime in XLS_MIME_TYPES or mime in {'application/octet-stream', ''} or has_ole_signature:
            return {'kind': 'xls', 'reason': None}
        return {'kind': 'unsupported', 'reason': 'Unsupported Excel format'}

    if has_zip_signature and mime in XLSX_MIME_TYPES or mime in {'application/zip', 'application/octet-stream', ''}:
        return {'kind': 'xlsx', 'reason': None}

    if has_ole_signature and mime in XLS_MIME_TYPES or mime in {'application/octet-stream', ''}:
        return {'kind': 'xls', 'reason': None}

    return {'kind': 'unsupported', 'reason': 'Unsupported Excel format'}


def _to_cell_text(value: Any) -> str:
    if value is None:
        return ''
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).rstrip('0').rstrip('.') if '.' in str(value) else str(value)
    if isinstance(value, bool):
        return str(value)
    return str(value).strip()


def parse_excel_workbook(file_bytes: bytes, filename: Optional[str], mime_type: Optional[str]) -> tuple[List[str], List[Dict[str, str]], str]:
    detection = detect_excel_format(filename, mime_type, file_bytes)
    if detection['kind'] == 'unsupported':
        raise ValueError(detection['reason'])

    print(f"[UPLOAD] Detected type: {detection['kind']}")

    if detection['kind'] == 'xlsx':
        print('[PARSER] Parser selected: openpyxl')
        try:
            workbook = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
            worksheet = workbook.active
            sheet_name = worksheet.title or 'Sheet1'
            if sheet_name == 'Sheet':
                sheet_name = 'Sheet1'
            headers: List[str] = []
            rows: List[Dict[str, str]] = []

            for row_index, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
                values = [_to_cell_text(value) for value in row]
                if row_index == 1:
                    headers = values
                    continue
                if not any(values):
                    continue
                row_data = {}
                for col_idx, header in enumerate(headers):
                    row_data[header or f'Column{col_idx + 1}'] = values[col_idx] if col_idx < len(values) else ''
                rows.append(row_data)

            workbook.close()
            print(f"[UPLOAD] Rows extracted: {len(rows)}")
            return headers, rows, sheet_name
        except Exception as exc:
            print(f"[PARSER] XLSX parse failed: {exc}")
            raise ValueError('Uploaded Excel file is corrupted.') from exc

    print('[PARSER] Parser selected: xlrd')
    try:
        import xlrd

        workbook = xlrd.open_workbook(file_contents=file_bytes)
        worksheet = workbook.sheet_by_index(0)
        headers: List[str] = []
        rows: List[Dict[str, str]] = []

        for row_index in range(worksheet.nrows):
            values = [_to_cell_text(worksheet.cell_value(row_index, col_index)) for col_index in range(worksheet.ncols)]
            if row_index == 0:
                headers = values
                continue
            if not any(values):
                continue
            row_data = {}
            for col_idx, header in enumerate(headers):
                row_data[header or f'Column{col_idx + 1}'] = values[col_idx] if col_idx < len(values) else ''
            rows.append(row_data)

        print(f"[UPLOAD] Rows extracted: {len(rows)}")
        return headers, rows, worksheet.name
    except Exception as exc:
        print(f"[PARSER] XLS parse failed: {exc}")
        raise ValueError('Uploaded Excel file is corrupted.') from exc


PROFILE_HEADERS = {
    'raast': [
        'Amount', 'BeneficiaryBankCode', 'BeneficiaryIBAN', 'BeneficiaryName', 'BeneficiaryCode',
        'ReferenceField1', 'ReferenceField2', 'BeneficiaryEmail', 'BeneficiaryMobile',
        'BeneficiaryIdentificationType', 'BeneficiaryIdentificationNo', 'ProductTypeCode'
    ],
    'web_based': [
        'Amount', 'BeneficiaryBankCode', 'BeneficiaryAccountNo', 'BeneficiaryName', 'BeneficiaryCode',
        'ReferenceField1', 'ReferenceField2', 'BeneficiaryEmail', 'BeneficiaryMobile', 'ProductTypeCode'
    ],
    '1link': [
        'Amount', 'BankName', 'BeneficiaryAccountNumber', 'BeneficiaryName', 'BeneficiaryCode',
        'ReferenceField1', 'ReferenceField2', 'BeneficiaryEmail', 'BeneficiaryNumber'
    ],
}

def clean_cell_value(cell) -> str:
    val = cell.value
    if val is None:
        return ""
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        val_str = f"{val:.20f}"
        if '.' in val_str:
            val_str = val_str.rstrip('0').rstrip('.')
        return val_str
    if isinstance(val, int):
        return str(val)
    return str(val).strip()

def detect_profile(headers: List[str]) -> str:
    normalized = {h.strip(): h for h in headers if h}
    
    # Check for direct indicators of profile
    if any(h in normalized for h in ['BeneficiaryIBAN', 'BeneficiaryIdentificationType', 'BeneficiaryIdentificationNo']):
        return 'raast'
    if any(h in normalized for h in ['BeneficiaryAccountNo']):
        return 'web_based'
    if any(h in normalized for h in ['BeneficiaryAccountNumber', 'BeneficiaryNumber']):
        return '1link'
        
    # Check exact matching fallback
    for profile_name, expected_headers in PROFILE_HEADERS.items():
        if all(header in normalized for header in expected_headers):
            return profile_name
            
    return 'unknown'


def validate_rows(rows: List[Dict[str, str]], profile: str, rules: Dict[str, Any], debug: bool = False) -> List[Dict[str, Any]]:
    headers = PROFILE_HEADERS.get(profile, [])
    return validate_excel_data(rows, headers, profile, rules, debug=debug)


def validate_excel_data(rows: List[Dict[str, str]], headers: List[str], profile: str, rules: Dict[str, Any], debug: bool = False) -> List[Dict[str, Any]]:
    header_map = {h.strip(): i for i, h in enumerate(headers) if h}
    
    # Precompute duplicates
    duplicate_tracker = {
        "BeneficiaryAccountNo": {},
        "BeneficiaryAccountNumber": {}
    }
    for idx, row in enumerate(rows):
        for field in duplicate_tracker.keys():
            val = row.get(field, '').strip()
            if val:
                duplicate_tracker[field].setdefault(val, []).append(idx)

    validation_results = []
    
    for idx, row in enumerate(rows):
        errors = []
        profile_rules = rules['profiles'][profile]
        mandatory_fields = profile_rules['mandatory']
        
        # 1. Mandatory Fields Check
        for field in mandatory_fields:
            val = row.get(field, '').strip()
            if not val:
                errors.append({
                    "col": header_map.get(field),
                    "field": field,
                    "msg": f"{field} is required",
                    "expected": "Non-empty value",
                    "actual": "(Empty)"
                })

        # 2. Amount Check
        amount_val = row.get('Amount', '').strip()
        if amount_val:
            try:
                amount_float = float(amount_val)
                if amount_float <= 0:
                    errors.append({
                        "col": header_map.get('Amount'),
                        "field": "Amount",
                        "msg": "Amount must be greater than 0",
                        "expected": "Number > 0",
                        "actual": amount_val
                    })
            except ValueError:
                errors.append({
                    "col": header_map.get('Amount'),
                    "field": "Amount",
                    "msg": "Amount must be a numeric value",
                    "expected": "Numeric value",
                    "actual": amount_val
                })

        # 3. Mobile/Telecom check
        mobile_field = 'BeneficiaryNumber' if profile == '1link' else 'BeneficiaryMobile'
        mobile_val = row.get(mobile_field, '').strip()
        if mobile_val:
            if not re.fullmatch(r'\+?\d{10,15}', mobile_val):
                errors.append({
                    "col": header_map.get(mobile_field),
                    "field": mobile_field,
                    "msg": "Mobile number must be 11 digits",
                    "expected": "Telecom format (11 digits)",
                    "actual": mobile_val
                })

        # 4. Identity validation (Profile A/raast only)
        if profile == 'raast':
            id_type = row.get('BeneficiaryIdentificationType', '').strip()
            id_no = row.get('BeneficiaryIdentificationNo', '').strip()
            iban = row.get('BeneficiaryIBAN', '').strip()
            
            if id_type:
                if id_type not in ['CNIC', 'NTN']:
                    errors.append({
                        "col": header_map.get('BeneficiaryIdentificationType'),
                        "field": "BeneficiaryIdentificationType",
                        "msg": "Identification Type must be CNIC or NTN",
                        "expected": "CNIC or NTN",
                        "actual": id_type
                    })
                if id_no:
                    if id_type == 'CNIC':
                        if not id_no.isdigit():
                            errors.append({
                                "col": header_map.get('BeneficiaryIdentificationNo'),
                                "field": "BeneficiaryIdentificationNo",
                                "msg": "CNIC must be numeric",
                                "expected": "Numeric value",
                                "actual": id_no
                            })
                        elif len(id_no) != 13:
                            errors.append({
                                "col": header_map.get('BeneficiaryIdentificationNo'),
                                "field": "BeneficiaryIdentificationNo",
                                "msg": "CNIC must be exactly 13 digits",
                                "expected": "13 digits (numeric)",
                                "actual": id_no
                            })
                    elif id_type == 'NTN':
                        if not re.fullmatch(r'[A-Za-z0-9]{8}', id_no):
                            if len(id_no) != 8:
                                errors.append({
                                    "col": header_map.get('BeneficiaryIdentificationNo'),
                                    "field": "BeneficiaryIdentificationNo",
                                    "msg": "NTN must be exactly 8 characters",
                                    "expected": "8 chars",
                                    "actual": id_no
                                })
                            else:
                                errors.append({
                                    "col": header_map.get('BeneficiaryIdentificationNo'),
                                    "field": "BeneficiaryIdentificationNo",
                                    "msg": "NTN must be alphanumeric",
                                    "expected": "Alphanumeric",
                                    "actual": id_no
                                })
                else:
                    errors.append({
                        "col": header_map.get('BeneficiaryIdentificationNo'),
                        "field": "BeneficiaryIdentificationNo",
                        "msg": "Identification Number is required",
                        "expected": "ID Number",
                        "actual": "(Empty)"
                    })
            elif id_no:
                errors.append({
                    "col": header_map.get('BeneficiaryIdentificationType'),
                    "field": "BeneficiaryIdentificationType",
                    "msg": "Identification Type is required",
                    "expected": "CNIC or NTN",
                    "actual": "(Empty)"
                })

            # IBAN check
            if iban:
                if len(iban) != 24 or not iban.isalnum() or not iban.isupper():
                    errors.append({
                        "col": header_map.get('BeneficiaryIBAN'),
                        "field": "BeneficiaryIBAN",
                        "msg": "IBAN must be uppercase and 24 characters long",
                        "expected": "24-char uppercase",
                        "actual": iban
                    })

        # 5. Duplicate Checks
        for field in duplicate_tracker.keys():
            val = row.get(field, '').strip()
            if val and len(duplicate_tracker[field][val]) > 1:
                errors.append({
                    "col": header_map.get(field),
                    "field": field,
                    "msg": f"Duplicate {field} detected",
                    "expected": "Unique value",
                    "actual": val,
                    "is_duplicate": True
                })

        if debug:
            profile_label = {
                'raast': 'Raast',
                'web_based': 'Web-Based',
                '1link': '1Link'
            }.get(profile, profile)
            name = row.get('BeneficiaryName', '').strip() or '(blank)'
            if profile == 'raast':
                identifier = row.get('BeneficiaryIdentificationNo', '').strip() or '(blank)'
                id_type = row.get('BeneficiaryIdentificationType', '').strip() or '(blank)'
                detail = f"Name: {name}, Type: {id_type}, ID: {identifier}"
            elif profile == 'web_based':
                identifier = row.get('BeneficiaryAccountNo', '').strip() or '(blank)'
                detail = f"Name: {name}, Account: {identifier}"
            else:
                identifier = row.get('BeneficiaryAccountNumber', '').strip() or '(blank)'
                detail = f"Name: {name}, Account: {identifier}"

            if errors:
                reason = '; '.join(err['msg'] for err in errors)
                print(f"[DEBUG] Processing Row {idx + 1} [Profile: {profile_label}] -> {detail} -> STATUS: FAIL [Reason: {reason}]")
            else:
                bypass_suffix = ' (Identity check bypassed)' if profile in {'web_based', '1link'} else ''
                print(f"[DEBUG] Processing Row {idx + 1} [Profile: {profile_label}] -> {detail} -> STATUS: PASS{bypass_suffix}")

        validation_results.append({
            "status": "PASS" if not errors else "FAIL",
            "errors": errors
        })
        
    return validation_results

def generate_highlighted_validation_report(rows, results, headers, profile):
    """Generate a validation report with all rows, highlighting error cells with red background and comments."""
    normalized_headers = list(headers or [])
    if not normalized_headers:
        normalized_headers = ['Column1']

    # Create the report workbook
    wb_report = openpyxl.Workbook()
    ws_report = wb_report.active
    ws_report.title = 'Validation Report'
    
    # Add headers
    ws_report.append(normalized_headers)
    
    # Define styling
    red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
    white_bold_font = Font(color='FFFFFFFF', bold=True)
    
    # Create a mapping from field names to column indices for easy error cell highlighting
    field_to_col = {header: idx + 1 for idx, header in enumerate(normalized_headers)}
    
    # Add all rows with error highlighting
    for idx, row in enumerate(rows):
        if idx >= len(results):
            break
        
        result = results[idx]
        row_data = [row.get(header, '') for header in normalized_headers]
        
        # Add the row
        ws_report.append(row_data)
        current_row = idx + 2  # +2 because row 1 is headers, +1 for 1-based indexing
        
        # Highlight error cells and add comments
        if result['errors']:
            for error in result['errors']:
                field_name = error.get('field', '')
                col_idx = field_to_col.get(field_name)
                
                if col_idx:
                    cell = ws_report.cell(row=current_row, column=col_idx)
                    # Apply red background and white bold font
                    cell.fill = red_fill
                    cell.font = white_bold_font
                    # Add comment with error message
                    error_msg = error.get('msg', 'Validation Error')
                    if error.get('expected'):
                        error_msg += f"\nExpected: {error['expected']}"
                    if error.get('actual'):
                        error_msg += f"\nActual: {error['actual']}"
                    cell.comment = Comment(error_msg, 'Validator')
    
    # Format the report
    for col in ws_report.columns:
        max_len = max(len(str(cell.value or '')) for cell in col) if col else 12
        col_letter = get_column_letter(col[0].column) if col else 'A'
        ws_report.column_dimensions[col_letter].width = max(max_len + 3, 12)
    
    if ws_report.max_row > 1:
        ws_report.auto_filter.ref = f"A1:{get_column_letter(ws_report.max_column)}{ws_report.max_row}"
        ws_report.views.sheetView[0].showGridLines = True
        ws_report.freeze_panes = 'A2'
    
    # Serialize to bytes
    report_stream = io.BytesIO()
    wb_report.save(report_stream)
    report_stream.seek(0)
    
    return report_stream.getvalue()

def generate_validated_excel_streams(rows, results, headers, profile):
    """Generate XLSX output workbooks from normalized rows and validation results.
    
    Both .xlsx and .xls inputs are normalized into the same row structure by the parser,
    so this function always outputs modern .xlsx format regardless of input format.
    """
    normalized_headers = list(headers or [])
    if not normalized_headers:
        normalized_headers = ['Column1']

    # Create workbooks for passed and rejected records
    wb_passed = openpyxl.Workbook()
    ws_passed = wb_passed.active
    ws_passed.title = 'Sheet'
    ws_passed.append(normalized_headers + ['Validation Status'])

    wb_rejected = openpyxl.Workbook()
    ws_rejected = wb_rejected.active
    ws_rejected.title = 'Sheet'
    ws_rejected.append(normalized_headers + ['Validation Status', 'Validation Errors'])

    red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
    white_bold_font = Font(color='FFFFFFFF', bold=True)

    # Build the passed and rejected workbooks row by row
    for idx, row in enumerate(rows):
        if idx >= len(results):
            break
        
        result = results[idx]
        row_values = [row.get(h, '') for h in normalized_headers]

        if result['status'] == 'PASS':
            ws_passed.append(row_values + ['PASS'])
        else:
            # Build error message from all errors in this row
            err_messages = [err['msg'] for err in result.get('errors', [])]
            ws_rejected.append(row_values + ['FAIL', '; '.join(err_messages)])
            
            # Color the failed cells red and add comments
            for err in result.get('errors', []):
                col_idx = err.get('col')
                if col_idx is not None and col_idx < len(normalized_headers):
                    cell = ws_rejected.cell(row=idx + 2, column=col_idx + 1)
                    cell.fill = red_fill
                    cell.font = white_bold_font
                    comment_text = f"Rule: {err.get('field', 'Validation')}\nExpected: {err['expected']}\nActual: {err['actual']}"
                    cell.comment = Comment(comment_text, 'SEVS')

    # Format both workbooks
    for ws in [ws_passed, ws_rejected]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col) if col else 12
            col_letter = get_column_letter(col[0].column) if col else 'A'
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
        if ws.max_row > 1:
            ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
            ws.views.sheetView[0].showGridLines = True
            ws.freeze_panes = 'A2'

    # Serialize to bytes
    passed_stream = io.BytesIO()
    wb_passed.save(passed_stream)
    passed_stream.seek(0)

    rejected_stream = io.BytesIO()
    wb_rejected.save(rejected_stream)
    rejected_stream.seek(0)

    return passed_stream.getvalue(), rejected_stream.getvalue()

@app.route('/')
def index():
    return send_from_directory(ROOT, 'index.html')

@app.route('/style.css')
def serve_css():
    return send_from_directory(ROOT, 'style.css')

@app.route('/tailwind.generated.css')
def serve_tailwind_css():
    return send_from_directory(ROOT, 'tailwind.generated.css')

@app.route('/main.js')
def serve_js():
    return send_from_directory(ROOT, 'main.js')

@app.route('/rules.json')
def serve_rules():
    return send_from_directory(ROOT, 'rules.json')

@app.route('/health', methods=['GET'])
def health():
    return jsonify({"status": "ok"})

@app.route('/analyze', methods=['POST'])
def analyze_endpoint():
    print('[UPLOAD] Received upload')
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "No file uploaded"}), 400

    uploaded_file = request.files['file']
    if uploaded_file.filename == '':
        return jsonify({"success": False, "message": "No selected file"}), 400

    try:
        file_bytes = uploaded_file.read()
    except Exception:
        return jsonify({"success": False, "message": "Uploaded file could not be read"}), 400

    if not file_bytes:
        return jsonify({"success": False, "message": "Uploaded file is empty"}), 400

    filename = uploaded_file.filename or ''
    mime_type = uploaded_file.mimetype or ''
    extension = Path(filename).suffix.lower()
    print(f"[UPLOAD] Extension: {extension} | MIME: {mime_type} | Size: {len(file_bytes)}")

    try:
        headers, rows, sheet_name = parse_excel_workbook(file_bytes, filename, mime_type)
    except ValueError as exc:
        return jsonify({"success": False, "message": str(exc)}), 400
    except Exception as exc:
        print(f"[UPLOAD] Excel read failed: {exc}")
        return jsonify({"success": False, "message": "Excel format unsupported"}), 400

    detected_profile = detect_profile(headers)
    print(f"[UPLOAD] Workbook loaded: sheet={sheet_name}, rows={len(rows)}")
    return jsonify({
        "success": True,
        "filename": uploaded_file.filename,
        "row_count": len(rows),
        "sheet_name": sheet_name,
        "profile": detected_profile
    })

@app.route('/validate', methods=['POST'])
def validate_endpoint():
    if 'file' not in request.files:
        return jsonify({"success": False, "message": "No file part in request"}), 400

    uploaded_file = request.files['file']
    if uploaded_file.filename == '':
        return jsonify({"success": False, "message": "No selected file"}), 400

    selected_profile = request.form.get('profile', 'auto')

    # Read rules dynamically
    global RULES
    RULES = load_rules()

    try:
        file_bytes = uploaded_file.read()
    except Exception:
        return jsonify({"success": False, "message": "Uploaded file could not be read"}), 400

    if not file_bytes:
        return jsonify({"success": False, "message": "Uploaded file is empty"}), 400

    filename = uploaded_file.filename or ''
    mime_type = uploaded_file.mimetype or ''
    print(f"[VALIDATION] Received upload: filename={filename}, mime={mime_type}, size={len(file_bytes)}")

    try:
        headers, rows, _ = parse_excel_workbook(file_bytes, filename, mime_type)
    except ValueError as exc:
        return jsonify({"success": False, "message": str(exc)}), 400
    except Exception as exc:
        print(f"[VALIDATION] Excel read failed: {exc}")
        return jsonify({"success": False, "message": "Excel format unsupported"}), 400

    print(f"[VALIDATION] Workbook loaded: rows={len(rows)}")

    # Profile Detection
    detected_profile = detect_profile(headers)
    active_profile = selected_profile if selected_profile != 'auto' else detected_profile

    if active_profile == 'unknown':
        return jsonify({
            "error": "Could not auto-detect file profile. Please select profile manually from the dropdown."
        }), 400

    # Run validations
    start_time = time.time()
    results = validate_excel_data(rows, headers, active_profile, RULES)
    end_time = time.time()
    duration = f"{end_time - start_time:.2f} sec"

    # Generate the two files from the normalized rows
    passed_file, rejected_file = generate_validated_excel_streams(rows, results, headers, active_profile)
    
    # Generate the highlighted validation report
    highlighted_report = generate_highlighted_validation_report(rows, results, headers, active_profile)

    # Extract original filename without extension, and get file extension
    original_filename = uploaded_file.filename or 'ValidationReport'
    file_name_base = Path(original_filename).stem  # Name without extension
    file_extension = Path(original_filename).suffix or '.xlsx'  # Keep original extension or default to .xlsx

    # Cache the resulting byte streams with original filename
    file_id = str(uuid.uuid4())
    VALIDATED_FILES_CACHE[file_id] = {
        "passed": passed_file,
        "rejected": rejected_file,
        "highlighted": highlighted_report,
        "filename_base": file_name_base,
        "file_extension": file_extension,
        "passed_count": sum(1 for r in results if r['status'] == 'PASS'),
        "rejected_count": sum(1 for r in results if r['status'] == 'FAIL'),
        "total_count": len(rows),
        "full_results": results  # Store full results, not just first 20
    }

    # Generate summary lists
    error_breakdown = {
        "Invalid IBAN": 0,
        "Invalid Mobile": 0,
        "Invalid CNIC": 0,
        "Duplicate Records": 0,
        "Missing Required Fields": 0
    }

    for res in results:
        for err in res['errors']:
            msg = err['msg']
            field = err['field']
            if field == 'BeneficiaryIBAN':
                error_breakdown["Invalid IBAN"] += 1
            elif field in ['BeneficiaryMobile', 'BeneficiaryNumber']:
                error_breakdown["Invalid Mobile"] += 1
            elif field == 'BeneficiaryIdentificationNo' and 'CNIC' in msg:
                error_breakdown["Invalid CNIC"] += 1
            elif err.get('is_duplicate'):
                error_breakdown["Duplicate Records"] += 1
            elif 'is required' in msg or 'missing' in msg:
                error_breakdown["Missing Required Fields"] += 1

    return jsonify({
        "file_id": file_id,
        "filename": original_filename,
        "filename_base": file_name_base,
        "file_extension": file_extension,
        "profile": active_profile,
        "headers": headers,
        "rows": rows[:20],
        "results": results[:20],
        "duration": duration,
        "summary": {
            "total": len(rows),
            "passed": VALIDATED_FILES_CACHE[file_id]["passed_count"],
            "rejected": VALIDATED_FILES_CACHE[file_id]["rejected_count"],
            "errors_breakdown": error_breakdown
        }
    })

@app.route('/api/analyze', methods=['POST'])
def analyze_api_endpoint():
    return analyze_endpoint()

@app.route('/api/validate', methods=['POST'])
def validate_api_endpoint():
    return validate_endpoint()

@app.route('/download/passed/<file_id>', methods=['GET'])
def download_passed(file_id):
    if file_id not in VALIDATED_FILES_CACHE:
        return "File not found or expired", 404
    data = VALIDATED_FILES_CACHE[file_id]["passed"]
    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name="Passed_Records.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route('/api/download/passed/<file_id>', methods=['GET'])
def download_passed_api(file_id):
    return download_passed(file_id)

@app.route('/download/rejected/<file_id>', methods=['GET'])
def download_rejected(file_id):
    if file_id not in VALIDATED_FILES_CACHE:
        return "File not found or expired", 404
    data = VALIDATED_FILES_CACHE[file_id]["rejected"]
    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name="Rejected_Records.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route('/api/download/rejected/<file_id>', methods=['GET'])
def download_rejected_api(file_id):
    return download_rejected(file_id)

if __name__ == '__main__':
    print('Server started on port 5001')
    app.run(host='127.0.0.1', port=5001, debug=True)
