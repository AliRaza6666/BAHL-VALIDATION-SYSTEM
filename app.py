from backend.app import app as flask_app, ensure_runtime_directories, load_rules
from backend.app import (
    detect_excel_format,
    generate_highlighted_validation_report,
    generate_validation_summary_workbook,
    generate_validated_excel_streams,
    parse_excel_workbook,
    validate_excel_data,
    validate_rows,
)

app = flask_app


if __name__ == '__main__':
    print('Backend module executed directly.')
    print('This project is configured for serverless deployment via backend/api/index.py.')
    ensure_runtime_directories()
    load_rules()
    print('Runtime directories and rules are ready.')
    print('Starting local backend server at http://127.0.0.1:5000')
    app.run(host='127.0.0.1', port=5000, debug=True)
