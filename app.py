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
