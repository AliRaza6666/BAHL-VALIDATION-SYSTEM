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
    try:
        print('Backend starting')
        ensure_runtime_directories()
        load_rules()
        print('Rules loaded')
        print('Upload folder ready')
        print('Generated directory ready')
        print('Server running on: http://127.0.0.1:5001')
        flask_app.run(host='127.0.0.1', port=5001, debug=True, use_reloader=False)
    except Exception as exc:
        print(f"[ERROR] Startup failed: {exc}")
        raise
