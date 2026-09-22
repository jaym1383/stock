from pathlib import Path
from io import BytesIO
from tempfile import TemporaryDirectory

from flask import Flask, jsonify, make_response, request, send_file


app = Flask(__name__)

ALLOWED_ORIGINS = {
    "https://jaym1383.github.io",
    "http://localhost:8000",
    "http://127.0.0.1:8000",
}


@app.after_request
def add_cors_headers(response):
    origin = request.headers.get("Origin")
    if origin in ALLOWED_ORIGINS:
        response.headers["Access-Control-Allow-Origin"] = origin
        response.headers["Vary"] = "Origin"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return response


@app.get("/")
def index():
    return jsonify({"ok": True, "service": "stock pdf converter"})


@app.get("/health")
def health():
    return jsonify({"ok": True})


@app.route("/api/convert", methods=["POST", "OPTIONS"])
def convert_pdf():
    if request.method == "OPTIONS":
        return make_response("", 204)

    uploaded = request.files.get("file")
    if uploaded is None or uploaded.filename == "":
        return jsonify({"detail": "No PDF file was uploaded."}), 400

    filename = uploaded.filename or ""
    content_type = uploaded.content_type or ""
    if content_type != "application/pdf" and not filename.lower().endswith(".pdf"):
        return jsonify({"detail": "Only PDF files are supported."}), 400

    try:
        from pdf_to_stock import build_rows, extract_records, write_workbook
    except SystemExit as exc:
        return jsonify({"detail": str(exc) or "Converter dependency is missing."}), 500
    except ImportError as exc:
        return jsonify({"detail": f"Converter dependency is missing: {exc}"}), 500

    try:
        with TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            pdf_path = tmp_dir / "input.pdf"
            xlsx_path = tmp_dir / "stock.xlsx"

            uploaded.save(pdf_path)
            records = extract_records(str(pdf_path))
            rows = build_rows(records)
            write_workbook(rows, str(xlsx_path))
            output = BytesIO(xlsx_path.read_bytes())
            output.seek(0)

            response = send_file(
                output,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="stock.xlsx",
            )
            response.headers["X-Stock-Rows"] = str(len(rows))
            response.headers["X-Stock-Quantity"] = str(sum(row[4] for row in rows))
            return response
    except Exception as exc:
        return jsonify({"detail": f"Failed to process PDF: {exc}"}), 500


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
