from pathlib import Path
from io import BytesIO
from tempfile import TemporaryDirectory
import os
import json
from datetime import datetime, timezone, timedelta
import sqlite3
from threading import RLock
from contextlib import contextmanager
from urllib.parse import quote

try:
    import fcntl
except ImportError:
    fcntl = None

STOCK_LOCK = RLock()

from flask import Flask, jsonify, make_response, request, send_file
from openpyxl import load_workbook


app = Flask(__name__)
STOCK_PATH = Path(os.environ.get("STOCK_DATA_DIR", Path(__file__).resolve().parent.parent / "data")) / "stock.xlsx"


@contextmanager
def stock_lock():
    # PythonAnywhere can serve requests in multiple worker processes.
    with STOCK_LOCK:
        STOCK_PATH.parent.mkdir(parents=True, exist_ok=True)
        with STOCK_PATH.with_suffix(".lock").open("a+b") as lock_file:
            if fcntl is not None:
                fcntl.flock(lock_file, fcntl.LOCK_EX)
            try:
                yield
            finally:
                if fcntl is not None:
                    fcntl.flock(lock_file, fcntl.LOCK_UN)

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
    response.headers["Access-Control-Expose-Headers"] = "X-Stock-Version, X-Stock-Rows, X-Stock-Quantity, X-Stock-Source, X-Stock-Imported-At"
    return response


@app.get("/")
def index():
    return jsonify({"ok": True, "service": "stock pdf converter"})


@app.get("/health")
def health():
    return jsonify({"ok": True})


def stock_version():
    stat = STOCK_PATH.stat()
    return f"{stat.st_mtime_ns}-{stat.st_size}"


def no_store(response):
    response.headers["Cache-Control"] = "no-store"
    return response


def publish_stock(content, source=None):
    STOCK_PATH.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(dir=STOCK_PATH.parent) as tmp:
        staged = Path(tmp) / "stock.xlsx"
        staged.write_bytes(content)
        os.replace(staged, STOCK_PATH)
    if source is not None:
        metadata = {"source": source, "imported_at": datetime.now(timezone.utc).isoformat()}
        STOCK_PATH.with_suffix(".metadata.json").write_text(json.dumps(metadata), encoding="utf-8")
    return stock_version()


def stock_metadata():
    try:
        return json.loads(STOCK_PATH.with_suffix(".metadata.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"source": STOCK_PATH.name, "imported_at": datetime.fromtimestamp(STOCK_PATH.stat().st_mtime, timezone.utc).isoformat()}


@app.get("/api/stock-version")
def get_stock_version():
    version = stock_version() if STOCK_PATH.exists() else None
    return no_store(jsonify({"version": version}))


@app.get("/api/stock")
def get_stock():
    with stock_lock():
        if not STOCK_PATH.exists():
            return no_store(jsonify({"detail": "No shared stock has been uploaded yet."})), 404
        content = STOCK_PATH.read_bytes()
        version = stock_version()
        metadata = stock_metadata()
    response = send_file(BytesIO(content), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response.headers["X-Stock-Version"] = version
    response.headers["X-Stock-Source"] = quote(metadata["source"])
    response.headers["X-Stock-Imported-At"] = metadata.get("imported_at") or ""
    return no_store(response)


def business_day():
    return datetime.now(timezone(timedelta(hours=7))).date().isoformat()


def daily_database():
    STOCK_PATH.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(STOCK_PATH.with_name("daily-stock.sqlite3"), timeout=15)
    connection.row_factory = sqlite3.Row
    connection.execute("""CREATE TABLE IF NOT EXISTS movements (
        id INTEGER PRIMARY KEY AUTOINCREMENT, request_id TEXT UNIQUE NOT NULL,
        day TEXT NOT NULL, stock_version TEXT NOT NULL, product TEXT NOT NULL,
        delta INTEGER NOT NULL, quantity INTEGER NOT NULL, created_at TEXT NOT NULL
    )""")
    connection.execute("CREATE INDEX IF NOT EXISTS movements_day ON movements(day, stock_version)")
    connection.commit()
    return connection


def daily_state(connection, version):
    day = business_day()
    events = []
    products = {}
    for row in connection.execute("SELECT * FROM movements WHERE day = ? ORDER BY id", (day,)):
        event = {"id": row["id"], "path": json.loads(row["product"]), "delta": row["delta"],
                 "quantity": row["quantity"], "at": row["created_at"], "stock_version": row["stock_version"]}
        events.append(event)
        if row["stock_version"] == version:
            products[row["product"]] = {"path": event["path"], "quantity": event["quantity"]}
    return {"day": day, "version": version, "revision": events[-1]["id"] if events else 0,
            "products": list(products.values()), "events": events}


@app.get("/api/daily-stock")
def get_daily_stock():
    with stock_lock():
        if not STOCK_PATH.exists():
            return no_store(jsonify({"detail": "Stock not found."})), 404
        with daily_database() as connection:
            state = daily_state(connection, stock_version())
        connection.close()
    return no_store(jsonify(state))


@app.route("/api/stock/quantity", methods=["POST", "OPTIONS"])
@app.route("/api/sell-out", methods=["POST", "OPTIONS"])
def update_quantity():
    if request.method == "OPTIONS":
        return make_response("", 204)
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return jsonify({"detail": "Invalid request."}), 400
    keys, qty = data.get("path"), data.get("quantity")
    expected = data.get("expected_quantity")
    request_id = data.get("request_id")
    if (not isinstance(keys, list) or len(keys) != 4 or not all(isinstance(k, str) and 0 < len(k) <= 500 for k in keys)
            or type(qty) is not int or not 0 <= qty <= 1000000
            or type(expected) is not int or not 0 <= expected <= 1000000
            or not isinstance(request_id, str) or not 8 <= len(request_id) <= 100):
        return jsonify({"detail": "Invalid product, quantity or request ID."}), 400
    with stock_lock():
        if not STOCK_PATH.exists():
            return jsonify({"detail": "Stock not found."}), 404
        version = stock_version()
        if data.get("version") != version or data.get("day") != business_day():
            return no_store(jsonify({"detail": "ข้อมูลหรือวันที่เปลี่ยนแล้ว กรุณา Sync และตรวจสอบยอดใหม่"})), 409
        product = json.dumps(keys, ensure_ascii=False)
        connection = daily_database()
        try:
            connection.execute("BEGIN IMMEDIATE")
            previous = connection.execute("SELECT * FROM movements WHERE request_id = ?", (request_id,)).fetchone()
            if previous:
                if previous["product"] != product or previous["quantity"] != qty or previous["stock_version"] != version or previous["day"] != business_day():
                    return jsonify({"detail": "Request ID already used."}), 409
                return no_store(jsonify({"success": True, "quantity": qty, **daily_state(connection, version)}))
            workbook = load_workbook(STOCK_PATH, read_only=True, data_only=True)
            base_quantity = 0
            found = False
            try:
                for sheet in workbook.worksheets:
                    if sheet.title.strip().lower() in {"คำแนะนำ", "instructions", "instruction", "readme", "info"}:
                        continue
                    values = sheet.iter_rows(values_only=True)
                    headers = {str(v or "").strip().lower(): i for i, v in enumerate(next(values, []))}
                    if not {"model", "storage", "color", "qty"}.issubset(headers):
                        continue
                    for row in values:
                        def value(name):
                            return str(row[headers[name]] if row[headers[name]] is not None else "").strip()
                        key = [value("category") if "category" in headers else sheet.title.strip(), value("model"), value("storage"), value("color")]
                        if key == keys:
                            found = True
                            base_quantity += max(0, int(float(row[headers["qty"]] or 0)))
            finally:
                workbook.close()
            if not found:
                return jsonify({"detail": "Product not found."}), 404
            latest = connection.execute("SELECT quantity FROM movements WHERE day = ? AND stock_version = ? AND product = ? ORDER BY id DESC LIMIT 1", (business_day(), version, product)).fetchone()
            current = latest["quantity"] if latest else base_quantity
            if expected != current:
                return no_store(jsonify({"detail": "มีผู้ใช้ปรับยอดแล้ว กรุณาตรวจสอบยอดล่าสุด", "quantity": current})), 409
            delta = qty - current
            if delta == 0:
                return jsonify({"detail": "Quantity has not changed."}), 400
            connection.execute("INSERT INTO movements(request_id, day, stock_version, product, delta, quantity, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
                               (request_id, business_day(), version, product, delta, qty, datetime.now(timezone.utc).isoformat()))
            connection.commit()
            return no_store(jsonify({"success": True, "quantity": qty, **daily_state(connection, version)}))
        finally:
            connection.close()


@app.route("/api/upload-stock", methods=["POST", "OPTIONS"])
def upload_stock():
    if request.method == "OPTIONS":
        return make_response("", 204)

    uploaded = request.files.get("file")
    if uploaded is None or not (uploaded.filename or "").lower().endswith(".xlsx"):
        return jsonify({"detail": "Upload an XLSX workbook."}), 400

    content = uploaded.read()
    try:
        workbook = load_workbook(BytesIO(content), read_only=True, data_only=True)
        usable = 0
        for sheet in workbook.worksheets:
            if sheet.title.strip().lower() in {"คำแนะนำ", "instructions", "instruction", "readme", "info"}:
                continue
            values = sheet.iter_rows(values_only=True)
            headers = next(values, None)
            if not headers:
                continue
            columns = {str(value or "").strip().lower(): i for i, value in enumerate(headers)}
            if not {"model", "storage", "color", "qty"}.issubset(columns):
                continue
            for row in values:
                try:
                    category = row[columns["category"]] if "category" in columns else sheet.title
                    required = [category, row[columns["model"]], row[columns["storage"]], row[columns["color"]]]
                    quantity = float(row[columns["qty"]])
                    if all(str(value or "").strip() for value in required) and quantity > 0:
                        usable += 1
                except (IndexError, TypeError, ValueError):
                    continue
        workbook.close()
        if not usable:
            return jsonify({"detail": "No usable stock rows found."}), 400
    except Exception:
        return jsonify({"detail": "Invalid XLSX workbook."}), 400

    with stock_lock():
        version = publish_stock(content, request.form.get("source_name") or uploaded.filename)
    return no_store(jsonify({"success": True, "version": version, "rows": usable}))


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
            if not rows:
                return jsonify({"detail": "No usable stock rows found in the PDF."}), 400
            write_workbook(rows, str(xlsx_path))
            output = BytesIO(xlsx_path.read_bytes())
            output.seek(0)
            with stock_lock():
                version = publish_stock(output.getvalue(), filename)

            response = send_file(
                output,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="stock.xlsx",
            )
            response.headers["X-Stock-Rows"] = str(len(rows))
            response.headers["X-Stock-Quantity"] = str(sum(row[4] for row in rows))
            response.headers["X-Stock-Version"] = version
            return no_store(response)
    except Exception as exc:
        return jsonify({"detail": f"Failed to process PDF: {exc}"}), 500


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
