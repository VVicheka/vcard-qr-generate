"""Flask Dashboard"""

import json
import os

from flask import Flask, render_template, request, send_file, jsonify

from qr_generator import (
    build_vcard,
    qr_png_bytes,
    read_excel_headers,
    read_excel_rows,
    build_excel_with_qr,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB cap

@app.route("/")
def index():
    return render_template("index.html")

@app.post("/api/generate-single")
def generate_single():
    first = request.form.get("first", "").strip()
    last = request.form.get("last", "").strip()
    if not first and not last:
        return jsonify({"error": "First or Last name is required"}), 400
    
    try:
        fields = json.loads(request.form.get("fields_json", "[]"))
    except json.JSONDecodeError:
        return jsonify({"error": "Invalid fields Json"}), 400
    
    logo_bytes = None
    if "logo" in request.files and request.files["logo"].filename:
        logo_bytes = request.files["logo"].read()

    vcard = build_vcard(first, last, fields)
    png = qr_png_bytes(vcard, logo_bytes=logo_bytes)

    filename = (request.form.get("filename") or f"{first}_{last}".strip("_") or "qr").strip()
    return send_file(png, mimetype="image/png", as_attachment=True, download_name=f"{filename}.png")

@app.post("/api/excel-columns")
def excel_columns():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    try:
        headers = read_excel_headers(request.files["file"])
    except Exception as e:
        return jsonify({"error": f"Could not read Excel: {e}"}), 400
    return jsonify({"columns": headers})


@app.post("/api/generate-bulk")
def generate_bulk():
    if "file" not in request.files:
        return jsonify({"error": "Excel file required"}), 400

    try:
        mapping = json.loads(request.form.get("mapping_json", "{}"))
    except json.JSONDecodeError:
        return jsonify({"error": "Invalid mapping JSON"}), 400

    if not mapping.get("first_col") or not mapping.get("last_col"):
        return jsonify({"error": "First Name and Last Name columns are required"}), 400

    logo_bytes = None
    if "logo" in request.files and request.files["logo"].filename:
        logo_bytes = request.files["logo"].read()

    try:
        _, rows = read_excel_rows(request.files["file"])
    except Exception as e:
        return jsonify({"error": f"Could not read Excel: {e}"}), 400

    if not rows:
        return jsonify({"error": "No data rows found"}), 400

    xlsx = build_excel_with_qr(rows, mapping, logo_bytes=logo_bytes)
    filename = (request.form.get("filename") or "contacts_with_qr").strip()
    return send_file(
        xlsx,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"{filename}.xlsx",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)