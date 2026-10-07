
import io
import os
import re
from pathlib import Path

import pandas as pd
import pdfplumber
from flask import Flask, jsonify, render_template, request, send_file

MAX_MB = int(os.getenv("MAX_FILE_MB", "10"))
MAX_BYTES = MAX_MB * 1024 * 1024

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_BYTES


def clean_cell(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def extract_pdf(pdf_bytes: bytes):
    tables = []
    text_pages = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text() or ""
            text_pages.append((page_num, page_text.strip()))

            page_tables = page.extract_tables() or []
            for table_idx, table in enumerate(page_tables, start=1):
                cleaned = [[clean_cell(c) for c in row] for row in table]
                cleaned = [row for row in cleaned if any(cell for cell in row)]
                if not cleaned:
                    continue

                # Normalize rows to the same width.
                width = max(len(row) for row in cleaned)
                normalized = [row + [""] * (width - len(row)) for row in cleaned]

                # Use first non-empty row as a header where possible.
                header = normalized[0]
                if not any(header):
                    header = [f"Column {i+1}" for i in range(width)]
                    data = normalized
                else:
                    # Make duplicate/blank column names safe for Excel.
                    safe_header = []
                    seen = {}
                    for i, h in enumerate(header):
                        h = h or f"Column {i+1}"
                        seen[h] = seen.get(h, 0) + 1
                        safe_header.append(h if seen[h] == 1 else f"{h}_{seen[h]}")
                    header = safe_header
                    data = normalized[1:]

                tables.append({
                    "page": page_num,
                    "index": table_idx,
                    "header": header,
                    "rows": data,
                })
    return tables, text_pages


def build_xlsx(tables, text_pages):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        if tables:
            for n, table in enumerate(tables, start=1):
                df = pd.DataFrame(table["rows"], columns=table["header"])
                sheet = f"Table {n}"[:31]
                df.to_excel(writer, sheet_name=sheet, index=False)
        else:
            pd.DataFrame(
                [{"Page": p, "Extracted text": text} for p, text in text_pages]
            ).to_excel(writer, sheet_name="Text", index=False)

        # Always include a small processing summary.
        summary = pd.DataFrame([
            {"Metric": "Tables found", "Value": len(tables)},
            {"Metric": "Pages processed", "Value": len(text_pages)},
            {"Metric": "Text pages with content", "Value": sum(bool(t) for _, t in text_pages)},
        ])
        summary.to_excel(writer, sheet_name="Summary", index=False)

        workbook = writer.book
        for ws in workbook.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col_cells in ws.columns:
                max_len = 0
                col_letter = col_cells[0].column_letter
                for cell in col_cells:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = min(max(max_len, len(value)), 45)
                ws.column_dimensions[col_letter].width = max(10, max_len + 2)

    out.seek(0)
    return out


@app.get("/")
def home():
    return render_template(
        "index.html",
        payment_link=os.getenv("STRIPE_PAYMENT_LINK", ""),
        max_mb=MAX_MB,
    )


@app.post("/api/convert")
def convert():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify({"error": "Please select a PDF file."}), 400

    if not uploaded.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are supported."}), 400

    pdf_bytes = uploaded.read()
    if not pdf_bytes:
        return jsonify({"error": "The PDF file is empty."}), 400
    if len(pdf_bytes) > MAX_BYTES:
        return jsonify({"error": f"File is larger than {MAX_MB} MB."}), 400

    try:
        tables, text_pages = extract_pdf(pdf_bytes)
        workbook = build_xlsx(tables, text_pages)
    except Exception as exc:
        app.logger.exception("Conversion failed")
        return jsonify({"error": "Could not process this PDF. Try another file."}), 500

    base = Path(uploaded.filename).stem
    filename = f"{base}_dataforge.xlsx"
    response = send_file(
        workbook,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response.headers["X-DataForge-Tables"] = str(len(tables))
    response.headers["X-DataForge-Pages"] = str(len(text_pages))
    return response


@app.get("/privacy")
def privacy():
    return render_template("privacy.html")

@app.get("/terms")
def terms():
    return render_template("terms.html")

@app.get("/health")
def health():
    return {"status": "ok"}


@app.errorhandler(413)
def too_large(_):
    return jsonify({"error": f"File is larger than {MAX_MB} MB."}), 413


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")), debug=True)
