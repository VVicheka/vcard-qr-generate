"""Reusable QR + vCard generation logic.

Used by both the CLI script and the Flask web app.
"""

import io
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image
import qrcode


# vCard field type -> how it renders in the vCard text
VCARD_RENDERERS = {
    "phone":   lambda label, val: f"TEL;TYPE=CELL:{val}",
    "email":   lambda label, val: f"EMAIL;TYPE=INTERNET:{val}",
    "url":     lambda label, val: f"URL;TYPE={label or 'Website'}:{val}",
    "note":    lambda label, val: f"NOTE:{val}",
    "title":   lambda label, val: f"TITLE:{val}",
    "org":     lambda label, val: f"ORG:{val}",
}


def build_vcard(first, last, fields):
    """Build a vCard string.

    fields: list of {"label": str, "value": str, "type": "phone|email|url|note"}
    """
    lines = [
        "BEGIN:VCARD",
        "VERSION:3.0",
        f"N:{last};{first};;;",
        f"FN:{first} {last}".strip(),
    ]
    for f in fields:
        val = (f.get("value") or "").strip()
        if not val:
            continue
        ftype = f.get("type", "url")
        label = (f.get("label") or "").strip()
        renderer = VCARD_RENDERERS.get(ftype, VCARD_RENDERERS["url"])
        lines.append(renderer(label, val))
    lines.append("END:VCARD")
    return "\n".join(lines)


def make_qr_with_logo(data, logo_bytes=None, qr_size=1000, logo_ratio=0.22):
    """Return a PIL Image of the QR with an optional centered logo."""
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=20,
        border=3,
    )
    qr.add_data(data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGBA")
    qr_img = qr_img.resize((qr_size, qr_size), Image.NEAREST)

    if logo_bytes:
        logo = Image.open(io.BytesIO(logo_bytes)).convert("RGBA")
        logo_size = int(qr_size * logo_ratio)
        logo = logo.resize((logo_size, logo_size), Image.LANCZOS)

        pad = max(12, logo_size // 12)
        bg_size = logo_size + pad * 2
        bg = Image.new("RGBA", (bg_size, bg_size), (255, 255, 255, 255))
        bg.paste(logo, (pad, pad), logo)

        pos = ((qr_size - bg_size) // 2, (qr_size - bg_size) // 2)
        qr_img.paste(bg, pos, bg)

    return qr_img.convert("RGB")


def qr_png_bytes(data, logo_bytes=None, qr_size=1000):
    img = make_qr_with_logo(data, logo_bytes=logo_bytes, qr_size=qr_size)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def read_excel_headers(file_stream):
    wb = load_workbook(file_stream, data_only=True, read_only=True)
    ws = wb.active
    headers = []
    for cell in next(ws.iter_rows(min_row=1, max_row=1)):
        headers.append(str(cell.value).strip() if cell.value else "")
    return [h for h in headers if h]


def read_excel_rows(file_stream):
    """Return (headers, rows) where rows is list of dicts {header: value}."""
    wb = load_workbook(file_stream, data_only=True)
    ws = wb.active
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    rows = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        if not any(raw):
            continue
        row = {headers[i]: ("" if v is None else str(v)) for i, v in enumerate(raw) if i < len(headers)}
        rows.append(row)
    return headers, rows


def build_excel_with_qr(rows, mapping, logo_bytes=None):
    """Build a styled .xlsx with a QR column.

    rows: list of dicts (from read_excel_rows)
    mapping: dict describing how to build the vCard for each row, shape:
        {
            "first_col": "<column name for first>",
            "last_col":  "<column name for last>",
            "fields": [
                {"col": "<column name>", "label": "<label>", "type": "phone|email|url|note"},
                ...
            ],
        }

    Returns BytesIO of the new workbook.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Contacts"

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", start_color="2D6A4F")
    cell_font = Font(name="Arial", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Build header row: first, last, each mapped field, then QR
    headers = ["First Name", "Last Name"]
    headers += [f.get("label") or f["col"] for f in mapping["fields"]]
    headers.append("QR Code")

    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border
        ws.column_dimensions[get_column_letter(col)].width = max(14, min(40, len(h) + 6))
    ws.row_dimensions[1].height = 30

    first_col = mapping["first_col"]
    last_col = mapping["last_col"]
    qr_col_idx = len(headers)

    for i, row in enumerate(rows, start=2):
        first = row.get(first_col, "")
        last = row.get(last_col, "")

        ws.cell(row=i, column=1, value=first).alignment = center
        ws.cell(row=i, column=2, value=last).alignment = center
        ws.cell(row=i, column=1).border = border
        ws.cell(row=i, column=2).border = border
        ws.cell(row=i, column=1).font = cell_font
        ws.cell(row=i, column=2).font = cell_font

        vcard_fields = []
        for col_idx, f in enumerate(mapping["fields"], start=3):
            val = row.get(f["col"], "")
            cell = ws.cell(row=i, column=col_idx, value=val)
            cell.alignment = center
            cell.border = border
            cell.font = cell_font
            vcard_fields.append({"label": f.get("label") or f["col"], "value": val, "type": f["type"]})

        vcard = build_vcard(first, last, vcard_fields)
        img = make_qr_with_logo(vcard, logo_bytes=logo_bytes, qr_size=700)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)

        xl_img = XLImage(buf)
        xl_img.width = 140
        xl_img.height = 140
        ws.add_image(xl_img, f"{get_column_letter(qr_col_idx)}{i}")
        ws.cell(row=i, column=qr_col_idx).border = border
        ws.row_dimensions[i].height = 110

    ws.freeze_panes = "A2"
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out
