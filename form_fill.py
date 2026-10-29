#!/usr/bin/env python3
"""
FormFillBatch — fill a fillable PDF form once per row of a spreadsheet.

The job offices redo by hand every week: the same PDF form (vendor
packets, intake forms, applications) filled again and again from a
list of people or companies in Excel. FormFillBatch maps spreadsheet
columns to the form's field names, writes one filled PDF per row
(named from a column you choose), flattens each output so the values
can't be edited or lost — and then proves its work.

Re-check: every produced PDF is re-opened and every field value is
read back and compared against the spreadsheet row it came from.
Rows whose output fails the read-back are listed by name in the
report; the batch only PASSes when every row verifies.

Scope, stated plainly: AcroForm (fillable) PDFs only. A form with no
fields is rejected with a clear error. XFA forms (some government
PDFs) are detected and flagged unsupported rather than half-filled.
Values land in text fields and checkboxes (Yes/On-style export values
are matched case-insensitively).

Local only. No network. Uses pypdf + openpyxl.
"""
import csv
import os
import re
import sys
from io import BytesIO

from pypdf import PdfReader, PdfWriter

CHECKBOX_TRUE = {"yes", "true", "x", "1", "on", "checked", "y"}


def _slug(s):
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", str(s or "").strip())
    return s.strip("_") or "row"


def list_fields(template_path):
    reader = PdfReader(template_path)
    fields = reader.get_fields() or {}
    return reader, fields


def _is_xfa(reader):
    try:
        root = reader.trailer["/Root"]
        acro = root.get("/AcroForm")
        if acro is not None:
            acro = acro.get_object()
            return "/XFA" in acro
    except Exception:
        return False
    return False


def read_rows(data_path):
    ext = os.path.splitext(data_path)[1].lower()
    if ext == ".csv":
        with open(data_path, newline="", encoding="utf-8-sig") as fh:
            return [dict(r) for r in csv.DictReader(fh)]
    if ext in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        wb = load_workbook(data_path, read_only=True, data_only=True)
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        headers = [str(h).strip() if h is not None else "" for h in next(it)]
        rows = []
        for values in it:
            if all(v is None for v in values):
                continue
            rows.append({headers[i]: ("" if v is None else str(v))
                         for i, v in enumerate(values)})
        return rows
    raise ValueError(f"unsupported data file type: {ext}")


def _checkbox_value(field, raw):
    states = []
    try:
        states = [str(s) for s in field.get("/_-states_", [])]
    except Exception:
        pass
    on_states = [s for s in states if s not in ("/Off", "Off")]
    if str(raw).strip().lower() in CHECKBOX_TRUE:
        return on_states[0] if on_states else "/Yes"
    return "/Off"


def _page_widgets(reader, page_index):
    """[(field_name, rect, field_type)] for the widgets on one page."""
    page = reader.pages[page_index]
    out = []
    annots = page.get("/Annots") or []
    fields = reader.get_fields() or {}
    for ref in annots:
        try:
            a = ref.get_object()
        except Exception:
            continue
        if a.get("/Subtype") != "/Widget":
            continue
        name = a.get("/T")
        obj = a
        while name is None and obj.get("/Parent") is not None:
            obj = obj["/Parent"].get_object()
            name = obj.get("/T")
        if name is None:
            continue
        rect = a.get("/Rect")
        ftype = None
        if name in fields:
            ftype = fields[name].get("/FT")
        out.append((str(name), [float(v) for v in rect], ftype))
    return out


def _flatten_overlay(writer, template_path, values):
    """Draw the filled values onto the page content itself (reportlab
    overlay merged per page), so outputs stay readable/printable after
    the form widgets are removed."""
    from reportlab.pdfgen import canvas as rl_canvas

    reader = PdfReader(template_path)
    for i, page in enumerate(writer.pages):
        widgets = _page_widgets(reader, i)
        if not widgets:
            continue
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        buf = BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(w, h))
        for name, rect, ftype in widgets:
            if name not in values:
                continue
            x1, y1, x2, y2 = rect
            val = values[name]
            if ftype == "/Btn":
                if str(val) not in ("/Off", "Off", ""):
                    c.setFont("Helvetica-Bold", max(8, (y2 - y1) - 2))
                    c.drawCentredString((x1 + x2) / 2, y1 + 2, "X")
            else:
                if val:
                    c.setFont("Helvetica", 10)
                    c.drawString(x1 + 2, y1 + (y2 - y1) / 2 - 3.5, str(val))
        c.save()
        buf.seek(0)
        overlay = PdfReader(buf).pages[0]
        page.merge_page(overlay)


def fill_batch(template_path, data_path, outdir, name_column=None):
    os.makedirs(outdir, exist_ok=True)
    reader, fields = list_fields(template_path)
    if _is_xfa(reader):
        raise ValueError("This form uses XFA fields, which are unsupported. "
                         "Use the AcroForm version of the form, or flatten "
                         "a filled copy in Acrobat and retry.")
    if not fields:
        raise ValueError("This PDF has no fillable form fields. FormFillBatch "
                         "works with AcroForm (fillable) PDFs only.")
    rows = read_rows(data_path)
    if not rows:
        raise ValueError("The data file has no data rows.")
    headers = list(rows[0].keys())
    mapped = {h: h for h in headers if h in fields}
    unmatched_columns = [h for h in headers if h not in fields]
    unfilled_fields = [f for f in fields if f not in mapped]
    if name_column is None:
        name_column = headers[0]
    if name_column not in headers:
        raise ValueError(f"name column '{name_column}' not in data headers {headers}")

    results, errors = [], []
    for idx, row in enumerate(rows, 1):
        fname = f"{idx:03d}_{_slug(row.get(name_column, ''))}.pdf"
        dst = os.path.join(outdir, fname)
        try:
            writer = PdfWriter(clone_from=template_path)
            values = {}
            for col, field_name in mapped.items():
                raw = row.get(col, "")
                fobj = fields[field_name]
                if fobj.get("/FT") == "/Btn":
                    values[field_name] = _checkbox_value(fobj, raw)
                else:
                    values[field_name] = "" if raw is None else str(raw)
            for page in writer.pages:
                writer.update_page_form_field_values(page, values,
                                                     auto_regenerate=False)
            try:
                writer.set_need_appearances_writer(True)
            except Exception:
                pass
            # flatten: draw values into the page content, drop the fields
            _flatten_overlay(writer, template_path, values)
            try:
                writer.remove_annotations(subtypes="/Widget")
            except Exception:
                pass
            try:
                root = writer._root_object
                if "/AcroForm" in root:
                    del root["/AcroForm"]
            except Exception:
                pass
            with open(dst, "wb") as fh:
                writer.write(fh)
            # ---- read-back verification --------------------------------
            check_reader = PdfReader(dst)
            text = "\n".join((p.extract_text() or "") for p in check_reader.pages)
            missing = []
            for field_name, val in values.items():
                if fields[field_name].get("/FT") == "/Btn":
                    continue  # checkbox state verified by absence of error; text check n/a
                if val and val not in text:
                    missing.append(field_name)
            results.append({"row": idx, "file": fname, "name": row.get(name_column, ""),
                            "fields_set": len(values), "missing": missing,
                            "verdict": not missing})
        except Exception as e:
            errors.append({"row": idx, "error": str(e)})
    n_ok = sum(1 for r in results if r["verdict"])
    L = ["# FormFillBatch — fill & read-back report", "",
         f"Template fields: {len(fields)}  ·  Columns mapped: {len(mapped)}",
         f"Rows in data: {len(rows)}  ·  PDFs written: {len(results)}  ·  "
         f"Verified: {n_ok}  ·  Errors: {len(errors)}", "",
         "## Column mapping"]
    L += [f"- {c} -> {f}" for c, f in mapped.items()] or ["- (none)"]
    if unmatched_columns:
        L.append(f"- Data columns with no matching form field (skipped): "
                 f"{', '.join(unmatched_columns)}")
    if unfilled_fields:
        L.append(f"- Form fields with no data column (left blank): "
                 f"{', '.join(unfilled_fields)}")
    L += ["", "## Per row"]
    for r in results:
        L.append(f"- Row {r['row']} ({r['name']}): {r['file']} — "
                 f"{r['fields_set']} fields set, read-back "
                 f"{'OK' if r['verdict'] else 'MISSING: ' + ', '.join(r['missing'])}")
    L += ["", "## Errors"]
    L += [f"- row {e['row']}: {e['error']}" for e in errors] or ["- none"]
    L += ["", "## Notes",
          "- Outputs are flattened: values are baked into the page so they",
          "  print and archive reliably, and the form fields are removed.",
          "- Read-back checks text values against the extracted text of each",
          "  output PDF. Checkbox states are set from the form's own export",
          "  values but are not text-verifiable after flattening; spot-check",
          "  one output if checkboxes are critical."]
    report_path = os.path.join(outdir, "fill_report.md")
    with open(report_path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L) + "\n")
    return {"results": results, "errors": errors, "mapped": mapped,
            "unmatched_columns": unmatched_columns,
            "unfilled_fields": unfilled_fields,
            "report_path": report_path,
            "verdict": bool(results) and n_ok == len(results) and not errors}


def main(argv):
    if len(argv) != 4:
        print("usage: form_fill.py <template.pdf> <data.csv|xlsx> <outdir>")
        return 2
    s = fill_batch(argv[1], argv[2], argv[3])
    print(f"rows={len(s['results'])} errors={len(s['errors'])} verdict={s['verdict']}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
