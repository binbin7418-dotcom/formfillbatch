# FormFillBatch

Fill a fillable PDF form once per spreadsheet row: one flattened PDF
per row, with every output re-opened and read back for verification.

## Run (dev)

```
python gui.py                              # desktop app (pypdf + reportlab)
python form_fill.py template.pdf data.csv out/
<venv python> test/acceptance_test.py
```

## How it works

`form_fill.py` reads the template's AcroForm fields (XFA is detected
and rejected clearly), maps CSV/XLSX columns to field names, fills
each row with pypdf, then flattens *properly*: values are drawn into
the page content via a reportlab overlay merged per page (widget
rects come from the template), after which widgets and the AcroForm
dictionary are removed. Verification re-opens each output and checks
every text value against the row it came from; the report lists
unmatched columns, unfilled fields, and per-row verdicts.

## Scope

AcroForm text fields + checkboxes. Checkbox states use the form's
own export values and are not text-verifiable after flattening (the
report says so). Digital signatures are not preserved by flattening.
