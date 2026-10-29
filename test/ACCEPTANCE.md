# FormFillBatch — Acceptance

Seeded test (pagesmith venv python: reportlab + pypdf + openpyxl):
`test/acceptance_test.py` — full output: `test/acceptance_results.txt`

**Result: 12/12 checks passed.**

The test builds a real fillable template with reportlab's AcroForm
(text fields FullName + City, checkbox Agree) and a 3-row CSV with
one deliberately unmatched column (Nickname).

Verified:
- Template field detection exact (3 fields)
- Batch verdict PASS, 0 errors; 3 outputs written
- Unmatched column reported (Nickname); no unfilled fields
- Filenames from the name column (001_Alice_Nguyen.pdf …)
- Every row's values readable in its output's extracted text
- Every output fully flattened (no form fields remain)
- GUI + engine compile clean (`py_compile`)

Note: the first test run caught a real defect — pypdf field values
were not surviving flattening (they lived only in widget appearance
streams). Fixed by drawing values into the page content via a
per-page reportlab overlay before removing the widgets; the re-run
is the 12/12 above.
