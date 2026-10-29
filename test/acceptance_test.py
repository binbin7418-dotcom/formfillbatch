#!/usr/bin/env python3
"""FormFillBatch seeded acceptance test (run with the pagesmith venv
python, which provides reportlab + pypdf + openpyxl).

Builds a real fillable template with reportlab's AcroForm (text
fields FullName + City, checkbox Agree), seeds a 3-row CSV (plus one
column that matches no field, to prove it is reported, not silently
dropped), then asserts: 3 flattened PDFs, values readable in the
extracted text of each output, form fields gone after flattening,
filenames from the name column, mapping report correct, verdict PASS.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from pypdf import PdfReader
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas

from form_fill import fill_batch

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
CHECKS = []
ROWS = [("Alice Nguyen", "Austin", "yes"),
        ("Brian Osei", "Denver", "no"),
        ("Carla Mendes", "Miami", "yes")]


def check(name, ok, detail=""):
    CHECKS.append((name, bool(ok), detail))
    print(("PASS " if ok else "FAIL ") + name + (f" — {detail}" if detail else ""))


os.makedirs(OUT, exist_ok=True)
tpl = os.path.join(OUT, "template.pdf")
c = canvas.Canvas(tpl, pagesize=LETTER)
form = c.acroForm
c.setFont("Helvetica", 11)
c.drawString(72, 720, "Vendor Intake Form (seed)")
c.drawString(72, 690, "Full name:")
form.textfield(name="FullName", x=160, y=682, width=260, height=18)
c.drawString(72, 655, "City:")
form.textfield(name="City", x=160, y=647, width=260, height=18)
c.drawString(72, 620, "Agrees to terms:")
form.checkbox(name="Agree", x=160, y=614, size=14)
c.save()

check("template really has 3 fields",
      set((PdfReader(tpl).get_fields() or {}).keys()) == {"FullName", "City", "Agree"})

csv_path = os.path.join(OUT, "people.csv")
with open(csv_path, "w") as fh:
    fh.write("FullName,City,Agree,Nickname\n")
    for n, city, a in ROWS:
        fh.write(f"{n},{city},{a},{n.split()[0]}\n")

s = fill_batch(tpl, csv_path, os.path.join(OUT, "filled"), name_column="FullName")
check("batch verdict PASS", s["verdict"], f"errors={s['errors']}")
check("3 outputs written", len(s["results"]) == 3)
check("unmatched column reported", s["unmatched_columns"] == ["Nickname"],
      f"{s['unmatched_columns']}")
check("no unfilled fields", s["unfilled_fields"] == [])
names = [r["file"] for r in s["results"]]
check("filenames from name column",
      any("Alice_Nguyen" in n for n in names) and any("Carla_Mendes" in n for n in names),
      f"{names}")
for r, (n, city, _a) in zip(s["results"], ROWS):
    rd = PdfReader(os.path.join(OUT, "filled", r["file"]))
    text = "\n".join((p.extract_text() or "") for p in rd.pages)
    check(f"row {r['row']} values readable", n in text and city in text)
    check(f"row {r['row']} flattened (no fields left)",
          not (rd.get_fields() or {}))

passed = sum(1 for _, ok, _ in CHECKS if ok)
print(f"\n{passed}/{len(CHECKS)} checks passed")
sys.exit(0 if passed == len(CHECKS) else 1)
