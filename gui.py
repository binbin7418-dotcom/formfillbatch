#!/usr/bin/env python3
"""FormFillBatch desktop shell (tkinter).
Pick a fillable PDF + a spreadsheet -> one filled, flattened PDF per row."""
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from form_fill import fill_batch, list_fields


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FormFillBatch — Fill PDF Forms from a Spreadsheet")
        self.geometry("700x600")
        self.template = tk.StringVar()
        self.data = tk.StringVar()
        self.namecol = tk.StringVar()
        tk.Label(self, text="Fillable PDF form (template):").pack(anchor="w", padx=12, pady=(12, 2))
        row = tk.Frame(self)
        row.pack(fill="x", padx=12)
        tk.Entry(row, textvariable=self.template).pack(side="left", fill="x", expand=True)
        tk.Button(row, text="Browse…", command=self.pick_template).pack(side="left", padx=6)
        tk.Label(self, text="Data file (CSV or Excel) — column headers must match field names:").pack(
            anchor="w", padx=12, pady=(10, 2))
        row2 = tk.Frame(self)
        row2.pack(fill="x", padx=12)
        tk.Entry(row2, textvariable=self.data).pack(side="left", fill="x", expand=True)
        tk.Button(row2, text="Browse…", command=self.pick_data).pack(side="left", padx=6)
        tk.Label(self, text="Output files are named from this column:").pack(anchor="w", padx=12, pady=(10, 2))
        self.name_combo = ttk.Combobox(self, textvariable=self.namecol, state="readonly")
        self.name_combo.pack(fill="x", padx=12)
        tk.Button(self, text="List the form's fields", command=self.show_fields).pack(pady=(10, 0))
        self.btn = tk.Button(self, text="Fill All Rows & Verify", command=self.run, height=2)
        self.btn.pack(pady=10)
        self.log = scrolledtext.ScrolledText(self, wrap="word")
        self.log.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.write("FormFillBatch fills one copy of your PDF form per spreadsheet row,\n"
                   "flattens each copy, then re-opens every output and reads the\n"
                   "values back to prove they landed. Everything runs locally.\n")

    def write(self, s):
        self.log.insert("end", s)
        self.log.see("end")

    def pick_template(self):
        p = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if p:
            self.template.set(p)

    def pick_data(self):
        p = filedialog.askopenfilename(
            filetypes=[("Data files", "*.csv *.xlsx *.xlsm"), ("All files", "*.*")])
        if p:
            self.data.set(p)
            try:
                from form_fill import read_rows
                rows = read_rows(p)
                if rows:
                    self.name_combo["values"] = list(rows[0].keys())
                    self.namecol.set(list(rows[0].keys())[0])
            except Exception as e:
                self.write(f"\nCould not read data headers: {e}\n")

    def show_fields(self):
        if not self.template.get():
            messagebox.showerror("FormFillBatch", "Pick a template PDF first.")
            return
        try:
            _r, fields = list_fields(self.template.get())
            names = ", ".join(sorted(fields)) or "(no fields found)"
            self.write(f"\nForm fields ({len(fields)}): {names}\n")
        except Exception as e:
            self.write(f"\nError reading fields: {e}\n")

    def run(self):
        if not self.template.get() or not self.data.get():
            messagebox.showerror("FormFillBatch", "Pick both a template PDF and a data file.")
            return
        self.btn.config(state="disabled")
        threading.Thread(target=self._work, daemon=True).start()

    def _work(self):
        try:
            outdir = os.path.join(os.path.dirname(os.path.abspath(self.data.get())),
                                  "formfill_output")
            s = fill_batch(self.template.get(), self.data.get(), outdir,
                           name_column=self.namecol.get() or None)
            self.write(f"\nDone. {len(s['results'])} PDF(s) written, "
                       f"{len(s['mapped'])} columns mapped, errors: {len(s['errors'])}.\n"
                       f"Read-back verdict: {'PASS' if s['verdict'] else 'CHECK REPORT'}\n"
                       f"Output folder: {outdir}\nReport: {s['report_path']}\n")
        except Exception as e:
            self.write(f"\nError: {e}\n")
        finally:
            self.btn.config(state="normal")


if __name__ == "__main__":
    App().mainloop()
