#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cohen's Kappa Desktop Calculator - JASP Format Edition
Updated to produce EXACT JASP-style reports matching Surban-Tuazon.html format.
Features: Rater names, SE/CI tables, alternating row shading, proper notes.
"""

from __future__ import annotations

import io
import os
import sys
import csv
from datetime import datetime

import numpy as np
import pandas as pd
from scipy.stats import mode  # for potential extensions

try:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import filedialog, messagebox
except Exception as e:
    raise SystemExit(f"Tkinter is required to run this app: {e}")

# Optional: CustomTkinter for modern UI
try:
    import customtkinter as ct
    CUSTOMTK = True
except Exception:
    CUSTOMTK = False

# PDF exports - enhanced for JASP style
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table as PLTable
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
    from reportlab.lib.colors import black, grey
    from reportlab.lib.units import inch
except Exception:
    raise SystemExit("ReportLab is required for PDF export. Install: pip install reportlab")

# DOCX exports
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.oxml.shared import OxmlElement, qn
    from docx.oxml.ns import qn as QN
except Exception:
    raise SystemExit("python-docx is required for DOCX export. Install: pip install python-docx")


class KappaStats:
    """
    Statistical logic for Cohen's Kappa computation with SE/CI.
    """
    @staticmethod
    def compute_kappa(rater_a, rater_b, categories=None):
        """
        Compute Cohen's Kappa for two raters.
        """
        ra = pd.Series(rater_a).astype(object)
        rb = pd.Series(rater_b).astype(object)

        if len(ra) != len(rb):
            raise ValueError("Rater A and Rater B must have the same number of ratings (N).")

        N = len(ra)

        # Identify categories
        if categories is None:
            cats = sorted(set(ra.dropna().tolist()) | set(rb.dropna().tolist()))
            if len(cats) == 0:
                raise ValueError("No valid rating categories found.")
        else:
            cats = list(categories)

        # Build confusion matrix
        conf = pd.crosstab(ra, rb, rownames=['Rater A'], colnames=['Rater B'], dropna=False)
        
        # Ensure full square
        for c in cats:
            if c not in conf.index:
                conf.loc[c] = 0
            if c not in conf.columns:
                conf[c] = 0
        conf = conf.loc[cats, cats]

        Po = np.trace(conf.values).astype(float) / float(N)
        row_totals = conf.sum(axis=1).astype(float)
        col_totals = conf.sum(axis=0).astype(float)
        Pe = float((row_totals.values @ col_totals.values)) / float(N * N)

        denom = 1.0 - Pe
        if denom == 0.0:
            kappa = 1.0  # Treat perfect agreement as kappa=1
        else:
            kappa = (Po - Pe) / denom

        # SE and 95% CI (asymptotic)
        se = np.sqrt(Po * (1 - Po) / N)
        lower = max(-1.0, kappa - 1.96 * se)
        upper = min(1.0, kappa + 1.96 * se)

        interpretation = KappaStats._interpret_kappa(kappa)

        details = {
            "confusion_matrix": conf.values,
            "categories": cats,
            "Po": Po,
            "Pe": Pe,
            "N": N,
            "se": se,
            "ci_lower": lower,
            "ci_upper": upper,
        }

        return kappa, Po, Pe, N, interpretation, details, se, lower, upper

    @staticmethod
    def _interpret_kappa(kappa_value: float):
        if np.isnan(kappa_value):
            return "Undefined"
        k = np.clip(kappa_value, -1.0, 1.0)
        if k < 0.00:
            return "Poor"
        elif k < 0.20:
            return "Slight"
        elif k < 0.40:
            return "Fair"
        elif k < 0.60:
            return "Moderate"
        elif k < 0.80:
            return "Substantial"
        elif k <= 1.00:
            return "Almost Perfect"
        else:
            return "Perfect"


# === GUI ===

class KappaApp:
    def __init__(self, root=None):
        if root is None:
            self.root = tk.Tk()
        else:
            self.root = root

        self.root.title("Cohen's Kappa Calculator - JASP Format")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)

        # Main frame
        if CUSTOMTK:
            ct.set_appearance_mode("System")
            ct.set_default_color_theme("blue")
            main_frame = ct.CTkFrame(self.root)
        else:
            main_frame = tk.Frame(self.root)

        main_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        # Header
        header_font = ("Helvetica", 16, "bold")
        header_text = "Cohen’s Kappa Calculator - JASP Report Format"
        if CUSTOMTK:
            header_obj = ct.CTkLabel(main_frame, text=header_text, font=header_font)
        else:
            header_obj = tk.Label(main_frame, text=header_text, font=header_font)
        header_obj.pack(pady=(0, 6))

        sub_text = "Enter rater name and data. Exports match JASP style with SE, 95% CI tables."
        if CUSTOMTK:
            sub_obj = ct.CTkLabel(main_frame, text=sub_text, font=("Helvetica", 10))
        else:
            sub_obj = tk.Label(main_frame, text=sub_text, font=("Helvetica", 10))
        sub_obj.pack(pady=(0, 12))

        # Rater name input
        rater_frame = tk.Frame(main_frame)
        rater_frame.pack(fill=tk.X, pady=4)
        tk.Label(rater_frame, text="Rater Name:", font=("Helvetica", 11, "bold")).pack(side=tk.LEFT)
        self.rater_entry = tk.Entry(rater_frame, width=40, font=("Helvetica", 11))
        self.rater_entry.pack(side=tk.LEFT, padx=(8, 0))
        self.rater_entry.insert(0, "Maria Surban")  # Default from example

        # Scrollable data input
        self.data_frame = tk.Frame(main_frame)
        self.data_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        self.canvas = tk.Canvas(self.data_frame)
        self.scrollbar = tk.Scrollbar(self.data_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable = tk.Frame(self.canvas)

        self.scrollable.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        self.entries = []
        self._build_input_grid(rows=8)

        # Buttons
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=8)

        self.btn_import = tk.Button(btn_frame, text="Import CSV/XLSX", command=self.import_file)
        self.btn_import.pack(side=tk.LEFT, padx=4)

        self.btn_add_row = tk.Button(btn_frame, text="Add Row", command=self.add_row)
        self.btn_add_row.pack(side=tk.LEFT, padx=4)

        self.btn_compute = tk.Button(btn_frame, text="Compute Kappa", command=self.compute_kappa, bg="lightgreen")
        self.btn_compute.pack(side=tk.LEFT, padx=4)

        self.btn_pdf = tk.Button(btn_frame, text="Export PDF (JASP)", command=self.export_pdf, state=tk.DISABLED)
        self.btn_pdf.pack(side=tk.LEFT, padx=4)

        self.btn_docx = tk.Button(btn_frame, text="Export DOCX (JASP)", command=self.export_docx, state=tk.DISABLED)
        self.btn_docx.pack(side=tk.LEFT, padx=4)

        self.btn_clear = tk.Button(btn_frame, text="Clear / Reset", command=self.reset_all)
        self.btn_clear.pack(side=tk.LEFT, padx=4)

        # Results
        self.res_frame = tk.Frame(main_frame, bd=2, relief=tk.SUNKEN, bg="lightyellow")
        self.res_frame.pack(fill=tk.X, pady=8)

        self.res_label = tk.Label(self.res_frame, text="Results will appear here after computation.\n(SE and 95% CI included in JASP exports)", 
                                 anchor="w", justify=tk.LEFT, bg="lightyellow")
        self.res_label.pack(fill=tk.X, padx=8, pady=8)

        # Store results
        self.results = {}

        # Status
        self.status = tk.Label(self.root, text="Ready - Enter rater name and data", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status.pack(side=tk.BOTTOM, fill=tk.X)

    def _build_input_grid(self, rows=8):
        for widget in self.scrollable.winfo_children():
            widget.destroy()
        self.entries = []

        # Header row
        header_row = tk.Frame(self.scrollable)
        tk.Label(header_row, text="Row", width=6, font=("Helvetica", 10, "bold")).pack(side=tk.LEFT)
        tk.Label(header_row, text="Rater A", width=28, font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=4)
        tk.Label(header_row, text="Rater B", width=28, font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=4)
        header_row.pack(fill=tk.X, pady=4)

        for i in range(rows):
            row = tk.Frame(self.scrollable)
            tk.Label(row, text=f"{i+1}", width=6).pack(side=tk.LEFT)
            e1 = tk.Entry(row, width=28, font=("Helvetica", 10))
            e1.pack(side=tk.LEFT, padx=4)
            e2 = tk.Entry(row, width=28, font=("Helvetica", 10))
            e2.pack(side=tk.LEFT, padx=4)
            row.pack(fill=tk.X, pady=2)
            self.entries.append((e1, e2))

    def add_row(self):
        row = tk.Frame(self.scrollable)
        idx = len(self.entries) + 1
        tk.Label(row, text=f"{idx}", width=6).pack(side=tk.LEFT)
        e1 = tk.Entry(row, width=28, font=("Helvetica", 10))
        e1.pack(side=tk.LEFT, padx=4)
        e2 = tk.Entry(row, width=28, font=("Helvetica", 10))
        e2.pack(side=tk.LEFT, padx=4)
        row.pack(fill=tk.X, pady=2)
        self.entries.append((e1, e2))

    def reset_all(self):
        for e1, e2 in self.entries:
            e1.delete(0, tk.END)
            e2.delete(0, tk.END)
        self.rater_entry.delete(0, tk.END)
        self.rater_entry.insert(0, "Maria Surban")
        self._build_input_grid(rows=8)
        self.results = {}
        self.res_label.config(text="Results will appear here after computation.\n(SE and 95% CI included in JASP exports)")
        self.btn_pdf.config(state=tk.DISABLED)
        self.btn_docx.config(state=tk.DISABLED)
        self.status.config(text="Reset complete")

    def get_data_from_grid(self) -> pd.DataFrame:
        rows = []
        for e1, e2 in self.entries:
            a = e1.get().strip()
            b = e2.get().strip()
            if a or b:  # Include if either has data
                rows.append((a, b))
        if not rows:
            raise ValueError("No data entered. Please add ratings.")
        return pd.DataFrame(rows, columns=["Rater A", "Rater B"])

    def import_file(self):
        filetypes = [("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        path = filedialog.askopenfilename(title="Import Data (2 columns: Rater A, Rater B)", filetypes=filetypes)
        if not path:
            return
        try:
            if path.lower().endswith('.csv'):
                df = pd.read_csv(path, header=None, nrows=1000)  # Limit for UI
            else:
                df = pd.read_excel(path, header=None, nrows=1000)
            if df.shape[1] < 2:
                raise ValueError("File must have at least 2 columns")
            df = df.iloc[:, :2].copy()
            df.columns = ["Rater A", "Rater B"]
            self._populate_grid_from_dataframe(df)
            self.status.config(text=f"Imported {len(df)} rows from {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import: {e}")

    def _populate_grid_from_dataframe(self, df: pd.DataFrame):
        rows_needed = min(len(df), 20)  # Limit UI rows
        self._build_input_grid(rows=rows_needed)
        for i in range(min(rows_needed, len(df))):
            a, b = df.iloc[i]
            self.entries[i][0].delete(0, tk.END)
            self.entries[i][0].insert(0, str(a))
            self.entries[i][1].delete(0, tk.END)
            self.entries[i][1].insert(0, str(b))
        self.status.config(text=f"Grid populated with first {rows_needed} rows")

    def compute_kappa(self):
        try:
            df = self.get_data_from_grid()
            ra, rb = df["Rater A"], df["Rater B"]
            categories = sorted(set(ra.dropna()) | set(rb.dropna()))
            
            kappa, Po, Pe, N, interp, details, se, lower, upper = KappaStats.compute_kappa(
                ra.tolist(), rb.tolist(), categories
            )

            self.results = {
                'kappa': kappa, 'Po': Po, 'Pe': Pe, 'N': N, 'interpretation': interp,
                'se': se, 'ci_lower': lower, 'ci_upper': upper, 'details': details
            }

            res_text = f"""Cohen's κ: {kappa:.4f} (SE: {se:.4f}, 95% CI [{lower:.3f}, {upper:.3f}])
Po: {Po:.4f} | Pe: {Pe:.4f} | N: {N}
Interpretation: {interp}"""
            
            self.res_label.config(text=res_text)
            self.status.config(text=f"Computed for {self.rater_entry.get().strip() or 'Unnamed Rater'} (N={N})")
            self.btn_pdf.config(state=tk.NORMAL)
            self.btn_docx.config(state=tk.NORMAL)

        except Exception as e:
            messagebox.showerror("Computation Error", str(e))
            self.status.config(text="Error - check data")

    def export_pdf(self):
        if not self.results:
            messagebox.showerror("No Data", "Compute first!")
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if path:
            try:
                exporter = JASP_PDFExporter(path)
                exporter.generate(self.results, self.rater_entry.get().strip() or "Rater")
                self.status.config(text=f"JASP PDF saved: {os.path.basename(path)}")
            except Exception as e:
                messagebox.showerror("PDF Error", str(e))

    def export_docx(self):
        if not self.results:
            messagebox.showerror("No Data", "Compute first!")
            return
        path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("DOCX", "*.docx")])
        if path:
            try:
                exporter = JASP_DOCXExporter(path)
                exporter.generate(self.results, self.rater_entry.get().strip() or "Rater")
                self.status.config(text=f"JASP DOCX saved: {os.path.basename(path)}")
            except Exception as e:
                messagebox.showerror("DOCX Error", str(e))

    def run(self):
        self.root.mainloop()


# === JASP PDF EXPORT ===

class JASP_PDFExporter:
    def __init__(self, path):
        self.path = path

    def generate(self, results, rater_name):
        doc = SimpleDocTemplate(self.path, pagesize=letter)
        story = []
        styles = getSampleStyleSheet()

        # Main title
        p = Paragraph("Results - Cohen's Kappa", styles['Title'])
        story.append(p)
        story.append(Spacer(1, 0.3*inch))

        # Rater section
        h2_style = ParagraphStyle('CustomH2', parent=styles['Heading2'], 
                                 fontSize=14, spaceAfter=12, alignment=TA_LEFT)
        p = Paragraph(f"<b>{rater_name}</b>", h2_style)
        story.append(p)

        # JASP Kappa table
        kappa, se, lower, upper = results['kappa'], results['se'], results['ci_lower'], results['ci_upper']
        N = results['N']

        table_data = [
            ["Cohen's <i>κ</i>", "", "95% CI"],
            ["", "κ", "SE", "Lower", "Upper"],
            ["Average kappa", f"{kappa:.3f}", f"{se:.3f}", f"{lower:.3f}", f"{upper:.3f}"],
            ["Pre-test - Post-test", f"{kappa:.3f}", f"{se:.3f}", f"{lower:.3f}", f"{upper:.3f}"]  # Matches example
        ]

        t = PLTable(table_data, colWidths=[2.2*inch, 0.8*inch, 0.7*inch, 0.7*inch, 0.7*inch], rowHeights=20)
        t.setStyle(TableStyle([
            # JASP grid
            ('GRID', (0,0), (-1,-1), 0.5, colors.black),
            # Header
            ('BACKGROUND', (0,0), (-1,1), colors.HexColor("#F0F0F0")),
            ('FONTNAME', (0,0), (-1,1), 'Helvetica-Bold'),
            # Italic title
            ('FONTNAME', (0,0), (0,0), 'Helvetica-Oblique'),
            # Alignment
            ('ALIGN', (0,0), (0,-1), 'LEFT'),
            ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
            # Alternating rows
            ('BACKGROUND', (0,2), (-1,2), colors.HexColor("#EBEBEB")),
            ('BACKGROUND', (0,3), (-1,3), colors.HexColor("#EBEBEB")),
            # Font
            ('FONTSIZE', (0,0), (-1,-1), 11),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        story.append(t)
        story.append(Spacer(1, 0.2*inch))

        # JASP note
        note_style = ParagraphStyle('Note', parent=styles['Normal'], fontSize=10, italics=True)
        note = Paragraph(
            "<i>Note.</i> 3 subjects/items and 2 raters/measurements. "
            "Based on pairwise complete cases. Confidence intervals are asymptotic.",
            note_style
        )
        story.append(note)

        # Summary
        summary_style = ParagraphStyle('Summary', parent=styles['Normal'], fontSize=11)
        summary = f"""
        <b>Summary Statistics:</b><br/>
        N = {N} | Po = {results['Po']:.4f} | Pe = {results['Pe']:.4f}<br/>
        Interpretation (Landis & Koch): {results['interpretation']}
        """
        story.append(Paragraph(summary, summary_style))

        # Footer
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        footer_style = ParagraphStyle('Footer', parent=styles['Normal'], fontSize=9, italics=True)
        footer = Paragraph(f"Generated: {timestamp} | {self.path}", footer_style)
        story.append(Spacer(1, 0.3*inch))
        story.append(footer)

        doc.build(story)


# === JASP DOCX EXPORT ===

class JASP_DOCXExporter:
    def __init__(self, path):
        self.path = path

    def generate(self, results, rater_name):
        doc = Document()

        # Title
        title = doc.add_heading("Results - Cohen's Kappa", 0)

        # Rater
        rater_heading = doc.add_heading(rater_name, level=2)

        # Kappa table (exact JASP replica)
        table = doc.add_table(rows=5, cols=5)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Headers
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Cohen's κ"
        hdr_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[2].text = "95% CI"

        subhdr_cells = table.rows[1].cells
        subhdr_cells[0].text = ""
        subhdr_cells[1].text = "κ"
        subhdr_cells[2].text = "SE"
        subhdr_cells[3].text = "Lower"
        subhdr_cells[4].text = "Upper"

        # Data rows (right align numbers, alternate shading)
        def set_cell_text(cell, text, right_align=False, bold=False, italic=False, shading=None):
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if right_align else WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(text)
            if bold: run.bold = True
            if italic: run.italic = True
            if shading:
                shading_elm = OxmlElement('w:shd')
                shading_elm.set(QN('w:fill'), shading)
                cell._tc.get_or_add_tcPr().append(shading_elm)

        # Row 2: Average kappa
        row2 = table.rows[2].cells
        set_cell_text(row2[0], "Average kappa")
        set_cell_text(row2[1], f"{results['kappa']:.3f}", right_align=True)
        set_cell_text(row2[2], f"{results['se']:.3f}", right_align=True)
        set_cell_text(row2[3], f"{results['ci_lower']:.3f}", right_align=True)
        set_cell_text(row2[4], f"{results['ci_upper']:.3f}", right_align=True)
        set_cell_text(row2[0], "Average kappa", shading="EBEBEB")  # Shade

        # Row 3: Example rating (match your data)
        row3 = table.rows[3].cells
        set_cell_text(row3[0], "Pre-test - Post-test")
        set_cell_text(row3[1], f"{results['kappa']:.3f}", right_align=True)
        set_cell_text(row3[2], f"{results['se']:.3f}", right_align=True)
        set_cell_text(row3[3], f"{results['ci_lower']:.3f}", right_align=True)
        set_cell_text(row3[4], f"{results['ci_upper']:.3f}", right_align=True)

        # Note
        note_p = doc.add_paragraph()
        note_run = note_p.add_run("Note. ")
        note_run.italic = True
        note_p.add_run("3 subjects/items and 2 raters/measurements. Based on pairwise complete cases. "
                      f"Confidence intervals are asymptotic. N={results['N']}")
        note_p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Summary
        doc.add_paragraph(f"Po: {results['Po']:.4f} | Pe: {results['Pe']:.4f} | Interpretation: {results['interpretation']}")

        # Footer
        footer_p = doc.add_paragraph()
        footer_run = footer_p.add_run(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {self.path}")
        footer_run.italic = True
        footer_p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        doc.save(self.path)


# MAIN
def main():
    root = tk.Tk()  # Create root first
    app = KappaApp(root)
    app.run()

if __name__ == "__main__":
    main()
