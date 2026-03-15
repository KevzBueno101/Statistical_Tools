"""
T-Test Analysis — Modern Sleek UI
Matches the ANOVA Analyzer / Cohen's Kappa aesthetic.

Implements:
- One-Sample t-test
- Independent Samples t-test
- Paired Samples t-test
- APA-style reporting
- DOCX / PDF export
- Excel / CSV import
"""

import customtkinter as ctk
from tkinter import messagebox, filedialog
import numpy as np
import pandas as pd
from scipy import stats
from datetime import datetime
import re
import os
from app_settings import SettingsManager, SettingsWindow

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False


# ─── Palette (identical to ANOVA Analyzer) ───────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BG_DEEP   = "#0d1117"
BG_CARD   = "#161b22"
BG_PANEL  = "#1c2230"
BG_INPUT  = "#1e2736"
ACCENT    = "#00c9a7"
ACCENT2   = "#4e9eff"
DANGER    = "#ef4444"
WARN      = "#f59e0b"
SUCCESS   = "#22c55e"
PURPLE    = "#a855f7"
TEXT_PRI  = "#e6edf3"
TEXT_SEC  = "#8b949e"
BORDER    = "#30363d"

FONT_HEAD = ("Segoe UI", 26, "bold")
FONT_CARD = ("Segoe UI", 15, "bold")
FONT_BODY = ("Segoe UI", 13)
FONT_MONO = ("Consolas", 12)
FONT_BTN  = ("Segoe UI", 13, "bold")
FONT_TINY = ("Segoe UI", 11)
FONT_LBL  = ("Segoe UI", 12, "bold")


# ─── Shared widget helpers ────────────────────────────────────────────────────

def divider(parent):
    ctk.CTkFrame(parent, height=1, fg_color=BORDER, corner_radius=0).pack(fill="x")


def styled_entry(parent, placeholder="", width=0, height=34):
    kw = dict(placeholder_text=placeholder, fg_color=BG_INPUT,
              border_color=BORDER, text_color=TEXT_PRI,
              placeholder_text_color=TEXT_SEC, border_width=1,
              corner_radius=6, height=height, font=FONT_BODY)
    e = ctk.CTkEntry(parent, **kw)
    if width:
        e.configure(width=width)
    return e


def sidebar_btn(parent, text, fg, hover, text_color=TEXT_PRI,
                font=FONT_BTN, height=36, state="normal"):
    return ctk.CTkButton(parent, text=text, fg_color=fg, hover_color=hover,
                         text_color=text_color, font=font, height=height,
                         corner_radius=8, state=state)


def card(parent, title="", **kw):
    frame = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12,
                         border_width=1, border_color=BORDER, **kw)
    if title:
        ctk.CTkLabel(frame, text=title, font=FONT_CARD,
                     text_color=TEXT_PRI).pack(anchor="w", padx=16, pady=(12, 4))
    return frame


def section_label(parent, text):
    ctk.CTkLabel(parent, text=text, font=("Segoe UI", 11, "bold"),
                 text_color=TEXT_SEC).pack(anchor="w", padx=18, pady=(14, 3))


# ─── Sidebar ──────────────────────────────────────────────────────────────────

class Sidebar(ctk.CTkFrame):
    def __init__(self, master, **kw):
        super().__init__(master, width=230, fg_color=BG_CARD,
                         corner_radius=0, **kw)
        self.pack_propagate(False)
        self._build()

    def _build(self):
        # Logo
        logo = ctk.CTkFrame(self, fg_color=ACCENT, corner_radius=0, height=64)
        logo.pack(fill="x"); logo.pack_propagate(False)
        ctk.CTkLabel(logo, text="  t  ", font=("Segoe UI", 30, "bold"),
                     text_color="#0d1117", fg_color=ACCENT).pack(expand=True)

        ctk.CTkLabel(self, text="t-Test Analysis", font=("Segoe UI", 16, "bold"),
                     text_color=TEXT_PRI, fg_color=BG_CARD).pack(pady=(16, 2))
        ctk.CTkLabel(self, text="APA Format Calculator", font=FONT_TINY,
                     text_color=TEXT_SEC, fg_color=BG_CARD).pack(pady=(0, 10))

        divider(self)

        # Test type selector
        section_label(self, "TEST TYPE")
        self.test_type_var = ctk.StringVar(value="independent")
        self.test_menu = ctk.CTkOptionMenu(
            self, variable=self.test_type_var,
            values=["one-sample", "independent", "paired"],
            fg_color=BG_INPUT, button_color=ACCENT, button_hover_color="#009e82",
            text_color=TEXT_PRI, dropdown_fg_color=BG_PANEL,
            dropdown_text_color=TEXT_PRI, font=FONT_BODY, height=34,
            corner_radius=6
        )
        self.test_menu.pack(fill="x", padx=14, pady=(0, 6))

        # Alpha
        section_label(self, "ALPHA LEVEL")
        self.alpha_entry = styled_entry(self, placeholder="0.05")
        self.alpha_entry.insert(0, "0.05")
        self.alpha_entry.pack(fill="x", padx=14, pady=(0, 6))

        # Researcher name
        section_label(self, "RESEARCHER NAME")
        self.researcher_entry = styled_entry(self, placeholder="e.g. Dr. John Smith")
        self.researcher_entry.pack(fill="x", padx=14, pady=(0, 6))

        divider(self)

        pad = {"fill": "x", "padx": 14, "pady": 4}

        self.import_btn = sidebar_btn(self, "📁  Import Excel / CSV",
                                      fg=ACCENT2, hover="#3b7ddd")
        self.import_btn.pack(**pad)

        self.preview_btn = sidebar_btn(self, "👁  Preview Data",
                                       fg=BG_PANEL, hover=BORDER,
                                       text_color=TEXT_SEC, state="disabled")
        self.preview_btn.pack(**pad)

        divider(self)

        self.run_btn = sidebar_btn(self, "▶   Run t-Test",
                                   fg=ACCENT, hover="#009e82",
                                   text_color="#0d1117",
                                   font=("Segoe UI", 13, "bold"), height=44)
        self.run_btn.pack(**pad)

        self.pdf_btn = sidebar_btn(self, "📄  Export PDF",
                                   fg="#1d4ed8", hover="#1e3a8a", state="disabled")
        self.pdf_btn.pack(**pad)

        self.docx_btn = sidebar_btn(self, "💾  Export DOCX",
                                    fg=PURPLE, hover="#7e22ce", state="disabled")
        self.docx_btn.pack(**pad)

        self.clear_btn = sidebar_btn(self, "🗑  Clear All",
                                     fg=DANGER, hover="#b91c1c")
        self.clear_btn.pack(**pad)

        # Theme toggle
        divider(self)
        self.theme_btn = sidebar_btn(self, "☀️  Light Mode",
                                     fg="#374151", hover="#4b5563",
                                     font=FONT_BODY, height=32)
        self.theme_btn.pack(fill="x", padx=14, pady=8)

        # Settings
        divider(self)
        self.settings_btn = sidebar_btn(self, "⚙   Settings",
                                        fg="#374151", hover="#4b5563",
                                        font=FONT_BODY, height=32)
        self.settings_btn.pack(fill="x", padx=14, pady=8)

        # Status / footer
        self.status_label = ctk.CTkLabel(self, text="", font=FONT_TINY,
                                          text_color=ACCENT, fg_color=BG_CARD,
                                          wraplength=200)
        self.status_label.pack(side="bottom", padx=12, pady=8)


# ─── Main App ─────────────────────────────────────────────────────────────────

class TTestApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("t-Test Analysis")
        self.geometry("1220x800")
        self.minsize(1100, 700)
        self.configure(fg_color=BG_DEEP)

        self.results = None
        self.imported_data = None
        self.dark_mode = True

        self._build_ui()
        self._on_test_change("independent")

    # ── UI Build ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Sidebar
        self.sidebar = Sidebar(self)
        self.sidebar.pack(side="left", fill="y")

        # Wire sidebar
        self.sidebar.test_menu.configure(command=self._on_test_change)
        self.sidebar.import_btn.configure(command=self.import_data)
        self.sidebar.preview_btn.configure(command=self.show_preview)
        self.sidebar.run_btn.configure(command=self.run_analysis)
        self.sidebar.pdf_btn.configure(command=self.export_pdf)
        self.sidebar.docx_btn.configure(command=self.export_docx)
        self.sidebar.clear_btn.configure(command=self.clear_fields)
        self.sidebar.theme_btn.configure(command=self.toggle_theme)
        self.sidebar.settings_btn.configure(command=self.open_settings)

        # Content area
        content = ctk.CTkFrame(self, fg_color=BG_DEEP, corner_radius=0)
        content.pack(side="left", fill="both", expand=True)

        # ── Header bar ────────────────────────────────────────────────────────
        header = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=0, height=64)
        header.pack(fill="x"); header.pack_propagate(False)

        # LEFT side: Title + Subtitle entries
        meta = ctk.CTkFrame(header, fg_color=BG_CARD)
        meta.pack(side="left", padx=16)

        ctk.CTkLabel(meta, text="Title:", font=("Segoe UI", 13, "bold"),
                     text_color=TEXT_PRI).grid(row=0, column=0, padx=(0, 4))
        self.title_entry = styled_entry(meta, placeholder="Report Title", width=200)
        self.title_entry.grid(row=0, column=1, padx=4)

        ctk.CTkLabel(meta, text="Subtitle:", font=("Segoe UI", 13, "bold"),
                     text_color=TEXT_PRI).grid(row=0, column=2, padx=(12, 4))
        self.subtitle_entry = styled_entry(meta, placeholder="e.g. t-Test", width=170)
        self.subtitle_entry.insert(0, "t-Test")
        self.subtitle_entry.grid(row=0, column=3, padx=4)

        # RIGHT side: App name as subtitle label
        ctk.CTkLabel(header, text="t-Test Analysis",
                     font=("Segoe UI", 13), text_color=TEXT_SEC).pack(side="right", padx=24)

        # ── Body: resizable two-column ────────────────────────────────────────
        import tkinter as tk
        from tkinter import ttk

        outer = ctk.CTkFrame(content, fg_color=BG_DEEP)
        outer.pack(fill="both", expand=True, padx=16, pady=16)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Sash", sashthickness=6, sashrelief="flat",
                        background="#30363d")

        pane = ttk.PanedWindow(outer, orient=tk.HORIZONTAL)
        pane.pack(fill="both", expand=True)

        # LEFT
        left_wrap = ctk.CTkFrame(pane, fg_color=BG_DEEP)
        left = card(left_wrap, title="📋  Data Input")
        left.pack(fill="both", expand=True)
        self._build_input_panel(left)
        pane.add(left_wrap, weight=1)

        # RIGHT
        right_wrap = ctk.CTkFrame(pane, fg_color=BG_DEEP)
        right = card(right_wrap, title="📊  Analysis Results")
        right.pack(fill="both", expand=True)
        self._build_results_panel(right)
        pane.add(right_wrap, weight=1)

        def _set_sash(*_):
            total = pane.winfo_width()
            if total > 100:
                pane.sashpos(0, total // 2)
                pane.unbind("<Configure>")
        pane.bind("<Configure>", _set_sash)

        # ── Status bar ────────────────────────────────────────────────────────
        bar = ctk.CTkFrame(content, fg_color=BG_CARD, height=28, corner_radius=0)
        bar.pack(fill="x", side="bottom"); bar.pack_propagate(False)
        self.file_label = ctk.CTkLabel(bar, text="No file saved yet",
                                        font=FONT_TINY, text_color=TEXT_SEC)
        self.file_label.pack(side="left", padx=12)
        self.stat_label = ctk.CTkLabel(bar, text="", font=("Segoe UI", 9, "bold"),
                                        text_color=ACCENT)
        self.stat_label.pack(side="right", padx=12)

    def _build_input_panel(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color=BG_CARD,
                                         scrollbar_button_color=BORDER)
        scroll.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        # ── Test value (one-sample) ──
        self.test_val_frame = ctk.CTkFrame(scroll, fg_color=BG_PANEL,
                                            corner_radius=8, border_width=1,
                                            border_color=BORDER)
        ctk.CTkLabel(self.test_val_frame, text="TEST VALUE  (μ₀)",
                     font=("Segoe UI", 11, "bold"), text_color=TEXT_SEC).pack(
            anchor="w", padx=12, pady=(10, 3))
        self.test_value_entry = styled_entry(self.test_val_frame, placeholder="0", width=120)
        self.test_value_entry.insert(0, "0")
        self.test_value_entry.pack(anchor="w", padx=12, pady=(0, 10))

        # ── Group 1 ──
        self.g1_frame = ctk.CTkFrame(scroll, fg_color=BG_PANEL,
                                      corner_radius=8, border_width=1,
                                      border_color=BORDER)
        self.g1_frame.pack(fill="x", pady=(0, 8))

        g1_header = ctk.CTkFrame(self.g1_frame, fg_color="transparent")
        g1_header.pack(fill="x", padx=12, pady=(10, 4))

        # Editable badge
        self.g1_badge = ctk.CTkEntry(g1_header, width=80, height=28,
                                      font=("Segoe UI", 10, "bold"),
                                      fg_color=ACCENT, border_color=ACCENT,
                                      text_color="#0d1117", justify="center",
                                      border_width=0, corner_radius=6)
        self.g1_badge.insert(0, "Group 1")
        self.g1_badge.pack(side="left", padx=(0, 8))

        self.g1_name = styled_entry(g1_header, placeholder="Group name", height=28)
        self.g1_name.insert(0, "Group_1")
        self.g1_name.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(self.g1_frame, text="Data (comma-separated):",
                     font=FONT_TINY, text_color=TEXT_SEC).pack(anchor="w", padx=12, pady=(2, 2))
        self.g1_text = ctk.CTkTextbox(self.g1_frame, height=80,
                                       fg_color=BG_INPUT, text_color=TEXT_PRI,
                                       border_width=1, border_color=BORDER,
                                       font=FONT_BODY, corner_radius=6)
        self.g1_text.pack(fill="x", padx=12, pady=(0, 12))

        # ── Group 2 ──
        self.g2_frame = ctk.CTkFrame(scroll, fg_color=BG_PANEL,
                                      corner_radius=8, border_width=1,
                                      border_color=BORDER)

        g2_header = ctk.CTkFrame(self.g2_frame, fg_color="transparent")
        g2_header.pack(fill="x", padx=12, pady=(10, 4))

        self.g2_badge = ctk.CTkEntry(g2_header, width=80, height=28,
                                      font=("Segoe UI", 10, "bold"),
                                      fg_color=ACCENT2, border_color=ACCENT2,
                                      text_color="#0d1117", justify="center",
                                      border_width=0, corner_radius=6)
        self.g2_badge.insert(0, "Group 2")
        self.g2_badge.pack(side="left", padx=(0, 8))

        self.g2_name = styled_entry(g2_header, placeholder="Group name", height=28)
        self.g2_name.insert(0, "Group_2")
        self.g2_name.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(self.g2_frame, text="Data (comma-separated):",
                     font=FONT_TINY, text_color=TEXT_SEC).pack(anchor="w", padx=12, pady=(2, 2))
        self.g2_text = ctk.CTkTextbox(self.g2_frame, height=80,
                                       fg_color=BG_INPUT, text_color=TEXT_PRI,
                                       border_width=1, border_color=BORDER,
                                       font=FONT_BODY, corner_radius=6)
        self.g2_text.pack(fill="x", padx=12, pady=(0, 12))

        # ── Preview panel (hidden by default) ──
        self.preview_card = ctk.CTkFrame(scroll, fg_color=BG_PANEL,
                                          corner_radius=8, border_width=1,
                                          border_color=BORDER)

        ph = ctk.CTkFrame(self.preview_card, fg_color="transparent")
        ph.pack(fill="x", padx=12, pady=(8, 4))
        ctk.CTkLabel(ph, text="DATA PREVIEW", font=("Segoe UI", 11, "bold"),
                     text_color=TEXT_SEC).pack(side="left")
        ctk.CTkButton(ph, text="✕", width=22, height=22, corner_radius=4,
                      fg_color=BORDER, hover_color=DANGER,
                      command=self.hide_preview).pack(side="right")

        self.preview_text = ctk.CTkTextbox(self.preview_card, height=140,
                                            fg_color=BG_INPUT, text_color=TEXT_PRI,
                                            font=FONT_MONO, border_width=1,
                                            border_color=BORDER, corner_radius=6)
        self.preview_text.pack(fill="x", padx=12, pady=(0, 12))
        self.preview_text.configure(state="disabled")

    def _build_results_panel(self, parent):
        self.results_text = ctk.CTkTextbox(parent, fg_color=BG_INPUT,
                                            text_color=TEXT_PRI, font=FONT_MONO,
                                            wrap="none", border_width=1,
                                            border_color=BORDER, corner_radius=8)
        self.results_text.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.results_text.configure(state="disabled")

    # ── Test-type switching ────────────────────────────────────────────────────

    def _on_test_change(self, choice):
        # Reset packs
        self.test_val_frame.pack_forget()
        self.g2_frame.pack_forget()

        if choice == "one-sample":
            self.test_val_frame.pack(fill="x", pady=(0, 8),
                                     in_=self.g1_frame.master)
            self.g1_badge.delete(0, "end"); self.g1_badge.insert(0, "Sample")
            self.g1_badge.configure(fg_color=ACCENT, border_color=ACCENT)
        elif choice == "paired":
            self.g1_badge.delete(0, "end"); self.g1_badge.insert(0, "Pre")
            self.g1_badge.configure(fg_color=ACCENT, border_color=ACCENT)
            self.g2_badge.delete(0, "end"); self.g2_badge.insert(0, "Post")
            self.g2_badge.configure(fg_color=ACCENT2, border_color=ACCENT2)
            self.g2_frame.pack(fill="x", pady=(0, 8),
                               in_=self.g1_frame.master)
        else:  # independent
            self.g1_badge.delete(0, "end"); self.g1_badge.insert(0, "Group 1")
            self.g1_badge.configure(fg_color=ACCENT, border_color=ACCENT)
            self.g2_badge.delete(0, "end"); self.g2_badge.insert(0, "Group 2")
            self.g2_badge.configure(fg_color=ACCENT2, border_color=ACCENT2)
            self.g2_frame.pack(fill="x", pady=(0, 8),
                               in_=self.g1_frame.master)

    # ── Import ────────────────────────────────────────────────────────────────

    def import_data(self):
        fp = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("All", "*.*")],
            title="Import Data File"
        )
        if not fp: return
        try:
            df = pd.read_csv(fp) if fp.endswith(".csv") else pd.read_excel(fp)
            self.imported_data = df

            if len(df.columns) >= 1:
                d1 = df.iloc[:, 0].dropna().tolist()
                self.g1_text.delete("1.0", "end")
                self.g1_text.insert("1.0", ", ".join(map(str, d1)))
                self.g1_name.delete(0, "end")
                self.g1_name.insert(0, str(df.columns[0]))

            if len(df.columns) >= 2 and self.sidebar.test_type_var.get() != "one-sample":
                d2 = df.iloc[:, 1].dropna().tolist()
                self.g2_text.delete("1.0", "end")
                self.g2_text.insert("1.0", ", ".join(map(str, d2)))
                self.g2_name.delete(0, "end")
                self.g2_name.insert(0, str(df.columns[1]))

            self.sidebar.preview_btn.configure(state="normal")
            self.sidebar.status_label.configure(
                text=f"✓ Imported\n{os.path.basename(fp)}")
            messagebox.showinfo("Imported",
                                f"Data imported!\nColumns: {list(df.columns)}\nRows: {len(df)}")
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed:\n{e}")

    def show_preview(self):
        if self.imported_data is None:
            messagebox.showinfo("No Data", "Import a file first."); return

        self.preview_card.pack(fill="x", pady=(0, 8),
                               in_=self.g1_frame.master)
        self.preview_text.configure(state="normal")
        self.preview_text.delete("1.0", "end")
        df = self.imported_data
        txt = (f"Shape: {df.shape[0]} rows × {df.shape[1]} cols\n"
               f"Columns: {list(df.columns)}\n\n"
               f"First 10 rows:\n{'─'*50}\n{df.head(10).to_string()}")
        self.preview_text.insert("1.0", txt)
        self.preview_text.configure(state="disabled")

    def hide_preview(self):
        self.preview_card.pack_forget()

    # ── Parsing ───────────────────────────────────────────────────────────────

    def _parse(self, text):
        text = re.sub(r'[,\n\r\t]+', ' ', text)
        return [float(v) for v in re.findall(r'-?\d+\.?\d*', text)]

    def _fmt_p(self, p):
        return "< .001" if p < 0.001 else f"= {p:.3f}"

    def _fmt_p_tbl(self, p):
        if p < 0.001: return "< .001"
        s = f"{p:.3f}"
        return ("." + s[2:]) if s.startswith("0.") else s

    # ── Analysis ──────────────────────────────────────────────────────────────

    def run_analysis(self):
        try:
            alpha = float(self.sidebar.alpha_entry.get())
            if not 0 < alpha < 1:
                messagebox.showerror("Error", "Alpha must be between 0 and 1"); return
        except ValueError:
            messagebox.showerror("Error", "Invalid alpha value"); return

        self.alpha = alpha
        ttype = self.sidebar.test_type_var.get()

        g1_raw = self.g1_text.get("1.0", "end").strip()
        if not g1_raw:
            messagebox.showerror("Error", "Enter data for the first group/sample"); return
        g1 = self._parse(g1_raw)
        if len(g1) < 2:
            messagebox.showerror("Error", "Need at least 2 values"); return

        g1_name = self.g1_name.get().strip() or "Group_1"

        if ttype == "one-sample":
            self.results = self._one_sample(g1, g1_name)
        else:
            g2_raw = self.g2_text.get("1.0", "end").strip()
            if not g2_raw:
                messagebox.showerror("Error", "Enter data for the second group"); return
            g2 = self._parse(g2_raw)
            if len(g2) < 2:
                messagebox.showerror("Error", "Second group needs ≥ 2 values"); return
            g2_name = self.g2_name.get().strip() or "Group_2"

            if ttype == "paired":
                if len(g1) != len(g2):
                    messagebox.showerror("Error", "Paired samples must have equal size"); return
                self.results = self._paired(g1, g2, g1_name, g2_name)
            else:
                self.results = self._independent(g1, g2, g1_name, g2_name)

        self._display_results()
        self.sidebar.pdf_btn.configure(state="normal")
        self.sidebar.docx_btn.configure(state="normal")
        r = self.results
        self.stat_label.configure(
            text=f"t({r['df']:.2f}) = {r['t_statistic']:.2f}   p {self._fmt_p(r['p_value'])}   "
                 f"{'✓ Significant' if r['p_value'] < r['alpha'] else '✗ Not Significant'}"
        )

    def _one_sample(self, data, name):
        try: tv = float(self.test_value_entry.get())
        except ValueError: tv = 0
        n = len(data); m = np.mean(data); s = np.std(data, ddof=1)
        t, p = stats.ttest_1samp(data, tv)
        return dict(test_type="one-sample", test_name="One-Sample t-Test",
                    group1_name=name, group1_data=data, test_value=tv,
                    n=n, mean=m, std=s, se=s/np.sqrt(n),
                    t_statistic=t, df=n-1, p_value=p,
                    cohens_d=(m-tv)/s, alpha=self.alpha)

    def _independent(self, d1, d2, n1, n2):
        n_1, n_2 = len(d1), len(d2)
        m1, m2 = np.mean(d1), np.mean(d2)
        s1, s2 = np.std(d1, ddof=1), np.std(d2, ddof=1)
        t, p = stats.ttest_ind(d1, d2, equal_var=False)
        v1, v2 = s1**2, s2**2
        df = ((v1/n_1 + v2/n_2)**2) / ((v1/n_1)**2/(n_1-1) + (v2/n_2)**2/(n_2-1))
        ps = np.sqrt(((n_1-1)*s1**2 + (n_2-1)*s2**2) / (n_1+n_2-2))
        return dict(test_type="independent",
                    test_name="Independent Samples t-Test (Welch's)",
                    group1_name=n1, group2_name=n2,
                    group1_data=d1, group2_data=d2,
                    n1=n_1, n2=n_2, mean1=m1, mean2=m2, std1=s1, std2=s2,
                    se1=s1/np.sqrt(n_1), se2=s2/np.sqrt(n_2),
                    t_statistic=t, df=df, p_value=p,
                    cohens_d=(m1-m2)/ps, alpha=self.alpha)

    def _paired(self, d1, d2, n1, n2):
        n = len(d1); m1, m2 = np.mean(d1), np.mean(d2)
        diff = np.array(d1) - np.array(d2)
        md, sd = np.mean(diff), np.std(diff, ddof=1)
        t, p = stats.ttest_rel(d1, d2)
        return dict(test_type="paired", test_name="Paired Samples t-Test",
                    group1_name=n1, group2_name=n2,
                    group1_data=d1, group2_data=d2,
                    n=n, mean1=m1, mean2=m2, mean_diff=md, std_diff=sd,
                    se_diff=sd/np.sqrt(n), t_statistic=t, df=n-1, p_value=p,
                    cohens_d=md/sd, alpha=self.alpha)

    # ── Display ───────────────────────────────────────────────────────────────

    def _display_results(self):
        self.results_text.configure(state="normal")
        self.results_text.delete("1.0", "end")
        r = self.results
        line = "─" * 54

        ct = self.title_entry.get().strip()
        cs = self.subtitle_entry.get().strip()
        auth = self.sidebar.researcher_entry.get().strip()

        out = f"{'═'*54}\n"
        if ct: out += f"{ct.upper()}\n"
        if cs: out += f"{cs}\n"
        if auth: out += f"by: {auth}\n"
        out += f"{'═'*54}\n\n"

        out += f"{r['test_name'].upper()}\n{line}\n"
        out += f"α = {r['alpha']}   Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n"

        out += f"DESCRIPTIVE STATISTICS\n{line}\n"
        if r['test_type'] == 'one-sample':
            out += f"{r['group1_name']}:  M = {r['mean']:.2f},  SD = {r['std']:.2f},  N = {r['n']}\n"
            out += f"Test Value:  μ₀ = {r['test_value']}\n\n"
        elif r['test_type'] == 'independent':
            out += f"{r['group1_name']}:  M = {r['mean1']:.2f},  SD = {r['std1']:.2f},  N = {r['n1']}\n"
            out += f"{r['group2_name']}:  M = {r['mean2']:.2f},  SD = {r['std2']:.2f},  N = {r['n2']}\n\n"
        else:
            out += f"{r['group1_name']}:  M = {r['mean1']:.2f},  N = {r['n']}\n"
            out += f"{r['group2_name']}:  M = {r['mean2']:.2f},  N = {r['n']}\n"
            out += f"Mean Diff:  {r['mean_diff']:.2f},  SD = {r['std_diff']:.2f}\n\n"

        out += f"TEST STATISTICS\n{line}\n"
        out += f"t  = {r['t_statistic']:.4f}\n"
        out += f"df = {r['df']:.2f}\n"
        out += f"p  {self._fmt_p(r['p_value'])}\n"
        out += f"d  = {r['cohens_d']:.4f}\n\n"

        out += f"DECISION\n{line}\n"
        if r['p_value'] < r['alpha']:
            out += f"✓ REJECT H₀  (p {self._fmt_p(r['p_value'])} < α = {r['alpha']})\n\n"
        else:
            out += f"✗ FAIL TO REJECT H₀  (p {self._fmt_p(r['p_value'])} ≥ α = {r['alpha']})\n\n"

        out += f"INTERPRETATION\n{line}\n"
        out += self._interpretation() + "\n\n"
        out += "═" * 54 + "\n"

        self.results_text.insert("1.0", out)
        self.results_text.configure(state="disabled")

    def _interpretation(self):
        r = self.results
        pf = self._fmt_p(r['p_value'])
        sig = r['p_value'] < r['alpha']
        if r['test_type'] == 'one-sample':
            w = "was significantly different from" if sig else "was not significantly different from"
            return (f"A one-sample t-test revealed that the sample mean (M = {r['mean']:.2f}) "
                    f"{w} the test value ({r['test_value']}), "
                    f"t({r['df']:.0f}) = {r['t_statistic']:.2f}, p {pf}, d = {r['cohens_d']:.2f}.")
        elif r['test_type'] == 'independent':
            w = "revealed a statistically significant difference" if sig else "showed no statistically significant difference"
            return (f"An independent samples t-test {w} between "
                    f"{r['group1_name']} (M = {r['mean1']:.2f}) and "
                    f"{r['group2_name']} (M = {r['mean2']:.2f}), "
                    f"t({r['df']:.2f}) = {r['t_statistic']:.2f}, p {pf}, d = {r['cohens_d']:.2f}.")
        else:
            w = "revealed a statistically significant difference" if sig else "showed no statistically significant difference"
            return (f"A paired samples t-test {w} between "
                    f"{r['group1_name']} (M = {r['mean1']:.2f}) and "
                    f"{r['group2_name']} (M = {r['mean2']:.2f}), "
                    f"mean difference = {r['mean_diff']:.2f}, "
                    f"t({r['df']:.0f}) = {r['t_statistic']:.2f}, p {pf}, d = {r['cohens_d']:.2f}.")

    # ── Theme toggle ──────────────────────────────────────────────────────────

    def toggle_theme(self):
        if self.dark_mode:
            ctk.set_appearance_mode("light")
            self.sidebar.theme_btn.configure(text="🌙  Dark Mode")
            self.dark_mode = False
        else:
            ctk.set_appearance_mode("dark")
            self.sidebar.theme_btn.configure(text="☀️  Light Mode")
            self.dark_mode = True

    # ── Export PDF ────────────────────────────────────────────────────────────

    def export_pdf(self):
        if not self.results:
            messagebox.showerror("Error", "Run analysis first"); return
        if not PDF_AVAILABLE:
            messagebox.showerror("Error", "Install reportlab: pip install reportlab"); return

        fp = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf"), ("All", "*.*")],
            initialfile=f"ttest_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
            title="Save PDF Report"
        )
        if not fp: return

        try:
            doc = SimpleDocTemplate(fp, pagesize=letter)
            story, styles = [], getSampleStyleSheet()
            r = self.results
            ct = self.title_entry.get().strip()
            cs = self.subtitle_entry.get().strip()

            if ct:
                story.append(Paragraph(ct, styles['Title']))
                story.append(Spacer(1, 0.1*inch))
            if cs:
                sub_style = ParagraphStyle('Sub', parent=styles['Heading2'], fontSize=13,
                                           textColor=colors.HexColor('#555555'))
                story.append(Paragraph(cs, sub_style))
                story.append(Spacer(1, 0.15*inch))

            story.append(Paragraph(r['test_name'], styles['Title']))
            story.append(Spacer(1, 0.15*inch))
            story.append(Paragraph(f"Alpha: α = {r['alpha']}", styles['Normal']))
            story.append(Spacer(1, 0.25*inch))

            tbl_style = TableStyle([
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('LINEABOVE', (0,0), (-1,0), 1, colors.black),
                ('LINEBELOW', (0,0), (-1,0), 1, colors.black),
                ('LINEBELOW', (0,-1), (-1,-1), 1, colors.black),
            ])

            story.append(Paragraph("Descriptive Statistics", styles['Heading2']))
            if r['test_type'] == 'one-sample':
                d = [['Group','Mean','SD','N'],
                     [r['group1_name'], f"{r['mean']:.2f}", f"{r['std']:.2f}", str(r['n'])]]
            elif r['test_type'] == 'independent':
                d = [['Group','Mean','SD','N'],
                     [r['group1_name'], f"{r['mean1']:.2f}", f"{r['std1']:.2f}", str(r['n1'])],
                     [r['group2_name'], f"{r['mean2']:.2f}", f"{r['std2']:.2f}", str(r['n2'])]]
            else:
                d = [['Measurement','Mean','N'],
                     [r['group1_name'], f"{r['mean1']:.2f}", str(r['n'])],
                     [r['group2_name'], f"{r['mean2']:.2f}", str(r['n'])]]
            t1 = Table(d); t1.setStyle(tbl_style)
            story.append(t1); story.append(Spacer(1, 0.2*inch))

            story.append(Paragraph("Test Statistics", styles['Heading2']))
            sd = [['Statistic','Value'],
                  ['t', f"{r['t_statistic']:.4f}"],
                  ['df', f"{r['df']:.2f}"],
                  ['p', self._fmt_p_tbl(r['p_value'])],
                  ["Cohen's d", f"{r['cohens_d']:.4f}"]]
            t2 = Table(sd); t2.setStyle(tbl_style)
            story.append(t2); story.append(Spacer(1, 0.2*inch))

            story.append(Paragraph("Interpretation", styles['Heading2']))
            story.append(Paragraph(self._interpretation(), styles['Normal']))
            story.append(Spacer(1, 0.3*inch))

            fs = ParagraphStyle('Foot', parent=styles['Normal'], fontSize=7,
                                textColor=colors.grey, fontName='Helvetica-Oblique')
            story.append(Paragraph(f"Saved: {fp}", fs))
            story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", fs))

            doc.build(story)
            self.file_label.configure(text=f"Last saved: {fp}")
            self.sidebar.status_label.configure(text=f"✓ PDF Saved\n{os.path.basename(fp)}")
            messagebox.showinfo("Saved", f"PDF exported:\n{fp}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{e}")

    # ── Export DOCX ───────────────────────────────────────────────────────────

    def export_docx(self):
        if not self.results:
            messagebox.showerror("Error", "Run analysis first"); return
        if not DOCX_AVAILABLE:
            messagebox.showerror("Error", "Install python-docx: pip install python-docx"); return

        fp = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word", "*.docx"), ("All", "*.*")],
            initialfile=f"ttest_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
            title="Save Word Report"
        )
        if not fp: return

        try:
            from docx.oxml.ns import qn
            from docx.oxml import OxmlElement

            doc = Document()
            for section in doc.sections:
                section.top_margin    = Inches(0.75)
                section.bottom_margin = Inches(0.75)
                section.left_margin   = Inches(1.0)
                section.right_margin  = Inches(1.0)

            r = self.results
            ct = self.title_entry.get().strip()
            cs = self.subtitle_entry.get().strip()
            auth = self.sidebar.researcher_entry.get().strip()

            def apa_borders(table):
                tbl = table._tbl; tblPr = tbl.tblPr
                tb = OxmlElement("w:tblBorders")
                for bn in ["top","left","bottom","right","insideH","insideV"]:
                    b = OxmlElement(f"w:{bn}"); b.set(qn("w:val"), "none"); tb.append(b)
                for bn, sz in [("top","12"),("bottom","12")]:
                    b = OxmlElement(f"w:{bn}")
                    b.set(qn("w:val"), "single"); b.set(qn("w:sz"), sz); tb.append(b)
                tblPr.append(tb)

            def hdr_sep(table):
                for cell in table.rows[0].cells:
                    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
                    tcB = OxmlElement("w:tcBorders")
                    bot = OxmlElement("w:bottom")
                    bot.set(qn("w:val"), "single"); bot.set(qn("w:sz"), "6")
                    tcB.append(bot); tcPr.append(tcB)

            def cfmt(cell, text, bold=False, size=10, align="left"):
                p = cell.paragraphs[0]
                p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if align == "center"
                               else WD_ALIGN_PARAGRAPH.LEFT)
                run = p.add_run(text)
                run.font.size = Pt(size); run.bold = bold

            # Title
            if ct:
                h = doc.add_heading(ct.upper(), 0); h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if cs:
                sp = doc.add_paragraph(cs); sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                sp.runs[0].italic = True; sp.runs[0].font.size = Pt(12)
            if auth:
                ap = doc.add_paragraph(f"by: {auth}"); ap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                ap.runs[0].italic = True; ap.runs[0].font.size = Pt(10)

            doc.add_heading(r['test_name'], 1)
            doc.add_paragraph(f"Alpha: α = {r['alpha']}")
            doc.add_paragraph()

            # Descriptive stats table
            doc.add_heading('Descriptive Statistics', 2)
            if r['test_type'] == 'one-sample':
                dt = doc.add_table(rows=2, cols=4); apa_borders(dt); hdr_sep(dt)
                for i, h in enumerate(["Group","M","SD","N"]):
                    cfmt(dt.rows[0].cells[i], h, bold=True, align="center")
                cfmt(dt.rows[1].cells[0], r['group1_name'])
                cfmt(dt.rows[1].cells[1], f"{r['mean']:.2f}", align="center")
                cfmt(dt.rows[1].cells[2], f"{r['std']:.2f}", align="center")
                cfmt(dt.rows[1].cells[3], str(r['n']), align="center")
            elif r['test_type'] == 'independent':
                dt = doc.add_table(rows=3, cols=4); apa_borders(dt); hdr_sep(dt)
                for i, h in enumerate(["Group","M","SD","N"]):
                    cfmt(dt.rows[0].cells[i], h, bold=True, align="center")
                cfmt(dt.rows[1].cells[0], r['group1_name'])
                cfmt(dt.rows[1].cells[1], f"{r['mean1']:.2f}", align="center")
                cfmt(dt.rows[1].cells[2], f"{r['std1']:.2f}", align="center")
                cfmt(dt.rows[1].cells[3], str(r['n1']), align="center")
                cfmt(dt.rows[2].cells[0], r['group2_name'])
                cfmt(dt.rows[2].cells[1], f"{r['mean2']:.2f}", align="center")
                cfmt(dt.rows[2].cells[2], f"{r['std2']:.2f}", align="center")
                cfmt(dt.rows[2].cells[3], str(r['n2']), align="center")
            else:
                dt = doc.add_table(rows=3, cols=3); apa_borders(dt); hdr_sep(dt)
                for i, h in enumerate(["Measurement","M","N"]):
                    cfmt(dt.rows[0].cells[i], h, bold=True, align="center")
                cfmt(dt.rows[1].cells[0], r['group1_name'])
                cfmt(dt.rows[1].cells[1], f"{r['mean1']:.2f}", align="center")
                cfmt(dt.rows[1].cells[2], str(r['n']), align="center")
                cfmt(dt.rows[2].cells[0], r['group2_name'])
                cfmt(dt.rows[2].cells[1], f"{r['mean2']:.2f}", align="center")
                cfmt(dt.rows[2].cells[2], str(r['n']), align="center")

            doc.add_paragraph()

            # Test stats table
            doc.add_heading('Test Statistics', 2)
            st = doc.add_table(rows=5, cols=2); apa_borders(st); hdr_sep(st)
            for i, h in enumerate(["Statistic","Value"]):
                cfmt(st.rows[0].cells[i], h, bold=True, align="center")
            for row, (stat, val) in enumerate([
                ("t", f"{r['t_statistic']:.4f}"),
                ("df", f"{r['df']:.2f}"),
                ("p", self._fmt_p_tbl(r['p_value'])),
                ("Cohen's d", f"{r['cohens_d']:.4f}")
            ], 1):
                cfmt(st.rows[row].cells[0], stat)
                cfmt(st.rows[row].cells[1], val, align="center")

            doc.add_paragraph()

            # Decision
            doc.add_heading('Decision', 2)
            if r['p_value'] < r['alpha']:
                doc.add_paragraph(f"Reject the null hypothesis (p {self._fmt_p(r['p_value'])} < α = {r['alpha']}).")
            else:
                doc.add_paragraph(f"Fail to reject the null hypothesis (p {self._fmt_p(r['p_value'])} ≥ α = {r['alpha']}).")

            # Interpretation
            doc.add_heading('Interpretation', 2)
            doc.add_paragraph(self._interpretation())

            doc.add_paragraph()
            fp_p = doc.add_paragraph(f"Saved: {fp}")
            fp_p.runs[0].italic = True; fp_p.runs[0].font.size = Pt(7)
            gp = doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            gp.runs[0].italic = True; gp.runs[0].font.size = Pt(7)

            doc.save(fp)
            self.file_label.configure(text=f"Last saved: {fp}")
            self.sidebar.status_label.configure(text=f"✓ DOCX Saved\n{os.path.basename(fp)}")
            messagebox.showinfo("Saved", f"Word document exported:\n{fp}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{e}")

    # ── Clear ─────────────────────────────────────────────────────────────────

    def open_settings(self):
        SettingsWindow(self, self)

    def apply_settings(self):
        sm = SettingsManager()
        fb, fc, fh, fm, fbt, ft = sm.fonts
        ff = sm.font_family
        self.results_text.configure(font=(ff, fm), wrap=sm.wrap_mode)
        self.sidebar.status_label.configure(font=(ff, ft))
        self.stat_label.configure(font=(ff, ft, "bold"))
        self.file_label.configure(font=(ff, ft))
        self.sidebar.run_btn.configure(fg_color=sm.accent, hover_color=sm.accent_hover)
        self.sidebar.configure(width=sm.sidebar_width)

    def clear_fields(self):
        self.g1_text.delete("1.0", "end")
        self.g2_text.delete("1.0", "end")
        self.g1_name.delete(0, "end"); self.g1_name.insert(0, "Group_1")
        self.g2_name.delete(0, "end"); self.g2_name.insert(0, "Group_2")
        self.title_entry.delete(0, "end")
        self.subtitle_entry.delete(0, "end"); self.subtitle_entry.insert(0, "t-Test")
        self.results_text.configure(state="normal")
        self.results_text.delete("1.0", "end")
        self.results_text.configure(state="disabled")
        self.results = None; self.imported_data = None
        self.stat_label.configure(text="")
        self.sidebar.status_label.configure(text="")
        self.sidebar.pdf_btn.configure(state="disabled")
        self.sidebar.docx_btn.configure(state="disabled")
        self.sidebar.preview_btn.configure(state="disabled")
        self.hide_preview()


# ─── Entry point ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = TTestApp()
    app.mainloop()