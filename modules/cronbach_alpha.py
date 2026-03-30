"""
cronbach_alpha_app.py  —  Cronbach's Alpha Reliability Test
UPDATED: Session Manager integrated (multi-part save / export).
FIX: Auto-label now preserves Part number and only increments the letter.
"""

import tempfile

import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import os
from app_settings import SettingsManager, SettingsWindow

# ── Session manager ───────────────────────────────────────────────────────────
from cronbach_session_manager import (
    SessionManagerPanel,
    _next_part_label,
    export_all_to_docx,
    export_all_to_pdf,
)
# ─────────────────────────────────────────────────────────────────────────────

try:
    import winsound
    winsound.MessageBeep = lambda *a, **kw: None
except ImportError:
    pass

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BG_DEEP  = "#0d1117"
BG_CARD  = "#161b22"
BG_PANEL = "#1c2230"
BG_INPUT = "#1e2736"
ACCENT   = "#00c9a7"
ACCENT2  = "#4e9eff"
DANGER   = "#ef4444"
WARN     = "#f59e0b"
SUCCESS  = "#22c55e"
PURPLE   = "#a855f7"
TEXT_PRI = "#e6edf3"
TEXT_SEC = "#8b949e"
BORDER   = "#30363d"

FONT_HEAD = ("Segoe UI", 22, "bold")
FONT_CARD = ("Segoe UI", 15, "bold")
FONT_BODY = ("Segoe UI", 13)
FONT_MONO = ("Consolas", 12)
FONT_BTN  = ("Segoe UI", 13, "bold")
FONT_TINY = ("Segoe UI", 11)
FONT_SML  = ("Segoe UI", 9, "bold")


def divider(parent):
    ctk.CTkFrame(parent, height=1, fg_color=BORDER, corner_radius=0).pack(fill="x")


def styled_entry(parent, placeholder="", width=0, height=34):
    kw = dict(placeholder_text=placeholder, fg_color=BG_INPUT,
              border_color=BORDER, text_color=TEXT_PRI,
              placeholder_text_color=TEXT_SEC, border_width=1,
              corner_radius=6, height=height, font=FONT_BODY)
    e = ctk.CTkEntry(parent, **kw)
    if width: e.configure(width=width)
    return e


def sidebar_btn(parent, text, fg, hover, text_color=TEXT_PRI,
                font=FONT_BTN, height=36, state="normal"):
    return ctk.CTkButton(parent, text=text, fg_color=fg, hover_color=hover,
                         text_color=text_color, font=font, height=height,
                         corner_radius=8, state=state)


def card(parent, title="", **kw):
    f = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12,
                     border_width=1, border_color=BORDER, **kw)
    if title:
        ctk.CTkLabel(f, text=title, font=FONT_CARD,
                     text_color=TEXT_PRI).pack(anchor="w", padx=16, pady=(12, 4))
    return f


def sec_label(parent, text):
    ctk.CTkLabel(parent, text=text, font=("Segoe UI", 11, "bold"),
                 text_color=TEXT_SEC).pack(anchor="w", padx=18, pady=(12, 3))


# ─── Statistics ───────────────────────────────────────────────────────────────

class CronbachAlphaCalculator:
    @staticmethod
    def compute_alpha_ci(avg_r, k, n, alpha_level=0.05):
        if avg_r >= 0.999: return None, None, None
        z = 0.5 * np.log((1 + avg_r) / (1 - avg_r))
        se_z = 1 / np.sqrt(n - 3)
        z_crit = 1.96
        z_lo, z_hi = z - z_crit * se_z, z + z_crit * se_z
        r_lo = (np.exp(2*z_lo) - 1) / (np.exp(2*z_lo) + 1)
        r_hi = (np.exp(2*z_hi) - 1) / (np.exp(2*z_hi) + 1)
        a_lo = (k * r_lo) / (1 + (k-1) * r_lo)
        a_hi = (k * r_hi) / (1 + (k-1) * r_hi)
        se_a = (a_hi - a_lo) / (2 * z_crit)
        return se_a, a_lo, a_hi

    @staticmethod
    def compute(data):
        num = data.select_dtypes(include=[np.number])
        if num.empty: raise ValueError("No numeric columns found")
        clean = num.dropna()
        if len(clean) == 0: raise ValueError("No valid rows after removing missing values")
        k, n = clean.shape[1], clean.shape[0]
        if k < 2: raise ValueError("At least 2 items required")
        corr = clean.corr()
        sum_corr = corr.sum().sum() - k
        avg_r = sum_corr / (k * (k - 1))
        alpha = (k * avg_r) / (1 + (k-1) * avg_r)
        se, ci_lo, ci_hi = (None, None, None) if n <= 3 else \
            CronbachAlphaCalculator.compute_alpha_ci(avg_r, k, n)
        perfect = [(clean.columns[i], clean.columns[j])
                   for i in range(k) for j in range(i+1, k)
                   if abs(corr.iloc[i, j] - 1.0) < 0.0001]
        interp = ("Excellent" if alpha >= 0.9 else "Good" if alpha >= 0.8 else
                  "Acceptable" if alpha >= 0.7 else "Questionable" if alpha >= 0.6 else
                  "Poor" if alpha >= 0.5 else "Unacceptable")
        return dict(alpha=alpha, avg_interitem_corr=avg_r, std_error=se,
                    ci_lower=ci_lo, ci_upper=ci_hi, n_items=k, n_respondents=n,
                    interpretation=interp, item_names=list(clean.columns),
                    perfect_correlations=perfect)


class LikertExpander:
    @staticmethod
    def expand(freq_dict):
        out = []
        for v in sorted(freq_dict.keys(), reverse=True):
            out.extend([v] * freq_dict[v])
        return out

    @staticmethod
    def validate(freq_dict, scale_size):
        expected = set(range(1, scale_size + 1))
        if set(freq_dict.keys()) != expected:
            return False, f"Missing scale points. Expected {expected}"
        for v, f in freq_dict.items():
            if not isinstance(f, int) or f < 0:
                return False, f"Invalid frequency for scale {v}: {f}"
        if sum(freq_dict.values()) < 2:
            return False, "Too few respondents (need ≥ 2)"
        return True, None


# ─── PDF ──────────────────────────────────────────────────────────────────────

class PDFReport:
    @staticmethod
    def generate(results, description, filename, title=None, subtitle=None, byline=None):
        doc = SimpleDocTemplate(filename, pagesize=letter,
                                rightMargin=72, leftMargin=72,
                                topMargin=72, bottomMargin=18)
        els, styles = [], getSampleStyleSheet()
        ts = ParagraphStyle('T', parent=styles['Heading1'], fontSize=16,
                            textColor=colors.black, spaceAfter=6,
                            alignment=TA_CENTER, fontName='Helvetica-Bold')
        ss = ParagraphStyle('S', parent=styles['Normal'], fontSize=13,
                            textColor=colors.HexColor('#444444'),
                            alignment=TA_CENTER, spaceAfter=12, fontName='Helvetica')
        bs = ParagraphStyle('B', parent=styles['Normal'], fontSize=12,
                            textColor=colors.black, alignment=TA_CENTER,
                            spaceAfter=20, fontName='Helvetica')
        hs = ParagraphStyle('H', parent=styles['Heading2'], fontSize=12,
                            textColor=colors.black, spaceAfter=6, spaceBefore=12,
                            fontName='Helvetica-Bold', alignment=TA_LEFT)
        ns = ParagraphStyle('N', parent=styles['Normal'], fontSize=10,
                            spaceAfter=6, alignment=TA_LEFT, fontName='Helvetica')
        its = ParagraphStyle('I', parent=styles['Normal'], fontSize=9,
                             spaceAfter=6, alignment=TA_LEFT, fontName='Helvetica-Oblique')
        tts = ParagraphStyle('TT', parent=styles['Normal'], fontSize=11,
                             textColor=colors.black, spaceAfter=6, spaceBefore=12,
                             fontName='Helvetica-Oblique', alignment=TA_LEFT)

        els.append(Paragraph(title or "Unidimensional Reliability", ts))
        if subtitle and subtitle.strip():
            els.append(Paragraph(subtitle, ss))
        else:
            els.append(Spacer(1, 0.05*inch))
        if byline and byline.strip():
            els.append(Paragraph(byline, bs))
        els.append(Spacer(1, 0.2*inch))
        if description and description.strip():
            els.append(Paragraph(description, ns))
            els.append(Spacer(1, 0.2*inch))

        els.append(Paragraph("<i>Frequentist Scale Reliability Statistics</i>", tts))

        def fmt(v): return f"{v:.3f}" if v is not None else "—"

        tbl_data = [
            ['', '', '', '95% CI', ''],
            ['Coefficient', 'Estimate', 'Std. Error', 'Lower', 'Upper'],
            ['Coefficient α', fmt(results['alpha']), fmt(results['std_error']),
             fmt(results['ci_lower']), fmt(results['ci_upper'])],
        ]
        cw = [1.8*inch, 1.0*inch, 1.0*inch, 0.9*inch, 0.9*inch]
        at = Table(tbl_data, colWidths=cw)
        at.setStyle(TableStyle([
            ('FONTNAME', (0,0), (-1,1), 'Helvetica'), ('FONTSIZE', (0,0), (-1,1), 10),
            ('ALIGN', (0,0), (0,-1), 'LEFT'), ('ALIGN', (1,0), (-1,-1), 'CENTER'),
            ('SPAN', (3,0), (4,0)),
            ('LINEABOVE', (0,0), (-1,0), 0.5, colors.black),
            ('LINEBELOW', (0,1), (-1,1), 0.5, colors.black),
            ('LINEBELOW', (0,-1), (-1,-1), 0.5, colors.black),
            ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('LEFTPADDING', (0,0), (-1,-1), 6), ('RIGHTPADDING', (0,0), (-1,-1), 6),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        els.append(at)

        if results.get('perfect_correlations'):
            pairs = [f"{p[0]} and {p[1]}" for p in results['perfect_correlations']]
            els.append(Spacer(1, 0.1*inch))
            els.append(Paragraph(f"<i>Note.</i> Variables {' and '.join(pairs)} correlated perfectly.", its))

        els.append(Spacer(1, 0.3*inch))
        els.append(Paragraph("Summary Statistics", hs))
        sd = [['Statistic', 'Value'],
              ['Number of Items', str(results['n_items'])],
              ['Number of Respondents', str(results['n_respondents'])],
              ['Average Inter-item Correlation', f"{results['avg_interitem_corr']:.3f}"],
              ['Reliability Interpretation', results['interpretation']]]
        st = Table(sd, colWidths=[2.5*inch, 2.5*inch])
        st.setStyle(TableStyle([
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('FONTSIZE', (0,0), (-1,-1), 10),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('LINEABOVE', (0,0), (-1,0), 0.5, colors.black),
            ('LINEBELOW', (0,0), (-1,0), 0.5, colors.black),
            ('LINEBELOW', (0,-1), (-1,-1), 0.5, colors.black),
            ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ]))
        els.append(st)
        els.append(Spacer(1, 0.3*inch))

        els.append(Paragraph("Interpretation Guide", hs))
        guide = ("<b>Cronbach's Alpha Interpretation Scale:</b><br/>"
                 "• α ≥ 0.9: Excellent internal consistency<br/>"
                 "• 0.8 ≤ α &lt; 0.9: Good internal consistency<br/>"
                 "• 0.7 ≤ α &lt; 0.8: Acceptable internal consistency<br/>"
                 "• 0.6 ≤ α &lt; 0.7: Questionable internal consistency<br/>"
                 "• 0.5 ≤ α &lt; 0.6: Poor internal consistency<br/>"
                 "• α &lt; 0.5: Unacceptable internal consistency")
        els.append(Paragraph(guide, ns))
        els.append(Spacer(1, 0.5*inch))

        fs = ParagraphStyle('F', parent=styles['Normal'], fontSize=8,
                            textColor=colors.grey, alignment=TA_LEFT,
                            fontName='Helvetica-Oblique')
        ts2 = datetime.now().strftime("%B %d, %Y at %H:%M:%S")
        els.append(Paragraph(f"File: {os.path.abspath(filename)}<br/>Generated: {ts2}", fs))
        doc.build(els)


# ─── Data Table ───────────────────────────────────────────────────────────────

class DataTableFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)

    def display_data(self, df, max_rows=None):
        for w in self.winfo_children():
            w.destroy()
        if df is None or df.empty:
            ctk.CTkLabel(self, text="No data loaded", font=FONT_BODY,
                         text_color=TEXT_SEC).pack(pady=30)
            return
        ctk.CTkLabel(self, text=f"📊 {len(df)} rows × {len(df.columns)} cols",
                     font=("Segoe UI", 11, "bold"), text_color=ACCENT).pack(pady=(6, 8))
        preview_df = df if max_rows is None else df.head(max_rows)
        hdr = ctk.CTkFrame(self, fg_color=BG_PANEL, corner_radius=6)
        hdr.pack(fill="x", padx=4, pady=(0, 2))
        ctk.CTkLabel(hdr, text="#", font=("Segoe UI", 10, "bold"),
                     width=36, text_color=TEXT_SEC).pack(side="left", padx=2, pady=3)
        for col in preview_df.columns:
            ctk.CTkLabel(hdr, text=str(col)[:12], font=("Segoe UI", 10, "bold"),
                         width=78, text_color=TEXT_PRI).pack(side="left", padx=2, pady=3)
        for idx, row in preview_df.iterrows():
            rf = ctk.CTkFrame(self, fg_color="transparent")
            rf.pack(fill="x", padx=4, pady=1)
            ctk.CTkLabel(rf, text=str(idx+1), font=("Segoe UI", 10),
                         width=36, text_color=TEXT_SEC).pack(side="left", padx=2)
            for val in row:
                ctk.CTkLabel(rf, text=(str(val)[:12] if pd.notna(val) else ""),
                             font=("Segoe UI", 10), width=78,
                             text_color=TEXT_PRI).pack(side="left", padx=2)


# ─── Sidebar ──────────────────────────────────────────────────────────────────

class Sidebar(ctk.CTkFrame):
    def __init__(self, master, **kw):
        super().__init__(master, width=236, fg_color=BG_CARD,
                         corner_radius=0, **kw)
        self.pack_propagate(False)
        self._build()

    def _build(self):
        logo = ctk.CTkFrame(self, fg_color=ACCENT, corner_radius=0, height=64)
        logo.pack(fill="x"); logo.pack_propagate(False)
        ctk.CTkLabel(logo, text="  α  ", font=("Segoe UI", 30, "bold"),
                     text_color="#0d1117", fg_color=ACCENT).pack(expand=True)

        ctk.CTkLabel(self, text="Cronbach's Alpha", font=("Segoe UI", 14, "bold"),
                     text_color=TEXT_PRI, fg_color=BG_CARD).pack(pady=(16, 2))
        ctk.CTkLabel(self, text="Reliability Calculator", font=FONT_TINY,
                     text_color=TEXT_SEC, fg_color=BG_CARD).pack(pady=(0, 10))

        divider(self)

        sec_label(self, "REPORT TITLE")
        self.title_entry = styled_entry(self, placeholder="Part 1-A")
        self.title_entry.insert(0, "Part 1-A")
        self.title_entry.pack(fill="x", padx=14, pady=(0, 4))

        sec_label(self, "SUBTITLE")
        self.subtitle_entry = styled_entry(self, placeholder="Reliability Test")
        self.subtitle_entry.pack(fill="x", padx=14, pady=(0, 4))
        self.subtitle_entry.insert(0, "Reliability Test")

        sec_label(self, "AUTHOR / RESEARCHER")
        self.author_entry = styled_entry(self, placeholder="e.g. Dr. John Smith")
        self.author_entry.pack(fill="x", padx=14, pady=(0, 4))

        divider(self)

        pad = {"fill": "x", "padx": 14, "pady": 4}

        self.import_btn = sidebar_btn(self, "📁  Import CSV / Excel",
                                      fg=ACCENT2, hover="#3b7ddd")
        self.import_btn.pack(**pad)

        self.likert_btn = sidebar_btn(self, "▶  Frequency Expander",
                                      fg="#6d28d9", hover="#5b21b6")
        self.likert_btn.pack(**pad)

        divider(self)

        self.compute_btn = sidebar_btn(self, "▶   Compute α",
                                       fg=ACCENT, hover="#009e82",
                                       text_color="#0d1117",
                                       font=("Segoe UI", 13, "bold"), height=44)
        self.compute_btn.pack(**pad)

        self.save_part_btn = sidebar_btn(self, "💾  Save Part",
                                          fg=SUCCESS, hover="#16a34a",
                                          text_color="#0d1117", state="disabled")
        self.save_part_btn.pack(**pad)

        self.session_btn = sidebar_btn(self, "📁  Session Manager",
                                        fg=PURPLE, hover="#9333ea")
        self.session_btn.pack(**pad)

        self.export_btn = sidebar_btn(self, "📄  Export PDF",
                                      fg="#1d4ed8", hover="#1e3a8a", state="disabled")
        self.export_btn.pack(**pad)
        self.print_btn = sidebar_btn(self, "🖨️  Print Report",
                              fg="#7c3aed", hover="#6d28d9", state="disabled")
        self.print_btn.pack(**pad)

        self.export_data_btn = sidebar_btn(self, "💾  Export Dataset",
                                            fg="#0f766e", hover="#134e4a", state="disabled")
        self.export_data_btn.pack(**pad)

        self.clear_btn = sidebar_btn(self, "🗑  Clear / Reset",
                                     fg=DANGER, hover="#b91c1c")
        self.clear_btn.pack(**pad)

        divider(self)

        self.theme_btn = sidebar_btn(self, "☀️  Light Mode",
                                     fg="#374151", hover="#4b5563",
                                     font=FONT_BODY, height=32)
        self.theme_btn.pack(fill="x", padx=14, pady=8)

        divider(self)
        self.settings_btn = sidebar_btn(self, "⚙   Settings",
                                        fg="#374151", hover="#4b5563",
                                        font=FONT_BODY, height=32)
        self.settings_btn.pack(fill="x", padx=14, pady=8)

        self.status_label = ctk.CTkLabel(self, text="", font=FONT_TINY,
                                          text_color=ACCENT, fg_color=BG_CARD,
                                          wraplength=206)
        self.status_label.pack(side="bottom", padx=12, pady=8)


# ─── Likert Popup ─────────────────────────────────────────────────────────────

class LikertWindow(ctk.CTkToplevel):
    def __init__(self, master, on_generate, **kw):
        super().__init__(master, **kw)
        self.title("Frequency Data Expander")
        self.geometry("820x640")
        self.configure(fg_color=BG_DEEP)
        self.on_generate = on_generate
        self.scale_size = 4
        self.entries = {}
        self.transient(master)
        self.lift()
        self.focus_force()
        self.attributes("-topmost", True)
        self._build()

    def _build(self):
        hdr = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=56)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="▶  Generate from Frequency Data",
                     font=FONT_HEAD, text_color=TEXT_PRI).pack(side="left", padx=20)

        body = ctk.CTkFrame(self, fg_color=BG_DEEP)
        body.pack(fill="both", expand=True, padx=16, pady=16)

        cfg = card(body, "")
        cfg.pack(fill="x", pady=(0, 12))
        cfg_row = ctk.CTkFrame(cfg, fg_color="transparent")
        cfg_row.pack(fill="x", padx=16, pady=12)

        ctk.CTkLabel(cfg_row, text="Scale Type:", font=FONT_BODY,
                     text_color=TEXT_SEC).pack(side="left", padx=(0, 8))
        self.scale_menu = ctk.CTkOptionMenu(
            cfg_row, values=["4-point", "5-point", "7-point"],
            command=self._on_scale, fg_color=BG_INPUT, button_color=ACCENT,
            button_hover_color="#009e82", text_color=TEXT_PRI,
            dropdown_fg_color=BG_PANEL, dropdown_text_color=TEXT_PRI,
            font=FONT_BODY, height=32, corner_radius=6, width=140)
        self.scale_menu.set("4-point")
        self.scale_menu.pack(side="left", padx=(0, 24))

        ctk.CTkLabel(cfg_row, text="Number of Items:", font=FONT_BODY,
                     text_color=TEXT_SEC).pack(side="left", padx=(0, 8))
        self.num_entry = styled_entry(cfg_row, placeholder="e.g. 10", width=90, height=32)
        self.num_entry.pack(side="left", padx=(0, 12))

        ctk.CTkButton(cfg_row, text="Create Fields", command=self._create_fields,
                      fg_color=ACCENT2, hover_color="#3b7ddd", text_color=TEXT_PRI,
                      font=FONT_BTN, height=32, corner_radius=6,
                      width=130).pack(side="left")

        grid_card = card(body, "")
        grid_card.pack(fill="both", expand=True, pady=(0, 12))
        self.grid_scroll = ctk.CTkScrollableFrame(grid_card, fg_color=BG_CARD,
                                                   scrollbar_button_color=BORDER)
        self.grid_scroll.pack(fill="both", expand=True, padx=12, pady=8)

        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(fill="x")

        ctk.CTkButton(btn_row, text="✓  Generate Dataset", command=self._generate,
                      fg_color=SUCCESS, hover_color="#16a34a", text_color="#0d1117",
                      font=FONT_BTN, height=40, corner_radius=8).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_row, text="✕  Clear Fields", command=self._clear,
                      fg_color=DANGER, hover_color="#b91c1c",
                      font=FONT_BTN, height=40, corner_radius=8).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_row, text="Close", command=self._close,
                      fg_color=BG_PANEL, hover_color=BORDER,
                      font=FONT_BODY, height=40, corner_radius=8).pack(side="right")

    def _close(self):
        self.attributes("-topmost", False)
        self.destroy()

    def _on_scale(self, choice):
        self.scale_size = int(choice.split("-")[0])

    def _create_fields(self):
        try:
            n = int(self.num_entry.get())
            if not 2 <= n <= 50:
                messagebox.showwarning("Invalid", "Enter 2–50 items.", parent=self); return
        except ValueError:
            messagebox.showerror("Error", "Enter a valid number.", parent=self); return

        for w in self.grid_scroll.winfo_children():
            w.destroy()
        self.entries = {}

        hdr = ctk.CTkFrame(self.grid_scroll, fg_color=BG_PANEL, corner_radius=6)
        hdr.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(hdr, text="Item", font=("Segoe UI", 11, "bold"),
                     width=64, text_color=TEXT_SEC).pack(side="left", padx=4, pady=4)
        for i in range(self.scale_size, 0, -1):
            ctk.CTkLabel(hdr, text=f"{i}→", font=("Segoe UI", 11, "bold"),
                         width=72, text_color=ACCENT).pack(side="left", padx=4, pady=4)

        for idx in range(n):
            name = f"I{idx+1}"
            rf = ctk.CTkFrame(self.grid_scroll, fg_color="transparent")
            rf.pack(fill="x", pady=2)
            ctk.CTkLabel(rf, text=name, font=("Segoe UI", 11, "bold"),
                         width=64, text_color=TEXT_PRI).pack(side="left", padx=4)
            self.entries[name] = {}
            for sv in range(self.scale_size, 0, -1):
                e = ctk.CTkEntry(rf, width=72, height=28, fg_color=BG_INPUT,
                                 border_color=BORDER, text_color=TEXT_PRI,
                                 placeholder_text="0", font=FONT_BODY,
                                 border_width=1, corner_radius=4)
                e.pack(side="left", padx=4)
                self.entries[name][sv] = e

    def _clear(self):
        for nm in self.entries:
            for sv in self.entries[nm]:
                self.entries[nm][sv].delete(0, "end")

    def _generate(self):
        if not self.entries:
            messagebox.showwarning("No Fields", "Create fields first.", parent=self); return
        try:
            items_data = {}
            for nm in self.entries:
                fd = {}
                for sv in range(1, self.scale_size + 1):
                    txt = self.entries[nm][sv].get().strip()
                    fd[sv] = int(txt) if txt else 0
                ok, err = LikertExpander.validate(fd, self.scale_size)
                if not ok:
                    messagebox.showerror("Error", f"{nm}: {err}", parent=self); return
                items_data[nm] = fd

            totals = {nm: sum(fd.values()) for nm, fd in items_data.items()}
            if len(set(totals.values())) > 1:
                summary = "\n".join(f"  {nm}: {t}" for nm, t in totals.items())
                if not messagebox.askyesno("Warning",
                        f"Inconsistent respondent counts:\n\n{summary}\n\nProceed?",
                        parent=self):
                    return

            max_len = max(sum(fd.values()) for fd in items_data.values())
            expanded = {}
            for nm, fd in items_data.items():
                exp = LikertExpander.expand(fd)
                if len(exp) < max_len:
                    exp += [np.nan] * (max_len - len(exp))
                expanded[nm] = exp

            df = pd.DataFrame(expanded)
            self.on_generate(df)
            messagebox.showinfo("Done",
                f"Dataset generated!\n{len(df)} respondents × {len(df.columns)} items",
                parent=self)
            self._close()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self)


# ─── Main App ─────────────────────────────────────────────────────────────────

class CronbachAlphaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Cronbach's Alpha Reliability Test")
        self.geometry("1380x820")
        self.minsize(1200, 700)
        self.configure(fg_color=BG_DEEP)

        self.df      = None
        self.results = None
        self.dark_mode = True

        self.saved_parts    = []
        self._session_panel = None

        self._build_ui()

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.sidebar = Sidebar(self)
        self.sidebar.pack(side="left", fill="y")

        self.sidebar.import_btn.configure(command=self.import_data)
        self.sidebar.likert_btn.configure(command=self.open_likert)
        self.sidebar.compute_btn.configure(command=self.compute_alpha)
        self.sidebar.export_btn.configure(command=self.export_pdf)
        self.sidebar.print_btn.configure(command=self.print_report)
        self.sidebar.export_data_btn.configure(command=self.export_dataset)
        self.sidebar.clear_btn.configure(command=self.clear_all)
        self.sidebar.theme_btn.configure(command=self.toggle_theme)
        self.sidebar.settings_btn.configure(command=self.open_settings)
        self.sidebar.save_part_btn.configure(command=self.save_part)
        self.sidebar.session_btn.configure(command=self.open_session_manager)

        content = ctk.CTkFrame(self, fg_color=BG_DEEP, corner_radius=0)
        content.pack(side="left", fill="both", expand=True)

        hdr = ctk.CTkFrame(content, fg_color=BG_CARD, corner_radius=0, height=64)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="Cronbach's Alpha Reliability Test",
                     font=FONT_HEAD, text_color=TEXT_PRI).pack(side="left", padx=24)
        ctk.CTkLabel(hdr, text="JASP-Compatible Formula",
                     font=("Segoe UI", 11), text_color=TEXT_SEC).pack(side="right", padx=24)

        import tkinter as tk
        from tkinter import ttk

        outer = ctk.CTkFrame(content, fg_color=BG_DEEP)
        outer.pack(fill="both", expand=True, padx=14, pady=14)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Sash", sashthickness=6, sashrelief="flat",
                        background="#30363d")

        pane = ttk.PanedWindow(outer, orient=tk.HORIZONTAL)
        pane.pack(fill="both", expand=True)

        left_wrap = ctk.CTkFrame(pane, fg_color=BG_DEEP)
        left = card(left_wrap, title="📋  Report Description")
        left.pack(fill="both", expand=True)
        self._build_desc_panel(left)
        pane.add(left_wrap, weight=3)

        center_wrap = ctk.CTkFrame(pane, fg_color=BG_DEEP)
        center = card(center_wrap, title="📊  Dataset Preview")
        center.pack(fill="both", expand=True)
        self._build_preview_panel(center)
        pane.add(center_wrap, weight=2)

        right_wrap = ctk.CTkFrame(pane, fg_color=BG_DEEP)
        right = card(right_wrap, title="📈  Analysis Results")
        right.pack(fill="both", expand=True)
        self._build_results_panel(right)
        pane.add(right_wrap, weight=3)

        def _set_sash(*_):
            total = pane.winfo_width()
            if total > 100:
                pane.sashpos(0, int(total * 3 / 8))
                pane.sashpos(1, int(total * 6 / 8))
                pane.unbind("<Configure>")
        pane.bind("<Configure>", _set_sash)

        bar = ctk.CTkFrame(content, fg_color=BG_CARD, height=28, corner_radius=0)
        bar.pack(fill="x", side="bottom"); bar.pack_propagate(False)
        self.file_label = ctk.CTkLabel(bar, text="No file saved yet",
                                        font=FONT_TINY, text_color=TEXT_SEC)
        self.file_label.pack(side="left", padx=12)
        self.stat_label = ctk.CTkLabel(bar, text="", font=("Segoe UI", 11, "bold"),
                                        text_color=ACCENT)
        self.stat_label.pack(side="right", padx=12)

    def _build_desc_panel(self, parent):
        scroll = ctk.CTkScrollableFrame(parent, fg_color=BG_CARD,
                                         scrollbar_button_color=BORDER)
        scroll.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        ctk.CTkLabel(scroll, text="DESCRIPTION", font=("Segoe UI", 11, "bold"),
                     text_color=TEXT_SEC).pack(anchor="w", padx=4, pady=(8, 3))
        self.desc_text = ctk.CTkTextbox(scroll, height=120, fg_color=BG_INPUT,
                                         text_color=TEXT_PRI, border_width=1,
                                         border_color=BORDER, font=FONT_BODY,
                                         corner_radius=6)
        self.desc_text.pack(fill="x", padx=4, pady=(0, 12))

        guide = ctk.CTkFrame(scroll, fg_color=BG_PANEL, corner_radius=8,
                              border_width=1, border_color=BORDER)
        guide.pack(fill="x", padx=4, pady=(0, 12))
        ctk.CTkLabel(guide, text="INTERPRETATION GUIDE",
                     font=("Segoe UI", 11, "bold"), text_color=TEXT_SEC).pack(
            anchor="w", padx=12, pady=(10, 6))
        levels = [
            (ACCENT,  "α ≥ 0.90", "Excellent"),
            (SUCCESS, "α ≥ 0.80", "Good"),
            (ACCENT2, "α ≥ 0.70", "Acceptable"),
            (WARN,    "α ≥ 0.60", "Questionable"),
            ("#f97316","α ≥ 0.50", "Poor"),
            (DANGER,  "α < 0.50",  "Unacceptable"),
        ]
        for color, rng, lbl in levels:
            row = ctk.CTkFrame(guide, fg_color="transparent")
            row.pack(fill="x", padx=12, pady=2)
            ctk.CTkFrame(row, width=10, height=10, fg_color=color,
                         corner_radius=5).pack(side="left", padx=(0, 8))
            ctk.CTkLabel(row, text=rng, font=("Segoe UI", 11, "bold"),
                         text_color=TEXT_PRI, width=72).pack(side="left")
            ctk.CTkLabel(row, text=lbl, font=FONT_BODY,
                         text_color=TEXT_SEC).pack(side="left", padx=6)
        ctk.CTkFrame(guide, height=8, fg_color="transparent").pack()

    def _build_preview_panel(self, parent):
        self.data_table = DataTableFrame(parent, fg_color=BG_CARD,
                                          scrollbar_button_color=BORDER)
        self.data_table.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.data_table.display_data(None)

    def _build_results_panel(self, parent):
        self.results_text = ctk.CTkTextbox(parent, fg_color=BG_INPUT,
                                            text_color=TEXT_PRI, font=FONT_MONO,
                                            wrap="word", border_width=1,
                                            border_color=BORDER, corner_radius=8)
        self.results_text.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self._set_results(
            "═══════════════════════════════════\n"
            "  Cronbach's Alpha Calculator\n"
            "  JASP-Compatible Formula\n"
            "═══════════════════════════════════\n\n"
            "Ready to analyze.\n\n"
            "Steps:\n"
            " 1. Import data  OR\n"
            "    use Frequency Expander\n"
            " 2. Click  ▶ Compute α\n"
            " 3. Review results here\n"
            " 4. 💾 Save Part  to session\n"
            " 5. Repeat for next part\n"
            " 6. 📁 Session Manager → Export All\n\n"
            "Waiting for data…"
        )

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _set_results(self, text):
        self.results_text.configure(state="normal")
        self.results_text.delete("1.0", "end")
        self.results_text.insert("1.0", text)
        self.results_text.configure(state="disabled")

    def _current_title_entry(self) -> str:
        """Return whatever is currently typed in the Report Title entry field."""
        return self.sidebar.title_entry.get().strip()

    # ── Actions ───────────────────────────────────────────────────────────────

    def import_data(self):
        fp = filedialog.askopenfilename(
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx"), ("All", "*.*")],
            title="Select data file"
        )
        if not fp: return
        try:
            self.df = pd.read_csv(fp) if fp.endswith(".csv") else pd.read_excel(fp)
            self.data_table.display_data(self.df)
            self._set_results(
                f"═══════════════════════════════════\n"
                f"  Data Imported Successfully!\n"
                f"═══════════════════════════════════\n\n"
                f"File:    {os.path.basename(fp)}\n"
                f"Rows:    {len(self.df)}\n"
                f"Columns: {len(self.df.columns)}\n\n"
                f"Columns:\n{', '.join(str(c) for c in self.df.columns)}\n\n"
                f"✓ Click  ▶ Compute α  to proceed."
            )
            self.sidebar.export_data_btn.configure(state="normal")
            self.sidebar.status_label.configure(text=f"✓ Imported\n{os.path.basename(fp)}")
            messagebox.showinfo("Imported",
                f"Data imported!\nRows: {len(self.df)}  Columns: {len(self.df.columns)}")
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed:\n{e}")

    def open_likert(self):
        LikertWindow(self, on_generate=self._on_likert_generated)

    def _on_likert_generated(self, df):
        self.df = df
        self.data_table.display_data(df)
        self.sidebar.export_data_btn.configure(state="normal")
        self._set_results(
            f"═══════════════════════════════════\n"
            f"  Dataset Generated!\n"
            f"═══════════════════════════════════\n\n"
            f"Method:      Likert Expansion\n"
            f"Respondents: {len(df)}\n"
            f"Items:       {len(df.columns)}\n\n"
            f"Items: {', '.join(df.columns)}\n\n"
            f"✓ Click  ▶ Compute α  to proceed."
        )
        self.sidebar.status_label.configure(text=f"✓ Generated\n{len(df)} respondents")

    def compute_alpha(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Import or generate data first."); return
        try:
            self.results = CronbachAlphaCalculator.compute(self.df)
            r = self.results
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            def fmt(v): return f"{v:.4f}" if v is not None else "N/A"

            txt = (
                f"╔══════════════════════════════════════╗\n"
                f"║   CRONBACH'S ALPHA RESULTS           ║\n"
                f"║   JASP-Compatible Formula            ║\n"
                f"╚══════════════════════════════════════╝\n\n"
                f"Timestamp: {ts}\n\n"
                f"┌──────────────────────────────────────┐\n"
                f"│ RELIABILITY COEFFICIENT              │\n"
                f"├──────────────────────────────────────┤\n"
                f"  Cronbach's Alpha (α): {r['alpha']:.4f}\n"
                f"  Interpretation:       {r['interpretation']}\n\n"
                f"  Std. Error:           {fmt(r['std_error'])}\n"
                f"  95% CI Lower:         {fmt(r['ci_lower'])}\n"
                f"  95% CI Upper:         {fmt(r['ci_upper'])}\n"
                f"└──────────────────────────────────────┘\n\n"
                f"┌──────────────────────────────────────┐\n"
                f"│ DESCRIPTIVE STATISTICS               │\n"
                f"├──────────────────────────────────────┤\n"
                f"  Number of Items:      {r['n_items']}\n"
                f"  Number of Respondents:{r['n_respondents']}\n"
                f"  Avg Inter-item r:     {r['avg_interitem_corr']:.4f}\n"
                f"└──────────────────────────────────────┘\n\n"
                f"FORMULA:\n"
                f"  α = (k × r̄) / [1 + (k-1) × r̄]\n\n"
            )

            if r.get('perfect_correlations'):
                txt += "⚠ PERFECT CORRELATIONS DETECTED:\n"
                for p in r['perfect_correlations']:
                    txt += f"  • {p[0]} and {p[1]}\n"
                txt += "\n"

            txt += (f"{'═'*40}\n"
                    f"ITEMS ({r['n_items']}):\n"
                    f"{'═'*40}\n")
            txt += "\n".join(f"  {i+1}. {nm}" for i, nm in enumerate(r['item_names']))
            txt += "\n\n✓ Click  💾 Save Part  to add to session."

            self._set_results(txt)
            self.sidebar.export_btn.configure(state="normal")
            self.sidebar.print_btn.configure(state="normal")
            self.sidebar.save_part_btn.configure(state="normal")
            self.stat_label.configure(
                text=f"α = {r['alpha']:.4f}   {r['interpretation']}   "
                     f"k={r['n_items']}   N={r['n_respondents']}"
            )
            self.sidebar.status_label.configure(
                text=f"✓ α = {r['alpha']:.4f}\n{r['interpretation']}")
            messagebox.showinfo("Done",
                f"Cronbach's Alpha computed!\n\n"
                f"α = {r['alpha']:.4f}\n"
                f"Interpretation: {r['interpretation']}\n\n"
                f"Click  💾 Save Part  to add to the session.")
        except Exception as e:
            messagebox.showerror("Error", f"Computation failed:\n{e}")

    def export_pdf(self):
        if self.results is None:
            messagebox.showwarning("No Results", "Compute alpha first."); return
        fp = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"CronbachAlpha_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        )
        if not fp: return
        try:
            PDFReport.generate(
                self.results,
                self.desc_text.get("1.0", "end-1c"),
                fp,
                title=self.sidebar.title_entry.get().strip() or "Reliability Test",
                subtitle=self.sidebar.subtitle_entry.get().strip(),
                byline=self.sidebar.author_entry.get().strip()
            )
            self.file_label.configure(text=f"Last saved: {fp}")
            self.sidebar.status_label.configure(text=f"✓ PDF Saved\n{os.path.basename(fp)}")
            messagebox.showinfo("Saved", f"PDF exported:\n{fp}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{e}")

    def print_report(self):
        if self.results is None:
            messagebox.showwarning("No Results", "Compute alpha first."); return
        import subprocess, sys

        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp_path = tmp.name
        tmp.close()

        try:
            PDFReport.generate(
                self.results,
                self.desc_text.get("1.0", "end-1c"),
                tmp_path,
                title=self.sidebar.title_entry.get().strip() or "Reliability Test",
                subtitle=self.sidebar.subtitle_entry.get().strip(),
                byline=self.sidebar.author_entry.get().strip()
            )
            if sys.platform == "win32":
                import ctypes
                ctypes.windll.shell32.ShellExecuteW(
                    None, "print", tmp_path, None, None, 0)
            elif sys.platform == "darwin":
                subprocess.run(["lpr", tmp_path])
            else:
                result = subprocess.run(["lpr", tmp_path])
                if result.returncode != 0:
                    subprocess.run(["xdg-open", tmp_path])
            self.sidebar.status_label.configure(text="🖨️ Sent to printer")
        except Exception as e:
            messagebox.showerror("Print Error", f"Printing failed:\n{e}")
            try: os.unlink(tmp_path)
            except: pass

    def export_dataset(self):
        if self.df is None:
            messagebox.showwarning("No Data", "No dataset to export."); return
        fp = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
            initialfile=f"Dataset_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        if not fp: return
        try:
            self.df.to_csv(fp, index=False) if fp.endswith(".csv") else \
                self.df.to_excel(fp, index=False)
            messagebox.showinfo("Saved", f"Dataset exported:\n{os.path.basename(fp)}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{e}")

    # ── Session actions ────────────────────────────────────────────────────────

    def save_part(self):
        """Snapshot the current results and metadata into saved_parts."""
        if self.results is None:
            messagebox.showwarning("No Results", "Compute alpha first."); return

        # ── FIX: read the current title entry so the Part number is preserved ──
        current = self._current_title_entry()
        label = _next_part_label(
            [p["label"] for p in self.saved_parts],
            current_label=current
        )

        report_title = current or label

        part = {
            **self.results,
            "label":           label,
            "report_title":    report_title,
            "report_subtitle": self.sidebar.subtitle_entry.get().strip(),
            "researcher_name": self.sidebar.author_entry.get().strip(),
            "description":     self.desc_text.get("1.0", "end-1c").strip(),
            "saved_at":        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }

        self.saved_parts.append(part)

        # ── FIX: advance only the letter, keeping the same Part number ──────
        next_label = _next_part_label(
            [p["label"] for p in self.saved_parts],
            current_label=label          # base off the label just saved
        )
        self.sidebar.title_entry.delete(0, "end")
        self.sidebar.title_entry.insert(0, next_label)

        # Refresh panel if open
        if self._session_panel and self._session_panel.winfo_exists():
            self._session_panel.refresh()

        n = len(self.saved_parts)
        self.sidebar.status_label.configure(
            text=f"✓ Saved as {label}\n{n} part(s) in session")
        messagebox.showinfo("Part Saved",
            f"'{label}' added to session.\n\n"
            f"Total parts in session: {n}\n\n"
            f"Open 📁 Session Manager to view or export all.")

    def open_session_manager(self):
        """Open (or bring to front) the session manager panel."""
        if self._session_panel and self._session_panel.winfo_exists():
            self._session_panel.lift()
            self._session_panel.focus_force()
            return

        self._session_panel = SessionManagerPanel(
            master=self,
            saved_parts=self.saved_parts,
            on_delete_part=self._delete_part,
            on_export_all_docx=self._export_all_docx,
            on_export_all_pdf=self._export_all_pdf,
        )

    def _delete_part(self, label: str):
        for i, p in enumerate(self.saved_parts):
            if p["label"] == label:
                self.saved_parts.pop(i)
                break
        n = len(self.saved_parts)
        self.sidebar.status_label.configure(
            text=f"🗑 Deleted {label}\n{n} part(s) remain")

    def _export_all_docx(self):
        if not self.saved_parts:
            messagebox.showwarning("Empty Session",
                "No parts saved yet. Save at least one part first."); return
        fp = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            initialfile=f"CronbachAlpha_Session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        )
        if not fp: return
        try:
            export_all_to_docx(self.saved_parts, fp)
            messagebox.showinfo("Exported",
                f"All {len(self.saved_parts)} part(s) exported to DOCX:\n{fp}")
            self.sidebar.status_label.configure(
                text=f"✓ DOCX exported\n{os.path.basename(fp)}")
        except Exception as e:
            messagebox.showerror("Export Error", f"DOCX export failed:\n{e}")

    def _export_all_pdf(self):
        if not self.saved_parts:
            messagebox.showwarning("Empty Session",
                "No parts saved yet. Save at least one part first."); return
        fp = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"CronbachAlpha_Session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        )
        if not fp: return
        try:
            export_all_to_pdf(self.saved_parts, fp)
            messagebox.showinfo("Exported",
                f"All {len(self.saved_parts)} part(s) exported to PDF:\n{fp}")
            self.sidebar.status_label.configure(
                text=f"✓ PDF exported\n{os.path.basename(fp)}")
        except Exception as e:
            messagebox.showerror("Export Error", f"PDF export failed:\n{e}")

    # ── Misc ──────────────────────────────────────────────────────────────────

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
        self.sidebar.compute_btn.configure(fg_color=sm.accent, hover_color=sm.accent_hover)
        self.sidebar.configure(width=sm.sidebar_width)

    def clear_all(self):
        self.df = None
        self.results = None
        self.data_table.display_data(None)
        self.desc_text.delete("1.0", "end")
        self._set_results("Cleared. Import or generate data to begin.")
        self.stat_label.configure(text="")
        self.sidebar.status_label.configure(text="")
        self.sidebar.export_btn.configure(state="disabled")
        self.sidebar.export_data_btn.configure(state="disabled")
        self.sidebar.save_part_btn.configure(state="disabled")
        # ── FIX: preserve Part number when clearing, only advance letter ──────
        current = self._current_title_entry()
        next_label = _next_part_label(
            [p["label"] for p in self.saved_parts],
            current_label=current
        )
        self.sidebar.title_entry.delete(0, "end")
        self.sidebar.title_entry.insert(0, next_label)

    def toggle_theme(self):
        if self.dark_mode:
            ctk.set_appearance_mode("light")
            self.sidebar.theme_btn.configure(text="🌙  Dark Mode")
            self.dark_mode = False
        else:
            ctk.set_appearance_mode("dark")
            self.sidebar.theme_btn.configure(text="☀️  Light Mode")
            self.dark_mode = True


# ─── Entry ────────────────────────────────────────────────────────────────────

def main():
    app = CronbachAlphaApp()
    app.mainloop()


if __name__ == "__main__":
    main()