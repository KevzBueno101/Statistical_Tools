"""
Pearson r Correlation Analysis — Extended Multi-Method Module
Matches the Cronbach's Alpha Analyzer aesthetic (dark sidebar, teal accent).

New in this version:
────────────────────────────────────────────────────────────────
✦ Scrollable Sidebar
✦ Composite Score Expander  (Item-level → Respondent-level)
✦ Scatter Plot Viewer
✦ Dataset Viewer
✦ Dataset Editor
✦ Variable Descriptions
✦ [NEW] Free Pair Selector
    • Freely choose any combination of variable pairs
    • Add / Remove individual pair rows
    • "Add All Pairs" convenience button
    • Duplicate and self-pair validation
────────────────────────────────────────────────────────────────
"""

import os, sys, tempfile, subprocess, warnings, copy
from datetime import datetime

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import pandas as pd
import numpy as np
from scipy import stats
from scipy.stats import norm

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer)
from reportlab.lib.enums import TA_CENTER, TA_LEFT

warnings.filterwarnings("ignore", category=RuntimeWarning)

try:
    from app_settings import SettingsManager, SettingsWindow
    HAS_SETTINGS = True
except ImportError:
    HAS_SETTINGS = False

try:
    import winsound
    winsound.MessageBeep = lambda *a, **kw: None
except ImportError:
    pass

# ─── Palette ──────────────────────────────────────────────────────────────────
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
ORANGE   = "#f97316"
CYAN     = "#06b6d4"
PINK     = "#ec4899"
TEXT_PRI = "#e6edf3"
TEXT_SEC = "#8b949e"
BORDER   = "#30363d"

FONT_HEAD = ("Segoe UI", 22, "bold")
FONT_CARD = ("Segoe UI", 15, "bold")
FONT_BODY = ("Segoe UI", 13)
FONT_MONO = ("Consolas", 12)
FONT_BTN  = ("Segoe UI", 13, "bold")
FONT_TINY = ("Segoe UI", 11)

# ─── Method Registry ──────────────────────────────────────────────────────────
METHODS = {
    "pearson": {
        "label": "Pearson Product-Moment", "symbol": "r", "color": ACCENT2,
        "desc": (
            "Pearson Product-Moment Correlation (r)\n\n"
            "Classical linear association between two continuous variables.\n\n"
            "Assumptions:\n"
            "  • Both variables are continuous (interval/ratio)\n"
            "  • Bivariate normality\n  • Linear relationship\n"
            "  • No significant outliers\n\n"
            "Range: −1 to +1\n"
            "Use when: Both variables are normally distributed."
        ),
    },
    "spearman": {
        "label": "Spearman Rank-Order", "symbol": "ρ", "color": SUCCESS,
        "desc": (
            "Spearman Rank-Order Correlation (ρ)\n\n"
            "Non-parametric. Converts scores to ranks first.\n\n"
            "Assumptions:\n"
            "  • Ordinal, interval, or ratio data\n"
            "  • Does NOT require normality\n"
            "  • Robust to outliers\n\n"
            "Range: −1 to +1\n"
            "Use when: Data is ordinal, skewed, or has outliers."
        ),
    },
    "kendall": {
        "label": "Kendall's Tau-b", "symbol": "τb", "color": PURPLE,
        "desc": (
            "Kendall's Tau-b (τb)\n\n"
            "Non-parametric. Concordant vs discordant pairs.\n\n"
            "Assumptions:\n"
            "  • Ordinal or continuous data\n"
            "  • More conservative than Spearman ρ\n"
            "  • Better with small N or many ties\n\n"
            "Range: −1 to +1\n"
            "Use when: Small samples, many tied ranks."
        ),
    },
    "point_biserial": {
        "label": "Point-Biserial", "symbol": "rpb", "color": WARN,
        "desc": (
            "Point-Biserial Correlation (rpb)\n\n"
            "Pearson r when one variable is a natural dichotomy.\n\n"
            "Assumptions:\n"
            "  • One continuous + one true binary (0/1)\n"
            "  • Binary = natural category (e.g. sex)\n\n"
            "Range: −1 to +1\n"
            "Use when: One variable is a natural dichotomy."
        ),
    },
    "biserial": {
        "label": "Biserial", "symbol": "rb", "color": ORANGE,
        "desc": (
            "Biserial Correlation (rb)\n\n"
            "Corrects rpb for artificially dichotomised variables.\n"
            "Formula: rb = rpb × √(pq) / z(p)\n\n"
            "Assumptions:\n"
            "  • One continuous + one artificially binary\n\n"
            "Range: can exceed ±1 theoretically\n"
            "Use when: Binary variable was originally continuous."
        ),
    },
    "phi": {
        "label": "Phi Coefficient", "symbol": "φ", "color": ACCENT,
        "desc": (
            "Phi Coefficient (φ)\n\n"
            "Pearson r for two genuine dichotomous variables.\n"
            "Formula: φ = (ad−bc) / √[(a+b)(c+d)(a+c)(b+d)]\n\n"
            "Assumptions:\n"
            "  • Both variables are natural binary (0/1)\n\n"
            "Range: −1 to +1\n"
            "Use when: Both variables are true dichotomies."
        ),
    },
    "tetrachoric": {
        "label": "Tetrachoric", "symbol": "rt", "color": DANGER,
        "desc": (
            "Tetrachoric Correlation (rt)\n\n"
            "Estimates r if both artificially binary variables\n"
            "were actually observed as continuous normals.\n"
            "Approx: rt ≈ cos(π / (1 + √(bc/ad)))\n\n"
            "Range: −1 to +1\n"
            "Use when: Both binary vars were cut from normals."
        ),
    },
    "partial": {
        "label": "Partial Correlation", "symbol": "r·z", "color": CYAN,
        "desc": (
            "Partial Correlation (r·z)\n\n"
            "Pearson r between X and Y after removing the linear\n"
            "influence of control variable(s) from BOTH.\n\n"
            "Range: −1 to +1\n"
            "Use when: Isolating X–Y free from confounds."
        ),
    },
    "semi_partial": {
        "label": "Semi-Partial (Part)", "symbol": "sr", "color": PINK,
        "desc": (
            "Semi-Partial (Part) Correlation (sr)\n\n"
            "Removes covariate influence from Y only.\n"
            "sr² = unique variance in Y explained by X.\n\n"
            "Range: −1 to +1\n"
            "Use when: X's unique contribution to Y."
        ),
    },
}

# ─── UI helpers ───────────────────────────────────────────────────────────────

def divider(parent):
    ctk.CTkFrame(parent, height=1, fg_color=BORDER,
                 corner_radius=0).pack(fill="x")

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
    f = ctk.CTkFrame(parent, fg_color=BG_CARD, corner_radius=12,
                     border_width=1, border_color=BORDER, **kw)
    if title:
        ctk.CTkLabel(f, text=title, font=FONT_CARD,
                     text_color=TEXT_PRI).pack(anchor="w", padx=16, pady=(12, 4))
    return f

def sec_label(parent, text):
    ctk.CTkLabel(parent, text=text, font=("Segoe UI", 11, "bold"),
                 text_color=TEXT_SEC).pack(anchor="w", padx=18, pady=(12, 3))


# ─── Statistics Engine ────────────────────────────────────────────────────────

class CorrelationEngine:
    @staticmethod
    def _fisher_ci(r, n):
        if n < 4 or abs(r) >= 0.9999: return None, None
        z = 0.5 * np.log((1 + r) / (1 - r))
        se = 1.0 / np.sqrt(n - 3)
        z_lo, z_hi = z - 1.95996*se, z + 1.95996*se
        r_lo = (np.exp(2*z_lo)-1)/(np.exp(2*z_lo)+1)
        r_hi = (np.exp(2*z_hi)-1)/(np.exp(2*z_hi)+1)
        return r_lo, r_hi

    @staticmethod
    def _effect(r):
        a = abs(r)
        if a >= .90: return "Very High"
        if a >= .70: return "High"
        if a >= .50: return "Moderate"
        if a >= .30: return "Low"
        return "Negligible"

    @staticmethod
    def _stars(p):
        if p < .001: return "***"
        if p < .01:  return "**"
        if p < .05:  return "*"
        return "ns"

    @staticmethod
    def _direction(r):
        if r >  .01: return "Positive"
        if r < -.01: return "Negative"
        return "None"

    @staticmethod
    def pearson_critical_r(n, alpha=0.05):
        df = n - 2
        if df < 1:
            raise ValueError("n must be at least 3.")
        t_crit = stats.t.ppf(1 - alpha / 2, df)
        r_crit = float(np.sqrt(t_crit**2 / (t_crit**2 + df)))
        return r_crit, df, float(t_crit)

    @classmethod
    def _base(cls, xn, yn, r, p, n, ci_lo, ci_hi):
        t = r*np.sqrt(n-2)/np.sqrt(max(1-r**2,1e-15)) if abs(r)<.9999 else np.inf
        return {"x":xn,"y":yn,"r":r,"r_sq":r**2,"p":p,"t_stat":t,
                "n":n,"ci_lower":ci_lo,"ci_upper":ci_hi,
                "effect":cls._effect(r),"direction":cls._direction(r),
                "sig":p<.05,"stars":cls._stars(p)}

    @classmethod
    def pearson(cls, df, cols):
        pairs=[]
        for i in range(len(cols)):
            for j in range(i+1,len(cols)):
                xn,yn=cols[i],cols[j]; sub=df[[xn,yn]].dropna(); n=len(sub)
                if n<3: continue
                r,p=stats.pearsonr(sub[xn],sub[yn])
                pairs.append(cls._base(xn,yn,r,p,n,*cls._fisher_ci(r,n)))
        return cls._wrap(pairs,df[cols].corr("pearson"),cols,"pearson")

    @classmethod
    def spearman(cls, df, cols):
        pairs=[]
        for i in range(len(cols)):
            for j in range(i+1,len(cols)):
                xn,yn=cols[i],cols[j]; sub=df[[xn,yn]].dropna(); n=len(sub)
                if n<3: continue
                r,p=stats.spearmanr(sub[xn],sub[yn])
                pairs.append(cls._base(xn,yn,r,p,n,*cls._fisher_ci(r,n)))
        return cls._wrap(pairs,df[cols].corr("spearman"),cols,"spearman")

    @classmethod
    def kendall(cls, df, cols):
        pairs=[]
        for i in range(len(cols)):
            for j in range(i+1,len(cols)):
                xn,yn=cols[i],cols[j]; sub=df[[xn,yn]].dropna(); n=len(sub)
                if n<3: continue
                tau,p=stats.kendalltau(sub[xn],sub[yn])
                se=np.sqrt((2*(2*n+5))/(9*n*(n-1)))
                pair=cls._base(xn,yn,tau,p,n,tau-1.95996*se,tau+1.95996*se)
                pair["t_stat"]=tau/se if se>0 else 0; pairs.append(pair)
        return cls._wrap(pairs,df[cols].corr("kendall"),cols,"kendall")

    @classmethod
    def point_biserial(cls, df, cols):
        pairs=[]
        for i in range(len(cols)):
            for j in range(i+1,len(cols)):
                xn,yn=cols[i],cols[j]; sub=df[[xn,yn]].dropna(); n=len(sub)
                if n<3: continue
                xu,yu=sub[xn].nunique(),sub[yn].nunique()
                if yu==2: cont,binn=xn,yn
                elif xu==2: cont,binn=yn,xn
                else:
                    r,p=stats.pearsonr(sub[xn],sub[yn])
                    pairs.append(cls._base(xn,yn,r,p,n,*cls._fisher_ci(r,n))); continue
                rpb,p=stats.pointbiserialr(sub[binn],sub[cont])
                pair=cls._base(xn,yn,rpb,p,n,*cls._fisher_ci(rpb,n))
                pair["notes"]=f"Binary:{binn}  Cont:{cont}"; pairs.append(pair)
        return cls._wrap(pairs,df[cols].corr("pearson"),cols,"point_biserial")

    @classmethod
    def biserial(cls, df, cols):
        pairs=[]
        for i in range(len(cols)):
            for j in range(i+1,len(cols)):
                xn,yn=cols[i],cols[j]; sub=df[[xn,yn]].dropna(); n=len(sub)
                if n<3: continue
                xu,yu=sub[xn].nunique(),sub[yn].nunique()
                if yu==2: cont,binn=xn,yn
                elif xu==2: cont,binn=yn,xn
                else:
                    r,p=stats.pearsonr(sub[xn],sub[yn])
                    pairs.append(cls._base(xn,yn,r,p,n,*cls._fisher_ci(r,n))); continue
                rpb,p=stats.pointbiserialr(sub[binn],sub[cont])
                prop=sub[binn].mean(); q=1-prop
                z_p=norm.pdf(norm.ppf(prop))
                rb=rpb*np.sqrt(prop*q)/z_p if z_p>0 else rpb
                pair=cls._base(xn,yn,rb,p,n,*cls._fisher_ci(np.clip(rb,-0.9999,.9999),n))
                pair["notes"]=f"Binary:{binn}  p(1)={prop:.3f}"; pairs.append(pair)
        return cls._wrap(pairs,df[cols].corr("pearson"),cols,"biserial")

    @classmethod
    def phi(cls, df, cols):
        pairs=[]
        for i in range(len(cols)):
            for j in range(i+1,len(cols)):
                xn,yn=cols[i],cols[j]; sub=df[[xn,yn]].dropna(); n=len(sub)
                if n<3: continue
                r,p=stats.pearsonr(sub[xn],sub[yn])
                try:
                    ct=pd.crosstab(sub[xn],sub[yn])
                    if ct.shape==(2,2):
                        a,b,c_,d=ct.iloc[0,0],ct.iloc[0,1],ct.iloc[1,0],ct.iloc[1,1]
                        den=np.sqrt((a+b)*(c_+d)*(a+c_)*(b+d))
                        if den>0: r=(a*d-b*c_)/den; p=1-stats.chi2.cdf(r**2*n,df=1)
                except: pass
                pairs.append(cls._base(xn,yn,r,p,n,*cls._fisher_ci(np.clip(r,-.9999,.9999),n)))
        return cls._wrap(pairs,df[cols].corr("pearson"),cols,"phi")

    @classmethod
    def tetrachoric(cls, df, cols):
        pairs=[]
        for i in range(len(cols)):
            for j in range(i+1,len(cols)):
                xn,yn=cols[i],cols[j]; sub=df[[xn,yn]].dropna(); n=len(sub)
                if n<4: continue
                try:
                    ct=pd.crosstab(sub[xn],sub[yn])
                    if ct.shape!=(2,2): raise ValueError
                    a,b,c_,d=ct.iloc[0,0]+.5,ct.iloc[0,1]+.5,ct.iloc[1,0]+.5,ct.iloc[1,1]+.5
                    rt=np.cos(np.pi/(1+np.sqrt(b*c_/(a*d))))
                    p1,p2=(a+b)/n,(a+c_)/n
                    z1,z2=norm.pdf(norm.ppf(p1)),norm.pdf(norm.ppf(p2))
                    se=np.sqrt(p1*(1-p1)*p2*(1-p2)/max(n*z1**2*z2**2,1e-15))
                    zs=rt/se if se>0 else 0; pv=2*(1-norm.cdf(abs(zs)))
                    pair=cls._base(xn,yn,rt,pv,n,rt-1.95996*se,rt+1.95996*se)
                    pair["t_stat"]=zs; pairs.append(pair)
                except:
                    r,p=stats.pearsonr(sub[xn],sub[yn])
                    pairs.append(cls._base(xn,yn,r,p,n,*cls._fisher_ci(r,n)))
        return cls._wrap(pairs,df[cols].corr("pearson"),cols,"tetrachoric")

    @classmethod
    def partial(cls, df, cols, controls):
        if not controls: raise ValueError("Need ≥1 control variable.")
        pairs=[]
        for i in range(len(cols)):
            for j in range(i+1,len(cols)):
                xn,yn=cols[i],cols[j]
                sub=df[[xn,yn]+controls].dropna(); n=len(sub)
                if n<len(controls)+4: continue
                Z=np.column_stack([np.ones(n),sub[controls].values])
                rx=sub[xn].values-Z@np.linalg.lstsq(Z,sub[xn].values,rcond=None)[0]
                ry=sub[yn].values-Z@np.linalg.lstsq(Z,sub[yn].values,rcond=None)[0]
                r,_=stats.pearsonr(rx,ry); df_r=n-len(controls)-2
                t=r*np.sqrt(df_r)/np.sqrt(max(1-r**2,1e-15))
                p=2*(1-stats.t.cdf(abs(t),df=df_r))
                pair=cls._base(xn,yn,r,p,n,*cls._fisher_ci(r,n-len(controls)))
                pair.update({"t_stat":t,"df":df_r,"notes":f"Ctrl:{','.join(controls)}"})
                pairs.append(pair)
        return cls._wrap(pairs,df[cols].corr("pearson"),cols,"partial",
                         notes=f"Controls: {', '.join(controls)}")

    @classmethod
    def semi_partial(cls, df, cols, controls):
        if not controls: raise ValueError("Need ≥1 control variable.")
        pairs=[]
        for i in range(len(cols)):
            for j in range(i+1,len(cols)):
                xn,yn=cols[i],cols[j]
                sub=df[[xn,yn]+controls].dropna(); n=len(sub)
                if n<len(controls)+4: continue
                Z=np.column_stack([np.ones(n),sub[controls].values])
                ry=sub[yn].values-Z@np.linalg.lstsq(Z,sub[yn].values,rcond=None)[0]
                r,_=stats.pearsonr(sub[xn].values,ry); df_r=n-len(controls)-2
                t=r*np.sqrt(df_r)/np.sqrt(max(1-r**2,1e-15))
                p=2*(1-stats.t.cdf(abs(t),df=df_r))
                pair=cls._base(xn,yn,r,p,n,*cls._fisher_ci(r,n-len(controls)))
                pair.update({"t_stat":t,"df":df_r,
                             "notes":f"Y residualised on:{','.join(controls)} sr²={r**2:.4f}"})
                pairs.append(pair)
        return cls._wrap(pairs,df[cols].corr("pearson"),cols,"semi_partial",
                         notes=f"Controls: {', '.join(controls)}")

    @classmethod
    def compute(cls, df, method_key, cols, controls=None):
        controls=controls or []
        if method_key=="partial": return cls.partial(df,cols,controls)
        if method_key=="semi_partial": return cls.semi_partial(df,cols,controls)
        return {"pearson":cls.pearson,"spearman":cls.spearman,"kendall":cls.kendall,
                "point_biserial":cls.point_biserial,"biserial":cls.biserial,
                "phi":cls.phi,"tetrachoric":cls.tetrachoric}[method_key](df,cols)

    @staticmethod
    def _wrap(pairs, corr_mat, cols, method_key, notes=""):
        if not pairs: raise ValueError("Not enough valid paired observations.")
        return {"pairs":pairs,"corr_matrix":corr_mat,"variables":cols,
                "n_vars":len(cols),"method_key":method_key,
                "method_info":METHODS[method_key],"notes":notes}


# ─── Composite Score Expander ─────────────────────────────────────────────────

class CompositeExpanderWindow(ctk.CTkToplevel):

    def __init__(self, master, on_generate, **kw):
        super().__init__(master, **kw)
        self.title("Composite Score Expander  —  Item → Respondent Level")
        self.geometry("1100x760")
        self.configure(fg_color=BG_DEEP)
        self.on_generate  = on_generate
        self.scale_size   = 4
        self.likert_items = {}
        self.test_items   = {}
        self.part_names   = {}
        self.transient(master)
        self.lift(); self.focus_force()
        self.attributes("-topmost", True)
        self._build()

    def _build(self):
        hdr = ctk.CTkFrame(self, fg_color=ACCENT2, corner_radius=0, height=60)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="  ▶  Composite Score Expander",
                     font=FONT_HEAD, text_color="#0d1117",
                     fg_color=ACCENT2).pack(side="left", padx=20)
        body = ctk.CTkFrame(self, fg_color=BG_DEEP)
        body.pack(fill="both", expand=True, padx=16, pady=14)
        cfg = card(body, ""); cfg.pack(fill="x", pady=(0, 10))
        crow = ctk.CTkFrame(cfg, fg_color="transparent"); crow.pack(fill="x", padx=14, pady=10)
        ctk.CTkLabel(crow, text="Scale:", font=FONT_BODY, text_color=TEXT_SEC).pack(side="left", padx=(0,6))
        self.scale_menu = ctk.CTkOptionMenu(
            crow, values=["4-point","5-point","7-point"],
            command=lambda c: setattr(self,"scale_size",int(c.split("-")[0])),
            fg_color=BG_INPUT, button_color=ACCENT2, button_hover_color="#3b7ddd",
            text_color=TEXT_PRI, dropdown_fg_color=BG_PANEL, dropdown_text_color=TEXT_PRI,
            font=FONT_BODY, height=32, corner_radius=6, width=120)
        self.scale_menu.set("4-point"); self.scale_menu.pack(side="left",padx=(0,20))
        ctk.CTkLabel(crow, text="No. of Items:", font=FONT_BODY, text_color=TEXT_SEC).pack(side="left", padx=(0,6))
        self.num_entry = styled_entry(crow, placeholder="e.g. 20", width=80, height=32)
        self.num_entry.pack(side="left", padx=(0,12))
        ctk.CTkLabel(crow, text="Parts:", font=FONT_BODY, text_color=TEXT_SEC).pack(side="left", padx=(0,6))
        self.parts_entry = styled_entry(crow, placeholder="e.g. 2", width=60, height=32)
        self.parts_entry.insert(0,"2"); self.parts_entry.pack(side="left",padx=(0,12))
        ctk.CTkButton(crow, text="Create Fields", command=self._create_fields,
                      fg_color=ACCENT2, hover_color="#3b7ddd", text_color=TEXT_PRI,
                      font=FONT_BTN, height=32, corner_radius=6, width=140).pack(side="left",padx=(0,14))
        ctk.CTkLabel(crow, text="Score:", font=FONT_BODY, text_color=TEXT_SEC).pack(side="left", padx=(0,6))
        self.score_mode = ctk.CTkOptionMenu(
            crow, values=["Mean","Total"],
            fg_color=BG_INPUT, button_color=PURPLE, button_hover_color="#7c3aed",
            text_color=TEXT_PRI, dropdown_fg_color=BG_PANEL, dropdown_text_color=TEXT_PRI,
            font=FONT_BODY, height=32, corner_radius=6, width=100)
        self.score_mode.set("Mean"); self.score_mode.pack(side="left")
        grid_card = card(body, ""); grid_card.pack(fill="both", expand=True, pady=(0,10))
        self.grid_scroll = ctk.CTkScrollableFrame(grid_card, fg_color=BG_CARD, scrollbar_button_color=BORDER)
        self.grid_scroll.pack(fill="both", expand=True, padx=12, pady=8)
        brow = ctk.CTkFrame(body, fg_color="transparent"); brow.pack(fill="x")
        ctk.CTkButton(brow, text="✓  Generate Composite Scores", command=self._generate,
                      fg_color=SUCCESS, hover_color="#16a34a", text_color="#0d1117",
                      font=FONT_BTN, height=42, corner_radius=8).pack(side="left",padx=(0,8))
        ctk.CTkButton(brow, text="✕  Clear", command=self._clear,
                      fg_color=DANGER, hover_color="#b91c1c", font=FONT_BTN,
                      height=42, corner_radius=8).pack(side="left",padx=(0,8))
        ctk.CTkButton(brow, text="Close", command=self.destroy,
                      fg_color=BG_PANEL, hover_color=BORDER, font=FONT_BODY,
                      height=42, corner_radius=8).pack(side="right")

    def _create_fields(self):
        try:
            n_items = int(self.num_entry.get()); n_parts = int(self.parts_entry.get())
            if not 2<=n_items<=80: messagebox.showwarning("Invalid","Enter 2–80 items.",parent=self); return
            if not 1<=n_parts<=10: messagebox.showwarning("Invalid","Enter 1–10 parts.",parent=self); return
        except ValueError: messagebox.showerror("Error","Enter valid numbers.",parent=self); return
        self.parts_list = [f"Part {i+1}" for i in range(n_parts)]
        for w in self.grid_scroll.winfo_children(): w.destroy()
        self.entries={}; self.part_vars={}
        hdr = ctk.CTkFrame(self.grid_scroll, fg_color=BG_PANEL, corner_radius=6); hdr.pack(fill="x", pady=(0,4))
        ctk.CTkLabel(hdr, text="Item", font=("Segoe UI",11,"bold"), width=60, text_color=TEXT_SEC).pack(side="left",padx=4,pady=4)
        ctk.CTkLabel(hdr, text="Assign to Part", font=("Segoe UI",11,"bold"), width=110, text_color=ACCENT2).pack(side="left",padx=4,pady=4)
        for sv in range(self.scale_size,0,-1):
            ctk.CTkLabel(hdr, text=f"f({sv})", font=("Segoe UI",11,"bold"), width=64, text_color=ACCENT).pack(side="left",padx=4,pady=4)
        for idx in range(n_items):
            name = f"I{idx+1}"
            rf = ctk.CTkFrame(self.grid_scroll, fg_color="transparent"); rf.pack(fill="x", pady=2)
            ctk.CTkLabel(rf, text=name, font=("Segoe UI",11,"bold"), width=60, text_color=TEXT_PRI).pack(side="left",padx=4)
            pv = tk.StringVar(value=self.parts_list[0]); self.part_vars[name] = pv
            ctk.CTkOptionMenu(rf, variable=pv, values=self.parts_list, fg_color=BG_INPUT,
                              button_color=PURPLE, button_hover_color="#7c3aed", text_color=TEXT_PRI,
                              dropdown_fg_color=BG_PANEL, dropdown_text_color=TEXT_PRI,
                              font=("Segoe UI",11), height=28, corner_radius=4, width=110).pack(side="left",padx=4)
            self.entries[name]={}
            for sv in range(self.scale_size,0,-1):
                e = ctk.CTkEntry(rf, width=64, height=28, fg_color=BG_INPUT, border_color=BORDER,
                                 text_color=TEXT_PRI, placeholder_text="0", font=FONT_BODY,
                                 border_width=1, corner_radius=4)
                e.pack(side="left",padx=4); self.entries[name][sv] = e

    def _clear(self):
        for nm in self.entries:
            for sv in self.entries[nm]: self.entries[nm][sv].delete(0,"end")

    def _generate(self):
        if not self.entries: messagebox.showwarning("No Fields","Create fields first.",parent=self); return
        mode = self.score_mode.get()
        try:
            items_data = {}
            for nm in self.entries:
                fd={}
                for sv in range(1,self.scale_size+1):
                    txt=self.entries[nm][sv].get().strip(); fd[sv]=int(txt) if txt else 0
                if sum(fd.values())<2: messagebox.showerror("Error",f"{nm}: Total frequency must be ≥ 2",parent=self); return
                items_data[nm]=fd
            totals={nm:sum(fd.values()) for nm,fd in items_data.items()}; max_n=max(totals.values())
            expanded={}
            for nm,fd in items_data.items():
                scores=[]
                for sv in sorted(fd.keys(),reverse=True): scores.extend([sv]*fd[sv])
                if len(scores)<max_n: scores+=[np.nan]*(max_n-len(scores))
                expanded[nm]=scores
            raw_df=pd.DataFrame(expanded)
            part_assignment={nm:self.part_vars[nm].get() for nm in self.entries}
            composite={}
            for part in self.parts_list:
                part_items=[nm for nm,pt in part_assignment.items() if pt==part]
                if not part_items: continue
                sub=raw_df[part_items].apply(pd.to_numeric,errors="coerce")
                composite[part]=sub.mean(axis=1) if mode=="Mean" else sub.sum(axis=1)
            if len(composite)<2: messagebox.showwarning("Too Few Parts","At least 2 parts with items are needed.",parent=self); return
            out_df=pd.DataFrame(composite)
            mapping_info={part:[nm for nm,pt in part_assignment.items() if pt==part] for part in self.parts_list}
            self.on_generate(out_df, mapping_info, mode, raw_df)
            messagebox.showinfo("Done",f"Composite scores generated!\nRespondents: {len(out_df)}\nParts: {len(composite)}\nScore type: {mode}",parent=self)
            self.destroy()
        except Exception as e: messagebox.showerror("Error",str(e),parent=self)


# ─── Dataset Viewer Window ────────────────────────────────────────────────────

class DatasetViewerWindow(ctk.CTkToplevel):
    def __init__(self, master, df, on_edit_callback=None, **kw):
        super().__init__(master, **kw)
        self.title("Dataset Viewer"); self.geometry("1200x780"); self.configure(fg_color=BG_DEEP)
        self.df = df.copy(); self.display_df = df.copy()
        self.on_edit_callback = on_edit_callback
        self._sort_col = None; self._sort_asc = True
        self.transient(master); self.lift(); self.focus_force(); self._build()

    def _build(self):
        hdr = ctk.CTkFrame(self, fg_color=CYAN, corner_radius=0, height=58)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="  🔍  Dataset Viewer", font=("Segoe UI",20,"bold"), text_color="#0d1117").pack(side="left", padx=20)
        ctk.CTkLabel(hdr, text=f"  {len(self.df)} rows × {len(self.df.columns)} columns", font=("Segoe UI",12), text_color="#0d4a52").pack(side="left", padx=10)
        tb = ctk.CTkFrame(self, fg_color=BG_PANEL, corner_radius=0, height=48); tb.pack(fill="x"); tb.pack_propagate(False)
        ctk.CTkLabel(tb, text="Search:", font=FONT_BODY, text_color=TEXT_SEC).pack(side="left", padx=(12,4))
        self.search_var = tk.StringVar(); self.search_var.trace_add("write", self._on_search)
        ctk.CTkEntry(tb, textvariable=self.search_var, placeholder_text="Filter any column…",
                     fg_color=BG_INPUT, border_color=BORDER, text_color=TEXT_PRI,
                     placeholder_text_color=TEXT_SEC, height=30, width=220, font=FONT_BODY,
                     border_width=1, corner_radius=6).pack(side="left", padx=(0,16))
        ctk.CTkLabel(tb, text="Column:", font=FONT_BODY, text_color=TEXT_SEC).pack(side="left", padx=(0,4))
        self.filter_col_var = tk.StringVar(value="All")
        ctk.CTkOptionMenu(tb, variable=self.filter_col_var, values=["All"]+list(self.df.columns),
                          command=lambda _: self._on_search(), fg_color=BG_INPUT, button_color=CYAN,
                          button_hover_color="#0891b2", text_color=TEXT_PRI, dropdown_fg_color=BG_PANEL,
                          dropdown_text_color=TEXT_PRI, font=FONT_BODY, height=30, corner_radius=6,
                          width=160, dynamic_resizing=False).pack(side="left", padx=(0,16))
        ctk.CTkButton(tb, text="↺ Reset", command=self._reset_view, fg_color=BG_INPUT,
                      hover_color=BORDER, text_color=TEXT_PRI, font=FONT_BODY,
                      height=30, corner_radius=6, width=80).pack(side="left", padx=(0,8))
        self.row_count_lbl = ctk.CTkLabel(tb, text="", font=("Segoe UI",11), text_color=ACCENT)
        self.row_count_lbl.pack(side="left", padx=8)
        if self.on_edit_callback:
            ctk.CTkButton(tb, text="✏️  Open Editor", command=self._open_editor,
                          fg_color=ORANGE, hover_color="#c2410c", text_color="#0d1117",
                          font=FONT_BTN, height=30, corner_radius=6).pack(side="right", padx=12)
        ctk.CTkButton(tb, text="💾 Export View", command=self._export_view,
                      fg_color="#1d4ed8", hover_color="#1e3a8a", text_color=TEXT_PRI,
                      font=FONT_BTN, height=30, corner_radius=6).pack(side="right", padx=(0,8))
        self.stats_scroll = ctk.CTkScrollableFrame(self, fg_color=BG_CARD, corner_radius=0,
                                                    height=90, orientation="horizontal",
                                                    scrollbar_button_color=BORDER)
        self.stats_scroll.pack(fill="x"); self._build_stats_bar()
        tbl_frame = ctk.CTkFrame(self, fg_color=BG_DEEP, corner_radius=0); tbl_frame.pack(fill="both", expand=True)
        style = ttk.Style(); style.theme_use("default")
        style.configure("Viewer.Treeview", background=BG_CARD, foreground=TEXT_PRI,
                        fieldbackground=BG_CARD, rowheight=24, font=("Consolas",10))
        style.configure("Viewer.Treeview.Heading", background=BG_PANEL, foreground=ACCENT2,
                        font=("Segoe UI",10,"bold"), relief="flat")
        style.map("Viewer.Treeview", background=[("selected","#264f78")], foreground=[("selected",TEXT_PRI)])
        vsb = ttk.Scrollbar(tbl_frame, orient="vertical"); vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tbl_frame, orient="horizontal"); hsb.pack(side="bottom", fill="x")
        self.tree = ttk.Treeview(tbl_frame, style="Viewer.Treeview",
                                  yscrollcommand=vsb.set, xscrollcommand=hsb.set, selectmode="browse")
        self.tree.pack(fill="both", expand=True)
        vsb.configure(command=self.tree.yview); hsb.configure(command=self.tree.xview)
        self.tree.tag_configure("odd", background="#161b22")
        self.tree.tag_configure("even", background="#1c2230")
        self.tree.tag_configure("nan", foreground=DANGER)
        self._setup_columns(); self._populate_table()
        sb = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=26)
        sb.pack(fill="x", side="bottom"); sb.pack_propagate(False)
        self.status_lbl = ctk.CTkLabel(sb, text="Click a column header to sort", font=FONT_TINY, text_color=TEXT_SEC)
        self.status_lbl.pack(side="left", padx=12)

    def _build_stats_bar(self):
        for w in self.stats_scroll.winfo_children(): w.destroy()
        num_cols = self.df.select_dtypes(include=[np.number]).columns
        COLORS = [ACCENT2,SUCCESS,PURPLE,WARN,ORANGE,CYAN,PINK,ACCENT,DANGER]
        for i, col in enumerate(self.df.columns):
            cc = COLORS[i % len(COLORS)]
            f = ctk.CTkFrame(self.stats_scroll, fg_color=BG_INPUT, corner_radius=8,
                              border_width=1, border_color=cc, width=136)
            f.pack(side="left", padx=4, pady=6); f.pack_propagate(False)
            ctk.CTkLabel(f, text=str(col)[:16], font=("Segoe UI",9,"bold"), text_color=cc).pack(pady=(4,1), padx=4)
            n_valid=int(self.df[col].notna().sum()); n_nan=int(self.df[col].isna().sum())
            ctk.CTkLabel(f, text=f"N={n_valid}"+(f"  ⚠{n_nan}" if n_nan else ""),
                         font=("Segoe UI",8), text_color=WARN if n_nan else TEXT_SEC).pack()
            if col in num_cols:
                ctk.CTkLabel(f, text=f"μ={self.df[col].mean():.2f}  σ={self.df[col].std():.2f}",
                             font=("Consolas",8), text_color=ACCENT).pack()
                ctk.CTkLabel(f, text=f"[{self.df[col].min():.2f}, {self.df[col].max():.2f}]",
                             font=("Segoe UI",8), text_color=TEXT_SEC).pack(pady=(0,4))
            else:
                ctk.CTkLabel(f, text=f"{self.df[col].nunique()} unique",
                             font=("Segoe UI",8), text_color=TEXT_SEC).pack(pady=(0,4))

    def _setup_columns(self):
        self.tree["columns"] = list(self.display_df.columns); self.tree["show"] = "headings"
        self.tree.heading("#0", text="#"); self.tree.column("#0", width=46, minwidth=30, stretch=False, anchor="center")
        for col in self.display_df.columns:
            self.tree.heading(col, text=col, command=lambda c=col: self._sort_by(c))
            self.tree.column(col, width=max(80,min(160,len(str(col))*9+20)), minwidth=50, anchor="center")

    def _populate_table(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        for i, (_, row) in enumerate(self.display_df.iterrows()):
            vals=[]; has_nan=False
            for col in self.display_df.columns:
                v = row[col]
                if pd.isna(v): vals.append("NaN"); has_nan=True
                else:
                    try: vals.append(f"{float(v):.4f}".rstrip("0").rstrip("."))
                    except: vals.append(str(v)[:20])
            tags = ("nan",) if has_nan else (("odd" if i%2==0 else "even"),)
            self.tree.insert("", "end", text=str(i+1), values=vals, tags=tags)
        self.row_count_lbl.configure(text=f"Showing {len(self.display_df)} of {len(self.df)} rows")

    def _sort_by(self, col):
        if self._sort_col==col: self._sort_asc=not self._sort_asc
        else: self._sort_col=col; self._sort_asc=True
        arrow = "▲" if self._sort_asc else "▼"
        for c in self.display_df.columns: self.tree.heading(c, text=c if c!=col else f"{c} {arrow}")
        try: self.display_df=self.display_df.sort_values(col,ascending=self._sort_asc,key=lambda x:pd.to_numeric(x,errors="coerce"))
        except: self.display_df=self.display_df.sort_values(col,ascending=self._sort_asc)
        self._populate_table()

    def _on_search(self, *_):
        query=self.search_var.get().strip().lower(); col_filter=self.filter_col_var.get()
        if not query: self.display_df=self.df.copy()
        else:
            if col_filter=="All":
                mask=self.df.apply(lambda col: col.astype(str).str.lower().str.contains(query,na=False)).any(axis=1)
            else: mask=self.df[col_filter].astype(str).str.lower().str.contains(query,na=False)
            self.display_df=self.df[mask].copy()
        self._populate_table()

    def _reset_view(self):
        self.search_var.set(""); self.filter_col_var.set("All"); self.display_df=self.df.copy()
        self._sort_col=None; self._sort_asc=True
        for col in self.display_df.columns: self.tree.heading(col, text=col)
        self._populate_table()

    def _open_editor(self):
        if self.on_edit_callback: self.on_edit_callback()

    def _export_view(self):
        fp=filedialog.asksaveasfilename(defaultextension=".csv",filetypes=[("CSV","*.csv"),("Excel","*.xlsx")],title="Export current view",parent=self)
        if not fp: return
        try:
            self.display_df.to_excel(fp,index=False) if fp.endswith(".xlsx") else self.display_df.to_csv(fp,index=False)
            messagebox.showinfo("Exported",f"View saved:\n{fp}",parent=self)
        except Exception as e: messagebox.showerror("Error",str(e),parent=self)


# ─── Dataset Editor Window ────────────────────────────────────────────────────

class DatasetEditorWindow(ctk.CTkToplevel):
    MAX_UNDO = 50

    def __init__(self, master, df, on_apply, **kw):
        super().__init__(master, **kw)
        self.title("Dataset Editor"); self.geometry("1200x780"); self.configure(fg_color=BG_DEEP)
        self.df=df.copy(); self.on_apply=on_apply
        self._undo_stack=[]; self._redo_stack=[]
        self._edit_entry=None; self._edit_item=None; self._edit_col=None
        self.transient(master); self.lift(); self.focus_force(); self._build()

    def _push_undo(self):
        self._undo_stack.append(self.df.copy())
        if len(self._undo_stack)>self.MAX_UNDO: self._undo_stack.pop(0)
        self._redo_stack.clear(); self._update_undo_btns()

    def _update_undo_btns(self):
        self.undo_btn.configure(state="normal" if self._undo_stack else "disabled")
        self.redo_btn.configure(state="normal" if self._redo_stack else "disabled")

    def _build(self):
        hdr=ctk.CTkFrame(self,fg_color=ORANGE,corner_radius=0,height=58); hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr,text="  ✏️  Dataset Editor",font=("Segoe UI",20,"bold"),text_color="#0d1117").pack(side="left",padx=20)
        ctk.CTkLabel(hdr,text="Double-click any cell to edit  •  Changes do not auto-save",font=("Segoe UI",11),text_color="#7c3303").pack(side="left",padx=10)
        tb=ctk.CTkFrame(self,fg_color=BG_PANEL,corner_radius=0,height=50); tb.pack(fill="x"); tb.pack_propagate(False)
        self.undo_btn=ctk.CTkButton(tb,text="↩ Undo",command=self._undo,fg_color=BG_INPUT,hover_color=BORDER,text_color=TEXT_PRI,font=FONT_BODY,height=32,corner_radius=6,width=84,state="disabled")
        self.undo_btn.pack(side="left",padx=(12,4),pady=8)
        self.redo_btn=ctk.CTkButton(tb,text="↪ Redo",command=self._redo,fg_color=BG_INPUT,hover_color=BORDER,text_color=TEXT_PRI,font=FONT_BODY,height=32,corner_radius=6,width=84,state="disabled")
        self.redo_btn.pack(side="left",padx=(0,16),pady=8)
        ctk.CTkButton(tb,text="+ Row",command=self._add_row,fg_color=SUCCESS,hover_color="#16a34a",text_color="#0d1117",font=FONT_BODY,height=32,corner_radius=6,width=78).pack(side="left",padx=4)
        ctk.CTkButton(tb,text="− Row",command=self._del_row,fg_color=DANGER,hover_color="#b91c1c",text_color=TEXT_PRI,font=FONT_BODY,height=32,corner_radius=6,width=78).pack(side="left",padx=4)
        ctk.CTkButton(tb,text="+ Col",command=self._add_col,fg_color=ACCENT2,hover_color="#3b7ddd",text_color="#0d1117",font=FONT_BODY,height=32,corner_radius=6,width=78).pack(side="left",padx=4)
        ctk.CTkButton(tb,text="− Col",command=self._del_col,fg_color="#7c3aed",hover_color="#6d28d9",text_color=TEXT_PRI,font=FONT_BODY,height=32,corner_radius=6,width=78).pack(side="left",padx=4)
        ctk.CTkButton(tb,text="Rename Col",command=self._rename_col,fg_color=CYAN,hover_color="#0891b2",text_color="#0d1117",font=FONT_BODY,height=32,corner_radius=6,width=110).pack(side="left",padx=4)
        ctk.CTkButton(tb,text="Fill NaN→0",command=self._fill_nan,fg_color=WARN,hover_color="#d97706",text_color="#0d1117",font=FONT_BODY,height=32,corner_radius=6,width=100).pack(side="left",padx=4)
        ctk.CTkButton(tb,text="✓  Apply Changes",command=self._apply,fg_color=SUCCESS,hover_color="#16a34a",text_color="#0d1117",font=FONT_BTN,height=36,corner_radius=8).pack(side="right",padx=12)
        ctk.CTkButton(tb,text="Discard & Close",command=self.destroy,fg_color=DANGER,hover_color="#b91c1c",text_color=TEXT_PRI,font=FONT_BODY,height=36,corner_radius=8).pack(side="right",padx=(0,8))
        tbl_frame=ctk.CTkFrame(self,fg_color=BG_DEEP,corner_radius=0); tbl_frame.pack(fill="both",expand=True)
        style=ttk.Style()
        style.configure("Editor.Treeview",background=BG_CARD,foreground=TEXT_PRI,fieldbackground=BG_CARD,rowheight=26,font=("Consolas",10))
        style.configure("Editor.Treeview.Heading",background=BG_PANEL,foreground=ORANGE,font=("Segoe UI",10,"bold"),relief="flat")
        style.map("Editor.Treeview",background=[("selected","#3b1f0a")],foreground=[("selected",TEXT_PRI)])
        vsb=ttk.Scrollbar(tbl_frame,orient="vertical"); vsb.pack(side="right",fill="y")
        hsb=ttk.Scrollbar(tbl_frame,orient="horizontal"); hsb.pack(side="bottom",fill="x")
        self.tree=ttk.Treeview(tbl_frame,style="Editor.Treeview",yscrollcommand=vsb.set,xscrollcommand=hsb.set,selectmode="browse")
        self.tree.pack(fill="both",expand=True)
        vsb.configure(command=self.tree.yview); hsb.configure(command=self.tree.xview)
        self.tree.tag_configure("odd",background="#161b22"); self.tree.tag_configure("even",background="#1c2230")
        self.tree.tag_configure("nan",foreground=DANGER); self.tree.tag_configure("edited",foreground=SUCCESS)
        self.tree.bind("<Double-1>",self._on_double_click); self.tree.bind("<Escape>",lambda e:self._cancel_edit())
        sb=ctk.CTkFrame(self,fg_color=BG_CARD,corner_radius=0,height=26); sb.pack(fill="x",side="bottom"); sb.pack_propagate(False)
        self.status_lbl=ctk.CTkLabel(sb,text="Double-click any cell to edit it",font=FONT_TINY,text_color=TEXT_SEC); self.status_lbl.pack(side="left",padx=12)
        self.cell_lbl=ctk.CTkLabel(sb,text="",font=("Consolas",10),text_color=ACCENT2); self.cell_lbl.pack(side="right",padx=12)
        self._refresh_table()

    def _refresh_table(self):
        self._cancel_edit()
        for item in self.tree.get_children(): self.tree.delete(item)
        self.tree["columns"]=list(self.df.columns); self.tree["show"]="headings"
        for col in self.df.columns:
            self.tree.heading(col,text=col,command=lambda c=col:self._select_col(c))
            self.tree.column(col,width=max(80,min(160,len(str(col))*9+24)),minwidth=50,anchor="center")
        for i,(_, row) in enumerate(self.df.iterrows()):
            vals=[]; has_nan=False
            for col in self.df.columns:
                v=row[col]
                if pd.isna(v): vals.append("NaN"); has_nan=True
                else:
                    try: vals.append(f"{float(v):.4f}".rstrip("0").rstrip("."))
                    except: vals.append(str(v)[:24])
            tag="nan" if has_nan else ("odd" if i%2==0 else "even")
            self.tree.insert("","end",iid=str(i),text=str(i+1),values=vals,tags=(tag,))
        nan_count=int(self.df.isna().sum().sum())
        self.status_lbl.configure(text=f"  {len(self.df)} rows × {len(self.df.columns)} cols{'  ⚠'+str(nan_count)+' NaN' if nan_count else ''}")

    def _on_double_click(self, event):
        region=self.tree.identify("region",event.x,event.y)
        if region!="cell": return
        col_id=self.tree.identify_column(event.x); item_id=self.tree.identify_row(event.y)
        if not item_id or not col_id: return
        col_idx=int(col_id.replace("#",""))-1
        if col_idx<0 or col_idx>=len(self.df.columns): return
        col_name=self.df.columns[col_idx]; row_idx=int(item_id)
        current=self.df.iloc[row_idx,col_idx]; cur_str="" if pd.isna(current) else str(current)
        bbox=self.tree.bbox(item_id,col_id)
        if not bbox: return
        x,y,w,h=bbox; self._cancel_edit()
        self._edit_item=item_id; self._edit_col=col_name; self._edit_row=row_idx
        self._edit_entry=tk.Entry(self.tree,bg=BG_INPUT,fg=TEXT_PRI,insertbackground=TEXT_PRI,
                                   font=("Consolas",10),relief="flat",bd=1,highlightthickness=2,highlightcolor=ORANGE)
        self._edit_entry.insert(0,cur_str); self._edit_entry.select_range(0,"end")
        self._edit_entry.place(x=x,y=y,width=w,height=h); self._edit_entry.focus_set()
        self._edit_entry.bind("<Return>",self._commit_edit); self._edit_entry.bind("<Tab>",self._commit_edit)
        self._edit_entry.bind("<Escape>",lambda e:self._cancel_edit()); self._edit_entry.bind("<FocusOut>",self._commit_edit)
        self.cell_lbl.configure(text=f"Editing  [{row_idx+1}, {col_name}]  current: {cur_str}")

    def _commit_edit(self, event=None):
        if self._edit_entry is None: return
        new_val_str=self._edit_entry.get().strip(); row_idx=self._edit_row; col_name=self._edit_col
        if new_val_str in ("","nan","NaN","NA","N/A","."): new_val=np.nan
        else:
            try: new_val=float(new_val_str)
            except: new_val=new_val_str
        old_val=self.df.iloc[row_idx,self.df.columns.get_loc(col_name)]
        self._push_undo(); self.df.iloc[row_idx,self.df.columns.get_loc(col_name)]=new_val
        self._cancel_edit()
        col_idx=list(self.df.columns).index(col_name); vals=list(self.tree.item(str(row_idx),"values"))
        if pd.isna(new_val): vals[col_idx]="NaN"; self.tree.item(str(row_idx),values=vals,tags=("nan",))
        else:
            try: vals[col_idx]=f"{float(new_val):.4f}".rstrip("0").rstrip(".")
            except: vals[col_idx]=str(new_val)[:24]
            self.tree.item(str(row_idx),values=vals,tags=("edited",))
        self.cell_lbl.configure(text=f"  ✓  [{row_idx+1}, {col_name}]  {old_val!r} → {new_val!r}")

    def _cancel_edit(self, event=None):
        if self._edit_entry:
            try: self._edit_entry.destroy()
            except: pass
            self._edit_entry=None

    def _undo(self):
        if not self._undo_stack: return
        self._redo_stack.append(self.df.copy()); self.df=self._undo_stack.pop()
        self._update_undo_btns(); self._refresh_table(); self.status_lbl.configure(text="Undo applied")

    def _redo(self):
        if not self._redo_stack: return
        self._undo_stack.append(self.df.copy()); self.df=self._redo_stack.pop()
        self._update_undo_btns(); self._refresh_table(); self.status_lbl.configure(text="Redo applied")

    def _add_row(self):
        self._push_undo()
        self.df=pd.concat([self.df,pd.DataFrame([[np.nan]*len(self.df.columns)],columns=self.df.columns)],ignore_index=True)
        self._refresh_table()

    def _del_row(self):
        sel=self.tree.selection()
        if not sel: messagebox.showwarning("No selection","Select a row first.",parent=self); return
        idx=int(sel[0])
        if messagebox.askyesno("Delete Row",f"Delete row {idx+1}?",parent=self):
            self._push_undo(); self.df=self.df.drop(index=idx).reset_index(drop=True); self._refresh_table()

    def _add_col(self):
        name=simpledialog.askstring("New Column","Enter column name:",parent=self)
        if not name or not name.strip(): return
        name=name.strip()
        if name in self.df.columns: messagebox.showwarning("Duplicate",f"Column '{name}' already exists.",parent=self); return
        self._push_undo(); self.df[name]=np.nan; self._refresh_table()

    def _del_col(self):
        col_name=simpledialog.askstring("Delete Column",f"Column name to delete (from: {', '.join(self.df.columns)}):",parent=self)
        if not col_name or col_name not in self.df.columns: return
        if len(self.df.columns)<=2: messagebox.showwarning("Too Few","Need at least 2 columns.",parent=self); return
        if messagebox.askyesno("Delete Column",f"Delete column '{col_name}'?",parent=self):
            self._push_undo(); self.df=self.df.drop(columns=[col_name]); self._refresh_table()

    def _rename_col(self):
        old=simpledialog.askstring("Rename Column",f"Current name (from: {', '.join(self.df.columns)}):",parent=self)
        if not old or old not in self.df.columns: messagebox.showwarning("Not found",f"'{old}' is not a column.",parent=self); return
        new=simpledialog.askstring("Rename Column",f"New name for '{old}':",parent=self)
        if not new or not new.strip(): return
        new=new.strip()
        if new in self.df.columns: messagebox.showwarning("Duplicate",f"'{new}' already exists.",parent=self); return
        self._push_undo(); self.df=self.df.rename(columns={old:new}); self._refresh_table()

    def _fill_nan(self):
        n=int(self.df.isna().sum().sum())
        if n==0: messagebox.showinfo("No NaN","No missing values found.",parent=self); return
        if messagebox.askyesno("Fill NaN",f"Replace {n} NaN value(s) with 0?",parent=self):
            self._push_undo(); self.df=self.df.fillna(0); self._refresh_table()

    def _select_col(self, col): self.status_lbl.configure(text=f"Column selected: {col}")

    def _apply(self):
        if messagebox.askyesno("Apply Changes","Apply edited dataset to the main application?\nThis will replace the current working data.",parent=self):
            self.on_apply(self.df.copy())
            messagebox.showinfo("Applied",f"Dataset updated:\n{len(self.df)} rows × {len(self.df.columns)} cols",parent=self)
            self.destroy()


# ─── Variable Descriptions Manager ───────────────────────────────────────────

class VariableDescriptionsWindow(ctk.CTkToplevel):
    def __init__(self, master, columns, existing_descs, on_save, **kw):
        super().__init__(master, **kw)
        self.title("Variable Descriptions"); self.geometry("760x640"); self.configure(fg_color=BG_DEEP)
        self.columns=columns; self.on_save=on_save; self.label_entries={}; self.desc_entries={}
        self.transient(master); self.lift(); self.focus_force()
        self.attributes("-topmost",True); self._build(existing_descs)

    def _build(self, existing):
        hdr=ctk.CTkFrame(self,fg_color=PURPLE,corner_radius=0,height=56); hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr,text="  📝  Variable Descriptions",font=("Segoe UI",18,"bold"),text_color="#0d1117").pack(side="left",padx=20)
        ctk.CTkLabel(hdr,text="Descriptions appear in the PDF report",font=("Segoe UI",11),text_color="#3b0764").pack(side="right",padx=20)
        hint=ctk.CTkFrame(self,fg_color=BG_PANEL,corner_radius=0,height=32); hint.pack(fill="x"); hint.pack_propagate(False)
        ctk.CTkLabel(hint,text="  Label = short display name  |  Description = full explanation for PDF",font=FONT_TINY,text_color=TEXT_SEC).pack(side="left",padx=12)
        scroll=ctk.CTkScrollableFrame(self,fg_color=BG_DEEP,scrollbar_button_color=BORDER); scroll.pack(fill="both",expand=True,padx=12,pady=8)
        COLORS=[ACCENT2,SUCCESS,PURPLE,WARN,ORANGE,CYAN,PINK,ACCENT,DANGER]
        for i,col in enumerate(self.columns):
            cc=COLORS[i%len(COLORS)]; ex=existing.get(col,{})
            cf=ctk.CTkFrame(scroll,fg_color=BG_CARD,corner_radius=10,border_width=1,border_color=cc); cf.pack(fill="x",pady=5)
            strip=ctk.CTkFrame(cf,fg_color=cc,corner_radius=0,height=32); strip.pack(fill="x"); strip.pack_propagate(False)
            ctk.CTkLabel(strip,text=f"  {col}",font=("Segoe UI",12,"bold"),text_color="#0d1117").pack(side="left",padx=8)
            body=ctk.CTkFrame(cf,fg_color="transparent"); body.pack(fill="x",padx=12,pady=8)
            lrow=ctk.CTkFrame(body,fg_color="transparent"); lrow.pack(fill="x",pady=(0,4))
            ctk.CTkLabel(lrow,text="Display Label:",font=FONT_BODY,text_color=TEXT_SEC,width=110).pack(side="left")
            le=ctk.CTkEntry(lrow,placeholder_text="e.g. Reading Interest",fg_color=BG_INPUT,border_color=BORDER,text_color=TEXT_PRI,border_width=1,corner_radius=6,height=30,font=FONT_BODY)
            le.pack(side="left",fill="x",expand=True,padx=(8,0))
            if ex.get("label"): le.insert(0,ex["label"])
            self.label_entries[col]=le
            drow=ctk.CTkFrame(body,fg_color="transparent"); drow.pack(fill="x")
            ctk.CTkLabel(drow,text="Description:",font=FONT_BODY,text_color=TEXT_SEC,width=110).pack(side="left",anchor="n",pady=4)
            de=ctk.CTkTextbox(drow,height=56,fg_color=BG_INPUT,text_color=TEXT_PRI,border_width=1,border_color=BORDER,font=FONT_BODY,corner_radius=6)
            de.pack(side="left",fill="x",expand=True,padx=(8,0))
            if ex.get("desc"): de.insert("1.0",ex["desc"])
            self.desc_entries[col]=de
        brow=ctk.CTkFrame(self,fg_color="transparent"); brow.pack(fill="x",padx=12,pady=(0,12))
        ctk.CTkButton(brow,text="✓  Save Descriptions",command=self._save,fg_color=PURPLE,hover_color="#7c3aed",text_color=TEXT_PRI,font=FONT_BTN,height=40,corner_radius=8).pack(side="left",padx=(0,8))
        ctk.CTkButton(brow,text="✕  Clear All",command=self._clear_all,fg_color=DANGER,hover_color="#b91c1c",text_color=TEXT_PRI,font=FONT_BODY,height=40,corner_radius=8).pack(side="left",padx=(0,8))
        ctk.CTkButton(brow,text="Close",command=self.destroy,fg_color=BG_PANEL,hover_color=BORDER,font=FONT_BODY,height=40,corner_radius=8).pack(side="right")

    def _save(self):
        result={}
        for col in self.columns:
            lbl=self.label_entries[col].get().strip(); dsc=self.desc_entries[col].get("1.0","end-1c").strip()
            result[col]={"label":lbl,"desc":dsc}
        self.on_save(result); messagebox.showinfo("Saved",f"Descriptions saved for {len(result)} variable(s).",parent=self); self.destroy()

    def _clear_all(self):
        if messagebox.askyesno("Clear All","Clear all labels and descriptions?",parent=self):
            for col in self.columns: self.label_entries[col].delete(0,"end"); self.desc_entries[col].delete("1.0","end")


# ─── [NEW] Free Pair Selector Window ─────────────────────────────────────────

class FreePairSelectorWindow(ctk.CTkToplevel):
    """
    Freely build any combination of variable pairs to correlate.
    Each row = one pair (Variable A ↔ Variable B).
    Supports: Add Pair, Add All Pairs, Remove rows, Clear All.
    Validates self-pairs and duplicates before confirming.
    """

    def __init__(self, master, columns, existing_pairs, on_save, **kw):
        super().__init__(master, **kw)
        self.title("Free Pair Selector  —  Choose Any Pairs")
        self.geometry("680x620")
        self.configure(fg_color=BG_DEEP)
        self.columns   = columns
        self.on_save   = on_save
        self.pair_rows = []   # list of (StringVar A, StringVar B, frame)
        self.transient(master)
        self.lift(); self.focus_force()
        self.attributes("-topmost", True)
        self._build(existing_pairs)

    def _build(self, existing_pairs):
        # Header
        hdr = ctk.CTkFrame(self, fg_color="#0f766e", corner_radius=0, height=60)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr, text="  🔗  Free Pair Selector",
                     font=("Segoe UI", 20, "bold"), text_color="#0d1117").pack(side="left", padx=20)
        ctk.CTkLabel(hdr, text="Choose any combination of variable pairs",
                     font=("Segoe UI", 11), text_color="#134e4a").pack(side="right", padx=20)

        # Toolbar
        tb = ctk.CTkFrame(self, fg_color=BG_PANEL, corner_radius=0, height=52)
        tb.pack(fill="x"); tb.pack_propagate(False)
        ctk.CTkButton(tb, text="＋  Add Pair", command=self._add_pair_row,
                      fg_color=SUCCESS, hover_color="#16a34a", text_color="#0d1117",
                      font=FONT_BTN, height=34, corner_radius=6, width=130).pack(side="left", padx=(12,6), pady=8)
        ctk.CTkButton(tb, text="⊕  Add All Pairs", command=self._add_all_pairs,
                      fg_color=ACCENT2, hover_color="#3b7ddd", text_color="#0d1117",
                      font=FONT_BTN, height=34, corner_radius=6, width=148).pack(side="left", padx=(0,6))
        ctk.CTkButton(tb, text="✕  Clear All", command=self._clear_all,
                      fg_color=DANGER, hover_color="#b91c1c", text_color=TEXT_PRI,
                      font=FONT_BODY, height=34, corner_radius=6, width=110).pack(side="left", padx=(0,6))
        self.count_lbl = ctk.CTkLabel(tb, text="0 pair(s)", font=("Segoe UI",12,"bold"), text_color=ACCENT)
        self.count_lbl.pack(side="right", padx=16)

        # Column header row
        col_hdr = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=32)
        col_hdr.pack(fill="x"); col_hdr.pack_propagate(False)
        ctk.CTkLabel(col_hdr, text="  #", font=("Segoe UI",10,"bold"), text_color=TEXT_SEC, width=36).pack(side="left", padx=4)
        ctk.CTkLabel(col_hdr, text="Variable A", font=("Segoe UI",10,"bold"), text_color=ACCENT2, width=200).pack(side="left", padx=4)
        ctk.CTkLabel(col_hdr, text="  ↔  ", font=("Segoe UI",13,"bold"), text_color=TEXT_SEC, width=40).pack(side="left")
        ctk.CTkLabel(col_hdr, text="Variable B", font=("Segoe UI",10,"bold"), text_color=SUCCESS, width=200).pack(side="left", padx=4)
        ctk.CTkLabel(col_hdr, text="Del", font=("Segoe UI",10,"bold"), text_color=DANGER, width=48).pack(side="left", padx=4)

        # Scrollable pair rows
        self.scroll = ctk.CTkScrollableFrame(self, fg_color=BG_DEEP, scrollbar_button_color=BORDER)
        self.scroll.pack(fill="both", expand=True)

        # Restore existing pairs or start with one blank row
        if existing_pairs:
            for a, b in existing_pairs:
                self._add_pair_row(default_a=a, default_b=b)
        else:
            self._add_pair_row()

        # Bottom buttons
        brow = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=58)
        brow.pack(fill="x", side="bottom"); brow.pack_propagate(False)
        ctk.CTkButton(brow, text="✓  Confirm & Use These Pairs", command=self._confirm,
                      fg_color="#0f766e", hover_color="#134e4a", text_color=TEXT_PRI,
                      font=FONT_BTN, height=42, corner_radius=8, width=260).pack(side="left", padx=14, pady=8)
        ctk.CTkButton(brow, text="Cancel", command=self.destroy,
                      fg_color=BG_PANEL, hover_color=BORDER, text_color=TEXT_PRI,
                      font=FONT_BODY, height=42, corner_radius=8, width=100).pack(side="right", padx=14)

    def _add_pair_row(self, default_a=None, default_b=None):
        idx   = len(self.pair_rows) + 1
        COLORS = [ACCENT2,SUCCESS,PURPLE,WARN,ORANGE,CYAN,PINK,ACCENT,DANGER]
        cc    = COLORS[(idx-1) % len(COLORS)]

        rf = ctk.CTkFrame(self.scroll, fg_color=BG_CARD, corner_radius=8, border_width=1, border_color=cc)
        rf.pack(fill="x", pady=3, padx=4)

        ctk.CTkLabel(rf, text=f" {idx} ", font=("Segoe UI",11,"bold"),
                     text_color=cc, fg_color=BG_DEEP, corner_radius=4,
                     width=32).pack(side="left", padx=(6,4), pady=8)

        va = tk.StringVar(value=default_a or self.columns[0])
        ctk.CTkOptionMenu(rf, variable=va, values=self.columns,
                          fg_color=BG_INPUT, button_color=ACCENT2, button_hover_color="#3b7ddd",
                          text_color=TEXT_PRI, dropdown_fg_color=BG_PANEL, dropdown_text_color=TEXT_PRI,
                          font=FONT_BODY, height=34, corner_radius=6, width=200).pack(side="left", padx=4, pady=6)

        ctk.CTkLabel(rf, text="↔", font=("Segoe UI",14,"bold"), text_color=cc, width=36).pack(side="left")

        vb_default = default_b or (self.columns[1] if len(self.columns) > 1 else self.columns[0])
        vb = tk.StringVar(value=vb_default)
        ctk.CTkOptionMenu(rf, variable=vb, values=self.columns,
                          fg_color=BG_INPUT, button_color=SUCCESS, button_hover_color="#16a34a",
                          text_color=TEXT_PRI, dropdown_fg_color=BG_PANEL, dropdown_text_color=TEXT_PRI,
                          font=FONT_BODY, height=34, corner_radius=6, width=200).pack(side="left", padx=4, pady=6)

        def _remove(frame=rf):
            self.pair_rows = [r for r in self.pair_rows if r[2] is not frame]
            frame.destroy(); self._update_count()

        ctk.CTkButton(rf, text="✕", command=_remove,
                      fg_color=DANGER, hover_color="#b91c1c", text_color=TEXT_PRI,
                      font=("Segoe UI",12,"bold"), height=34, width=44,
                      corner_radius=6).pack(side="left", padx=(6,8), pady=6)

        self.pair_rows.append((va, vb, rf))
        self._update_count()

    def _add_all_pairs(self):
        existing_keys = set()
        for va, vb, _ in self.pair_rows:
            a, b = va.get(), vb.get()
            if a != b: existing_keys.add((min(a,b), max(a,b)))
        added = 0
        for i in range(len(self.columns)):
            for j in range(i+1, len(self.columns)):
                a, b = self.columns[i], self.columns[j]
                key = (min(a,b), max(a,b))
                if key not in existing_keys:
                    self._add_pair_row(default_a=a, default_b=b)
                    existing_keys.add(key); added += 1
        if added == 0:
            messagebox.showinfo("All Pairs Already Added", "Every possible pair is already in the list.", parent=self)

    def _clear_all(self):
        if not self.pair_rows: return
        if messagebox.askyesno("Clear All", f"Remove all {len(self.pair_rows)} pair(s)?", parent=self):
            for _, _, frame in self.pair_rows: frame.destroy()
            self.pair_rows.clear(); self._update_count()

    def _update_count(self):
        n = len(self.pair_rows)
        self.count_lbl.configure(text=f"{n} pair{'s' if n!=1 else ''}",
                                  text_color=ACCENT if n>0 else DANGER)

    def _confirm(self):
        if not self.pair_rows:
            messagebox.showwarning("No Pairs", "Add at least one pair before confirming.", parent=self); return
        pairs=[]; seen=set(); warnings_list=[]
        for i, (va, vb, _) in enumerate(self.pair_rows, start=1):
            a, b = va.get().strip(), vb.get().strip()
            if a == b:
                warnings_list.append(f"Row {i}: Variable A and B are the same ({a}) — skipped."); continue
            key = (min(a,b), max(a,b))
            if key in seen:
                warnings_list.append(f"Row {i}: Duplicate pair ({a} ↔ {b}) — skipped."); continue
            seen.add(key); pairs.append((a, b))
        if warnings_list:
            msg = "\n".join(warnings_list)
            if pairs: msg += f"\n\n✓ {len(pairs)} valid pair(s) will be used."
            messagebox.showwarning("Pair Issues", msg, parent=self)
        if not pairs:
            messagebox.showerror("No Valid Pairs", "All rows were skipped.\nFix the pairs and try again.", parent=self); return
        self.on_save(pairs); self.destroy()


# ─── Scatter Plot Window ──────────────────────────────────────────────────────

class ScatterPlotWindow(ctk.CTkToplevel):
    def __init__(self, master, results, df, **kw):
        super().__init__(master, **kw)
        self.title("Scatter Plot Viewer"); self.geometry("960x660"); self.configure(fg_color=BG_DEEP)
        self.results=results; self.df=df
        self.transient(master); self.lift(); self.focus_force(); self._build()

    def _build(self):
        hdr=ctk.CTkFrame(self,fg_color=BG_CARD,corner_radius=0,height=52); hdr.pack(fill="x"); hdr.pack_propagate(False)
        mi=self.results["method_info"]
        ctk.CTkLabel(hdr,text=f"  📉  Scatter Plot Viewer  —  {mi['label']}  ({mi['symbol']})",font=("Segoe UI",15,"bold"),text_color=TEXT_PRI).pack(side="left",padx=20)
        ctk.CTkButton(hdr,text="💾  Save All Plots",command=self._save_all,fg_color=ACCENT2,hover_color="#3b7ddd",text_color="#0d1117",font=FONT_BTN,height=34,corner_radius=6).pack(side="right",padx=16)
        style=ttk.Style(); style.theme_use("default")
        style.configure("TNotebook",background=BG_DEEP,borderwidth=0)
        style.configure("TNotebook.Tab",background=BG_PANEL,foreground=TEXT_SEC,padding=[14,6],font=("Segoe UI",11))
        style.map("TNotebook.Tab",background=[("selected",BG_CARD)],foreground=[("selected",TEXT_PRI)])
        self.nb=ttk.Notebook(self); self.nb.pack(fill="both",expand=True,padx=10,pady=10)
        self._figs=[]; pairs=self.results["pairs"]
        PAL=[mi["color"],ACCENT,SUCCESS,PURPLE,WARN,ORANGE,CYAN,PINK]
        for idx,pair in enumerate(pairs):
            tab=ctk.CTkFrame(self.nb,fg_color=BG_CARD); self.nb.add(tab,text=f"  {pair['x']} ↔ {pair['y']}  ")
            self._build_scatter_tab(tab,pair,PAL[idx%len(PAL)])

    def _hex_to_rgb(self, hex_color):
        h=hex_color.lstrip("#"); return tuple(int(h[i:i+2],16)/255 for i in (0,2,4))

    def _build_scatter_tab(self, parent, pair, color_hex):
        fig=plt.Figure(figsize=(8.5,5.2),dpi=100,facecolor="#0d1117"); self._figs.append(fig)
        ax=fig.add_subplot(111); ax.set_facecolor("#161b22")
        for spine in ax.spines.values(): spine.set_color("#30363d")
        ax.tick_params(colors=TEXT_SEC,labelsize=10); ax.xaxis.label.set_color(TEXT_SEC); ax.yaxis.label.set_color(TEXT_SEC)
        sub=self.df[[pair["x"],pair["y"]]].dropna(); x,y=sub[pair["x"]].values,sub[pair["y"]].values
        rgb=self._hex_to_rgb(color_hex); dot_c=(*rgb,0.65); line_c=(*rgb,1.0)
        ax.scatter(x,y,color=dot_c,edgecolors=(*rgb,0.9),linewidths=0.6,s=52,zorder=3)
        if len(x)>=2:
            m,b=np.polyfit(x,y,1); x_line=np.linspace(x.min(),x.max(),200); y_line=m*x_line+b
            ax.plot(x_line,y_line,color=line_c,linewidth=2.0,zorder=4,label=f"y = {m:.3f}x + {b:.3f}")
            n=len(x); x_mean=x.mean()
            se_fit=np.std(y-(m*x+b),ddof=2)*np.sqrt(1/n+(x_line-x_mean)**2/((n-1)*np.var(x,ddof=1)))
            ax.fill_between(x_line,y_line-1.96*se_fit,y_line+1.96*se_fit,alpha=0.15,color=line_c,zorder=2)
        sym=self.results["method_info"]["symbol"]; r_val=pair["r"]; p_val=pair["p"]; n_val=pair["n"]
        ci_lo=pair.get("ci_lower"); ci_hi=pair.get("ci_upper")
        ci_str=f"95% CI [{ci_lo:.3f}, {ci_hi:.3f}]" if ci_lo is not None else ""
        ann=(f"{sym} = {r_val:.4f}  {pair['stars']}\np = {p_val:.4f}\nEffect: {pair['effect']}\nN = {n_val}\n{ci_str}")
        ax.text(0.03,0.97,ann,transform=ax.transAxes,fontsize=9.5,verticalalignment="top",fontfamily="monospace",
                color=TEXT_PRI,bbox=dict(boxstyle="round,pad=0.5",facecolor="#1c2230",alpha=0.88,edgecolor=color_hex,linewidth=1.2))
        ax.set_xlabel(pair["x"],fontsize=11); ax.set_ylabel(pair["y"],fontsize=11)
        ax.set_title(f"{pair['x']}  ↔  {pair['y']}",color=TEXT_PRI,fontsize=13,pad=10)
        ax.grid(True,color="#30363d",linewidth=0.5,alpha=0.6)
        if len(x)>=2: ax.legend(fontsize=9,facecolor="#1c2230",edgecolor="#30363d",labelcolor=TEXT_SEC)
        fig.tight_layout(pad=1.8)
        canvas=FigureCanvasTkAgg(fig,master=parent); canvas.draw()
        canvas.get_tk_widget().pack(fill="both",expand=True,padx=6,pady=6)
        tb_frame=ctk.CTkFrame(parent,fg_color=BG_PANEL,corner_radius=0); tb_frame.pack(fill="x")
        toolbar=NavigationToolbar2Tk(canvas,tb_frame); toolbar.config(background=BG_PANEL)
        for btn in toolbar.winfo_children():
            try: btn.config(background=BG_PANEL,foreground=TEXT_SEC)
            except: pass
        toolbar.update()

    def _save_all(self):
        folder=filedialog.askdirectory(title="Select folder to save plots")
        if not folder: return
        saved=0
        for idx,(fig,pair) in enumerate(zip(self._figs,self.results["pairs"])):
            fname=os.path.join(folder,f"scatter_{pair['x']}_vs_{pair['y']}_{idx+1}.png")
            try: fig.savefig(fname,dpi=150,bbox_inches="tight",facecolor=fig.get_facecolor()); saved+=1
            except Exception as e: messagebox.showwarning("Save Error",f"Could not save plot {idx+1}:\n{e}",parent=self)
        if saved: messagebox.showinfo("Saved",f"Saved {saved} plot(s) to:\n{folder}",parent=self)


# ─── Variable Selector Popup ──────────────────────────────────────────────────

class VariableSelectorWindow(ctk.CTkToplevel):
    def __init__(self, master, columns, method_key, on_select, **kw):
        super().__init__(master, **kw)
        self.title("Select Variables"); self.geometry("520x640"); self.configure(fg_color=BG_DEEP)
        self.columns=columns; self.method_key=method_key; self.on_select=on_select
        self.check_vars={}; self.ctrl_vars={}; self.needs_ctrl=method_key in ("partial","semi_partial")
        self.transient(master); self.lift(); self.focus_force(); self.attributes("-topmost",True); self._build()

    def _build(self):
        hdr=ctk.CTkFrame(self,fg_color=BG_CARD,corner_radius=0,height=56); hdr.pack(fill="x"); hdr.pack_propagate(False)
        mi=METHODS[self.method_key]
        ctk.CTkLabel(hdr,text=f"  ⚙  Select Variables — {mi['label']}",font=("Segoe UI",14,"bold"),text_color=TEXT_PRI).pack(side="left",padx=20)
        body=ctk.CTkFrame(self,fg_color=BG_DEEP); body.pack(fill="both",expand=True,padx=16,pady=16)
        av=card(body,"Analysis Variables  (≥ 2)"); av.pack(fill="both",expand=True,pady=(0,8))
        sc=ctk.CTkScrollableFrame(av,fg_color=BG_CARD,scrollbar_button_color=BORDER); sc.pack(fill="both",expand=True,padx=12,pady=(4,8))
        for col in self.columns:
            v=ctk.BooleanVar(value=True); self.check_vars[col]=v
            ctk.CTkCheckBox(sc,text=col,variable=v,fg_color=mi["color"],hover_color=BG_PANEL,text_color=TEXT_PRI,font=FONT_BODY).pack(anchor="w",padx=8,pady=3)
        if self.needs_ctrl:
            cc=card(body,"Control Variables  (≥ 1)"); cc.pack(fill="x",pady=(0,8))
            cs=ctk.CTkScrollableFrame(cc,fg_color=BG_CARD,height=110,scrollbar_button_color=BORDER); cs.pack(fill="x",padx=12,pady=(4,8))
            for col in self.columns:
                v=ctk.BooleanVar(value=False); self.ctrl_vars[col]=v
                ctk.CTkCheckBox(cs,text=col,variable=v,fg_color=PURPLE,hover_color="#7c3aed",text_color=TEXT_PRI,font=FONT_BODY).pack(anchor="w",padx=8,pady=2)
        br=ctk.CTkFrame(body,fg_color="transparent"); br.pack(fill="x",pady=(4,0))
        ctk.CTkButton(br,text="✓  Confirm",command=self._confirm,fg_color=mi["color"],hover_color=BG_PANEL,text_color="#0d1117",font=FONT_BTN,height=40,corner_radius=8).pack(side="left",padx=(0,8))
        ctk.CTkButton(br,text="All",command=lambda:[v.set(True) for v in self.check_vars.values()],fg_color=SUCCESS,hover_color="#16a34a",font=FONT_BODY,height=40,corner_radius=8).pack(side="left",padx=(0,8))
        ctk.CTkButton(br,text="None",command=lambda:[v.set(False) for v in self.check_vars.values()],fg_color=DANGER,hover_color="#b91c1c",font=FONT_BODY,height=40,corner_radius=8).pack(side="left",padx=(0,8))
        ctk.CTkButton(br,text="Close",command=self.destroy,fg_color=BG_PANEL,hover_color=BORDER,font=FONT_BODY,height=40,corner_radius=8).pack(side="right")

    def _confirm(self):
        chosen=[c for c,v in self.check_vars.items() if v.get()]
        controls=[c for c,v in self.ctrl_vars.items() if v.get()] if self.needs_ctrl else []
        if len(chosen)<2: messagebox.showwarning("Too Few","Select ≥ 2 analysis variables.",parent=self); return
        if self.needs_ctrl and not controls: messagebox.showwarning("No Controls","Select ≥ 1 control variable.",parent=self); return
        if set(chosen)&set(controls): messagebox.showwarning("Overlap","A variable cannot be in both groups.",parent=self); return
        self.on_select(chosen,controls); self.destroy()


# ─── Data Table Widget ────────────────────────────────────────────────────────

class DataTableFrame(ctk.CTkFrame):
    def __init__(self, master, **kw):
        super().__init__(master, fg_color=BG_CARD, corner_radius=0, **kw)
        self._df = None
        self.info_bar = ctk.CTkScrollableFrame(self, fg_color=BG_PANEL, corner_radius=6,
                                                height=82, orientation="horizontal", scrollbar_button_color=BORDER)
        self.info_bar.pack(fill="x", padx=6, pady=(6,4))
        ctk.CTkLabel(self.info_bar, text="No variables loaded", font=FONT_TINY, text_color=TEXT_SEC).pack(pady=10)
        tbl_frame=ctk.CTkFrame(self,fg_color=BG_CARD,corner_radius=0); tbl_frame.pack(fill="both",expand=True,padx=6,pady=(0,6))
        tbl_frame.columnconfigure(0,weight=1); tbl_frame.rowconfigure(0,weight=1)
        self._canvas=tk.Canvas(tbl_frame,bg=BG_CARD,highlightthickness=0); self._canvas.grid(row=0,column=0,sticky="nsew")
        vsb=tk.Scrollbar(tbl_frame,orient="vertical",command=self._canvas.yview); vsb.grid(row=0,column=1,sticky="ns")
        hsb=tk.Scrollbar(tbl_frame,orient="horizontal",command=self._canvas.xview); hsb.grid(row=1,column=0,sticky="ew")
        self._canvas.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
        self._inner=tk.Frame(self._canvas,bg=BG_CARD)
        self._canvas_window=self._canvas.create_window((0,0),window=self._inner,anchor="nw")
        self._inner.bind("<Configure>",self._on_inner_configure)
        self._canvas.bind("<Configure>",self._on_canvas_configure)
        self._canvas.bind_all("<MouseWheel>",self._on_mousewheel)
        self._canvas.bind_all("<Shift-MouseWheel>",self._on_shift_mousewheel)

    def _on_inner_configure(self, event=None): self._canvas.configure(scrollregion=self._canvas.bbox("all"))
    def _on_canvas_configure(self, event=None): self._canvas.itemconfig(self._canvas_window,width=max(event.width,self._inner.winfo_reqwidth()))
    def _on_mousewheel(self, event):
        try: self._canvas.yview_scroll(int(-1*(event.delta/120)),"units")
        except: pass
    def _on_shift_mousewheel(self, event):
        try: self._canvas.xview_scroll(int(-1*(event.delta/120)),"units")
        except: pass

    def display_data(self, df, selected_cols=None, max_rows=200):
        self._df=df
        for w in self.info_bar.winfo_children(): w.destroy()
        for w in self._inner.winfo_children(): w.destroy()
        if df is None or df.empty:
            ctk.CTkLabel(self.info_bar,text="No data loaded",font=FONT_TINY,text_color=TEXT_SEC).pack(pady=16)
            tk.Label(self._inner,text="No data loaded",bg=BG_CARD,fg=TEXT_SEC,font=("Segoe UI",12)).grid(row=0,column=0,padx=40,pady=40)
            return
        num_df=df.select_dtypes(include=[np.number]); hi_cols=set(selected_cols or [])
        COLORS=[ACCENT2,SUCCESS,PURPLE,WARN,ORANGE,CYAN,PINK,ACCENT,DANGER]
        for i,col in enumerate(df.columns):
            is_num=col in num_df.columns; is_sel=col in hi_cols
            cc=COLORS[i%len(COLORS)] if is_sel else (ACCENT if is_num else TEXT_SEC)
            cf=ctk.CTkFrame(self.info_bar,fg_color=BG_INPUT,corner_radius=8,border_width=1,border_color=cc,width=120)
            cf.pack(side="left",padx=4,pady=4); cf.pack_propagate(False)
            ctk.CTkLabel(cf,text=str(col)[:14],font=("Segoe UI",10,"bold"),text_color=cc).pack(pady=(4,1))
            n_valid=int(df[col].notna().sum()); n_nan=int(df[col].isna().sum())
            ctk.CTkLabel(cf,text=f"N={n_valid}{'  ⚠'+str(n_nan) if n_nan else ''}",font=("Segoe UI",9),text_color=WARN if n_nan else TEXT_SEC).pack(pady=(2,0))
            if is_num:
                ctk.CTkLabel(cf,text=f"μ={df[col].mean():.2f}",font=("Consolas",9,"bold"),text_color=ACCENT).pack()
                ctk.CTkLabel(cf,text=f"[{df[col].min():.2f}, {df[col].max():.2f}]",font=("Segoe UI",8),text_color=TEXT_SEC).pack(pady=(0,4))
            else: ctk.CTkLabel(cf,text=f"{df[col].nunique()} unique",font=("Segoe UI",9),text_color=TEXT_SEC).pack(pady=(0,4))
        preview=df if max_rows is None else df.head(max_rows); cols=list(preview.columns)
        def lbl(parent,text,fg,bg,bold=False,width=None,anchor="center"):
            font=("Segoe UI",10,"bold") if bold else ("Segoe UI",10)
            kw=dict(text=text,fg=fg,bg=bg,font=font,anchor=anchor,relief="flat",padx=4,pady=2)
            if width: kw["width"]=width
            return tk.Label(parent,**kw)
        lbl(self._inner,"#",TEXT_SEC,"#0d1117",bold=True,width=5).grid(row=0,column=0,sticky="nsew",padx=1,pady=1)
        for ci,col in enumerate(cols):
            lbl(self._inner,str(col)[:14],ACCENT2 if col in hi_cols else TEXT_SEC,"#0d1117",bold=True).grid(row=0,column=ci+1,sticky="nsew",padx=1,pady=1)
        for ri,(idx,row) in enumerate(preview.iterrows()):
            bg=BG_CARD if ri%2==0 else BG_PANEL
            lbl(self._inner,str(ri+1),TEXT_SEC,bg,width=4).grid(row=ri+1,column=0,sticky="nsew",padx=1,pady=0)
            for ci,col in enumerate(cols):
                val=row[col]
                if pd.isna(val): disp="NaN"; fg=DANGER
                else:
                    try: disp=f"{float(val):.4f}".rstrip("0").rstrip(".")
                    except: disp=str(val)[:14]
                    fg=TEXT_PRI
                lbl(self._inner,disp,fg,bg).grid(row=ri+1,column=ci+1,sticky="nsew",padx=1,pady=0)
        if max_rows and len(df)>max_rows:
            tk.Label(self._inner,text=f"  … showing {max_rows} of {len(df)} rows",fg=TEXT_SEC,bg="#0d1117",font=("Segoe UI",9,"italic"),anchor="w").grid(row=max_rows+2,column=0,columnspan=len(cols)+1,sticky="w",padx=6,pady=4)
        self._inner.update_idletasks(); self._on_inner_configure()


# ─── PDF Report ───────────────────────────────────────────────────────────────

class PDFReport:
    @staticmethod
    def generate(results, description, filename,
                 title=None, subtitle=None, byline=None,
                 composite_info=None, var_descriptions=None,
                 critical_value=None):
        doc = SimpleDocTemplate(filename, pagesize=letter,
                                rightMargin=40, leftMargin=40,
                                topMargin=36, bottomMargin=18)
        PAGE_W = letter[0] - 80
        els, sty = [], getSampleStyleSheet()
        def ps(name, **kw): return ParagraphStyle(name, parent=sty['Normal'], **kw)
        ts  = ps('T',  fontSize=11, textColor=colors.black, spaceAfter=2, alignment=TA_CENTER, fontName='Helvetica-Bold')
        ss  = ps('S',  fontSize=9,  textColor=colors.HexColor('#444444'), alignment=TA_CENTER, spaceAfter=2, fontName='Helvetica')
        bs  = ps('B',  fontSize=8,  textColor=colors.black, alignment=TA_CENTER, spaceAfter=4, fontName='Helvetica')
        hs  = ps('H',  fontSize=8,  textColor=colors.black, spaceAfter=2, spaceBefore=4, fontName='Helvetica-Bold', alignment=TA_LEFT)
        ns  = ps('N',  fontSize=7,  spaceAfter=2, alignment=TA_LEFT, fontName='Helvetica')
        its = ps('I',  fontSize=6.5, spaceAfter=2, alignment=TA_LEFT, fontName='Helvetica-Oblique')
        tts = ps('TT', fontSize=8,  textColor=colors.black, spaceAfter=2, spaceBefore=4, fontName='Helvetica-Oblique', alignment=TA_LEFT)
        fs  = ps('F',  fontSize=6,  textColor=colors.grey, alignment=TA_LEFT, fontName='Helvetica-Oblique')
        vhs = ps('VH', fontSize=7,  textColor=colors.HexColor('#1d4ed8'), spaceAfter=1, spaceBefore=3, fontName='Helvetica-Bold', alignment=TA_LEFT)
        vds = ps('VD', fontSize=6.5, spaceAfter=1, alignment=TA_LEFT, fontName='Helvetica', leftIndent=6, textColor=colors.HexColor('#333333'))
        vhs_head = ps('VHH', fontSize=8, textColor=colors.black, spaceAfter=2, spaceBefore=0, fontName='Helvetica-Bold', alignment=TA_LEFT)
        mi=results["method_info"]; sym=mi["symbol"]; meth=mi["label"]
        els.append(Paragraph(title or f"{meth} Pearson R", ts))
        if subtitle and subtitle.strip(): els.append(Paragraph(subtitle, ss))
        if byline and byline.strip(): els.append(Paragraph(byline, bs))
        els.append(Spacer(1, 3)); els.append(Paragraph(f"<i>Method: {meth}  (coefficient: {sym})</i>", its))
        if critical_value:
            cv_n=critical_value.get("n"); cv_df=critical_value.get("df"); cv_a=critical_value.get("alpha",0.05)
            cv_t=critical_value.get("t_critical"); cv_r=critical_value.get("critical_r")
            cv_t_str=f"{cv_t:.6f}" if cv_t is not None else "—"; cv_r_str=f"{cv_r:.6f}" if cv_r is not None else "—"
            els.append(Paragraph(f"<i>Critical Value</i> = {cv_r_str}  (α = {cv_a}, n = {cv_n}, df = {cv_df}, t_critical = {cv_t_str})", ns))
        if results.get("notes"): els.append(Paragraph(f"<i>{results['notes']}</i>", its))
        if description and description.strip(): els.append(Spacer(1, 2)); els.append(Paragraph(description, ns))
        if composite_info:
            els.append(Spacer(1, 3)); els.append(Paragraph("Composite Score Construction", hs))
            for part, info in composite_info.items():
                els.append(Paragraph(f"<b>{part}</b> ({info.get('mode','Mean')} of {len(info.get('items',[]))} items): {', '.join(info.get('items',[]))}", ns))
        def fmt(v, d=3): return f"{v:.{d}f}" if v is not None else "—"
        els.append(Spacer(1, 4)); els.append(Paragraph(f"<i>Frequentist Correlation Results — {meth}</i>", tts))
        hdr_row=['Variables',sym,f'{sym}²','t / z','p','Sig.','95% CI','Effect','N']
        tbl_data=[hdr_row]
        for p in results["pairs"]:
            ci_str=(f"[{fmt(p['ci_lower'])}, {fmt(p['ci_upper'])}]" if p['ci_lower'] is not None else "—")
            tbl_data.append([f"{p['x']} vs {p['y']}",fmt(p['r']),fmt(p['r_sq']),fmt(p['t_stat'],3),fmt(p['p'],4),p['stars'],ci_str,p['effect'],str(p['n'])])
        cw=[1.30*inch,.38*inch,.38*inch,.50*inch,.55*inch,.35*inch,1.00*inch,.58*inch,.30*inch]
        at=Table(tbl_data,colWidths=cw,repeatRows=1)
        at.setStyle(TableStyle([
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTNAME',(0,1),(-1,-1),'Helvetica'),
            ('FONTSIZE',(0,0),(-1,-1),7),('ALIGN',(0,0),(0,-1),'LEFT'),('ALIGN',(1,0),(-1,-1),'CENTER'),
            ('LINEABOVE',(0,0),(-1,0),.5,colors.black),('LINEBELOW',(0,0),(-1,0),.5,colors.black),
            ('LINEBELOW',(0,-1),(-1,-1),.5,colors.black),('TOPPADDING',(0,0),(-1,-1),2),
            ('BOTTOMPADDING',(0,0),(-1,-1),2),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f7f7f7')]),
        ]))
        els.append(at); els.append(Spacer(1, 2))
        els.append(Paragraph("<i>Note.</i> * p &lt; .05  ** p &lt; .01  *** p &lt; .001  ns=not significant.  CIs via Fisher z-transformation.", its))
        for p in [p for p in results["pairs"] if p.get("notes")]:
            els.append(Paragraph(f"<i>{p['x']} vs {p['y']}:</i> {p['notes']}", its))
        els.append(Spacer(1, 6))
        vars_=results["variables"]; mat=results["corr_matrix"]; n_vars=len(vars_)
        label_w=min(0.80*inch,1.0*inch); val_w=0.46*inch; left_w=label_w+val_w*n_vars+4
        mat_data=[[''] + [str(v)[:8] for v in vars_]]
        for v in vars_: mat_data.append([str(v)[:10]]+[fmt(mat.loc[v,v2]) for v2 in vars_])
        mt=Table(mat_data,colWidths=[label_w]+[val_w]*n_vars,repeatRows=1)
        mt.setStyle(TableStyle([
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTNAME',(0,1),(-1,-1),'Helvetica'),
            ('FONTSIZE',(0,0),(-1,-1),7),('ALIGN',(0,0),(-1,-1),'CENTER'),('ALIGN',(0,1),(0,-1),'LEFT'),
            ('LINEABOVE',(0,0),(-1,0),.4,colors.black),('LINEBELOW',(0,0),(-1,0),.4,colors.black),
            ('LINEBELOW',(0,-1),(-1,-1),.4,colors.black),('TOPPADDING',(0,0),(-1,-1),2),
            ('BOTTOMPADDING',(0,0),(-1,-1),2),('LEFTPADDING',(0,0),(-1,-1),3),('RIGHTPADDING',(0,0),(-1,-1),3),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f7f7f7')]),
        ]))
        right_w=PAGE_W-left_w-10
        COLORS_HEX=["#1d4ed8","#15803d","#7c3aed","#b45309","#c2410c","#0891b2","#be185d","#0f766e","#b91c1c"]
        right_content=[Paragraph("Variable Descriptions", vhs_head)]
        has_var_descs=False
        if var_descriptions:
            for i,(col,info) in enumerate(var_descriptions.items()):
                lbl=info.get("label","").strip(); desc=info.get("desc","").strip()
                if not lbl and not desc: continue
                has_var_descs=True; hex_c=COLORS_HEX[i%len(COLORS_HEX)]
                right_content.append(Paragraph(f"<font color='{hex_c}'><b>{col}</b></font>"+(f"  <i>{lbl}</i>" if lbl else ""), vhs))
                if desc: right_content.append(Paragraph(desc, vds))
        if not has_var_descs: right_content.append(Paragraph("<i>No variable descriptions provided.</i>", vds))
        inner_t=Table([[p] for p in right_content],colWidths=[right_w-8])
        inner_t.setStyle(TableStyle([('LEFTPADDING',(0,0),(-1,-1),2),('RIGHTPADDING',(0,0),(-1,-1),2),
                                      ('TOPPADDING',(0,0),(-1,-1),0),('BOTTOMPADDING',(0,0),(-1,-1),1),('VALIGN',(0,0),(-1,-1),'TOP')]))
        two_col=Table([[Paragraph("Correlation Matrix",hs),Paragraph("",hs)],[mt,inner_t]],colWidths=[left_w,right_w],hAlign='LEFT')
        two_col.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('LEFTPADDING',(0,0),(-1,-1),0),
                                      ('RIGHTPADDING',(0,0),(-1,-1),4),('TOPPADDING',(0,0),(-1,-1),0),
                                      ('BOTTOMPADDING',(0,0),(-1,-1),0),('LINEAFTER',(0,0),(0,-1),.5,colors.HexColor('#cccccc')),
                                      ('LEFTPADDING',(1,0),(1,-1),8)]))
        els.append(two_col); els.append(Spacer(1, 5))
        els.append(Paragraph("Correlation Interpretation Guide (APA-style)", hs))
        interp_data=[[Paragraph("<b>|r|</b>",ns),Paragraph("<b>Interpretation</b>",ns)],
                     [Paragraph("|r| ≥ .90",ns),Paragraph("Very strong",ns)],[Paragraph(".70 ≤ |r| < .90",ns),Paragraph("Strong",ns)],
                     [Paragraph(".50 ≤ |r| < .70",ns),Paragraph("Moderate",ns)],[Paragraph(".30 ≤ |r| < .50",ns),Paragraph("Weak",ns)],
                     [Paragraph("|r| < .30",ns),Paragraph("Negligible",ns)]]
        interp_t=Table(interp_data,colWidths=[2.0*inch,2.7*inch])
        interp_t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#f7f7f7")),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
                                       ("FONTSIZE",(0,0),(-1,-1),7),("ALIGN",(0,0),(-1,-1),"LEFT"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                                       ("LINEBELOW",(0,0),(-1,0),0.5,colors.black),("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#d1d5db")),
                                       ("INNERGRID",(0,0),(-1,-1),0.25,colors.HexColor("#d1d5db")),
                                       ("TOPPADDING",(0,0),(-1,-1),1),("BOTTOMPADDING",(0,0),(-1,-1),1)]))
        els.append(interp_t)
        els.append(Paragraph("<i>Note.</i> Negative values of r indicate an inverse relationship.", ns))
        els.append(Spacer(1, 4))
        els.append(Paragraph(f"File: {os.path.abspath(filename)}   Generated: {datetime.now().strftime('%B %d, %Y  %H:%M:%S')}", fs))
        doc.build(els)


# ─── Manual X / Y Entry Window ───────────────────────────────────────────────

class ManualEntryWindow(ctk.CTkToplevel):
    def __init__(self, master, on_load, **kw):
        super().__init__(master, **kw)
        self.title("Manual Data Entry  —  X / Y Values"); self.geometry("920x700"); self.configure(fg_color=BG_DEEP)
        self.on_load=on_load; self.col_data=[]
        self.transient(master); self.lift(); self.focus_force(); self.attributes("-topmost",True); self._build()

    def _build(self):
        hdr=ctk.CTkFrame(self,fg_color=ORANGE,corner_radius=0,height=58); hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr,text="  ✎  Manual X / Y Data Entry",font=("Segoe UI",18,"bold"),text_color="#0d1117",fg_color=ORANGE).pack(side="left",padx=20)
        ctk.CTkLabel(hdr,text="comma  •  space  •  newline  •  tab — all accepted",font=("Segoe UI",11),text_color="#7c3303",fg_color=ORANGE).pack(side="right",padx=20)
        body=ctk.CTkFrame(self,fg_color=BG_DEEP); body.pack(fill="both",expand=True,padx=14,pady=12)
        cfg=ctk.CTkFrame(body,fg_color=BG_CARD,corner_radius=10,border_width=1,border_color=BORDER); cfg.pack(fill="x",pady=(0,10))
        crow=ctk.CTkFrame(cfg,fg_color="transparent"); crow.pack(fill="x",padx=14,pady=10)
        ctk.CTkLabel(crow,text="Variables:",font=FONT_BODY,text_color=TEXT_SEC).pack(side="left",padx=(0,8))
        self.n_vars_menu=ctk.CTkOptionMenu(crow,values=["2","3","4","5","6","7","8"],command=self._rebuild_columns,
                                            fg_color=BG_INPUT,button_color=ORANGE,button_hover_color="#c2410c",
                                            text_color=TEXT_PRI,dropdown_fg_color=BG_PANEL,dropdown_text_color=TEXT_PRI,
                                            font=FONT_BODY,height=32,corner_radius=6,width=72)
        self.n_vars_menu.set("2"); self.n_vars_menu.pack(side="left",padx=(0,20))
        ctk.CTkLabel(crow,text="Variable names (comma-sep, optional):",font=FONT_BODY,text_color=TEXT_SEC).pack(side="left",padx=(0,8))
        self.names_entry=ctk.CTkEntry(crow,placeholder_text="e.g. Part1, Part11",fg_color=BG_INPUT,border_color=BORDER,
                                       text_color=TEXT_PRI,placeholder_text_color=TEXT_SEC,border_width=1,corner_radius=6,
                                       height=32,font=FONT_BODY,width=240)
        self.names_entry.pack(side="left",padx=(0,10))
        ctk.CTkButton(crow,text="Apply",command=self._apply_names,fg_color=ORANGE,hover_color="#c2410c",text_color="#0d1117",font=FONT_BTN,height=32,corner_radius=6,width=80).pack(side="left")
        hint=ctk.CTkFrame(body,fg_color=BG_PANEL,corner_radius=8,border_width=1,border_color=BORDER); hint.pack(fill="x",pady=(0,10))
        ctk.CTkLabel(hint,text="💡  Accepted formats: 3, 5, 2, 4 | 3 5 2 4 | one per line | paste Excel column. Blank/NA/. = NaN. Unequal lengths padded with NaN.",
                     font=("Segoe UI",10),text_color=TEXT_SEC,wraplength=860).pack(padx=12,pady=8)
        self.cols_frame=ctk.CTkFrame(body,fg_color="transparent"); self.cols_frame.pack(fill="both",expand=True,pady=(0,10))
        self._build_columns(2)
        prev=ctk.CTkFrame(body,fg_color=BG_CARD,corner_radius=8,border_width=1,border_color=BORDER,height=42); prev.pack(fill="x",pady=(0,10)); prev.pack_propagate(False)
        ctk.CTkLabel(prev,text="Preview:",font=("Segoe UI",11,"bold"),text_color=TEXT_SEC).pack(side="left",padx=12)
        self.preview_lbl=ctk.CTkLabel(prev,text="Enter values to see N and first rows",font=("Consolas",11),text_color=TEXT_SEC,wraplength=720); self.preview_lbl.pack(side="left",padx=6)
        brow=ctk.CTkFrame(body,fg_color="transparent"); brow.pack(fill="x")
        ctk.CTkButton(brow,text="✓  Load Data for Correlation",command=self._load,fg_color=SUCCESS,hover_color="#16a34a",text_color="#0d1117",font=FONT_BTN,height=44,corner_radius=8).pack(side="left",padx=(0,8))
        ctk.CTkButton(brow,text="👁  Preview",command=self._preview,fg_color=ACCENT2,hover_color="#3b7ddd",text_color="#0d1117",font=FONT_BTN,height=44,corner_radius=8).pack(side="left",padx=(0,8))
        ctk.CTkButton(brow,text="✕  Clear All",command=self._clear,fg_color=DANGER,hover_color="#b91c1c",font=FONT_BTN,height=44,corner_radius=8).pack(side="left",padx=(0,8))
        ctk.CTkButton(brow,text="Close",command=self.destroy,fg_color=BG_PANEL,hover_color=BORDER,font=FONT_BODY,height=44,corner_radius=8).pack(side="right")

    def _build_columns(self, n):
        for w in self.cols_frame.winfo_children(): w.destroy()
        self.col_data=[]; DN=["X","Y","Z","W","V","U","T","S"]; DC=[ACCENT2,SUCCESS,PURPLE,WARN,ORANGE,CYAN,PINK,ACCENT]
        for i in range(n):
            accent=DC[i%len(DC)]
            cf=ctk.CTkFrame(self.cols_frame,fg_color=BG_CARD,corner_radius=10,border_width=1,border_color=BORDER); cf.pack(side="left",fill="both",expand=True,padx=4)
            strip=ctk.CTkFrame(cf,fg_color=accent,corner_radius=0,height=38); strip.pack(fill="x"); strip.pack_propagate(False)
            name_var=tk.StringVar(value=DN[i])
            ctk.CTkEntry(strip,textvariable=name_var,fg_color=accent,border_color=accent,text_color="#0d1117",font=("Segoe UI",13,"bold"),corner_radius=0,height=36,border_width=0,justify="center").pack(fill="x",padx=6,pady=2)
            count_var=tk.StringVar(value="0 values")
            ctk.CTkLabel(cf,textvariable=count_var,font=("Segoe UI",10,"bold"),text_color=accent).pack(pady=(4,2))
            txt=ctk.CTkTextbox(cf,fg_color=BG_INPUT,text_color=TEXT_PRI,font=("Consolas",12),border_width=1,border_color=BORDER,corner_radius=6,wrap="none")
            txt.pack(fill="both",expand=True,padx=8,pady=(0,8))
            def _upd(event=None,cv=count_var,tb=txt): cv.set(f"{len(self._parse(tb.get('1.0','end')))} values")
            txt.bind("<KeyRelease>",_upd); txt.bind("<<Paste>>",lambda e,fn=_upd:self.after(60,fn))
            self.col_data.append({"name_var":name_var,"count_var":count_var,"textbox":txt})

    def _rebuild_columns(self, choice):
        saved=[{"name":cd["name_var"].get(),"text":cd["textbox"].get("1.0","end").strip()} for cd in self.col_data]
        self._build_columns(int(choice))
        for i,cd in enumerate(self.col_data):
            if i<len(saved):
                cd["name_var"].set(saved[i]["name"])
                if saved[i]["text"]: cd["textbox"].insert("1.0",saved[i]["text"])

    def _apply_names(self):
        raw=self.names_entry.get().strip()
        if not raw: return
        names=[n.strip() for n in raw.replace(","," ").split() if n.strip()]
        for i,cd in enumerate(self.col_data):
            if i<len(names): cd["name_var"].set(names[i])

    @staticmethod
    def _parse(raw):
        import re; cleaned=re.sub(r"[\t,;]+"," ",raw); result=[]
        for tok in cleaned.split():
            tok=tok.strip()
            if tok in ("","na","NA","nan","NaN","N/A","n/a","."): result.append(np.nan)
            else:
                try: result.append(float(tok))
                except: result.append(np.nan)
        return result

    def _build_df(self):
        series={}; lengths=[]
        for cd in self.col_data:
            name=cd["name_var"].get().strip() or f"Var{len(series)+1}"
            vals=self._parse(cd["textbox"].get("1.0","end"))
            if not vals: raise ValueError(f"Column '{name}' is empty.")
            series[name]=vals; lengths.append(len(vals))
        if len(set(lengths))>1:
            max_n=max(lengths)
            for k in series: series[k]+=[np.nan]*(max_n-len(series[k]))
        df=pd.DataFrame(series)
        if df.select_dtypes(include=[np.number]).empty: raise ValueError("No numeric values found.")
        return df

    def _preview(self):
        try:
            df=self._build_df()
            preview="  |  ".join(f"{col}: [{', '.join(str(round(v,2)) if not pd.isna(v) else 'NaN' for v in df[col].head(4).tolist())}...] (N={df[col].notna().sum()})" for col in df.columns)
            self.preview_lbl.configure(text=f"N={len(df)} rows  |  {preview}",text_color=ACCENT)
        except Exception as e: self.preview_lbl.configure(text=f"⚠  {e}",text_color=DANGER)

    def _load(self):
        try:
            df=self._build_df(); self.on_load(df)
            messagebox.showinfo("Loaded",f"Data loaded!\nRows: {len(df)}\nVariables: {len(df.columns)}\n\nColumns: {', '.join(df.columns)}\n\nClick  ▶ Compute Correlation  to proceed.",parent=self)
            self.destroy()
        except Exception as e: messagebox.showerror("Error",str(e),parent=self)

    def _clear(self):
        for cd in self.col_data: cd["textbox"].delete("1.0","end"); cd["count_var"].set("0 values")
        self.preview_lbl.configure(text="Enter values to see N and first rows",text_color=TEXT_SEC)


# ─── Scrollable Sidebar ───────────────────────────────────────────────────────

class Sidebar(ctk.CTkFrame):
    def __init__(self, master, **kw):
        super().__init__(master, width=262, fg_color=BG_CARD, corner_radius=0, **kw)
        self.pack_propagate(False)
        self._method_key="pearson"; self._on_method_cb=None; self._build()

    @property
    def method_key(self): return self._method_key

    def _build(self):
        self.logo_frame=ctk.CTkFrame(self,fg_color=ACCENT2,corner_radius=0,height=64); self.logo_frame.pack(fill="x"); self.logo_frame.pack_propagate(False)
        self.logo_label=ctk.CTkLabel(self.logo_frame,text="  r  ",font=("Segoe UI",30,"bold"),text_color="#0d1117",fg_color=ACCENT2); self.logo_label.pack(expand=True)
        self._scroll=ctk.CTkScrollableFrame(self,fg_color=BG_CARD,scrollbar_button_color=BORDER,corner_radius=0); self._scroll.pack(fill="both",expand=True)
        s=self._scroll
        self.method_name_lbl=ctk.CTkLabel(s,text="Pearson Product-Moment",font=("Segoe UI",12,"bold"),text_color=TEXT_PRI,fg_color=BG_CARD,wraplength=240); self.method_name_lbl.pack(pady=(10,2))
        ctk.CTkLabel(s,text="Correlation Analysis Suite",font=FONT_TINY,text_color=TEXT_SEC,fg_color=BG_CARD).pack(pady=(0,8))
        divider(s)
        sec_label(s,"CORRELATION METHOD")
        self.method_menu=ctk.CTkOptionMenu(s,values=[v["label"] for v in METHODS.values()],command=self._on_method_change,
                                            fg_color=BG_INPUT,button_color=ACCENT2,button_hover_color="#3b7ddd",text_color=TEXT_PRI,
                                            dropdown_fg_color=BG_PANEL,dropdown_text_color=TEXT_PRI,font=("Segoe UI",12),height=36,
                                            corner_radius=6,width=238,dynamic_resizing=False)
        self.method_menu.set("Pearson Product-Moment"); self.method_menu.pack(fill="x",padx=12,pady=(0,8))
        divider(s)
        sec_label(s,"REPORT TITLE")
        self.title_entry=styled_entry(s,placeholder="Correlation Analysis"); self.title_entry.insert(0,"Pearson R"); self.title_entry.pack(fill="x",padx=14,pady=(0,4))
        sec_label(s,"SUBTITLE")
        self.subtitle_entry=styled_entry(s,placeholder="Optional subtitle"); self.subtitle_entry.pack(fill="x",padx=14,pady=(0,4))
        sec_label(s,"AUTHOR / RESEARCHER")
        self.author_entry=styled_entry(s,placeholder="e.g. Dr. Jane Reyes"); self.author_entry.pack(fill="x",padx=14,pady=(0,4))
        divider(s)
        pad={"fill":"x","padx":14,"pady":4}
        self.import_btn=sidebar_btn(s,"📁  Import CSV / Excel",fg=ACCENT2,hover="#3b7ddd"); self.import_btn.pack(**pad)
        self.expander_btn=sidebar_btn(s,"▶  Composite Score Expander",fg="#6d28d9",hover="#5b21b6"); self.expander_btn.pack(**pad)
        self.manual_btn=sidebar_btn(s,"✎  Manual X / Y Entry",fg=ORANGE,hover="#c2410c"); self.manual_btn.pack(**pad)
        self.select_btn=sidebar_btn(s,"⚙   Select Variables",fg="#0f766e",hover="#134e4a",state="disabled"); self.select_btn.pack(**pad)
        divider(s)
        sec_label(s,"DATASET")
        self.view_data_btn=sidebar_btn(s,"🔍  View Dataset",fg=CYAN,hover="#0891b2",text_color="#0d1117",state="disabled"); self.view_data_btn.pack(**pad)
        self.edit_data_btn=sidebar_btn(s,"✏️  Edit Dataset",fg=ORANGE,hover="#c2410c",text_color="#0d1117",state="disabled"); self.edit_data_btn.pack(**pad)
        self.var_desc_btn=sidebar_btn(s,"📝  Variable Descriptions",fg=PURPLE,hover="#7c3aed",text_color=TEXT_PRI,state="disabled"); self.var_desc_btn.pack(**pad)
        divider(s)
        self.compute_btn=sidebar_btn(s,"▶   Compute Correlation",fg=ACCENT2,hover="#3b7ddd",text_color="#0d1117",font=("Segoe UI",13,"bold"),height=44); self.compute_btn.pack(**pad)
        divider(s)

        # ── [NEW] PAIRWISE COMPARE — Free Pair Selector ───────────────────
        sec_label(s, "PAIRWISE COMPARE")

        self.pair_mode_var = tk.BooleanVar(value=False)
        self.pair_mode_chk = ctk.CTkCheckBox(
            s, text="Enable free pair selection",
            variable=self.pair_mode_var,
            fg_color=BG_INPUT, hover_color=BG_PANEL,
            text_color=TEXT_PRI, font=FONT_BODY)
        self.pair_mode_chk.pack(fill="x", padx=14, pady=(0, 6))

        self.open_pair_selector_btn = sidebar_btn(
            s, "🔗  Choose Pairs…",
            fg="#0f766e", hover="#134e4a",
            text_color=TEXT_PRI, state="disabled")
        self.open_pair_selector_btn.pack(**pad)

        self.pair_summary_lbl = ctk.CTkLabel(
            s, text="No pairs selected",
            font=FONT_TINY, text_color=TEXT_SEC,
            fg_color=BG_CARD, wraplength=230)
        self.pair_summary_lbl.pack(padx=14, pady=(0, 8))

        # Backward-compat stubs (hidden, 1px)
        self.pair_seq_mode_var = tk.BooleanVar(value=False)
        self.pair_var1_menu = ctk.CTkOptionMenu(
            s, values=["(load data first)"],
            fg_color=BG_INPUT, button_color=ACCENT2, button_hover_color="#3b7ddd",
            text_color=TEXT_PRI, dropdown_fg_color=BG_PANEL, dropdown_text_color=TEXT_PRI,
            font=FONT_BODY, height=1, corner_radius=6, width=1)
        self.pair_var2_menu = ctk.CTkOptionMenu(
            s, values=["(load data first)"],
            fg_color=BG_INPUT, button_color=SUCCESS, button_hover_color="#16a34a",
            text_color=TEXT_PRI, dropdown_fg_color=BG_PANEL, dropdown_text_color=TEXT_PRI,
            font=FONT_BODY, height=1, corner_radius=6, width=1)
        self.pair_seq_entry = ctk.CTkEntry(
            s, height=1, width=1, fg_color=BG_INPUT,
            border_color=BORDER, text_color=TEXT_PRI, font=FONT_BODY)
        # ─────────────────────────────────────────────────────────────────

        sec_label(s,"PEARSON r — CRITICAL VALUE")
        crit_row=ctk.CTkFrame(s,fg_color="transparent"); crit_row.pack(fill="x",padx=14,pady=(0,6))
        ctk.CTkLabel(crit_row,text="n",font=FONT_BODY,text_color=TEXT_SEC,width=20).pack(side="left",padx=(0,6))
        self.critical_n_entry=styled_entry(crit_row,placeholder="sample size",width=100,height=34); self.critical_n_entry.pack(side="left",padx=(0,8))
        self.critical_r_btn=ctk.CTkButton(crit_row,text="Compute",fg_color=BG_PANEL,hover_color=BORDER,text_color=TEXT_PRI,font=FONT_BODY,height=34,width=118,corner_radius=6,border_width=1,border_color=BORDER)
        self.critical_r_btn.pack(side="left")
        self.plot_btn=sidebar_btn(s,"📉  View Scatter Plots",fg=PURPLE,hover="#7c3aed",state="disabled"); self.plot_btn.pack(**pad)
        self.export_btn=sidebar_btn(s,"📄  Export PDF",fg="#1d4ed8",hover="#1e3a8a",state="disabled"); self.export_btn.pack(**pad)
        self.print_btn=sidebar_btn(s,"🖨️  Print Report",fg="#7c3aed",hover="#6d28d9",state="disabled"); self.print_btn.pack(**pad)
        self.export_data_btn=sidebar_btn(s,"💾  Export Dataset",fg="#0f766e",hover="#134e4a",state="disabled"); self.export_data_btn.pack(**pad)
        self.clear_btn=sidebar_btn(s,"🗑  Clear / Reset",fg=DANGER,hover="#b91c1c"); self.clear_btn.pack(**pad)
        divider(s)
        self.theme_btn=sidebar_btn(s,"☀️  Light Mode",fg="#374151",hover="#4b5563",font=FONT_BODY,height=32); self.theme_btn.pack(fill="x",padx=14,pady=8)
        if HAS_SETTINGS:
            divider(s)
            self.settings_btn=sidebar_btn(s,"⚙   Settings",fg="#374151",hover="#4b5563",font=FONT_BODY,height=32); self.settings_btn.pack(fill="x",padx=14,pady=8)
        divider(s)
        self.status_label=ctk.CTkLabel(s,text="",font=FONT_TINY,text_color=ACCENT,fg_color=BG_CARD,wraplength=238); self.status_label.pack(padx=12,pady=12)

    def _on_method_change(self, choice):
        for k,v in METHODS.items():
            if v["label"]==choice:
                self._method_key=k; color=v["color"]; sym=v["symbol"]
                self.logo_frame.configure(fg_color=color); self.logo_label.configure(fg_color=color,text=f"  {sym}  ")
                self.method_name_lbl.configure(text=choice); self.compute_btn.configure(fg_color=color,hover_color=color)
                if self._on_method_cb: self._on_method_cb(k)
                break

    def set_method_callback(self, fn): self._on_method_cb=fn


# ─── Main App ─────────────────────────────────────────────────────────────────

class PearsonRApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Correlation Analysis Suite"); self.geometry("1440x860"); self.minsize(1200,740); self.configure(fg_color=BG_DEEP)
        self.df=None; self.raw_df=None; self.selected_cols=None; self.controls=[]
        self.results=None; self.critical_value_info=None; self.composite_info=None
        self.var_descriptions={}; self.dark_mode=True
        self.free_pairs=[]          # [NEW] list of (a, b) tuples
        self._build_ui()

    def _available_numeric_columns(self):
        if self.df is None: return []
        if self.selected_cols:
            return [c for c in self.selected_cols if c in self.df.columns and c in set(self.df.select_dtypes(include=[np.number]).columns)]
        return list(self.df.select_dtypes(include=[np.number]).columns)

    def _update_pair_compare_controls(self):
        cols=self._available_numeric_columns()
        if len(cols)<2:
            self.sidebar.pair_mode_var.set(False)
            self.sidebar.open_pair_selector_btn.configure(state="disabled")
            return
        if self.df is not None:
            self.sidebar.open_pair_selector_btn.configure(state="normal")

    def _parse_sequence_vars(self, raw, allowed_set):
        if raw is None: return []
        s=raw.strip()
        if not s: return []
        s=s.replace(";",",").replace("\n",",").replace("\r",",")
        tokens=[t.strip() for t in (s.split(",") if "," in s else s.split()) if t.strip()]
        out,seen=[],set()
        for t in tokens:
            if t in allowed_set and t not in seen: out.append(t); seen.add(t)
        return out

    def _build_ui(self):
        self.sidebar=Sidebar(self); self.sidebar.pack(side="left",fill="y")
        self.sidebar.set_method_callback(self._on_method_changed)
        self.sidebar.import_btn.configure(command=self.import_data)
        self.sidebar.expander_btn.configure(command=self.open_expander)
        self.sidebar.manual_btn.configure(command=self.open_manual_entry)
        self.sidebar.select_btn.configure(command=self.open_variable_selector)
        self.sidebar.view_data_btn.configure(command=self.open_dataset_viewer)
        self.sidebar.edit_data_btn.configure(command=self.open_dataset_editor)
        self.sidebar.var_desc_btn.configure(command=self.open_var_descriptions)
        self.sidebar.compute_btn.configure(command=self.compute_r)
        self.sidebar.critical_r_btn.configure(command=self.compute_critical_r)
        self.sidebar.plot_btn.configure(command=self.open_plot)
        self.sidebar.export_btn.configure(command=self.export_pdf)
        self.sidebar.print_btn.configure(command=self.print_report)
        self.sidebar.export_data_btn.configure(command=self.export_dataset)
        self.sidebar.clear_btn.configure(command=self.clear_all)
        self.sidebar.theme_btn.configure(command=self.toggle_theme)
        self.sidebar.open_pair_selector_btn.configure(command=self.open_free_pair_selector)  # [NEW]
        if HAS_SETTINGS: self.sidebar.settings_btn.configure(command=self.open_settings)

        content=ctk.CTkFrame(self,fg_color=BG_DEEP,corner_radius=0); content.pack(side="left",fill="both",expand=True)
        hdr=ctk.CTkFrame(content,fg_color=BG_CARD,corner_radius=0,height=64); hdr.pack(fill="x"); hdr.pack_propagate(False)
        ctk.CTkLabel(hdr,text="Correlation Analysis Suite",font=FONT_HEAD,text_color=TEXT_PRI).pack(side="left",padx=24)
        self.hdr_method_lbl=ctk.CTkLabel(hdr,text="Pearson Product-Moment  |  r",font=("Segoe UI",11),text_color=ACCENT2); self.hdr_method_lbl.pack(side="right",padx=24)
        outer=ctk.CTkFrame(content,fg_color=BG_DEEP); outer.pack(fill="both",expand=True,padx=14,pady=14)
        sty=ttk.Style(); sty.theme_use("default"); sty.configure("Sash",sashthickness=6,sashrelief="flat",background="#30363d")
        pane=ttk.PanedWindow(outer,orient=tk.HORIZONTAL); pane.pack(fill="both",expand=True)
        lw=ctk.CTkFrame(pane,fg_color=BG_DEEP); lc=card(lw,title="📋  Description & Method Info"); lc.pack(fill="both",expand=True); self._build_desc_panel(lc); pane.add(lw,weight=3)
        cw=ctk.CTkFrame(pane,fg_color=BG_DEEP); cc=card(cw,title="📊  Dataset Preview"); cc.pack(fill="both",expand=True); self._build_preview_panel(cc); pane.add(cw,weight=2)
        rw=ctk.CTkFrame(pane,fg_color=BG_DEEP); rc=card(rw,title="📈  Analysis Results"); rc.pack(fill="both",expand=True); self._build_results_panel(rc); pane.add(rw,weight=3)
        def _set_sash(*_):
            total=pane.winfo_width()
            if total>100: pane.sashpos(0,int(total*3/8)); pane.sashpos(1,int(total*6/8)); pane.unbind("<Configure>")
        pane.bind("<Configure>",_set_sash)
        bar=ctk.CTkFrame(content,fg_color=BG_CARD,height=28,corner_radius=0); bar.pack(fill="x",side="bottom"); bar.pack_propagate(False)
        self.file_label=ctk.CTkLabel(bar,text="No file saved yet",font=FONT_TINY,text_color=TEXT_SEC); self.file_label.pack(side="left",padx=12)
        self.stat_label=ctk.CTkLabel(bar,text="",font=("Segoe UI",11,"bold"),text_color=ACCENT2); self.stat_label.pack(side="right",padx=12)

    def _build_desc_panel(self, parent):
        scroll=ctk.CTkScrollableFrame(parent,fg_color=BG_CARD,scrollbar_button_color=BORDER); scroll.pack(fill="both",expand=True,padx=12,pady=(0,12))
        ctk.CTkLabel(scroll,text="STUDY DESCRIPTION",font=("Segoe UI",11,"bold"),text_color=TEXT_SEC).pack(anchor="w",padx=4,pady=(8,3))
        self.desc_text=ctk.CTkTextbox(scroll,height=90,fg_color=BG_INPUT,text_color=TEXT_PRI,border_width=1,border_color=BORDER,font=FONT_BODY,corner_radius=6); self.desc_text.pack(fill="x",padx=4,pady=(0,10))
        ctk.CTkLabel(scroll,text="METHOD DESCRIPTION",font=("Segoe UI",11,"bold"),text_color=TEXT_SEC).pack(anchor="w",padx=4,pady=(4,3))
        self.method_desc_box=ctk.CTkTextbox(scroll,height=175,fg_color=BG_PANEL,text_color=TEXT_PRI,border_width=1,border_color=BORDER,font=("Consolas",11),corner_radius=6); self.method_desc_box.pack(fill="x",padx=4,pady=(0,10))
        self._update_method_desc("pearson")
        ref=ctk.CTkFrame(scroll,fg_color=BG_PANEL,corner_radius=8,border_width=1,border_color=BORDER); ref.pack(fill="x",padx=4,pady=(0,10))
        ctk.CTkLabel(ref,text="METHOD QUICK REFERENCE",font=("Segoe UI",11,"bold"),text_color=TEXT_SEC).pack(anchor="w",padx=12,pady=(10,6))
        for k,v in METHODS.items():
            row=ctk.CTkFrame(ref,fg_color="transparent"); row.pack(fill="x",padx=12,pady=2)
            ctk.CTkFrame(row,width=10,height=10,fg_color=v["color"],corner_radius=5).pack(side="left",padx=(0,8))
            ctk.CTkLabel(row,text=v["symbol"],font=("Consolas",11,"bold"),text_color=v["color"],width=32).pack(side="left")
            ctk.CTkLabel(row,text=v["label"],font=FONT_TINY,text_color=TEXT_PRI).pack(side="left",padx=6)
        ctk.CTkFrame(ref,height=6,fg_color="transparent").pack()
        guide=ctk.CTkFrame(scroll,fg_color=BG_PANEL,corner_radius=8,border_width=1,border_color=BORDER); guide.pack(fill="x",padx=4,pady=(0,8))
        ctk.CTkLabel(guide,text="EFFECT SIZE GUIDE (Cohen, 1988)",font=("Segoe UI",11,"bold"),text_color=TEXT_SEC).pack(anchor="w",padx=12,pady=(10,6))
        for col,rng,lbl in [(ACCENT,"|r| ≥ 0.90","Very High"),(ACCENT2,"|r| ≥ 0.70","High"),(SUCCESS,"|r| ≥ 0.50","Moderate"),(WARN,"|r| ≥ 0.30","Low"),(TEXT_SEC,"|r| < 0.30","Negligible")]:
            row=ctk.CTkFrame(guide,fg_color="transparent"); row.pack(fill="x",padx=12,pady=2)
            ctk.CTkFrame(row,width=10,height=10,fg_color=col,corner_radius=5).pack(side="left",padx=(0,8))
            ctk.CTkLabel(row,text=rng,font=("Segoe UI",11,"bold"),text_color=TEXT_PRI,width=92).pack(side="left")
            ctk.CTkLabel(row,text=lbl,font=FONT_BODY,text_color=TEXT_SEC).pack(side="left",padx=6)
        ctk.CTkFrame(guide,height=6,fg_color="transparent").pack()

    def _build_preview_panel(self, parent):
        info_row=ctk.CTkFrame(parent,fg_color=BG_PANEL,corner_radius=6,height=28); info_row.pack(fill="x",padx=12,pady=(0,4)); info_row.pack_propagate(False)
        ctk.CTkLabel(info_row,text="Loaded:",font=("Segoe UI",10,"bold"),text_color=TEXT_SEC).pack(side="left",padx=8)
        self.data_info_lbl=ctk.CTkLabel(info_row,text="None",font=("Consolas",10),text_color=TEXT_SEC); self.data_info_lbl.pack(side="left",padx=4)
        self.vars_info_lbl=ctk.CTkLabel(info_row,text="",font=("Segoe UI",10),text_color=ACCENT2); self.vars_info_lbl.pack(side="right",padx=8)
        self.data_table=DataTableFrame(parent); self.data_table.pack(fill="both",expand=True,padx=12,pady=(0,12))
        self.data_table.display_data(None,selected_cols=None); self._update_data_info()

    def _build_results_panel(self, parent):
        self.results_text=ctk.CTkTextbox(parent,fg_color=BG_INPUT,text_color=TEXT_PRI,font=FONT_MONO,wrap="word",border_width=1,border_color=BORDER,corner_radius=8)
        self.results_text.pack(fill="both",expand=True,padx=12,pady=(0,12))
        self._set_results(
            "═══════════════════════════════════════\n"
            "  Correlation Analysis Suite\n"
            "  9 Methods  |  JASP-Compatible\n"
            "═══════════════════════════════════════\n\n"
            "[NEW] Free Pair Selector:\n"
            "  🔗 Click 'Choose Pairs…' in sidebar\n"
            "     to freely pick any variable pairs\n"
            "  ✓  Tick 'Enable free pair selection'\n"
            "     then click Compute\n\n"
            "Steps:\n"
            " 1. Select method from sidebar\n"
            " 2. Import data  OR  use Expander\n"
            " 3. View / Edit dataset as needed\n"
            " 4. Optionally choose specific pairs\n"
            " 5. Click  ▶ Compute Correlation\n"
            " 6. View plots / Export PDF\n\n"
            "Waiting for data…")

    def _set_results(self, text):
        self.results_text.configure(state="normal"); self.results_text.delete("1.0","end"); self.results_text.insert("1.0",text); self.results_text.configure(state="disabled")

    def _update_data_info(self, source_label=""):
        if self.df is None: self.data_info_lbl.configure(text="None",text_color=TEXT_SEC); self.vars_info_lbl.configure(text=""); return
        rows,cols_n=len(self.df),len(self.df.columns); num_n=len(self.df.select_dtypes(include=[np.number]).columns)
        nan_total=int(self.df.isna().sum().sum()); nan_txt=f"  ⚠{nan_total} NaN" if nan_total else ""
        self.data_info_lbl.configure(text=f"{rows} rows × {cols_n} cols  ({num_n} numeric){nan_txt}  [{source_label}]",text_color=ACCENT)
        sel=self.selected_cols or []
        if sel: self.vars_info_lbl.configure(text=f"Selected: {', '.join(sel[:6])}{'…' if len(sel)>6 else ''}",text_color=ACCENT2)
        else: self.vars_info_lbl.configure(text="",text_color=TEXT_SEC)

    def _update_method_desc(self, key):
        self.method_desc_box.configure(state="normal"); self.method_desc_box.delete("1.0","end"); self.method_desc_box.insert("1.0",METHODS[key]["desc"]); self.method_desc_box.configure(state="disabled")

    def _on_method_changed(self, key):
        self._update_method_desc(key); mi=METHODS[key]
        self.hdr_method_lbl.configure(text=f"{mi['label']}  |  {mi['symbol']}",text_color=mi["color"])
        self._set_results(f"Method changed: {mi['label']}  ({mi['symbol']})\n\nImport data and click  ▶ Compute Correlation.")

    def _activate_data_buttons(self):
        for btn in (self.sidebar.select_btn,self.sidebar.export_data_btn,self.sidebar.view_data_btn,
                    self.sidebar.edit_data_btn,self.sidebar.var_desc_btn,self.sidebar.open_pair_selector_btn):
            btn.configure(state="normal")

    def import_data(self):
        fp=filedialog.askopenfilename(filetypes=[("CSV","*.csv"),("Excel","*.xlsx"),("All","*.*")],title="Select data file")
        if not fp: return
        try:
            self.df=pd.read_csv(fp) if fp.endswith(".csv") else pd.read_excel(fp)
            self.selected_cols=list(self.df.select_dtypes(include=[np.number]).columns)
            self.controls=[]; self.composite_info=None; self.var_descriptions={}; self.free_pairs=[]
            self.data_table.display_data(self.df,selected_cols=self.selected_cols)
            self._update_pair_compare_controls(); self._activate_data_buttons()
            mi=METHODS[self.sidebar.method_key]
            self._set_results(f"═══════════════════════════════════\n  Data Imported!\n═══════════════════════════════════\n\nFile:    {os.path.basename(fp)}\nRows:    {len(self.df)}\nCols:    {len(self.df.columns)}\n\nNumeric Vars ({len(self.selected_cols)}):\n{', '.join(self.selected_cols)}\n\nMethod: {mi['label']}\n\n✓ Use 🔍 View / ✏️ Edit / 🔗 Choose Pairs before computing.")
            self.sidebar.status_label.configure(text=f"✓ Imported\n{os.path.basename(fp)}"); self._update_data_info(f"CSV/Excel: {os.path.basename(fp)}")
        except Exception as e: messagebox.showerror("Import Error",f"Failed:\n{e}")

    def open_expander(self): CompositeExpanderWindow(self,on_generate=self._on_composite_generated)
    def open_manual_entry(self): ManualEntryWindow(self,on_load=self._on_manual_loaded)

    def _on_manual_loaded(self, df):
        self.df=df; self.raw_df=None
        self.selected_cols=list(df.select_dtypes(include=["float64","int64","float32","int32","float","int"]).columns)
        self.controls=[]; self.composite_info=None; self.var_descriptions={}; self.free_pairs=[]
        self.data_table.display_data(df,selected_cols=self.selected_cols)
        self._update_pair_compare_controls(); self._activate_data_buttons()
        mi=METHODS[self.sidebar.method_key]
        self._set_results(f"{'='*37}\n  Manual Data Loaded!\n{'='*37}\n\nRows (respondents): {len(df)}\nVariables:          {len(df.columns)}\n\nColumns: {', '.join(df.columns)}\n\nMethod: {mi['label']}\n\n✓ Click  ▶ Compute Correlation  to proceed.")
        self.sidebar.status_label.configure(text=f"✓ Manual entry\n{len(df)} respondents"); self._update_data_info("Manual Entry")

    def _on_composite_generated(self, composite_df, mapping_info, mode, raw_df):
        self.df=composite_df; self.raw_df=raw_df
        self.selected_cols=list(composite_df.columns); self.controls=[]
        self.composite_info={part:{"items":items,"mode":mode} for part,items in mapping_info.items()}
        self.var_descriptions={}; self.free_pairs=[]
        self.data_table.display_data(composite_df,selected_cols=self.selected_cols)
        self._update_pair_compare_controls(); self._activate_data_buttons()
        summary="\n".join(f"  {part} ({mode} of {len(info['items'])} items): {', '.join(info['items'])}" for part,info in self.composite_info.items())
        self._set_results(f"{'='*37}\n  Composite Scores Generated!\n{'='*37}\n\nScore Type:   {mode}\nRespondents:  {len(composite_df)}\nParts:        {len(composite_df.columns)}\n\nPart Composition:\n{summary}\n\n✓ Click  ▶ Compute Correlation  to proceed.")
        self.sidebar.status_label.configure(text=f"✓ Composite\n{len(composite_df)} respondents"); self._update_data_info("Composite Expander")

    def open_dataset_viewer(self):
        if self.df is None: messagebox.showwarning("No Data","Import data first."); return
        DatasetViewerWindow(self,self.df,on_edit_callback=self.open_dataset_editor)

    def open_dataset_editor(self):
        if self.df is None: messagebox.showwarning("No Data","Import data first."); return
        DatasetEditorWindow(self,self.df,on_apply=self._on_dataset_edited)

    def _on_dataset_edited(self, new_df):
        self.df=new_df; num_cols=list(new_df.select_dtypes(include=[np.number]).columns)
        if self.selected_cols: self.selected_cols=[c for c in self.selected_cols if c in new_df.columns]
        if not self.selected_cols: self.selected_cols=num_cols
        self.var_descriptions={k:v for k,v in self.var_descriptions.items() if k in new_df.columns}
        self.free_pairs=[(a,b) for a,b in self.free_pairs if a in new_df.columns and b in new_df.columns]
        self.results=None
        for btn in (self.sidebar.export_btn,self.sidebar.print_btn,self.sidebar.plot_btn): btn.configure(state="disabled")
        self.data_table.display_data(new_df,selected_cols=self.selected_cols)
        self._update_pair_compare_controls(); self._update_data_info("Edited")
        self._set_results(f"{'='*37}\n  Dataset Updated  (Edited)\n{'='*37}\n\nRows:  {len(new_df)}\nCols:  {len(new_df.columns)}\n\nColumns: {', '.join(new_df.columns)}\n\n✓ Re-run  ▶ Compute Correlation  with the updated data.")
        self.sidebar.status_label.configure(text=f"✓ Dataset edited\n{len(new_df)} rows")

    def open_var_descriptions(self):
        if self.df is None: messagebox.showwarning("No Data","Import data first."); return
        VariableDescriptionsWindow(self,list(self.df.columns),existing_descs=self.var_descriptions,on_save=self._on_var_descs_saved)

    def _on_var_descs_saved(self, descs):
        self.var_descriptions=descs
        n_filled=sum(1 for v in descs.values() if v.get("label") or v.get("desc"))
        self.sidebar.status_label.configure(text=f"✓ {n_filled} var descriptions\nsaved")
        self._set_results(f"{'='*37}\n  Variable Descriptions Saved\n{'='*37}\n\n"+"\n".join(f"  {col}:\n    Label: {info.get('label','—')}\n    Desc:  {(info.get('desc','—') or '—')[:60]}{'…' if len(info.get('desc',''))>60 else ''}" for col,info in descs.items() if info.get("label") or info.get("desc"))+"\n\nDescriptions will appear in the next PDF export.")

    # ── [NEW] Free Pair Selector ──────────────────────────────────────────────

    def open_free_pair_selector(self):
        if self.df is None: messagebox.showwarning("No Data","Import data first."); return
        cols=list(self.df.select_dtypes(include=[np.number]).columns)
        if len(cols)<2: messagebox.showwarning("Not Enough","Need ≥ 2 numeric variables."); return
        FreePairSelectorWindow(self, cols, existing_pairs=self.free_pairs, on_save=self._on_free_pairs_saved)

    def _on_free_pairs_saved(self, pairs):
        self.free_pairs=pairs; self.sidebar.pair_mode_var.set(True)
        summary="  |  ".join(f"{a} ↔ {b}" for a,b in pairs)
        self.sidebar.pair_summary_lbl.configure(
            text=f"{len(pairs)} pair(s): {summary[:120]}{'…' if len(summary)>120 else ''}",
            text_color=ACCENT)
        self.sidebar.status_label.configure(text=f"✓ {len(pairs)} free pair(s)\nselected")
        self._set_results(f"{'='*37}\n  Free Pairs Selected\n{'='*37}\n\n"+"\n".join(f"  {i+1}. {a}  ↔  {b}" for i,(a,b) in enumerate(pairs))+f"\n\n✓ {len(pairs)} pair(s) ready.\nClick  ▶ Compute Correlation  to proceed.")

    # ─────────────────────────────────────────────────────────────────────────

    def open_plot(self):
        if self.results is None: messagebox.showwarning("No Results","Compute correlation first."); return
        if self.df is None: messagebox.showwarning("No Data","No data available for plotting."); return
        ScatterPlotWindow(self,self.results,self.df)

    def open_variable_selector(self):
        if self.df is None: messagebox.showwarning("No Data","Import data first."); return
        num_cols=list(self.df.select_dtypes(include=[np.number]).columns)
        if len(num_cols)<2: messagebox.showwarning("Not Enough","Need ≥ 2 numeric variables."); return
        VariableSelectorWindow(self,num_cols,method_key=self.sidebar.method_key,on_select=self._on_vars_selected)

    def _on_vars_selected(self, chosen, controls):
        self.selected_cols=chosen; self.controls=controls
        ctrl_txt=f"Controls: {', '.join(controls)}" if controls else "No control variables"
        self._set_results(f"Variables Selected\n\nAnalysis ({len(chosen)}): {', '.join(chosen)}\n{ctrl_txt}\n\n✓ Click  ▶ Compute Correlation  to proceed.")
        self.sidebar.status_label.configure(text=f"✓ {len(chosen)} vars\n{ctrl_txt}")
        if self.df is not None:
            self.data_table.display_data(self.df,selected_cols=self.selected_cols)
            self._update_pair_compare_controls(); self._update_data_info()

    def compute_critical_r(self):
        raw=self.sidebar.critical_n_entry.get().strip()
        if not raw: messagebox.showwarning("Critical r","Enter sample size n.",parent=self); return
        try: n=int(float(raw))
        except ValueError: messagebox.showerror("Critical r","n must be a number.",parent=self); return
        try: r_crit,df,t_crit=CorrelationEngine.pearson_critical_r(n,alpha=0.05)
        except ValueError as e: messagebox.showwarning("Critical r",str(e),parent=self); return
        self.critical_value_info={"n":n,"df":df,"alpha":0.05,"t_critical":t_crit,"critical_r":r_crit,"tail":"two-tailed"}
        messagebox.showinfo("Pearson r — critical value",f"Two-tailed test, α = 0.05\n\nn = {n}\ndf = {df}\nt_critical = {t_crit:.6f}\n\nCritical |r| = {r_crit:.6f}\n\nA sample correlation is significant at α = 0.05 (two-tailed)\nif |r| exceeds this value.",parent=self)

    def compute_r(self):
        if self.df is None: messagebox.showwarning("No Data","Import data first."); return
        try:
            mkey=self.sidebar.method_key; mi=METHODS[mkey]; ctrl=self.controls or []
            cols=self.selected_cols or list(self.df.select_dtypes(include=[np.number]).columns)
            self.results=None

            # ── [NEW] Free Pair Mode ──────────────────────────────────────
            if self.sidebar.pair_mode_var.get():
                free_pairs=self.free_pairs
                if not free_pairs:
                    messagebox.showwarning("No Pairs","Click '🔗 Choose Pairs…' to select pairs first.",parent=self); return
                if mkey in ("partial","semi_partial") and not ctrl:
                    messagebox.showwarning("Controls Required","Partial / Semi-Partial requires selecting control variables.",parent=self); return
                # Collect all unique vars involved in the selected pairs
                all_vars=list(dict.fromkeys(v for a,b in free_pairs for v in (a,b)))
                r_all=CorrelationEngine.compute(self.df,mkey,all_vars,controls=ctrl)
                filtered=[]
                for (a,b) in free_pairs:
                    match=next((p for p in r_all["pairs"] if (p["x"]==a and p["y"]==b) or (p["x"]==b and p["y"]==a)),None)
                    if match: filtered.append(match)
                if not filtered:
                    messagebox.showwarning("No Results","None of the selected pairs had enough valid observations.",parent=self); return
                r_all["pairs"]=filtered; r_all["variables"]=all_vars; self.results=r_all
            else:
                self.results=CorrelationEngine.compute(self.df,mkey,cols,controls=ctrl)
            # ─────────────────────────────────────────────────────────────

            r=self.results; ts=datetime.now().strftime("%Y-%m-%d %H:%M:%S"); pairs=r["pairs"]; sym=mi["symbol"]
            def fmt(v,d=4): return f"{v:.{d}f}" if v is not None else "N/A"
            txt=(f"╔══════════════════════════════════════╗\n║  CORRELATION ANALYSIS RESULTS        ║\n╚══════════════════════════════════════╝\n\n"
                 f"Method:    {mi['label']}  ({sym})\nTimestamp: {ts}\nVariables: {r['n_vars']}   Pairs: {len(pairs)}\n")
            if r.get("notes"): txt+=f"Notes:     {r['notes']}\n"
            if self.composite_info:
                txt+="\nComposite Scores Used:\n"
                for pt,info in self.composite_info.items(): txt+=f"  {pt} ({info['mode']} of {len(info['items'])} items)\n"
            if self.sidebar.pair_mode_var.get() and self.free_pairs:
                txt+=f"\nFree Pair Mode: {len(self.free_pairs)} pair(s) selected\n"
            active_descs={c:self.var_descriptions[c] for c in r["variables"] if c in self.var_descriptions and (self.var_descriptions[c].get("label") or self.var_descriptions[c].get("desc"))}
            if active_descs:
                txt+="\nVariable Descriptions:\n"
                for col,info in active_descs.items():
                    lbl=info.get("label",""); desc=info.get("desc","")
                    txt+=f"  {col}"
                    if lbl: txt+=f"  ({lbl})"
                    if desc: txt+=f"\n    {desc[:80]}{'…' if len(desc)>80 else ''}"
                    txt+="\n"
            txt+="\n┌──────────────────────────────────────┐\n│ PAIRWISE RESULTS                     │\n├──────────────────────────────────────┤\n"
            for p in pairs:
                flag="✓ Significant" if p["sig"] else "✗ Not Significant"
                ci_str=(f"[{fmt(p['ci_lower'])}, {fmt(p['ci_upper'])}]" if p['ci_lower'] is not None else "N/A")
                txt+=(f"\n  {p['x']}  ↔  {p['y']}\n  {'─'*36}\n"
                      f"  {sym}:               {fmt(p['r'])}\n  {sym}² (shared var): {fmt(p['r_sq'])}\n"
                      f"  t / z-stat:      {fmt(p['t_stat'],3)}\n  p-value:         {fmt(p['p'])}  {p['stars']}\n"
                      f"  95% CI:          {ci_str}\n  Effect Size:     {p['effect']}\n"
                      f"  Direction:       {p['direction']}\n  N:               {p['n']}\n  Result:          {flag}\n")
                if p.get("notes"): txt+=f"  Notes:           {p['notes']}\n"
            txt+="\n└──────────────────────────────────────┘\n\n"
            vars_=r["variables"]; mat=r["corr_matrix"]; col_w=10
            txt+=f"┌──────────────────────────────────────┐\n│ CORRELATION MATRIX  ({sym})           │\n├──────────────────────────────────────┤\n"
            hdr_r="  "+"".ljust(14)
            for v in vars_: hdr_r+=str(v)[:col_w].ljust(col_w+2)
            txt+=hdr_r+"\n"
            for v in vars_:
                row_str=f"  {str(v)[:14].ljust(14)}"
                for v2 in vars_: row_str+=("  1.000  " if v==v2 else f"{mat.loc[v,v2]:+.3f}").ljust(col_w+2)
                txt+=row_str+"\n"
            txt+="└──────────────────────────────────────┘\n\n* p<.05  ** p<.01  *** p<.001  ns=not significant\n\n✓ Ready to export PDF  |  Click 📉 View Scatter Plots"
            self._set_results(txt)
            self.sidebar.export_btn.configure(state="normal"); self.sidebar.print_btn.configure(state="normal"); self.sidebar.plot_btn.configure(state="normal")
            if self.df is not None: self.data_table.display_data(self.df,selected_cols=r["variables"])
            best=max(pairs,key=lambda x:abs(x['r']))
            self.stat_label.configure(text=f"{mi['label']}  |  Strongest: {best['x']} ↔ {best['y']}  {sym}={best['r']:.4f}  {best['stars']}  {best['effect']}  N={best['n']}")
            self.sidebar.status_label.configure(text=f"✓ {len(pairs)} pair(s)\n{mi['label'][:24]}")
            messagebox.showinfo("Done",f"Correlation computed!\nMethod: {mi['label']}\nPairs:  {len(pairs)}\n\nStrongest: {best['x']} ↔ {best['y']}\n{sym} = {best['r']:.4f}  ({best['effect']}, {best['stars']})")
        except Exception as e: messagebox.showerror("Error",f"Computation failed:\n{e}")

    def _pdf_kwargs(self):
        mi=METHODS[self.sidebar.method_key]
        return dict(title=self.sidebar.title_entry.get().strip() or f"{mi['label']} Analysis",
                    subtitle=self.sidebar.subtitle_entry.get().strip(),byline=self.sidebar.author_entry.get().strip(),
                    composite_info=self.composite_info,var_descriptions=self.var_descriptions if self.var_descriptions else None,
                    critical_value=self.critical_value_info)

    def export_pdf(self):
        if self.results is None: messagebox.showwarning("No Results","Compute first."); return
        mi=METHODS[self.sidebar.method_key]
        fp=filedialog.asksaveasfilename(defaultextension=".pdf",filetypes=[("PDF","*.pdf")],
                                         initialfile=f"Correlation_{mi['symbol']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        if not fp: return
        try:
            PDFReport.generate(self.results,self.desc_text.get("1.0","end-1c"),fp,**self._pdf_kwargs())
            self.file_label.configure(text=f"Last saved: {fp}"); self.sidebar.status_label.configure(text=f"✓ PDF Saved\n{os.path.basename(fp)}")
            messagebox.showinfo("Saved",f"PDF exported:\n{fp}")
        except Exception as e: messagebox.showerror("Error",f"Export failed:\n{e}")

    def print_report(self):
        if self.results is None: messagebox.showwarning("No Results","Compute first."); return
        tmp=tempfile.NamedTemporaryFile(suffix=".pdf",delete=False); tmp_path=tmp.name; tmp.close()
        try:
            PDFReport.generate(self.results,self.desc_text.get("1.0","end-1c"),tmp_path,**self._pdf_kwargs())
            if sys.platform=="win32":
                import ctypes; ctypes.windll.shell32.ShellExecuteW(None,"print",tmp_path,None,None,0)
            elif sys.platform=="darwin": subprocess.run(["lpr",tmp_path])
            else:
                res=subprocess.run(["lpr",tmp_path])
                if res.returncode!=0: subprocess.run(["xdg-open",tmp_path])
            self.sidebar.status_label.configure(text="🖨️ Sent to printer")
        except Exception as e:
            messagebox.showerror("Print Error",f"Printing failed:\n{e}")
            try: os.unlink(tmp_path)
            except: pass

    def export_dataset(self):
        if self.df is None: messagebox.showwarning("No Data","No dataset to export."); return
        fp=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx"),("CSV","*.csv")],initialfile=f"Dataset_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        if not fp: return
        try:
            self.df.to_csv(fp,index=False) if fp.endswith(".csv") else self.df.to_excel(fp,index=False)
            messagebox.showinfo("Saved",f"Dataset exported:\n{os.path.basename(fp)}")
        except Exception as e: messagebox.showerror("Error",f"Export failed:\n{e}")

    def open_settings(self):
        if HAS_SETTINGS: SettingsWindow(self,self)

    def apply_settings(self):
        if not HAS_SETTINGS: return
        sm=SettingsManager(); fb,fc,fh,fm,fbt,ft=sm.fonts; ff=sm.font_family
        self.results_text.configure(font=(ff,fm),wrap=sm.wrap_mode)
        self.sidebar.status_label.configure(font=(ff,ft)); self.stat_label.configure(font=(ff,ft,"bold"))
        self.file_label.configure(font=(ff,ft)); self.sidebar.configure(width=sm.sidebar_width)

    def clear_all(self):
        self.df=None; self.raw_df=None; self.selected_cols=None; self.controls=[]
        self.results=None; self.composite_info=None; self.critical_value_info=None
        self.var_descriptions={}; self.free_pairs=[]
        self.sidebar.pair_mode_var.set(False)
        self.sidebar.pair_summary_lbl.configure(text="No pairs selected",text_color=TEXT_SEC)
        self._update_pair_compare_controls()
        self.data_table.display_data(None); self.desc_text.delete("1.0","end")
        self._set_results("Cleared. Import data to begin.")
        self.stat_label.configure(text=""); self.sidebar.status_label.configure(text="")
        for btn in (self.sidebar.export_btn,self.sidebar.print_btn,self.sidebar.export_data_btn,
                    self.sidebar.select_btn,self.sidebar.plot_btn,self.sidebar.view_data_btn,
                    self.sidebar.edit_data_btn,self.sidebar.var_desc_btn,self.sidebar.open_pair_selector_btn):
            btn.configure(state="disabled")

    def toggle_theme(self):
        if self.dark_mode:
            ctk.set_appearance_mode("light"); self.sidebar.theme_btn.configure(text="🌙  Dark Mode"); self.dark_mode=False
        else:
            ctk.set_appearance_mode("dark"); self.sidebar.theme_btn.configure(text="☀️  Light Mode"); self.dark_mode=True


# ─── Entry ────────────────────────────────────────────────────────────────────

def main():
    app = PearsonRApp()
    app.mainloop()

if __name__ == "__main__":
    main()