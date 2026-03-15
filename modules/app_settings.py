"""
app_settings.py — Shared Settings Module
Place this file in the same folder as the three stat apps.
"""

import json, os, tkinter as tk
import customtkinter as ctk

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              ".stat_suite_settings.json")

DEFAULT_SETTINGS = {
    "theme":          "Dark",
    "accent_color":   "Teal",
    "font_size":      "Medium",
    "font_family":    "Segoe UI",
    "decimal_places": "2",
    "apa_pvalue":     True,
    "sidebar_width":  "Normal",
    "results_wrap":   "Word",
    "show_tips":      True,
}

ACCENTS = {
    "Teal":   {"accent": "#00c9a7", "hover": "#009e82"},
    "Blue":   {"accent": "#4e9eff", "hover": "#3b7ddd"},
    "Purple": {"accent": "#a855f7", "hover": "#7e22ce"},
    "Orange": {"accent": "#f59e0b", "hover": "#d97706"},
    "Rose":   {"accent": "#f43f5e", "hover": "#be123c"},
    "Green":  {"accent": "#22c55e", "hover": "#16a34a"},
}

FONT_SCALES = {
    "Small":       {"body": 11, "card": 13, "head": 20, "mono": 10, "btn": 11, "tiny": 9},
    "Medium":      {"body": 13, "card": 15, "head": 22, "mono": 12, "btn": 13, "tiny": 11},
    "Large":       {"body": 15, "card": 17, "head": 24, "mono": 13, "btn": 15, "tiny": 12},
    "Extra Large": {"body": 17, "card": 19, "head": 26, "mono": 14, "btn": 17, "tiny": 13},
}

SIDEBAR_WIDTHS = {"Compact": 190, "Normal": 230, "Wide": 280}
WRAP_MODES     = {"None": "none", "Word": "word", "Char": "char"}

BG   = "#0d1117"
CARD = "#161b22"
PNL  = "#1c2230"
INP  = "#1e2736"
PRI  = "#e6edf3"
SEC  = "#8b949e"
BDR  = "#30363d"
DNG  = "#ef4444"


# ─── SettingsManager ──────────────────────────────────────────────────────────

class SettingsManager:
    _instance  = None
    _callbacks = []

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._data = dict(DEFAULT_SETTINGS)
            cls._instance._load()
        return cls._instance

    def _load(self):
        try:
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE) as f:
                    self._data.update(json.load(f))
        except Exception:
            pass

    def save(self):
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(self._data, f, indent=2)
        except Exception:
            pass

    def reset(self):
        self._data = dict(DEFAULT_SETTINGS)
        self.save()
        self._notify()

    def get(self, key, default=None):
        return self._data.get(key, default if default is not None
                              else DEFAULT_SETTINGS.get(key))

    def set(self, key, value):
        self._data[key] = value
        self.save()
        self._notify()

    @property
    def accent(self):
        return ACCENTS.get(self._data.get("accent_color", "Teal"), ACCENTS["Teal"])["accent"]

    @property
    def accent_hover(self):
        return ACCENTS.get(self._data.get("accent_color", "Teal"), ACCENTS["Teal"])["hover"]

    @property
    def font_sizes(self):
        return FONT_SCALES.get(self._data.get("font_size", "Medium"), FONT_SCALES["Medium"])

    @property
    def fonts(self):
        """Returns (body, card, head, mono, btn, tiny) as a tuple for easy unpacking."""
        s = self.font_sizes
        return (s["body"], s["card"], s["head"], s["mono"], s["btn"], s["tiny"])

    @property
    def font_family(self):
        return self._data.get("font_family", "Segoe UI")

    @property
    def sidebar_width(self):
        return SIDEBAR_WIDTHS.get(self._data.get("sidebar_width", "Normal"), 230)

    @property
    def wrap_mode(self):
        return WRAP_MODES.get(self._data.get("results_wrap", "Word"), "word")

    @property
    def decimal_places(self):
        try:
            return int(self._data.get("decimal_places", 2))
        except Exception:
            return 2

    @property
    def apa_pvalue(self):
        v = self._data.get("apa_pvalue", True)
        # Handle saved as bool, int (0/1), or string
        if isinstance(v, bool): return v
        if isinstance(v, int): return v != 0
        if isinstance(v, str): return v.lower() not in ("false", "0", "no", "")
        return True

    @property
    def show_tips(self):
        v = self._data.get("show_tips", True)
        if isinstance(v, bool): return v
        if isinstance(v, int): return v != 0
        if isinstance(v, str): return v.lower() not in ("false", "0", "no", "")
        return True

    def register(self, cb):
        if cb not in self._callbacks:
            self._callbacks.append(cb)

    def unregister(self, cb):
        if cb in self._callbacks:
            self._callbacks.remove(cb)

    def _notify(self):
        for cb in list(self._callbacks):
            try: cb()
            except Exception: pass

    def fmt(self, value):
        dp = self.decimal_places
        return f"{round(float(value), dp):.{dp}f}"

    def fmt_p(self, p):
        if self.apa_pvalue:
            return "< .001" if p < 0.001 else f"= {p:.3f}"
        return f"= {p:.{self.decimal_places}f}"


# ─── SettingsWindow ───────────────────────────────────────────────────────────

class SettingsWindow(ctk.CTkToplevel):

    def __init__(self, master, app_instance, **kw):
        super().__init__(master, **kw)
        self.title("Settings")
        self.geometry("580x720")
        self.resizable(True, True)
        self.configure(bg=BG)
        self.minsize(500, 520)

        self.app = app_instance
        self.sm  = SettingsManager()

        # local StringVar/BooleanVar — only committed on Apply
        self.v_theme   = ctk.StringVar(value=self.sm.get("theme"))
        self.v_accent  = ctk.StringVar(value=self.sm.get("accent_color"))
        self.v_family  = ctk.StringVar(value=self.sm.get("font_family"))
        self.v_fsize   = ctk.StringVar(value=self.sm.get("font_size"))
        self.v_sidebar = ctk.StringVar(value=self.sm.get("sidebar_width"))
        self.v_decimal = ctk.StringVar(value=str(self.sm.get("decimal_places")))
        self.v_wrap    = ctk.StringVar(value=self.sm.get("results_wrap"))
        self.v_apa     = ctk.BooleanVar(value=bool(self.sm.get("apa_pvalue")))
        self.v_tips    = ctk.BooleanVar(value=bool(self.sm.get("show_tips")))

        self.transient(master)
        self.lift()
        self.focus_force()
        self.attributes("-topmost", True)

        self._build()

    def _close(self):
        try: self.attributes("-topmost", False)
        except Exception: pass
        self.destroy()

    # ── Build skeleton ────────────────────────────────────────────────────────

    def _build(self):
        sm = self.sm

        # ── Header (fixed) ────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=sm.accent, height=54)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="⚙   Settings",
                 font=("Segoe UI", 20, "bold"),
                 fg="#0d1117", bg=sm.accent).pack(side="left", padx=20)

        # ── Bottom button bar (fixed) ─────────────────────────────────────────
        btn_bar = ctk.CTkFrame(self, fg_color=CARD, corner_radius=0, height=58)
        btn_bar.pack(fill="x", side="bottom")
        btn_bar.pack_propagate(False)

        ctk.CTkButton(btn_bar, text="↺  Defaults",
                      fg_color=PNL, hover_color=BDR,
                      text_color=SEC, font=("Segoe UI", 12),
                      height=36, corner_radius=7,
                      command=self._reset
                      ).pack(side="left", padx=12, pady=11)

        ctk.CTkButton(btn_bar, text="✕  Cancel",
                      fg_color=PNL, hover_color=DNG,
                      text_color=SEC, font=("Segoe UI", 12),
                      height=36, corner_radius=7,
                      command=self._close
                      ).pack(side="right", padx=(4, 12), pady=11)

        ctk.CTkButton(btn_bar, text="✓  Apply & Close",
                      fg_color=sm.accent, hover_color=sm.accent_hover,
                      text_color="#0d1117", font=("Segoe UI", 13, "bold"),
                      height=36, corner_radius=7,
                      command=self._apply
                      ).pack(side="right", padx=4, pady=11)

        # ── Canvas + scrollbar (fills remaining space) ────────────────────────
        wrapper = tk.Frame(self, bg=BG)
        wrapper.pack(fill="both", expand=True)

        self._canvas = tk.Canvas(wrapper, bg=BG, highlightthickness=0,
                                 bd=0)
        vsb = tk.Scrollbar(wrapper, orient="vertical",
                           command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        # Inner frame that holds all content
        self._inner = tk.Frame(self._canvas, bg=BG)
        self._win_id = self._canvas.create_window(
            (0, 0), window=self._inner, anchor="nw"
        )

        # Keep inner frame width = canvas width
        self._canvas.bind("<Configure>", self._on_canvas_resize)
        self._inner.bind("<Configure>", self._on_inner_resize)

        # Mouse-wheel scrolling
        self._canvas.bind_all("<MouseWheel>",
                              lambda e: self._canvas.yview_scroll(
                                  int(-1*(e.delta/120)), "units"))

        # ── Populate content ──────────────────────────────────────────────────
        self._populate()

    def _on_canvas_resize(self, event):
        self._canvas.itemconfig(self._win_id, width=event.width)

    def _on_inner_resize(self, event):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    # ── Content ───────────────────────────────────────────────────────────────

    def _populate(self):
        p = self._inner   # shorthand

        # ══════════════════════════════════════════
        # SECTION 1 — Appearance
        # ══════════════════════════════════════════
        self._sec_title(p, "🎨   Appearance")

        self._field_label(p, "Theme")
        self._seg_row(p, self.v_theme, ["Dark", "Light", "System"])

        self._field_label(p, "Accent Color")
        self._accent_row(p)

        self._field_label(p, "Font Family")
        self._seg_row(p, self.v_family,
                      ["Segoe UI", "Calibri", "Helvetica", "Arial", "Courier New"])

        self._field_label(p, "Font Size")
        self._seg_row(p, self.v_fsize,
                      ["Small", "Medium", "Large", "Extra Large"])

        self._field_label(p, "Sidebar Width")
        self._seg_row(p, self.v_sidebar, ["Compact", "Normal", "Wide"])

        self._gap(p, 14)

        # ══════════════════════════════════════════
        # SECTION 2 — Output
        # ══════════════════════════════════════════
        self._sec_title(p, "📐   Output & Statistics")

        self._field_label(p, "Decimal Places")
        self._seg_row(p, self.v_decimal, ["2", "3", "4"])

        self._field_label(p, "Results Text Wrap")
        self._seg_row(p, self.v_wrap, ["None", "Word", "Char"])

        self._gap(p, 14)

        # ══════════════════════════════════════════
        # SECTION 3 — Options
        # ══════════════════════════════════════════
        self._sec_title(p, "🔧   Options")

        self._toggle(p, self.v_apa,
                     "APA p-value format",
                     "Shows  p < .001  instead of  p = 0.0003")

        self._toggle(p, self.v_tips,
                     "Show inline tips",
                     "Display helper hints inside input panels")

        self._gap(p, 24)

    # ── Widget builders ────────────────────────────────────────────────────────

    def _sec_title(self, parent, text):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", padx=14, pady=(18, 6))
        # accent bar
        tk.Frame(row, bg=self.sm.accent, width=4, height=22).pack(
            side="left", padx=(0, 10))
        tk.Label(row, text=text,
                 font=("Segoe UI", 14, "bold"),
                 fg=PRI, bg=BG).pack(side="left")

    def _field_label(self, parent, text):
        tk.Label(parent, text=text,
                 font=("Segoe UI", 11),
                 fg=SEC, bg=BG).pack(anchor="w", padx=18, pady=(8, 3))

    def _seg_row(self, parent, var, options):
        """Horizontal segmented button row using ctk buttons inside a tk frame."""
        container = tk.Frame(parent, bg=PNL, pady=4)
        container.pack(fill="x", padx=18, pady=(0, 6))

        buttons = []

        def activate(opt):
            var.set(opt)
            for b in buttons:
                active = b.cget("text") == opt
                b.configure(
                    fg_color=self.sm.accent if active else "transparent",
                    text_color="#0d1117" if active else SEC,
                    font=("Segoe UI", 11, "bold") if active
                         else ("Segoe UI", 11),
                )

        for opt in options:
            active = (var.get() == opt)
            b = ctk.CTkButton(
                container,
                text=opt,
                fg_color=self.sm.accent if active else "transparent",
                hover_color=self.sm.accent_hover,
                text_color="#0d1117" if active else SEC,
                font=("Segoe UI", 11, "bold") if active else ("Segoe UI", 11),
                height=36, corner_radius=6, border_width=0,
                command=lambda o=opt: activate(o),
            )
            b.pack(side="left", padx=4, pady=4)
            buttons.append(b)

    def _accent_row(self, parent):
        container = tk.Frame(parent, bg=PNL)
        container.pack(fill="x", padx=18, pady=(0, 6))

        self._dot_widgets = {}

        def pick(name):
            self.v_accent.set(name)
            for n, (dot, lbl) in self._dot_widgets.items():
                sel = (n == name)
                dot.configure(
                    relief="solid" if sel else "flat",
                    bd=3 if sel else 0,
                    highlightbackground=PRI if sel else PNL,
                    highlightthickness=3 if sel else 0,
                )

        for name, pal in ACCENTS.items():
            col = tk.Frame(container, bg=PNL)
            col.pack(side="left", padx=12, pady=10)

            selected = (self.v_accent.get() == name)

            # Use a tk Canvas circle as the dot
            dot_canvas = tk.Canvas(col, width=36, height=36,
                                   bg=PNL, highlightthickness=3,
                                   highlightbackground=PRI if selected else PNL,
                                   cursor="hand2")
            dot_canvas.pack()
            dot_canvas.create_oval(2, 2, 34, 34,
                                   fill=pal["accent"], outline="")
            dot_canvas.bind("<Button-1>", lambda e, n=name: pick(n))

            lbl = tk.Label(col, text=name,
                           font=("Segoe UI", 9),
                           fg=SEC, bg=PNL)
            lbl.pack(pady=(3, 0))
            lbl.bind("<Button-1>", lambda e, n=name: pick(n))

            self._dot_widgets[name] = (dot_canvas, lbl)

    def _toggle(self, parent, var, label, hint=""):
        row = tk.Frame(parent, bg=PNL)
        row.pack(fill="x", padx=18, pady=(0, 8))

        txt = tk.Frame(row, bg=PNL)
        txt.pack(side="left", fill="x", expand=True, padx=12, pady=10)

        tk.Label(txt, text=label,
                 font=("Segoe UI", 12, "bold"),
                 fg=PRI, bg=PNL).pack(anchor="w")
        if hint:
            tk.Label(txt, text=hint,
                     font=("Segoe UI", 10),
                     fg=SEC, bg=PNL).pack(anchor="w", pady=(2, 0))

        ctk.CTkSwitch(row, variable=var, text="",
                      progress_color=self.sm.accent,
                      button_color=PRI,
                      onvalue=True, offvalue=False,
                      width=46).pack(side="right", padx=14, pady=10)

    def _gap(self, parent, h=10):
        tk.Frame(parent, bg=BG, height=h).pack(fill="x")

    # ── Actions ────────────────────────────────────────────────────────────────

    def _apply(self):
        sm = self.sm

        # Save all values — cast booleans explicitly to avoid 0/1 int issues
        sm.set("theme",          self.v_theme.get())
        sm.set("accent_color",   self.v_accent.get())
        sm.set("font_family",    self.v_family.get())
        sm.set("font_size",      self.v_fsize.get())
        sm.set("sidebar_width",  self.v_sidebar.get())
        sm.set("decimal_places", str(self.v_decimal.get()))
        sm.set("results_wrap",   self.v_wrap.get())
        sm.set("apa_pvalue",     bool(self.v_apa.get()))
        sm.set("show_tips",      bool(self.v_tips.get()))

        # Apply theme globally
        ctk.set_appearance_mode(
            {"Dark": "dark", "Light": "light", "System": "system"}
            .get(sm.get("theme"), "dark")
        )

        # Notify the host app to refresh its widgets
        if hasattr(self.app, "apply_settings"):
            self.app.apply_settings()

        self._close()

    def _reset(self):
        self.sm.reset()
        self._close()
        SettingsWindow(self.master, self.app)