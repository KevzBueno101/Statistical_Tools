"""
Modern Statistical Analysis App with Regression (Column-Based Entry Version)
Features: 
- CustomTkinter Modern GUI
- JASP-style APA Tables (Model Summary, ANOVA, Coefficients)
- Column-Based Raw Data Entry & CSV Import
- Metadata (Title/Author)
- Timestamped & Source-tracked PDF Footer
- Robust Statistical Engine with Error Handling
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
from scipy import stats
from sklearn.linear_model import LinearRegression
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.colors import black, grey
import os
from datetime import datetime

# Appearance Settings
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class RegressionApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Modern Statistical Analysis System (APA Style)")
        self.geometry("1200x850")

        # Data State
        self.data = None
        self.results = {}
        self.current_file_path = "No Data Loaded"
        
        # Column entry widgets
        self.column_entries = {}
        self.num_rows = 10  # Default number of rows

        self.setup_ui()

    def setup_ui(self):
        # Grid Configuration
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # --- Sidebar ---
        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar.grid_rowconfigure(10, weight=1)

        ctk.CTkLabel(self.sidebar, text="REGRESSION APP", font=ctk.CTkFont(size=20, weight="bold")).grid(row=0, column=0, padx=20, pady=(20, 20))

        # Metadata
        ctk.CTkLabel(self.sidebar, text="Report Metadata:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")
        self.title_entry = ctk.CTkEntry(self.sidebar, placeholder_text="Report Title")
        self.title_entry.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        self.author_entry = ctk.CTkEntry(self.sidebar, placeholder_text="Author Name")
        self.author_entry.grid(row=3, column=0, padx=20, pady=5, sticky="ew")

        # Regression Type
        ctk.CTkLabel(self.sidebar, text="Regression Type:", font=ctk.CTkFont(weight="bold")).grid(row=4, column=0, padx=20, pady=(20, 0), sticky="w")
        self.reg_type_menu = ctk.CTkOptionMenu(
            self.sidebar,
            values=["Single Regression", "Multiple Regression", "Mediating Regression"],
            command=self.on_reg_type_change
        )
        self.reg_type_menu.grid(row=5, column=0, padx=20, pady=5, sticky="ew")

        # Actions
        ctk.CTkLabel(self.sidebar, text="Actions:", font=ctk.CTkFont(weight="bold")).grid(row=6, column=0, padx=20, pady=(20, 0), sticky="w")
        self.import_btn = ctk.CTkButton(self.sidebar, text="Import CSV File", command=self.import_csv)
        self.import_btn.grid(row=7, column=0, padx=20, pady=10, sticky="ew")
        self.run_btn = ctk.CTkButton(self.sidebar, text="Run Analysis", fg_color="#2c3e50", hover_color="#34495e", command=self.run_regression)
        self.run_btn.grid(row=8, column=0, padx=20, pady=10, sticky="ew")
        self.export_btn = ctk.CTkButton(self.sidebar, text="Export PDF Report", command=self.export_pdf)
        self.export_btn.grid(row=9, column=0, padx=20, pady=10, sticky="ew")

        # Theme
        ctk.CTkLabel(self.sidebar, text="Appearance:", font=ctk.CTkFont(weight="bold")).grid(row=11, column=0, padx=20, pady=(10, 0), sticky="w")
        self.theme_menu = ctk.CTkOptionMenu(self.sidebar, values=["System", "Light", "Dark"], command=lambda m: ctk.set_appearance_mode(m))
        self.theme_menu.grid(row=12, column=0, padx=20, pady=(5, 20), sticky="ew")

        # --- Main Content ---
        self.main = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(1, weight=1)

        # Top Panel: Data Input & Variable Selection
        self.top_panel = ctk.CTkFrame(self.main)
        self.top_panel.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 20))
        self.top_panel.grid_columnconfigure(0, weight=3)
        self.top_panel.grid_columnconfigure(1, weight=2)

        # Raw Data Input with Columns
        self.data_input_frame = ctk.CTkFrame(self.top_panel, fg_color="transparent")
        self.data_input_frame.grid(row=0, column=0, sticky="nsew", padx=15, pady=15)
        
        # Header with controls
        header_frame = ctk.CTkFrame(self.data_input_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 5))
        ctk.CTkLabel(header_frame, text="Raw Data Entry:", font=ctk.CTkFont(weight="bold")).pack(side="left")
        
        # Column setup controls
        ctk.CTkButton(header_frame, text="Setup Columns", width=120, command=self.setup_columns_dialog).pack(side="right", padx=5)
        
        # Scrollable frame for data entry
        self.data_scroll = ctk.CTkScrollableFrame(self.data_input_frame, height=180)
        self.data_scroll.pack(fill="both", expand=True, pady=(5, 10))
        
        # Initial setup
        self.create_default_columns()
        
        # Load & Clear buttons
        btn_row = ctk.CTkFrame(self.data_input_frame, fg_color="transparent")
        btn_row.pack(anchor="e")
        self.load_raw_btn = ctk.CTkButton(btn_row, text="Load Manual Data", height=32, command=self.load_manual_data)
        self.load_raw_btn.pack(side="left", padx=(0, 5))
        self.clear_btn = ctk.CTkButton(btn_row, text="Clear Data", height=32, fg_color="#c0392b", hover_color="#e74c3c", command=self.clear_data)
        self.clear_btn.pack(side="left")

        # Variable Selection
        self.var_select_frame = ctk.CTkFrame(self.top_panel, fg_color="transparent")
        self.var_select_frame.grid(row=0, column=1, sticky="nsew", padx=15, pady=15)
        
        ctk.CTkLabel(self.var_select_frame, text="Dependent (Y):", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.dep_menu = ctk.CTkOptionMenu(self.var_select_frame, values=["None"])
        self.dep_menu.pack(fill="x", pady=(5, 15))

        ctk.CTkLabel(self.var_select_frame, text="Independent (X):", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.ind_scroll = ctk.CTkScrollableFrame(self.var_select_frame, height=120)
        self.ind_scroll.pack(fill="both", expand=True, pady=5)
        self.ind_checks = {}

        # Mediator selector (hidden by default, shown for Mediating Regression)
        self.med_frame = ctk.CTkFrame(self.var_select_frame, fg_color="transparent")
        ctk.CTkLabel(self.med_frame, text="Mediator (M):", font=ctk.CTkFont(weight="bold")).pack(anchor="w")
        self.med_menu = ctk.CTkOptionMenu(self.med_frame, values=["None"])
        self.med_menu.pack(fill="x", pady=(5, 0))

        # Bottom Panel: Results
        self.results_frame = ctk.CTkFrame(self.main)
        self.results_frame.grid(row=1, column=0, sticky="nsew")
        self.results_frame.grid_columnconfigure(0, weight=1)
        self.results_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self.results_frame, text="Analysis Results (JASP/APA Style)", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=15, pady=(15, 5), sticky="w")
        self.output_box = ctk.CTkTextbox(self.results_frame, font=("Consolas", 13))
        self.output_box.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")

    # --- Column Setup ---
    
    def setup_columns_dialog(self):
        """Dialog to configure column names and number of rows"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Setup Data Columns")
        dialog.geometry("400x300")
        dialog.grab_set()
        
        ctk.CTkLabel(dialog, text="Column Names (comma-separated):", font=ctk.CTkFont(weight="bold")).pack(padx=20, pady=(20, 5), anchor="w")
        
        # Get current column names
        current_cols = ",".join(self.column_entries.keys()) if self.column_entries else "StudyHours,Attendance,ExamScore"
        cols_entry = ctk.CTkEntry(dialog, width=360)
        cols_entry.pack(padx=20, pady=5)
        cols_entry.insert(0, current_cols)
        
        ctk.CTkLabel(dialog, text="Number of Rows:", font=ctk.CTkFont(weight="bold")).pack(padx=20, pady=(20, 5), anchor="w")
        rows_entry = ctk.CTkEntry(dialog, width=360)
        rows_entry.pack(padx=20, pady=5)
        rows_entry.insert(0, str(self.num_rows))
        
        def apply_setup():
            try:
                col_names = [c.strip() for c in cols_entry.get().split(",") if c.strip()]
                num_rows = int(rows_entry.get())
                
                if not col_names:
                    messagebox.showwarning("Input Error", "Please enter at least one column name.")
                    return
                if num_rows < 1 or num_rows > 100:
                    messagebox.showwarning("Input Error", "Number of rows must be between 1 and 100.")
                    return
                
                self.num_rows = num_rows
                self.create_data_columns(col_names)
                dialog.destroy()
            except ValueError:
                messagebox.showerror("Input Error", "Number of rows must be a valid integer.")
        
        ctk.CTkButton(dialog, text="Apply", command=apply_setup).pack(pady=20)
    
    def create_default_columns(self):
        """Create default column setup"""
        default_cols = ["StudyHours", "Attendance", "ExamScore"]
        default_data = [
            [10, 95, 85],
            [5, 80, 70],
            [15, 100, 95],
            [2, 60, 55],
            [8, 85, 75],
            [12, 90, 88],
            [4, 70, 62],
            ["", "", ""],
            ["", "", ""],
            ["", "", ""]
        ]
        self.create_data_columns(default_cols, default_data)
    
    def create_data_columns(self, column_names, default_data=None):
        """Create entry boxes organized by columns"""
        # Clear existing entries
        for widget in self.data_scroll.winfo_children():
            widget.destroy()
        self.column_entries = {}
        
        # Create container frame
        container = ctk.CTkFrame(self.data_scroll, fg_color="transparent")
        container.pack(fill="both", expand=True)
        
        # Create columns
        for col_idx, col_name in enumerate(column_names):
            col_frame = ctk.CTkFrame(container, fg_color="transparent")
            col_frame.grid(row=0, column=col_idx, padx=5, sticky="nsew")
            
            # Column header
            ctk.CTkLabel(col_frame, text=col_name, font=ctk.CTkFont(weight="bold")).pack(pady=(0, 5))
            
            # Entry boxes for this column
            entries = []
            for row_idx in range(self.num_rows):
                entry = ctk.CTkEntry(col_frame, width=100)
                entry.pack(pady=2)
                
                # Fill with default data if provided
                if default_data and row_idx < len(default_data) and col_idx < len(default_data[row_idx]):
                    value = default_data[row_idx][col_idx]
                    if value != "":
                        entry.insert(0, str(value))
                
                entries.append(entry)
            
            self.column_entries[col_name] = entries

    # --- Regression Type Handler ---

    def on_reg_type_change(self, reg_type):
        """Show/hide mediator selector based on regression type."""
        if reg_type == "Mediating Regression":
            self.med_frame.pack(anchor="w", fill="x", pady=(10, 0))
        else:
            self.med_frame.pack_forget()

    def clear_data(self):
        """Clear all entry boxes and reset loaded data."""
        for entries in self.column_entries.values():
            for entry in entries:
                entry.delete(0, "end")
        self.data = None
        self.results = {}
        self.output_box.delete("1.0", "end")

    # --- Data Logic ---

    def load_manual_data(self):
        """Load data from the column entry boxes"""
        if not self.column_entries:
            messagebox.showwarning("Input Error", "No columns configured.")
            return
        
        try:
            data_dict = {}
            
            for col_name, entries in self.column_entries.items():
                values = []
                for entry in entries:
                    val = entry.get().strip()
                    if val:  # Only include non-empty values
                        try:
                            values.append(float(val))
                        except ValueError:
                            messagebox.showerror("Data Error", f"Invalid numeric value in column '{col_name}': {val}")
                            return
                data_dict[col_name] = values
            
            # Check if all columns have data
            if not any(data_dict.values()):
                messagebox.showwarning("Input Error", "Please enter some data.")
                return
            
            # Check if all columns have the same length
            lengths = [len(v) for v in data_dict.values()]
            if len(set(lengths)) > 1:
                messagebox.showerror("Data Error", "All columns must have the same number of values.\n" + 
                                   "\n".join([f"{k}: {len(v)} values" for k, v in data_dict.items()]))
                return
            
            self.data = pd.DataFrame(data_dict)
            self.current_file_path = "Manual Entry"
            self.refresh_vars()
            messagebox.showinfo("Success", f"Manual data loaded successfully.\n{len(self.data)} rows, {len(self.data.columns)} columns.")
        
        except Exception as e:
            messagebox.showerror("Data Error", f"Could not load data:\n{str(e)}")

    def import_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not path: return
        try:
            self.data = pd.read_csv(path)
            self.current_file_path = path
            self.refresh_vars()
            messagebox.showinfo("Success", f"Imported: {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to read file:\n{str(e)}")

    def refresh_vars(self):
        if self.data is None: return
        cols = self.data.select_dtypes(include=[np.number]).columns.tolist()
        if not cols:
            messagebox.showerror("Data Error", "No numeric columns found in the dataset.")
            return
        
        self.dep_menu.configure(values=cols)
        self.dep_menu.set(cols[0])

        self.med_menu.configure(values=cols)
        self.med_menu.set(cols[-1] if len(cols) > 1 else cols[0])
        
        for w in self.ind_scroll.winfo_children(): w.destroy()
        self.ind_checks = {}
        for c in cols:
            cb = ctk.CTkCheckBox(self.ind_scroll, text=c)
            cb.pack(anchor="w", padx=10, pady=2)
            self.ind_checks[c] = cb

    # --- Regression Engine ---

    def run_regression(self):
        if self.data is None:
            self.load_manual_data()
            if self.data is None:
                return

        reg_type = self.reg_type_menu.get()
        if reg_type == "Mediating Regression":
            self.run_mediation()
        else:
            self.run_standard(reg_type)

    def run_standard(self, reg_type):
        """Single or Multiple Regression."""
        y_var = self.dep_menu.get()
        x_vars = [name for name, cb in self.ind_checks.items() if cb.get()]

        if not x_vars:
            messagebox.showwarning("Selection Error", "Please select at least one Independent Variable.")
            return
        if y_var in x_vars:
            messagebox.showerror("Selection Error", "Dependent variable cannot be an Independent variable.")
            return
        if reg_type == "Single Regression" and len(x_vars) != 1:
            messagebox.showerror("Selection Error", "Single Regression requires exactly 1 Independent Variable.")
            return

        try:
            # 1. Data Preparation
            df = self.data[[y_var] + x_vars].dropna()
            if len(df) < len(x_vars) + 2:
                messagebox.showerror("Stats Error", f"Insufficient data. Need at least {len(x_vars)+2} rows after removing NaNs.")
                return

            X = df[x_vars].values
            y = df[y_var].values
            n, p = X.shape

            # 2. Model Fitting
            model = LinearRegression().fit(X, y)
            y_hat = model.predict(X)

            # 3. Sum of Squares
            sse = np.sum((y - y_hat)**2)
            ssr = np.sum((y_hat - np.mean(y))**2)
            sst = np.sum((y - np.mean(y))**2)
            
            # 4. Model Metrics
            r2 = model.score(X, y)
            adj_r2 = 1 - (1 - r2) * (n - 1) / (n - p - 1)
            rmse = np.sqrt(sse / (n - p - 1))

            # 5. ANOVA (F-test)
            msr = ssr / p
            mse = sse / (n - p - 1)
            f_stat = msr / mse if mse != 0 else 0
            f_p = stats.f.sf(f_stat, p, n - p - 1)

            # 6. Coefficients (T-test)
            # Add intercept for matrix math
            X_mat = np.hstack([np.ones((n, 1)), X])
            # Use pseudo-inverse for robustness against multicollinearity
            xtx_inv = np.linalg.pinv(X_mat.T @ X_mat)
            se_b = np.sqrt(np.diagonal(xtx_inv) * mse)
            
            params = np.append(model.intercept_, model.coef_)
            t_stats = params / se_b
            p_vals = [2 * (1 - stats.t.cdf(np.abs(t), n - p - 1)) for t in t_stats]

            # 7. Store Results
            self.results = {
                "type": self.reg_type_menu.get(),
                "title": self.title_entry.get() or "Regression Analysis Report",
                "author": self.author_entry.get() or "Anonymous",
                "dep": y_var, "ind": x_vars, "n": n,
                "r": np.sqrt(r2), "r2": r2, "adj_r2": adj_r2, "rmse": rmse,
                "f": f_stat, "f_p": f_p, "df_reg": p, "df_res": n - p - 1,
                "b": params, "se": se_b, "t": t_stats, "p": p_vals
            }
            self.show_results()

        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred during regression:\n{str(e)}")

    def run_mediation(self):
        """Mediation Analysis: X → M → Y (Baron & Kenny / Sobel test)."""
        y_var = self.dep_menu.get()
        x_var = None
        for name, cb in self.ind_checks.items():
            if cb.get():
                x_var = name
                break
        m_var = self.med_menu.get()

        if not x_var:
            messagebox.showwarning("Selection Error", "Please select exactly 1 Independent Variable for mediation.")
            return
        if x_var == y_var:
            messagebox.showerror("Selection Error", "Independent and Dependent variables cannot be the same.")
            return
        if m_var == y_var or m_var == x_var:
            messagebox.showerror("Selection Error", "Mediator must be different from both X and Y.")
            return

        try:
            df = self.data[[y_var, x_var, m_var]].dropna()
            n = len(df)
            if n < 5:
                messagebox.showerror("Stats Error", "Insufficient data for mediation analysis (need at least 5 rows).")
                return

            X = df[x_var].values
            M = df[m_var].values
            Y = df[y_var].values

            # Path c: Total effect  Y ~ X
            slope_c, intercept_c, r_c, p_c, se_c = stats.linregress(X, Y)
            total_effect = slope_c

            # Path a: X ~ M  (X predicts M)
            slope_a, intercept_a, r_a, p_a, se_a = stats.linregress(X, M)

            # Path b + c': M + X ~ Y  (both predict Y)
            X_mat = np.column_stack([np.ones(n), X, M])
            beta = np.linalg.lstsq(X_mat, Y, rcond=None)[0]
            direct_effect = beta[1]   # c' = effect of X on Y controlling M
            slope_b = beta[2]         # b = effect of M on Y controlling X

            # Residual standard errors for Sobel
            Y_hat_full = X_mat @ beta
            mse_full = np.sum((Y - Y_hat_full) ** 2) / (n - 3)
            var_c = mse_full * np.linalg.inv(X_mat.T @ X_mat)
            se_direct = np.sqrt(var_c[1, 1])
            se_b_med = np.sqrt(var_c[2, 2])

            # Indirect effect
            indirect = slope_a * slope_b
            se_sobel = np.sqrt(slope_b**2 * se_a**2 + slope_a**2 * se_b_med**2)
            z_sobel = indirect / se_sobel if se_sobel != 0 else 0
            p_sobel = 2 * (1 - stats.norm.cdf(abs(z_sobel)))

            # R² for each path
            r2_c = r_c ** 2
            r2_a = r_a ** 2
            Y_hat_pathb = intercept_a + slope_a * X
            ss_res_b = np.sum((M - Y_hat_pathb) ** 2)
            ss_tot_b = np.sum((M - np.mean(M)) ** 2)
            # R² for full model (X+M → Y)
            ss_res_full = np.sum((Y - Y_hat_full) ** 2)
            ss_tot_y = np.sum((Y - np.mean(Y)) ** 2)
            r2_full = 1 - ss_res_full / ss_tot_y if ss_tot_y != 0 else 0

            # Proportion mediated
            prop_mediated = indirect / total_effect if total_effect != 0 else 0

            self.results = {
                "type": "Mediating Regression",
                "title": self.title_entry.get() or "Mediation Analysis Report",
                "author": self.author_entry.get() or "Anonymous",
                "dep": y_var, "ind": [x_var], "mediator": m_var, "n": n,
                # Path c (total effect)
                "total_b": total_effect, "total_se": se_c, "total_t": slope_c / se_c if se_c != 0 else 0,
                "total_p": p_c, "total_r2": r2_c,
                # Path a
                "a_b": slope_a, "a_se": se_a, "a_t": slope_a / se_a if se_a != 0 else 0,
                "a_p": p_a, "a_r2": r2_a,
                # Path b
                "b_b": slope_b, "b_se": se_b_med,
                # Path c' (direct effect)
                "direct_b": direct_effect, "direct_se": se_direct,
                "direct_t": direct_effect / se_direct if se_direct != 0 else 0,
                # Sobel test
                "indirect": indirect, "se_sobel": se_sobel,
                "z_sobel": z_sobel, "p_sobel": p_sobel,
                "r2_full": r2_full, "prop_mediated": prop_mediated,
            }
            self.show_results()

        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred during mediation analysis:\n{str(e)}")

    def show_results(self):
        res = self.results
        self.output_box.delete("1.0", "end")

        if res.get("type") == "Mediating Regression":
            self._show_mediation_results(res)
        else:
            self._show_standard_results(res)

    def _show_standard_results(self, res):
        out = f"REPORT: {res['title']}\nBy: {res['author']}\n"
        out += f"Type: {res['type']}\n"
        out += "="*65 + "\n\n"
        
        # Model Summary
        out += "Model Summary\n" + "-"*65 + "\n"
        out += f"{'Model':<10} {'R':<10} {'R²':<10} {'Adj. R²':<10} {'RMSE':<10}\n"
        out += f"{'1':<10} {res['r']:<10.3f} {res['r2']:<10.3f} {res['adj_r2']:<10.3f} {res['rmse']:<10.3f}\n\n"
        
        # ANOVA
        out += "ANOVA\n" + "-"*65 + "\n"
        out += f"{'Source':<15} {'df':<10} {'F':<10} {'p':<10}\n"
        fp_str = "< .001" if res['f_p'] < 0.001 else f"{res['f_p']:.3f}"
        out += f"{'Regression':<15} {res['df_reg']:<10} {res['f']:<10.3f} {fp_str:<10}\n"
        out += f"{'Residual':<15} {res['df_res']:<10}\n\n"
        
        # Coefficients
        out += "Coefficients\n" + "-"*75 + "\n"
        out += f"{'Variable':<20} {'B':<10} {'Std. Error':<15} {'t':<10} {'p':<10}\n"
        labels = ["(Intercept)"] + res['ind']
        for i, label in enumerate(labels):
            p_str = "< .001" if res['p'][i] < 0.001 else f"{res['p'][i]:.3f}"
            out += f"{label:<20} {res['b'][i]:<10.3f} {res['se'][i]:<15.3f} {res['t'][i]:<10.3f} {p_str:<10}\n"
        
        self.output_box.insert("1.0", out)

    def _show_mediation_results(self, res):
        out = f"REPORT: {res['title']}\nBy: {res['author']}\n"
        out += f"Type: Mediating Regression\n"
        out += f"Model: {res['ind'][0]} (X) → {res['mediator']} (M) → {res['dep']} (Y)\n"
        out += f"N = {res['n']}\n"
        out += "="*70 + "\n\n"

        # Path summaries
        def _fmt(val, width=10):
            return f"{val:<{width}.3f}"
        def _fp(p):
            return "< .001" if p < 0.001 else f"{p:.3f}"

        out += "Path Analysis\n" + "-"*70 + "\n"
        out += f"{'Path':<12} {'B':<10} {'SE':<10} {'t/z':<10} {'p':<10} {'R²':<10}\n"
        out += f"{'c  (Total)':<12} {_fmt(res['total_b'])} {_fmt(res['total_se'])} {_fmt(res['total_t'])} {_fp(res['total_p']):<10} {_fmt(res['total_r2'])}\n"
        out += f"{'a  (X→M)':<12} {_fmt(res['a_b'])} {_fmt(res['a_se'])} {_fmt(res['a_t'])} {_fp(res['a_p']):<10} {_fmt(res['a_r2'])}\n"
        out += f"{'b  (M→Y)':<12} {_fmt(res['b_b'])} {_fmt(res['b_se'])}       \n"
        out += f"{'c\' (Direct)':<12} {_fmt(res['direct_b'])} {_fmt(res['direct_se'])} {_fmt(res['direct_t'])}       \n\n"

        # Sobel test
        out += "Indirect Effect (Sobel Test)\n" + "-"*70 + "\n"
        out += f"  Indirect effect (a × b) : {res['indirect']:.4f}\n"
        out += f"  SE (Sobel)              : {res['se_sobel']:.4f}\n"
        out += f"  z-statistic             : {res['z_sobel']:.4f}\n"
        out += f"  p-value                 : {_fp(res['p_sobel'])}\n"
        out += f"  Proportion mediated     : {res['prop_mediated']:.4f}\n"
        out += f"  Full model R²           : {res['r2_full']:.4f}\n\n"

        # Interpretation
        sig = "significant" if res['p_sobel'] < 0.05 else "not significant"
        out += f"Interpretation: The indirect effect is {sig} (p = {_fp(res['p_sobel'])}).\n"
        out += f"  {res['prop_mediated']*100:.1f}% of the total effect of {res['ind'][0]} on {res['dep']}\n"
        out += f"  is mediated through {res['mediator']}.\n"

        self.output_box.insert("1.0", out)

    # --- PDF Export ---

    def export_pdf(self):
        if not self.results:
            messagebox.showwarning("Warning", "No results to export. Run analysis first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not path: return
        try:
            self.generate_pdf(path)
            messagebox.showinfo("Success", "PDF Report generated successfully.")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to create PDF:\n{str(e)}")

    def generate_pdf(self, filename):
        doc = SimpleDocTemplate(filename, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = []
        res = self.results

        # Header
        elements.append(Paragraph(f"<b>{res['title']}</b>", styles["Title"]))
        elements.append(Paragraph(f"Author: {res['author']} | Type: {res.get('type', 'Regression')}", styles["Normal"]))
        elements.append(Spacer(1, 20))

        # APA Table Style
        apa_table_style = TableStyle([
            ('LINEABOVE', (0, 0), (-1, 0), 1, black),
            ('LINEBELOW', (0, 0), (-1, 0), 0.5, black),
            ('LINEBELOW', (0, -1), (-1, -1), 1, black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ])

        if res.get("type") == "Mediating Regression":
            self._pdf_mediation(doc, elements, res, styles, apa_table_style)
        else:
            self._pdf_standard(elements, res, styles, apa_table_style)

        # Footer (shared)
        elements.append(Spacer(1, 200))
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        footer = f"<i><font color='grey' size='8'>File: {self.current_file_path} | Generated: {ts}</font></i>"
        elements.append(Paragraph(footer, styles["Normal"]))

        doc.build(elements)

    def _pdf_standard(self, elements, res, styles, apa_table_style):
        # Table 1: Model Summary
        elements.append(Paragraph(f"<i>Model Summary - Dependent Variable: {res['dep']}</i>", styles["Normal"]))
        data1 = [["Model", "R", "R²", "Adjusted R²", "RMSE"], ["1", f"{res['r']:.3f}", f"{res['r2']:.3f}", f"{res['adj_r2']:.3f}", f"{res['rmse']:.3f}"]]
        t1 = Table(data1, colWidths=[60, 80, 80, 100, 80])
        t1.setStyle(apa_table_style); elements.append(t1); elements.append(Spacer(1, 20))

        # Table 2: ANOVA
        elements.append(Paragraph("<i>ANOVA Table</i>", styles["Normal"]))
        fp = "< .001" if res['f_p'] < 0.001 else f"{res['f_p']:.3f}"
        data2 = [["Source", "df", "F", "p"], ["Regression", str(res['df_reg']), f"{res['f']:.3f}", fp], ["Residual", str(res['df_res']), "", ""]]
        t2 = Table(data2, colWidths=[100, 60, 80, 80])
        t2.setStyle(apa_table_style); elements.append(t2); elements.append(Spacer(1, 20))

        # Table 3: Coefficients
        elements.append(Paragraph("<i>Coefficients</i>", styles["Normal"]))
        data3 = [["Variable", "B", "Std. Error", "t", "p"]]
        labels = ["(Intercept)"] + res['ind']
        for i, label in enumerate(labels):
            p_val = "< .001" if res['p'][i] < 0.001 else f"{res['p'][i]:.3f}"
            data3.append([label, f"{res['b'][i]:.3f}", f"{res['se'][i]:.3f}", f"{res['t'][i]:.3f}", p_val])
        t3 = Table(data3, colWidths=[140, 80, 80, 80, 80])
        t3.setStyle(apa_table_style); elements.append(t3)

    def _pdf_mediation(self, doc, elements, res, styles, apa_table_style):
        elements.append(Paragraph(f"<i>Mediation Model: {res['ind'][0]} (X) → {res['mediator']} (M) → {res['dep']} (Y) | N = {res['n']}</i>", styles["Normal"]))
        elements.append(Spacer(1, 10))

        # Path Analysis Table
        elements.append(Paragraph("<i>Path Analysis</i>", styles["Normal"]))
        fp_c = "< .001" if res['total_p'] < 0.001 else f"{res['total_p']:.3f}"
        fp_a = "< .001" if res['a_p'] < 0.001 else f"{res['a_p']:.3f}"
        data_pa = [
            ["Path", "B", "SE", "t/z", "p", "R²"],
            ["c  (Total effect)", f"{res['total_b']:.3f}", f"{res['total_se']:.3f}", f"{res['total_t']:.3f}", fp_c, f"{res['total_r2']:.3f}"],
            ["a  (X → M)", f"{res['a_b']:.3f}", f"{res['a_se']:.3f}", f"{res['a_t']:.3f}", fp_a, f"{res['a_r2']:.3f}"],
            ["b  (M → Y)", f"{res['b_b']:.3f}", f"{res['b_se']:.3f}", "", "", ""],
            ["c' (Direct effect)", f"{res['direct_b']:.3f}", f"{res['direct_se']:.3f}", f"{res['direct_t']:.3f}", "", ""],
        ]
        t_pa = Table(data_pa, colWidths=[120, 70, 70, 60, 60, 60])
        t_pa.setStyle(apa_table_style); elements.append(t_pa); elements.append(Spacer(1, 20))

        # Sobel Test Table
        elements.append(Paragraph("<i>Sobel Test for Indirect Effect</i>", styles["Normal"]))
        fp_s = "< .001" if res['p_sobel'] < 0.001 else f"{res['p_sobel']:.3f}"
        data_sobel = [
            ["Indirect Effect", "SE", "z", "p", "Prop. Mediated"],
            [f"{res['indirect']:.4f}", f"{res['se_sobel']:.4f}", f"{res['z_sobel']:.4f}", fp_s, f"{res['prop_mediated']:.4f}"],
        ]
        t_sobel = Table(data_sobel, colWidths=[100, 80, 80, 80, 100])
        t_sobel.setStyle(apa_table_style); elements.append(t_sobel); elements.append(Spacer(1, 15))

        sig = "significant" if res['p_sobel'] < 0.05 else "not significant"
        interp = (f"The indirect effect is {sig} (p = {fp_s}). "
                  f"{res['prop_mediated']*100:.1f}% of the total effect of {res['ind'][0]} on {res['dep']} "
                  f"is mediated through {res['mediator']}.")
        elements.append(Paragraph(f"<i>Interpretation: {interp}</i>", styles["Normal"]))

if __name__ == "__main__":
    app = RegressionApp()
    app.mainloop()