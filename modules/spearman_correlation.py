"""
Spearman's Rank Correlation Analyzer
A comprehensive GUI application for statistical correlation analysis

Requirements:
pip install customtkinter pandas numpy scipy matplotlib openpyxl seaborn pillow
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import numpy as np
from scipy import stats
from scipy.stats import spearmanr, pearsonr, kendalltau
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
from datetime import datetime
import json
import os

# Set appearance mode and default color theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class SpearmanAnalyzer(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Window configuration
        self.title("Spearman's Rank Correlation Analyzer")
        self.geometry("1400x800")
        
        # Data storage
        self.data = None
        self.results = {}
        self.current_file = None
        
        # Create main layout
        self.create_menu_bar()
        self.create_main_layout()
        self.create_status_bar()
        
    def create_menu_bar(self):
        """Create menu bar"""
        menubar = ctk.CTkFrame(self, height=40, fg_color="transparent")
        menubar.pack(side="top", fill="x", padx=10, pady=5)
        
        # File menu
        file_btn = ctk.CTkButton(menubar, text="File", width=80,
                                 command=self.show_file_menu)
        file_btn.pack(side="left", padx=5)
        
        # Analysis menu
        analysis_btn = ctk.CTkButton(menubar, text="Analysis", width=80,
                                     command=self.show_analysis_menu)
        analysis_btn.pack(side="left", padx=5)
        
        # View menu
        view_btn = ctk.CTkButton(menubar, text="View", width=80,
                                command=self.show_view_menu)
        view_btn.pack(side="left", padx=5)
        
        # Help menu
        help_btn = ctk.CTkButton(menubar, text="Help", width=80,
                                command=self.show_help)
        help_btn.pack(side="left", padx=5)
        
        # Theme toggle
        self.theme_switch = ctk.CTkSwitch(menubar, text="Dark Mode",
                                         command=self.toggle_theme)
        self.theme_switch.pack(side="right", padx=10)
        self.theme_switch.select()
        
    def create_main_layout(self):
        """Create main application layout"""
        # Main container
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Left sidebar (30% width)
        self.sidebar = ctk.CTkFrame(main_container, width=300)
        self.sidebar.pack(side="left", fill="y", padx=(0, 5))
        self.sidebar.pack_propagate(False)
        
        # Center workspace (50% width)
        self.workspace = ctk.CTkFrame(main_container)
        self.workspace.pack(side="left", fill="both", expand=True, padx=5)
        
        # Right panel (20% width)
        self.right_panel = ctk.CTkFrame(main_container, width=250)
        self.right_panel.pack(side="right", fill="y", padx=(5, 0))
        self.right_panel.pack_propagate(False)
        
        # Setup panels
        self.setup_sidebar()
        self.setup_workspace()
        self.setup_right_panel()
        
    def setup_sidebar(self):
        """Setup left sidebar with data controls"""
        # Title
        title = ctk.CTkLabel(self.sidebar, text="Data Input",
                            font=ctk.CTkFont(size=18, weight="bold"))
        title.pack(pady=10)
        
        # Data input section
        input_frame = ctk.CTkFrame(self.sidebar)
        input_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(input_frame, text="Load Data:",
                    font=ctk.CTkFont(size=14)).pack(pady=5)
        
        ctk.CTkButton(input_frame, text="📁 Import CSV",
                     command=self.import_csv).pack(fill="x", padx=10, pady=2)
        
        ctk.CTkButton(input_frame, text="📊 Import Excel",
                     command=self.import_excel).pack(fill="x", padx=10, pady=2)
        
        ctk.CTkButton(input_frame, text="✏️ Manual Entry",
                     command=self.manual_entry).pack(fill="x", padx=10, pady=2)
        
        ctk.CTkButton(input_frame, text="🎲 Generate Sample Data",
                     command=self.generate_sample_data).pack(fill="x", padx=10, pady=2)
        
        # Variable selection
        var_frame = ctk.CTkFrame(self.sidebar)
        var_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(var_frame, text="Variable Selection:",
                    font=ctk.CTkFont(size=14)).pack(pady=5)
        
        ctk.CTkLabel(var_frame, text="X Variable:").pack(anchor="w", padx=10)
        self.x_var = ctk.CTkComboBox(var_frame, values=["No data loaded"])
        self.x_var.pack(fill="x", padx=10, pady=2)
        
        ctk.CTkLabel(var_frame, text="Y Variable:").pack(anchor="w", padx=10, pady=(5,0))
        self.y_var = ctk.CTkComboBox(var_frame, values=["No data loaded"])
        self.y_var.pack(fill="x", padx=10, pady=2)
        
        # Analysis options
        options_frame = ctk.CTkFrame(self.sidebar)
        options_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(options_frame, text="Analysis Options:",
                    font=ctk.CTkFont(size=14)).pack(pady=5)
        
        self.correlation_method = ctk.CTkComboBox(options_frame,
            values=["Spearman", "Pearson", "Kendall"])
        self.correlation_method.set("Spearman")
        self.correlation_method.pack(fill="x", padx=10, pady=2)
        
        self.alternative = ctk.CTkComboBox(options_frame,
            values=["two-sided", "less", "greater"])
        self.alternative.set("two-sided")
        self.alternative.pack(fill="x", padx=10, pady=2)
        
        ctk.CTkLabel(options_frame, text="Significance Level:").pack(anchor="w", padx=10)
        self.alpha = ctk.CTkEntry(options_frame, placeholder_text="0.05")
        self.alpha.insert(0, "0.05")
        self.alpha.pack(fill="x", padx=10, pady=2)
        
        # Run analysis button
        ctk.CTkButton(self.sidebar, text="🔬 Run Analysis",
                     command=self.run_analysis,
                     height=40,
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=20, padx=10, fill="x")
        
    def setup_workspace(self):
        """Setup center workspace for data display and plots"""
        # Tabview for different views
        self.tabview = ctk.CTkTabview(self.workspace)
        self.tabview.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Add tabs
        self.tabview.add("Data View")
        self.tabview.add("Scatter Plot")
        self.tabview.add("Rank Plot")
        self.tabview.add("Distribution")
        
        # Data view tab
        self.setup_data_view()
        
    def setup_data_view(self):
        """Setup data view table"""
        data_frame = self.tabview.tab("Data View")
        
        # Create treeview for data display
        tree_frame = ctk.CTkFrame(data_frame)
        tree_frame.pack(fill="both", expand=True)
        
        # Scrollbars
        vsb = ctk.CTkScrollbar(tree_frame, orientation="vertical")
        vsb.pack(side="right", fill="y")
        
        hsb = ctk.CTkScrollbar(tree_frame, orientation="horizontal")
        hsb.pack(side="bottom", fill="x")
        
        # Create data info label
        self.data_info = ctk.CTkLabel(data_frame, text="No data loaded",
                                     font=ctk.CTkFont(size=12))
        self.data_info.pack(pady=5)
        
    def setup_right_panel(self):
        """Setup right panel for results display"""
        title = ctk.CTkLabel(self.right_panel, text="Analysis Results",
                           font=ctk.CTkFont(size=18, weight="bold"))
        title.pack(pady=10)
        
        # Results textbox
        self.results_text = ctk.CTkTextbox(self.right_panel, wrap="word")
        self.results_text.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Export buttons
        export_frame = ctk.CTkFrame(self.right_panel)
        export_frame.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkButton(export_frame, text="📄 Export Report",
                     command=self.export_report).pack(fill="x", pady=2)
        
        ctk.CTkButton(export_frame, text="💾 Save Results",
                     command=self.save_results).pack(fill="x", pady=2)
        
    def create_status_bar(self):
        """Create status bar at bottom"""
        self.status_bar = ctk.CTkLabel(self, text="Ready",
                                      anchor="w",
                                      font=ctk.CTkFont(size=10))
        self.status_bar.pack(side="bottom", fill="x", padx=10, pady=5)
        
    # Data handling methods
    def import_csv(self):
        """Import CSV file"""
        filename = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            try:
                self.data = pd.read_csv(filename)
                self.current_file = filename
                self.update_data_view()
                self.update_variable_lists()
                self.status_bar.configure(text=f"Loaded: {os.path.basename(filename)}")
                messagebox.showinfo("Success", f"Loaded {len(self.data)} rows")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV: {str(e)}")
                
    def import_excel(self):
        """Import Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            try:
                self.data = pd.read_excel(filename)
                self.current_file = filename
                self.update_data_view()
                self.update_variable_lists()
                self.status_bar.configure(text=f"Loaded: {os.path.basename(filename)}")
                messagebox.showinfo("Success", f"Loaded {len(self.data)} rows")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel: {str(e)}")
                
    def manual_entry(self):
        """Open manual data entry window"""
        ManualEntryWindow(self)
        
    def generate_sample_data(self):
        """Open sample data generator"""
        SampleDataGenerator(self)
        
    def update_data_view(self):
        """Update the data view with current data"""
        if self.data is not None:
            info = f"Data: {len(self.data)} rows × {len(self.data.columns)} columns"
            self.data_info.configure(text=info)
            
    def update_variable_lists(self):
        """Update variable selection dropdowns"""
        if self.data is not None:
            numeric_cols = self.data.select_dtypes(include=[np.number]).columns.tolist()
            self.x_var.configure(values=numeric_cols)
            self.y_var.configure(values=numeric_cols)
            if len(numeric_cols) >= 2:
                self.x_var.set(numeric_cols[0])
                self.y_var.set(numeric_cols[1])
                
    # Analysis methods
    def run_analysis(self):
        """Run correlation analysis"""
        if self.data is None:
            messagebox.showwarning("No Data", "Please load data first")
            return
            
        x_col = self.x_var.get()
        y_col = self.y_var.get()
        
        if x_col == y_col:
            messagebox.showwarning("Invalid Selection", "Please select different variables")
            return
            
        try:
            # Get data
            x = self.data[x_col].dropna()
            y = self.data[y_col].dropna()
            
            # Align data (remove rows where either is NA)
            valid_idx = self.data[[x_col, y_col]].dropna().index
            x = self.data.loc[valid_idx, x_col]
            y = self.data.loc[valid_idx, y_col]
            
            if len(x) < 3:
                messagebox.showerror("Error", "Need at least 3 valid data points")
                return
                
            # Calculate correlation
            method = self.correlation_method.get()
            alpha_val = float(self.alpha.get())
            
            if method == "Spearman":
                corr, pval = spearmanr(x, y)
                method_name = "Spearman's Rho"
            elif method == "Pearson":
                corr, pval = pearsonr(x, y)
                method_name = "Pearson's r"
            else:  # Kendall
                corr, pval = kendalltau(x, y)
                method_name = "Kendall's Tau"
                
            # Store results
            self.results = {
                'method': method_name,
                'correlation': corr,
                'p_value': pval,
                'n': len(x),
                'alpha': alpha_val,
                'x_col': x_col,
                'y_col': y_col,
                'x_data': x,
                'y_data': y
            }
            
            # Display results
            self.display_results()
            
            # Update plots
            self.update_plots()
            
            self.status_bar.configure(text="Analysis complete")
            
        except Exception as e:
            messagebox.showerror("Analysis Error", f"Error: {str(e)}")
            
    def display_results(self):
        """Display analysis results"""
        r = self.results
        
        # Interpret correlation strength
        corr_abs = abs(r['correlation'])
        if corr_abs < 0.3:
            strength = "weak"
        elif corr_abs < 0.7:
            strength = "moderate"
        else:
            strength = "strong"
            
        direction = "positive" if r['correlation'] > 0 else "negative"
        
        # Significance
        significant = "significant" if r['p_value'] < r['alpha'] else "not significant"
        
        # Format results text
        results_text = f"""
{r['method']} Analysis Results
{'=' * 40}

Variables:
  X: {r['x_col']}
  Y: {r['y_col']}

Sample Size: n = {r['n']}

Correlation Coefficient: {r['correlation']:.4f}
P-value: {r['p_value']:.4f}
Significance Level: α = {r['alpha']}

Interpretation:
  Strength: {strength.capitalize()}
  Direction: {direction.capitalize()}
  Significance: {significant.capitalize()}

Conclusion:
There is a {strength} {direction} correlation
between {r['x_col']} and {r['y_col']}.

The result is {significant} at α = {r['alpha']}.
"""
        
        self.results_text.delete("1.0", "end")
        self.results_text.insert("1.0", results_text)
        
    def update_plots(self):
        """Update all plots"""
        self.create_scatter_plot()
        self.create_rank_plot()
        self.create_distribution_plot()
        
    def create_scatter_plot(self):
        """Create scatter plot"""
        tab = self.tabview.tab("Scatter Plot")
        
        # Clear previous plot
        for widget in tab.winfo_children():
            widget.destroy()
            
        r = self.results
        
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.scatter(r['x_data'], r['y_data'], alpha=0.6, s=50)
        
        # Add regression line
        z = np.polyfit(r['x_data'], r['y_data'], 1)
        p = np.poly1d(z)
        x_line = np.linspace(r['x_data'].min(), r['x_data'].max(), 100)
        ax.plot(x_line, p(x_line), "r--", alpha=0.8, label='Trend line')
        
        ax.set_xlabel(r['x_col'], fontsize=12)
        ax.set_ylabel(r['y_col'], fontsize=12)
        ax.set_title(f"{r['method']}: ρ = {r['correlation']:.4f}, p = {r['p_value']:.4f}",
                    fontsize=14, fontweight='bold')
        ax.legend()
        ax.grid(True, alpha=0.3)
        
        canvas = FigureCanvasTkAgg(fig, tab)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        
    def create_rank_plot(self):
        """Create rank plot"""
        tab = self.tabview.tab("Rank Plot")
        
        for widget in tab.winfo_children():
            widget.destroy()
            
        r = self.results
        
        # Calculate ranks
        x_ranks = stats.rankdata(r['x_data'])
        y_ranks = stats.rankdata(r['y_data'])
        
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.scatter(x_ranks, y_ranks, alpha=0.6, s=50, c='green')
        
        # Add diagonal line
        max_rank = max(x_ranks.max(), y_ranks.max())
        ax.plot([0, max_rank], [0, max_rank], 'r--', alpha=0.5)
        
        ax.set_xlabel(f"Rank of {r['x_col']}", fontsize=12)
        ax.set_ylabel(f"Rank of {r['y_col']}", fontsize=12)
        ax.set_title("Rank Scatter Plot", fontsize=14, fontweight='bold')
        ax.grid(True, alpha=0.3)
        
        canvas = FigureCanvasTkAgg(fig, tab)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        
    def create_distribution_plot(self):
        """Create distribution plots"""
        tab = self.tabview.tab("Distribution")
        
        for widget in tab.winfo_children():
            widget.destroy()
            
        r = self.results
        
        fig, axes = plt.subplots(2, 2, figsize=(10, 8))
        
        # X distribution
        axes[0, 0].hist(r['x_data'], bins=20, alpha=0.7, color='blue', edgecolor='black')
        axes[0, 0].set_title(f"Distribution of {r['x_col']}")
        axes[0, 0].set_xlabel(r['x_col'])
        axes[0, 0].set_ylabel("Frequency")
        
        # Y distribution
        axes[0, 1].hist(r['y_data'], bins=20, alpha=0.7, color='green', edgecolor='black')
        axes[0, 1].set_title(f"Distribution of {r['y_col']}")
        axes[0, 1].set_xlabel(r['y_col'])
        axes[0, 1].set_ylabel("Frequency")
        
        # Q-Q plot for X
        stats.probplot(r['x_data'], dist="norm", plot=axes[1, 0])
        axes[1, 0].set_title(f"Q-Q Plot: {r['x_col']}")
        
        # Q-Q plot for Y
        stats.probplot(r['y_data'], dist="norm", plot=axes[1, 1])
        axes[1, 1].set_title(f"Q-Q Plot: {r['y_col']}")
        
        plt.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, tab)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        
    # Export methods
    def export_report(self):
        """Export full report"""
        if not self.results:
            messagebox.showwarning("No Results", "Run analysis first")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    f.write(self.results_text.get("1.0", "end"))
                messagebox.showinfo("Success", "Report exported successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {str(e)}")
                
    def save_results(self):
        """Save results to JSON"""
        if not self.results:
            messagebox.showwarning("No Results", "Run analysis first")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                # Prepare serializable results
                save_data = {
                    'method': self.results['method'],
                    'correlation': float(self.results['correlation']),
                    'p_value': float(self.results['p_value']),
                    'n': int(self.results['n']),
                    'alpha': float(self.results['alpha']),
                    'x_col': self.results['x_col'],
                    'y_col': self.results['y_col'],
                    'timestamp': datetime.now().isoformat()
                }
                
                with open(filename, 'w') as f:
                    json.dump(save_data, f, indent=2)
                    
                messagebox.showinfo("Success", "Results saved successfully")
            except Exception as e:
                messagebox.showerror("Error", f"Save failed: {str(e)}")
                
    # Menu methods
    def show_file_menu(self):
        """Show file menu options"""
        menu = ctk.CTkToplevel(self)
        menu.title("File Menu")
        menu.geometry("200x250")
        
        ctk.CTkButton(menu, text="Import CSV", command=self.import_csv).pack(pady=5)
        ctk.CTkButton(menu, text="Import Excel", command=self.import_excel).pack(pady=5)
        ctk.CTkButton(menu, text="Export Report", command=self.export_report).pack(pady=5)
        ctk.CTkButton(menu, text="Save Results", command=self.save_results).pack(pady=5)
        
    def show_analysis_menu(self):
        """Show analysis menu"""
        if self.data is None:
            messagebox.showinfo("Info", "Load data first to run analysis")
        else:
            self.run_analysis()
            
    def show_view_menu(self):
        """Show view options"""
        messagebox.showinfo("View", "Use tabs to switch between different views")
        
    def show_help(self):
        """Show help information"""
        help_text = """
Spearman's Rank Correlation Analyzer
====================================

Quick Start:
1. Load data (CSV or Excel)
2. Select X and Y variables
3. Click 'Run Analysis'

Features:
- Multiple correlation methods
- Interactive plots
- Export results

For more help, visit the documentation.
"""
        messagebox.showinfo("Help", help_text)
        
    def toggle_theme(self):
        """Toggle between light and dark mode"""
        if self.theme_switch.get():
            ctk.set_appearance_mode("dark")
        else:
            ctk.set_appearance_mode("light")


class ManualEntryWindow(ctk.CTkToplevel):
    """Window for manual data entry"""
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        self.title("Manual Data Entry")
        self.geometry("600x400")
        
        # Instructions
        ctk.CTkLabel(self, text="Enter data (comma-separated for each row)",
                    font=ctk.CTkFont(size=14)).pack(pady=10)
        
        # Entry area
        self.text_area = ctk.CTkTextbox(self, wrap="word")
        self.text_area.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Example
        example = "X,Y\n1,2\n2,4\n3,6\n4,8\n5,10"
        self.text_area.insert("1.0", example)
        
        # Buttons
        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(pady=10)
        
        ctk.CTkButton(btn_frame, text="Load Data",
                     command=self.load_data).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel",
                     command=self.destroy).pack(side="left", padx=5)
        
    def load_data(self):
        """Load manually entered data"""
        try:
            from io import StringIO
            data_str = self.text_area.get("1.0", "end")
            self.parent.data = pd.read_csv(StringIO(data_str))
            self.parent.update_data_view()
            self.parent.update_variable_lists()
            messagebox.showinfo("Success", "Data loaded successfully")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load data: {str(e)}")


class SampleDataGenerator(ctk.CTkToplevel):
    """Window for generating sample data"""
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        self.title("Generate Sample Data")
        self.geometry("400x350")
        
        # Sample size
        ctk.CTkLabel(self, text="Sample Size:").pack(pady=5)
        self.sample_size = ctk.CTkEntry(self, placeholder_text="100")
        self.sample_size.insert(0, "100")
        self.sample_size.pack(pady=5)
        
        # Correlation strength
        ctk.CTkLabel(self, text="Correlation Strength:").pack(pady=5)
        self.correlation = ctk.CTkComboBox(self,
            values=["Strong Positive (0.8)", "Moderate Positive (0.5)",
                   "Weak Positive (0.2)", "No Correlation (0)",
                   "Weak Negative (-0.2)", "Moderate Negative (-0.5)",
                   "Strong Negative (-0.8)"])
        self.correlation.set("Moderate Positive (0.5)")
        self.correlation.pack(pady=5)
        
        # Distribution
        ctk.CTkLabel(self, text="Distribution:").pack(pady=5)
        self.distribution = ctk.CTkComboBox(self,
            values=["Normal", "Uniform", "Exponential"])
        self.distribution.set("Normal")
        self.distribution.pack(pady=5)
        
        # Generate button
        ctk.CTkButton(self, text="Generate Data",
                     command=self.generate).pack(pady=20)
        
    def generate(self):
        """Generate sample data"""
        try:
            n = int(self.sample_size.get())
            
            # Extract correlation value
            corr_str = self.correlation.get()
            rho = float(corr_str.split('(')[1].split(')')[0])
            
            # Generate correlated data
            mean = [0, 0]
            cov = [[1, rho], [rho, 1]]
            
            if self.distribution.get() == "Normal":
                data = np.random.multivariate_normal(mean, cov, n)
            else:
                # For other distributions, generate and rank
                x = np.random.randn(n)
                y = rho * x + np.sqrt(1 - rho**2) * np.random.randn(n)
                data = np.column_stack([x, y])
            
            # Create DataFrame
            self.parent.data = pd.DataFrame(data, columns=['X', 'Y'])
            self.parent.update_data_view()
            self.parent.update_variable_lists()
            
            messagebox.showinfo("Success", f"Generated {n} data points")
            self.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Generation failed: {str(e)}")


def main():
    """Main entry point"""
    app = SpearmanAnalyzer()
    app.mainloop()


if __name__ == "__main__":
    main()