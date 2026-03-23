"""
Unified Statistical Analysis Desktop Application with Database Integration
Main Menu - Entry Point

Features:
- Multiple statistical analysis tools
- SQLite database for storing analysis history
- History viewing and management
- Export functionality

Author: Statistical Analysis Suite  
Version: 2.0.0 (Database Integrated)
"""

import customtkinter as ctk
from tkinter import messagebox, filedialog
import sys
import os
from datetime import datetime

# Add modules directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'modules'))

# Import statistical modules
from modules import anova_analyzer
from modules import cronbach_alpha
from modules import ttest_analyzer
from modules import spearman_correlation
from modules import cohen_kappa
from modules import chi_square_test
from modules import regression_analysis
from modules import pearson_r_module
from modules import tally_module
from modules import kr20_reliability_app

# Import database module
try:
    from modules import database
    DATABASE_AVAILABLE = True
except ImportError:
    DATABASE_AVAILABLE = False
    print("⚠️ Warning: Database module not found. History features will be disabled.")


# Global theme manager
class ThemeManager:
    """Global theme manager that all modules will use"""
    _instance = None
    _current_mode = "dark"
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(ThemeManager, cls).__new__(cls)
        return cls._instance
    
    @classmethod
    def set_theme(cls, mode):
        """Set theme globally for all windows"""
        cls._current_mode = mode
        ctk.set_appearance_mode(mode)
    
    @classmethod
    def get_theme(cls):
        """Get current theme"""
        return cls._current_mode
    
    @classmethod
    def toggle_theme(cls):
        """Toggle between dark and light mode"""
        if cls._current_mode == "dark":
            cls.set_theme("light")
            return "light"
        else:
            cls.set_theme("dark")
            return "dark"


# Initialize theme
theme_manager = ThemeManager()
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class MainMenuApp(ctk.CTk):
    """Main Menu Application Window with Database Integration"""
    
    def __init__(self):
        super().__init__()
        
        # Configure main window
        self.title("Statistical Analysis Suite (Database Enabled)")
        self.geometry("920x680")
        self.resizable(True, True)
        
        # Set minimum size
        self.minsize(860, 620)
        
        # Center window on screen
        self.center_window()
        
        # Database state
        self.db_path = os.path.join(os.path.dirname(__file__), "stats_app.db")
        self.database_ready = False
        self.db_stats = {}

        # Initialize database
        if DATABASE_AVAILABLE:
            self.init_database()
        
        # Create UI
        self.create_ui()
        
    def center_window(self):
        """Center the window on screen"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')
    
    def init_database(self):
        """Initialize the database system"""
        try:
            database.initialize_database(self.db_path)
            self.db_stats = database.get_database_stats(self.db_path)
            self.database_ready = True
        except Exception as e:
            print(f"Database initialization error: {e}")
            self.database_ready = False
            self.db_stats = {}
        
    def create_ui(self):
        """Create the main menu UI"""
        
        # Main scrollable container (THIS IS KEY!)
        main_scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        main_scroll.pack(fill="both", expand=True, padx=30, pady=30)
        
        # Header with theme toggle
        header_frame = ctk.CTkFrame(main_scroll, fg_color="#1f538d", corner_radius=10)
        header_frame.pack(fill="x", pady=(0, 20))
        
        # Theme toggle button (top right)
        self.theme_btn = ctk.CTkButton(
            header_frame,
            text="☀️ Light Mode",
            command=self.toggle_theme,
            width=130,
            height=32,
            font=ctk.CTkFont(size=12),
            fg_color="#2d5a8a",
            hover_color="#1e3a5f"
        )
        self.theme_btn.pack(side="right", padx=15, pady=10)
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="Statistical Analysis Suite",
            font=ctk.CTkFont(size=32, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=(20, 5))
        
        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Professional Statistical Tools with Database History",
            font=ctk.CTkFont(size=14),
            text_color="#e0e0e0"
        )
        subtitle_label.pack(pady=(0, 20))
        
        # Database status indicator
        if self.database_ready:
            db_status = ctk.CTkFrame(main_scroll, fg_color="#2e7d32", corner_radius=8)
            db_status.pack(fill="x", pady=(0, 15))
            
            status_text = f"📊 Database Active  |  {self.db_stats.get('total_records', 0)} Saved Analyses"
            ctk.CTkLabel(
                db_status,
                text=status_text,
                font=ctk.CTkFont(size=11),
                text_color="white"
            ).pack(pady=8)
        
        # Modules section: single combo launcher
        modules_label = ctk.CTkLabel(
            main_scroll,
            text="Select Analysis Tool:",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        modules_label.pack(anchor="w", pady=(10, 12))

        selector_card = ctk.CTkFrame(main_scroll, corner_radius=12)
        selector_card.pack(fill="x", pady=(0, 14))

        ctk.CTkLabel(
            selector_card,
            text="All modules are now available from one list",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        ).pack(anchor="w", padx=16, pady=(12, 8))

        self.module_actions = {
            "ANOVA — One-Way ANOVA": self.launch_anova,
            "Cronbach's Alpha": self.launch_cronbach,
            "KR-20 Reliability": self.launch_kr20,
            "Independent t-Test": self.launch_ttest,
            "Pearson R Correlation": self.launch_pearson_r,
            "Spearman Correlation": self.launch_spearman,
            "Cohen's Kappa": self.launch_kappa,
            "Chi-Square Test": self.launch_chi_square,
            "Regression Analysis": self.launch_regression,
            "Tally Module": self.launch_tally,
            "Help & Docs": self.show_help,
        }
        module_names = list(self.module_actions.keys())

        self.module_combo = ctk.CTkComboBox(
            selector_card,
            values=module_names,
            width=560,
            height=42,
            font=ctk.CTkFont(size=14),
            dropdown_font=ctk.CTkFont(size=13),
            state="readonly"
        )
        self.module_combo.set("Pearson R Correlation")
        self.module_combo.pack(fill="x", padx=16, pady=(0, 10))

        launch_row = ctk.CTkFrame(selector_card, fg_color="transparent")
        launch_row.pack(fill="x", padx=16, pady=(0, 14))

        ctk.CTkButton(
            launch_row,
            text="▶ Launch Selected Module",
            command=self.launch_selected_module,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#1976d2",
            hover_color="#1565c0"
        ).pack(side="left")

        ctk.CTkButton(
            launch_row,
            text="📚 Help",
            command=self.show_help,
            width=110,
            height=40,
            fg_color="#455a64",
            hover_color="#37474f"
        ).pack(side="left", padx=(10, 0))
        
        # Database management section (if available)
        if self.database_ready:
            self.create_database_section(main_scroll)
        
        # Footer
        footer_frame = ctk.CTkFrame(main_scroll, fg_color="transparent")
        footer_frame.pack(side="bottom", fill="x", pady=(20, 0))
        
        footer_label = ctk.CTkLabel(
            footer_frame,
            text="© 2024 Statistical Analysis Suite | Database Version 2.0",
            font=ctk.CTkFont(size=10),
            text_color="gray"
        )
        footer_label.pack()

        # Exit button
        exit_btn = ctk.CTkButton(
            footer_frame,
            text="Exit Application",
            command=self.quit_app,
            width=150,
            height=35,
            fg_color="#d32f2f",
            hover_color="#b71c1c"
        )
        exit_btn.pack(pady=(10, 0))

    def launch_selected_module(self):
        """Launch currently selected module from combo box."""
        selected = self.module_combo.get().strip()
        action = self.module_actions.get(selected)
        if not action:
            messagebox.showwarning("No Selection", "Please select a module first.")
            return
        action()
    
    def create_database_section(self, parent):
        """Create database management section"""
        db_frame = ctk.CTkFrame(parent)
        db_frame.pack(fill="x", pady=(20, 10))
        
        ctk.CTkLabel(
            db_frame,
            text="📚 Database Management:",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(10, 10))
        
        # Button container
        btn_container = ctk.CTkFrame(db_frame, fg_color="transparent")
        btn_container.pack(fill="x", padx=15, pady=(0, 10))
        
        # View History button
        ctk.CTkButton(
            btn_container,
            text="📜 View Analysis History",
            command=self.view_history,
            width=200,
            height=35,
            font=ctk.CTkFont(size=13),
            fg_color="#1976d2",
            hover_color="#1565c0"
        ).pack(side="left", padx=5)
        
        # Search button
        ctk.CTkButton(
            btn_container,
            text="🔍 Search Analyses",
            command=self.search_analyses,
            width=180,
            height=35,
            font=ctk.CTkFont(size=13),
            fg_color="#7b1fa2",
            hover_color="#6a1b9a"
        ).pack(side="left", padx=5)
        
        # Export button
        ctk.CTkButton(
            btn_container,
            text="💾 Export to CSV",
            command=self.export_history,
            width=160,
            height=35,
            font=ctk.CTkFont(size=13),
            fg_color="#388e3c",
            hover_color="#2e7d32"
        ).pack(side="left", padx=5)
        
    def create_module_card(self, parent, module, row, col):
        """Create a styled card for each analysis module in grid layout"""
        
        # Card frame with hover effects
        card = ctk.CTkFrame(
            parent,
            fg_color="#2b2b2b",
            corner_radius=12,
            cursor="hand2"
        )
        card.grid(row=row, column=col, padx=8, pady=8, sticky="nsew")
        
        # Make card height consistent
        card.grid_propagate(False)
        card.configure(height=180)
        
        # Color accent bar at top
        accent_bar = ctk.CTkFrame(
            card,
            fg_color=module["color"],
            height=4,
            corner_radius=0
        )
        accent_bar.pack(fill="x", side="top")
        
        # Icon (large and centered)
        icon_label = ctk.CTkLabel(
            card,
            text=module["icon"],
            font=ctk.CTkFont(size=48),
            fg_color="transparent"
        )
        icon_label.pack(pady=(20, 5))
        
        # Module name (main title)
        name_label = ctk.CTkLabel(
            card,
            text=module["name"],
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="transparent"
        )
        name_label.pack(pady=(0, 2))
        
        # Full name (smaller, subtle)
        if module["name"] != module["full_name"]:
            full_name_label = ctk.CTkLabel(
                card,
                text=module["full_name"],
                font=ctk.CTkFont(size=10),
                text_color="gray",
                fg_color="transparent"
            )
            full_name_label.pack(pady=(0, 5))
        
        # Description
        desc_label = ctk.CTkLabel(
            card,
            text=module["description"],
            font=ctk.CTkFont(size=10),
            text_color="#b0b0b0",
            fg_color="transparent",
            wraplength=200
        )
        desc_label.pack(pady=(0, 15))
        
        # Hover effects
        def on_enter(event):
            card.configure(fg_color="#3a3a3a")
            accent_bar.configure(height=6)
        
        def on_leave(event):
            card.configure(fg_color="#2b2b2b")
            accent_bar.configure(height=4)
        
        def on_click(event):
            # Visual feedback
            card.configure(fg_color="#4a4a4a")
            self.after(100, lambda: card.configure(fg_color="#3a3a3a"))
            # Execute command
            module["command"]()
        
        # Bind events to all elements
        for widget in [card, accent_bar, icon_label, name_label, desc_label]:
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
            widget.bind("<Button-1>", on_click)
            if hasattr(widget, 'configure'):
                try:
                    widget.configure(cursor="hand2")
                except:
                    pass
        
        # Bind to full name label if it exists
        if module["name"] != module["full_name"]:
            full_name_label.bind("<Enter>", on_enter)
            full_name_label.bind("<Leave>", on_leave)
            full_name_label.bind("<Button-1>", on_click)
            try:
                full_name_label.configure(cursor="hand2")
            except:
                pass
    
    # ========================================================================
    # MODULE LAUNCHERS (with database integration hooks)
    # ========================================================================
    
    def show_help(self):
        """Show help and documentation window"""
        help_window = ctk.CTkToplevel(self)
        help_window.title("Help & Documentation")
        help_window.geometry("700x600")
        
        # Header
        header = ctk.CTkLabel(
            help_window,
            text="📚 Statistical Analysis Suite - Help",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        header.pack(pady=20)
        
        # Help content
        help_text = ctk.CTkTextbox(help_window, font=ctk.CTkFont(size=11), wrap="word")
        help_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        help_content = """
STATISTICAL ANALYSIS SUITE - USER GUIDE
========================================

GETTING STARTED:
1. Select an analysis tool from the module dropdown
2. Import your data or enter manually
3. Configure test parameters
4. Run the analysis
5. Export results as PDF or Word document

AVAILABLE TOOLS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

📊 ONE-WAY ANOVA
   • Compare means across 3+ groups
   • Post-hoc tests included
   • APA-formatted reports

📈 CRONBACH'S ALPHA
   • Assess scale reliability
   • Internal consistency testing
   • JASP-compatible formulas

📉 INDEPENDENT T-TEST
   • Compare two independent groups
   • Welch's correction automatic
   • Effect sizes (Cohen's d)

🔗 SPEARMAN'S CORRELATION
   • Non-parametric correlation
   • Rank-order relationships
   • Visual scatter plots

🤝 COHEN'S KAPPA
   • Inter-rater agreement
   • Categorical data
   • Confidence intervals

✕ CHI-SQUARE TEST
   • Categorical associations
   • McNemar's test option
   • Expected frequencies

📐 REGRESSION ANALYSIS
   • Simple & multiple regression
   • Model fit statistics
   • Coefficient tables

DATABASE FEATURES:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✓ Automatic saving of all analyses
✓ Search and filter history
✓ Export to CSV/JSON
✓ View detailed results

KEYBOARD SHORTCUTS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Ctrl+N : New Analysis
Ctrl+O : Open Data File
Ctrl+S : Save Report
Ctrl+E : Export Results

SUPPORT:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
For technical support or feature requests,
please contact the development team.

Version: 2.0.0 (Database Integrated)
© 2024 Statistical Analysis Suite
        """
        
        help_text.insert("1.0", help_content)
        help_text.configure(state="disabled")
        
        # Close button
        ctk.CTkButton(
            help_window,
            text="Close",
            command=help_window.destroy,
            width=150,
            height=35
        ).pack(pady=(0, 20))
    
    def launch_anova(self):
        """Launch ANOVA Analyzer"""
        try:
            app = anova_analyzer.ANOVAAnalyzer()
            # Pass database save function
            if DATABASE_AVAILABLE:
                app.db_save_function = self.save_analysis_to_db
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch ANOVA Analyzer:\n{str(e)}")
    
    def launch_cronbach(self):
        """Launch Cronbach's Alpha"""
        try:
            app = cronbach_alpha.CronbachAlphaApp()
            if DATABASE_AVAILABLE:
                app.db_save_function = self.save_analysis_to_db
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Cronbach's Alpha:\n{str(e)}")
    
    def launch_ttest(self):
        """Launch Independent t-test"""
        try:
            app = ttest_analyzer.TTestApp()
            if DATABASE_AVAILABLE:
                app.db_save_function = self.save_analysis_to_db
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch t-test Analyzer:\n{str(e)}")
    
    def launch_spearman(self):
        """Launch Spearman's Correlation"""
        try:
            app = spearman_correlation.SpearmanAnalyzer()
            if DATABASE_AVAILABLE:
                app.db_save_function = self.save_analysis_to_db
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Spearman's Correlation:\n{str(e)}")
    
    def launch_kappa(self):
        """Launch Cohen's Kappa"""
        try:
            app = cohen_kappa.KappaApp()
            if DATABASE_AVAILABLE:
                app.db_save_function = self.save_analysis_to_db
            app.run()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Cohen's Kappa:\n{str(e)}")

    def launch_chi_square(self):
        """Launch Chi-Square Test"""
        try:
            app = chi_square_test.ChiSquareTestApp()
            if DATABASE_AVAILABLE:
                app.db_save_function = self.save_analysis_to_db
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Chi-Square Test:\n{str(e)}")

    def launch_regression(self):
        """Launch Regression Analysis"""
        try:
            app = regression_analysis.RegressionApp()
            if DATABASE_AVAILABLE:
                app.db_save_function = self.save_analysis_to_db
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Regression Analysis:\n{str(e)}")

    def launch_pearson_r(self):
        """Launch Pearson R module."""
        try:
            app = pearson_r_module.PearsonRApp()
            if DATABASE_AVAILABLE:
                app.db_save_function = self.save_analysis_to_db
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Pearson R module:\n{str(e)}")

    def launch_tally(self):
        """Launch Tally module."""
        try:
            app = tally_module.TallyApp()
            if DATABASE_AVAILABLE:
                app.db_save_function = self.save_analysis_to_db
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Tally module:\n{str(e)}")

    def launch_kr20(self):
        """Launch KR-20 reliability module."""
        try:
            app = kr20_reliability_app.KR20ReliabilityApp()
            if DATABASE_AVAILABLE:
                app.db_save_function = self.save_analysis_to_db
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch KR-20 module:\n{str(e)}")
    
    # ========================================================================
    # DATABASE OPERATIONS
    # ========================================================================
    
    def save_analysis_to_db(self, analysis_type, input_data, result, interpretation="", metadata=None):
        """
        Save analysis result to database.
        This function is passed to analysis modules for automatic saving.
        """
        if not self.database_ready:
            return None
        try:
            record_id = database.save_result(
                analysis_type=analysis_type,
                input_data=input_data,
                result=result,
                interpretation=interpretation,
                metadata=metadata or {},
                db_path=self.db_path
            )
            
            if record_id:
                # Update stats
                self.db_stats = database.get_database_stats(self.db_path)
                print(f"✓ Analysis saved to database (ID: {record_id})")
                return record_id
            else:
                print("✗ Failed to save analysis to database")
                return None
                
        except Exception as e:
            print(f"✗ Error saving to database: {e}")
            return None
    
    def view_history(self):
        """Open history viewer window"""
        if not self.database_ready:
            messagebox.showwarning("Database", "Database is not available.")
            return
        HistoryViewer(self)
    
    def search_analyses(self):
        """Open search window"""
        if not self.database_ready:
            messagebox.showwarning("Database", "Database is not available.")
            return
        SearchWindow(self)
    
    def export_history(self):
        """Export analysis history to CSV"""
        if not self.database_ready:
            messagebox.showwarning("Database", "Database is not available.")
            return
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=f"analysis_history_{datetime.now().strftime('%Y%m%d')}.csv"
            )
            
            if filename:
                success = database.export_to_csv(filename, db_path=self.db_path)
                if success:
                    messagebox.showinfo("Success", f"History exported to:\n{filename}")
                else:
                    messagebox.showerror("Error", "Failed to export history")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{str(e)}")
    
    def quit_app(self):
        """Exit the application"""
        if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
            self.quit()
            self.destroy()
    
    def toggle_theme(self):
        """Toggle between dark and light mode"""
        new_theme = theme_manager.toggle_theme()
        
        if new_theme == "light":
            self.theme_btn.configure(text="🌙 Dark Mode")
            messagebox.showinfo(
                "Theme Changed", 
                "Light mode activated!\n\nAll opened and future windows will use light mode."
            )
        else:
            self.theme_btn.configure(text="☀️ Light Mode")
            messagebox.showinfo(
                "Theme Changed", 
                "Dark mode activated!\n\nAll opened and future windows will use dark mode."
            )


# ============================================================================
# HISTORY VIEWER WINDOW
# ============================================================================

class HistoryViewer(ctk.CTkToplevel):
    """Window for viewing analysis history"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.parent_app = parent
        
        self.title("Analysis History")
        self.geometry("1000x600")
        
        # Header
        header = ctk.CTkLabel(
            self,
            text="📚 Analysis History",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        header.pack(pady=20)
        
        # Control panel
        control_frame = ctk.CTkFrame(self)
        control_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        ctk.CTkButton(
            control_frame,
            text="🔄 Refresh",
            command=self.load_history,
            width=100
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            control_frame,
            text="🗑️ Delete Selected",
            command=self.delete_selected,
            width=140,
            fg_color="#d32f2f"
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            control_frame,
            text="👁️ View Details",
            command=self.view_details,
            width=120
        ).pack(side="left", padx=5)
        
        # Results display
        self.results_text = ctk.CTkTextbox(self, font=ctk.CTkFont(size=11))
        self.results_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # Load history
        self.load_history()
        self.current_records = []
    
    def load_history(self):
        """Load and display analysis history"""
        try:
            records = database.get_all_results(limit=50, db_path=self.parent_app.db_path)
            self.current_records = records
            
            self.results_text.delete("1.0", "end")
            
            if not records:
                self.results_text.insert("1.0", "No analysis history found.")
                return
            
            output = "=" * 90 + "\n"
            output += f"ANALYSIS HISTORY (Showing {len(records)} most recent)\n"
            output += "=" * 90 + "\n\n"
            
            for record in records:
                output += f"ID: {record['id']} | Type: {record['analysis_type']}\n"
                output += f"Date: {record['date_created']}\n"
                
                if record.get('interpretation'):
                    interp = record['interpretation'][:100]
                    output += f"Summary: {interp}...\n"
                
                output += "-" * 90 + "\n\n"
            
            self.results_text.insert("1.0", output)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load history:\n{str(e)}")
    
    def delete_selected(self):
        """Delete selected record"""
        try:
            # Simple input dialog for ID
            dialog = ctk.CTkInputDialog(
                text="Enter the ID of the record to delete:",
                title="Delete Record"
            )
            id_input = dialog.get_input()
            
            if id_input:
                record_id = int(id_input)
                
                if messagebox.askyesno("Confirm", f"Delete record ID {record_id}?"):
                    success = database.delete_result(record_id, db_path=self.parent_app.db_path)
                    if success:
                        messagebox.showinfo("Success", "Record deleted")
                        self.load_history()
                    else:
                        messagebox.showerror("Error", "Failed to delete record")
        except ValueError:
            messagebox.showerror("Error", "Invalid ID")
        except Exception as e:
            messagebox.showerror("Error", f"Delete failed:\n{str(e)}")
    
    def view_details(self):
        """View detailed information about a record"""
        try:
            dialog = ctk.CTkInputDialog(
                text="Enter the ID of the record to view:",
                title="View Details"
            )
            id_input = dialog.get_input()
            
            if id_input:
                record_id = int(id_input)
                record = database.get_result_by_id(record_id, db_path=self.parent_app.db_path)
                
                if record:
                    DetailViewer(self, record)
                else:
                    messagebox.showerror("Error", f"Record {record_id} not found")
        except ValueError:
            messagebox.showerror("Error", "Invalid ID")
        except Exception as e:
            messagebox.showerror("Error", f"View failed:\n{str(e)}")


# ============================================================================
# SEARCH WINDOW
# ============================================================================

class SearchWindow(ctk.CTkToplevel):
    """Window for searching analyses"""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.parent_app = parent
        
        self.title("Search Analyses")
        self.geometry("700x500")
        
        # Header
        ctk.CTkLabel(
            self,
            text="🔍 Search Analyses",
            font=ctk.CTkFont(size=20, weight="bold")
        ).pack(pady=15)
        
        # Search type
        ctk.CTkLabel(self, text="Analysis Type:").pack(pady=5)
        
        types = database.get_analysis_types(db_path=self.parent_app.db_path)
        self.type_menu = ctk.CTkOptionMenu(self, values=["All"] + types)
        self.type_menu.pack(pady=5)
        
        # Search button
        ctk.CTkButton(
            self,
            text="Search",
            command=self.search,
            width=150,
            height=35
        ).pack(pady=15)
        
        # Results
        self.results_text = ctk.CTkTextbox(self, font=ctk.CTkFont(size=11))
        self.results_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))
    
    def search(self):
        """Perform search"""
        try:
            analysis_type = self.type_menu.get()
            
            if analysis_type == "All":
                results = database.get_all_results(db_path=self.parent_app.db_path)
            else:
                results = database.search_by_analysis_type(
                    analysis_type, db_path=self.parent_app.db_path
                )
            
            self.results_text.delete("1.0", "end")
            
            if not results:
                self.results_text.insert("1.0", "No results found.")
                return
            
            output = f"Found {len(results)} results:\n\n"
            output += "=" * 70 + "\n\n"
            
            for record in results:
                output += f"ID: {record['id']} | {record['analysis_type']}\n"
                output += f"Date: {record['date_created']}\n"
                if record.get('interpretation'):
                    output += f"Summary: {record['interpretation'][:80]}...\n"
                output += "-" * 70 + "\n\n"
            
            self.results_text.insert("1.0", output)
            
        except Exception as e:
            messagebox.showerror("Error", f"Search failed:\n{str(e)}")


# ============================================================================
# DETAIL VIEWER WINDOW
# ============================================================================

class DetailViewer(ctk.CTkToplevel):
    """Window for viewing detailed record information"""
    
    def __init__(self, parent, record):
        super().__init__(parent)
        
        self.title(f"Record Details - ID {record['id']}")
        self.geometry("800x600")
        
        # Display record details
        details_text = ctk.CTkTextbox(self, font=ctk.CTkFont(size=11))
        details_text.pack(fill="both", expand=True, padx=20, pady=20)
        
        output = "=" * 70 + "\n"
        output += f"RECORD DETAILS - ID {record['id']}\n"
        output += "=" * 70 + "\n\n"
        
        output += f"Analysis Type: {record['analysis_type']}\n"
        output += f"Date Created: {record['date_created']}\n\n"
        
        output += "INTERPRETATION:\n"
        output += "-" * 70 + "\n"
        output += f"{record.get('interpretation', 'N/A')}\n\n"
        
        output += "RESULTS:\n"
        output += "-" * 70 + "\n"
        output += f"{record.get('result', 'N/A')}\n\n"
        
        output += "INPUT DATA:\n"
        output += "-" * 70 + "\n"
        output += f"{record.get('input_data', 'N/A')}\n"
        
        details_text.insert("1.0", output)
        details_text.configure(state="disabled")


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """Main entry point"""
    # Create output directory if it doesn't exist
    output_dir = os.path.join(os.path.dirname(__file__), 'output')
    os.makedirs(output_dir, exist_ok=True)
    
    # Create regression reports subdirectory
    regression_output = os.path.join(output_dir, 'regression_reports')
    os.makedirs(regression_output, exist_ok=True)
    
    # Launch main menu
    app = MainMenuApp()
    app.mainloop()


if __name__ == "__main__":
    main()