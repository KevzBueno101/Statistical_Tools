"""
Unified Statistical Analysis Desktop Application
Main Menu - Entry Point

This application integrates multiple statistical analysis tools:
- One-Way ANOVA Analyzer
- Cronbach's Alpha Reliability Test
- Independent Samples t-test
- Spearman's Rank Correlation
- Cohen's Kappa Calculator

Author: Statistical Analysis Suite
Version: 1.0.0
"""

import customtkinter as ctk
from tkinter import messagebox
import sys
import os

# Add modules directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'modules'))

# Import all modules
from modules import anova_analyzer
from modules import cronbach_alpha
from modules import ttest_analyzer
from modules import spearman_correlation
from modules import cohen_kappa

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
    """Main Menu Application Window"""
    
    def __init__(self):
        super().__init__()
        
        # Configure main window
        self.title("Statistical Analysis Suite")
        self.geometry("800x650")
        self.resizable(True, True)  # Make window resizable
        
        # Set minimum size
        self.minsize(700, 600)
        
        # Center window on screen
        self.center_window()
        
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
        
    def create_ui(self):
        """Create the main menu UI"""
        
        # Main container
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=30, pady=30)
        
        # Header with theme toggle
        header_frame = ctk.CTkFrame(main_frame, fg_color="#1f538d", corner_radius=10)
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
            text="Professional Statistical Tools for Research & Analysis",
            font=ctk.CTkFont(size=14),
            text_color="#e0e0e0"
        )
        subtitle_label.pack(pady=(0, 20))
        
        # Modules section
        modules_label = ctk.CTkLabel(
            main_frame,
            text="Select Analysis Tool:",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        modules_label.pack(anchor="w", pady=(10, 15))
        
        # Create module buttons
        modules = [
            {
                "name": "One-Way ANOVA Analyzer",
                "description": "Compare means across multiple groups",
                "icon": "📊",
                "command": self.launch_anova,
                "color": "#2e7d32"
            },
            {
                "name": "Cronbach's Alpha Test",
                "description": "Assess internal consistency reliability",
                "icon": "📈",
                "command": self.launch_cronbach,
                "color": "#1976d2"
            },
            {
                "name": "Independent t-test",
                "description": "Compare means between two groups",
                "icon": "📉",
                "command": self.launch_ttest,
                "color": "#d32f2f"
            },
            {
                "name": "Spearman's Correlation",
                "description": "Analyze rank-order relationships",
                "icon": "🔗",
                "command": self.launch_spearman,
                "color": "#7b1fa2"
            },
            {
                "name": "Cohen's Kappa Calculator",
                "description": "Measure inter-rater agreement",
                "icon": "🤝",
                "command": self.launch_kappa,
                "color": "#f57c00"
            }
        ]
        
        for module in modules:
            self.create_module_button(main_frame, module)
        
        # Footer
        footer_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        footer_frame.pack(side="bottom", fill="x", pady=(20, 0))
        
        footer_label = ctk.CTkLabel(
            footer_frame,
            text="© 2024 Statistical Analysis Suite | All Rights Reserved",
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
        
    def create_module_button(self, parent, module):
        """Create a styled module button - fully clickable"""
        
        # Outer container frame
        container = ctk.CTkFrame(parent, fg_color="transparent")
        container.pack(fill="x", pady=5)
        
        # Main clickable frame (acts as button)
        button_frame = ctk.CTkFrame(
            container, 
            fg_color="#2b2b2b", 
            corner_radius=8,
            height=70,
            cursor="hand2"
        )
        button_frame.pack(fill="x", padx=2, pady=2)
        button_frame.pack_propagate(False)  # Maintain fixed height
        
        # Bind click events to frame and all its children
        def on_click(event):
            module["command"]()
        
        def on_enter(event):
            button_frame.configure(fg_color="#3a3a3a")
        
        def on_leave(event):
            button_frame.configure(fg_color="#2b2b2b")
        
        button_frame.bind("<Button-1>", on_click)
        button_frame.bind("<Enter>", on_enter)
        button_frame.bind("<Leave>", on_leave)
        
        # Icon - clickable
        icon_label = ctk.CTkLabel(
            button_frame,
            text=module["icon"],
            font=ctk.CTkFont(size=32),
            fg_color="transparent",
            cursor="hand2"
        )
        icon_label.place(relx=0.05, rely=0.5, anchor="w")
        icon_label.bind("<Button-1>", on_click)
        icon_label.bind("<Enter>", on_enter)
        icon_label.bind("<Leave>", on_leave)
        
        # Name label - clickable
        name_label = ctk.CTkLabel(
            button_frame,
            text=module["name"],
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w",
            fg_color="transparent",
            cursor="hand2"
        )
        name_label.place(relx=0.15, rely=0.35, anchor="w")
        name_label.bind("<Button-1>", on_click)
        name_label.bind("<Enter>", on_enter)
        name_label.bind("<Leave>", on_leave)
        
        # Description label - clickable
        desc_label = ctk.CTkLabel(
            button_frame,
            text=module["description"],
            font=ctk.CTkFont(size=11),
            text_color="gray",
            anchor="w",
            fg_color="transparent",
            cursor="hand2"
        )
        desc_label.place(relx=0.15, rely=0.65, anchor="w")
        desc_label.bind("<Button-1>", on_click)
        desc_label.bind("<Enter>", on_enter)
        desc_label.bind("<Leave>", on_leave)
        
        # Arrow indicator - clickable
        indicator = ctk.CTkLabel(
            button_frame,
            text="▶",
            font=ctk.CTkFont(size=24),
            text_color=module["color"],
            fg_color="transparent",
            cursor="hand2"
        )
        indicator.place(relx=0.95, rely=0.5, anchor="e")
        indicator.bind("<Button-1>", on_click)
        indicator.bind("<Enter>", on_enter)
        indicator.bind("<Leave>", on_leave)
        
    # Module launchers
    def launch_anova(self):
        """Launch ANOVA Analyzer"""
        try:
            app = anova_analyzer.ANOVAAnalyzer()
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch ANOVA Analyzer:\n{str(e)}")
    
    def launch_cronbach(self):
        """Launch Cronbach's Alpha"""
        try:
            app = cronbach_alpha.CronbachAlphaApp()
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Cronbach's Alpha:\n{str(e)}")
    
    def launch_ttest(self):
        """Launch Independent t-test"""
        try:
            root = ctk.CTk()
            app = ttest_analyzer.IndependentTTestApp(root)
            root.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch t-test Analyzer:\n{str(e)}")
    
    def launch_spearman(self):
        """Launch Spearman's Correlation"""
        try:
            app = spearman_correlation.SpearmanAnalyzer()
            app.mainloop()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Spearman's Correlation:\n{str(e)}")
    
    def launch_kappa(self):
        """Launch Cohen's Kappa"""
        try:
            app = cohen_kappa.KappaApp()
            app.run()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to launch Cohen's Kappa:\n{str(e)}")
    
    def quit_app(self):
        """Exit the application"""
        if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
            self.quit()
            self.destroy()
    
    def toggle_theme(self):
        """Toggle between dark and light mode - applies to ALL modules"""
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


def main():
    """Main entry point"""
    # Create output directory if it doesn't exist
    output_dir = os.path.join(os.path.dirname(__file__), 'output')
    os.makedirs(output_dir, exist_ok=True)
    
    # Launch main menu
    app = MainMenuApp()
    app.mainloop()


if __name__ == "__main__":
    main()