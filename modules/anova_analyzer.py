import customtkinter as ctk
from tkinter import messagebox, filedialog
from scipy.stats import f_oneway
from statsmodels.stats.multicomp import pairwise_tukeyhsd
import numpy as np
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime
import pandas as pd

# Set appearance mode and default color theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ANOVAAnalyzer(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure main window
        self.title("One-Way ANOVA Analyzer")
        self.geometry("900x750")  # Increased height to accommodate new fields
        
        # Store group input widgets
        self.group_widgets = []
        self.anova_results = None
        
        # Create main layout
        self.create_layout()
        
        # Add initial 3 groups
        for i in range(3):
            self.add_group()
    
    def create_layout(self):
        """Create the main GUI layout"""
        
        # Title
        title_label = ctk.CTkLabel(
            self, 
            text="One-Way ANOVA Analysis", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=15)
        
        # Report Info Frame (Title, Subtitle, Name)
        info_frame = ctk.CTkFrame(self)
        info_frame.pack(fill="x", padx=20, pady=(0, 10))
        
        # Report Title
        title_container = ctk.CTkFrame(info_frame)
        title_container.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(
            title_container, 
            text="Report Title:", 
            font=ctk.CTkFont(size=12, weight="bold"),
            width=100
        ).pack(side="left", padx=(5, 5))
        
        self.report_title_entry = ctk.CTkEntry(
            title_container, 
            placeholder_text="e.g., ANOVA ANALYSIS RESULTS"
        )
        self.report_title_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        
        # Report Subtitle
        subtitle_container = ctk.CTkFrame(info_frame)
        subtitle_container.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(
            subtitle_container, 
            text="Subtitle:", 
            font=ctk.CTkFont(size=12, weight="bold"),
            width=100
        ).pack(side="left", padx=(5, 5))
        
        self.report_subtitle_entry = ctk.CTkEntry(
            subtitle_container, 
            placeholder_text="e.g. Variables A. VS B."
        )
        self.report_subtitle_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        
        # Researcher Name
        name_container = ctk.CTkFrame(info_frame)
        name_container.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkLabel(
            name_container, 
            text="by:", 
            font=ctk.CTkFont(size=12, weight="bold"),
            width=100
        ).pack(side="left", padx=(5, 5))
        
        self.researcher_name_entry = ctk.CTkEntry(
            name_container, 
            placeholder_text="e.g., Dr. John Smith"
        )
        self.researcher_name_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # Main container with two columns
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        # Left column - Input section
        left_frame = ctk.CTkFrame(main_frame)
        left_frame.pack(side="left", fill="both", expand=True, padx=(10, 5), pady=10)
        
        input_label = ctk.CTkLabel(
            left_frame, 
            text="Input Data (comma-separated values)", 
            font=ctk.CTkFont(size=16, weight="bold")
        )
        input_label.pack(pady=(10, 10))
        
        # Scrollable frame for groups
        self.groups_frame = ctk.CTkScrollableFrame(left_frame, height=300)
        self.groups_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Control buttons for groups
        control_frame = ctk.CTkFrame(left_frame)
        control_frame.pack(fill="x", padx=10, pady=10)

        add_group_btn = ctk.CTkButton(
            control_frame, 
            text="➕ Add Group", 
            command=self.add_group,
            width=120
        )
        add_group_btn.pack(side="left", padx=5)

        import_excel_btn = ctk.CTkButton(
            control_frame, 
            text="📁 Import Excel", 
            command=self.import_excel,
            width=120,
            fg_color="#1976d2",
            hover_color="#0d47a1"
        )
        import_excel_btn.pack(side="left", padx=5)

        clear_btn = ctk.CTkButton(
            control_frame, 
            text="🗑️ Clear All", 
            command=self.clear_all,
            width=120,
            fg_color="#d32f2f",
            hover_color="#b71c1c"
        )
        clear_btn.pack(side="left", padx=5)
        
        # Right column - Results and actions
        right_frame = ctk.CTkFrame(main_frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=(5, 10), pady=10)
        
        results_label = ctk.CTkLabel(
            right_frame, 
            text="Analysis Results", 
            font=ctk.CTkFont(size=16, weight="bold")
        )
        results_label.pack(pady=(10, 10))
        
        # Results textbox
        self.results_text = ctk.CTkTextbox(right_frame, height=300, wrap="word")
        self.results_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Action buttons
        action_frame = ctk.CTkFrame(right_frame)
        action_frame.pack(fill="x", padx=10, pady=10)
        
        run_btn = ctk.CTkButton(
            action_frame, 
            text="▶️ Run ANOVA", 
            command=self.run_anova,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2e7d32",
            hover_color="#1b5e20"
        )
        run_btn.pack(fill="x", pady=(0, 10))
        
        save_btn = ctk.CTkButton(
            action_frame, 
            text="💾 Save as DOCX", 
            command=self.save_to_docx,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        save_btn.pack(fill="x", pady=(0, 10))
        
        reset_btn = ctk.CTkButton(
            action_frame, 
            text="🔄 Reset", 
            command=self.reset_all,
            height=40,
            fg_color="#f57c00",
            hover_color="#e65100"
        )
        reset_btn.pack(fill="x")
        
        # File location label at the bottom
        self.file_location_label = ctk.CTkLabel(
            self,
            text="No file saved yet",
            font=ctk.CTkFont(size=9, slant="italic"),
            text_color="gray"
        )
        self.file_location_label.pack(side="bottom", pady=5)
        
    def import_excel(self):
        """Import group data from an Excel file"""
        # Open file dialog
        filepath = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not filepath:
            return
        
        try:
            # Read Excel file using pandas
            df = pd.read_excel(filepath)
            
            # Get number of columns with data
            num_columns = len(df.columns)
            
            if num_columns == 0:
                messagebox.showerror("Error", "The Excel file has no data columns!")
                return
            
            # Clear existing entries first
            self.clear_all()
            
            # Determine how many groups we need
            current_groups = len(self.group_widgets)
            
            # Add more groups if Excel has more columns
            if num_columns > current_groups:
                for _ in range(num_columns - current_groups):
                    self.add_group()
            
            # Process each column
            successfully_imported = 0
            
            for col_idx, column_name in enumerate(df.columns):
                if col_idx >= len(self.group_widgets):
                    break
                
                # Get column data and remove NaN values
                column_data = df[column_name].dropna()
                
                # Validate that all values are numeric
                if not pd.api.types.is_numeric_dtype(column_data):
                    # Try to convert to numeric
                    try:
                        column_data = pd.to_numeric(column_data, errors='coerce').dropna()
                    except:
                        messagebox.showwarning(
                            "Warning", 
                            f"Column '{column_name}' contains non-numeric values and will be skipped."
                        )
                        continue
                
                # Check if column has data after cleaning
                if len(column_data) == 0:
                    messagebox.showwarning(
                        "Warning", 
                        f"Column '{column_name}' has no valid numeric data and will be skipped."
                    )
                    continue
                
                # Convert to comma-separated string
                values_str = ', '.join([str(val) for val in column_data.values])
                
                # Get the entry widget for this group
                _, _, entry, _ = self.group_widgets[col_idx]
                
                # Clear and insert new values
                entry.delete(0, 'end')
                entry.insert(0, values_str)
                
                successfully_imported += 1
            
            # Show success message
            if successfully_imported > 0:
                messagebox.showinfo(
                    "Success", 
                    f"Successfully imported {successfully_imported} group(s) from Excel!\n\n"
                    f"File: {filepath.split('/')[-1]}"
                )
            else:
                messagebox.showwarning(
                    "Warning", 
                    "No valid numeric data was found in the Excel file."
                )
        
        except FileNotFoundError:
            messagebox.showerror("Error", "Excel file not found!")
        
        except Exception as e:
            messagebox.showerror(
                "Error", 
                f"Failed to import Excel file:\n{str(e)}\n\n"
                f"Make sure the file is a valid Excel (.xlsx) file."
            )
        
    def add_group(self):
        """Add a new group input field"""
        group_num = len(self.group_widgets) + 1
        
        # Frame for single group
        group_frame = ctk.CTkFrame(self.groups_frame)
        group_frame.pack(fill="x", pady=5)
        
        # Label
        label = ctk.CTkLabel(
            group_frame, 
            text=f"Group {group_num}:", 
            font=ctk.CTkFont(size=13, weight="bold"),
            width=80
        )
        label.pack(side="left", padx=(5, 5))
        
        # Entry field
        entry = ctk.CTkEntry(group_frame, placeholder_text="e.g., 12, 15, 14, 17")
        entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # Remove button
        remove_btn = ctk.CTkButton(
            group_frame, 
            text="✖", 
            width=30,
            command=lambda: self.remove_group(group_frame, label, entry, remove_btn),
            fg_color="#d32f2f",
            hover_color="#b71c1c"
        )
        remove_btn.pack(side="right", padx=5)
        
        # Store references
        self.group_widgets.append((group_frame, label, entry, remove_btn))
    
    def remove_group(self, group_frame, label, entry, remove_btn):
        """Remove a group input field"""
        if len(self.group_widgets) <= 2:
            messagebox.showwarning("Warning", "You must have at least 2 groups for ANOVA!")
            return
        
        # Remove from GUI
        group_frame.destroy()
        
        # Remove from list
        self.group_widgets = [
            widget for widget in self.group_widgets 
            if widget[0] != group_frame
        ]
        
        # Renumber remaining groups
        for i, (frame, lbl, ent, btn) in enumerate(self.group_widgets, 1):
            lbl.configure(text=f"Group {i}:")
    
    def clear_all(self):
        """Clear all input fields"""
        for _, _, entry, _ in self.group_widgets:
            entry.delete(0, 'end')
    
    def reset_all(self):
        """Reset everything including results"""
        self.clear_all()
        self.results_text.delete("1.0", "end")
        self.anova_results = None
        # Don't clear report info fields on reset
    
    def validate_and_parse_inputs(self):
        """Validate and parse all group inputs"""
        if len(self.group_widgets) < 2:
            messagebox.showerror("Error", "You must have at least 2 groups!")
            return None
        
        groups = []
        group_names = []
        
        for i, (_, _, entry, _) in enumerate(self.group_widgets, 1):
            text = entry.get().strip()
            
            if not text:
                messagebox.showerror("Error", f"Group {i} is empty! Please enter values.")
                return None
            
            try:
                # Parse comma-separated values
                values = [float(x.strip()) for x in text.split(',')]
                
                if len(values) < 2:
                    messagebox.showerror("Error", f"Group {i} must have at least 2 values!")
                    return None
                
                groups.append(values)
                group_names.append(f"Group {i}")
                
            except ValueError:
                messagebox.showerror("Error", f"Group {i} contains invalid values! Use only numbers.")
                return None
        
        return groups, group_names
    
    def run_anova(self):
        """Perform One-Way ANOVA analysis"""
        # Validate and parse inputs
        result = self.validate_and_parse_inputs()
        if result is None:
            return
        
        groups, group_names = result
        
        try:
            # Perform ANOVA using scipy
            F_statistic, p_value = f_oneway(*groups)
            
            # Manual ANOVA calculations for detailed output
            all_data = np.concatenate(groups)
            grand_mean = np.mean(all_data)
            k = len(groups)
            N = len(all_data)
            
            # Sum of Squares
            SS_between = sum(len(g) * (np.mean(g) - grand_mean)**2 for g in groups)
            SS_total = np.sum((all_data - grand_mean)**2)
            SS_within = SS_total - SS_between
            
            # Degrees of Freedom
            df_between = k - 1
            df_within = N - k
            
            # Mean Squares
            MS_between = SS_between / df_between
            MS_within = SS_within / df_within
            
            # Alpha level
            alpha = 0.05
            
            # Decision
            decision = "Reject H₀" if p_value < alpha else "Fail to Reject H₀"
            is_significant = p_value < alpha
            
            # Conclusion
            if is_significant:
                conclusion = "There is a statistically significant difference among the group means."
            else:
                conclusion = "There is no statistically significant difference among the group means."
            
            # Store results for DOCX export
            self.anova_results = {
                'groups': groups,
                'group_names': group_names,
                'F_statistic': F_statistic,
                'p_value': p_value,
                'alpha': alpha,
                'decision': decision,
                'conclusion': conclusion,
                'is_significant': is_significant,
                'SS_between': SS_between,
                'SS_within': SS_within,
                'SS_total': SS_total,
                'df_between': df_between,
                'df_within': df_within,
                'MS_between': MS_between,
                'MS_within': MS_within,
                'all_data': all_data,
                'report_title': self.report_title_entry.get().strip() or "ANOVA ANALYSIS RESULTS",
                'report_subtitle': self.report_subtitle_entry.get().strip(),
                'researcher_name': self.researcher_name_entry.get().strip()
            }
            
            # Display results
            self.display_results()
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during analysis:\n{str(e)}")
    
    def display_results(self):
        """Display ANOVA results in the textbox"""
        self.results_text.delete("1.0", "end")
        
        if self.anova_results is None:
            return
        
        r = self.anova_results
        
        # Build output text
        output = "=" * 50 + "\n"
        output += f"{r['report_title']}\n"
        if r['report_subtitle']:
            output += f"{r['report_subtitle']}\n"
        if r['researcher_name']:
            output += f"by: {r['researcher_name']}\n"
        output += "=" * 50 + "\n\n"
        
        # Descriptive Statistics
        output += "DESCRIPTIVE STATISTICS\n"
        output += "-" * 50 + "\n"
        for i, (name, group) in enumerate(zip(r['group_names'], r['groups'])):
            output += f"{name}:\n"
            output += f"  n = {len(group)}\n"
            output += f"  Mean = {np.mean(group):.2f}\n"
            output += f"  SD = {np.std(group, ddof=1):.4f}\n\n"
        
        # ANOVA Table
        output += "\nANOVA TABLE\n"
        output += "-" * 50 + "\n"
        output += f"{'Source':<15} {'SS':<12} {'df':<6} {'MS':<12} {'F':<10}\n"
        output += f"{'Between':<15} {r['SS_between']:<12.4f} {r['df_between']:<6} {r['MS_between']:<12.2f} {r['F_statistic']:<10.3f}\n"
        output += f"{'Within':<15} {r['SS_within']:<12.4f} {r['df_within']:<6} {r['MS_within']:<12.4f}\n"
        output += f"{'Total':<15} {r['SS_total']:<12.4f} {r['df_between'] + r['df_within']:<6}\n"
        
        # Test Results
        output += "\n" + "=" * 50 + "\n"
        output += "TEST RESULTS\n"
        output += "=" * 50 + "\n"
        output += f"F-statistic: {r['F_statistic']:.4f}\n"
        output += f"p-value: {r['p_value']:.6f}\n"
        output += f"Alpha level: {r['alpha']}\n"
        output += f"Degrees of Freedom: ({r['df_between']}, {r['df_within']})\n\n"
        
        output += f"Decision: {r['decision']}\n\n"
        
        output += "CONCLUSION:\n"
        output += f"{r['conclusion']}\n"
        
        # Post-hoc if significant
        if r['is_significant']:
            output += "\n\n" + "=" * 50 + "\n"
            output += "POST HOC ANALYSIS (Tukey HSD)\n"
            output += "=" * 50 + "\n"
            
            try:
                # Prepare data for Tukey
                labels = []
                for i, group in enumerate(r['groups']):
                    labels.extend([r['group_names'][i]] * len(group))
                
                tukey = pairwise_tukeyhsd(r['all_data'], labels)
                output += "\n" + str(tukey) + "\n"
                
                # Store tukey results
                self.anova_results['tukey'] = tukey
            except Exception as e:
                output += f"\nCould not perform post-hoc analysis: {str(e)}\n"
        
        self.results_text.insert("1.0", output)
    
    def save_to_docx(self):
        """Save ANOVA results to a DOCX file"""
        if self.anova_results is None:
            messagebox.showwarning("Warning", "No results to save! Run ANOVA first.")
            return
        
        # Ask user for save location
        filepath = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")],
            title="Save ANOVA Report"
        )
        
        if not filepath:
            return
        
        try:
            r = self.anova_results
            
            # Create document
            doc = Document()
            
            # Title
            title = doc.add_heading(r['report_title'], 1)
            title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            for run in title.runs:
                run.font.size = Pt(16)
            
            # Subtitle (if provided)
            if r['report_subtitle']:
                subtitle = doc.add_heading(r['report_subtitle'], level=2)
                subtitle.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                for run in subtitle.runs:
                    run.font.size = Pt(12)
            
            # Researcher name (if provided)
            if r['researcher_name']:
                name_para = doc.add_paragraph(f"By: {r['researcher_name']}")
                name_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                for run in name_para.runs:
                    run.font.size = Pt(10)
            
            doc.add_paragraph()
            
            # Descriptive Statistics
            doc.add_heading('Descriptive Statistics', level=2)
            for run in doc.paragraphs[-1].runs:
                run.font.size = Pt(11)
            
            desc_table = doc.add_table(rows=len(r['groups']) + 1, cols=4)
            desc_table.style = 'Table Grid'
            
            headers = ['Group', 'n', 'Mean', 'Std. Dev']
            for i, header in enumerate(headers):
                cell = desc_table.rows[0].cells[i]
                cell.text = header
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
            
            for i, (name, group) in enumerate(zip(r['group_names'], r['groups']), 1):
                desc_table.rows[i].cells[0].text = name
                desc_table.rows[i].cells[1].text = str(len(group))
                desc_table.rows[i].cells[2].text = f"{np.mean(group):.4f}"
                desc_table.rows[i].cells[3].text = f"{np.std(group, ddof=1):.4f}"
            
            # Set font size for table
            for row in desc_table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(9)
            
            doc.add_paragraph()
            
            # ANOVA Table
            doc.add_heading('ANOVA Table', level=2)
            for run in doc.paragraphs[-1].runs:
                run.font.size = Pt(11)
            
            anova_table = doc.add_table(rows=4, cols=6)
            anova_table.style = 'Table Grid'
            
            headers = ['Source', 'SS', 'df', 'MS', 'F', 'p-value']
            for i, header in enumerate(headers):
                cell = anova_table.rows[0].cells[i]
                cell.text = header
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
            
            # Between Groups
            anova_table.rows[1].cells[0].text = 'Between Groups'
            anova_table.rows[1].cells[1].text = f"{r['SS_between']:.4f}"
            anova_table.rows[1].cells[2].text = str(r['df_between'])
            anova_table.rows[1].cells[3].text = f"{r['MS_between']:.4f}"
            anova_table.rows[1].cells[4].text = f"{r['F_statistic']:.4f}"
            anova_table.rows[1].cells[5].text = f"{r['p_value']:.6f}"
            
            # Within Groups
            anova_table.rows[2].cells[0].text = 'Within Groups'
            anova_table.rows[2].cells[1].text = f"{r['SS_within']:.4f}"
            anova_table.rows[2].cells[2].text = str(r['df_within'])
            anova_table.rows[2].cells[3].text = f"{r['MS_within']:.4f}"
            
            # Total
            anova_table.rows[3].cells[0].text = 'Total'
            anova_table.rows[3].cells[1].text = f"{r['SS_total']:.4f}"
            anova_table.rows[3].cells[2].text = str(r['df_between'] + r['df_within'])
            
            # Set font size for ANOVA table
            for row in anova_table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(9)
            
            doc.add_paragraph()
            
            # Test Results
            doc.add_heading('Test Results', level=2)
            for run in doc.paragraphs[-1].runs:
                run.font.size = Pt(11)
            
            results_para = doc.add_paragraph()
            results_para.add_run('F-statistic: ').bold = True
            results_para.add_run(f"{r['F_statistic']:.4f}\n")
            results_para.add_run('p-value: ').bold = True
            results_para.add_run(f"{r['p_value']:.6f}\n")
            results_para.add_run('Alpha level: ').bold = True
            results_para.add_run(f"{r['alpha']}\n")
            results_para.add_run('Degrees of Freedom: ').bold = True
            results_para.add_run(f"({r['df_between']}, {r['df_within']})\n\n")
            results_para.add_run('Decision: ').bold = True
            results_para.add_run(f"{r['decision']}\n")
            
            # Set font size for results
            for run in results_para.runs:
                run.font.size = Pt(10)
            
            # Conclusion
            doc.add_heading('Conclusion', level=2)
            for run in doc.paragraphs[-1].runs:
                run.font.size = Pt(11)
            
            conclusion_para = doc.add_paragraph()
            conclusion_para.add_run(r['conclusion']).bold = True
            for run in conclusion_para.runs:
                run.font.size = Pt(10)
            
            # Post-hoc if significant
            if r['is_significant'] and 'tukey' in r:
                doc.add_paragraph()
                doc.add_heading('Post Hoc Analysis (Tukey HSD)', level=2)
                for run in doc.paragraphs[-1].runs:
                    run.font.size = Pt(11)
                
                tukey = r['tukey']
                tukey_data = tukey.summary().data
                
                posthoc_table = doc.add_table(rows=len(tukey_data), cols=len(tukey_data[0]))
                posthoc_table.style = 'Table Grid'
                
                for i, row in enumerate(tukey_data):
                    for j, val in enumerate(row):
                        cell = posthoc_table.rows[i].cells[j]
                        cell.text = str(val)
                        if i == 0:
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    run.bold = True
                
                # Set font size for post-hoc table
                for row in posthoc_table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = Pt(8)
            
            # Input Data Table
            doc.add_heading('Raw Data', level=2)
            for run in doc.paragraphs[-1].runs:
                run.font.size = Pt(11)
            
            input_table = doc.add_table(rows=len(r['groups']) + 1, cols=3)
            input_table.style = 'Table Grid'
            
            # Headers
            headers = ['Group', 'n', 'Values']
            for i, header in enumerate(headers):
                cell = input_table.rows[0].cells[i]
                cell.text = header
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
            
            # Data rows
            for i, (name, group) in enumerate(zip(r['group_names'], r['groups']), 1):
                input_table.rows[i].cells[0].text = name
                input_table.rows[i].cells[1].text = str(len(group))
                input_table.rows[i].cells[2].text = ', '.join([f"{v:.2f}" for v in group])
            
            # Set font size for input data table
            for row in input_table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(8)
            
            doc.add_paragraph()
            # Save document
            doc.save(filepath)
            
            # Add file location and date at the bottom
            last_para = doc.add_paragraph()
            run = last_para.add_run(f"File Location: {filepath}     Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(128, 128, 128)  # Grey color
            run.italic = True
            
            # Save document again with the footer
            doc.save(filepath)
            
            # Update file location label
            self.file_location_label.configure(text=f"Last saved: {filepath}")
            
            messagebox.showinfo("Success", f"Report saved successfully!\n\n{filepath}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save document:\n{str(e)}")


def main():
    """Main function to run the application"""
    app = ANOVAAnalyzer()
    app.mainloop()


if __name__ == "__main__":
    main()