import customtkinter as ctk
from tkinter import messagebox, filedialog
from scipy.stats import f_oneway
from statsmodels.stats.multicomp import pairwise_tukeyhsd
import numpy as np
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
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
        self.geometry("900x750")
        
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
        filepath = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not filepath:
            return
        
        try:
            df = pd.read_excel(filepath)
            num_columns = len(df.columns)
            
            if num_columns == 0:
                messagebox.showerror("Error", "The Excel file has no data columns!")
                return
            
            self.clear_all()
            current_groups = len(self.group_widgets)
            
            if num_columns > current_groups:
                for _ in range(num_columns - current_groups):
                    self.add_group()
            
            successfully_imported = 0
            
            for col_idx, column_name in enumerate(df.columns):
                if col_idx >= len(self.group_widgets):
                    break
                
                column_data = df[column_name].dropna()
                
                if not pd.api.types.is_numeric_dtype(column_data):
                    try:
                        column_data = pd.to_numeric(column_data, errors='coerce').dropna()
                    except:
                        messagebox.showwarning(
                            "Warning", 
                            f"Column '{column_name}' contains non-numeric values and will be skipped."
                        )
                        continue
                
                if len(column_data) == 0:
                    messagebox.showwarning(
                        "Warning", 
                        f"Column '{column_name}' has no valid numeric data and will be skipped."
                    )
                    continue
                
                values_str = ', '.join([str(val) for val in column_data.values])
                _, _, entry, _ = self.group_widgets[col_idx]
                entry.delete(0, 'end')
                entry.insert(0, values_str)
                successfully_imported += 1
            
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
        
        except Exception as e:
            messagebox.showerror(
                "Error", 
                f"Failed to import Excel file:\n{str(e)}\n\n"
                f"Make sure the file is a valid Excel (.xlsx) file."
            )
        
    def add_group(self):
        """Add a new group input field"""
        group_num = len(self.group_widgets) + 1
        
        group_frame = ctk.CTkFrame(self.groups_frame)
        group_frame.pack(fill="x", pady=5)
        
        label = ctk.CTkLabel(
            group_frame, 
            text=f"Group {group_num}:", 
            font=ctk.CTkFont(size=13, weight="bold"),
            width=80
        )
        label.pack(side="left", padx=(5, 5))
        
        entry = ctk.CTkEntry(group_frame, placeholder_text="e.g., 12, 15, 14, 17")
        entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        remove_btn = ctk.CTkButton(
            group_frame, 
            text="✖", 
            width=30,
            command=lambda: self.remove_group(group_frame, label, entry, remove_btn),
            fg_color="#d32f2f",
            hover_color="#b71c1c"
        )
        remove_btn.pack(side="right", padx=5)
        
        self.group_widgets.append((group_frame, label, entry, remove_btn))
    
    def remove_group(self, group_frame, label, entry, remove_btn):
        """Remove a group input field"""
        if len(self.group_widgets) <= 2:
            messagebox.showwarning("Warning", "You must have at least 2 groups for ANOVA!")
            return
        
        group_frame.destroy()
        
        self.group_widgets = [
            widget for widget in self.group_widgets 
            if widget[0] != group_frame
        ]
        
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
        result = self.validate_and_parse_inputs()
        if result is None:
            return
        
        groups, group_names = result
        
        try:
            F_statistic, p_value = f_oneway(*groups)
            
            all_data = np.concatenate(groups)
            grand_mean = np.mean(all_data)
            k = len(groups)
            N = len(all_data)
            
            SS_between = sum(len(g) * (np.mean(g) - grand_mean)**2 for g in groups)
            SS_total = np.sum((all_data - grand_mean)**2)
            SS_within = SS_total - SS_between
            
            df_between = k - 1
            df_within = N - k
            
            MS_between = SS_between / df_between
            MS_within = SS_within / df_within
            
            eta_squared = SS_between / SS_total
            
            alpha = 0.05
            decision = "Reject H₀" if p_value < alpha else "Fail to Reject H₀"
            is_significant = p_value < alpha
            
            if is_significant:
                conclusion = "There is a statistically significant difference among the group means."
            else:
                conclusion = "There is no statistically significant difference among the group means."
            
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
                'eta_squared': eta_squared,
                'all_data': all_data,
                'report_title': self.report_title_entry.get().strip() or "ANOVA ANALYSIS RESULTS",
                'report_subtitle': self.report_subtitle_entry.get().strip(),
                'researcher_name': self.researcher_name_entry.get().strip()
            }
            
            self.display_results()
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during analysis:\n{str(e)}")
    
    def display_results(self):
        """Display ANOVA results in the textbox"""
        self.results_text.delete("1.0", "end")
        
        if self.anova_results is None:
            return
        
        r = self.anova_results
        
        output = "=" * 50 + "\n"
        output += f"{r['report_title']}\n"
        if r['report_subtitle']:
            output += f"{r['report_subtitle']}\n"
        if r['researcher_name']:
            output += f"by: {r['researcher_name']}\n"
        output += "=" * 50 + "\n\n"
        
        output += "DESCRIPTIVE STATISTICS\n"
        output += "-" * 50 + "\n"
        for i, (name, group) in enumerate(zip(r['group_names'], r['groups'])):
            output += f"{name}:\n"
            output += f"  n = {len(group)}\n"
            output += f"  Mean = {np.mean(group):.2f}\n"
            output += f"  SD = {np.std(group, ddof=1):.4f}\n\n"
        
        output += "\nANOVA TABLE\n"
        output += "-" * 50 + "\n"
        output += f"{'Source':<15} {'SS':<12} {'df':<6} {'MS':<12} {'F':<10}\n"
        output += f"{'Between':<15} {r['SS_between']:<12.4f} {r['df_between']:<6} {r['MS_between']:<12.2f} {r['F_statistic']:<10.3f}\n"
        output += f"{'Within':<15} {r['SS_within']:<12.4f} {r['df_within']:<6} {r['MS_within']:<12.4f}\n"
        output += f"{'Total':<15} {r['SS_total']:<12.4f} {r['df_between'] + r['df_within']:<6}\n"
        
        output += "\n" + "=" * 50 + "\n"
        output += "TEST RESULTS\n"
        output += "=" * 50 + "\n"
        output += f"F-statistic: {r['F_statistic']:.4f}\n"
        output += f"p-value: {r['p_value']:.6f}\n"
        output += f"Eta Squared (η²): {r['eta_squared']:.4f}\n"
        output += f"Alpha level: {r['alpha']}\n"
        output += f"Degrees of Freedom: ({r['df_between']}, {r['df_within']})\n\n"
        
        output += f"Decision: {r['decision']}\n\n"
        
        output += "CONCLUSION:\n"
        output += f"{r['conclusion']}\n"
        
        if r['is_significant']:
            output += "\n\n" + "=" * 50 + "\n"
            output += "POST HOC ANALYSIS (Tukey HSD)\n"
            output += "=" * 50 + "\n"
            
            try:
                labels = []
                for i, group in enumerate(r['groups']):
                    labels.extend([r['group_names'][i]] * len(group))
                
                tukey = pairwise_tukeyhsd(r['all_data'], labels)
                output += "\n" + str(tukey) + "\n"
                
                self.anova_results['tukey'] = tukey
            except Exception as e:
                output += f"\nCould not perform post-hoc analysis: {str(e)}\n"
        
        self.results_text.insert("1.0", output)
    
    def set_cell_border(self, cell, **kwargs):
        """Set cell borders for APA table formatting"""
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        
        tcBorders = OxmlElement('w:tcBorders')
        for edge in ('top', 'left', 'bottom', 'right'):
            if edge in kwargs:
                border = OxmlElement(f'w:{edge}')
                for key, value in kwargs[edge].items():
                    border.set(qn(f'w:{key}'), str(value))
                tcBorders.append(border)
        
        tcPr.append(tcBorders)
    
    def create_apa_table(self, doc, data, headers, caption):
        """Create an APA 7th Edition formatted table"""
        caption_para = doc.add_paragraph()
        caption_run = caption_para.add_run(caption)
        caption_run.font.name = 'Times New Roman'
        caption_run.font.size = Pt(8)
        caption_run.bold = True
        caption_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        caption_para.space_after = Pt(2)
        
        table = doc.add_table(rows=len(data) + 1, cols=len(headers))
        table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        for i, header in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = header
            
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(7)
                    run.bold = True
            
            self.set_cell_border(
                cell,
                top={"sz": 12, "val": "single", "color": "000000"},
                bottom={"sz": 8, "val": "single", "color": "000000"},
                left={"sz": 0, "val": "none"},
                right={"sz": 0, "val": "none"}
            )
        
        for row_idx, row_data in enumerate(data, start=1):
            for col_idx, value in enumerate(row_data):
                cell = table.rows[row_idx].cells[col_idx]
                cell.text = str(value)
                
                for paragraph in cell.paragraphs:
                    if col_idx == 0:
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    else:
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(9)
                
                if row_idx == len(data):
                    self.set_cell_border(
                        cell,
                        top={"sz": 0, "val": "none"},
                        bottom={"sz": 12, "val": "single", "color": "000000"},
                        left={"sz": 0, "val": "none"},
                        right={"sz": 0, "val": "none"}
                    )
                else:
                    self.set_cell_border(
                        cell,
                        top={"sz": 0, "val": "none"},
                        bottom={"sz": 0, "val": "none"},
                        left={"sz": 0, "val": "none"},
                        right={"sz": 0, "val": "none"}
                    )
        
        space_para = doc.add_paragraph()
        space_para.space_before = Pt(0)
        space_para.space_after = Pt(2)
        
        return table
    
    def format_p_value(self, p):
        """Format p-value according to APA style"""
        if p < 0.001:
            return "< .001"
        else:
            return f"{p:.3f}".lstrip('0')
    
    def save_to_docx(self):
        """Save ANOVA results to a DOCX file with TWO-COLUMN layout"""
        if self.anova_results is None:
            messagebox.showwarning("Warning", "No results to save! Run ANOVA first.")
            return
        
        filepath = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")],
            title="Save ANOVA Report"
        )
        
        if not filepath:
            return
        
        try:
            r = self.anova_results
            
            doc = Document()
            
            # Set narrow margins for more space
            sections = doc.sections
            for section in sections:
                section.top_margin = Inches(0.5)
                section.bottom_margin = Inches(0.5)
                section.left_margin = Inches(0.5)
                section.right_margin = Inches(0.5)
            
            # Set document defaults
            style = doc.styles['Normal']
            font = style.font
            font.name = 'Times New Roman'
            font.size = Pt(12)
            
            # Title (centered, spans full width)
            title = doc.add_heading(r['report_title'], 0)
            title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            title.space_after = Pt(2)
            for run in title.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(12)
                run.bold = True
            
            # Subtitle
            if r['report_subtitle']:
                subtitle = doc.add_paragraph(r['report_subtitle'])
                subtitle.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                subtitle.space_after = Pt(1)
                for run in subtitle.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(10)
            
            # Researcher name
            if r['researcher_name']:
                name_para = doc.add_paragraph(f"By: {r['researcher_name']}")
                name_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                name_para.space_after = Pt(6)
                for run in name_para.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(9)
            
            # CREATE TWO-COLUMN TABLE FOR LAYOUT
            layout_table = doc.add_table(rows=1, cols=2)
            layout_table.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            # Remove all borders from layout table
            for row in layout_table.rows:
                for cell in row.cells:
                    self.set_cell_border(
                        cell,
                        top={"sz": 0, "val": "none"},
                        bottom={"sz": 0, "val": "none"},
                        left={"sz": 0, "val": "none"},
                        right={"sz": 0, "val": "none"}
                    )
            
            # Set column widths
            layout_table.columns[0].width = Inches(3.5)
            layout_table.columns[1].width = Inches(3.5)
            
            # Get left and right cells
            left_cell = layout_table.rows[0].cells[0]
            right_cell = layout_table.rows[0].cells[1]
            
            # LEFT COLUMN CONTENT
            left_doc_content = []
            
            # Descriptive Statistics
            desc_heading = left_cell.add_paragraph()
            desc_run = desc_heading.add_run("Descriptive Statistics")
            desc_run.font.name = 'Times New Roman'
            desc_run.font.size = Pt(12)
            desc_run.bold = True
            desc_heading.space_after = Pt(2)
            
            desc_headers = ['Group', 'n', 'M', 'SD']
            desc_data = []
            
            for name, group in zip(r['group_names'], r['groups']):
                desc_data.append([
                    name,
                    str(len(group)),
                    f"{np.mean(group):.2f}",
                    f"{np.std(group, ddof=1):.2f}"
                ])
            
            self.create_mini_table(left_cell, desc_data, desc_headers)
            
            # ANOVA Table
            anova_heading = left_cell.add_paragraph()
            anova_heading.space_before = Pt(4)
            anova_run = anova_heading.add_run("One-Way ANOVA Results")
            anova_run.font.name = 'Times New Roman'
            anova_run.font.size = Pt(12)
            anova_run.bold = True
            anova_heading.space_after = Pt(2)
            
            anova_headers = ['Source', 'SS', 'df', 'MS', 'F', 'p', 'η²']
            anova_data = []
            
            anova_data.append([
                'Between',
                f"{r['SS_between']:.2f}",
                str(r['df_between']),
                f"{r['MS_between']:.2f}",
                f"{r['F_statistic']:.2f}",
                self.format_p_value(r['p_value']),
                f"{r['eta_squared']:.3f}"
            ])
            
            anova_data.append([
                'Within',
                f"{r['SS_within']:.2f}",
                str(r['df_within']),
                f"{r['MS_within']:.2f}",
                '', '', ''
            ])
            
            anova_data.append([
                'Total',
                f"{r['SS_total']:.2f}",
                str(r['df_between'] + r['df_within']),
                '', '', '', ''
            ])
            
            anova_table = self.create_mini_table(left_cell, anova_data, anova_headers)
            
            # Note for ANOVA
            note_para = left_cell.add_paragraph()
            note_para.space_before = Pt(0)
            note_para.space_after = Pt(4)
            note_run = note_para.add_run(f"Note. α = {r['alpha']}. η² = effect size.")
            note_run.font.name = 'Times New Roman'
            note_run.font.size = Pt(10)
            note_run.italic = True
            
            # Raw Data Table
            raw_heading = left_cell.add_paragraph()
            raw_heading.space_before = Pt(4)
            raw_run = raw_heading.add_run("Raw Data by Group")
            raw_run.font.name = 'Times New Roman'
            raw_run.font.size = Pt(11)
            raw_run.bold = True
            raw_heading.space_after = Pt(2)
            
            raw_headers = ['Group', 'n', 'Values']
            raw_data = []
            
            for name, group in zip(r['group_names'], r['groups']):
                values_str = ', '.join([f"{v:.1f}" for v in group])
                if len(values_str) > 40:
                    values_str = values_str[:37] + "..."
                
                raw_data.append([
                    name,
                    str(len(group)),
                    values_str
                ])
            
            self.create_mini_table(left_cell, raw_data, raw_headers)
            
            # RIGHT COLUMN CONTENT
            # Conclusion
            conclusion_heading = right_cell.add_paragraph()
            conclusion_run = conclusion_heading.add_run("Conclusion")
            conclusion_run.font.name = 'Times New Roman'
            conclusion_run.font.size = Pt(10)
            conclusion_run.bold = True
            conclusion_heading.space_after = Pt(2)
            
            conclusion_para = right_cell.add_paragraph()
            conclusion_para.space_after = Pt(1)
            conclusion_text = conclusion_para.add_run(r['conclusion'])
            conclusion_text.font.name = 'Times New Roman'
            conclusion_text.font.size = Pt(11)
            
            # APA format statistical statement
            stat_para = right_cell.add_paragraph()
            stat_para.space_after = Pt(4)
            
            if r['is_significant']:
                apa_text = (f"The one-way ANOVA was significant, F({r['df_between']}, {r['df_within']}) = "
                           f"{r['F_statistic']:.2f}, {self.format_p_value(r['p_value'])}, "
                           f"η² = {r['eta_squared']:.3f}.")
            else:
                apa_text = (f"The one-way ANOVA was not significant, F({r['df_between']}, {r['df_within']}) = "
                           f"{r['F_statistic']:.2f}, {self.format_p_value(r['p_value'])}, "
                           f"η² = {r['eta_squared']:.3f}.")
            
            stat_run = stat_para.add_run(apa_text)
            stat_run.font.name = 'Times New Roman'
            stat_run.font.size = Pt(11)
            stat_run.italic = True
            
            # Post Hoc Analysis (if significant)
            if r['is_significant'] and 'tukey' in r:
                posthoc_heading = right_cell.add_paragraph()
                posthoc_heading.space_before = Pt(4)
                posthoc_run = posthoc_heading.add_run("Post Hoc Analysis (Tukey HSD)")
                posthoc_run.font.name = 'Times New Roman'
                posthoc_run.font.size = Pt(11)
                posthoc_run.bold = True
                posthoc_heading.space_after = Pt(2)
                
                tukey = r['tukey']
                tukey_summary = tukey.summary()
                
                posthoc_headers = ['Comparison', 'Diff', 'CI Low', 'CI High', 'p', 'Sig']
                posthoc_data = []
                
                for row in tukey_summary.data[1:]:
                    group1 = str(row[0])
                    group2 = str(row[1])
                    meandiff = float(row[2])
                    lower = float(row[3])
                    upper = float(row[4])
                    reject = row[5]
                    p_adj = float(row[6]) if len(row) > 6 else 0.05
                    
                    comparison = f"{group1} vs {group2}"
                    significant = "Yes" if reject else "No"
                    
                    posthoc_data.append([
                        comparison,
                        f"{meandiff:.2f}",
                        f"{lower:.2f}",
                        f"{upper:.2f}",
                        self.format_p_value(p_adj),
                        significant
                    ])
                
                self.create_mini_table(right_cell, posthoc_data, posthoc_headers)
                
                # Note for post hoc
                note_para = right_cell.add_paragraph()
                note_para.space_before = Pt(0)
                note_run = note_para.add_run("Note. CI = 95% confidence interval.")
                note_run.font.name = 'Times New Roman'
                note_run.font.size = Pt(7)
                note_run.italic = True
            
            # Footer at bottom
            doc.add_paragraph()
            footer_para = doc.add_paragraph()
            footer_para.space_before = Pt(6)
            footer_run = footer_para.add_run(
                f"{datetime.now().strftime('%B %d, %Y at %I:%M %p')} | File: {filepath.split('/')[-1]}"
            )
            footer_run.font.name = 'Times New Roman'
            footer_run.font.size = Pt(7)
            footer_run.font.color.rgb = RGBColor(128, 128, 128)
            footer_run.italic = True
            footer_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            doc.save(filepath)
            
            self.file_location_label.configure(text=f"Last saved: {filepath}")
            
            messagebox.showinfo("Success", f"Two-column APA report saved!\n\n{filepath}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save document:\n{str(e)}")
    
    def create_mini_table(self, cell, data, headers):
        """Create a mini table inside a cell for two-column layout"""
        # Add the table to the cell
        table = cell.add_table(rows=len(data) + 1, cols=len(headers))
        
        # Set headers
        for i, header in enumerate(headers):
            header_cell = table.rows[0].cells[i]
            header_cell.text = header
            
            for paragraph in header_cell.paragraphs:
                paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                for run in paragraph.runs:
                    run.font.name = 'Times New Roman'
                    run.font.size = Pt(7)
                    run.bold = True
            
            self.set_cell_border(
                header_cell,
                top={"sz": 12, "val": "single", "color": "000000"},
                bottom={"sz": 8, "val": "single", "color": "000000"},
                left={"sz": 0, "val": "none"},
                right={"sz": 0, "val": "none"}
            )
        
        # Populate data
        for row_idx, row_data in enumerate(data, start=1):
            for col_idx, value in enumerate(row_data):
                data_cell = table.rows[row_idx].cells[col_idx]
                data_cell.text = str(value)
                
                for paragraph in data_cell.paragraphs:
                    if col_idx == 0:
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    else:
                        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(7)
                
                if row_idx == len(data):
                    self.set_cell_border(
                        data_cell,
                        top={"sz": 0, "val": "none"},
                        bottom={"sz": 12, "val": "single", "color": "000000"},
                        left={"sz": 0, "val": "none"},
                        right={"sz": 0, "val": "none"}
                    )
                else:
                    self.set_cell_border(
                        data_cell,
                        top={"sz": 0, "val": "none"},
                        bottom={"sz": 0, "val": "none"},
                        left={"sz": 0, "val": "none"},
                        right={"sz": 0, "val": "none"}
                    )
        
        return table


def main():
    """Main function to run the application"""
    app = ANOVAAnalyzer()
    app.mainloop()


if __name__ == "__main__":
    main()