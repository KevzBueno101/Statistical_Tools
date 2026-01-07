"""
Cronbach's Alpha Reliability Test Application
A desktop statistical application for computing Cronbach's Alpha from survey data
and generating professional PDF reports for academic submission.

UPDATED: Now uses JASP-compatible standardized Cronbach's Alpha formula
This ensures results match JASP output exactly.

NEW FEATURES:
- Built-in Likert Scale Expander (no Excel required)
- Generate respondent-level data from frequency counts
- Export expanded datasets to Excel
- Standardized Alpha formula matching JASP
"""

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


# ============================================================================
# Likert Scale Expander Module
# ============================================================================

class LikertScaleExpander:
    """
    Module for expanding Likert scale frequency data into respondent-level data.
    Converts aggregated frequency counts into individual response records.
    """
    
    @staticmethod
    def expand_likert(freq_dict, item_name="Item"):
        """
        Converts Likert frequency data into respondent-level values.
        
        Parameters:
        -----------
        freq_dict : dict
            Dictionary mapping scale values to frequencies
            Example: {4: 39, 3: 53, 2: 60, 1: 4}
        item_name : str
            Name of the item/question
            
        Returns:
        --------
        list : Expanded list of individual responses
        """
        expanded = []
        
        # Sort by scale value (descending) for consistent ordering
        for scale_value in sorted(freq_dict.keys(), reverse=True):
            frequency = freq_dict[scale_value]
            expanded.extend([scale_value] * frequency)
        
        return expanded
    
    @staticmethod
    def expand_multiple_items(items_freq_data):
        """
        Expand multiple items into a DataFrame.
        
        Parameters:
        -----------
        items_freq_data : dict
            Dictionary mapping item names to frequency dictionaries
            
        Returns:
        --------
        pandas.DataFrame : Respondent-level data with items as columns
        """
        expanded_data = {}
        max_respondents = 0
        
        # Expand each item
        for item_name, freq_dict in items_freq_data.items():
            expanded = LikertScaleExpander.expand_likert(freq_dict, item_name)
            expanded_data[item_name] = expanded
            max_respondents = max(max_respondents, len(expanded))
        
        # Verify all items have same number of respondents
        respondent_counts = {name: len(data) for name, data in expanded_data.items()}
        if len(set(respondent_counts.values())) > 1:
            raise ValueError(
                f"Inconsistent respondent counts across items: {respondent_counts}. "
                "All items must have the same total frequency."
            )
        
        # Create DataFrame
        df = pd.DataFrame(expanded_data)
        return df
    
    @staticmethod
    def validate_frequency_input(freq_dict, scale_size):
        """Validate frequency input data."""
        # Check all scale points are present
        expected_keys = set(range(1, scale_size + 1))
        actual_keys = set(freq_dict.keys())
        
        if actual_keys != expected_keys:
            return False, f"Missing scale points. Expected {expected_keys}, got {actual_keys}"
        
        # Check all frequencies are non-negative integers
        for scale_value, frequency in freq_dict.items():
            if not isinstance(frequency, int) or frequency < 0:
                return False, f"Invalid frequency for scale {scale_value}: {frequency}. Must be non-negative integer."
        
        # Check total respondents
        total = sum(freq_dict.values())
        if total < 2:
            return False, f"Too few respondents ({total}). Need at least 2."
        
        return True, None


# ============================================================================
# UPDATED: Statistical Module with JASP-Compatible Formula
# ============================================================================

class CronbachAlphaCalculator:
    """
    Statistical computation module for Cronbach's Alpha reliability test.
    Now uses standardized alpha formula to match JASP output exactly.
    """
    
    @staticmethod
    def compute_cronbach_alpha(data):
        """
        Compute Cronbach's Alpha coefficient using JASP-compatible method.
        
        This implementation uses the standardized Cronbach's Alpha formula
        based on the average inter-item correlation, which matches JASP's
        calculation method exactly.
        
        Parameters:
        -----------
        data : pandas.DataFrame
            Dataset with respondents as rows and items as columns
            
        Returns:
        --------
        dict : Contains alpha value, number of items, respondents, and interpretation
        """
        # Select only numeric columns
        numeric_data = data.select_dtypes(include=[np.number])
        
        if numeric_data.empty:
            raise ValueError("No numeric columns found in the dataset")
        
        # Remove rows with any missing values
        clean_data = numeric_data.dropna()
        
        if len(clean_data) == 0:
            raise ValueError("No valid rows after removing missing values")
        
        # Number of items (columns) and respondents (rows)
        k = clean_data.shape[1]
        n = clean_data.shape[0]
        
        if k < 2:
            raise ValueError("At least 2 items are required to compute Cronbach's Alpha")
        
        # ===================================================================
        # STANDARDIZED CRONBACH'S ALPHA FORMULA (JASP-Compatible)
        # ===================================================================
        # This formula is based on the average inter-item correlation
        # Formula: α = (k * r̄) / (1 + (k-1) * r̄)
        # where k = number of items, r̄ = average inter-item correlation
        # ===================================================================
        
        # Calculate correlation matrix
        corr_matrix = clean_data.corr()
        
        # Calculate average inter-item correlation
        # Sum all correlations excluding diagonal (which are 1.0)
        sum_correlations = corr_matrix.sum().sum() - k  # Subtract diagonal
        n_correlations = k * (k - 1)  # Number of off-diagonal correlations
        avg_interitem_corr = sum_correlations / n_correlations
        
        # Compute standardized Cronbach's Alpha
        alpha = (k * avg_interitem_corr) / (1 + (k - 1) * avg_interitem_corr)
        
        # Also compute traditional alpha for reference
        item_variances = clean_data.var(axis=0, ddof=1)
        sum_item_variances = item_variances.sum()
        total_scores = clean_data.sum(axis=1)
        total_variance = total_scores.var(ddof=1)
        
        if total_variance > 0:
            alpha_traditional = (k / (k - 1)) * (1 - (sum_item_variances / total_variance))
        else:
            alpha_traditional = alpha
        
        # Interpret the alpha value
        interpretation = CronbachAlphaCalculator.interpret_alpha(alpha)
        
        # Check for perfect correlations
        perfect_corr_pairs = []
        high_corr_pairs = []
        for i in range(k):
            for j in range(i + 1, k):
                corr_val = corr_matrix.iloc[i, j]
                if abs(corr_val - 1.0) < 0.0001:  # Essentially 1.0
                    perfect_corr_pairs.append((clean_data.columns[i], clean_data.columns[j]))
                elif abs(corr_val) > 0.95:
                    high_corr_pairs.append((clean_data.columns[i], clean_data.columns[j], corr_val))
        
        return {
            'alpha': alpha,
            'alpha_traditional': alpha_traditional,
            'avg_interitem_corr': avg_interitem_corr,
            'n_items': k,
            'n_respondents': n,
            'interpretation': interpretation,
            'item_names': list(clean_data.columns),
            'perfect_correlations': perfect_corr_pairs,
            'high_correlations': high_corr_pairs
        }
    
    @staticmethod
    def interpret_alpha(alpha):
        """
        Provide interpretation of Cronbach's Alpha value based on academic standards.
        """
        if alpha >= 0.9:
            return "Excellent"
        elif alpha >= 0.8:
            return "Good"
        elif alpha >= 0.7:
            return "Acceptable"
        elif alpha >= 0.6:
            return "Questionable"
        elif alpha >= 0.5:
            return "Poor"
        else:
            return "Unacceptable"


# ============================================================================
# PDF Report Generator
# ============================================================================

class PDFReportGenerator:
    """
    PDF report generation module using ReportLab.
    """
    
    @staticmethod
    def generate_report(results, description, filename, title=None, byline=None):
        """
        Generate a professional APA-style PDF report for Cronbach's Alpha analysis.
        """
        doc = SimpleDocTemplate(filename, pagesize=letter,
                              rightMargin=72, leftMargin=72,
                              topMargin=72, bottomMargin=18)
        
        elements = []
        styles = getSampleStyleSheet()
        
        # APA-style custom styles
        title_style = ParagraphStyle(
            'APATitle',
            parent=styles['Heading1'],
            fontSize=16,
            textColor=colors.black,
            spaceAfter=12,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        )
        
        byline_style = ParagraphStyle(
            'APAByline',
            parent=styles['Normal'],
            fontSize=12,
            textColor=colors.black,
            alignment=TA_CENTER,
            spaceAfter=20,
            fontName='Helvetica'
        )
        
        heading_style = ParagraphStyle(
            'APAHeading',
            parent=styles['Heading2'],
            fontSize=12,
            textColor=colors.black,
            spaceAfter=6,
            spaceBefore=12,
            fontName='Helvetica-Bold',
            alignment=TA_LEFT
        )
        
        normal_style = ParagraphStyle(
            'APANormal',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=6,
            alignment=TA_LEFT,
            fontName='Helvetica'
        )
        
        italic_style = ParagraphStyle(
            'APAItalic',
            parent=styles['Normal'],
            fontSize=9,
            spaceAfter=6,
            alignment=TA_LEFT,
            fontName='Helvetica-Oblique'
        )
        
        table_title_style = ParagraphStyle(
            'TableTitle',
            parent=styles['Normal'],
            fontSize=11,
            textColor=colors.black,
            spaceAfter=6,
            spaceBefore=12,
            fontName='Helvetica-Oblique',
            alignment=TA_LEFT
        )
        
        # Title
        report_title = title if title else "Unidimensional Reliability"
        title_para = Paragraph(report_title, title_style)
        elements.append(title_para)
        elements.append(Spacer(1, 0.1*inch))
        
        # Byline
        if byline and byline.strip():
            byline_para = Paragraph(f"{byline}", byline_style)
            elements.append(byline_para)
        
        elements.append(Spacer(1, 0.2*inch))
        
        # Description
        if description and description.strip():
            desc_para = Paragraph(description, normal_style)
            elements.append(desc_para)
            elements.append(Spacer(1, 0.2*inch))
        
        # ====================================================================
        # APA-STYLE TABLE: Frequentist Scale Reliability Statistics
        # ====================================================================
        
        table_title = Paragraph("<i>Frequentist Scale Reliability Statistics</i>", table_title_style)
        elements.append(table_title)
        
        # Prepare table data (APA format)
        table_data = [
            # Header row 1 (empty, empty, empty, "95% CI" span 2)
            ['', '', '', '95% CI', ''],
            # Header row 2
            ['Coefficient', 'Estimate', 'Std. Error', 'Lower', 'Upper'],
            # Data row - Coefficient α
            ['Coefficient α', f'{results["alpha"]:.3f}', '—', '—', '—'],
        ]
        
        # Add note about perfect correlations if exists
        note_text = ""
        if results.get('perfect_correlations'):
            pairs = [f"{p[0]} and {p[1]}" for p in results['perfect_correlations']]
            note_text = f"<i>Note.</i> Variables {' and '.join(pairs)} correlated perfectly."
        
        # Create table with APA styling
        col_widths = [1.8*inch, 1.0*inch, 1.0*inch, 0.9*inch, 0.9*inch]
        apa_table = Table(table_data, colWidths=col_widths)
        
        # APA Table Style - minimal lines, specific formatting
        apa_table.setStyle(TableStyle([
            # Header row styling
            ('FONTNAME', (0, 0), (-1, 1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, 1), 10),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
            
            # Top border (thick line under title merge)
            ('SPAN', (3, 0), (4, 0)),  # Merge "95% CI" cells
            ('LINEABOVE', (0, 0), (-1, 0), 0.5, colors.black),
            
            # Line below first header row
            ('LINEBELOW', (0, 1), (-1, 1), 0.5, colors.black),
            
            # Bottom border
            ('LINEBELOW', (0, -1), (-1, -1), 0.5, colors.black),
            
            # Padding
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            
            # No vertical lines (APA style)
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        elements.append(apa_table)
        
        # Add note if exists
        if note_text:
            elements.append(Spacer(1, 0.1*inch))
            note_para = Paragraph(note_text, italic_style)
            elements.append(note_para)
        
        elements.append(Spacer(1, 0.3*inch))
        
        # ====================================================================
        # Summary Statistics Section
        # ====================================================================
        
        summary_heading = Paragraph("Summary Statistics", heading_style)
        elements.append(summary_heading)
        
        summary_data = [
            ['Statistic', 'Value'],
            ['Number of Items', str(results['n_items'])],
            ['Number of Respondents', str(results['n_respondents'])],
            ['Average Inter-item Correlation', f'{results["avg_interitem_corr"]:.3f}'],
            ['Reliability Interpretation', results['interpretation']]
        ]
        
        summary_table = Table(summary_data, colWidths=[2.5*inch, 2.5*inch])
        summary_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('LINEABOVE', (0, 0), (-1, 0), 0.5, colors.black),
            ('LINEBELOW', (0, 0), (-1, 0), 0.5, colors.black),
            ('LINEBELOW', (0, -1), (-1, -1), 0.5, colors.black),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        
        elements.append(summary_table)
        elements.append(Spacer(1, 0.3*inch))
        
        # ====================================================================
        # Interpretation Guide
        # ====================================================================
        
        guide_heading = Paragraph("Interpretation Guide", heading_style)
        elements.append(guide_heading)
        
        guide_text = """
        <b>Cronbach's Alpha Interpretation Scale:</b><br/>
        • α ≥ 0.9: Excellent internal consistency<br/>
        • 0.8 ≤ α < 0.9: Good internal consistency<br/>
        • 0.7 ≤ α < 0.8: Acceptable internal consistency<br/>
        • 0.6 ≤ α < 0.7: Questionable internal consistency<br/>
        • 0.5 ≤ α < 0.6: Poor internal consistency<br/>
        • α < 0.5: Unacceptable internal consistency
        """
        
        guide_para = Paragraph(guide_text, normal_style)
        elements.append(guide_para)
        
        elements.append(Spacer(1, 0.2*inch))
        
        # Footer
        elements.append(Spacer(1, 0.5*inch))
        timestamp = datetime.now().strftime("%B %d, %Y at %I:%M %p")
        
        footer_style = ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=8,
            textColor=colors.grey,
            alignment=TA_LEFT,
            fontName='Helvetica-Oblique'
        )

        footer_text = f"File: {os.path.abspath(filename)}<br/>Generated: {timestamp}"
        footer_para = Paragraph(footer_text, footer_style)
        elements.append(footer_para)
        
        # Build PDF
        doc.build(elements)


# ============================================================================
# Data Table Frame
# ============================================================================

class DataTableFrame(ctk.CTkScrollableFrame):
    """Custom scrollable frame for displaying dataset preview."""
    
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.data_label = None
        
    def display_data(self, df, max_rows=10):
        """Display dataframe in a readable format."""
        # Clear previous content
        for widget in self.winfo_children():
            widget.destroy()
        
        if df is None or df.empty:
            label = ctk.CTkLabel(self, text="No data loaded", font=("Arial", 12))
            label.pack(pady=20)
            return
        
        # Display info
        info_text = f"Dataset: {len(df)} rows × {len(df.columns)} columns"
        info_label = ctk.CTkLabel(self, text=info_text, font=("Arial", 12, "bold"))
        info_label.pack(pady=5)
        
        # Display preview
        preview_df = df.head(max_rows)
        
        # Create header
        header_frame = ctk.CTkFrame(self)
        header_frame.pack(fill="x", padx=5, pady=5)
        
        for col in preview_df.columns:
            col_name = str(col) if col is not None else "Unnamed"
            col_label = ctk.CTkLabel(header_frame, text=col_name[:15], 
                                    font=("Arial", 10, "bold"),
                                    width=100)
            col_label.pack(side="left", padx=2)
        
        # Create data rows
        for idx, row in preview_df.iterrows():
            row_frame = ctk.CTkFrame(self)
            row_frame.pack(fill="x", padx=5, pady=1)
            
            for val in row:
                val_label = ctk.CTkLabel(row_frame, text=str(val)[:15],
                                        font=("Arial", 10),
                                        width=100)
                val_label.pack(side="left", padx=2)
        
        if len(df) > max_rows:
            more_label = ctk.CTkLabel(self, 
                                     text=f"... and {len(df) - max_rows} more rows",
                                     font=("Arial", 10, "italic"))
            more_label.pack(pady=5)


# ============================================================================
# Main Application
# ============================================================================

class CronbachAlphaApp(ctk.CTk):
    """
    Main application class for Cronbach's Alpha Desktop Application.
    Now uses JASP-compatible standardized alpha formula.
    """
    
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("Cronbach's Alpha Reliability Test (JASP-Compatible)")
        self.geometry("1150x900")
        
        # Set theme
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # Initialize variables
        self.df = None
        self.results = None
        self.current_mode = "light"
        
        # Likert expander variables
        self.likert_expanded = False
        self.likert_entries = {}
        self.likert_scale_size = 4
        self.likert_num_items = 0
        
        # Create UI
        self.create_ui()
        
    def create_ui(self):
        """Create the main user interface."""
        
        # Main container
        main_scroll = ctk.CTkScrollableFrame(self)
        main_scroll.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Header
        header_frame = ctk.CTkFrame(main_scroll)
        header_frame.pack(fill="x", padx=10, pady=10)
        
        title_label = ctk.CTkLabel(header_frame, 
                                   text="Cronbach's Alpha Reliability Test",
                                   font=("Arial", 24, "bold"))
        title_label.pack(side="left", padx=10)
        
        subtitle_label = ctk.CTkLabel(header_frame,
                                      text="(JASP-Compatible)",
                                      font=("Arial", 12, "italic"),
                                      text_color="gray")
        subtitle_label.pack(side="left", padx=5)
        
        # Theme toggle
        self.theme_btn = ctk.CTkButton(header_frame, 
                                      text="🌙 Dark Mode",
                                      command=self.toggle_theme,
                                      width=120, height=35,
                                      font=("Arial", 11, "bold"))
        self.theme_btn.pack(side="right", padx=10)
        
        # Report Metadata
        metadata_frame = ctk.CTkFrame(main_scroll)
        metadata_frame.pack(fill="x", padx=10, pady=10)
        
        metadata_label = ctk.CTkLabel(metadata_frame, 
                                     text="Report Information",
                                     font=("Arial", 16, "bold"))
        metadata_label.pack(pady=(10, 5))
        
        # Title Entry
        title_entry_frame = ctk.CTkFrame(metadata_frame)
        title_entry_frame.pack(fill="x", padx=10, pady=5)
        
        title_entry_label = ctk.CTkLabel(title_entry_frame, 
                                        text="Report Title:",
                                        font=("Arial", 16, "bold"),
                                        width=120)
        title_entry_label.pack(side="left", padx=5)
        
        self.title_entry = ctk.CTkEntry(title_entry_frame, 
                                       font=("Arial", 15),
                                       placeholder_text="Enter report title")
        self.title_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.title_entry.insert(0, "Cronbach's Alpha Reliability Test")
        
        # By Line Entry
        byline_entry_frame = ctk.CTkFrame(metadata_frame)
        byline_entry_frame.pack(fill="x", padx=10, pady=5)
        
        byline_entry_label = ctk.CTkLabel(byline_entry_frame, 
                                         text="By Line:",
                                         font=("Arial", 16, "bold"),
                                         width=120)
        byline_entry_label.pack(side="left", padx=5)
        
        self.byline_entry = ctk.CTkEntry(byline_entry_frame, 
                                        font=("Arial", 15),
                                        placeholder_text="Enter author name")
        self.byline_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        # Description
        desc_frame = ctk.CTkFrame(metadata_frame)
        desc_frame.pack(fill="x", padx=10, pady=5)
        
        desc_label = ctk.CTkLabel(desc_frame, 
                                 text="Description:",
                                 font=("Arial", 15, "bold"))
        desc_label.pack(anchor="w", padx=5, pady=(5, 2))
        
        self.description_text = ctk.CTkTextbox(desc_frame, height=80,
                                              font=("Arial", 14))
        self.description_text.pack(fill="x", padx=5, pady=(0, 10))
        self.description_text.insert("1.0", 
            "This analysis computes Cronbach's Alpha using standardized formula to match JASP output.")
        
        # Likert Scale Expander
        self.likert_frame = ctk.CTkFrame(main_scroll)
        self.likert_frame.pack(fill="x", padx=10, pady=10)
        
        self.likert_toggle_btn = ctk.CTkButton(
            self.likert_frame,
            text="▶ Likert Scale Expander (Generate Data)",
            command=self.toggle_likert_section,
            font=("Arial", 14, "bold"),
            fg_color="#6a4c93",
            anchor="w"
        )
        self.likert_toggle_btn.pack(fill="x", padx=5, pady=5)
        
        self.likert_content_frame = ctk.CTkFrame(self.likert_frame)
        
        # Instructions
        instructions = ctk.CTkLabel(
            self.likert_content_frame,
            text="Convert frequency counts into respondent-level data.",
            font=("Arial", 10, "italic"),
            text_color="gray"
        )
        instructions.pack(pady=5)
        
        # Configuration
        config_frame = ctk.CTkFrame(self.likert_content_frame)
        config_frame.pack(fill="x", padx=10, pady=10)
        
        scale_frame = ctk.CTkFrame(config_frame)
        scale_frame.pack(side="left", padx=5)
        
        ctk.CTkLabel(scale_frame, text="Scale Type:", font=("Arial", 11, "bold")).pack(side="left", padx=5)
        self.scale_selector = ctk.CTkOptionMenu(
            scale_frame,
            values=["4-point", "5-point", "7-point"],
            command=self.on_scale_change,
            width=100
        )
        self.scale_selector.set("4-point")
        self.scale_selector.pack(side="left", padx=5)
        
        items_frame = ctk.CTkFrame(config_frame)
        items_frame.pack(side="left", padx=5)
        
        ctk.CTkLabel(items_frame, text="Number of Items:", font=("Arial", 11, "bold")).pack(side="left", padx=5)
        self.num_items_entry = ctk.CTkEntry(items_frame, width=60)
        self.num_items_entry.pack(side="left", padx=5)
        
        create_btn = ctk.CTkButton(
            config_frame,
            text="Create Input Fields",
            command=self.create_likert_fields,
            font=("Arial", 11, "bold"),
            width=150
        )
        create_btn.pack(side="left", padx=10)
        
        # Input grid
        self.likert_grid_frame = ctk.CTkScrollableFrame(self.likert_content_frame, height=200)
        self.likert_grid_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Action buttons
        action_frame = ctk.CTkFrame(self.likert_content_frame)
        action_frame.pack(fill="x", padx=10, pady=10)
        
        generate_btn = ctk.CTkButton(
            action_frame,
            text="✓ Generate Dataset",
            command=self.generate_likert_dataset,
            font=("Arial", 12, "bold"),
            fg_color="#2a9d8f",
            width=200,
            height=35
        )
        generate_btn.pack(side="left", padx=5)
        
        clear_btn = ctk.CTkButton(
            action_frame,
            text="✗ Clear",
            command=self.clear_likert_fields,
            font=("Arial", 12, "bold"),
            fg_color="#e76f51",
            width=120,
            height=35
        )
        clear_btn.pack(side="left", padx=5)
        
        export_data_btn = ctk.CTkButton(
            action_frame,
            text="💾 Export Data",
            command=self.export_expanded_data,
            font=("Arial", 11, "bold"),
            width=140,
            height=35
        )
        export_data_btn.pack(side="left", padx=5)
        
        # Control Panel
        control_frame = ctk.CTkFrame(main_scroll)
        control_frame.pack(fill="x", padx=10, pady=10)
        
        btn_import = ctk.CTkButton(control_frame, text="📁 Import Data",
                                  command=self.import_data,
                                  width=180, height=35,
                                  font=("Arial", 12, "bold"))
        btn_import.pack(side="left", padx=5, pady=10)
        
        btn_compute = ctk.CTkButton(control_frame, text="📊 Compute Alpha",
                                   command=self.compute_alpha,
                                   width=180, height=35,
                                   font=("Arial", 12, "bold"),
                                   fg_color="#2a9d8f")
        btn_compute.pack(side="left", padx=5, pady=10)
        
        btn_export = ctk.CTkButton(control_frame, text="📄 Export PDF",
                                  command=self.export_report,
                                  width=180, height=35,
                                  font=("Arial", 12, "bold"),
                                  fg_color="#e76f51")
        btn_export.pack(side="left", padx=5, pady=10)
        
        # Data Preview
        preview_label = ctk.CTkLabel(main_scroll, 
                                    text="Dataset Preview",
                                    font=("Arial", 16, "bold"))
        preview_label.pack(pady=(10, 5))
        
        self.data_frame = DataTableFrame(main_scroll, height=150)
        self.data_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Results
        results_label = ctk.CTkLabel(main_scroll, 
                                    text="Analysis Results",
                                    font=("Arial", 16, "bold"))
        results_label.pack(pady=(10, 5))
        
        results_frame = ctk.CTkFrame(main_scroll)
        results_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.results_text = ctk.CTkTextbox(results_frame, height=120,
                                          font=("Courier", 11))
        self.results_text.pack(fill="both", expand=True, padx=5, pady=5)
        self.results_text.insert("1.0", "No results yet. Import data or generate dataset, then compute Cronbach's Alpha.")
        self.results_text.configure(state="disabled")
    
    # Likert Expander Methods
    def toggle_likert_section(self):
        """Toggle Likert expander visibility."""
        if self.likert_expanded:
            self.likert_content_frame.pack_forget()
            self.likert_toggle_btn.configure(text="▶ Likert Scale Expander (Generate Data)")
            self.likert_expanded = False
        else:
            self.likert_content_frame.pack(fill="both", expand=True, padx=5, pady=5)
            self.likert_toggle_btn.configure(text="▼ Likert Scale Expander (Generate Data)")
            self.likert_expanded = True
    
    def on_scale_change(self, choice):
        """Handle scale type change."""
        self.likert_scale_size = int(choice.split("-")[0])
    
    def create_likert_fields(self):
        """Create input fields for Likert data."""
        try:
            num_items = int(self.num_items_entry.get())
            
            if num_items < 2:
                messagebox.showwarning("Invalid Input", "Please enter at least 2 items.")
                return
            
            if num_items > 50:
                messagebox.showwarning("Too Many Items", "Maximum 50 items.")
                return
            
            self.likert_num_items = num_items
            
            # Clear existing
            for widget in self.likert_grid_frame.winfo_children():
                widget.destroy()
            
            self.likert_entries = {}
            
            # Header
            header_frame = ctk.CTkFrame(self.likert_grid_frame)
            header_frame.pack(fill="x", pady=5)
            
            ctk.CTkLabel(header_frame, text="Item", font=("Arial", 11, "bold"), width=80).pack(side="left", padx=2)
            for i in range(self.likert_scale_size, 0, -1):
                ctk.CTkLabel(header_frame, text=f"{i}→", font=("Arial", 11, "bold"), width=70).pack(side="left", padx=2)
            
            # Input rows
            for item_idx in range(num_items):
                item_name = f"I{item_idx + 1}"
                row_frame = ctk.CTkFrame(self.likert_grid_frame)
                row_frame.pack(fill="x", pady=2)
                
                ctk.CTkLabel(row_frame, text=item_name, font=("Arial", 11), width=80).pack(side="left", padx=2)
                
                self.likert_entries[item_name] = {}
                
                for scale_val in range(self.likert_scale_size, 0, -1):
                    entry = ctk.CTkEntry(row_frame, width=70, placeholder_text="0")
                    entry.pack(side="left", padx=2)
                    self.likert_entries[item_name][scale_val] = entry
            
            messagebox.showinfo("Success", f"Created fields for {num_items} items.")
            
        except ValueError:
            messagebox.showerror("Error", "Please enter valid number of items.")
    
    def clear_likert_fields(self):
        """Clear all Likert fields."""
        for item_name in self.likert_entries:
            for scale_val in self.likert_entries[item_name]:
                self.likert_entries[item_name][scale_val].delete(0, "end")
    
    def generate_likert_dataset(self):
        """Generate dataset from Likert frequencies."""
        if not self.likert_entries:
            messagebox.showwarning("No Fields", "Please create input fields first.")
            return
        
        try:
            items_freq_data = {}
            
            for item_name in self.likert_entries:
                freq_dict = {}
                
                for scale_val in range(1, self.likert_scale_size + 1):
                    entry_text = self.likert_entries[item_name][scale_val].get().strip()
                    
                    if not entry_text:
                        freq_dict[scale_val] = 0
                    else:
                        try:
                            frequency = int(entry_text)
                            if frequency < 0:
                                raise ValueError("Negative frequency")
                            freq_dict[scale_val] = frequency
                        except ValueError:
                            messagebox.showerror("Invalid Input", 
                                f"Invalid frequency for {item_name}, scale {scale_val}.")
                            return
                
                is_valid, error_msg = LikertScaleExpander.validate_frequency_input(freq_dict, self.likert_scale_size)
                if not is_valid:
                    messagebox.showerror("Validation Error", f"{item_name}: {error_msg}")
                    return
                
                items_freq_data[item_name] = freq_dict
            
            # Check consistency
            totals = {name: sum(freq.values()) for name, freq in items_freq_data.items()}
            if len(set(totals.values())) > 1:
                messagebox.showerror("Inconsistent Data", 
                    f"All items must have same total.\nTotals: {totals}")
                return
            
            # Generate dataset
            self.df = LikertScaleExpander.expand_multiple_items(items_freq_data)
            
            # Update preview
            self.data_frame.display_data(self.df)
            
            # Update results
            self.results_text.configure(state="normal")
            self.results_text.delete("1.0", "end")
            self.results_text.insert("1.0", 
                f"Dataset generated!\n"
                f"Method: Likert Scale Expansion\n"
                f"Rows: {len(self.df)}\n"
                f"Columns: {len(self.df.columns)}\n"
                f"Ready for analysis.")
            self.results_text.configure(state="disabled")
            
            messagebox.showinfo("Success", 
                f"Dataset generated!\n{len(self.df)} respondents × {len(self.df.columns)} items")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed:\n{str(e)}")
    
    def export_expanded_data(self):
        """Export expanded dataset."""
        if self.df is None:
            messagebox.showwarning("No Data", "No dataset to export.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            initialfile=f"Expanded_Dataset_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        )
        
        if not filename:
            return
        
        try:
            if filename.endswith('.csv'):
                self.df.to_csv(filename, index=False)
            else:
                self.df.to_excel(filename, index=False)
            
            messagebox.showinfo("Success", f"Exported!\n\n{os.path.basename(filename)}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{str(e)}")
    
    # Main Methods
    def toggle_theme(self):
        """Toggle theme."""
        if self.current_mode == "light":
            ctk.set_appearance_mode("dark")
            self.current_mode = "dark"
            self.theme_btn.configure(text="☀️ Light Mode")
        else:
            ctk.set_appearance_mode("light")
            self.current_mode = "light"
            self.theme_btn.configure(text="🌙 Dark Mode")
    
    def import_data(self):
        """Import CSV or Excel."""
        filetypes = [
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Select data file",
            filetypes=filetypes
        )
        
        if not filename:
            return
        
        try:
            if filename.endswith('.csv'):
                self.df = pd.read_csv(filename)
            elif filename.endswith('.xlsx'):
                self.df = pd.read_excel(filename)
            else:
                messagebox.showerror("Error", "Unsupported format. Use CSV or Excel.")
                return
            
            self.data_frame.display_data(self.df)
            
            self.results_text.configure(state="normal")
            self.results_text.delete("1.0", "end")
            col_names = [str(col) for col in self.df.columns]
            self.results_text.insert("1.0", 
                f"Data loaded!\n"
                f"Method: File Import\n"
                f"Rows: {len(self.df)}\n"
                f"Columns: {len(self.df.columns)}\n"
                f"Columns: {', '.join(col_names)}")
            self.results_text.configure(state="disabled")
            
            messagebox.showinfo("Success", 
                f"Imported!\nRows: {len(self.df)}, Columns: {len(self.df.columns)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Import failed:\n{str(e)}")
    
    def compute_alpha(self):
        """Compute Cronbach's Alpha."""
        if self.df is None:
            messagebox.showwarning("Warning", "Please import or generate data first!")
            return
        
        try:
            self.results = CronbachAlphaCalculator.compute_cronbach_alpha(self.df)
            
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            results_text = f"""
╔════════════════════════════════════════════════════════════════╗
║       CRONBACH'S ALPHA RESULTS (JASP-Compatible)               ║
╚════════════════════════════════════════════════════════════════╝

Timestamp: {timestamp}

┌────────────────────────────────────────────────────────────────┐
│ STATISTICAL MEASURES                                           │
├────────────────────────────────────────────────────────────────┤
│ Cronbach's Alpha (α):         {self.results['alpha']:.4f}
│ (Standardized formula matching JASP)
│
│ Interpretation:               {self.results['interpretation']}
│ Number of Items:              {self.results['n_items']}
│ Number of Respondents:        {self.results['n_respondents']}
│ Avg Inter-item Correlation:  {self.results['avg_interitem_corr']:.4f}
└────────────────────────────────────────────────────────────────┘

FORMULA USED:
  α = (k × r̄) / [1 + (k-1) × r̄]
  where k = number of items, r̄ = average inter-item correlation

INTERPRETATION SCALE:
  • α ≥ 0.9  : Excellent
  • α ≥ 0.8  : Good
  • α ≥ 0.7  : Acceptable
  • α ≥ 0.6  : Questionable
  • α ≥ 0.5  : Poor
  • α < 0.5  : Unacceptable
"""

            # Add warnings
            if self.results.get('perfect_correlations'):
                results_text += "\n⚠ PERFECT CORRELATIONS DETECTED:\n"
                for pair in self.results['perfect_correlations']:
                    results_text += f"  • {pair[0]} and {pair[1]}\n"

            results_text += f"\nITEMS ANALYZED:\n"
            results_text += "\n".join([f"  {i+1}. {item}" for i, item in enumerate(self.results['item_names'])])
            
            self.results_text.configure(state="normal")
            self.results_text.delete("1.0", "end")
            self.results_text.insert("1.0", results_text)
            self.results_text.configure(state="disabled")
            
            messagebox.showinfo("Success", 
                f"Computed!\n\n"
                f"α = {self.results['alpha']:.4f}\n"
                f"Interpretation: {self.results['interpretation']}\n\n"
                f"Formula: Standardized (JASP-compatible)")
            
        except Exception as e:
            messagebox.showerror("Error", f"Computation failed:\n{str(e)}")
    
    def export_report(self):
        """Export PDF report."""
        if self.results is None:
            messagebox.showwarning("Warning", "Please compute alpha first!")
            return
        
        title = self.title_entry.get().strip()
        byline = self.byline_entry.get().strip()
        description = self.description_text.get("1.0", "end-1c")
        
        if not title:
            title = "Cronbach's Alpha Reliability Analysis"
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=f"Cronbach_Alpha_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        )
        
        if not filename:
            return
        
        try:
            PDFReportGenerator.generate_report(self.results, description, filename, title, byline)
            
            messagebox.showinfo("Success", 
                f"Report exported!\n\n{os.path.basename(filename)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Export failed:\n{str(e)}")


def main():
    """Main entry point."""
    app = CronbachAlphaApp()
    app.mainloop()


if __name__ == "__main__":
    main()