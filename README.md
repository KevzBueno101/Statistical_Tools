# Statistical Analysis Suite

A unified desktop application integrating five professional statistical analysis tools for research and data analysis.

## 📊 Features

### Included Analysis Tools:
1. **One-Way ANOVA Analyzer** - Compare means across multiple groups with post-hoc tests
2. **Cronbach's Alpha Test** - Assess internal consistency reliability with Likert scale expander
3. **Independent t-test** - Compare means between two independent groups
4. **Spearman's Correlation** - Analyze rank-order relationships with visualization
5. **Cohen's Kappa Calculator** - Measure inter-rater agreement with JASP-style reports

## 🚀 Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

### Setup Steps

1. **Extract the application files** to a folder

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**:
   ```bash
   python main.py
   ```

## 📁 Project Structure

```
StatisticalAnalysisApp/
│
├── main.py                          # Main menu application
├── modules/
│   ├── __init__.py
│   ├── anova_analyzer.py           # ANOVA module
│   ├── cronbach_alpha.py           # Cronbach's Alpha module
│   ├── ttest_analyzer.py           # t-test module
│   ├── spearman_correlation.py     # Spearman module
│   └── cohen_kappa.py              # Cohen's Kappa module
│
├── output/                         # Generated reports
├── requirements.txt
└── README.md
```

## 🔨 Building Standalone Executable (.exe)

### Using PyInstaller

1. **Install PyInstaller**:
   ```bash
   pip install pyinstaller
   ```

2. **Create the executable**:
   ```bash
   pyinstaller --onefile --windowed --name "StatisticalAnalysisSuite" main.py
   ```

3. **For a cleaner build with icon** (if you have icon.ico):
   ```bash
   pyinstaller --onefile --windowed --name "StatsApp" --icon=assets/icon.ico main.py
   ```

4. **Advanced build with all dependencies**:
   ```bash
   pyinstaller --onefile ^
               --windowed ^
               --name "StatsApp" ^
               --icon=assets/stats.ico ^
               --add-data "modules;modules" ^
               --hidden-import=scipy ^
               --hidden-import=scipy.stats ^
               --hidden-import=statsmodels ^
               --hidden-import=matplotlib ^
               --hidden-import=seaborn ^
               --collect-all customtkinter ^
               main.py
   ```

5. **Find your executable**:
   - The `.exe` file will be in the `dist/` folder
   - You can distribute this single file to users (no Python needed)

### Build Notes
- First build may take 5-10 minutes
- Resulting .exe will be 50-100MB (includes all dependencies)
- Test the .exe on a clean machine without Python installed
- Include a README with the .exe for end users

### Troubleshooting Build Issues

**If modules are not found:**
```bash
pyinstaller --onefile --windowed --paths=./modules main.py
```

**If customtkinter assets are missing:**
```bash
pyinstaller --collect-all customtkinter main.py
```

**For debugging:**
Remove `--windowed` flag to see console output:
```bash
pyinstaller --onefile main.py
```

## 📖 Usage Guide

### Starting the Application
1. Run `main.py` or the compiled `.exe`
2. The main menu will appear with all available tools
3. Click any tool button to launch that module
4. Each module operates independently

### Data Input Methods
- **Manual Entry**: Type data directly into input fields
- **CSV Import**: Load data from comma-separated files
- **Excel Import**: Load data from .xlsx/.xls files
- **Clipboard Paste**: Paste data from spreadsheets

### Output Formats
- **PDF Reports**: Professional formatted reports
- **DOCX Reports**: Editable Microsoft Word documents
- **Excel Exports**: Data tables and results
- **JSON**: Results data for further processing

## 🛠️ Technical Details

### Dependencies
- **CustomTkinter**: Modern GUI framework
- **NumPy/Pandas**: Data manipulation
- **SciPy**: Statistical computations
- **Statsmodels**: Advanced statistical models
- **Matplotlib/Seaborn**: Data visualization
- **ReportLab**: PDF generation
- **python-docx**: Word document generation

### Platform Support
- Windows 10/11 (fully tested)
- macOS 10.14+ (compatible)
- Linux (Ubuntu 20.04+, compatible)

### Performance
- Handles datasets up to 10,000 rows efficiently
- Real-time computation for most analyses
- Multi-threaded where applicable

## 🐛 Troubleshooting

### Common Issues

**"Module not found" error:**
- Ensure all dependencies are installed: `pip install -r requirements.txt`
- Check Python version: `python --version` (3.8+ required)

**GUI not displaying correctly:**
- Update CustomTkinter: `pip install --upgrade customtkinter`
- Check display scaling settings (Windows)

**Import errors in compiled .exe:**
- Rebuild with `--collect-all` flag for missing packages
- Ensure all modules are in correct directories

**File save/export errors:**
- Check write permissions in output directory
- Ensure sufficient disk space

## 📞 Support

For issues, feature requests, or contributions:
- Check existing issues in documentation
- Verify all dependencies are correctly installed
- Test on latest Python version (3.11+ recommended)

## 📄 License

This application is provided for educational and research purposes.
Individual modules may have their own licensing requirements.

## 🎯 Version History

- **v1.0.0** - Initial release with 5 integrated modules
  - One-Way ANOVA Analyzer
  - Cronbach's Alpha Test
  - Independent t-test
  - Spearman's Correlation
  - Cohen's Kappa Calculator

## 🔮 Future Enhancements

Planned features for future releases:
- Batch processing mode
- Custom report templates
- Data visualization gallery
- Multiple language support
- Cloud storage integration

---

**Built with ❤️ for researchers, statisticians, and data analysts**