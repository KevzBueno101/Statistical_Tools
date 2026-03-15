# 📊 Statistical Analysis Suite

A modern, sleek desktop application integrating professional statistical analysis tools for research and academic use. Built with Python and CustomTkinter — featuring a dark/light themed UI, APA-style reporting, and shared settings across all modules.

---

## ✨ Modules

| # | Module | File | Description |
|---|--------|------|-------------|
| 1 | **One-Way ANOVA** | `anova_analyzer.py` | Compare means across multiple groups with Tukey HSD post-hoc test |
| 2 | **t-Test Analysis** | `ttest_analyzer.py` | One-sample, Independent, and Paired t-tests with Cohen's d effect size |
| 3 | **Cronbach's Alpha** | `cronbach_alpha.py` | Internal consistency reliability with Likert frequency expander |
| 4 | **Chi-Square Test** | `chi_square_test.py` | Goodness-of-fit and test of independence with contingency tables |
| 5 | **Spearman Correlation** | `spearman_correl.py` | Rank-order correlation analysis with scatter plot visualization |
| 6 | **Cohen's Kappa** | `cohen_kappa.py` | Inter-rater agreement with JASP-style APA output |
| 7 | **Regression Analysis** | `regression_analysis.py` | Simple and multiple linear regression with diagnostics |
| 8 | **Database** | `database.py` | Data management and storage utilities |

---

## 🎨 UI Highlights

- **Dark/Light theme** toggle
- **Teal accent** color scheme (customizable via Settings)
- **Resizable panels** — drag the divider between columns
- **Shared ⚙ Settings** across all modules:
  - Theme (Dark / Light / System)
  - Accent Color (6 options: Teal, Blue, Purple, Orange, Rose, Green)
  - Font Family & Size
  - Sidebar Width (Compact / Normal / Wide)
  - Decimal Places (2 / 3 / 4)
  - APA p-value formatting toggle
  - Results text wrap mode

---

## 🚀 Getting Started

### Prerequisites
- Python 3.8 or higher
- pip

### Installation

```bash
# 1. Clone the repository
git clone https://github.com/KevzBueno101/Statistical_Tools.git
cd Statistical_Tools

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the app
python main.py
```

---

## 📁 Project Structure

```
Statistical_Tools/
│
├── main.py                          # Main menu launcher
├── app_settings.py                  # Shared settings module (all apps)
├── requirements.txt
├── README.md
│
├── modules/
│   ├── __init__.py
│   ├── anova_analyzer.py            # One-Way ANOVA
│   ├── ttest_analyzer.py            # t-Test Analysis
│   ├── cronbach_alpha.py            # Cronbach's Alpha
│   ├── chi_square_test.py           # Chi-Square Test
│   ├── spearman_correl.py           # Spearman Correlation
│   ├── cohen_kappa.py               # Cohen's Kappa
│   ├── regression_analysis.py       # Regression Analysis
│   └── database.py                  # Data Management
│
├── assets/                          # Icons and images
└── output/                          # Generated reports (auto-created)
```

---

## 📦 Dependencies

```
customtkinter
numpy
pandas
scipy
statsmodels
matplotlib
seaborn
reportlab
python-docx
openpyxl
```

Install all at once:
```bash
pip install -r requirements.txt
```

---

## 📄 Export Options

Each module supports:
- **DOCX** — APA-formatted Word document report
- **PDF** — Professional PDF report (via ReportLab)
- **Excel / CSV** — Raw data export

---

## ⚙️ Settings

A shared settings panel (`app_settings.py`) persists preferences across all modules via `.stat_suite_settings.json`:

| Setting | Options |
|---------|---------|
| Theme | Dark / Light / System |
| Accent Color | Teal, Blue, Purple, Orange, Rose, Green |
| Font Family | Segoe UI, Calibri, Helvetica, Arial, Courier New |
| Font Size | Small, Medium, Large, Extra Large |
| Sidebar Width | Compact, Normal, Wide |
| Decimal Places | 2, 3, 4 |
| Results Wrap | None, Word, Char |
| APA p-value | Toggle `p < .001` format |
| Show Tips | Toggle inline helper hints |

---

## 🔨 Build Standalone Executable (.exe)

```bash
pip install pyinstaller

pyinstaller --onefile --windowed ^
  --name "StatisticalSuite" ^
  --add-data "modules;modules" ^
  --add-data "assets;assets" ^
  --collect-all customtkinter ^
  --hidden-import=scipy ^
  --hidden-import=scipy.stats ^
  --hidden-import=statsmodels ^
  --hidden-import=matplotlib ^
  main.py
```

The `.exe` will be in the `dist/` folder.

---

## 🐛 Troubleshooting

| Issue | Fix |
|-------|-----|
| `Module not found` | Run `pip install -r requirements.txt` |
| `app_settings` import error | Make sure `app_settings.py` is in the **root** folder, same level as `main.py` |
| Settings not saving | Check write permissions in app folder |
| GUI looks wrong | Run `pip install --upgrade customtkinter` |
| Push rejected on git | Run `git pull origin main` first, then push again |

---

## 📖 How to Use

1. Run `python main.py` — main menu appears
2. Click any tool button to launch that module
3. Each module opens independently in its own window
4. Click **⚙ Settings** in any sidebar to customize the UI
5. Settings are shared and saved automatically across all modules

---

## 🎯 Version History

| Version | Date | Changes |
|---------|------|---------|
| **v2.0.0** | Mar 2026 | Modern UI redesign — dark sidebar, teal accent, resizable panels, shared settings, 8 modules |
| **v1.0.0** | Jan 2026 | Initial release with 5 integrated modules |

---

## 🔮 Planned Features

- [ ] Batch processing mode
- [ ] Custom report templates
- [ ] Data visualization gallery
- [ ] Descriptive statistics module
- [ ] Multiple language support

---

## 📄 License

For educational and research purposes.

---

