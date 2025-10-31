# 📊 PDF to Excel Financial Data Extractor

A Python tool that automatically extracts **financial statements from PDF reports** (such as 10-Ks or annual reports) and converts them into **clean, structured Excel workbooks** — perfect for plugging into a **DCF (Discounted Cash Flow)** model.

---

## ✨ Features

- 🔍 **Intelligent PDF Detection** — Automatically detects text-based vs. scanned PDFs
- 📄 **Multi-Format Support** — Handles both native text PDFs and scanned/image-based documents with OCR
- 📊 **Financial Statement Extraction** — Extracts Income Statements, Balance Sheets, and Cash Flow Statements
- 🧹 **Data Cleaning** — Automatically cleans number formats, handles negative values in parentheses, removes commas
- 📑 **Organized Output** — Creates separate Excel sheets for each extracted table with auto-labeling (IS/BS/CF)
- 🎯 **DCF-Ready** — Outputs clean, structured data perfect for financial modeling

---

## 🧠 How It Works

### User Flow
1. **Place** your PDF file (e.g., `Apple_10K.pdf`) into the `input/` folder
2. **Run** the script with appropriate options
3. **Retrieve** your clean Excel file from the `output/` folder
4. **Link** the sheets directly into your financial model with Excel formulas

### Extraction Pipeline
1. **Detect Type** — Determines if the PDF is text-based (direct extraction) or scanned (requires OCR first)
2. **Extract Tables** — Uses `tabula-py` to read tables page by page, with `pdfplumber` as fallback for complex layouts
3. **Clean & Format** — Uses `pandas` and `numpy` to fix number formats, negative values in parentheses, and standardize column names
4. **Export to Excel** — Writes structured data to `.xlsx` file with `openpyxl`, placing each table on its own sheet
5. **Auto-Label Sheets** — Automatically detects and labels sheets as IS (Income Statement), BS (Balance Sheet), or CF (Cash Flow)

---

## ⚙️ Setup Instructions

### Prerequisites

You'll need the following installed on your system:

- **Python 3.9+** — [Download Python](https://www.python.org/downloads/)
- **Java Runtime (JRE or JDK)** — Required for `tabula-py`. [Download Java](https://www.java.com/en/download/)
- *(Optional)* **OCRmyPDF** — For processing scanned or image-based PDFs

### Installation Steps

#### 1. Clone or Download the Repository

Navigate to the project directory:
```bash
cd dcf-to-excel-app
```

#### 2. Create and Activate a Virtual Environment

Create a virtual environment to manage dependencies:

```bash
python -m venv venv
```

Activate the virtual environment:

| Operating System | Command |
| :--- | :--- |
| **Windows** | `venv\Scripts\activate` |
| **Mac / Linux** | `source venv/bin/activate` |

#### 3. Install Required Packages

Install all dependencies from `requirements.txt`:

```bash
pip install -r requirements.txt
```

The required packages include:
- `pandas` — Data manipulation and cleaning
- `numpy` — Numerical operations
- `tabula-py` — PDF table extraction
- `pdfplumber` — PDF text and table extraction fallback
- `PyMuPDF` — PDF metadata and text detection
- `openpyxl` — Excel file creation
- `ocrmypdf` — OCR for scanned PDFs
- `tqdm` — Progress bars

---

## 🚀 Usage

### Basic Usage

1. Place your source PDF file(s) in the `input/` folder
2. Run the script:

```bash
python pdf_to_excel.py "input/Company_10K.pdf" -o "output/Company_10K_output.xlsx"
```

### Command Line Options

| Flag | Description | Example |
| :--- | :--- | :--- |
| `-o`, `--output` | Output Excel file path (required) | `-o "output/file.xlsx"` |
| `--pages` | Extract only specific page ranges | `--pages "12-18,35"` |
| `--ocr` | Enable OCR for scanned/image PDFs | `--ocr` |

### Example Commands

**Extract all tables from a text-based PDF:**
```bash
python pdf_to_excel.py "input/Apple_10K.pdf" -o "output/Apple_10K.xlsx"
```

**Extract tables from specific pages:**
```bash
python pdf_to_excel.py "input/Company_Report.pdf" -o "output/Financials.xlsx" --pages "45-52,60-65"
```

**Process a scanned PDF with OCR:**
```bash
python pdf_to_excel.py "input/Scanned_Report.pdf" -o "output/Extracted.xlsx" --ocr
```

---

## 📁 Project Structure

```
dcf-to-excel-app/
├── input/              # Place your PDF files here
├── output/             # Extracted Excel files appear here
├── excel_models/       # Financial models directory
├── logs/               # Application logs
├── script/             # Script utilities
├── utils/              # Helper utilities
├── pdf_to_excel.py     # Main script
├── requirements.txt    # Python dependencies
└── README.md           # This file
```

---

## 🔧 Troubleshooting

### Common Issues

**Issue: `tabula-py` fails to extract tables**
- **Solution:** Ensure Java Runtime is installed and accessible from your PATH
- **Alternative:** The script will automatically fall back to `pdfplumber` for complex layouts

**Issue: OCR not working for scanned PDFs**
- **Solution:** Install `ocrmypdf` and ensure all dependencies are met: `pip install ocrmypdf`
- **Note:** OCR processing may take significantly longer than text-based extraction

**Issue: Numbers not formatted correctly**
- **Solution:** The script automatically cleans common number format issues, but you may need to manually adjust for unusual formats

**Issue: Tables split across multiple pages**
- **Solution:** The script attempts to detect and merge split tables, but complex layouts may require manual review

---

## 📝 Notes

- The tool works best with standard financial statement formats (10-K, 10-Q, annual reports)
- Processing time varies based on PDF size and complexity
- Scanned PDFs require significantly more processing time due to OCR
- Always review the extracted data before using it in financial models

---

## 📄 License

This project is provided as-is for educational and personal use.

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome!