# Slide2PDF Converter

A tool to convert PPTX/ODP files to PDF. It resolves common font spacing issues found in LibreOffice conversions, automatically removes watermarks, and fixes PDF metadata titles.

### 1. Environment Setup

#### **Install Python**
1.  Go to the [Python Official Download Page](https://www.python.org/downloads/windows/).
2.  Download the Latest Stable Release Windows Installer.
3.  Run the installer and **make sure to check "Add Python to PATH"** before clicking "Install Now".
4.  After installation, open `cmd` and type `python --version` to verify the installation.

#### **Install Required Packages**
Run the following command to install the conversion and PDF processing engines:

```bash
pip install aspose-slides pymupdf
```

*If the latest version fails, refer to the verified stable environment (Python 3.12.3):*
```bash
pip install -r requirements.txt
```

---

### 2. File Configuration

Download the following files from this GitHub repository and place them in the same folder:

*   **SlidesToPDF.py**: Core conversion and watermark cleaning logic.
*   **Convert.bat**: Windows launcher (if manual creation is needed):

```batch
@echo off
python "%~dp0SlidesToPDF.py" %*
pause
```

---

### 3. Usage and Security Notes

*   **How to Use**: Simply drag and drop PPTX/ODP files onto `Convert.bat` to generate the PDF.
*   **Local Processing**: The conversion logic runs entirely on your local CPU. No cloud uploads are involved, making it safe for processing sensitive documents.
*   **High-Quality Output**: Uses the `aspose-slides` engine to ensure correct layout and font spacing. `pymupdf` is used at the command level to precisely remove watermark tags without affecting underlying content.
