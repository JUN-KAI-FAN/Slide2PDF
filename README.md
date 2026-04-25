# Slide2PDF

A tool that converts your presentations to PDF with readable formatting, solving the LibreOffice compatibility issues you face when Office isn't available on your company or personal machine.

## Features

- **Pure-Python pipeline** (default): uses `python-pptx` + `reportlab` — no external dependencies required.
- **LibreOffice backend** (optional): pass `--libreoffice` to delegate conversion to LibreOffice headless for pixel-perfect fidelity.
- Handles `.pptx` and `.odp` files.
- Preserves slide titles, subtitles, body text, and bullet points.
- One output page per slide, landscape A4 format.

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```
python slide2pdf.py INPUT [OUTPUT] [--libreoffice]
```

| Argument | Description |
|---|---|
| `INPUT` | Path to the input presentation (`.pptx` or `.odp`). |
| `OUTPUT` | *(optional)* Output PDF path. Defaults to `<INPUT_STEM>.pdf` in the same directory. |
| `--libreoffice` | Use LibreOffice headless as the conversion backend (must be installed). |

### Examples

```bash
# Convert using the pure-Python pipeline
python slide2pdf.py slides/my_deck.pptx

# Specify an output path
python slide2pdf.py slides/my_deck.pptx output/my_deck.pdf

# Use LibreOffice for pixel-perfect conversion
python slide2pdf.py slides/my_deck.pptx --libreoffice
```

## Running the tests

```bash
pip install pytest
python -m pytest tests/ -v
```

## Why not just use LibreOffice directly?

LibreOffice conversion can produce poor results with fonts, layout, and bullet points when the source file contains Office-specific formatting. The default python-pptx + reportlab pipeline produces a clean, text-first PDF that is always readable, regardless of whether Office fonts are installed.
