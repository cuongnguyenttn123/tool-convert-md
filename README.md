# tool-convert-md

A Python tool to convert Word, Excel, or PDF files to Markdown, with image extraction and linking support.

## Features

- Convert a `.docx` file into a folder with the same base name, containing a `.md` file plus images.
  Supports headings, lists, tables (including nested tables, rendered as HTML), text boxes,
  headers/footers, footnotes/endnotes, hyperlinks (including internal bookmarks), and images.
- Convert an Excel file (`.xlsx`, `.xlsm`) into a folder with the same base name.
- Scan all sheets in an Excel workbook and export each sheet as a separate `.md` file.
- Convert a `.pdf` file into a folder with the same base name, containing a `.md` file plus images.
- Automatically extract images and insert Markdown image links.
- Keep output structure clean and easy to manage.

## Output Rules

### 1) DOCX

Input:

- `BaoCao.docx`

Output:

- `BaoCao/BaoCao.md`
- `BaoCao/images/` (created only if images are found)

### 2) Excel

Input:

- `DuLieu.xlsx`

Output:

- `DuLieu/`
- `DuLieu/<SheetName>.md` (one file per sheet)
- `DuLieu/images/` (created only if images are found)

### 3) PDF

Input:

- `TaiLieu.pdf`

Output:

- `TaiLieu/TaiLieu.md`
- `TaiLieu/images/` (created only if images are found)

## Requirements

- Python 3.9+

## Installation

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

### Convert a DOCX file

```bash
python3 convert_to_md.py /path/to/file.docx
```

### Convert an Excel file

```bash
python3 convert_to_md.py /path/to/file.xlsx
```

### Convert a PDF file

```bash
python3 convert_to_md.py /path/to/file.pdf
```

### Convert all supported files in a folder

```bash
python3 convert_to_md.py /path/to/folder
```

### Specify an output directory

```bash
python3 convert_to_md.py /path/to/file.xlsx -o /path/to/output
```

## Notes

- The tool skips temporary Office files with the `~$` prefix.
- For Excel, data is exported as Markdown tables based on the used cell range.
- Images (DOCX, Excel, PDF) are saved in each output's `images/` folder and linked with relative paths.
- PDF headings are inferred from font size; tables are detected automatically. Conversion quality
  depends on how the PDF was produced (scanned/image-only PDFs will have little to no extractable text).

## Main Files

- `convert_to_md.py`: conversion script
- `requirements.txt`: required dependencies
