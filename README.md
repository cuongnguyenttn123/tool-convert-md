# tool-convert-md

A Python tool to convert Word/Excel files to Markdown, with image extraction and linking support.

## Features

- Convert a `.docx` file into a `.md` file with the same base name.
- Convert an Excel file (`.xlsx`, `.xlsm`) into a folder with the same base name.
- Scan all sheets in an Excel workbook and export each sheet as a separate `.md` file.
- Automatically extract images and insert Markdown image links.
- Keep output structure clean and easy to manage.

## Output Rules

### 1) DOCX

Input:

- `BaoCao.docx`

Output:

- `BaoCao.md`
- `BaoCao_images/` (created only if images are found)

### 2) Excel

Input:

- `DuLieu.xlsx`

Output:

- `DuLieu/`
- `DuLieu/<SheetName>.md` (one file per sheet)
- `DuLieu/images/` (created only if images are found)

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
- Excel images are saved in the `images` folder and linked with relative paths in each sheet Markdown file.

## Main Files

- `convert_to_md.py`: conversion script
- `requirements.txt`: required dependencies
