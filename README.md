# Marksheet Generator

A lightweight Python utility to generate student marksheets from Excel ledger files.

## Features

- Extract student data from structured Excel ledgers
- Generate Word (.docx) marksheets for First Term, Second Term, and Final/Annual reports
- Automatic grade and GPA calculations
- Include school logo and student details
- GUI interface for easy operation

## Requirements

- Python 3.10+
- `openpyxl`
- `python-docx`
- `tkinter`

```bash
pip install openpyxl python-docx
```

## Usage

```bash
python ui.py
```

1. Select your Excel ledger file
2. Fill in student details
3. Click "Generate Marksheets"

## Project Structure

```
.
├── config.py
├── data_mapper.py
├── doc_builder
│   ├── common_functions.py
│   ├── doc_builder_final.py
│   └── doc_builder_terminal.py
├── excel_reader.py
├── README.md
├── requirements.txt
├── resources
│   ├── image.png
│   └── ledger.xlsx
├── ui.py
└── utils.py
```

## Notes

- Personal use tool with no warranty or official support
- Test data included in `resources/ledger.xlsx`
- Designed for single-user operation but easily customizable
