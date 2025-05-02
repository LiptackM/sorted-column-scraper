# sorted-column-scraper
This Python script extracts and sorts data from a specified column in an Excel spreadsheet. It's particularly useful for processing Excel files where data is organized in columns and you want to isolate and sort values from a specific section.

## Features

- Load and parse `.xlsx` Excel files
- Extract a specified column (by header title)
- Slice data between a header and an optional end cell
- Sorts the sliced data
- CLI support for direct usage
- Full unit test suite using ` pytest -rs tests/test_excel_column_sorted.py`

## Requirements

- Python 3.8+
- `openpyxl`
- `pytest` (for testing)

Install requirements:

```bash
pip install -r requirements.txt
