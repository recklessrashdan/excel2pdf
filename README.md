# excel2pdf

A simple Python utility to batch convert Excel files (.xlsx, .xls) to PDF format.

## Description

`excel2pdf` is a Windows-based automation tool that converts all Excel spreadsheets in a directory to PDF files. It uses Microsoft Excel's COM interface to perform high-quality conversions and maintains a detailed conversion log for tracking successes and failures.

## Features

- 🔄 **Batch Conversion**: Automatically converts all Excel files in the current directory
- 📝 **Detailed Logging**: Tracks all conversion attempts with timestamps
- ⚡ **Fast Processing**: Leverages Excel's native export functionality
- 🛡️ **Error Handling**: Gracefully handles conversion failures and logs them
- 🎯 **Simple to Use**: No complex configuration required

## Requirements

- **Python 3.x**
- **Microsoft Excel** (must be installed on your system)
- **pywin32**: Python library for Windows COM interface
  ```bash
  pip install pywin32
  ```

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/recklessrashdan/excel2pdf.git
   cd excel2pdf
   ```

2. Install required dependencies:
   ```bash
   pip install pywin32
   ```

## Usage

1. Place all Excel files (.xlsx, .xls) that you want to convert in the same directory as `converter.py`

2. Run the script:
   ```bash
   python converter.py
   ```

3. The script will:
   - Convert each Excel file to PDF format
   - Save PDFs with the same filename as the original (with .pdf extension)
   - Display conversion status in the console
   - Log all activities to `conversion_log.txt`

## Example

If your directory contains:
- `data.xlsx`
- `report.xls`
- `budget.xlsx`

Running `converter.py` will generate:
- `data.pdf`
- `report.pdf`
- `budget.pdf`

And create a `conversion_log.txt` with the conversion history.

## Output Log

The `conversion_log.txt` file tracks:
- Session start time
- Successful conversions: `SUCCESS: filename.xlsx -> filename.pdf`
- Failed conversions: `FAILED: filename.xlsx -> error details`

Example log:
```
===== New Session =====
Time: 2026-07-05 14:30:45.123456

SUCCESS: report.xlsx -> report.pdf
SUCCESS: data.xls -> data.pdf
FAILED: budget.xlsx -> [Excel error details]
```

## Notes

- ⚠️ **Windows Only**: This tool requires Microsoft Excel and Windows COM interface
- Excel application runs in the background (not visible)
- Original Excel files are not modified
- PDFs are saved in the same directory as the source files

## Troubleshooting

**"Excel is not installed"**: Ensure Microsoft Excel is properly installed on your system.

**"pywin32 module not found"**: Run `pip install pywin32` to install the required module.

**Files not converting**: Check `conversion_log.txt` for specific error messages.

## License

This project is open source and available for personal and commercial use.

## Author

[recklessrashdan](https://github.com/recklessrashdan)
