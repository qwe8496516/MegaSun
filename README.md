# Excel File Merger Tool

A simple desktop application that can import two Excel files and merge them into a new Excel file.

## Features

- Graphical User Interface (GUI)
- Supports .xlsx and .xls formats
- Real-time progress display
- Error handling and prompt messages
- Multi-threaded processing to prevent interface freezing

## Installation Requirements

### 1. Install Python
Ensure your system has Python 3.12 or later installed.

### 2. Install Required Packages
Run the following command in the project directory:

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install pandas openpyxl xlrd
```

## Usage

### 1. Run the Program
```bash
python excel_merger.py
```

### 2. Operation Steps
1. Click "Browse" button to select the first Excel file
2. Click "Browse" button to select the second Excel file
3. Click "Browse" button to select the output file path
4. Click "Start Merge" button to begin processing
5. Wait for processing to complete

## Notes

- Current merge logic is simple vertical merge (combining data rows from two files)
- Ensure both input files have the same column structure
- Output files will overwrite existing files with the same name

## Future Extensions

This basic framework can be easily extended to support:
- More complex merge logic (such as merging based on specific columns)
- Data validation and cleaning
- Multi-file batch processing
- Custom column mapping
- Data transformation and calculations

## Technical Architecture

- **GUI Framework**: Tkinter (built-in Python)
- **Excel Processing**: pandas + openpyxl
- **Multi-threading**: threading
- **File Dialogs**: tkinter.filedialog