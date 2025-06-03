# PySlip - Payslip Generator and Printer

A Python GUI application for generating and printing payslips from Excel sheets.

## Features

- Load and view Excel sheets containing employee payroll data
- Generate payslips based on selected employee data
- Preview payslips before printing
- Print individual payslips with print preview
- Bulk generate and save payslips to files
- Bulk print all payslips from a sheet
- Cross-platform support for printing (Windows, macOS, Linux)

## Requirements

- Python 3.8 or later
- Dependencies listed in requirements.txt

## Installation

1. Clone the repository
2. Install the dependencies:
```
pip install -r requirements.txt
```

## Usage

Run the application:
```
python v1.py
```

### Steps to Use:

1. Click "Select Excel File" to load an Excel file containing payroll data
2. Select the sheet containing the data
3. The table will be populated with the sheet data
4. Select a row to preview an individual payslip
5. Use the buttons at the bottom to:
   - Generate Selected Payslip: Preview the selected payslip
   - Print Selected Payslip: Print the currently selected payslip
   - Generate All Payslips: Save all payslips as text files
   - Print All Payslips: Print all payslips in the current sheet

## File Structure

- `v1.py`: Main application file with the GUI
- `slypGenarater.py`: Contains functions for generating payslips
- `print_manager.py`: Handles printing functionality
- `requirements.txt`: Lists dependencies for the project

## Notes on Printing

- The application supports printing on Windows, macOS, and Linux
- Printer detection is automatic
- Print preview is available for individual payslips
- Bulk printing shows a progress dialog
