# COATS Payslip Generator - User Guide

## Table of Contents
1. [Getting Started](#getting-started)
2. [Loading Your Excel File](#loading-your-excel-file)
3. [Adding and Mapping New Fields](#adding-and-mapping-new-fields)
4. [Managing Existing Fields](#managing-existing-fields)
5. [Generating and Printing Payslips](#generating-and-printing-payslips)

## Getting Started

1. Double-click the COATS Payslip Generator application to launch it
2. The main window will open with these main areas:
   - Top: File selection button
   - Left: Sheet list
   - Right: Payslip preview
   - Bottom: Excel data table
   - Bottom buttons: Generate PDF, Print, and Configure options

## Loading Your Excel File

1. Click "Select Excel File" at the top of the window
2. Browse to your payroll Excel file
3. Select the appropriate sheet from the left panel
4. Your employee data will appear in the bottom table
5. Click any row to preview that employee's payslip

## Adding and Mapping New Fields

### Opening the Configuration

1. Click the "Configure Custom Fields" button at the bottom of the window
2. A new window will open showing current mappings

### Adding New Earnings or Deductions

1. In the configuration window:
   - Choose "FIXED" or "FTC" from the Sheet Type dropdown
   - Select "Earnings" or "Deductions" from Field Type
   - Type the Display Name (how it should appear on payslip)
   - Type the Excel Header (exact column name from Excel)
   - Click "Add/Update"

### Understanding Single vs Double Column Fields

#### Single Column Fields
- Use for simple values like Basic Salary
- Example:
  - Display Name: `BASIC SAL`
  - Excel Header: `BASIC SAL`

#### Double Column Fields (for Hours and Amount)
- Use for values that have hours and amounts
- Enter with comma separation
- Example for Normal OT:
  - Display Name: `NORMAL OT Hours, NORMAL OT Amount`
  - Excel Header: `NOT HRS, NOT AMT`

### Tips for Adding New Fields
- Excel Header must match EXACTLY with your Excel column name
- Check for extra spaces in Excel Headers
- Use proper case (uppercase/lowercase) to match Excel

1. Click on a sheet name in the left panel to load its data
2. The data table will show all entries
3. Click on any row to preview the payslip
4. Use the buttons at the bottom to:
   - Generate PDF for selected payslip
   - Print selected payslip
   - Generate PDFs for all payslips
   - Print all payslips

## Managing Existing Fields

### Viewing Current Mappings
1. Open Configure Custom Fields
2. Select Sheet Type (FIXED or FTC)
3. All current mappings will be shown in the list

### Editing a Field
1. Click on the mapping you want to edit from the list
2. The details will appear in the input fields
3. Make your changes
4. Click "Add/Update" to save changes

### Removing a Field
1. Click on the mapping you want to remove
2. Click the "Remove" button
3. Confirm the deletion

### Common Field Types

#### Earnings Fields Examples:
- Basic Salary
- Allowances (BR, Medical, etc.)
- Overtime (Normal, Double, Triple)
- Special Payments (First Aid, Fire Team)
- Incentives and Bonuses

#### Deductions Fields Examples:
- EPF
- Loans
- Insurance
- Union Fees
- Advance Payments

1. Select the mapping from the list
2. The fields will populate with existing values
3. Make your changes
4. Click "Add/Update" to save

### Removing Fields

1. Select the mapping you want to remove
2. Click "Remove" button

### Single vs Double Column Format

- Single Column: Enter one display name and one Excel header
- Double Column: Use comma to separate values
  Example: "OT Hours, OT Amount" for display names
  
### Understanding Field Types

1. **Earnings Fields:**
   - Basic Salary
   - Allowances
   - Overtime
   - Bonuses
   - Other payments

2. **Deductions Fields:**
   - EPF
   - Taxes
   - Loans
   - Other deductions

## Generating and Printing Payslips

### Generating Single PDF
1. Select an employee from the data table
2. Click "Generate Payslip & Save PDF"
3. Choose where to save the PDF
4. The PDF will be created with the employee's payslip

### Generating Multiple PDFs
1. Click "Generate All Payslips & PDFs"
2. Select the folder where you want to save PDFs
3. Wait for the progress bar to complete
4. All payslips will be saved as separate PDFs

### Printing Options

### Print Single Payslip
1. Select the employee from the table
2. Click "Print Selected Payslip"
3. In the print preview window:
   - Check the payslip layout
   - Adjust printer settings if needed
   - Click Print to send to printer

### Print Multiple Payslips
1. Click "Print All Payslips"
2. Select your printer
3. Choose print settings:
   - Paper size (A4 recommended)
   - Orientation (Portrait)
4. Click Print to start printing all payslips

### Quick Tips
- Always preview before printing
- Check printer has enough paper
- For bulk printing, test with one page first

### Common Issues and Solutions

1. **Excel File Not Loading:**
   - Verify file format (.xlsx, .xls)
   - Check file permissions
   - Ensure file isn't open in Excel

2. **Missing Fields in Payslip:**
   - Check configuration settings
   - Verify Excel column names match exactly
   - Check for extra spaces in names

3. **PDF Generation Errors:**
   - Verify write permissions in output folder
   - Check available disk space
   - Close any open PDFs

4. **Printing Issues:**
   - Verify printer connection
   - Check printer queue
   - Ensure correct printer driver

### Getting Help

If you encounter issues not covered in this documentation:
1. Check the application logs
2. Contact your system administrator
3. Report issues to the development team

## Best Practices

1. **Data Management:**
   - Keep regular backups of configuration
   - Use consistent Excel formats
   - Verify data before bulk operations

2. **Configuration:**
   - Document custom field mappings
   - Test new configurations with sample data
   - Keep field names consistent

3. **Operations:**
   - Preview before printing
   - Verify PDF output
   - Monitor disk space for bulk operations

---

For additional support or feature requests, please contact your system administrator or the development team.
