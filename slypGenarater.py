import pandas as pd
from datetime import datetime

def filter_payslip_items(items):
    """Filter out zero values and empty strings from payslip items"""
    filtered = []
    for line in items:
        # Skip if line is empty or just whitespace
        if not line.strip():
            continue
            
        # Check if line contains a value part (after tab)
        parts = line.split('\t')
        if len(parts) >= 2:
            value_part = parts[-1].strip()
            # Skip if value is "0.00" or "0" or empty or None
            try:
                float_val = float(value_part.replace(',', ''))
                if float_val == 0:
                    continue
            except (ValueError, AttributeError):
                if value_part in ['', 'None', 'nan'] or not value_part.strip():
                    continue
        filtered.append(line)
    return filtered

import math
import re

def filter_payslip_item(items):
    """
    Filters out formatted payslip strings if all numbers (like 0.00, NaN) are invalid.
    Keeps lines with at least one non-zero, non-NaN value.
    """
    filtered = []
    for line in items:
        # Extract all monetary values like 1,200.00, -45.00, etc.
        matches = re.findall(r"[-+]?\d[\d,]*\.\d{2}", line)
        # Convert to float
        values = []
        for val in matches:
            try:
                val_float = float(val.replace(",", ""))
                values.append(val_float)
            except:
                continue
        # Keep line if at least one number is valid (not 0.00 or NaN)
        if any(v != 0.0 and not math.isnan(v) for v in values):
            filtered.append(line)
    return filtered


def format_date_only(val):
    """Format a date value to dd/mm/yyyy, or return as string if not a date."""
    if pd.isnull(val):
        return ""
    try:
        if isinstance(val, str):
            dt = pd.to_datetime(val)
        elif isinstance(val, (pd.Timestamp, datetime)):
            dt = val
        else:
            return str(val)
        return dt.strftime('%d/%m/%Y')
    except Exception:
        return str(val)

def generate_fixed_payslip(row):
    """Generate a payslip for FIXED April sheet."""
    width = 100
    headerWidth= 80

    header = (
        f"{'COATS THREAD EXPORTS (PVT) LTD - OPERATOR EMPLOYEES'.center(headerWidth)}\n"
        f"{'PAY SLIP FOR THE MONTH OF APRIL 2025'.center(headerWidth)}\n"
    )

    # Updated emp_info_lines to handle integers for emp no and account no
    emp_info_lines = [
        f"{'EMP NO':<12}  {str(int(float(row['EMP NO ']))):<23}{'NIC NO':<14}  {row['NIC No.']}",
        f"{'NAME':<12}  {row['NAME']:<23}{'DEPARTMENT':<14}  {row['DEPARTMENT']}",
        f"{'DESIGNATION':<12}  {row['DESIGNATION']:<23}{'D.O.B':<14}  {format_date_only(row['DOB'])}",
        f"{'D.O.J':<12}  {format_date_only(row['DOJ']):<23}{'E.P.F.NO':<14}  {str(int(float(row['EPF  NO'])))}",
        f"{'':<12}  {'':<23}{'SAP NO':<14}  WF{str(int(float(row['EMP NO '])))}",
    ]

    emp_info = '\n'.join(emp_info_lines)



# Earnings/Deductions header
    ed_header = (
        f"\n\n{'EARNINGS'.ljust(headerWidth // 2)}{'DEDUCTIONS'.ljust(headerWidth // 2)}\n"
    )

    # List of earnings
    earnings = [
        f"{'BASIC SAL':<20}{row['BASIC SAL']:>12,.2f}",
        f"{'B.R ALLOWA':<20}{row['B.R ALLOWA']:>12,.2f}",
        f"{'MEDICAL':<20}{row['MEDICAL']:>12,.2f}",
        f"{'ACTING AL':<20}{row['ACTING AL2']:>12,.2f}",
        f"{'INCENTIVE':<20}{row['INCENTIVE']:>12,.2f}",
        f"{'SHIFT ALLO':<20}{row['SHIFT ALLO']:>12,.2f}",
        f"{'DISCRETIONARY INC':<20}{row['Dis cre- Fuel']:>12,.2f}",
        f"{'NORMAL OT':<20}{row['NORMAL OT']:>12,.2f}{row[row.index[row.index.get_loc('NORMAL OT') + 1]]:>5,.2f}",
        f"{'TRIPPLE OT':<20}{row['TRIPPLE OT']:>12,.2f}{row[row.index[row.index.get_loc('TRIPPLE OT') + 1]]:>5,.2f}",
        f"{'DOUBBLE OT':<20}{row['DOUBBLE OT']:>12,.2f}{row[row.index[row.index.get_loc('DOUBBLE OT') + 1]]:>5,.2f}",
        f"{'FIRST AID':<20}{row['FIRST AID']:>12,.2f}",
        f"{'FIRE TEAM':<20}{row['FIRE TEAM']:>12,.2f}",
        f"{'RELOCATION':<20}{row['RELOCATION']:>12,.2f}",
        f"{'SOSU ALLOW':<20}{row['SOSU ALLOW']:>12,.2f}",
        f"{'NO PAY COR':<20}{row['NO PAY COR']:>12,.2f}",
        f"{'FE NIG SHI':<20}{row['FE NIG SHI']:>12,.2f}",
        f"{'BALANCE LEAVE':<20}{row['BALANCE LEAVE']:>12,.2f}",
        f"{'SPEC SOSU':<20}{row['SPEC SOSU']:>12,.2f}",
        f"{'TAX REFUD':<20}{row['TAX REFUD']:>12,.2f}",
        f"{'Sunday Wages':<20}{row['Sunday Wages']:>12,.2f}",
        f"{'Arrears Double OT':<20}{row['Arrears Double OT']:>12,.2f}{row[row.index[row.index.get_loc('Arrears Double OT') + 1]]:>5,.2f}",
    ]

    # List of deductions
    deductions = [
        f"{'EPF YEE':15}{row['EPF YEE']:>12,.2f}",
        f"{'NO PAY':15}{row['NO PAY']:>12,.2f}",
        f"{'LATE MINUTE':15}{row['Late Minute']:>12,.2f}",
        f"{'WELFARE':15}{row['WELFARE']:>12,.2f}",
        f"{'SPORTS CLU':15}{row['SPORTS CLU']:>12,.2f}",
        f"{'FAIR FIRST':15}{row['FAIR FIRST']:>12,.2f}",
        f"{'UNION ICE':15}{row['UNION ICE']:>12,.2f}",
        f"{'FES ADVANC':15}{row['FES ADVANC']:>12,.2f}{row[row.index[row.index.get_loc('FES ADVANC') + 1]]:>12,.2f}",
        f"{'MOTOR CYCL':15}{row['MOTOR CYCL']:>12,.2f}{row[row.index[row.index.get_loc('MOTOR CYCL') + 1]]:>12,.2f}",
        f"{'MOTOR CINT':15}{row['MOTOR CINT']:>12,.2f}{row[row.index[row.index.get_loc('MOTOR CINT') + 1]]:>12,.2f}",
        f"{'P.L.D.C.Sampath':15}{row['P.L.D.C.Sampath']:>12,.2f}",
        f"{'APIT TAX':15}{row['APIT TAX']:>12,.2f}",
        f"{'WIJAYA RADIO':15}{row['WIJAYA RADIO']:>12,.2f}{row[row.index[row.index.get_loc('WIJAYA RADIO') + 1]]:>12,.2f}",
        f"{'EDUCATIONAL ASSISTANT':15}{row['EDUCATIONAL ASSISTANT']:>12,.2f}{row[row.index[row.index.get_loc('EDUCATIONAL ASSISTANT') + 1]]:>12,.2f}",
        f"{'MOCY GURANTER':15}{row['MOCY GURANTER']:>12,.2f}",
        f"{'MOCY GURANTER INT':15}{row['MOCY GURANTER INT']:>12,.2f}",
    ]


    # Filter out 0.00 or blank values
    earnings = filter_payslip_item(earnings)
    deductions = filter_payslip_item(deductions)

    # Format side-by-side layout
    combined_lines = combine_lines_fixed(earnings, deductions)

    # Footer aligned
    footer_lines = [
        f"{'TOT EARNINGS':<18}  {row['TOT EARN']:>12,.2f}       {' TOT DEDUCTIONS':<18}  {row['TOT DED']:>8,.2f}",
        "",
        f"{'EPF YEE 8%':<18}  {row['EPF YEE']:>12,.2f}       {' NET PAY':<16}  {row['netpay']:>10,.2f}",
        f"{'ETF YER 3%':<18}  {row['ETF YER']:>12,.2f}       {' BANK PAYMENT':<16}  {row['netpay']:>10,.2f}",
        f"{'EPF YER 12%':<18}  {row['EPF YER']:>12,.2f}",
        f"{'TOTAL EPF':<18}  {row['TOTAL EPF']:>12,.2f}",
        "",  # Blank line
        f"{'BANK':<10}  {str(row[14])}   {row['BRANCH NAME']}       {'A/C NO':<12}  {str(int(float(row['A/C NO'])))}",
    ]

    # Combine all parts
    payslip = header + emp_info + ed_header + '\n' + '\n'.join(combined_lines) + '\n\n' + '\n'.join(footer_lines)

    return payslip

def generate_ftc_payslip(row):
    """Generate a payslip for FTC April sheet."""
    width = 100
    headerWidth= 80
    
    header = (
        f"{'COATS THREAD EXPORTS (PVT) LTD - OPERATOR EMPLOYEES'.center(headerWidth)}\n"
        f"{'PAY SLIP FOR THE MONTH OF APRIL 2025'.center(headerWidth)}\n"
    )

    # Updated emp_info_lines to handle integers for emp no and account no
    emp_info_lines = [
        f"{'EMP NO':<15}  {str(int(float(row['EMP NO ']))):<24}{'NIC NO':<16}  {row['NIC No.']}",
        f"{'NAME':<15}  {row['NAME']:<24}{'DEPARTMENT':<16}  {row['DEPARTMENT']}",
        f"{'DESIGNATION':<15}  {row['DESIGNATION']:<24}{'D.O.B':<16}  {format_date_only(row['DOB'])}",
        f"{'D.O.J':<15}  {format_date_only(row['DOJ']):<24}{'E.P.F.NO':<16}  {str(int(float(row['EPF  NO'])))}",
        f"{'RATE':<15}  {int(row.get('BASIC SAL', 0)):<24}{'SAP NO':<16}  WF{str(int(float(row['EMP NO '])))}",
    ]

    emp_info = '\n'.join(emp_info_lines)

    ed_header = (
        f"\n\n{'EARNINGS'.ljust(headerWidth // 2)}{' DEDUCTIONS'.ljust(headerWidth // 2)}\n"
    )

    # Update earnings format to match fixed payslip
    earnings = [
        f"{'B.R ALLOWA':<20}{row.get('B.R ALLOWA', 0):>12,.2f}",
        f"{'BASIC SAL':<20}{row.get('BASIC SAL', 0):>12,.2f}",
        f"{'MEDICAL':<20}{row.get('MEDICAL', 0):>12,.2f}",
        f"{'ACTING AL':<20}{row.get('ACTING AL2', 0):>12,.2f}",
        f"{'INCENTIVE':<20}{row.get('INCENTIVE', 0):>12,.2f}",
        f"{'SHIFT ALLO':<20}{row.get('SHIFT ALLO', 0):>12,.2f}",
        f"{'NORMAL OT':<20}{row.get('NORMAL OT', 0):>12,.2f}{row[row.index[row.index.get_loc('NORMAL OT') + 1]]:>5,.2f}",
        f"{'TRIPPLE OT':<20}{row.get('TRIPPLE OT', 0):>12,.2f}{row[row.index[row.index.get_loc('TRIPLE OT') + 1]]:>5,.2f}",
        f"{'DOUBBLE OT':<20}{row.get('DOUBBLE OT', 0):>12,.2f}{row[row.index[row.index.get_loc('DOUBBLE OT') + 1]]:>5,.2f}",
        f"{'FIRST AID':<20}{row.get('FIRST AID', 0):>12,.2f}",
        f"{'FIRE TEAM':<20}{row.get('FIRE TEAM', 0):>12,.2f}",
        f"{'RELOCATION':<20}{row.get('RELOCATION', 0):>12,.2f}",
        f"{'SOSU ALLOW':<20}{row.get('SOSU ALLOW', 0):>12,.2f}",
        f"{'Sunday Wages':<20}{row.get('Sunday Wages', 0):>12,.2f}",
        f"{'Arrears Double OT':<20}{row.get('Arrears Double OT', 0):>12,.2f}{row[row.index[row.index.get_loc('Arrears Double OT') + 1]]:>5,.2f}",
    ]

    # Update deductions format to match fixed payslip
    deductions = [
        f"{'EPF YEE':15}{row.get('EPF YEE', 0):>12,.2f}",
        f"{'NO PAY':15}{row.get('NO PAY', 0):>12,.2f}",
        f"{'WELFARE':15}{row.get('WELFARE', 0):>12,.2f}",
        f"{'SPORTS CLU':15}{row.get('SPORTS CLU', 0):>12,.2f}",
        f"{'FAIR FIRST':15}{row.get('FAIR FIRST', 0):>12,.2f}",
        f"{'UNION ICE':15}{row.get('UNION ICE', 0):>12,.2f}",
        f"{'FES ADVANC':15}{row.get('FES ADVANC', 0):>12,.2f}",
        f"{'MOTOR CYCL':15}{row.get('MOTOR CYCL', 0):>12,.2f}",
        f"{'MOTOR CINT':15}{row.get('MOTOR CINT', 0):>12,.2f}",
        f"{'P.L.D.C.Sampath':15}{row.get('P.L.D.C.Sampath', 0):>12,.2f}",
        f"{'APIT TAX':15}{row.get('APIT TAX', 0):>12,.2f}",
        f"{'WIJAYA RADIO':15}{row.get('WIJAYA RADIO', 0):>12,.2f}",
        f"{'EDUCATIONAL ASSISTANT':15}{row.get('EDUCATIONAL ASSISTANT', 0):>12,.2f}",
        f"{'FOOT CYCLE LOAN':15}{row.get('FOOT CYCLE LOAN', 0):>12,.2f}",
        f"{'Singer':15}{row.get('Singer', 0):>12,.2f}",
    ]

    # Filter out 0.00 or blank values
    earnings = filter_payslip_item(earnings)
    deductions = filter_payslip_item(deductions)

    # Use the same combine_lines function as fixed payslip
    combined_lines = combine_lines_ftc(earnings, deductions)

    # Footer aligned like fixed payslip
    footer_lines = [
        f"{'TOT EARNINGS':<18}  {row.get('TOT EARN', 0):>12,.2f}         {'TOT DEDUCTIONS':<15}  {row.get('total deduction', 0):>10,.2f}",
        "",
        f"{'EPF YEE 8%':<18}  {row.get('EPF YEE', 0):>12,.2f}         {'NET PAY':<15}  {row.get('NETPAY', 0):>10,.2f}",
        f"{'ETF YER 3%':<18}  {row.get('ETF YER', 0):>12,.2f}         {'BANK PAYMENT':<15}  {row.get('NETPAY', 0):>10,.2f}",
        f"{'EPF YER 12%':<18}  {row.get('EPF YER', 0):>12,.2f}",
        f"{'TOTAL EPF':<18}  {row.get('TOTAL EPF', 0):>12,.2f}",
        "",  # Blank line
        f"{'BANK':<10}  {row.get('BANK CODE', ''):<15}  {row.get('BRANCH', ''):<10}  {'A/C NO':<12}  {str(int(float(row.get('A/C NO', 0)))):<15}",
    ]

    # Combine all parts
    payslip = header + emp_info + ed_header + '\n' + '\n'.join(combined_lines) + '\n\n' + '\n'.join(footer_lines)

    return payslip

def combine_lines_fixed(earnings, deductions, left_width=25, value_width=12, spacer=3):
    """Combine earnings and deductions into formatted lines."""
    combined_lines = []
    max_length = max(len(earnings), len(deductions))
    
    for i in range(max_length):
        # Get earnings and deductions entries, or empty strings if past the end
        left = earnings[i] if i < len(earnings) else ""
        right = deductions[i] if i < len(deductions) else ""
        
        # Each item is already formatted as "{description} {value}"
        # Just need to combine them with proper spacing
        left_space = " " * spacer
        line = f"{left:<{left_width + value_width}}{left_space}{right}"
        combined_lines.append(line)
    
    return combined_lines

def combine_lines_ftc(earnings, deductions, left_width=25, value_width=12, spacer=2):
    """Combine earnings and deductions into formatted lines."""
    combined_lines = []
    max_length = max(len(earnings), len(deductions))

    for i in range(max_length):
        # Get earnings and deductions entries, or empty strings if past the end
        left = earnings[i] if i < len(earnings) else ""
        right = deductions[i] if i < len(deductions) else ""

        # Each item is already formatted as "{description} {value}"
        # Just need to combine them with proper spacing
        left_space = "  " * spacer
        line = f"{left:<{left_width + value_width}}{left_space}{right}"
        combined_lines.append(line)

    return combined_lines

def generate_payslip(row, sheet_name=""):
    """
    Main payslip generator that routes to the appropriate function based on sheet name.
    """
    if "FIXED" in sheet_name.upper():
        return generate_fixed_payslip(row)
    elif "FTC" in sheet_name.upper():
        return generate_ftc_payslip(row)
    else:
        raise ValueError(f"Unsupported sheet type: {sheet_name}")


