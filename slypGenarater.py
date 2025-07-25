import pandas as pd
import math
import re
from datetime import datetime
from config import PayslipConfig  
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
        if not line or not line.strip():
            continue
            
        # Extract all monetary values like 1,200.00, -45.00, 0.00, etc.
        # Updated regex to better match monetary formats
        matches = re.findall(r"[-+]?\d{1,3}(?:,\d{3})*\.\d{2}", line)
        
        # Convert to float and check for valid values
        has_valid_value = False
        for val in matches:
            try:
                val_float = float(val.replace(",", ""))
                if val_float != 0.0 and not math.isnan(val_float):
                    has_valid_value = True
                    break
            except (ValueError, TypeError):
                continue
        
        # Keep line if it has at least one valid (non-zero) value
        if has_valid_value:
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

def add_custom_entries(earnings, deductions, row, sheet_type):
    """Add custom mapped entries to earnings and deductions (single or double column, only one display name)"""
    config = PayslipConfig()
    mappings = config.get_mappings(sheet_type)

    def get_val(col):
        return row.get(col, 0) if col in row else 0

    def safe_fmt(val, width=12):
        try:
            fval = float(val)
            return f"{fval:>{width},.2f}"
        except Exception:
            return f"{str(val):>{width}}"
    
    def is_valid_value(val):
        """Check if value is valid (not zero, null, or NaN)"""
        if pd.isnull(val):
            return False
        try:
            float_val = float(val)
            return float_val != 0.0 and not math.isnan(float_val)
        except (ValueError, TypeError):
            return False

    # Add custom earnings
    for display_name, excel_header in mappings['earnings'].items():
        if isinstance(display_name, str) and display_name.startswith('[') and display_name.endswith(']'):
            try:
                import ast
                display_name = ast.literal_eval(display_name)
            except Exception:
                pass
        if isinstance(excel_header, (list, tuple)) and len(excel_header) == 2:
            col1, col2 = excel_header
            if col1 in row and col2 in row:
                v1, v2 = get_val(col1), get_val(col2)
                # Only add if at least one value is valid (non-zero)
                if is_valid_value(v1) or is_valid_value(v2):
                    earnings.append(f"{display_name:<20}{safe_fmt(v1,12)}{safe_fmt(v2,5)}")
        elif isinstance(excel_header, str):
            if excel_header in row:
                v = get_val(excel_header)
                # Only add if value is valid (non-zero)
                if is_valid_value(v):
                    earnings.append(f"{display_name:<20}{safe_fmt(v,12)}")

    # Add custom deductions
    for display_name, excel_header in mappings['deductions'].items():
        if isinstance(display_name, str) and display_name.startswith('[') and display_name.endswith(']'):
            try:
                import ast
                display_name = ast.literal_eval(display_name)
            except Exception:
                pass
        if isinstance(excel_header, (list, tuple)) and len(excel_header) == 2:
            col1, col2 = excel_header
            if col1 in row and col2 in row:
                v1, v2 = get_val(col1), get_val(col2)
                # Only add if at least one value is valid (non-zero)
                if is_valid_value(v1) or is_valid_value(v2):
                    deductions.append(f"{display_name:15}{safe_fmt(v1,12)}{safe_fmt(v2,5)}")
        elif isinstance(excel_header, str):
            if excel_header in row:
                v = get_val(excel_header)
                # Only add if value is valid (non-zero)
                if is_valid_value(v):
                    deductions.append(f"{display_name:15}{safe_fmt(v,12)}")

def get_payslip_month_year():
    """Return the payslip month and year string, e.g., 'MAY 2025'."""
    now = datetime.now()
    month_str = now.strftime('%B').upper()
    year_str = now.strftime('%Y')
    return f"{month_str} {year_str}"

def generate_earnings_from_config(row, sheet_type):
    """Generate earnings list from configuration"""
    config = PayslipConfig()
    mappings = config.get_mappings(sheet_type)
    earnings = []
    
    def is_valid_value(val):
        """Check if value is valid (not zero, null, or NaN)"""
        if pd.isnull(val):
            return False
        try:
            float_val = float(val)
            return float_val != 0.0 and not math.isnan(float_val)
        except (ValueError, TypeError):
            return False
    
    for display_name, excel_header in mappings['earnings'].items():
        if isinstance(excel_header, list) and len(excel_header) == 2:
            col1, col2 = excel_header
            if col1 in row:
                v1 = row.get(col1, 0)
                if col2 == "next_column":
                    # Get the next column after col1
                    try:
                        col1_idx = row.index.get_loc(col1)
                        v2 = row.iloc[col1_idx + 1] if col1_idx + 1 < len(row) else 0
                    except:
                        v2 = 0
                elif col2 == "prev_column":
                    # Get the previous column before col1
                    try:
                        col1_idx = row.index.get_loc(col1)
                        v2 = row.iloc[col1_idx - 1] if col1_idx - 1 >= 0 else 0
                    except:
                        v2 = 0
                else:
                    v2 = row.get(col2, 0)
                
                # Only add if at least one value is valid (non-zero)
                if is_valid_value(v1) or is_valid_value(v2):
                    earnings.append(f"{display_name:<20}{v1:>12,.2f}{v2:>5,.2f}")
        else:
            # Single column
            if excel_header in row:
                v = row.get(excel_header, 0)
                # Only add if value is valid (non-zero)
                if is_valid_value(v):
                    earnings.append(f"{display_name:<20}{v:>12,.2f}")
    
    return earnings

def generate_deductions_from_config(row, sheet_type):
    """Generate deductions list from configuration"""
    config = PayslipConfig()
    mappings = config.get_mappings(sheet_type)
    deductions = []
    
    def is_valid_value(val):
        """Check if value is valid (not zero, null, or NaN)"""
        if pd.isnull(val):
            return False
        try:
            float_val = float(val)
            return float_val != 0.0 and not math.isnan(float_val)
        except (ValueError, TypeError):
            return False
    
    for display_name, excel_header in mappings['deductions'].items():
        if isinstance(excel_header, list) and len(excel_header) == 2:
            col1, col2 = excel_header
            if col1 in row:
                v1 = row.get(col1, 0)
                if col2 == "next_column":
                    # Get the next column after col1
                    try:
                        col1_idx = row.index.get_loc(col1)
                        v2 = row.iloc[col1_idx + 1] if col1_idx + 1 < len(row) else 0
                    except:
                        v2 = 0
                elif col2 == "prev_column":
                    # Get the previous column before col1
                    try:
                        col1_idx = row.index.get_loc(col1)
                        v2 = row.iloc[col1_idx - 1] if col1_idx - 1 >= 0 else 0
                    except:
                        v2 = 0
                else:
                    v2 = row.get(col2, 0)
                
                # Only add if at least one value is valid (non-zero)
                if is_valid_value(v1) or is_valid_value(v2):
                    deductions.append(f"{display_name:15}{v1:>12,.2f}{v2:>12,.2f}")
        else:
            # Single column
            if excel_header in row:
                v = row.get(excel_header, 0)
                # Only add if value is valid (non-zero)
                if is_valid_value(v):
                    deductions.append(f"{display_name:15}{v:>12,.2f}")
    
    return deductions

def generate_fixed_payslip(row):
    """Generate a payslip for FIXED April sheet."""
    width = 100
    headerWidth= 80

    payslip_month = get_payslip_month_year()
    header = (
        f"{'COATS THREAD EXPORTS (PVT) LTD - OPERATOR EMPLOYEES'.center(headerWidth)}\n"
        f"{('PAY SLIP FOR THE MONTH OF ' + payslip_month).center(headerWidth)}\n"
    )

    # Updated emp_info_lines to handle integers for emp no and account no
    emp_info_lines = [
        f"{'EMP NO':<12}  {str(int(float(row['EMP NO ']))):<26}{'NIC NO':<14}  {row['NIC No.']}",
        f"{'NAME':<12}  {row['NAME']:<26}{'DEPARTMENT':<14}  {row['DEPARTMENT']}",
        f"{'DESIGNATION':<12}  {row['DESIGNATION']:<26}{'D.O.B':<14}  {format_date_only(row['DOB'])}",
        f"{'D.O.J':<12}  {format_date_only(row['DOJ']):<26}{'E.P.F.NO':<14}  {str(int(float(row['EPF  NO'])))}",
        f"{'':<12}  {'':<26}{'SAP NO':<14}  {str(row['REF NO'])}",    ]

    emp_info = '\n'.join(emp_info_lines)



    # Earnings/Deductions header
    ed_header = (
        f"\n\n{'EARNINGS'.ljust(headerWidth // 2)}{'DEDUCTIONS'.ljust(headerWidth // 2)}\n"
    )

    # Generate earnings and deductions from config instead of hardcoded lists
    earnings = generate_earnings_from_config(row, 'FIXED')
    deductions = generate_deductions_from_config(row, 'FIXED')

    # Add custom mapped entries (now includes filtering at source)
    add_custom_entries(earnings, deductions, row, 'FIXED')
    
    # Additional filtering as safety net (should be minimal now)
    earnings = filter_payslip_item(earnings)
    deductions = filter_payslip_item(deductions)
    # Format side-by-side layout
    combined_lines = combine_lines_fixed(earnings, deductions)

    # Footer aligned
    footer_lines = [
        f"{'TOT EARNINGS':<18}  {row['TOT EARN']:>11,.2f}       {' TOT DEDUCTIONS':<18}  {row['TOT DED']:>9,.2f}",        "",
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

    payslip_month = get_payslip_month_year()
    header = (
        f"{'COATS THREAD EXPORTS (PVT) LTD - FTC EMPLOYEES'.center(headerWidth)}\n"
        f"{('PAY SLIP FOR THE MONTH OF ' + payslip_month).center(headerWidth)}\n"
    )

    # Updated emp_info_lines to handle integers for emp no and account no
    emp_info_lines = [
        f"{'EMP NO':<15}  {str(int(float(row['EMP NO ']))):<24}{'NIC NO':<16}  {row['NIC No.']}",
        f"{'NAME':<15}  {row['NAME']:<24}{'DEPARTMENT':<16}  {row['DEPARTMENT']}",
        f"{'DESIGNATION':<15}  {row['DESIGNATION']:<24}{'D.O.B':<16}  {format_date_only(row['DOB'])}",
        f"{'D.O.J':<15}  {format_date_only(row['DOJ']):<24}{'E.P.F.NO':<16}  {str(int(float(row['EPF  NO'])))}",
        f"{'RATE':<15}  {str(int(float(row['RATE']))):<24}",    ]

    emp_info = '\n'.join(emp_info_lines)

    ed_header = (
        f"\n\n{'EARNINGS'.ljust(headerWidth // 2)}{' DEDUCTIONS'.ljust(headerWidth // 2)}\n"
    )

    # Generate earnings and deductions from config instead of hardcoded lists
    earnings = generate_earnings_from_config(row, 'FTC')
    deductions = generate_deductions_from_config(row, 'FTC')

    # Add custom mapped entries (now includes filtering at source)
    add_custom_entries(earnings, deductions, row, 'FTC')
    
    # Additional filtering as safety net (should be minimal now)
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
        f"{'TOTAL EPF':<18}  {row.get('TOTAL EPF', 0):>12,.2f}         {'NO OF DAYS WORKED':<15}  {row.get('NO OF DAYS WORKED', 0):>8,.2f}",
        "",  # Blank line
        f"{'BANK':<10}  {row.get('BANK CODE', ''):<15}  {row.get('BRANCH', ''):<10}  {'A/C NO':<15}  {str(int(float(row.get('A/C NO', 0)))):<15}",
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