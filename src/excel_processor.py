"""
Excel processing module for generating and formatting grade reports.
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers, PatternFill, Font, Border, Side, Alignment
from openpyxl.formatting.rule import CellIsRule, FormulaRule


def initialize_grade_weights(df: pd.DataFrame, mp_groups: dict[str, list[str]], mp_codes_with_em: list[str]) -> pd.Series:
    """
    Initialize weight values for RA columns and EM columns.
    
    Args:
        df: DataFrame containing the grade data
        mp_groups: Dictionary mapping MP codes to their RA codes
        mp_codes_with_em: List of MP codes that have EM entries
    
    Returns:
        Series with:
        - 0 for RA columns (will be formatted as percentages in Excel)
        - 10% for EM columns (MP code + EMPRESA)
        - Empty strings for other columns
    """
    percentages = pd.Series('', index=df.columns)
    
    # Set 0 for all RA columns (will be formatted as percentage in Excel)
    for mp_ras in mp_groups.values():
        for ra in mp_ras:
            percentages[ra] = 0
    
    # Set 10% for all EM columns
    for mp_code in mp_codes_with_em:
        em_column = f'{mp_code} EMPRESA'
        if em_column in percentages.index:
            percentages[em_column] = 0.1  # 10% in decimal form
            
    return percentages


def apply_row_formatting(
    workbook_path: str,
    mp_codes_with_em: list[str],
    mp_codes: list[str]
) -> None:
    """
    Apply formatting to the percentage row:
    - Add borders to all cells
    - Make the header cell bold
    - Make MP-related cells bold (CENTRE, EMPRESA, MP)
    - Set specific background colors for different types of cells
    - Set light turquoise background for headers and "PONDERACIÓ (%)" cell
    - Freeze first row and first column
    - Make student names bold
    - Center all cells vertically and horizontally (except first column which is left-aligned)
    - Set standard column widths based on column type
    - Format RA headers with line break after second underscore
    
    Args:
        workbook_path: Path to the Excel workbook
        mp_codes_with_em: List of MP codes that have EM entries
        mp_codes: List of all MP codes
    """
    wb = load_workbook(workbook_path)
    ws = wb.active
    last_row = ws.max_row
    
    # Freeze first row and first column
    ws.freeze_panes = 'B2'
    
    # Define styles
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    bold_font = Font(bold=True)
    mp_fill = PatternFill(start_color="E9D8FD", end_color="E9D8FD", fill_type="solid")
    ra_fill = PatternFill(start_color="C8DDFD", end_color="C8DDFD", fill_type="solid")
    header_fill = PatternFill(start_color="B8F2E6", end_color="B8F2E6", fill_type="solid")
    student_fill = PatternFill(start_color="D9F7F0", end_color="D9F7F0", fill_type="solid")  # 20% lighter than B8F2E6
    center_aligned = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_aligned = Alignment(horizontal='left', vertical='center')
    
    # Define standard column widths
    STUDENT_NAME_WIDTH = 40  # Width for student names column
    MP_COLUMN_WIDTH = 15     # Width for MP-related columns (CENTRE, EMPRESA, MP)
    RA_COLUMN_WIDTH = 12     # Width for RA grade columns (increased to fit headers)
    
    # Helper function to get column letter
    def get_column_for_header(header: str) -> str:
        for cell in ws[1]:  # First row
            if cell.value == header:
                return get_column_letter(cell.column)
        return None
    
    # Helper function to format RA header with line break
    def format_ra_header(header: str) -> str:
        if not header:
            return header
        
        # Count underscores to ensure we have at least two
        underscore_count = header.count('_')
        if underscore_count < 2:
            return header
            
        # Find the position of the second underscore
        first_pos = header.find('_')
        second_pos = header.find('_', first_pos + 1)
        
        # Insert line break after second underscore
        return f"{header[:second_pos + 1]}\n{header[second_pos + 1:]}"
    
    # Make first column header uppercase
    ws['A1'].value = str(ws['A1'].value).upper()
    
    # Apply formatting to all cells
    for row in ws.iter_rows(min_row=1, max_row=last_row):
        for cell in row:
            # Add borders to all cells
            cell.border = border
            
            if cell.column == 1:  # First column (student names and "PONDERACIÓ (%)")
                cell.alignment = left_aligned
                if cell.row < last_row:  # All rows except the last one (student names)
                    cell.font = bold_font
                    if cell.row > 1:  # Student names only (not header or last row)
                        cell.fill = student_fill
                if cell.row == 1 or cell.row == last_row:  # Header row and "PONDERACIÓ (%)" cell
                    cell.fill = header_fill
            else:  # All other columns
                cell.alignment = center_aligned
                if cell.row == 1:  # Header row
                    # Check if it's an RA header and format it
                    header_value = str(cell.value)
                    for mp_code in mp_codes:
                        if header_value.startswith(f'{mp_code}_'):
                            cell.value = format_ra_header(header_value)
                            break
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell.fill = header_fill
    
    # Apply formatting to the percentage row
    for cell in ws[last_row]:
        # Make the header cell bold (already left-aligned from above)
        if cell.column == 1:  # First column
            cell.font = bold_font
    
    # Apply background colors and bold text to MP-related cells
    for mp_code in mp_codes:
        if mp_code in mp_codes_with_em:
            # Color and bold MP CENTRE, EMPRESA, and MP columns
            for suffix in [' CENTRE', ' EMPRESA', '']:
                col = get_column_for_header(f'{mp_code}{suffix}')
                if col:
                    cell = ws[f'{col}{last_row}']
                    cell.fill = mp_fill
                    cell.font = bold_font
        else:
            # Color and bold MP column
            col = get_column_for_header(mp_code)
            if col:
                cell = ws[f'{col}{last_row}']
                cell.fill = mp_fill
                cell.font = bold_font
    
    # Apply light blue background to RA cells
    for mp_code in mp_codes:
        for ra in ws[1]:  # Check all columns
            if ra.value and str(ra.value).startswith(f'{mp_code}_'):
                col = get_column_letter(ra.column)
                cell = ws[f'{col}{last_row}']
                cell.fill = ra_fill
    
    # Set column widths based on column type
    # First, set width for student names column
    ws.column_dimensions['A'].width = STUDENT_NAME_WIDTH
    
    # Set widths for all other columns
    for cell in ws[1]:  # Iterate through header row
        if cell.column == 1:  # Skip first column (already handled)
            continue
            
        column = get_column_letter(cell.column)
        header_value = str(cell.value)
        
        # Check if it's an MP-related column
        is_mp_column = False
        for mp_code in mp_codes:
            if (mp_code in header_value and 
                (' CENTRE' in header_value or 
                 ' EMPRESA' in header_value or 
                 header_value == mp_code)):
                is_mp_column = True
                break
        
        # Set column width based on type
        if is_mp_column:
            ws.column_dimensions[column].width = MP_COLUMN_WIDTH
        else:  # RA columns
            ws.column_dimensions[column].width = RA_COLUMN_WIDTH
            # Adjust row height for header to ensure text wrapping works
            ws.row_dimensions[1].height = 30  # Increased height for header row
    
    wb.save(workbook_path)


def apply_conditional_formatting(
    workbook_path: str,
    mp_groups: dict[str, list[str]],
    mp_codes_with_em: list[str],
    mp_codes: list[str]
) -> None:
    """
    Apply conditional formatting rules to highlight:
    - Red text for RA columns if 0%
    - Red text for MP CENTRE columns if not 90%
    - Red text for MP columns if not 100%
    - Red background for any cell in the percentage row that is empty or not a percentage
    
    Args:
        workbook_path: Path to the Excel workbook
        mp_groups: Dictionary mapping MP codes to their RA codes
        mp_codes_with_em: List of MP codes that have EM entries
        mp_codes: List of all MP codes
    """
    wb = load_workbook(workbook_path)
    ws = wb.active
    last_row = ws.max_row
    
    # Helper function to get column letter
    def get_column_for_header(header: str) -> str:
        for cell in ws[1]:  # First row
            if cell.value == header:
                return get_column_letter(cell.column)
        return None
    
    red_font = Font(color="FF0000")  # Red color
    red_fill = PatternFill(start_color="FFD9D9", end_color="FFD9D9", fill_type="solid")  # Light red background
    
    # Get all columns except 'ESTUDIANT'
    all_columns = []
    for cell in ws[1]:  # First row
        if cell.column > 1:  # Skip the first column
            col_letter = get_column_letter(cell.column)
            all_columns.append(col_letter)
    
    # Add red background rule for empty or non-percentage cells in the last row
    for col in all_columns:  # Now this will skip the first column
        cell_range = f'{col}{last_row}'
        # Rule for empty cells
        empty_rule = CellIsRule(
            operator='equal',
            formula=['""'],
            stopIfTrue=True,
            fill=red_fill
        )
        ws.conditional_formatting.add(cell_range, empty_rule)
        
        # Rule for non-percentage cells (checks if the cell doesn't contain '%')
        non_percent_rule = FormulaRule(
            formula=[f'AND(NOT(ISBLANK({cell_range})), NOT(ISNUMBER({cell_range})))'],
            stopIfTrue=True,
            fill=red_fill
        )
        ws.conditional_formatting.add(cell_range, non_percent_rule)
    
    # Format RA columns - red text if 0%
    for mp_code in mp_codes:
        for ra in mp_groups[mp_code]:
            col = get_column_for_header(ra)
            if col:
                cell_range = f'{col}{last_row}'
                rule = CellIsRule(
                    operator='equal',
                    formula=['0'],
                    stopIfTrue=True,
                    font=red_font
                )
                ws.conditional_formatting.add(cell_range, rule)
    
    # Format MP CENTRE columns - red text if not 90%
    for mp_code in mp_codes_with_em:
        col = get_column_for_header(f'{mp_code} CENTRE')
        if col:
            cell_range = f'{col}{last_row}'
            rule = CellIsRule(
                operator='notEqual',
                formula=['0.9'],  # 90%
                stopIfTrue=True,
                font=red_font
            )
            ws.conditional_formatting.add(cell_range, rule)
    
    # Format MP columns - red text if not 100%
    for mp_code in mp_codes:
        col = get_column_for_header(mp_code)
        if col:
            cell_range = f'{col}{last_row}'
            rule = CellIsRule(
                operator='notEqual',
                formula=['1'],  # 100%
                stopIfTrue=True,
                font=red_font
            )
            ws.conditional_formatting.add(cell_range, rule)
    
    wb.save(workbook_path)


def apply_mp_sum_formulas(
    workbook_path: str,
    mp_groups: dict[str, list[str]],
    mp_codes_with_em: list[str],
    mp_codes: list[str]
) -> None:
    """
    Apply Excel formulas to calculate MP sums in the last row.
    For MPs with EM:
    - MP CENTRE column sums the RA percentages
    - MP EMPRESA column is initialized with 10%
    - MP column sums CENTRE and EMPRESA columns
    For other MPs:
    - MP column directly sums the RA percentages
    
    Args:
        workbook_path: Path to the Excel workbook
        mp_groups: Dictionary mapping MP codes to their RA codes
        mp_codes_with_em: List of MP codes that have EM entries
        mp_codes: List of all MP codes
    """
    # Load the workbook
    wb = load_workbook(workbook_path)
    ws = wb.active
    
    # Get the last row number (where the percentages are)
    last_row = ws.max_row
    
    # Helper function to convert column name to letter
    def get_column_for_header(header: str) -> str:
        for cell in ws[1]:  # First row
            if cell.value == header:
                return get_column_letter(cell.column)
        return None
    
    # For each MP, create appropriate formulas
    for mp_code in mp_codes:
        # Get RA codes for this MP
        ra_codes = mp_groups[mp_code]
        
        # Get the column letters for each RA
        ra_columns = [get_column_for_header(ra) for ra in ra_codes]
        ra_columns = [col for col in ra_columns if col is not None]
        
        if ra_columns:
            # Create the sum formula for RA percentages
            ra_sum_formula = f'=SUM({",".join(f"{col}{last_row}" for col in ra_columns)})'
            
            if mp_code in mp_codes_with_em:
                # For MPs with EM:
                # 1. CENTRE column gets the sum of RAs
                centre_column = get_column_for_header(f'{mp_code} CENTRE')
                if centre_column:
                    cell = ws[f'{centre_column}{last_row}']
                    cell.value = ra_sum_formula
                    cell.number_format = '0%'
                
                # 2. EMPRESA column is already set to 10% by initialize_grade_weights
                empresa_column = get_column_for_header(f'{mp_code} EMPRESA')
                if empresa_column:
                    cell = ws[f'{empresa_column}{last_row}']
                    cell.number_format = '0%'
                
                # 3. MP column sums CENTRE and EMPRESA
                mp_column = get_column_for_header(mp_code)
                if mp_column and centre_column and empresa_column:
                    cell = ws[f'{mp_column}{last_row}']
                    cell.value = f'=SUM({centre_column}{last_row},{empresa_column}{last_row})'
                    cell.number_format = '0%'
            else:
                # For regular MPs, just sum the RAs directly
                mp_column = get_column_for_header(mp_code)
                if mp_column:
                    cell = ws[f'{mp_column}{last_row}']
                    cell.value = ra_sum_formula
                    cell.number_format = '0%'
    
    # Format all RA cells and EM cells in the last row as percentages
    for mp_code in mp_codes:
        # Format RA cells
        for ra in mp_groups[mp_code]:
            col = get_column_for_header(ra)
            if col:
                cell = ws[f'{col}{last_row}']
                cell.number_format = '0%'
        
        # Format EM cell if applicable
        if mp_code in mp_codes_with_em:
            em_col = get_column_for_header(f'{mp_code} EMPRESA')
            if em_col:
                cell = ws[f'{em_col}{last_row}']
                cell.number_format = '0%'
    
    # Save the workbook
    wb.save(workbook_path)


def export_excel_with_spacing(
    df: pd.DataFrame,
    output_path: str,
    mp_codes_with_em: list[str],
    mp_codes: list[str]
) -> None:
    """
    Export DataFrame to Excel with specific column spacing after each MP's RAs.
    - 3 empty columns after MPs with EM (named: MP CENTRE, MP EMPRESA, MP)
    - 1 empty column after other MPs (named: MP)
    - Adds a "PONDERACIÓ (%)" row after student entries with percentage calculations
    - Uses Excel formulas to dynamically sum RA percentages for each MP
    - Applies conditional formatting for invalid percentages
    - Applies borders and background colors to the percentage row
    Uses pre-computed mp_codes list for efficient grouping.
    """
    # Get the current column order (excluding 'estudiant')
    ra_codes = [col for col in df.columns if col != 'estudiant']
    
    # Group RA codes by their MP using existing mp_codes
    mp_groups = {mp: [] for mp in mp_codes}  # Initialize with known MPs
    for ra_code in ra_codes:
        # Find which MP this RA belongs to
        mp_code = next(mp for mp in mp_codes if ra_code.startswith(mp + '_'))
        mp_groups[mp_code].append(ra_code)
    
    # Create new column order with spacing
    new_columns = ['estudiant']
    for mp_code in mp_codes:  # Use mp_codes to maintain consistent order
        # Add all RA codes for this MP
        new_columns.extend(mp_groups[mp_code])
        # Add empty columns with specific names based on whether MP has EM
        if mp_code in mp_codes_with_em:
            new_columns.extend([
                f'{mp_code} CENTRE',
                f'{mp_code} EMPRESA',
                f'{mp_code}'
            ])
        else:
            new_columns.append(f'{mp_code}')
    
    # Create new DataFrame with the desired column order
    export_df = pd.DataFrame(index=df.index)
    export_df['estudiant'] = df['estudiant']
    
    # Add RA columns with their values
    for col in ra_codes:
        export_df[col] = df[col]
    
    # Add empty spacing columns
    for col in new_columns:
        if col not in export_df.columns:
            export_df[col] = ''
    
    # Reorder columns
    export_df = export_df[new_columns]
    
    # Initialize percentages for the last row (RA columns and EM columns)
    percentages = initialize_grade_weights(export_df, mp_groups, mp_codes_with_em)
    
    # Add "PONDERACIÓ (%)" row with initial percentages
    ponderacio_row = pd.DataFrame([percentages], columns=new_columns)
    ponderacio_row.iloc[0, 0] = 'PONDERACIÓ (%)'
    export_df = pd.concat([export_df, ponderacio_row], ignore_index=True)
    
    # Export to Excel
    output_path = output_path.replace('.csv', '.xlsx')
    export_df.to_excel(output_path, index=False)
    
    # Apply Excel formulas for MP sums
    apply_mp_sum_formulas(output_path, mp_groups, mp_codes_with_em, mp_codes)
    
    # Apply row formatting (borders, bold, background colors)
    apply_row_formatting(output_path, mp_codes_with_em, mp_codes)
    
    # Apply conditional formatting
    apply_conditional_formatting(output_path, mp_groups, mp_codes_with_em, mp_codes)
    
    print(f"\t- Exported {len(export_df)-1} entries to {output_path}")


# Future Excel formatting functions can be added here:
# def apply_conditional_formatting(workbook: Workbook) -> None:
#     """Apply conditional formatting to grade cells."""
#     pass
#
# def protect_worksheet(worksheet: Worksheet) -> None:
#     """Protect specific ranges in the worksheet."""
#     pass
#
# def add_grade_formulas(worksheet: Worksheet) -> None:
#     """Add grade calculation formulas."""
#     pass 