"""
Excel processing module for generating and formatting grade reports.
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers, PatternFill, Font, Border, Side, Alignment, Protection
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation


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
    - Make "PONDERACIÓ (%)" cell bold
    - Set specific background colors for different types of cells
    - Set light turquoise background for headers and "PONDERACIÓ (%)" cell
    - Freeze first row and first column
    - Make student names bold
    - Center all cells vertically and horizontally (except first column which is left-aligned)
    - Set standard column widths based on column type
    - Format RA headers with line break after second underscore
    - Set appropriate row height to display multi-line text
    - Make all content in MP CENTRE, EMPRESA, and MP columns bold
    - Set 2 decimal places for numbers in MP CENTRE columns
    
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
    
    # Define number formats
    TWO_DECIMAL_FORMAT = '0.00'  # Format for CENTRE columns
    PERCENTAGE_FORMAT = '0%'     # Format for percentage cells
    
    # Define standard column widths
    STUDENT_NAME_WIDTH = 40  # Width for student names column
    MP_COLUMN_WIDTH = 15     # Width for MP-related columns (CENTRE, EMPRESA, MP)
    RA_COLUMN_WIDTH = 12     # Width for RA grade columns (increased to fit headers)
    
    # Define standard row height (in points) to accommodate "AVALUACIONS PENDENTS" in two lines
    STANDARD_ROW_HEIGHT = 30  # Increased height for all rows
    
    # Helper function to get column letter
    def get_column_for_header(header: str) -> str:
        for cell in ws[1]:  # First row
            if cell.value:  # Check if cell has a value
                # Remove any line breaks from both the cell value and the header for comparison
                cell_value = str(cell.value).replace('\n', '')
                header_clean = header.replace('\n', '')
                if cell_value == header_clean:
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
    
    # Set row height for all rows except the header (which was already set)
    for row in range(2, last_row + 1):
        ws.row_dimensions[row].height = STANDARD_ROW_HEIGHT
    
    # Set increased height for header row to accommodate wrapped RA headers
    ws.row_dimensions[1].height = 45  # Header row needs more height for wrapped RA codes
    
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
                elif cell.row == last_row:  # "PONDERACIÓ (%)" cell
                    cell.font = bold_font
                    cell.fill = header_fill
                if cell.row == 1:  # Header row
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
    
    # Apply formatting to MP-related columns
    for mp_code in mp_codes:
        if mp_code in mp_codes_with_em:
            # Format MP CENTRE column
            centre_col = get_column_for_header(f'{mp_code} CENTRE')
            if centre_col:
                # Apply to all cells in the column (except header)
                for row in range(2, last_row):  # Skip header and last row
                    cell = ws[f'{centre_col}{row}']
                    cell.font = bold_font
                    cell.number_format = TWO_DECIMAL_FORMAT
                
                # Format last row as percentage
                cell = ws[f'{centre_col}{last_row}']
                cell.font = bold_font
                cell.number_format = PERCENTAGE_FORMAT
                cell.fill = mp_fill
            
            # Format MP EMPRESA column
            empresa_col = get_column_for_header(f'{mp_code} EMPRESA')
            if empresa_col:
                # Apply to all cells in the column (except header)
                for row in range(2, last_row + 1):
                    cell = ws[f'{empresa_col}{row}']
                    cell.font = bold_font
                    if row == last_row:
                        cell.number_format = PERCENTAGE_FORMAT
                        cell.fill = mp_fill
            
            # Format MP column
            mp_col = get_column_for_header(mp_code)
            if mp_col:
                # Apply to all cells in the column (except header)
                for row in range(2, last_row + 1):
                    cell = ws[f'{mp_col}{row}']
                    cell.font = bold_font
                    if row == last_row:
                        cell.number_format = PERCENTAGE_FORMAT
                        cell.fill = mp_fill
        else:
            # Format MP column for Type B MPs
            mp_col = get_column_for_header(mp_code)
            if mp_col:
                # Apply to all cells in the column (except header)
                for row in range(2, last_row + 1):
                    cell = ws[f'{mp_col}{row}']
                    cell.font = bold_font
                    if row == last_row:
                        cell.number_format = PERCENTAGE_FORMAT
                        cell.fill = mp_fill
    
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
    
    wb.save(workbook_path)


def apply_conditional_formatting(
    workbook_path: str,
    mp_groups: dict[str, list[str]],
    mp_codes_with_em: list[str],
    mp_codes: list[str]
) -> None:
    """
    Apply conditional formatting rules to highlight:
    - For RA columns:
        * Red background if not a number
        * Red text if number is not in range 1-10
        * Red background in last row if percentage cell is blank
        * Red text in last row if percentage is 0%
    - Red text for percentage error messages in MP CENTRE and MP columns
    - Red text for "NO SUPERAT" message
    - Red text for "AVALUACIONS PENDENTS" in MP CENTRE and MP columns
    - Red text for MP CENTRE columns if not 90%
    - Red text for MP columns if not 100%
    
    Args:
        workbook_path: Path to the Excel workbook
        mp_groups: Dictionary mapping MP codes to their RA codes
        mp_codes_with_em: List of MP codes that have EM entries
        mp_codes: List of all MP codes
    """
    wb = load_workbook(workbook_path)
    ws = wb.active
    last_row = ws.max_row
    
    # Helper function to get column letter, handling line breaks in headers
    def get_column_for_header(header: str) -> str:
        for cell in ws[1]:  # First row
            if cell.value:  # Check if cell has a value
                # Remove any line breaks from both the cell value and the header for comparison
                cell_value = str(cell.value).replace('\n', '')
                header_clean = header.replace('\n', '')
                if cell_value == header_clean:
                    return get_column_letter(cell.column)
        return None
    
    red_font = Font(color="FF0000")  # Red color
    red_fill = PatternFill(start_color="FFD9D9", end_color="FFD9D9", fill_type="solid")  # Light red background
    
    # Format RA columns
    for mp_code in mp_codes:
        for ra in mp_groups[mp_code]:
            col = get_column_for_header(ra)
            if col:
                # Apply rules to student grade cells (all rows except header and last row)
                for row in range(2, last_row):  # Start from row 2 (skip header)
                    cell_range = f'{col}{row}'
                    
                    # Red background if not a number (including empty cells)
                    non_numeric_rule = FormulaRule(
                        formula=[f'OR(ISBLANK({cell_range}), NOT(ISNUMBER({cell_range})))'],
                        stopIfTrue=True,
                        fill=red_fill
                    )
                    ws.conditional_formatting.add(cell_range, non_numeric_rule)
                    
                    # Red text if number is not in range 1-10
                    out_of_range_rule = FormulaRule(
                        formula=[f'AND(ISNUMBER({cell_range}), OR({cell_range}<1, {cell_range}>10))'],
                        stopIfTrue=True,
                        font=red_font
                    )
                    ws.conditional_formatting.add(cell_range, out_of_range_rule)
                    
                    # Red text if number is less than 5
                    less_than_five_rule = FormulaRule(
                        formula=[f'AND(ISNUMBER({cell_range}), {cell_range}<5)'],
                        stopIfTrue=True,
                        font=red_font
                    )
                    ws.conditional_formatting.add(cell_range, less_than_five_rule)
                
                # Add red background rule for blank percentage cell in last row
                cell_range = f'{col}{last_row}'
                blank_rule = FormulaRule(
                    formula=[f'ISBLANK({cell_range})'],
                    stopIfTrue=True,
                    fill=red_fill
                )
                ws.conditional_formatting.add(cell_range, blank_rule)
                
                # Add red text rule for 0% in last row
                zero_percent_rule = FormulaRule(
                    formula=[f'{cell_range}=0'],
                    stopIfTrue=True,
                    font=red_font
                )
                ws.conditional_formatting.add(cell_range, zero_percent_rule)
    
    # Add red text rules for error messages in MP CENTRE and MP columns
    for mp_code in mp_codes:
        if mp_code in mp_codes_with_em:
            # For Type A MPs
            centre_col = get_column_for_header(f'{mp_code} CENTRE')
            if centre_col:
                # Red text for percentage error message, "NO SUPERAT", and "AVALUACIONS PENDENTS"
                for row in range(2, last_row):  # Skip header and percentage row
                    cell_range = f'{centre_col}{row}'
                    error_rule = FormulaRule(
                        formula=[f'OR({cell_range}="ELS RA HAN DE SUMAR 90%",{cell_range}="NO SUPERAT",{cell_range}="AVALUACIONS PENDENTS")'],
                        stopIfTrue=True,
                        font=red_font
                    )
                    ws.conditional_formatting.add(cell_range, error_rule)
                
                # Red text if not 90% in last row
                cell_range = f'{centre_col}{last_row}'
                rule = CellIsRule(
                    operator='notEqual',
                    formula=['0.9'],  # 90%
                    stopIfTrue=True,
                    font=red_font
                )
                ws.conditional_formatting.add(cell_range, rule)
        
        # For both Type A and B MPs
        mp_col = get_column_for_header(mp_code)
        if mp_col:
            # Red text for error messages
            for row in range(2, last_row):  # Skip header and percentage row
                cell_range = f'{mp_col}{row}'
                error_messages = []
                if mp_code in mp_codes_with_em:
                    error_messages.append('"AVALUACIONS PENDENTS"')
                else:
                    error_messages.extend(['"ELS RA HAN DE SUMAR 100%"', '"NO SUPERAT"', '"AVALUACIONS PENDENTS"'])
                
                error_rule = FormulaRule(
                    formula=[f'OR({",".join(f"{cell_range}={msg}" for msg in error_messages)})'],
                    stopIfTrue=True,
                    font=red_font
                )
                ws.conditional_formatting.add(cell_range, error_rule)
            
            # Red text if not 100% in last row
            cell_range = f'{mp_col}{last_row}'
            rule = CellIsRule(
                operator='notEqual',
                formula=['1'],  # 100%
                stopIfTrue=True,
                font=red_font
            )
            ws.conditional_formatting.add(cell_range, rule)
    
    wb.save(workbook_path)


def apply_data_validation(
    workbook_path: str,
    mp_groups: dict[str, list[str]],
    mp_codes: list[str]
) -> None:
    """
    Apply data validation rules:
    - For RA percentage cells in last row:
        * Must be a percentage between 0% and 100%
        * Cannot be left blank
        * Custom error messages in Catalan
    
    Args:
        workbook_path: Path to the Excel workbook
        mp_groups: Dictionary mapping MP codes to their RA codes
        mp_codes: List of all MP codes
    """
    from openpyxl.worksheet.datavalidation import DataValidation
    
    wb = load_workbook(workbook_path)
    ws = wb.active
    last_row = ws.max_row
    
    # Helper function to get column letter
    def get_column_for_header(header: str) -> str:
        for cell in ws[1]:  # First row
            if cell.value:
                cell_value = str(cell.value).replace('\n', '')
                header_clean = header.replace('\n', '')
                if cell_value == header_clean:
                    return get_column_letter(cell.column)
        return None
    
    # Create data validation for percentages
    dv = DataValidation(
        type="decimal",
        operator="between",
        formula1=0,  # 0%
        formula2=1,  # 100%
        allow_blank=False,
        showErrorMessage=True,
        errorTitle="Valor no vàlid",
        error="Si us plau, introduïu un percentatge.",
        promptTitle="Percentatge requerit",
        prompt="Introduïu un valor entre 0% i 100%"
    )
    ws.add_data_validation(dv)
    
    # Apply validation to RA percentage cells
    for mp_code in mp_codes:
        for ra in mp_groups[mp_code]:
            col = get_column_for_header(ra)
            if col:
                cell_range = f'{col}{last_row}'
                dv.add(cell_range)
    
    wb.save(workbook_path)


def apply_mp_sum_formulas(
    workbook_path: str,
    mp_groups: dict[str, list[str]],
    mp_codes_with_em: list[str],
    mp_codes: list[str]
) -> None:
    """
    Apply Excel formulas to calculate MP sums in the last row and SUMPRODUCT for student grades.
    For MPs with EM:
    - MP CENTRE column:
      * Checks if RA percentages sum to 90%, shows error if not
      * Shows "NO SUPERAT" if any RA grade is below 5
      * For students: checks if all values are valid numbers using COUNT
      * For percentage row: sums all RA percentages
    - MP EMPRESA column is set to 10%
    - MP column uses SUMPRODUCT between CENTRE and EMPRESA columns, rounded to integer
    For other MPs:
    - MP column:
      * Checks if RA percentages sum to 100%, shows error if not
      * Shows "NO SUPERAT" if any RA grade is below 5
      * Checks if all values are valid numbers using COUNT
      * Uses SUMPRODUCT for weighted calculation, rounded to integer
    
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
            if cell.value:
                cell_value = str(cell.value).replace('\n', '')
                header_clean = header.replace('\n', '')
                if cell_value == header_clean:
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
            # Get first and last RA columns for range references
            first_ra_col = ra_columns[0]
            last_ra_col = ra_columns[-1]
            
            # Create the failing grades check
            failing_grades_check = []
            for ra_col in ra_columns:
                failing_grades_check.append(f'AND(ISNUMBER({ra_col}{{row}}),{ra_col}{{row}}<5)')
            any_failing = f'OR({",".join(failing_grades_check)})'
            
            if mp_code in mp_codes_with_em:
                # For MPs with EM:
                # 1. CENTRE column gets the SUMPRODUCT formula with validation
                centre_column = get_column_for_header(f'{mp_code} CENTRE')
                if centre_column:
                    # Apply formula to each student row
                    for row in range(2, last_row):  # Skip header and percentage row
                        cell = ws[f'{centre_column}{row}']
                        # Create formula that checks percentages, failing grades, and valid numbers
                        cell.value = (
                            f'=IF(SUM({first_ra_col}{last_row}:{last_ra_col}{last_row})<>0.9,'
                            f'"ELS RA HAN DE SUMAR 90%",'
                            f'IF({any_failing.format(row=row)},'
                            f'"NO SUPERAT",'
                            f'IF(COUNT({first_ra_col}{row}:{last_ra_col}{row})=COLUMNS({first_ra_col}{row}:{last_ra_col}{row}),'
                            f'SUMPRODUCT({first_ra_col}{row}:{last_ra_col}{row},'
                            f'{first_ra_col}{last_row}:{last_ra_col}{last_row}),'
                            f'"AVALUACIONS PENDENTS")))'
                        )
                    
                    # Set percentage for CENTRE column (sum of RA percentages)
                    cell = ws[f'{centre_column}{last_row}']
                    cell.value = f'=SUM({first_ra_col}{last_row}:{last_ra_col}{last_row})'
                    cell.number_format = '0%'
                
                # 2. EMPRESA column is set to 10%
                empresa_column = get_column_for_header(f'{mp_code} EMPRESA')
                if empresa_column:
                    cell = ws[f'{empresa_column}{last_row}']
                    cell.value = 0.1  # 10%
                    cell.number_format = '0%'
                
                # 3. MP column uses SUMPRODUCT between CENTRE and EMPRESA columns, rounded to integer
                mp_column = get_column_for_header(mp_code)
                if mp_column and centre_column and empresa_column:
                    # For each student row
                    for row in range(2, last_row):  # Skip header and percentage row
                        cell = ws[f'{mp_column}{row}']
                        # Create SUMPRODUCT formula for CENTRE and EMPRESA, with rounding
                        cell.value = (
                            f'=IF(AND(NOT(ISTEXT({centre_column}{row})),NOT(ISTEXT({empresa_column}{row}))),'
                            f'ROUND(SUMPRODUCT('
                            f'CHOOSE({{1,2}},{centre_column}{row},{empresa_column}{row}),'
                            f'CHOOSE({{1,2}},{centre_column}{last_row},{empresa_column}{last_row})'
                            f'),0), "AVALUACIONS PENDENTS")'
                        )
                    
                    # Set percentage for MP column (should sum to 100%)
                    cell = ws[f'{mp_column}{last_row}']
                    cell.value = f'={centre_column}{last_row}+{empresa_column}{last_row}'
                    cell.number_format = '0%'
            else:
                # For regular MPs:
                # MP column gets the SUMPRODUCT formula with validation, rounded to integer
                mp_column = get_column_for_header(mp_code)
                if mp_column:
                    # Apply formula to each student row
                    for row in range(2, last_row):  # Skip header and percentage row
                        cell = ws[f'{mp_column}{row}']
                        # Create formula that checks percentages, failing grades, and valid numbers
                        cell.value = (
                            f'=IF(SUM({first_ra_col}{last_row}:{last_ra_col}{last_row})<>1,'
                            f'"ELS RA HAN DE SUMAR 100%",'
                            f'IF({any_failing.format(row=row)},'
                            f'"NO SUPERAT",'
                            f'IF(COUNT({first_ra_col}{row}:{last_ra_col}{row})=COLUMNS({first_ra_col}{row}:{last_ra_col}{row}),'
                            f'ROUND(SUMPRODUCT({first_ra_col}{row}:{last_ra_col}{row},'
                            f'{first_ra_col}{last_row}:{last_ra_col}{last_row}),0),'
                            f'"AVALUACIONS PENDENTS")))'
                        )
                    
                    # Set percentage for MP column (should sum to 100%)
                    cell = ws[f'{mp_column}{last_row}']
                    cell.value = f'=SUM({first_ra_col}{last_row}:{last_ra_col}{last_row})'
                    cell.number_format = '0%'
    
    # Format all RA cells in the last row as percentages
    for mp_code in mp_codes:
        for ra in mp_groups[mp_code]:
            col = get_column_for_header(ra)
            if col:
                cell = ws[f'{col}{last_row}']
                cell.number_format = '0%'
    
    wb.save(workbook_path)


def apply_cell_protection(
    workbook_path: str,
    mp_codes_with_em: list[str],
    mp_codes: list[str],
    mp_groups: dict[str, list[str]]
) -> None:
    """
    Apply cell protection to Excel file:
    - Lock all cells except RA percentage cells in the last row
    - Use password "edita'm" for protection
    - Allow editing only for RA percentage cells
    """
    wb = load_workbook(workbook_path)
    ws = wb.active
    last_row = ws.max_row
    
    # Lock all cells first
    for row in ws.iter_rows():
        for cell in row:
            cell.protection = Protection(locked=True)
    
    # Helper function to find column letter for a header
    def get_column_for_header(header: str) -> str:
        for cell in ws[1]:  # First row
            if cell.value:
                # Remove any line breaks from both the cell value and the header for comparison
                cell_value = str(cell.value).replace('\n', '')
                header_clean = header.replace('\n', '')
                if cell_value == header_clean:
                    return get_column_letter(cell.column)
        return None

    # Unlock only RA grade columns in last row
    for mp_code in mp_codes:
        # Find RA columns for this MP
        for ra in mp_groups[mp_code]:
            col = get_column_for_header(ra)
            if col:
                cell = ws[f'{col}{last_row}']
                cell.protection = Protection(locked=False)
    
    # Protect the sheet with password
    ws.protection.sheet = True
    ws.protection.password = "edita'm"
    
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
    - Adds a "PONDERACIÓ (%}" row after student entries with percentage calculations
    - Uses Excel formulas to dynamically sum RA percentages for each MP
    - Applies data validation for percentage cells
    - Applies conditional formatting for invalid grades
    - Protects all cells except RA percentage cells
    Uses pre-computed mp_codes list for efficient grouping.
    """
    # Get the current column order (excluding 'estudiant')
    ra_codes = [col for col in df.columns if col != 'estudiant']
    
    # Group RA codes by their MP using existing mp_codes
    mp_groups = {mp: [] for mp in mp_codes}  # Initialize with known MPs
    for ra_code in ra_codes:
        # Find which MP this RA belongs to
        for mp_code in mp_codes:
            if ra_code.startswith(mp_code):
                mp_groups[mp_code].append(ra_code)
                break
    
    # Add empty spacing columns with explicit numeric type for calculations
    new_columns = ['estudiant']  # Start with student name column
    for mp_code in mp_codes:
        # Add RA columns for this MP
        for ra in mp_groups[mp_code]:
            new_columns.append(ra)
        
        # Add spacing columns based on whether this MP has EM
        if mp_code in mp_codes_with_em:
            new_columns.extend([
                f'{mp_code} CENTRE',
                f'{mp_code} EMPRESA',
                f'{mp_code}'
            ])
        else:
            new_columns.append(f'{mp_code}')
    
    # Create DataFrame with proper column order
    export_df = pd.DataFrame(columns=new_columns)
    export_df['estudiant'] = df['estudiant']
    
    # Add RA columns with their values, preserving numeric types
    for col in ra_codes:
        export_df[col] = df[col]
    
    # Add empty spacing columns with explicit numeric type for calculations
    for col in new_columns:
        if col not in export_df.columns:
            # Create empty column with float64 type and fill with NaN
            export_df[col] = pd.Series(dtype='float64')
    
    # Reorder columns
    export_df = export_df[new_columns]
    
    # Initialize percentages for the last row (RA columns and EM columns)
    percentages = initialize_grade_weights(export_df, mp_groups, mp_codes_with_em)
    
    # Add "PONDERACIÓ (%}" row with initial percentages
    ponderacio_row = pd.DataFrame([percentages], columns=new_columns)
    ponderacio_row.iloc[0, 0] = 'PONDERACIÓ (%)'
    
    # Filter out empty or all-NA columns before concatenation
    non_empty_columns = [col for col in new_columns if not export_df[col].isna().all()]
    export_df = export_df[non_empty_columns]
    ponderacio_row = ponderacio_row[non_empty_columns]
    
    # Concatenate while preserving numeric types
    export_df = pd.concat([export_df, ponderacio_row], ignore_index=True)
    
    # Convert numeric columns to float64 explicitly
    numeric_cols = export_df.select_dtypes(include=['int64', 'float64']).columns
    for col in numeric_cols:
        # Convert to float64 with NaN for missing values
        export_df[col] = pd.to_numeric(export_df[col], errors='coerce').astype('float64')
    
    # Replace NaN with empty string only in non-numeric columns
    non_numeric_cols = export_df.columns.difference(numeric_cols)
    export_df[non_numeric_cols] = export_df[non_numeric_cols].fillna('')
    
    # Export to Excel
    output_path = output_path.replace('.csv', '.xlsx')
    export_df.to_excel(output_path, index=False)
    
    # Apply Excel formulas for MP sums
    apply_mp_sum_formulas(output_path, mp_groups, mp_codes_with_em, mp_codes)
    
    # Apply row formatting (borders, bold, background colors)
    apply_row_formatting(output_path, mp_codes_with_em, mp_codes)
    
    # Apply data validation for percentage cells
    apply_data_validation(output_path, mp_groups, mp_codes)
    
    # Apply conditional formatting (now without percentage validation)
    apply_conditional_formatting(output_path, mp_groups, mp_codes_with_em, mp_codes)
    
    print(f"\t- Exported {len(export_df)-1} entries to {output_path}")
    
    # Apply cell protection
    apply_cell_protection(output_path, mp_codes_with_em, mp_codes, mp_groups)
    """
    Export DataFrame to Excel with specific column spacing after each MP's RAs.
    - 3 empty columns after MPs with EM (named: MP CENTRE, MP EMPRESA, MP)
    - 1 empty column after other MPs (named: MP)
    - Adds a "PONDERACIÓ (%)" row after student entries with percentage calculations
    - Uses Excel formulas to dynamically sum RA percentages for each MP
    - Applies data validation for percentage cells
    - Applies conditional formatting for invalid grades
    - Protects all cells except RA percentage cells
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
    
    # Add RA columns with their values, preserving numeric types
    for col in ra_codes:
        export_df[col] = df[col]
    
    # Add empty spacing columns with explicit numeric type for calculations
    for col in new_columns:
        if col not in export_df.columns:
            # Create empty column with float64 type and fill with NaN
            export_df[col] = pd.Series(dtype='float64')
            
    # Reorder columns
    export_df = export_df[new_columns]
    
    # Initialize percentages for the last row (RA columns and EM columns)
    percentages = initialize_grade_weights(export_df, mp_groups, mp_codes_with_em)
    
    # Add "PONDERACIÓ (%)" row with initial percentages
    ponderacio_row = pd.DataFrame([percentages], columns=new_columns)
    ponderacio_row.iloc[0, 0] = 'PONDERACIÓ (%)'
    
    # Concatenate while preserving numeric types
    export_df = pd.concat([export_df, ponderacio_row], ignore_index=True)
    
    # Convert numeric columns to float64 explicitly
    numeric_cols = export_df.select_dtypes(include=['int64', 'float64']).columns
    for col in numeric_cols:
        # Convert to float64 with NaN for missing values
        export_df[col] = pd.to_numeric(export_df[col], errors='coerce').astype('float64')
    
    # Replace NaN with empty string only in non-numeric columns
    non_numeric_cols = export_df.columns.difference(numeric_cols)
    export_df[non_numeric_cols] = export_df[non_numeric_cols].fillna('')
    
    # Export to Excel
    output_path = output_path.replace('.csv', '.xlsx')
    export_df.to_excel(output_path, index=False)
    
    # Apply Excel formulas for MP sums
    apply_mp_sum_formulas(output_path, mp_groups, mp_codes_with_em, mp_codes)
    
    # Apply row formatting (borders, bold, background colors)
    apply_row_formatting(output_path, mp_codes_with_em, mp_codes)
    
    # Apply data validation for percentage cells
    apply_data_validation(output_path, mp_groups, mp_codes)
    
    # Apply conditional formatting (now without percentage validation)
    apply_conditional_formatting(output_path, mp_groups, mp_codes_with_em, mp_codes)
        
    # Apply cell protection
    apply_cell_protection(output_path, mp_codes_with_em, mp_codes, mp_groups)
