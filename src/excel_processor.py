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
    mp_codes: list[str],
    include_weighting: bool
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
    
    # Set row height for student data rows (all rows except header and potentially the last row)
    for row in range(2, last_row):
        ws.row_dimensions[row].height = STANDARD_ROW_HEIGHT
    
    # Set height for the last row (PONDERACIÓ) only if it's being included and needs the standard height
    if include_weighting and last_row >= 2: # Ensure last_row is at least 2 to avoid issues with very small sheets
        ws.row_dimensions[last_row].height = STANDARD_ROW_HEIGHT
    
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
    mp_codes: list[str],
    include_weighting: bool
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
                
                if include_weighting:
                    # Add red background rule for blank percentage cell in last row
                    cell_range_last_row_ra = f'{col}{last_row}'
                    blank_rule = FormulaRule(
                        formula=[f'ISBLANK({cell_range_last_row_ra})'],
                        stopIfTrue=True,
                        fill=red_fill
                    )
                    ws.conditional_formatting.add(cell_range_last_row_ra, blank_rule)
                    
                    # Add red text rule for 0% in last row
                    zero_percent_rule = FormulaRule(
                        formula=[f'{cell_range_last_row_ra}=0'],
                        stopIfTrue=True,
                        font=red_font
                    )
                    ws.conditional_formatting.add(cell_range_last_row_ra, zero_percent_rule)
    
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
                
                if include_weighting:
                    # Red text if not 90% in last row
                    cell_range_last_row_centre = f'{centre_col}{last_row}'
                    rule = CellIsRule(
                        operator='notEqual',
                        formula=['0.9'],  # 90%
                        stopIfTrue=True,
                        font=red_font
                    )
                    ws.conditional_formatting.add(cell_range_last_row_centre, rule)
        
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
            
            if include_weighting:
                # Red text if not 100% in last row
                cell_range_last_row_mp = f'{mp_col}{last_row}'
                rule = CellIsRule(
                    operator='notEqual',
                    formula=['1'],  # 100%
                    stopIfTrue=True,
                    font=red_font
                )
                ws.conditional_formatting.add(cell_range_last_row_mp, rule)
    
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
    mp_codes: list[str],
    include_weighting: bool
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
        include_weighting: Boolean to control inclusion of weighting logic
    """
    wb = load_workbook(workbook_path)
    ws = wb.active
    last_row = ws.max_row
    
    # Helper function to get column letter
    def get_column_for_header(header: str) -> str:
        # Ensure header is a string before attempting string operations
        header_str = str(header) if header is not None else ""
        for cell in ws[1]:  # First row
            if cell.value:
                # Ensure cell.value is a string before attempting string operations
                cell_value_str = str(cell.value) if cell.value is not None else ""
                # Replace both \n and \r (Windows newline)
                cell_value_clean = cell_value_str.replace('\n', '').replace('\r', '')
                header_clean = header_str.replace('\n', '').replace('\r', '')
                if cell_value_clean == header_clean:
                    return get_column_letter(cell.column)
        return None
    
    # For each MP, create appropriate formulas
    for mp_code in mp_codes:
        ra_codes = mp_groups.get(mp_code, []) # Use .get for safety
        ra_columns = [get_column_for_header(ra) for ra in ra_codes]
        ra_columns = [col for col in ra_columns if col is not None]
        
        if ra_columns:
            first_ra_col = ra_columns[0]
            last_ra_col = ra_columns[-1]
            
            failing_grades_check_parts = []
            for ra_col in ra_columns:
                failing_grades_check_parts.append(f'AND(ISNUMBER({ra_col}{{row}}),{ra_col}{{row}}<5)')
            any_failing = f'OR({",".join(failing_grades_check_parts)})'
            
            if mp_code in mp_codes_with_em:
                centre_column = get_column_for_header(f'{mp_code} CENTRE')
                empresa_column = get_column_for_header(f'{mp_code} EMPRESA')
                mp_column_em = get_column_for_header(mp_code) # Renamed to avoid conflict

                if centre_column:
                    for row_idx in range(2, last_row): # Student rows
                        cell = ws[f'{centre_column}{row_idx}']
                        if include_weighting:
                            cell.value = (
                                f'=IF(SUM({first_ra_col}{last_row}:{last_ra_col}{last_row})<>0.9,'
                                f'"ELS RA HAN DE SUMAR 90%",'
                                f'IF({any_failing.format(row=row_idx)},'
                                f'"NO SUPERAT",'
                                f'IF(COUNT({first_ra_col}{row_idx}:{last_ra_col}{row_idx})=COLUMNS({first_ra_col}{row_idx}:{last_ra_col}{row_idx}),'
                                f'SUMPRODUCT({first_ra_col}{row_idx}:{last_ra_col}{row_idx},'
                                f'{first_ra_col}{last_row}:{last_ra_col}{last_row}),'
                                f'"AVALUACIONS PENDENTS")))'
                            )
                        else:
                            cell.value = None
                    
                    if include_weighting and last_row >=1 : # Last row (weighting)
                        cell = ws[f'{centre_column}{last_row}']
                        cell.value = f'=SUM({first_ra_col}{last_row}:{last_ra_col}{last_row})'
                        cell.number_format = '0%'
                
                if empresa_column and include_weighting and last_row >=1: # Last row (weighting)
                    cell = ws[f'{empresa_column}{last_row}']
                    cell.value = 0.1  # 10%
                    cell.number_format = '0%'
                
                if mp_column_em and centre_column and empresa_column: # Ensure all columns found
                    for row_idx in range(2, last_row): # Student rows
                        cell = ws[f'{mp_column_em}{row_idx}']
                        if include_weighting:
                            cell.value = (
                                f'=IF(AND(NOT(ISTEXT({centre_column}{row_idx})),NOT(ISTEXT({empresa_column}{row_idx})),ISNUMBER({centre_column}{last_row}),ISNUMBER({empresa_column}{last_row})),' # Check if weights are numbers
                                f'ROUND(SUMPRODUCT('
                                f'CHOOSE({{1,2}},{centre_column}{row_idx},{empresa_column}{row_idx}),'
                                f'CHOOSE({{1,2}},{centre_column}{last_row},{empresa_column}{last_row})'
                                f'),0), "AVALUACIONS PENDENTS")'
                            )
                        else:
                            cell.value = None
                            
                    if include_weighting and last_row >=1: # Last row (weighting)
                        cell = ws[f'{mp_column_em}{last_row}']
                        cell.value = f'={centre_column}{last_row}+{empresa_column}{last_row}'
                        cell.number_format = '0%'
            else: # Regular MPs (Type B)
                mp_column_regular = get_column_for_header(mp_code) # Renamed to avoid conflict
                if mp_column_regular:
                    for row_idx in range(2, last_row): # Student rows
                        cell = ws[f'{mp_column_regular}{row_idx}']
                        if include_weighting:
                            cell.value = (
                                f'=IF(SUM({first_ra_col}{last_row}:{last_ra_col}{last_row})<>1,'
                                f'"ELS RA HAN DE SUMAR 100%",'
                                f'IF({any_failing.format(row=row_idx)},'
                                f'"NO SUPERAT",'
                                f'IF(COUNT({first_ra_col}{row_idx}:{last_ra_col}{row_idx})=COLUMNS({first_ra_col}{row_idx}:{last_ra_col}{row_idx}),'
                                f'ROUND(SUMPRODUCT({first_ra_col}{row_idx}:{last_ra_col}{row_idx},'
                                f'{first_ra_col}{last_row}:{last_ra_col}{last_row}),0),'
                                f'"AVALUACIONS PENDENTS")))'
                            )
                        else:
                            cell.value = None

                    if include_weighting and last_row >=1: # Last row (weighting)
                        cell = ws[f'{mp_column_regular}{last_row}']
                        cell.value = f'=SUM({first_ra_col}{last_row}:{last_ra_col}{last_row})'
                        cell.number_format = '0%'
    
    if include_weighting and last_row >=1:
        # Format all RA cells in the last row as percentages
        for mp_code_outer in mp_codes: 
            for ra_outer in mp_groups.get(mp_code_outer, []): # Use .get for safety
                col_outer = get_column_for_header(ra_outer)
                if col_outer:
                    cell = ws[f'{col_outer}{last_row}']
                    cell.number_format = '0%'
    
    wb.save(workbook_path)


def apply_cell_protection(
    workbook_path: str,
    mp_codes_with_em: list[str],
    mp_codes: list[str],
    mp_groups: dict[str, list[str]],
    include_weighting: bool
) -> None:
    """
    Apply cell protection to Excel file:
    - Lock all cells except RA percentage cells in the last row
    - Use password "edita'm" for protection
    - Allow editing only for RA percentage cells
    """
    if not include_weighting:
        return

    wb = load_workbook(workbook_path)
    ws = wb.active
    last_row = ws.max_row

    # Lock all cells first
    for row in ws.iter_rows():
        for cell in row:
            cell.protection = Protection(locked=True)

    # Helper function to find column letter for a header
    def get_column_for_header(header: str) -> str:
        header_str = str(header) if header is not None else ""
        for cell_obj in ws[1]:  # First row
            if cell_obj.value:
                cell_value_str = str(cell_obj.value) if cell_obj.value is not None else ""
                # Robust cleaning for cell value and header
                cell_value_clean = cell_value_str.replace('\n', '').replace('\r', '')
                header_clean = header_str.replace('\n', '').replace('\r', '')
                if cell_value_clean == header_clean:
                    return get_column_letter(cell_obj.column)
        return None

    # Unlock only RA grade columns in last row
    for mp_code in mp_codes:
        # Find RA columns for this MP
        for ra in mp_groups[mp_code]:
            col = get_column_for_header(ra)
            if col:
                cell_to_unlock = ws[f'{col}{last_row}']
                cell_to_unlock.protection = Protection(locked=False)

    # Protect the sheet with password
    ws.protection.sheet = True
    ws.protection.password = "edita'm"

    wb.save(workbook_path)


def export_excel_with_spacing(
    df: pd.DataFrame,
    output_path: str,
    mp_codes_with_em: list[str],
    mp_codes: list[str],
    include_weighting: bool
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
    
    # Group RA codes by their MP. Sort mp_codes by length descending to match longer MPs first (e.g., MP10 before MP1).
    # Use startswith(mp_code + '_') for more precise matching.
    mp_groups = {mp: [] for mp in mp_codes}
    sorted_mp_codes = sorted(mp_codes, key=len, reverse=True) # Sort MPs

    for ra_code in ra_codes:
        found_mp = False
        for mp_code in sorted_mp_codes:
            if ra_code.startswith(mp_code + '_'): # Precise matching
                mp_groups[mp_code].append(ra_code)
                found_mp = True
                break
        if not found_mp:
            # This case should ideally not happen if RAs are well-named (e.g. MPXX_RAX)
            # Consider logging a warning or handling as an error if an RA doesn't match any MP
            print(f"Warning: RA code '{ra_code}' did not match any known MP code prefix.")
    
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
    
    # Start with a copy of the original DataFrame's data for existing columns
    export_df = df.copy()

    # Add new placeholder columns (MP, CENTRE, EMPRESA) if they don't exist
    # These are the columns present in new_columns but not in original df
    for col_name in new_columns:
        if col_name not in export_df.columns:
            export_df[col_name] = pd.Series(dtype='float64', index=export_df.index) # Ensure correct index

    # Ensure all columns from new_columns are present and in the correct order
    # This also handles the case where some RA columns might not have been in the initial df
    # (though current logic implies df contains all ra_codes)
    export_df = export_df.reindex(columns=new_columns)
    
    
    # Initialize percentages for the last row (RA columns and EM columns)
    percentages = initialize_grade_weights(export_df, mp_groups, mp_codes_with_em)
    
    # Add "PONDERACIÓ (%}" row with initial percentages
    ponderacio_row = pd.DataFrame([percentages], columns=new_columns)
    ponderacio_row.iloc[0, 0] = 'PONDERACIÓ (%)'
    
    # export_df and ponderacio_row are already aligned with new_columns.
    # No explicit filtering here is needed as subsequent steps handle NaNs.
    
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
    apply_mp_sum_formulas(output_path, mp_groups, mp_codes_with_em, mp_codes, include_weighting)
    
    # Apply row formatting (borders, bold, background colors)
    apply_row_formatting(output_path, mp_codes_with_em, mp_codes, include_weighting)
    
    # Apply data validation for percentage cells
    apply_data_validation(output_path, mp_groups, mp_codes)
    
    # Apply conditional formatting (now without percentage validation)
    apply_conditional_formatting(output_path, mp_groups, mp_codes_with_em, mp_codes, include_weighting)
    
    #print(f"\t- Exported {len(export_df)-1} entries to {output_path}")
    
    # Apply cell protection
    apply_cell_protection(output_path, mp_codes_with_em, mp_codes, mp_groups, include_weighting)

    # If include_weighting is False, remove the last row (ponderacio row) from the Excel file
    if not include_weighting:
        wb = load_workbook(output_path)
        ws = wb.active
        if ws.max_row > 0: # Check if there are any rows to delete
            # Check if the first cell of the last row contains 'PONDERACIÓ (%)'
            # This is a safeguard to ensure we are deleting the correct row.
            if ws.cell(row=ws.max_row, column=1).value == 'PONDERACIÓ (%)':
                 ws.delete_rows(ws.max_row)
                 # print(f"\t- Removed 'PONDERACIÓ (%)' row as per include_weighting=False.") # Optional: uncomment for logging
            # else:
                # print(f"\t- Warning: Last row was not 'PONDERACIÓ (%)'. Did not remove last row.") # Optional: uncomment for logging
        wb.save(output_path)
    
    print(f"\t- Exported {len(export_df)-1} entries to {output_path}")
