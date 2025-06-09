"""
Excel processing module for generating and formatting grade reports.
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.formatting.rule import FormulaRule


def apply_row_formatting(
    workbook_path: str,
    mp_codes_with_em: list[str],
    mp_codes: list[str],
) -> None:
    """
    Apply formatting to the percentage row:
    - Add borders to all cells
    - Make the header cell bold
    - Make MP-related cells bold (CENTRE, EMPRESA, MP)
    - Set specific background colors for different types of cells
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
    
    # Set row height for student data rows (all rows except header)
    for row in range(2, last_row+1):
        ws.row_dimensions[row].height = STANDARD_ROW_HEIGHT
    
    # Set increased height for header row to accommodate wrapped RA headers
    ws.row_dimensions[1].height = 45  # Header row needs more height for wrapped RA codes
    
    # Apply formatting to all cells (only up to last_row, not last_row+1)
    for row in ws.iter_rows(min_row=1, max_row=last_row):
        for cell in row:
            # Add borders to all cells
            cell.border = border
            
            if cell.column == 1:  # First column (student names)
                cell.alignment = left_aligned
                cell.font = bold_font
                if cell.row > 1:  # Student names only (not header)
                    cell.fill = student_fill
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
                for row in range(2, last_row+1):  # Skip header
                    cell = ws[f'{centre_col}{row}']
                    cell.font = bold_font
                    cell.number_format = TWO_DECIMAL_FORMAT
            
            # Format MP EMPRESA column
            empresa_col = get_column_for_header(f'{mp_code} EMPRESA')
            if empresa_col:
                # Apply to all cells in the column (except header)
                for row in range(2, last_row+1):
                    cell = ws[f'{empresa_col}{row}']
                    cell.font = bold_font
            
            # Format MP column
            mp_col = get_column_for_header(mp_code)
            if mp_col:
                # Apply to all cells in the column (except header)
                for row in range(2, last_row+1):
                    cell = ws[f'{mp_col}{row}']
                    cell.font = bold_font
        else:
            # Format MP column for Type B MPs
            mp_col = get_column_for_header(mp_code)
            if mp_col:
                # Apply to all cells in the column (except header)
                for row in range(2, last_row+1):
                    cell = ws[f'{mp_col}{row}']
                    cell.font = bold_font
    
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
                for row in range(2, last_row+1):  # Start from row 2 (skip header)
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
    
    wb.save(workbook_path)


# def apply_mp_sum_formulas(
#     workbook_path: str,
#     mp_groups: dict[str, list[str]],
#     mp_codes_with_em: list[str],
#     mp_codes: list[str]
# ) -> None:
#     """
#     Apply Excel formulas to calculate MP sums in the last row and SUMPRODUCT for student grades.
#     For MPs with EM:
#     - MP CENTRE column:
#       * Checks if RA percentages sum to 90%, shows error if not
#       * Shows "NO SUPERAT" if any RA grade is below 5
#       * For students: checks if all values are valid numbers using COUNT
#       * For percentage row: sums all RA percentages    
#     Args:
#         workbook_path: Path to the Excel workbook
#         mp_groups: Dictionary mapping MP codes to their RA codes
#         mp_codes_with_em: List of MP codes that have EM entries
#         mp_codes: List of all MP codes
#     """
#     wb = load_workbook(workbook_path)
#     ws = wb.active
#     last_row = ws.max_row
    
#     # Helper function to get column letter
#     def get_column_for_header(header: str) -> str:
#         # Ensure header is a string before attempting string operations
#         header_str = str(header) if header is not None else ""
#         for cell in ws[1]:  # First row
#             if cell.value:
#                 # Ensure cell.value is a string before attempting string operations
#                 cell_value_str = str(cell.value) if cell.value is not None else ""
#                 # Replace both \n and \r (Windows newline)
#                 cell_value_clean = cell_value_str.replace('\n', '').replace('\r', '')
#                 header_clean = header_str.replace('\n', '').replace('\r', '')
#                 if cell_value_clean == header_clean:
#                     return get_column_letter(cell.column)
#         return None
    
#     # For each MP, create appropriate formulas
#     for mp_code in mp_codes:
#         ra_codes = mp_groups.get(mp_code, []) # Use .get for safety
#         ra_columns = [get_column_for_header(ra) for ra in ra_codes]
#         ra_columns = [col for col in ra_columns if col is not None]
        
#         if ra_columns:
#             first_ra_col = ra_columns[0]
#             last_ra_col = ra_columns[-1]
            
#             failing_grades_check_parts = []
#             for ra_col in ra_columns:
#                 failing_grades_check_parts.append(f'AND(ISNUMBER({ra_col}{{row}}),{ra_col}{{row}}<5)')
#             any_failing = f'OR({",".join(failing_grades_check_parts)})'
            
#             if mp_code in mp_codes_with_em:
#                 centre_column = get_column_for_header(f'{mp_code} CENTRE')
#                 empresa_column = get_column_for_header(f'{mp_code} EMPRESA')
#                 mp_column_em = get_column_for_header(mp_code) # Renamed to avoid conflict

#                 if centre_column:
#                     for row_idx in range(2, last_row+1): # Student rows
#                         cell = ws[f'{centre_column}{row_idx}']
#                         cell.value = None
                
#                 if mp_column_em and centre_column and empresa_column: # Ensure all columns found
#                     for row_idx in range(2, last_row+1): # Student rows
#                         cell = ws[f'{mp_column_em}{row_idx}']
#                         cell.value = None

#             else: # Regular MPs (Type B)
#                 mp_column_regular = get_column_for_header(mp_code) # Renamed to avoid conflict
#                 if mp_column_regular:
#                     for row_idx in range(2, last_row+1): # Student rows
#                         cell = ws[f'{mp_column_regular}{row_idx}']
#                         cell.value = None
    
#     wb.save(workbook_path)


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
    - Uses Excel formulas to dynamically sum RA percentages for each MP
    - Applies data validation for percentage cells
    - Applies conditional formatting for invalid grades
    - Protects all cells except RA percentage cells
    Uses pre-computed mp_codes list for efficient grouping.
    """
    # Filter out MP grade columns (those that don't end with EM or RA)
    # and exclude the 'estudiant' column from processing
    non_mp_columns = [col for col in df.columns 
                         if col.endswith(('EM', 'RA')) or col == 'estudiant']
    # Remove suffixes from MP grade columns
    df = df.rename(
        columns=lambda col: col.split('_')[0] if col not in non_mp_columns else col
    )
    # Get MP grade columns
    mp_grade_columns = [col for col in df.columns if not col.endswith('EM') and
                         not col.endswith('RA') and col != 'estudiant']
    # Create a filtered DataFrame with non-MP grade columns
    df_without_mp_grades = df[non_mp_columns].copy()
    
    # Get the current column order (excluding 'estudiant')
    ra_codes = [col for col in df_without_mp_grades.columns if col != 'estudiant']
    
    # Group RA codes by their MP. Sort mp_codes by length descending to match longer MPs first.
    # Use startswith(mp_code + '_') for more precise matching.
    mp_groups = {mp: [] for mp in mp_codes}
    sorted_mp_codes = sorted(mp_codes, key=len, reverse=True)  # Sort MPs by length

    for ra_code in ra_codes:
        found_mp = False
        for mp_code in sorted_mp_codes:
            if ra_code.startswith(mp_code + '_'):  # Precise matching
                mp_groups[mp_code].append(ra_code)
                found_mp = True
                break
        if not found_mp and not ra_code.endswith(('EM', 'RA')):
            # Only warn for non-EM, non-RA codes that don't match any MP
            print(f"Warning: Code '{ra_code}' did not match any known MP code prefix and is not an EM/RA code.")
    
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
    
    # Start with the filtered DataFrame that only contains valid grade columns
    export_df = df_without_mp_grades.copy()

    # Create a mapping of EM codes to their corresponding MP codes
    em_to_mp = {}
    for mp_code in mp_codes_with_em:
        # Find all EM codes for this MP (format: MP_CF_1EM or MP_CF_12EM)
        em_codes = [col for col in export_df.columns if col.startswith(f'{mp_code}_') and col.endswith('EM')]
        for em_code in em_codes:
            em_to_mp[em_code] = mp_code
    
    # Add new placeholder columns (MP, CENTRE, EMPRESA) if they don't exist
    for col_name in new_columns:
        if col_name not in export_df.columns:
            export_df[col_name] = pd.Series(dtype='float64', index=export_df.index)
    
    # Ensure all columns from new_columns are present and in the correct order
    export_df = export_df.reindex(columns=new_columns)

    print(mp_grade_columns)
    for mp_code in mp_grade_columns:
        print(mp_code)
        if mp_code in export_df.columns:
            print("Found")
            # `df` here is still the renamed DataFrame, which has the original
            # grade values under column `mp_code`
            export_df[mp_code] = df[mp_code]
    print(export_df.iloc[:, :14])
    # First, export the DataFrame to Excel
    output_path = output_path.replace('.csv', '.xlsx')
    export_df.to_excel(output_path, index=False)
    
    # Now process EM records in the Excel file
    wb = load_workbook(output_path)
    ws = wb.active
    
    # Helper function to get column letter for a header
    def get_col_letter(header):
        for cell in ws[1]:
            if str(cell.value).strip() == header:
                return get_column_letter(cell.column)
        return None
    
    # Process EM records if there are any
    if em_to_mp:
        for em_code, mp_code in em_to_mp.items():
            empresa_header = f'{mp_code} EMPRESA'
            em_col = get_col_letter(em_code)
            empresa_col = get_col_letter(empresa_header)
            
            if em_col and empresa_col:
                # Copy EM grades to EMPRESA column
                for row_idx in range(2, ws.max_row + 1):
                    em_cell = ws[f'{em_col}{row_idx}']
                    empresa_cell = ws[f'{empresa_col}{row_idx}']
                    if em_cell.value is not None:
                        empresa_cell.value = em_cell.value
                
                # Remove the original EM column
                ws.delete_cols(ws[em_col][0].column, 1)
        
        # Save the workbook after processing all EM records
        wb.save(output_path)
    
    # Remove EM codes from mp_groups to avoid processing them as RA columns
    for em_code in em_to_mp:
        mp_code = em_to_mp[em_code]
        if mp_code in mp_groups and em_code in mp_groups[mp_code]:
            mp_groups[mp_code].remove(em_code)
    
    # Reload the DataFrame to reflect any changes
    export_df = pd.read_excel(output_path)
    
    # Convert numeric columns to float64 explicitly
    numeric_cols = export_df.select_dtypes(include=['int64', 'float64']).columns
    for col in numeric_cols:
        # Convert to float64 with NaN for missing values
        export_df[col] = pd.to_numeric(export_df[col], errors='coerce').astype('float64')
    
    # Replace NaN with empty string only in non-numeric columns
    non_numeric_cols = export_df.columns.difference(numeric_cols)
    export_df[non_numeric_cols] = export_df[non_numeric_cols].fillna('')
    
    # Export the final DataFrame to Excel
    export_df.to_excel(output_path, index=False)

    # Apply row formatting (borders, bold, background colors)
    apply_row_formatting(output_path, mp_codes_with_em, mp_codes)
    
    # Apply conditional formatting (now without percentage validation)
    apply_conditional_formatting(output_path, mp_groups, mp_codes_with_em, mp_codes)
    
    print(f"\t- Exported {len(export_df)} entries to {output_path}")