"""
Excel processing module for generating and formatting grade reports.
"""

import pandas as pd


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
    
    # Export to Excel
    output_path = output_path.replace('.csv', '.xlsx')
    export_df.to_excel(output_path, index=False)
    print(f"\t- Exported {len(export_df)} entries to {output_path}")


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