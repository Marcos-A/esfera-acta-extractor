import os
import glob

from src.conversion_service import convert_input_directory


def main() -> None:
    """
    Run the original command-line workflow for local batch processing.

    This entrypoint predates the web UI and is still useful for staff who want to drop
    PDFs into the repository folders and generate both detailed and summary workbooks.
    """
    # Keep the folder names explicit so non-technical users can follow the repository
    # structure without having to trace configuration values.
    input_dir = '01_source_pdfs'
    output_dir = '02_extracted_data'
    os.makedirs(output_dir, exist_ok=True)
    pdf_files = glob.glob(os.path.join(input_dir, '*.pdf'))

    if not pdf_files:
        print("No PDF files found in '01_source_pdfs' directory.")
        return

    try:
        result = convert_input_directory(input_dir, output_dir)
        for artifact in result.artifacts:
            print(f"\t- Successfully extracted data from {artifact.source_name} -> {artifact.output_name}")
    except Exception as e:
        print(f"ERROR processing input directory: {str(e)}")
        return

    summary_output_dir = '03_final_grade_summaries'
    os.makedirs(summary_output_dir, exist_ok=True)

    source_files_pattern = os.path.join('02_extracted_data', '*.xlsx')
    all_potential_source_files = glob.glob(source_files_pattern)
    actual_source_xlsx_files = [f for f in all_potential_source_files if not os.path.basename(f).startswith('~$')]

    if not actual_source_xlsx_files:
        print("No valid processed XLSX files found in '02_extracted_data' to summarize.")
    else:
        print("\nGenerating summary reports...")
        # Imported lazily so the CLI can finish the extraction phase even if summary
        # generation changes its dependencies later.
        from src.summary_generator import generate_summary_report 
        for source_xlsx_file in actual_source_xlsx_files:
            summary_file_name = f"qualificacions_MP-{os.path.basename(source_xlsx_file)}"
            # summary_output_dir is '03_final_grade_summaries', defined above
            output_summary_path = os.path.join(summary_output_dir, summary_file_name)
            try:
                generate_summary_report(source_xlsx_file, output_summary_path)
                print(f"\t- Successfully generated summary: {output_summary_path}")
            except Exception as e:
                print(f"ERROR generating summary for {source_xlsx_file}: {str(e)}")


if __name__ == '__main__':
    main()
