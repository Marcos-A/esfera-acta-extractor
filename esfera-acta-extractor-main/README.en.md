[Llegiu-lo en català](README.md)

# Esfer@ Acta Extractor

A Python tool to extract and process grade records from Esfer@'s PDF grade reports. The tool parses PDF tables, processes student grades, and generates two types of Excel files:
1. A detailed file with all RA (Learning Achievement), EM (Workplace Training), and MP (Professional Module) grades.
2. A summary report for each MP with student grades and a clear results overview, including both center and company evaluations where applicable.

## Features

### PDF Processing
- Automatic extraction of grade reports from PDF files
- Intelligent processing of student names and grades
- Support for RA (Learning Achievement), EM (Workplace Training), and MP (Professional Module) grades
- Automatic identification of MPs with Workplace Training (EM) hours and their corresponding grades

### Detailed Excel File
- Comprehensive view of all RA, EM and MP grades
- Clear organization by student and MP
- Conditional formatting for better visualization

### Grade Summary
- Consolidated view of grades for each MP
- Different number formats for Type A (2 decimals) and Type B (integer) MPs
- Built-in legend explaining MP types
- Alternating row colors for better readability
- Data validation to ensure correct values
- Professional formatting ready for printing

## Requirements

- Docker (recommended)
- Python 3.9+ (if running locally)
- Required Python packages listed in `requirements.txt`

## Project Structure

```
.
├── 01_source_pdfs/           # Directory for input PDF files
│   └── .gitkeep              # Keeps directory in Git but ignores contents
├── 02_extracted_data/        # Directory for output Excel files
│   └── .gitkeep              # Keeps directory in Git but ignores contents
├── 03_final_grade_summaries/ # Directory for final grade summaries
│   └── .gitkeep              # Keeps directory in Git but ignores contents
├── rules/                    # Project configuration rules
│   └── column-context.md     # Column context rules
├── src/                      # Source code package
│   ├── __init__.py           # Package initialization and exports
│   ├── pdf_processor.py      # PDF extraction and processing
│   ├── data_processor.py     # Data cleaning and transformation
│   ├── grade_processor.py    # Grade-specific operations
│   ├── excel_processor.py    # Excel file generation and formatting
│   └── summary_generator.py  # Grade summary report generation
├── cursor.config.jsonc       # Cursor configuration
├── windsurf.config.jsonc     # Windsurf configuration
├── Dockerfile                # Docker configuration
├── README.md                 # Catalan version 
├── README.en.md              # This file (English)
├── requirements.txt          # Python dependencies
├── LICENSE                   # GNU GPL-3.0 license
└── esfera-acta-extractor.py  # Main script

```

### Module Description

- **pdf_processor.py**: Handles all PDF-related operations
  - Table extraction from PDFs
  - Group code extraction from first page
  
- **data_processor.py**: General data processing utilities
  - Header normalization
  - Column filtering and transformation
  - Data cleaning and reshaping
  
- **grade_processor.py**: Grade-specific logic
  - RA record extraction
  - MP code identification
  - EM entry detection
  - Record sorting
  
- **excel_processor.py**: Excel file handling
  - Excel file generation
  - Column spacing and organization
  - Formatting and formulas

- **summary_generator.py**: Grade summary report generation
  - Creates MP qualification summaries
  - Includes RA codes and student grades
  - Applies conditional formatting for better readability
  - Generates explanatory legend for MP types

## Installation & Usage

### Using Docker (Recommended)

**Important**: Place your Esfer@ PDF file in the root folder of the project before running any Docker commands.

1. Build the Docker image:
```bash
docker build -t esfera-acta-extractor .
```

2. You have two options to run the container:

   a. Direct execution:
   ```bash
   docker run -v $(pwd):/app esfera-acta-extractor python esfera-acta-extractor.py
   ```

   b. Interactive mode (recommended for debugging or multiple files):
   ```bash
   docker run --rm -it \
     -v "$(pwd)":/data \
     -w /data \
     esfera-acta-extractor
   ```
   Once inside the container, you can run:
   ```bash
   python esfera-acta-extractor.py
   ```
   To exit the interactive shell, simply type:
   ```bash
   exit
   ```

    Optional: control how the string `NA` is treated in Excel cells:

    - By default the tool **preserves** the literal `NA` string (it will appear as `NA` in the Excel and receive the corresponding conditional formatting).
    - To restore the previous behaviour (treat `NA` as a missing value and convert it to an empty cell), set the environment variable `PRESERVE_NA=0` before running.

    Example (keep `NA`):
    ```bash
    python esfera-acta-extractor.py
    ```

    Example (treat `NA` as empty, legacy behaviour):
    ```bash
    PRESERVE_NA=0 python esfera-acta-extractor.py
    ```

**Note**: The script will:
- Look for PDF files in the `01_source_pdfs` directory
- Process each PDF file individually
- Generate Excel output files in the `02_extracted_data` directory
- Name each output file based on the group code found in its corresponding PDF
- Skip any files that are not PDFs or cannot be processed
- Continue processing remaining files even if some fail

3. Container cleanup:
   - For interactive mode (`--rm` flag): Container is automatically removed upon exit
   - For background processes: Stop the container with:
   ```bash
   docker stop $(docker ps -q --filter ancestor=esfera-acta-extractor)
   ```

**Note**: The script will:
- Look for the PDF file in the current directory
- Generate the Excel output file in the same directory
- Name the output file based on the group code found in the PDF

### Local Installation

1. Clone the repository:
```bash
git clone https://github.com/Marcos-A/esfera-acta-extractor.git
cd esfera-acta-extractor
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the script:
```bash
python esfera-acta-extractor.py
```

## Input/Output

### Input
- PDF file from Esfer@ containing grade tables
- File should include "Codi del grup" field in the first page
- Expected grade formats: A#, PDT, EP, NA

### Output
- Excel file named with the group code (e.g., `CFPM_AG10101.xlsx`)
- Contains:
  - Student names
  - RA grades grouped by MP
  - 3 empty columns after MPs with EM entries (labeled as CENTRE, EMPRESA, MP)
  - 1 empty column after regular MPs (labeled as MP)

## Development

The project uses:
- `pandas` for data manipulation
- `pdfplumber` for PDF parsing
- `openpyxl` for Excel file generation
- `tabulate` for development/debugging output

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

This means you can:
- Use the software for any purpose
- Change the software to suit your needs
- Share the software with other people
- Share the changes you make

But you must:
- Share the source code when you share the software
- License any derivative work under GPL-3.0
- State significant changes made to the software
- Include the original license and copyright notices

## Acknowledgments

- Built for processing Esfer@'s educational grade reports
- Designed to handle specific PDF formats from the Catalan educational system
