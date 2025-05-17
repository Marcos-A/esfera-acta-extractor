[Llegiu-lo en català](README.md)

# Esfer@ Acta Extractor

A Python tool to extract and process grade records from Esfer@'s PDF grade reports. The tool parses PDF tables, processes student grades, and generates an Excel file with a structured view of RA (Learning Achievement) grades and MP (Professional Module) qualifications.

## Features

- Extracts tables from PDF grade reports
- Processes student names and grades
- Identifies MPs with EM (Workplace Training) entries
- Generates Excel files with:
  - Student grades for each RA
  - Properly spaced columns for MP qualifications
  - Special column layout for MPs with EM entries

## Requirements

- Docker (recommended)
- Python 3.9+ (if running locally)
- Required Python packages listed in `requirements.txt`

## Project Structure

```
.
├── src/                      # Source code package
│   ├── __init__.py           # Package initialization and exports
│   ├── pdf_processor.py      # PDF extraction and processing
│   ├── data_processor.py     # Data cleaning and transformation
│   ├── grade_processor.py    # Grade-specific operations
│   └── excel_processor.py    # Excel file generation and formatting
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
  - (Future) Formatting and formulas

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
git clone https://github.com/yourusername/esfera-acta-extractor.git
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
- Share the software with your friends and neighbors
- Share the changes you make

But you must:
- Share the source code when you share the software
- License any derivative work under GPL-3.0
- State significant changes made to the software
- Include the original license and copyright notices

## Acknowledgments

- Built for processing Esfer@'s educational grade reports
- Designed to handle specific PDF formats from the Catalan educational system