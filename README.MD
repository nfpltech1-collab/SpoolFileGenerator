# Spool File Generator (GST) v2

A Python-based GUI application to generate spool files for GST filing by extracting data from PDF invoices and matching it with Excel/CSV records.

## Tech Stack
- Python 3.x
- Tkinter (GUI)
- pandas (Data Processing)
- pdfplumber (PDF Extraction)
- PyMuPDF (Digital Signature Detection)

## Project Structure
- `spool_file_generator_v2.py`: Main monolithic application file.
- `modular_app/`: Directory containing the modular version of the application.
- `venv/`: Virtual environment directory.
- `SpoolOutput/`: Generated spool files are saved here.
- `Excel/`: Directory for Excel/CSV data files.
- `Invoice/`: Directory for Invoice PDF files.

## Installation

### Clone the repository
```bash
git clone <repository-url>
cd SpoolFileGenerator
```

### Setup Virtual Environment

1. **Create virtual environment**
   ```bash
   python -m venv venv
   ```

2. **Activate environment**
   - Windows:
     ```bash
     venv\Scripts\activate
     ```
   - Mac/Linux:
     ```bash
     source venv/bin/activate
     ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the application**
   ```bash
   python spool_file_generator_v2.py
   # OR for modular version
   python modular_app/main.py
   ```

## Creating Executable

To build a standalone `.exe` file:

```bash
# For Monolithic Version
pyinstaller --onefile --noconsole --name "SpoolGenerator" spool_file_generator_v2.py

# For Modular Version
pyinstaller --onefile --noconsole --name "SpoolGeneratorModular" modular_app/main.py
```

## Features
- **PDF Extraction**: Extracts Invoice No, Date, Vendor Code, GSTIN, IRN, and Line Items from PDF invoices.
- **Data Matching**: Matches extracted PDF data with Excel/CSV records based on Part Number.
- **Validation**:
  - IRN Number (64-char hex)
  - "Original for Recipient" check
  - Digital Signature verification
  - GSTIN format validation
- **Spool Generation**: Generates fixed-width spool files.
- **GUI**: User-friendly interface with preview table and bulk processing support.

## Environment Variables
Create a `.env` file if needed (currently not used for secrets, but good practice).

## License
Proprietary / Internal Use
