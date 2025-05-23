# Strongbox Parser

A Python tool for converting financial data from Strongbox Excel files into an Audit Sight compatible format.

## Features

- Converts Strongbox Excel financial data into the Audit Sight template format
- Processes transaction data from TXN-FY sheets
- Extracts trial balance data from the TB sheet
- Filters transactions based on a specified date range
- User-friendly GUI for file selection and date input
- Robust error handling for problematic Excel sheets

## Requirements

- Python 3.6+
  - pandas
  - openpyxl
- xlwings
- tkinter (included with standard Python installation)

## Usage

1. Run the script:
   ```
   python strongbox_parser.py
   ```
2. Select the Strongbox Excel file when prompted
3. Choose an output directory
4. Enter the start and end dates for the period you want to analyze
5. Enter a name for the output file
6. Click "Start Processing"

The program will create an Excel file containing:
- Journal Entries & Lines
- Comparative Trial Balances
- Additional template tabs from the Audit Sight Template

## Error Handling

The parser includes fallback mechanisms to handle problematic Excel sheets. If a sheet cannot be read with pandas, it will attempt to read it using openpyxl. If a sheet still cannot be processed, the program will display a warning message and skip that sheet.

## Setup

1. Create and activate a virtual environment:
   ```bash
   # Create virtual environment
   python3 -m venv venv
   
   # Activate virtual environment
   # On macOS/Linux:
   source venv/bin/activate
   # On Windows:
   .\venv\Scripts\activate
   ```

2. Install the required packages:
   ```bash
   pip install -r requirements.txt --index-url https://pypi.org/simple
   ```

## Output

The output file will be named `Audit_Sight_Output_YYYYMMDD_YYYYMMDD.xlsx` where the dates represent your selected date range.

## Notes

- The script will automatically handle multiple fiscal year tabs (TXN-FY1, TXN-FY2, etc.) based on your date range
- The Comparative Trial Balances tab will use the end of the month before your start date for beginning balances
- All formatting from the Audit Sight Template will be preserved in the output file

## Troubleshooting

If you encounter any issues with package installation:
1. Make sure your virtual environment is activated (you should see `(venv)` at the start of your command prompt)
2. Try installing packages with the public PyPI repository:
   ```bash
   pip install -r requirements.txt --index-url https://pypi.org/simple
   ``` 