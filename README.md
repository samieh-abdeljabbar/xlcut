# XLCut - XML to Excel Converter

A simple Python tool that converts XML files to Excel format with smart auto-detection of XML structure.

## Features

- Converts multiple XML files at once
- Auto-detects repeating elements in XML (rows)
- Flattens nested XML structures into columns
- Handles XML attributes
- Adds source file tracking when processing multiple files
- Formatted Excel output with headers and alternating row colors
- Timestamped output files (no overwrites)

## Setup & Usage

### Mac

**First-Time Setup:**
1. Open Terminal
2. Navigate to the xlcut folder:
   ```bash
   cd ~/Desktop/xlcut
   ```
3. Create and activate the virtual environment (one-time):
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   pip install -r requirements.txt
   ```

**Running the Script:**
1. Put your XML files in the `source` folder
2. Open Terminal and run:
   ```bash
   cd ~/Desktop/xlcut
   source venv/bin/activate
   python xlcut.py
   ```
3. Find your Excel file in the `output` folder

### Windows

**First-Time Setup:**
1. Open Command Prompt or PowerShell
2. Navigate to the xlcut folder:
   ```cmd
   cd %USERPROFILE%\Desktop\xlcut
   ```
   Or in PowerShell:
   ```powershell
   cd $env:USERPROFILE\Desktop\xlcut
   ```
3. Create and activate the virtual environment (one-time):
   ```cmd
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   ```

**Running the Script:**
1. Put your XML files in the `source` folder
2. Open Command Prompt or PowerShell and run:
   ```cmd
   cd %USERPROFILE%\Desktop\xlcut
   venv\Scripts\activate
   python xlcut.py
   ```
3. Find your Excel file in the `output` folder

### Output

The script creates an Excel file with:
- **"Items Sold" sheet** - Each item sold as a row with Date, Register, Cashier, Description, Quantity, Unit Price, Total, Department, UPC
- **Additional sheets** - Grouped by transaction type (journal, sale, etc.)

## Example

Given this XML file:
```xml
<transactions>
    <transaction id="001">
        <date>2024-01-15</date>
        <amount>125.50</amount>
        <customer>
            <name>John Doe</name>
            <email>john@example.com</email>
        </customer>
        <status>completed</status>
    </transaction>
    <transaction id="002">
        <date>2024-01-16</date>
        <amount>89.99</amount>
        <customer>
            <name>Jane Smith</name>
            <email>jane@example.com</email>
        </customer>
        <status>pending</status>
    </transaction>
</transactions>
```

The tool produces an Excel file with these columns:

| @id | date | amount | customer.name | customer.email | status |
|-----|------|--------|---------------|----------------|--------|
| 001 | 2024-01-15 | 125.50 | John Doe | john@example.com | completed |
| 002 | 2024-01-16 | 89.99 | Jane Smith | jane@example.com | pending |

## How It Works

1. **Auto-detection**: The tool finds repeating XML elements (like `<transaction>`) and treats each as a row
2. **Flattening**: Nested elements become dot-notation columns (e.g., `customer.name`)
3. **Attributes**: XML attributes are prefixed with `@` (e.g., `@id`)
4. **Multiple files**: When processing multiple XML files, a `_source_file` column is added

## Folder Structure

```
xlcut/
├── source/          ← Put XML files here
├── output/          ← Excel files generated here
├── xlcut.py         ← Main script
├── venv/            ← Python virtual environment
├── requirements.txt ← Dependencies
└── README.md        ← This file
```

## Requirements

- Python 3.11+
- lxml
- openpyxl

## Troubleshooting

**"No XML files found"**
- Make sure your XML files are in the `source` folder
- Check that files have `.xml` extension

**"No data found"**
- The XML file may be empty or have no repeating elements
- Check that your XML is valid

**Module not found errors**
- Make sure you activated the virtual environment: `source venv/bin/activate`
- Reinstall dependencies: `pip install -r requirements.txt`
