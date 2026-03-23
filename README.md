# Arabic Inventory Item Card Generator

This Python script generates three file formats with an Arabic inventory item card layout using RTL (Right-to-Left) formatting:
- Excel file (`Item_Card.xlsx`)
- PDF file (`Item_Card.pdf`)
- Word document (`Item_Card.docx`)

## Requirements

- Python 3.x
- openpyxl library
- reportlab library
- arabic-reshaper library
- python-bidi library
- python-docx library

## Installation

Install the required libraries using pip:

```bash
pip install openpyxl reportlab arabic-reshaper python-bidi python-docx
```

## Usage

Run the script:

```bash
python create_excel.py
```

This will generate three files in the same directory:
- `Item_Card.xlsx` (Excel format)
- `Item_Card.pdf` (PDF format)
- `Item_Card.docx` (Word format)

## Output

The generated files include:
- **RTL Layout**: Proper right-to-left formatting for Arabic text
- **Title**: كارت صنف (Item Card)
- **Header Fields**: Serial number, item name, item number, and unit fields
- **Data Table**: 
  - التاريخ (Date)
  - رقم الإذن (Permit Number)
  - البيان (Description)
  - الوارد (Incoming)
  - المنصرف (Outgoing)
  - الرصيد (Balance)
  - ملاحظات (Notes)
- **25 Empty Data Rows**: Ready for data entry with dotted horizontal borders

## Features

- Professional Arabic formatting
- Proper column widths optimized for content
- Thick borders for headers
- Dotted borders for data rows (matching traditional Arabic forms)
- Centered alignment for table data
- Bold headers with appropriate font sizes
