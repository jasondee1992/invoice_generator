# Invoice Generator

Python script for generating a professional invoice PDF from Excel data and an evidence images folder.

## Features

- Reads invoice data from an Excel workbook
- Generates a clean invoice summary PDF page
- Appends evidence images starting on page 2 onward
- Supports up to 4 evidence images per page in a 2x2 grid
- Preserves image aspect ratio
- Repeats the invoice table header on continuation pages
- Moves `Subtotal`, `Total`, and `Payment Details` to the final invoice page only
- Uses configurable input and output paths
- Can be packaged as a Windows `.exe` for double-click use

## Project Structure

```text
invoice_generator/
├─ generate_invoice.py
├─ build_exe.bat
├─ icon/
│  └─ icon.png
├─ requirements.txt
├─ README.md
├─ data/
│  └─ invoice_data.xlsx
├─ evidences/
│  ├─ image1.jpg
│  ├─ image2.png
│  └─ ...
└─ output/
   └─ invoice_output.pdf
```

## Requirements

- Python 3.10+
- `pandas`
- `openpyxl`
- `reportlab`
- `pillow`

Install dependencies:

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install pandas openpyxl reportlab pillow
```

## Excel Format

The script expects **2 worksheets** in the Excel file:

- `invoice_info`
- `line_items`

### Sheet 1: `invoice_info`

This sheet should contain **one row only** for one invoice per run.

Required columns:

```text
invoice_number
invoice_date
due_date
sender_name
sender_address
client_name
client_address
payment_details
contact_email
```

Optional columns:

```text
sender_phone
sender_email
client_company
client_email
currency_symbol
subtotal
total
notes
```

Example:

| invoice_number | invoice_date | due_date | sender_name | sender_address | client_name | client_address | payment_details | contact_email | currency_symbol |
|---|---|---|---|---|---|---|---|---|---|
| INV-2026-001 | March 20, 2026 | March 27, 2026 | Jason Dela Cruz | 123 Main Street | ABC Client | 456 Market Road | Bank: BDO | juan@example.com | Php |

### Sheet 2: `line_items`

This sheet should contain one or more service rows.

Required columns:

```text
description
service_date
rate
```

Optional columns:

```text
quantity
amount
```

Example:

| description | service_date | quantity | rate | amount |
|---|---|---:|---:|---:|
| Graphic design work | 2026-03-10 | 1 | 300 | 300 |
| Revision and final export | 2026-03-15 | 1 | 120 | 120 |

## Multi-line Description in Excel

If you want a single line item description to appear on multiple lines in the PDF:

1. Open the description cell in Excel
2. Press `Alt + Enter`
3. Type the next line

Example inside one Excel cell:

```text
Graphic design work
Graphic design work
Graphic design work
```

The script preserves those line breaks in the PDF and automatically increases the row height.

## Evidence Images

Place all evidence images inside the `evidences/` folder.

Supported formats:

- `.jpg`
- `.jpeg`
- `.png`
- `.webp`
- `.bmp`

Behavior:

- Page 1 and possible continuation invoice pages contain invoice details and line items
- Evidence images start after the final invoice page
- Maximum 4 images per page
- Images are arranged in a 2x2 grid
- Filenames are shown as small captions below each image

## Running the Script

Default run:

```bash
python generate_invoice.py
```

Using the project virtual environment on Windows:

```powershell
.\venv\Scripts\python.exe generate_invoice.py
```

Custom paths:

```powershell
.\venv\Scripts\python.exe generate_invoice.py --input-excel-path data/invoice_data.xlsx --evidence-folder-path evidences --output-pdf-path output/invoice_output.pdf
```

## Windows Executable (.exe)

You can package the script into a Windows executable so users can just double-click it.

### Build the `.exe`

Run:

```powershell
.\build_exe.bat
```

This will:

- install `PyInstaller` inside the project virtual environment
- convert `icon/icon.png` into `icon/icon.ico`
- build `InvoiceGenerator.exe`
- place the `.exe` in the project root

After build, your project can look like this:

```text
invoice_generator/
├─ InvoiceGenerator.exe
├─ data/
├─ evidences/
└─ output/
```

### Use the `.exe`

1. Put the Excel file in `data/invoice_data.xlsx`
2. Put evidence images in `evidences/`
3. Double-click `InvoiceGenerator.exe`
4. The PDF will be generated in `output/invoice_output.pdf`

Important:

- the `.exe` uses paths relative to its own folder
- keep `data/`, `evidences/`, and `output/` beside the `.exe`
- if an error happens, the packaged app shows a popup message
- the executable icon comes from `icon/icon.png`

### Sharing with Others

If someone clones the repo, they have two options:

1. Build the `.exe` themselves using `build_exe.bat`
2. Use a copy of the already-built `InvoiceGenerator.exe`

If you want them to double-click immediately after cloning, you need to include `InvoiceGenerator.exe` in the repo or provide it separately.

## Command-line Arguments

- `--input-excel-path`
  Path to the Excel workbook
- `--evidence-folder-path`
  Path to the evidence images folder
- `--output-pdf-path`
  Path of the generated PDF

Example:

```bash
python generate_invoice.py --input-excel-path data/invoice_data.xlsx --evidence-folder-path evidences --output-pdf-path output/my_invoice.pdf
```

## Output Behavior

- Invoice details are drawn first
- If the line items do not fit on one page, the script creates continuation invoice pages
- The table header is repeated on each continuation invoice page
- `Subtotal`, `Total`, and `Payment Details` appear only on the final invoice page
- `Payment Details` is anchored near the bottom of the final invoice page

## Currency Notes

The currency display comes from the `currency_symbol` column in `invoice_info`.

Examples:

- `Php`
- `PHP`
- `$`

Note:

Using the actual peso sign `₱` may show as a square if the default PDF font does not support it. If that happens, use `Php` or `PHP` unless the script is updated to use a Unicode font.

## Common Errors

### Missing worksheet(s): invoice_info, line_items

Cause:

Your Excel file does not contain the expected sheet names.

Fix:

Rename or create these worksheets exactly:

- `invoice_info`
- `line_items`

### Missing columns in sheet

Cause:

One or more required columns are missing.

Fix:

Compare your Excel columns with the required format above.

### No image files found in the evidence folder

Cause:

The evidence folder is empty or contains unsupported file types only.

Fix:

Add supported image files to the evidence folder.

### Black square instead of `₱`

Cause:

The default PDF font does not support the peso symbol.

Fix:

Use `Php` or `PHP` in the `currency_symbol` column, or update the script to use a Unicode font.

## Notes for Debugging in PyCharm / IntelliJ

If you right-click `generate_invoice.py` and choose `Debug`, the script uses its default arguments unless you add custom script parameters.

To set script arguments:

1. Open `Run | Edit Configurations`
2. Select the Python configuration for `generate_invoice.py`
3. Add values in `Script parameters`, for example:

```text
--input-excel-path data/invoice_data.xlsx --evidence-folder-path evidences --output-pdf-path output/debug_invoice.pdf
```

## Maintainer Notes

Main file:

- [generate_invoice.py](D:/Projects/invoice_generator/generate_invoice.py)

Main responsibilities inside the script:

- load and validate Excel data
- calculate line item row heights
- paginate invoice table rows
- draw invoice summary and payment section
- draw evidence image pages
