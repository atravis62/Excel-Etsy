# Excel Etsy Purchase Organizer

This repository contains a small Python utility that works as an Excel add-on to organize and catalog your Etsy purchases. It processes an Excel workbook and creates a catalog sheet with your unique laser engraving files, helping you avoid duplicate purchases.

## Requirements
- Python 3.11 or higher
- `openpyxl`

Install the dependency with:

```bash
pip install openpyxl
```

## Usage
1. Create an Excel workbook (e.g., `purchases.xlsx`) with a sheet named **Purchases**. The sheet should have the following columns starting in row 1:
   - `Item ID`
   - `Item Name`
   - `File Path`
   - `Purchase Date`
2. Run the script with the path to your workbook:

```bash
python src/etsy_excel_addon.py purchases.xlsx
```

The script will:
- Generate or update a `Catalog` sheet with unique items sorted by name.
- Mark duplicate entries in the `Purchases` sheet with a `Duplicate` column.

After running the script, open your workbook in Excel to browse the organized catalog.

## Notes
This utility does not connect directly to the Etsy API. You will need to manually export or enter your purchase information into the workbook.
