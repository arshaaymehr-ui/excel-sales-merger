# excel-sales-merger

Merge multiple branch Excel sales files into one clean report with a summary sheet.

Python · pandas · openpyxl

[نسخه فارسی](README.fa.md)

## What it does

- Reads every `.xlsx` file in a folder and merges them
- Cleans messy data: empty cells, text in number columns, Persian digits, negative values, duplicates
- Builds a summary of sales by branch and by product
- Exports a formatted Excel report with three sheets

## Usage

```bash
pip install -r requirements.txt
python generate_sample_data.py   # optional: creates sample data
python merge_sales_reports.py
```

Change `DATA_FOLDER` and `OUTPUT_FILE` at the top of `merge_sales_reports.py` to use your own files.

## Expected columns

`تاریخ`, `شعبه`, `محصول`, `تعداد`, `قیمت واحد`

## Output

`sales_report.xlsx` with sheets: merged data, sales by branch, sales by product.