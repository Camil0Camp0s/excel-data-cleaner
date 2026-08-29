# Excel/CSV Data Cleaner & Report Generator

Python script that takes a CSV or Excel file with "dirty" data (a typical
export from a sales system, CRM, or form) and produces a clean, ready-to-use
file plus a summary report.

## What it solves

Real-world data almost never comes clean. This script automatically fixes:

- **Inconsistent text**: mixed-case names, extra whitespace.
- **Dates in multiple formats** (`15/03/2024`, `2024-03-16`, `2024.03.17`...)
  → all normalized to `YYYY-MM-DD`.
- **Duplicate rows** (the same customer/date/product logged more than once).
- **Missing values**: empty prices are filled with the average for that same
  product; empty cities are flagged as "Unknown" instead of left blank.

It then generates an Excel file with two sheets:
1. **Clean Data** — the corrected table, with formatted headers.
2. **Summary** — total sales, top city, best-selling product, and a
   breakdown by city and by product.

## Usage

```bash
pip install -r requirements.txt
python clean_data.py sample_messy_data.csv clean_report.xlsx
```

The script prints a console log of everything it fixed (useful to show a
client exactly what work was done):

```
=== Cleaning Report ===
Original rows:           12
Duplicates removed:      4
Missing prices filled (with product average): 1
Unknown cities flagged:  2
Final rows:              8
```

## Adaptable to other cases

The pipeline (clean text → normalize dates → remove duplicates → fill
missing values → generate summary) can be adapted to any tabular dataset:
inventory, sales leads, attendance records, contact lists, etc. Only the
column names and the summary's business rules need to change.

## Portfolio note

This project was built as a portfolio piece for freelancing (data
automation / Excel cleaning), using fictional sample data — it does not
contain real information from any client.

## License

Published for portfolio and educational purposes only — see [LICENSE](LICENSE).
