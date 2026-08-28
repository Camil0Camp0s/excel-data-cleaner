"""
Data Cleaner & Report Generator
--------------------------------
Takes a "dirty" CSV/Excel file (typical export from a sales system) and
produces:
  1. A clean, formatted Excel file, ready to use.
  2. A "Summary" sheet with key metrics (sales by city, top product, etc).
  3. A console log of everything that was fixed, so the client can see the
     work that was done.

Usage:
    python clean_data.py sample_messy_data.csv clean_report.xlsx
"""

import sys
import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


def load_data(path: str) -> pd.DataFrame:
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    return pd.read_excel(path)


def clean_text_columns(df: pd.DataFrame) -> pd.DataFrame:
    df["Name"] = df["Name"].str.strip().str.title()
    df["Email"] = df["Email"].str.strip().str.lower()
    df["City"] = df["City"].fillna("Unknown").str.strip().str.title()
    df["City"] = df["City"].replace("", "Unknown")
    df["Product"] = df["Product"].str.strip().str.title()
    return df


def normalize_dates(df: pd.DataFrame) -> pd.DataFrame:
    # Dates arrive in several formats (DD/MM/YYYY, MM/DD/YYYY, YYYY-MM-DD,
    # YYYY.MM.DD). pandas can't guess the right format for all of them at
    # once, so each known format is tried in order and filled in wherever
    # it matches.
    formats = ["%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%Y.%m.%d", "%d-%m-%Y"]
    parsed = pd.Series(pd.NaT, index=df.index)
    for fmt in formats:
        mask = parsed.isna()
        parsed[mask] = pd.to_datetime(df.loc[mask, "Purchase_Date"], format=fmt, errors="coerce")
    df["Purchase_Date"] = parsed
    return df


def handle_missing_values(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    stats = {
        "missing_prices": int(df["Price"].isna().sum()),
        "unknown_cities": int((df["City"] == "Unknown").sum()),
    }
    # Missing price is filled with the average price of that same product
    df["Price"] = df.groupby("Product")["Price"].transform(lambda s: s.fillna(s.mean()))
    return df, stats


def remove_duplicates(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    before = len(df)
    df = df.drop_duplicates(subset=["Email", "Purchase_Date", "Product"])
    after = len(df)
    return df, before - after


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    df["Total"] = df["Quantity"] * df["Price"]
    sales_by_city = df.groupby("City")["Total"].sum().sort_values(ascending=False)
    units_by_product = df.groupby("Product")["Quantity"].sum().sort_values(ascending=False)

    rows = [
        ("Total sales", f"${df['Total'].sum():,.0f}"),
        ("Number of transactions", len(df)),
        ("Top city by sales", sales_by_city.index[0]),
        ("Best-selling product (units)", units_by_product.index[0]),
        ("", ""),
        ("Sales by city", ""),
    ]
    for city, total in sales_by_city.items():
        rows.append((city, f"${total:,.0f}"))
    rows.append(("", ""))
    rows.append(("Units by product", ""))
    for product, units in units_by_product.items():
        rows.append((product, int(units)))

    return pd.DataFrame(rows, columns=["Metric", "Value"])


def style_sheet(ws):
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
    for col_cells in ws.columns:
        length = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = length + 4


def main():
    if len(sys.argv) != 3:
        print("Usage: python clean_data.py <input_file> <output_file.xlsx>")
        sys.exit(1)

    input_path, output_path = sys.argv[1], sys.argv[2]

    df = load_data(input_path)
    original_rows = len(df)

    df = clean_text_columns(df)
    df = normalize_dates(df)
    df, duplicates_removed = remove_duplicates(df)
    df, missing_stats = handle_missing_values(df)
    df = df.sort_values("Purchase_Date").reset_index(drop=True)

    summary = build_summary(df)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Clean Data", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        for sheet in writer.sheets.values():
            style_sheet(sheet)

    print("=== Cleaning Report ===")
    print(f"Original rows:           {original_rows}")
    print(f"Duplicates removed:      {duplicates_removed}")
    print(f"Missing prices filled (with product average): {missing_stats['missing_prices']}")
    print(f"Unknown cities flagged:  {missing_stats['unknown_cities']}")
    print(f"Final rows:              {len(df)}")
    print(f"\nFile generated: {output_path}")


if __name__ == "__main__":
    main()
