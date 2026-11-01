import argparse
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

REQUIRED_COLUMNS = ["product", "salesperson", "amount"]
CURRENCY_FORMAT = '₹#,##0.00'

def load_and_clean(filename):
    path = Path(filename)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {filename}")

    df = pd.read_csv(path)
    # Normalize headers so variations like ' Product ' or 'PRODUCT' are accepted.
    df.columns = df.columns.astype("string").str.strip().str.lower()
    if df.columns.duplicated().any():
        duplicates = df.columns[df.columns.duplicated()].tolist()
        raise ValueError(f"Duplicate column names after normalization: {', '.join(duplicates)}")

    missing_columns = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")

    original_rows = len(df)
    duplicate_rows = int(df.duplicated().sum())
    df = df.drop_duplicates().copy()

    df["amount"] = (
        df["amount"].astype("string")
        .str.replace("₹", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    invalid_amount_rows = int(df["amount"].isna().sum())
    df = df.dropna(subset=["amount"]).copy()

    for col in ["product", "salesperson"]:
        df[col] = df[col].replace(r"^\s*$", pd.NA, regex=True)
        df[col] = df[col].fillna("Unknown").astype("string").str.strip().str.title()
        df.loc[df[col] == "", col] = "Unknown"

    if df.empty:
        raise ValueError("No valid transactions remain after cleaning.")

    cleaning = {
        "records_received": original_rows,
        "duplicate_rows": duplicate_rows,
        "invalid_amount_rows": invalid_amount_rows,
        "valid_records": len(df),
        "records_removed": original_rows - len(df),
    }
    return df, cleaning

def build_report(df, cleaning, output_file):
    total_sales = float(df["amount"].sum())
    avg_sale = float(df["amount"].mean())
    median_sale = float(df["amount"].median())
    highest_sale = float(df["amount"].max())
    lowest_sale = float(df["amount"].min())
    transactions = len(df)

    product_summary = df.groupby("product")["amount"].agg(
        total_sales="sum", average_sale="mean", transactions="count"
    ).sort_values("total_sales", ascending=False).reset_index()

    person_summary = df.groupby("salesperson")["amount"].agg(
        total_sales="sum", average_sale="mean", transactions="count"
    ).sort_values("total_sales", ascending=False).reset_index()

    top_product = product_summary.iloc[0]
    top_person = person_summary.iloc[0]

    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    products_ws = wb.create_sheet("Sales by Product")
    people_ws = wb.create_sheet("Sales by Salesperson")
    clean_ws = wb.create_sheet("Cleaned Data")

    navy = "17365D"
    blue = "1F4E78"
    light_blue = "D9EAF7"
    pale = "F3F6FA"
    white = "FFFFFF"
    gray = "667085"
    thin_gray = Side(style="thin", color="D9E1F2")

    # Dashboard title and subtitle
    ws.merge_cells("A1:H1")
    ws["A1"] = "SALES PERFORMANCE DASHBOARD"
    ws["A1"].font = Font(name="Aptos Display", size=20, bold=True, color=white)
    ws["A1"].fill = PatternFill("solid", fgColor=navy)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 38

    ws.merge_cells("A2:H2")
    ws["A2"] = "Automated business reporting | Generated from cleaned transaction data"
    ws["A2"].font = Font(size=10, italic=True, color=gray)
    ws["A2"].alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 22

    # KPI cards (two-cell label/value blocks)
    kpis = [
        ("A4:B4", "A5:B6", "TOTAL SALES", total_sales, CURRENCY_FORMAT),
        ("C4:D4", "C5:D6", "TRANSACTIONS", transactions, "0"),
        ("E4:F4", "E5:F6", "AVERAGE SALE", avg_sale, CURRENCY_FORMAT),
        ("G4:H4", "G5:H6", "TOP PRODUCT", str(top_product["product"]), None),
    ]
    for label_range, value_range, label, value, fmt in kpis:
        ws.merge_cells(label_range)
        ws.merge_cells(value_range)
        lc = ws[label_range.split(":")[0]]
        vc = ws[value_range.split(":")[0]]
        lc.value = label
        lc.font = Font(size=9, bold=True, color=white)
        lc.fill = PatternFill("solid", fgColor=blue)
        lc.alignment = Alignment(horizontal="center", vertical="center")
        vc.value = value
        vc.font = Font(size=15, bold=True, color=navy)
        vc.fill = PatternFill("solid", fgColor=light_blue)
        vc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if fmt:
            vc.number_format = fmt
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 24
    ws.row_dimensions[6].height = 18

    # Business insights
    ws.merge_cells("A8:H8")
    ws["A8"] = "BUSINESS HIGHLIGHTS"
    ws["A8"].font = Font(bold=True, color=white, size=11)
    ws["A8"].fill = PatternFill("solid", fgColor=navy)
    highlights = [
        ("A9", f"Top product: {top_product['product']} (₹{top_product['total_sales']:,.2f})"),
        ("A10", f"Top salesperson: {top_person['salesperson']} (₹{top_person['total_sales']:,.2f})"),
        ("A11", f"Median transaction: ₹{median_sale:,.2f} | Highest: ₹{highest_sale:,.2f} | Lowest: ₹{lowest_sale:,.2f}"),
        ("A12", f"Data quality: {cleaning['valid_records']} valid records from {cleaning['records_received']} received"),
    ]
    for cell, value in highlights:
        row = ws[cell].row
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        ws[cell] = value
        ws[cell].font = Font(size=10, color=navy)
        ws[cell].fill = PatternFill("solid", fgColor=pale if row % 2 else white)
        ws[cell].alignment = Alignment(vertical="center")
        ws.row_dimensions[row].height = 21

    # Cleaning summary
    ws["A14"] = "DATA CLEANING CHECK"
    ws["A14"].font = Font(bold=True, color=white)
    ws["A14"].fill = PatternFill("solid", fgColor=navy)
    ws.merge_cells("A14:H14")
    cleaning_rows = [
        ("Records received", cleaning["records_received"]),
        ("Duplicate rows removed", cleaning["duplicate_rows"]),
        ("Invalid/missing amounts removed", cleaning["invalid_amount_rows"]),
        ("Valid records analysed", cleaning["valid_records"]),
        ("Total records removed", cleaning["records_removed"]),
    ]
    for i, (label, value) in enumerate(cleaning_rows, start=15):
        ws.cell(i, 1, label)
        ws.cell(i, 2, value)
        ws.cell(i, 1).font = Font(bold=True, color=navy)
        ws.cell(i, 2).font = Font(color=navy)
        ws.cell(i, 1).fill = PatternFill("solid", fgColor=pale)
        ws.cell(i, 2).fill = PatternFill("solid", fgColor=white)

    # Detail sheets
    def write_table(sheet, frame):
        for c, name in enumerate(frame.columns, start=1):
            cell = sheet.cell(1, c, name.replace("_", " ").title())
            cell.font = Font(bold=True, color=white)
            cell.fill = PatternFill("solid", fgColor=blue)
            cell.alignment = Alignment(horizontal="center")
        for r, row in enumerate(frame.itertuples(index=False, name=None), start=2):
            for c, value in enumerate(row, start=1):
                sheet.cell(r, c, value)
                if r % 2 == 0:
                    sheet.cell(r, c).fill = PatternFill("solid", fgColor=pale)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for col in sheet.columns:
            letter = get_column_letter(col[0].column)
            width = min(max(max(len(str(x.value or "")) for x in col) + 3, 14), 30)
            sheet.column_dimensions[letter].width = width
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = Border(bottom=thin_gray)

    write_table(products_ws, product_summary)
    write_table(people_ws, person_summary)
    write_table(clean_ws, df)

    for sheet in [products_ws, people_ws]:
        for r in range(2, sheet.max_row + 1):
            sheet.cell(r, 2).number_format = CURRENCY_FORMAT
            sheet.cell(r, 3).number_format = CURRENCY_FORMAT

    # Charts on dashboard
    chart1 = BarChart()
    chart1.type = "bar"
    chart1.style = 10
    chart1.title = "Revenue by Product"
    chart1.x_axis.title = "Revenue (₹)"
    chart1.y_axis.title = "Product"
    chart1.height = 7
    chart1.width = 13
    data = Reference(products_ws, min_col=2, min_row=1, max_row=products_ws.max_row)
    cats = Reference(products_ws, min_col=1, min_row=2, max_row=products_ws.max_row)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.legend = None
    ws.add_chart(chart1, "A22")

    chart2 = BarChart()
    chart2.type = "bar"
    chart2.style = 11
    chart2.title = "Revenue by Salesperson"
    chart2.x_axis.title = "Revenue (₹)"
    chart2.y_axis.title = "Salesperson"
    chart2.height = 7
    chart2.width = 13
    data2 = Reference(people_ws, min_col=2, min_row=1, max_row=people_ws.max_row)
    cats2 = Reference(people_ws, min_col=1, min_row=2, max_row=people_ws.max_row)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.legend = None
    ws.add_chart(chart2, "E22")

    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 17
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    for sheet in [products_ws, people_ws, clean_ws]:
        sheet.sheet_view.showGridLines = False

    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return {
        "total_sales": total_sales,
        "transactions": transactions,
        "top_product": str(top_product["product"]),
        "top_salesperson": str(top_person["salesperson"]),
        "output_file": str(output_file),
    }

def main():
    parser = argparse.ArgumentParser(description="Clean sales CSV data and generate an Excel business report.")
    parser.add_argument("input_csv", nargs="?", help="Path to the input CSV file")
    parser.add_argument("-o", "--output", default="sales_business_report.xlsx", help="Output Excel workbook")
    args = parser.parse_args()

    filename = args.input_csv or input("Enter CSV filename: ").strip()
    try:
        df, cleaning = load_and_clean(filename)
        result = build_report(df, cleaning, args.output)
    except (FileNotFoundError, ValueError, pd.errors.ParserError) as exc:
        print(f"Error: {exc}")
        raise SystemExit(1)

    print("\n========== SALES REPORT ==========")
    print(f"Total sales:       ₹{result['total_sales']:,.2f}")
    print(f"Transactions:      {result['transactions']}")
    print(f"Top product:       {result['top_product']}")
    print(f"Top salesperson:   {result['top_salesperson']}")
    print(f"Records received:  {cleaning['records_received']}")
    print(f"Valid records:     {cleaning['valid_records']}")
    print(f"Records removed:   {cleaning['records_removed']}")
    print(f"\nExcel report saved: {result['output_file']}")

if __name__ == "__main__":
    main()
