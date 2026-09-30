import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference


# ============================================================
# 1. GET INPUT FILE
# ============================================================

filename = input("Enter CSV filename: ")

if not filename.endswith(".csv"):
    print("❌ Please provide a CSV file.")
    exit()

try:
    df = pd.read_csv(filename)
except FileNotFoundError:
    print("❌ File not found. Please check the filename.")
    exit()


# ============================================================
# 2. VALIDATE REQUIRED COLUMNS
# ============================================================

required_columns = ["product", "salesperson", "amount"]

for column in required_columns:
    if column not in df.columns:
        print(f"❌ Required column missing: {column}")
        exit()


# ============================================================
# 3. RECORD ORIGINAL ROW COUNT
# ============================================================

original_rows = len(df)


# ============================================================
# 4. CLEAN DATA
# ============================================================

# Remove duplicate rows
df = df.drop_duplicates()

# Clean amount column
df["amount"] = (
    df["amount"]
    .astype("string")
    .str.replace("₹", "", regex=False)
    .str.replace(",", "", regex=False)
    .str.strip()
)

# Convert amount to numeric
# Invalid values become NaN
df["amount"] = pd.to_numeric(
    df["amount"],
    errors="coerce"
)

# Remove transactions with invalid/missing amounts
df = df.dropna(subset=["amount"])


# Clean empty/whitespace-only salesperson values
df["salesperson"] = df["salesperson"].replace(
    r"^\s*$",
    pd.NA,
    regex=True
)

# Clean empty/whitespace-only product values
df["product"] = df["product"].replace(
    r"^\s*$",
    pd.NA,
    regex=True
)

# Replace missing names
df["salesperson"] = df["salesperson"].fillna("Unknown")
df["product"] = df["product"].fillna("Unknown")

# Normalize text
df["salesperson"] = (
    df["salesperson"]
    .astype("string")
    .str.strip()
    .str.title()
)

df["product"] = (
    df["product"]
    .astype("string")
    .str.strip()
    .str.title()
)


# ============================================================
# 5. CLEANING METRICS
# ============================================================

cleaned_rows = len(df)

removed_rows = original_rows - cleaned_rows


# ============================================================
# 6. BUSINESS ANALYSIS
# ============================================================

total_sales = df["amount"].sum()

average_sale = df["amount"].mean()

median_sale = df["amount"].median()

highest_sale = df["amount"].max()

lowest_sale = df["amount"].min()

number_of_transactions = len(df)

sales_by_person = (
    df.groupby("salesperson")["amount"]
    .sum()
)

sales_for_each_product = (
    df.groupby("product")["amount"]
    .sum()
)


# ============================================================
# 7. BUSINESS INSIGHTS
# ============================================================

top_product = sales_for_each_product.idxmax()
top_product_sales = sales_for_each_product.max()

top_salesperson = sales_by_person.idxmax()
top_salesperson_sales = sales_by_person.max()


# ============================================================
# 8. TERMINAL REPORT
# ============================================================

print()
print("========================================")
print("           SALES REPORT")
print("========================================")
print()

print(f"Total Sales:          ₹{total_sales:,.2f}")
print(f"Average Sale:         ₹{average_sale:,.2f}")
print(f"Median Sale:          ₹{median_sale:,.2f}")
print(f"Highest Sale:         ₹{highest_sale:,.2f}")
print(f"Lowest Sale:          ₹{lowest_sale:,.2f}")
print(f"Transactions:         {number_of_transactions}")
print()

print("Sales by Salesperson")
print("----------------------------------------")
print(sales_by_person)
print()

print("Sales by Product")
print("----------------------------------------")
print(sales_for_each_product)
print()

print("Business Insights")
print("----------------------------------------")
print(
    f"{top_product} generated the most revenue: "
    f"₹{top_product_sales:,.2f}"
)
print(
    f"{top_salesperson} generated the most sales: "
    f"₹{top_salesperson_sales:,.2f}"
)
print()

print("Data Cleaning")
print("----------------------------------------")
print(f"Records Received:     {original_rows}")
print(f"Valid Records:        {cleaned_rows}")
print(f"Records Removed:      {removed_rows}")

print()
print("========================================")


# ============================================================
# 9. CREATE EXCEL WORKBOOK
# ============================================================

workbook = Workbook()

summary_sheet = workbook.active
summary_sheet.title = "Summary"

product_sheet = workbook.create_sheet(
    "Sales by Product"
)

person_sheet = workbook.create_sheet(
    "Sales by Salesperson"
)

data_sheet = workbook.create_sheet(
    "Cleaned Data"
)


# ============================================================
# 10. SUMMARY SHEET
# ============================================================

summary_sheet["A1"] = "Metric"
summary_sheet["B1"] = "Value"

summary_sheet["A2"] = "Total Sales"
summary_sheet["B2"] = total_sales

summary_sheet["A3"] = "Average Sale"
summary_sheet["B3"] = average_sale

summary_sheet["A4"] = "Median Sale"
summary_sheet["B4"] = median_sale

summary_sheet["A5"] = "Highest Sale"
summary_sheet["B5"] = highest_sale

summary_sheet["A6"] = "Lowest Sale"
summary_sheet["B6"] = lowest_sale

summary_sheet["A7"] = "Number of Transactions"
summary_sheet["B7"] = number_of_transactions

summary_sheet["A8"] = "Records Removed"
summary_sheet["B8"] = removed_rows

summary_sheet["A10"] = "Business Insights"

summary_sheet["A11"] = "Top Product"
summary_sheet["B11"] = top_product

summary_sheet["A12"] = "Top Product Revenue"
summary_sheet["B12"] = top_product_sales

summary_sheet["A13"] = "Top Salesperson"
summary_sheet["B13"] = top_salesperson

summary_sheet["A14"] = "Top Salesperson Revenue"
summary_sheet["B14"] = top_salesperson_sales


# ============================================================
# 11. SALES BY PRODUCT SHEET
# ============================================================

product_sheet["A1"] = "Product"
product_sheet["B1"] = "Total Sales"
product_sheet["C1"] = "Average Sale"
product_sheet["D1"] = "Number of Transactions"

for product in sales_for_each_product.index:

    row = product_sheet.max_row + 1

    product_sheet[f"A{row}"] = product

    product_sheet[f"B{row}"] = (
        sales_for_each_product[product]
    )

    product_sheet[f"C{row}"] = (
        df[df["product"] == product]["amount"].mean()
    )

    product_sheet[f"D{row}"] = (
        df["product"].value_counts()[product]
    )


# ============================================================
# 12. SALES BY SALESPERSON SHEET
# ============================================================

person_sheet["A1"] = "Salesperson"
person_sheet["B1"] = "Total Sales"
person_sheet["C1"] = "Average Sale"
person_sheet["D1"] = "Number of Transactions"

for person in sales_by_person.index:

    row = person_sheet.max_row + 1

    person_sheet[f"A{row}"] = person

    person_sheet[f"B{row}"] = (
        sales_by_person[person]
    )

    person_sheet[f"C{row}"] = (
        df[df["salesperson"] == person]["amount"].mean()
    )

    person_sheet[f"D{row}"] = (
        df["salesperson"].value_counts()[person]
    )


# ============================================================
# 13. CLEANED DATA SHEET
# ============================================================

for column_index, column_name in enumerate(
    df.columns,
    start=1
):
    data_sheet.cell(
        row=1,
        column=column_index,
        value=column_name
    )


for row_index, row_data in enumerate(
    df.itertuples(index=False),
    start=2
):

    for column_index, value in enumerate(
        row_data,
        start=1
    ):

        data_sheet.cell(
            row=row_index,
            column=column_index,
            value=value
        )


# ============================================================
# 14. FORMAT HEADERS
# ============================================================

header_fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

header_font = Font(
    bold=True,
    color="FFFFFF"
)

for sheet in workbook.worksheets:

    for cell in sheet[1]:

        cell.fill = header_fill

        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center"
        )


# ============================================================
# 15. FORMAT CURRENCY COLUMNS
# ============================================================

currency_format = '₹#,##0.00'

# Summary
for cell in ["B2", "B3", "B4", "B5", "B6", "B12", "B14"]:
    summary_sheet[cell].number_format = currency_format

# Product sheet
for row in range(2, product_sheet.max_row + 1):
    product_sheet[f"B{row}"].number_format = currency_format
    product_sheet[f"C{row}"].number_format = currency_format

# Person sheet
for row in range(2, person_sheet.max_row + 1):
    person_sheet[f"B{row}"].number_format = currency_format
    person_sheet[f"C{row}"].number_format = currency_format


# ============================================================
# 16. AUTO-SIZE COLUMNS
# ============================================================

for sheet in workbook.worksheets:

    for column in sheet.columns:

        max_length = 0

        column_letter = column[0].column_letter

        for cell in column:

            if cell.value is not None:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        sheet.column_dimensions[
            column_letter
        ].width = max_length + 2


# ============================================================
# 17. FREEZE HEADER ROWS
# ============================================================

for sheet in workbook.worksheets:
    sheet.freeze_panes = "A2"


# ============================================================
# ============================================================
# ============================================================
# ============================================================
# 18. PRODUCT SALES CHART
# ============================================================

from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import Series

product_chart = BarChart()

product_chart.type = "col"
product_chart.title = "Sales by Product"
product_chart.y_axis.title = "Total Sales"
product_chart.x_axis.title = "Product"

# Product names → X-axis
product_categories = Reference(
    product_sheet,
    min_col=1,
    min_row=2,
    max_row=product_sheet.max_row
)

# Sales values → ONE series
product_values = Reference(
    product_sheet,
    min_col=2,
    min_row=2,
    max_row=product_sheet.max_row
)

# Create exactly one series
product_series = Series(
    product_values,
    title="Total Sales"
)

product_chart.append(product_series)

product_chart.set_categories(product_categories)

product_chart.height = 8
product_chart.width = 14

product_sheet.add_chart(
    product_chart,
    "F2"
)
# 19. SALESPERSON SALES CHART
# ============================================================

person_chart = BarChart()

person_chart.type = "col"

person_chart.title = "Sales by Salesperson"

person_chart.y_axis.title = "Total Sales"

person_chart.x_axis.title = "Salesperson"


person_data = Reference(
    person_sheet,
    min_col=2,
    min_row=1,
    max_row=person_sheet.max_row
)

person_categories = Reference(
    person_sheet,
    min_col=1,
    min_row=2,
    max_row=person_sheet.max_row
)

person_chart.add_data(
    person_data,
    titles_from_data=True
)

person_chart.set_categories(
    person_categories
)

person_chart.height = 8
person_chart.width = 14

person_sheet.add_chart(
    person_chart,
    "F2"
)


# ============================================================
# 20. SAVE FINAL REPORT
# ============================================================

workbook.save(
    "sales_business_report.xlsx"
)

print()
print("✅ Excel report created successfully:")
print("   sales_business_report.xlsx")