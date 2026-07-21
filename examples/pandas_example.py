"""
Pandas DataFrame Example
========================
"""
try:
    import pandas as pd
    import numpy as np
except ImportError:
    print("Install pandas: pip install pandas numpy")
    exit(1)

from rvgsrust_xlsxwriter import Workbook
from rvgsrust_xlsxwriter.dataframe import write_pandas_dataframe

# Create sample Pandas DataFrame
np.random.seed(42)
df = pd.DataFrame({
    "Product": [f"Item_{i:03d}" for i in range(1, 21)],
    "Category": np.random.choice(["Electronics", "Clothing", "Food", "Books"], 20),
    "Price": np.random.uniform(10, 500, 20).round(2),
    "Quantity": np.random.randint(1, 100, 20),
    "In Stock": np.random.choice([True, False], 20),
})

# Calculate Revenue
df["Revenue"] = (df["Price"] * df["Quantity"]).round(2)

# Create workbook
wb = Workbook("pandas_output.xlsx")
ws = wb.add_worksheet("Inventory")

# Formats
header_fmt = wb.add_format()
header_fmt.set_bold()
header_fmt.set_background_color("#203764")
header_fmt.set_font_color("white")
header_fmt.set_align("center")
header_fmt.set_border("thin")

price_fmt = wb.add_format()
price_fmt.set_num_format("$#,##0.00")
price_fmt.set_border("thin")
price_fmt.set_border_color("#D9D9D9")

revenue_fmt = wb.add_format()
revenue_fmt.set_num_format("$#,##0.00")
revenue_fmt.set_bold()
revenue_fmt.set_border("thin")
revenue_fmt.set_border_color("#D9D9D9")

stock_true = wb.add_format()
stock_true.set_background_color("#C6E0B4")
stock_true.set_align("center")
stock_true.set_border("thin")

stock_false = wb.add_format()
stock_false.set_background_color("#F8CBAD")
stock_false.set_align("center")
stock_false.set_border("thin")

# Column formats
col_formats = {
    "Price": price_fmt,
    "Revenue": revenue_fmt,
}

# Write DataFrame
write_pandas_dataframe(ws, df, row=0, col=0, header_format=header_fmt, column_formats=col_formats)

# Apply conditional formatting for In Stock column
for row_idx in range(len(df)):
    in_stock = df.iloc[row_idx]["In Stock"]
    fmt = stock_true if in_stock else stock_false
    ws.write(row_idx + 1, 4, in_stock, fmt)

# Summary statistics
summary_row = len(df) + 2
ws.write(summary_row, 0, "SUMMARY", wb.add_format().set_bold().set_font_size(12))

stats = [
    ("Total Products", len(df)),
    ("Avg Price", f"=AVERAGE(C2:C{len(df)+1})"),
    ("Total Revenue", f"=SUM(F2:F{len(df)+1})"),
    ("In Stock Count", f"=COUNTIF(E2:E{len(df)+1},TRUE)"),
]

stat_fmt = wb.add_format()
stat_fmt.set_bold()
stat_fmt.set_background_color("#E7E6E6")
stat_fmt.set_border("thin")

for i, (label, value) in enumerate(stats):
    ws.write(summary_row + 1 + i, 0, label, stat_fmt)
    if isinstance(value, str) and value.startswith("="):
        ws.write_formula(summary_row + 1 + i, 1, value, price_fmt)
    else:
        ws.write(summary_row + 1 + i, 1, value, stat_fmt)

# Set widths
ws.set_column_width(0, 14)
ws.set_column_width(1, 14)
ws.set_column_width(2, 12)
ws.set_column_width(3, 12)
ws.set_column_width(4, 12)
ws.set_column_width(5, 14)

ws.freeze_panes(1, 0)
ws.autofit()

wb.close()
print("Created: pandas_output.xlsx")
