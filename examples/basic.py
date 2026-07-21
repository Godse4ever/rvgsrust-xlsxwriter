"""
Basic Example - All Core Features
=================================
Demonstrates: writing, formatting, borders, colors, fonts, merging
"""
from rvgsrust_xlsxwriter import Workbook

# Create workbook
wb = Workbook()
ws = wb.add_worksheet("Sales Report")

# Create formats
header_format = wb.add_format()
header_format.set_bold()
header_format.set_font_size(14)
header_format.set_background_color("#4472C4")  # Blue
header_format.set_font_color("white")
header_format.set_align("center")
header_format.set_border("thin")
header_format.set_border_color("#2F5496")

data_format = wb.add_format()
data_format.set_border("thin")
data_format.set_border_color("#D9D9D9")

money_format = wb.add_format()
money_format.set_num_format("$#,##0.00")
money_format.set_border("thin")
money_format.set_border_color("#D9D9D9")

total_format = wb.add_format()
total_format.set_bold()
total_format.set_background_color("#E7E6E6")
total_format.set_border("medium")
total_format.set_top_border("double")
total_format.set_num_format("$#,##0.00")

# Write merged header
ws.merge_range(0, 0, 0, 4, "QUARTERLY SALES REPORT", header_format)

# Write column headers
headers = ["Product", "Q1", "Q2", "Q3", "Total"]
for col, header in enumerate(headers):
    ws.write(2, col, header, header_format)

# Write data
data = [
    ["Widgets", 12000, 15000, 18000],
    ["Gadgets", 8000, 9500, 11000],
    ["Thingamajigs", 5000, 6200, 7500],
    ["Doohickeys", 3000, 4100, 5200],
]

for row_idx, row_data in enumerate(data):
    ws.write(3 + row_idx, 0, row_data[0], data_format)
    for col_idx, value in enumerate(row_data[1:], 1):
        fmt = money_format if col_idx > 0 else data_format
        ws.write(3 + row_idx, col_idx, value, fmt)

# Write formulas for totals
for row in range(3, 7):
    ws.write_formula(row, 4, f"=SUM(B{row+1}:D{row+1})", money_format)

# Write grand total row
ws.write(7, 0, "GRAND TOTAL", total_format)
for col in range(1, 5):
    col_letter = chr(ord('A') + col)
    ws.write_formula(7, col, f"=SUM({col_letter}4:{col_letter}7)", total_format)

# Set column widths
ws.set_column_width(0, 18)
for col in range(1, 5):
    ws.set_column_width(col, 14)

# Freeze header row
ws.freeze_panes(3, 0)

# Auto-fit (optional, but we already set widths)
ws.autofit()

# Close and save
wb.close("basic_example.xlsx")
print("Created: basic_example.xlsx")
