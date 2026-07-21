"""
Polars DataFrame Example
========================
"""
try:
    import polars as pl
except ImportError:
    print("Install polars: pip install polars")
    exit(1)

from rvgsrust_xlsxwriter import Workbook
from rvgsrust_xlsxwriter.dataframe import write_polars_dataframe

# Create sample Polars DataFrame
df = pl.DataFrame({
    "Employee": ["Alice Johnson", "Bob Smith", "Carol White", "David Brown", "Eve Davis"],
    "Department": ["Sales", "Engineering", "Marketing", "Sales", "Engineering"],
    "Salary": [75000.0, 95000.0, 65000.0, 82000.0, 105000.0],
    "Bonus": [5000.0, 8000.0, 3000.0, 6000.0, 10000.0],
    "Start Date": ["2020-01-15", "2019-03-22", "2021-06-01", "2018-11-10", "2020-09-05"],
})

# Create workbook
wb = Workbook("polars_output.xlsx")
ws = wb.add_worksheet("Employees")

# Header format
header_fmt = wb.add_format()
header_fmt.set_bold()
header_fmt.set_font_size(11)
header_fmt.set_background_color("#4472C4")
header_fmt.set_font_color("white")
header_fmt.set_align("center")
header_fmt.set_border("thin")

# Money format
money_fmt = wb.add_format()
money_fmt.set_num_format("$#,##0")
money_fmt.set_border("thin")
money_fmt.set_border_color("#D9D9D9")

# Department colors
dept_colors = {
    "Sales": wb.add_format().set_background_color("#E2EFDA").set_border("thin"),
    "Engineering": wb.add_format().set_background_color("#DDEBF7").set_border("thin"),
    "Marketing": wb.add_format().set_background_color("#FCE4D6").set_border("thin"),
}

# Column formats
col_formats = {
    "Salary": money_fmt,
    "Bonus": money_fmt,
}

# Write DataFrame
write_polars_dataframe(ws, df, row=0, col=0, header_format=header_fmt, column_formats=col_formats)

# Apply department colors manually (row by row)
for row_idx in range(len(df)):
    dept = df[row_idx, "Department"]
    if dept in dept_colors:
        ws.write(row_idx + 1, 1, dept, dept_colors[dept])

# Add total row
total_row = len(df) + 1
total_fmt = wb.add_format()
total_fmt.set_bold()
total_fmt.set_background_color("#C6E0B4")
total_fmt.set_border("medium")
total_fmt.set_top_border("double")
total_fmt.set_num_format("$#,##0")

ws.write(total_row, 0, "TOTAL", total_fmt)
ws.write_formula(total_row, 2, f"=SUM(C2:C{len(df)+1})", total_fmt)
ws.write_formula(total_row, 3, f"=SUM(D2:D{len(df)+1})", total_fmt)

# Set column widths
ws.set_column_width(0, 18)
ws.set_column_width(1, 15)
ws.set_column_width(2, 14)
ws.set_column_width(3, 14)
ws.set_column_width(4, 14)

# Freeze panes
ws.freeze_panes(1, 0)

# Auto-fit
ws.autofit()

wb.close()
print("Created: polars_output.xlsx")
