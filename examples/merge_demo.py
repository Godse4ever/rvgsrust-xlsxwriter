"""
Cell Merging Demo
=================
Shows all merge_range capabilities
"""
from rvgsrust_xlsxwriter import Workbook

wb = Workbook("merge_demo.xlsx")
ws = wb.add_worksheet("Merging")

# Title
title_fmt = wb.add_format()
title_fmt.set_bold()
title_fmt.set_font_size(16)
title_fmt.set_font_color("white")
title_fmt.set_background_color("#203764")
title_fmt.set_align("center")
title_fmt.set_border("medium")

ws.merge_range(0, 0, 0, 6, "CELL MERGING DEMONSTRATION", title_fmt)

# Section headers
section_fmt = wb.add_format()
section_fmt.set_bold()
section_fmt.set_font_size(12)
section_fmt.set_background_color("#4472C4")
section_fmt.set_font_color("white")

# Horizontal merge
ws.merge_range(2, 0, 2, 6, "Horizontal Merge: 7 columns wide", section_fmt)

# Vertical merge
ws.write(4, 0, "Vertical:", wb.add_format().set_bold())
vert_fmt = wb.add_format()
vert_fmt.set_background_color("#E7E6E6")
vert_fmt.set_align("center")
vert_fmt.set_vertical_align("center")
vert_fmt.set_border("thin")
ws.merge_range(4, 1, 8, 1, "5 rows tall", vert_fmt)

# Grid merge
ws.write(4, 3, "Grid:", wb.add_format().set_bold())
grid_fmt = wb.add_format()
grid_fmt.set_background_color("#FFF2CC")
grid_fmt.set_align("center")
grid_fmt.set_vertical_align("center")
grid_fmt.set_border("medium")
ws.merge_range(4, 4, 7, 6, "4 rows x 3 columns", grid_fmt)

# Data with merged headers
data_header = wb.add_format()
data_header.set_bold()
data_header.set_background_color("#70AD47")
data_header.set_font_color("white")
data_header.set_border("thin")

data_cell = wb.add_format()
data_cell.set_border("thin")
data_cell.set_border_color("#D9D9D9")

# Quarter headers (merged)
ws.merge_range(10, 1, 10, 2, "Q1", data_header)
ws.merge_range(10, 3, 10, 4, "Q2", data_header)
ws.merge_range(10, 5, 10, 6, "Q3", data_header)

# Sub-headers
sub_headers = ["Revenue", "Profit", "Revenue", "Profit", "Revenue", "Profit"]
for col, header in enumerate(sub_headers, 1):
    ws.write(11, col, header, data_header)

# Data rows
products = ["Product A", "Product B", "Product C"]
for row, product in enumerate(products, 12):
    ws.write(row, 0, product, data_cell)
    for col in range(1, 7):
        ws.write(row, col, f"=RANDBETWEEN(1000,5000)", data_cell)

# Total row (merged label)
total_fmt = wb.add_format()
total_fmt.set_bold()
total_fmt.set_background_color("#C6E0B4")
total_fmt.set_border("medium")
total_fmt.set_top_border("double")

ws.merge_range(15, 0, 15, 1, "TOTAL", total_fmt)
for col in range(2, 7):
    col_letter = chr(ord('A') + col)
    ws.write_formula(15, col, f"=SUM({col_letter}13:{col_letter}15)", total_fmt)

# Set widths
for col in range(7):
    ws.set_column_width(col, 14)

wb.close()
print("Created: merge_demo.xlsx")
