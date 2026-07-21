"""
Formatting Demo - All Format Features
=====================================
Shows every formatting option available
"""
from rvgsrust_xlsxwriter import Workbook

wb = Workbook()
ws = wb.add_worksheet("Formats")

# Title
title = wb.add_format()
title.set_bold()
title.set_font_size(18)
title.set_font_color("#C00000")
title.set_align("center")
ws.merge_range(0, 0, 0, 5, "COMPLETE FORMATTING DEMO", title)

# Section: Font Styles
row = 2
ws.write(row, 0, "FONT STYLES", wb.add_format().set_bold().set_font_size(12))

styles = [
    ("Normal", wb.add_format()),
    ("Bold", wb.add_format().set_bold()),
    ("Italic", wb.add_format().set_italic()),
    ("Bold + Italic", wb.add_format().set_bold().set_italic()),
    ("Underline", wb.add_format().set_underline()),
    ("Bold + Underline", wb.add_format().set_bold().set_underline()),
]

for i, (label, fmt) in enumerate(styles):
    ws.write(row + 1 + i, 0, label)
    ws.write(row + 1 + i, 1, "Sample Text", fmt)

# Section: Font Colors
row = 9
ws.write(row, 0, "FONT COLORS", wb.add_format().set_bold().set_font_size(12))

colors = ["red", "green", "blue", "orange", "purple", "navy", "magenta", "brown"]
for i, color in enumerate(colors):
    fmt = wb.add_format()
    fmt.set_font_color(color)
    fmt.set_bold()
    ws.write(row + 1 + i, 0, color.title())
    ws.write(row + 1 + i, 1, f"This is {color} text", fmt)

# Section: Background Colors
row = 18
ws.write(row, 0, "BACKGROUND COLORS", wb.add_format().set_bold().set_font_size(12))

bg_colors = [
    ("Yellow", "#FFFF00"),
    ("Light Blue", "#DDEBF7"),
    ("Light Green", "#E2EFDA"),
    ("Light Red", "#FCE4D6"),
    ("Light Purple", "#E1D5E7"),
    ("Light Orange", "#FCE4D6"),
]

for i, (name, color) in enumerate(bg_colors):
    fmt = wb.add_format()
    fmt.set_background_color(color)
    fmt.set_border("thin")
    ws.write(row + 1 + i, 0, name)
    ws.write(row + 1 + i, 1, "Colored Cell", fmt)

# Section: Borders
row = 25
ws.write(row, 0, "BORDERS", wb.add_format().set_bold().set_font_size(12))

border_styles = ["thin", "medium", "thick", "dashed", "dotted", "double"]
for i, style in enumerate(border_styles):
    fmt = wb.add_format()
    fmt.set_border(style)
    fmt.set_border_color("#FF0000")
    ws.write(row + 1 + i, 0, style.title())
    ws.write(row + 1 + i, 1, "Bordered", fmt)

# Section: Individual Borders
row = 32
ws.write(row, 0, "INDIVIDUAL BORDERS", wb.add_format().set_bold().set_font_size(12))

individual = [
    ("Top Only", lambda f: f.set_top_border("thick")),
    ("Bottom Only", lambda f: f.set_bottom_border("thick")),
    ("Left Only", lambda f: f.set_left_border("thick")),
    ("Right Only", lambda f: f.set_right_border("thick")),
    ("Top + Bottom", lambda f: f.set_top_border("thick").set_bottom_border("thick")),
    ("All Different", lambda f: f.set_top_border("thin").set_bottom_border("medium").set_left_border("thick").set_right_border("double")),
]

for i, (label, apply) in enumerate(individual):
    fmt = wb.add_format()
    apply(fmt)
    fmt.set_border_color("#0000FF")
    ws.write(row + 1 + i, 0, label)
    ws.write(row + 1 + i, 1, "Bordered", fmt)

# Section: Alignment
row = 39
ws.write(row, 0, "ALIGNMENT", wb.add_format().set_bold().set_font_size(12))

alignments = [
    ("Left", "left"),
    ("Center", "center"),
    ("Right", "right"),
    ("Fill", "fill"),
    ("Justify", "justify"),
]

for i, (label, align) in enumerate(alignments):
    fmt = wb.add_format()
    fmt.set_align(align)
    fmt.set_border("thin")
    ws.write(row + 1 + i, 0, label)
    ws.write(row + 1 + i, 1, "Aligned Text", fmt)

# Section: Number Formats
row = 45
ws.write(row, 0, "NUMBER FORMATS", wb.add_format().set_bold().set_font_size(12))

num_formats = [
    ("General", None, 1234.567),
    ("2 Decimals", "0.00", 1234.567),
    ("Currency", "$#,##0.00", 1234.567),
    ("Percentage", "0.00%", 0.1234),
    ("Scientific", "0.00E+00", 1234.567),
    ("Fraction", "# ?/?", 1.25),
    ("Thousands", "#,##0", 1234567),
]

for i, (label, fmt_str, value) in enumerate(num_formats):
    fmt = wb.add_format()
    if fmt_str:
        fmt.set_num_format(fmt_str)
    ws.write(row + 1 + i, 0, label)
    ws.write(row + 1 + i, 1, value, fmt)

# Section: Merged Cells
row = 53
ws.write(row, 0, "MERGED CELLS", wb.add_format().set_bold().set_font_size(12))

merge_fmt = wb.add_format()
merge_fmt.set_bold()
merge_fmt.set_background_color("#7030A0")
merge_fmt.set_font_color("white")
merge_fmt.set_align("center")
merge_fmt.set_border("medium")

ws.merge_range(row + 1, 0, row + 1, 3, "This cell spans 4 columns", merge_fmt)
ws.merge_range(row + 2, 0, row + 3, 1, "This cell spans 2 rows and 2 columns", merge_fmt)

# Set column widths
ws.set_column_width(0, 20)
ws.set_column_width(1, 20)
ws.set_column_width(2, 15)
ws.set_column_width(3, 15)
ws.set_column_width(4, 15)
ws.set_column_width(5, 15)

wb.close("formatting_demo.xlsx")
print("Created: formatting_demo.xlsx")
