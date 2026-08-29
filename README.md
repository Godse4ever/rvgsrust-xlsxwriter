# RVGSRust-XLSXWriter

> **A Rust-powered XLSX library for Python, with a Pythonic API and Polars/Pandas/PyArrow support.**
>
> Built on the official [`rust_xlsxwriter`](https://github.com/jmcnamara/rust_xlsxwriter) crate. Not affiliated with the Python [`XlsxWriter`](https://xlsxwriter.readthedocs.io/) package (a separate, pure-Python project it shares a similar name and purpose with) -- this library's Python-facing API is inspired by it but is not a drop-in replacement; see [Quick Start](#quick-start) below for the real differences.

[![PyPI version](https://badge.fury.io/py/rvgsrust-xlsxwriter.svg)](https://pypi.org/project/rvgsrust-xlsxwriter/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Rust](https://img.shields.io/badge/Rust-1.85%2B-orange.svg)](https://www.rust-lang.org)
[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org)

---

## Why RVGSRust-XLSXWriter?

| Feature | `xlsxwriter` (Python) | `rustpy-xlsxwriter` | `rvgsrust-xlsxwriter` |
|---------|----------------------|---------------------|----------------------|
| Speed (100k rows) | 1× (baseline) | — | **6–8× faster** ⚡ |
| **GIL released during `save()`** | ❌ blocks | ❌ blocks | ✅ **Yes** (~7–48ms stall vs ~1,270–5,980ms on a wide workload) |
| **Cell Merging** | ✅ Yes | ❌ No | ✅ **Yes** |
| **Full Format API** | ✅ Yes | ⚠️ Limited | ✅ **Complete** |
| **Borders (all sides)** | ✅ Yes | ⚠️ Basic | ✅ **All sides + colors** |
| **Cell/Font Colors** | ✅ Yes | ✅ Yes | ✅ **Yes** |
| **Formulas** | ✅ Yes | ❌ No | ✅ **Yes** |
| **Hyperlinks (text/tip)** | ✅ Yes | ❌ No | ✅ **Yes** |
| **Date/Time** | ✅ Yes | ✅ Yes | ✅ **Yes** |
| **Images** | ✅ Yes | ❌ No | ✅ **Yes** |
| **Freeze Panes** | ✅ Yes | ✅ Yes | ✅ **Yes** |
| **Sheet Protection** | ✅ Yes | ✅ Yes | ✅ **Yes** |
| **Worksheet Tables** | ✅ Yes | ❌ No | ✅ **Yes** |
| **Rich Strings** | ✅ Yes | ❌ No | ✅ **Yes** |
| **Polars/Pandas/PyArrow** | Manual | ✅ Zero-copy | ✅ **Zero-copy** |
| **Charts** | ✅ Yes | ❌ No | ✅ **Yes** |
| **Conditional Format** | ✅ Yes | ❌ No | ✅ **Yes** |

**We win on completeness.** Both `rustpy-xlsxwriter` and `rvgsrust-xlsxwriter` use the same Rust core, but we expose *every* feature so you never have to fall back to Python.

The **6–8×** figure is measured against pure-Python `xlsxwriter` at 100,000 rows × 8 columns (tall, narrow shape) — see [Performance & Benchmarks](#performance--benchmarks) for full numbers and methodology. Against another Rust-backed writer with the same `rust_xlsxwriter` core (`rustpy-xlsxwriter`), an independent evaluation on a wide workload (76,480 columns × 1,261 rows, 14 sheets) measured roughly **2× faster on total write time and ~3× faster on the save phase specifically**. Where this library won decisively on that same wide workload was responsiveness during `save()`: a background thread polling every 5ms saw a worst-case stall of **7–48ms** across runs, versus **1,270–5,980ms** for the comparison library — the GIL is released during save, so a GUI or async app stays responsive instead of freezing. That's arguably a stronger selling point than raw throughput for desktop-app use cases, even though it's less prominent in the numbers above. Machine load moved absolute timings by 3–4× across runs on identical work in that evaluation; the ratios above held up better than the absolutes, which is why they're given as ranges/ratios rather than single numbers.

---

## Installation

```bash
pip install rvgsrust-xlsxwriter
```

### Optional Dependencies

```bash
# For Polars DataFrame support (recommended)
pip install rvgsrust-xlsxwriter[polars]

# For Pandas DataFrame support
pip install rvgsrust-xlsxwriter[pandas]

# For everything
pip install rvgsrust-xlsxwriter[all]
```

---

## Quick Start

```python
from rvgsrust_xlsxwriter import Workbook

# Create workbook -- pass the path up front so close() needs no argument later
wb = Workbook("report.xlsx")
ws = wb.add_worksheet("Sales")

# Create a rich format
header = wb.add_format()
header.set_bold()
header.set_font_size(14)
header.set_background_color("#4472C4")
header.set_font_color("white")
header.set_align("center")
header.set_border("thin")

# Write with format
ws.write(0, 0, "Product", header)
ws.write(0, 1, "Q1 Sales", header)
ws.write(0, 2, "Q2 Sales", header)

# Merge cells
ws.merge_range(0, 3, 0, 4, "Total", header)

# Write data
data_format = wb.add_format()
data_format.set_border("thin")
data_format.set_border_color("#D9D9D9")

ws.write(1, 0, "Widgets", data_format)
ws.write(1, 1, 12000, data_format)
ws.write(1, 2, 15000, data_format)
ws.write_formula(1, 3, "=B2+C2", data_format)

# Auto-fit columns
ws.autofit()

# Save -- close() with no argument uses the path given to the constructor.
# Passing a path here still works too, and overrides the constructor path.
wb.close()
```

The context-manager form does the same thing automatically on `__exit__`:

```python
with Workbook("report.xlsx") as wb:
    ws = wb.add_worksheet("Sales")
    ws.write(0, 0, "Product")
# saved to report.xlsx here, close() called for you
```

`save_to_buffer()` skips the file entirely and returns the xlsx as
`bytes` -- useful for a web response or an in-memory pipeline, with no
temp-file round trip:

```python
wb = Workbook()
ws = wb.add_worksheet()
ws.write(0, 0, "In memory")
xlsx_bytes = wb.save_to_buffer()
# e.g. return Response(xlsx_bytes, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
```

---

## Features

### Bulk Writes (recommended for large datasets)

```python
data = [
    {"Name": "Alice", "Department": "Engineering", "Salary": 95000},
    {"Name": "Bob", "Department": "Sales", "Salary": 72000},
    # ... thousands more rows
]

header_fmt = wb.add_format()
header_fmt.set_bold()

# One call for the entire dataset, instead of one write() call per cell
ws.write_records(0, 0, data, header_format=header_fmt)

# Column order/subset and header row are both optional:
ws.write_records(0, 0, data, headers=["Name", "Salary"], write_header=False)
```

### Complete Formatting

```python
fmt = wb.add_format()

# Font
fmt.set_bold()
fmt.set_italic()
fmt.set_underline()
fmt.set_font_name("Calibri")
fmt.set_font_size(12)
fmt.set_font_color("#FF0000")      # Hex or named colors

# Background
fmt.set_background_color("#FFFF00")
fmt.set_pattern("solid")           # Pattern fills

# Borders (all sides or individual)
fmt.set_border("thin")
fmt.set_border_color("#000000")
fmt.set_top_border("thick")
fmt.set_bottom_border("double")
fmt.set_left_border("dashed")
fmt.set_right_border("dotted")

# Alignment
fmt.set_align("center")            # left, center, right, fill, justify
fmt.set_vertical_align("vcenter")  # top, vcenter, bottom
fmt.set_text_wrap()
fmt.set_rotation(45)

# Numbers
fmt.set_num_format("$#,##0.00")
fmt.set_num_format("0.00%")
fmt.set_num_format("yyyy-mm-dd")

# Less common, but all there: quote prefix, hyperlink style, checkboxes,
# font family/charset/script, reading direction -- and every set_*()
# above has an unset_*() inverse.
fmt.set_quote_prefix()             # force text display, e.g. "007"
fmt.set_hyperlink()                # Excel's built-in hyperlink look
fmt.set_checkbox()                 # for a boolean cell
fmt.set_font_script("superscript") # "none", "superscript", "subscript"
fmt.set_reading_direction(1)       # 0=context, 1=left-to-right, 2=right-to-left
fmt.unset_bold()                   # reverse any set_*() above

# All setters return self, so they chain:
wb.add_format().set_bold().set_font_size(12).set_background_color("#4472C4")
```

### Cell Merging

```python
# Merge range: first_row, first_col, last_row, last_col
merge_fmt = wb.add_format()
merge_fmt.set_bold()
merge_fmt.set_background_color("#7030A0")
merge_fmt.set_font_color("white")
merge_fmt.set_align("center")

ws.merge_range(0, 0, 0, 4, "Quarterly Report", merge_fmt)
ws.merge_range(2, 0, 4, 0, "Vertical Label", merge_fmt)
```

### Formulas & Hyperlinks

```python
# Formula
ws.write_formula(5, 0, "=SUM(A1:A4)")
ws.write_formula(5, 1, "=AVERAGE(B1:B4)", money_format)

# Hyperlink
ws.write_url(6, 0, "https://github.com/Godse4ever/rvgsrust-xlsxwriter")
```

### Dates & Times

```python
# Write date
ws.write_date(0, 0, 2024, 1, 15)  # year, month, day

# Write datetime
ws.write_datetime(0, 1, 2024, 1, 15, 14, 30, 0)  # + hour, min, sec

# Use date format
date_fmt = wb.add_format()
date_fmt.set_num_format("yyyy-mm-dd")
ws.write_date(0, 0, 2024, 1, 15, date_fmt)
```

### Images

```python
ws.insert_image(0, 0, "logo.png")
```

### Sheet Operations

```python
ws.set_column_width(0, 20.0)       # Set column width
ws.set_row_height(0, 30.0)         # Set row height
ws.freeze_panes(1, 0)              # Freeze header row
ws.hide()                          # Hide sheet
ws.set_tab_color("red")            # Color the sheet tab
ws.protect("password")             # Password protect
ws.autofit()                       # Auto-fit columns
ws.autofilter(0, 0, 100, 4)        # Add filter dropdowns over A1:E101
```

### Defined Names

```python
# Workbook-global name
wb.define_name("SalesTotal", "Sheet1!$B$2:$B$100")

# Sheet-scoped name (only visible/usable within that sheet)
wb.define_name("Sheet1!LocalTotal", "Sheet1!$B$2:$B$100")
```

### Tables

Write the data first with the normal `write_*()` methods, then call
`add_table()` over the range it occupies -- matches `rust_xlsxwriter`'s
own usage pattern:

```python
from rvgsrust_xlsxwriter import Table, TableColumn

ws.write_row(0, 0, ["Product", "Q1", "Q2"])
ws.write_column(1, 0, ["Apples", "Pears"])
ws.write_column(1, 1, [10, 20])
ws.write_column(1, 2, [15, 25])

columns = [
    TableColumn().set_header("Product").set_total_label("Total"),
    TableColumn().set_header("Q1").set_total_function("sum"),
    TableColumn().set_header("Q2").set_total_function("sum"),
]

table = (
    Table()
    .set_columns(columns)
    .set_name("SalesTable")
    .set_style("medium9")
    .set_total_row(True)
    .set_banded_rows(True)
)

ws.add_table(0, 0, 3, 2, table)
```

`TableColumn` also supports `set_formula()` (per-row calculated
column, using Excel's structured-reference syntax like
`"SUM(SalesTable[@[Q1]:[Q2]])"`), `set_format()`/`set_header_format()`
(per-column cell/header formatting), and `set_total_function()` accepts
either a built-in keyword (`"sum"`, `"average"`, `"count"`,
`"count_numbers"`, `"max"`, `"min"`, `"stddev"`, `"var"`, `"none"`) or,
for anything else, treats the string as a custom formula.
`Table.set_style()` accepts `"none"`, `"light1"`-`"light21"`,
`"medium1"`-`"medium28"`, or `"dark1"`-`"dark11"`.

---

## Multithreading

**Cross-sheet parallelism is already active, automatically, for every multi-sheet workbook.** The underlying `rust_xlsxwriter` crate spawns one real OS thread per worksheet to assemble its XML when you call `wb.close()`:

```rust
// From rust_xlsxwriter's own source (src/packager.rs) -- unconditional,
// no feature flag, runs on every save() with 2+ worksheets:
thread::scope(|scope| {
    for worksheet in &mut workbook.worksheets {
        scope.spawn(|| {
            worksheet.assemble_xml_file();
        });
    }
});
```

| Worksheets | Threads spawned |
|---|---|
| 1 | 1 |
| 5 | 5 |

This is automatic and scales to one thread per worksheet.

---

## Polars / Pandas Integration

`write_polars_dataframe()` / `write_pandas_dataframe()` automatically use the zero-copy Arrow path under the hood whenever possible.

### Polars (Zero-Copy)

```python
import polars as pl
from rvgsrust_xlsxwriter import Workbook
from rvgsrust_xlsxwriter.dataframe import write_polars_dataframe

df = pl.DataFrame({
    "Name": ["Alice", "Bob", "Charlie"],
    "Sales": [100.5, 200.75, 150.0],
})

wb = Workbook()
ws = wb.add_worksheet()

header_fmt = wb.add_format()
header_fmt.set_bold()
header_fmt.set_background_color("#4472C4")
header_fmt.set_font_color("white")

write_polars_dataframe(ws, df, header_format=header_fmt)
ws.autofit()
wb.close("output.xlsx")
```

### Pandas

```python
import pandas as pd
from rvgsrust_xlsxwriter import Workbook
from rvgsrust_xlsxwriter.dataframe import write_pandas_dataframe

df = pd.DataFrame({
    "Product": ["Widget", "Gadget"],
    "Price": [9.99, 19.99],
})

wb = Workbook()
ws = wb.add_worksheet()

write_pandas_dataframe(ws, df)
wb.close("output.xlsx")
```

### Direct Arrow Access

```python
import pyarrow as pa

table = pa.table({"id": [1, 2, 3], "name": ["Alice", "Bob", "Carol"]})

ws.write_dataframe(0, 0, table, header_format=header_fmt)
```

**Current type support:** `int8`/`int16`/`int32`/`int64`, `uint8`/`uint16`/`uint32`/`uint64`, `float32`/`float64`, `string`/`utf8`, `large_utf8`, `utf8view` (Polars default), `bool`, `date32`/`date64`, and `timestamp` in all four units (`s`/`ms`/`us`/`ns` -- `ns` is pandas' default `datetime64[ns]`).

Date and timestamp columns are written as real Excel dates with a `yyyy-mm-dd` / `yyyy-mm-dd hh:mm:ss` number format applied automatically, rather than as raw serial numbers. Timezone-aware timestamps are written as UTC wall-clock time (Excel has no timezone concept) and emit a `UserWarning`; call `.dt.tz_convert(None)` first if you want to choose the offset yourself.

---

## Column, row and range formats

```python
from rvgsrust_xlsxwriter import Workbook, Format

wb = Workbook()
ws = wb.add_worksheet()

money = Format()
money.set_num_format("#,##0.00")

ws.set_column_format(2, money)              # whole column
ws.set_column_range_format(3, 5, money)     # columns D-F
ws.set_column_range_width(3, 5, 14.0)       # width across the same range
ws.set_row_format(0, money)                 # whole row
ws.set_cell_format(0, 0, money)             # single cell
ws.set_range_format(1, 0, 10, 4, money)     # a rectangular range
ws.clear_cell_format(0, 0)                  # remove a cell's format, keep its value

# Border around the outside of a range, with interior styling too --
# builds the up-to-9 per-position format combinations (corners, edges,
# interior) internally instead of you tracking them by hand.
border = Format()
border.set_border("thin")
ws.set_range_format_with_border(1, 0, 10, 4, money, border)
```

Column and row formats apply to cells that don't carry a format of their
own, which makes them the way to reformat data written by
`write_dataframe`.

### Applying a format across many columns

Putting a border (or any format) on a wide data grid has five routes, with
very different costs:

| API | Cost | Coverage |
|---|---|---|
| `set_column_range_format(first, last, fmt)` | Column-level style, negligible | Applies only where the cell has no format of its own — cells with a date or custom number format keep their own and show no border |
| `write_dataframe(..., column_formats={...})` | Merged into each written cell | Full, including number-formatted columns. Nulls are always skipped, never materialized as formatted blank cells |
| `set_range_format_with_border(r1, c1, r2, c2, cell_fmt, border_fmt)` | Per cell, built internally | Full -- purpose-built for "fill/style the range and put a border around the outside" in one call |
| `set_range_format(r1, c1, r2, c2, fmt)` | Per cell | Full, but a 25,000×76,000 range is ~1.9 billion cells — not viable |
| `set_cell_format(row, col, fmt)` | Per cell | Full, for targeted use only |

Practical guidance:

- Uniform styling across a grid, no number-formatted columns →
  `set_column_range_format`. Cheapest by a wide margin.
- Grids containing date, datetime, or custom-numeric columns →
  `column_formats` on `write_dataframe`, because OOXML gives a cell's own
  format precedence over the column format — a border applied only at
  column level will be silently missing on exactly those columns.
- Setting a uniform width across many columns → `set_column_range_width`,
  not a `set_column_width` loop. On a wide export the loop cost several
  seconds purely in FFI overhead.
- A single border wrapped around a filled range (a summary table, a
  legend box) → `set_range_format_with_border`. It's the one built
  specifically for that shape, and it does the corner/edge/interior
  format bookkeeping so you don't have to.

## Page setup and printing

```python
ws.set_landscape()
ws.set_paper_size(9)                                  # 9 = A4, 1 = Letter
ws.set_margins(0.5, 0.5, 0.75, 0.75, 0.3, 0.3)        # inches
ws.set_print_area(0, 0, 49, 7)
ws.set_repeat_rows(0, 0)                              # header row on every page
ws.set_print_fit_to_pages(1, 0)                       # 1 page wide, any height
ws.set_print_gridlines(True)
ws.set_print_center_horizontally(True)
ws.set_page_breaks([25, 50])
```

Also `set_portrait`, `set_page_order`, `set_repeat_columns`,
`set_print_scale`, `set_vertical_page_breaks`, `set_print_headings`,
`set_print_center_vertically`, `set_print_black_and_white`,
`set_print_draft` and `set_print_first_page_number`.

These set worksheet metadata rather than writing cells, so they are
unaffected by `constant_memory` row ordering even where they take row
numbers.

### Headers and footers

```python
ws.set_header("&CPage &[Page] of &[Pages]")
ws.set_footer("&LConfidential&RGenerated by rvgsrust-xlsxwriter")
```

`&L`, `&C`, `&R` align text to the left/center/right section.
`&[Page]`, `&[Pages]`, `&[File]`, `&[Tab]`, `&[Date]` are Excel
placeholders expanded when the file is opened. This library doesn't
validate the string, but it isn't a raw pass-through either:
`rust_xlsxwriter` normalizes the bracket syntax to the older
single-letter codes before writing (`&[Page]` becomes `&P`, `&[Pages]`
becomes `&N`, `&[Tab]` becomes `&A`) -- both spellings are valid Excel
codes and render identically once opened, the file on disk just won't
contain whichever bracket form you typed. Excel's 256-character limit
on the combined string (including control characters) isn't checked at
write time; an oversized string is silently truncated by Excel rather
than rejected. Full syntax reference:
https://rustxlsxwriter.github.io/worksheet/headers.html

## Conditional Formatting

Twelve rule types are supported. Build a rule object, then attach it to a
range with `Worksheet.add_conditional_format(first_row, first_col,
last_row, last_col, rule)`.

```python
from rvgsrust_xlsxwriter import (
    Workbook, Format, ConditionalFormatCell, ConditionalFormat3ColorScale,
    ConditionalFormatDataBar,
)

wb = Workbook()
ws = wb.add_worksheet()
for i, v in enumerate([10, 20, 30, 40, 50]):
    ws.write(i, 0, v)

bad = Format()
bad.set_background_color("#FFC7CE")

high = ConditionalFormatCell()
high.set_rule_greater_than(35)
high.set_format(bad)
ws.add_conditional_format(0, 0, 4, 0, high)

scale = ConditionalFormat3ColorScale()
scale.set_minimum_color("#F8696B")
scale.set_midpoint_color("#FFEB84")
scale.set_maximum_color("#63BE7B")
ws.add_conditional_format(0, 1, 4, 1, scale)

bar = ConditionalFormatDataBar()
bar.set_fill_color("#638EC6")
bar.set_direction("left_to_right")
ws.add_conditional_format(0, 2, 4, 2, bar)

wb.close("cf.xlsx")
```

| Class | Rule setters |
|---|---|
| `ConditionalFormatCell` | `set_rule_greater_than`, `..._less_than`, `..._equal_to`, `..._not_equal_to`, `..._greater_than_or_equal_to`, `..._less_than_or_equal_to`, `set_rule_between`, `set_rule_not_between` |
| `ConditionalFormatBlank` | `invert()` for non-blank |
| `ConditionalFormatDuplicate` | `invert()` for unique |
| `ConditionalFormatError` | `invert()` for non-error |
| `ConditionalFormatFormula` | `set_rule("=$A1>50")` |
| `ConditionalFormatAverage` | `set_rule("above" \| "below" \| "equal_or_above" \| "equal_or_below" \| "{1,2,3}_std_dev_above" \| "..._below")` |
| `ConditionalFormatTop` | `set_rule("top" \| "bottom" \| "top_percent" \| "bottom_percent", n)` |
| `ConditionalFormatText` | `set_rule("contains" \| "does_not_contain" \| "begins_with" \| "ends_with", text)` |
| `ConditionalFormatDate` | `set_rule("yesterday" \| "today" \| "tomorrow" \| "last_7_days" \| "last_week" \| "this_week" \| "next_week" \| "last_month" \| "this_month" \| "next_month")` |
| `ConditionalFormat2ColorScale` | `set_minimum/set_maximum(type, value)`, `set_minimum_color/set_maximum_color` |
| `ConditionalFormat3ColorScale` | as above plus `set_midpoint`, `set_midpoint_color` |
| `ConditionalFormatDataBar` | `set_minimum/set_maximum`, `set_fill_color`, `set_border_color`, `set_negative_fill_color`, `set_negative_border_color`, `set_axis_color`, `set_solid_fill`, `set_border_off`, `set_bar_only`, `set_direction`, `set_axis_position`, `use_classic_style` |

Rule-type strings for `set_minimum`/`set_maximum`/`set_midpoint` are
`automatic`, `lowest`/`min`, `highest`/`max`, `number`, `percent`,
`percentile`, `formula`. Values may be numbers, or strings when the type
is `formula`. Every class also has `set_multi_range(range)` and
`set_stop_if_true(enable)`.

Color scales and data bars have no `set_format()`: Excel renders those
from the scale or bar definition itself rather than from a format record.
Icon sets are the same way.

### Icon Sets

```python
from rvgsrust_xlsxwriter import ConditionalFormatIconSet, ConditionalFormatCustomIcon

cf = ConditionalFormatIconSet()
cf.set_icon_type("three_traffic_lights")   # the default; 20 styles total
cf.reverse_icons(True)                     # lowest value gets the "highest" icon
cf.show_icons_only(True)                   # hide the cell's own value
ws.add_conditional_format(0, 0, 9, 0, cf)
```

Override individual thresholds/icons/direction with `set_icons()`,
passing one `ConditionalFormatCustomIcon` per icon in the set (3, 4, or
5, matching `set_icon_type`):

```python
icons = [ConditionalFormatCustomIcon() for _ in range(3)]
icons[0].set_rule("percent", 0)
icons[1].set_rule("percent", 33)
icons[2].set_rule("percent", 67)
icons[2].set_icon_type("five_boxes", 4)   # borrow one icon from a different set
icons[2].set_greater_than(True)           # ">" instead of Excel's default ">="

cf = ConditionalFormatIconSet()
cf.set_icons(icons)
```

`ConditionalFormatIconSet()` is valid on its own with no other calls --
internally it already calls `set_icon_type("three_traffic_lights")`,
working around a real `rust_xlsxwriter` gotcha where a freshly
constructed icon set otherwise fails validation (its default icon-rules
list starts empty, and Excel requires exactly 3/4/5 entries matching
the type).

## Data Validation

```python
from rvgsrust_xlsxwriter import DataValidation

wb = Workbook()
ws = wb.add_worksheet()

# Dropdown from a fixed list -- the most common case
status = DataValidation()
status.allow_list_strings(["Pending", "In Progress", "Done"])
status.show_input_message(True)
status.set_input_title("Status")
status.set_input_message("Pick one from the list")
ws.add_data_validation(0, 0, 99, 0, status)

# Dropdown sourced from a range written elsewhere -- no 255-char limit,
# unlike allow_list_strings()
ws.write_column(0, 5, ["Small", "Medium", "Large"])
size = DataValidation()
size.allow_list_formula("F1:F3")
ws.add_data_validation(0, 1, 99, 1, size)

# Numeric range, with an error dialog on invalid input
rating = DataValidation()
rating.allow_whole_number("between", 1, 5)
rating.show_error_message(True)
rating.set_error_title("Invalid rating")
rating.set_error_message("Enter a number from 1 to 5")
rating.set_error_style("stop")  # "stop", "warning", or "information"
ws.add_data_validation(0, 2, 99, 2, rating)

# Arbitrary formula rule
cross_check = DataValidation()
cross_check.allow_custom("=A1<>\"\"")
ws.add_data_validation(0, 3, 99, 3, cross_check)
```

`allow_whole_number`, `allow_decimal_number`, and `allow_text_length`
all take the same 8 comparison types: `equal_to`, `not_equal_to`,
`greater_than`, `greater_than_or_equal_to`, `less_than`,
`less_than_or_equal_to`, `between`, `not_between` (the last two take a
second value). `set_multi_range(range)` replaces the range given to
`add_data_validation()` entirely rather than adding to it -- put every
range you want validated into that one call. `allow_any_value()` clears
a rule while keeping the input/error messages, useful for a validation
that's purely informational.

Not yet implemented: `allow_date`/`allow_time` range rules, and the
cell-reference formula variants of the numeric/text-length/date/time
rules (comparing against a cell instead of a literal value) -- see
MISSING.md.

## Row and Column Grouping

Collapsible outline sections, the kind used for expandable summary
reports:

```python
ws.write(0, 0, "Region")
ws.write(1, 0, "  North")
ws.write(2, 0, "  South")
ws.write(3, 0, "Total")

ws.group_rows(1, 2)              # rows 1-2 collapse under row 3
ws.group_symbols_above(True)     # [+] toggle above the group, not below

ws.group_columns(1, 3)           # same idea for columns
ws.group_symbols_to_left(True)   # toggle to the left, not the right
```

`group_rows_collapsed`/`group_columns_collapsed` start the group
already collapsed. Groups can nest up to Excel's 7-level limit --
group the outer range first, then the inner range, the normal way
multi-level outlines are built.

**Known limitation:** when a worksheet is created with
`constant_memory=True`, `group_rows()`/`group_rows_collapsed()` don't
apply per-row grouping to the output -- verified against actual
output, not assumed. The call succeeds and the sheet records the
correct maximum outline level, but Excel won't show the per-row
collapse/expand behavior. This appears to be a `constant_memory`
streaming limitation upstream, not something fixable from this
binding without writing worksheet XML directly. `group_columns()` is
unaffected, since columns aren't part of the row-streaming mechanism.

## Sparklines

```python
from rvgsrust_xlsxwriter import Workbook, Sparkline

wb = Workbook()
ws = wb.add_worksheet()
for r in range(5):
    for c in range(5):
        ws.write(r, c, (r + 1) * (c + 1))

sp = Sparkline()
sp.set_range("Sheet1!A1:E1")
sp.set_type("column")
sp.show_high_point(True)
sp.set_sparkline_color("#638EC6")
ws.add_sparkline(0, 5, sp)

# A grouped sparkline shares one set of options across a range. Its data
# range must be 2D, one row per sparkline in the group.
group = Sparkline()
group.set_range("Sheet1!A1:E5")
group.set_group_max(True)
ws.add_sparkline_group(0, 6, 4, 6, group)

wb.close("sparklines.xlsx")
```

Types are `line` (default), `column`, and `win_lose`. Point markers:
`show_high_point`, `show_low_point`, `show_first_point`,
`show_last_point`, `show_negative_points`, `show_markers`, `show_axis`,
`show_hidden_data`. Colors: `set_sparkline_color`, plus
`set_high_point_color`, `set_low_point_color`, `set_first_point_color`,
`set_last_point_color`, `set_negative_points_color`, `set_markers_color`.
Scaling: `set_line_weight`, `set_custom_max`, `set_custom_min`,
`set_group_max`, `set_group_min`, `set_style`. Also `set_date_range`,
`set_right_to_left`, `set_column_order`, and `show_empty_cells_as` with
`gaps`, `zero` or `connected`.

## Charts

```python
from rvgsrust_xlsxwriter import Workbook, Chart, ChartSeries

wb = Workbook()
ws = wb.add_worksheet()
for i, (label, value) in enumerate([("Jan", 10), ("Feb", 40), ("Mar", 20)]):
    ws.write(i, 0, label)
    ws.write(i, 1, value)

series = ChartSeries()
series.set_categories("Sheet1!$A$1:$A$3")
series.set_values("Sheet1!$B$1:$B$3")
series.set_name("Revenue")

chart = Chart("column")
chart.push_series(series)
chart.set_title_name("Revenue by month")
chart.set_x_axis_name("Month")
chart.set_y_axis_name("USD")
chart.set_y_axis_min(0.0)
chart.set_legend_position("bottom")
ws.insert_chart(0, 3, chart, 10, 10)   # last two args are pixel offsets

wb.close("chart.xlsx")
```

The 23 chart types are `area`, `area_stacked`, `area_percent_stacked`,
`bar`, `bar_stacked`, `bar_percent_stacked`, `column`, `column_stacked`,
`column_percent_stacked`, `doughnut`, `line`, `line_stacked`,
`line_percent_stacked`, `pie`, `radar`, `radar_with_markers`,
`radar_filled`, `scatter`, `scatter_straight`,
`scatter_straight_with_markers`, `scatter_smooth`,
`scatter_smooth_with_markers`, `stock`.

Series are built separately and attached with `Chart.push_series()`.
`ChartSeries` supports `set_values`, `set_categories`, `set_name`,
`set_secondary_axis`, `set_overlap`, `set_gap`, `set_smooth`,
`set_invert_if_negative`, `set_invert_if_negative_color`,
`delete_from_legend` and `set_point_colors`.

Axes, titles and legends are **not** separate objects, because their
constructors are `pub(crate)` upstream. They are flattened onto `Chart` as
`set_title_name` / `set_title_hidden` / `set_title_overlay`,
`set_x_axis_*` and `set_y_axis_*` (`name`, `min`, `max`, `major_unit`,
`minor_unit`, `log_base`, `num_format`, `hidden`, `reverse`,
`major_gridlines`, `minor_gridlines`, plus `date_axis` / `text_axis` on x),
and `set_legend_position` / `set_legend_hidden` / `set_legend_overlay`.

Legend positions are `right`, `left`, `top`, `bottom`, `top_right`. There
is no overlay position; use `set_legend_overlay(True)` alongside one of
those.

Note that `set_title_hidden()`, `set_legend_hidden()`,
`set_x_axis_reverse()`, `set_y_axis_reverse()`, `show_hidden_data()` and
`show_na_as_empty_cell()` take no arguments, matching upstream.

Chart formatting (`ChartFormat`, `ChartFont`) and series decorations
(`ChartMarker`, `ChartTrendline`, `ChartDataLabel`) are not exposed yet.

### Chart formatting

`ChartFormat` handles fills, lines and borders; `ChartFont` handles text.

```python
from rvgsrust_xlsxwriter import Chart, ChartSeries, ChartFormat, ChartFont

bar_style = ChartFormat()
bar_style.set_fill_color("#4472C4")
bar_style.set_border_color("#000000")
bar_style.set_border_width(1.0)

title_font = ChartFont()
title_font.set_bold()
title_font.set_size(16.0)
title_font.set_color("#333333")

series = ChartSeries()
series.set_values("Sheet1!$B$1:$B$5")
series.set_format(bar_style)

chart = Chart("column")
chart.push_series(series)
chart.set_title_name("Styled")
chart.set_title_font(title_font)
```

`ChartLine` and `ChartSolidFill` are not separate classes. Upstream they
exist only to be handed to `ChartFormat`, so they are flattened into it:
`set_line_color`, `set_line_width`, `set_line_dash_type`,
`set_line_transparency`, `set_line_hidden`, `set_no_line`, the same six
under `set_border_*` / `set_no_border`, and `set_fill_color`,
`set_fill_transparency`, `set_no_fill`. Line state is kept per format
object, so successive calls compose.

Dash types are `solid`, `round_dot`, `square_dot`, `dash`, `dash_dot`,
`long_dash`, `long_dash_dot`, `long_dash_dot_dot`.

`ChartFont` supports `set_bold`, `unset_bold`, `set_default_bold`,
`set_italic`, `set_underline`, `set_strikethrough`, `set_color`,
`set_name`, `set_size`, `set_rotation`, `set_right_to_left`,
`set_pitch_family` and `set_character_set`. Note `set_bold`, `set_italic`,
`set_underline` and `set_strikethrough` take no arguments, matching
upstream.

Attachment points: `ChartSeries.set_format`, and on `Chart`
`set_title_font` / `set_title_format`, `set_x_axis_font` /
`set_x_axis_name_font` / `set_x_axis_format` / `set_x_axis_name_format`,
the same four for `y_axis`, and `set_legend_font` / `set_legend_format`.

Pattern and gradient fills are not exposed yet.

### Series markers, trendlines and data labels

```python
from rvgsrust_xlsxwriter import (
    Chart, ChartSeries, ChartMarker, ChartTrendline, ChartDataLabel,
)

series = ChartSeries()
series.set_values("Sheet1!$B$1:$B$5")

marker = ChartMarker()
marker.set_type("circle")
marker.set_size(7)
series.set_marker(marker)

trend = ChartTrendline()
trend.set_type("moving_average", 3)
trend.set_display_r_squared(True)
series.set_trendline(trend)

label = ChartDataLabel()
label.show_value()
label.set_position("above")
series.set_data_label(label)

chart = Chart("line")
chart.push_series(series)
```

Markers set on a series survive being pushed to a chart, including on
line, radar and scatter chart types where Excel defaults markers to off.
Those defaults still apply when no marker is set explicitly.

Marker types are `square`, `diamond`, `triangle`, `x`, `star`,
`short_dash`, `long_dash`, `circle`, `plus_sign`. Automatic and no-marker
are **not** types: use `set_automatic()` or `set_none()`.

Trendline types are `none`, `linear`, `exponential`, `logarithmic`
(not `logarithm`), `power`, `polynomial`, `moving_average`. The last two
take a period as the second argument to `set_type`, defaulting to 2.
`set_display_equation` and `set_display_r_squared` are named with a `set_`
prefix here for consistency, though upstream omits it.

Data labels support `show_value`, `show_category_name`,
`show_series_name`, `show_leader_lines`, `show_legend_key`,
`show_percentage`, `show_x_value`, `show_y_value`, `set_hidden`,
`set_position`, `set_num_format`, `set_separator`, `set_value`,
`set_font`, `set_format` and `set_custom`. Positions are `default`,
`center`, `right`, `left`, `above`, `below`, `inside_base`, `inside_end`,
`outside_end`, `best_fit`.

For per-point labels, build one `ChartDataLabel` per point, call
`set_custom()` on the ones that should differ, and pass the list to
`ChartSeries.set_custom_data_labels()`.

## Performance & Benchmarks

See [PERFORMANCE.md](PERFORMANCE.md) for full tables and methodology.

**100,000 rows × 8 columns (800,000 cells) — measured on Intel i5-6267U, Python 3.12, `rust_xlsxwriter` 0.96:**

| Method | Mean | vs xlsxwriter |
|---|---|---|
| **rvgs** `write_dataframe()` pandas | 1.43s | 🏆 **8.3× faster** |
| **rvgs** `write_dataframe()` polars | 1.53s | 7.8× faster |
| **rvgs** `write_dataframe()` pyarrow | 1.65s | 7.2× faster |
| **rvgs** `write_records()` (bulk dict) | 1.94s | 6.1× faster |
| **rvgs** `write_rows()` (list-of-lists) | 2.03s | 5.8× faster |
| xlsxwriter `write_row()` | 11.87s | 1× (baseline) |
| openpyxl write-only | 18.67s | 1.6× slower |
| pandas `to_excel()` openpyxl | 25.03s | 2.1× slower |

At 10,000 rows: **6.8× faster** than xlsxwriter. At 1,000 rows: **5.6× faster**.

The `write_dataframe()` path uses the Arrow C-stream zero-copy interface — pandas, polars, and pyarrow DataFrames all feed through this route with no Python-side data conversion.

**Run the benchmark on your machine:**

```bash
maturin develop --release
pip install xlsxwriter openpyxl pandas pyarrow polars
python benchmarks/run_benchmarks.py --runs 7 --warmup 2
```

### What the numbers depend on

Throughput is dominated by cell count and by how sparse the data is, not by
row count alone. The 100k×8 benchmark above and a wide-workload evaluation
(76,480 columns × 1,261 rows, 14 sheets) tell different stories:

| Comparison | Shape | Result |
|---|---|---|
| vs. pure-Python `xlsxwriter` | 100,000 rows × 8 cols | 6–8× faster |
| vs. another Rust-backed writer with zero-copy Arrow (`rustpy-xlsxwriter`) | wide export, same shape both sides | ~1.4× faster on total write time (latest measured run: 7.0s → 5.1s) |

If you're moving from pure-Python `xlsxwriter`, expect a large win. If you're
moving from another `rust_xlsxwriter` binding, expect a more modest one on
throughput — the bigger differences are GIL behaviour and memory, below.
Machine load moves absolute timings and even the ratio itself run to run —
an earlier pass on the same comparison measured ~2× rather than ~1.4× on
total write time for the same kind of workload. Treat the ratio as
directionally right (meaningfully faster, not dramatically faster) rather
than a number to hold this library to exactly; the output size and GIL
numbers below have been more stable across runs.

### `save()` releases the GIL

Long writes don't block other Python threads. Measured on the same
1,261 × 76,480 export across 14 sheets, a background thread polling at 5ms
saw a worst-case stall of **7–48ms** across runs, against **1,270–5,980ms**
for a comparable Rust-backed writer that holds the GIL for the duration.
The most recent measured run landed at 15ms vs 1,364ms — squarely inside
that range, not an outlier.

In practice this is the difference between a desktop GUI staying responsive
during a multi-second export and freezing solid for it. If you're writing
large workbooks from a Qt/Tk/wx application, or from a request handler that
also serves other work, this matters more than raw throughput.

### When `constant_memory` helps

`constant_memory=True` flushes rows to disk as they're written instead of
buffering the whole worksheet, so the saving scales with **row count**, not
column count. The table below is total process peak memory -- the writer
is one allocator among several in a real pipeline (dataframe, Arrow
buffers, the writer itself), so this understates what `constant_memory`
does to the writer's own footprint specifically:

| Rows | Peak, default | Peak, `constant_memory=True` |
|---|---|---|
| ~1,300 | 2,151 MB | 2,131 MB (no measurable gain) |
| ~25,220 | 7,613 MB | 6,837 MB (**10% lower**) |

Latest measured run: **10% off total process peak** (table above), but
**16% off the writer's own memory footprint** specifically once the
other allocators are factored out -- no absolute figure for that one,
just the percentage. An earlier, cruder pass had reported 26% without
that distinction, measuring something closer to the total-peak number
above but not quite the same benchmark; treat 10%/16% as the current,
more precise reading rather than reconciling it against the older 26%.

Rules of thumb:

- **Tall data (many rows): use it.** The row buffer is the dominant cost.
- **Wide but short data: it won't help much.** The per-row buffer was never
  the bottleneck.
- The constraint: rows must be written in non-decreasing order — see below.

The writer is often not the largest allocation in the pipeline, either: in
the ~25,000-row case above, the source dataframe itself accounts for a
substantial share of peak memory before the writer is even called --
part of why the *total* process peak drops by less (10%) than the
writer's own footprint does (16%).

### `constant_memory` fails loudly, not silently

`rust_xlsxwriter` itself does not error when rows are written out of order
in constant-memory mode — it produces a corrupt or incomplete `.xlsx`. This
binding adds an explicit check and raises `ValueError` naming the row
involved, and the workbook still closes cleanly afterward.

If you're producing files that go to clients or external systems, this is
the difference between an exception in your logs and a broken deliverable
nobody notices until the recipient opens it.

### A note on very wide workbooks

Survey and panel exports are routinely thousands of columns wide and
sparse — most variables are not asked of most respondents. Two
consequences:

- Excel's hard limit is 16,384 columns per worksheet, so anything wider
  must be split across sheets or files by the caller.
- Sparse data means most cells are genuinely absent from the XML. Any API
  that materializes a cell per column — to attach a format, for instance —
  multiplies both file size and write time by the sparsity ratio. `null`
  values are always skipped rather than written as formatted blank cells
  in `write_dataframe()`, specifically to avoid this (see
  [Applying a format across many columns](#applying-a-format-across-many-columns)).

If you're working at this shape, measure with your own data before
adopting an API — row-count-based benchmarks won't predict your results.

Output size is also consistently smaller than the comparison writer's on
identical data in that evaluation: 16.8 MB vs 18.3 MB, roughly 8%. Cause
not investigated on our end, so treat this as an observation rather than a
deliberate design win.

---

## Roadmap

| Version | Features |
|---------|----------|
| **v0.1** | ✅ Core writing, formatting, merging, formulas, dates, images, Polars/Pandas support |
| **v0.2** | ✅ Bulk `write_records()` / `write_rows()`; Arrow zero-copy `write_dataframe()`; `constant_memory` streaming mode; autofilter; defined names; worksheet tables; charts (phases 1-3); conditional formatting; sparklines; page setup; extended Arrow types |
| **v0.2.1** | ✅ Audit release: correctness fixes, GIL released during `save()`, allocation-free Arrow string path, streamed `write_dataframe()`, cross-platform CI |
| **v0.2.2** | ✅ Patch release: `Workbook.close()` / `with Workbook(path) as wb:` now work with no argument, using the constructor-provided path; version metadata alignment; `set_column_range_width()`; canonical `set_border_top/bottom/left/right()` names (old names kept as aliases); `.pyi` type stubs + `py.typed` marker; `Cargo.lock` committed; `write_dataframe(column_formats=...)` -- a true per-cell merge, so a border survives on a date column alongside its own number format, not the column-scoped workaround this shipped with first; `annotations` no longer leaks into the module namespace |
| **v0.2.3** | ✅ Patch release: `Workbook.save_to_buffer() -> bytes`; `Worksheet.set_header()`/`set_footer()` (text only, images still need an `Image` pyclass); `write_rows()` no longer double-materializes the dataset before writing |
| **v0.2.4** | ✅ Patch release: `write_dataframe()` no longer materializes a formatted blank cell for every null when `column_formats` is used (was up to 84x slower / 29x larger on sparse wide data); upgraded `rust_xlsxwriter` 0.96 -> 0.98.2 (MSRV 1.85 -> 1.88); `Format` at full parity except `set_font_scheme()` (`set_quote_prefix`, `set_hyperlink`, `set_checkbox`, `set_font_family/charset/script`, `set_reading_direction`, and matching `unset_*` inverses); `set_range_format_with_border()` and `clear_cell_format()` on `Worksheet` |
| **v0.2.5** | ✅ Patch release: `DataValidation` and `Worksheet.add_data_validation()` -- dropdown lists (from a string list or a cell range), whole-number/decimal-number/text-length range rules, custom formula rules, and every input/error-message setting. Date/time rules and cell-reference formula variants still open. |
| **v0.2.6** | ✅ Patch release: row/column outline grouping (`group_rows`/`group_columns`/`*_collapsed`/`group_symbols_above`/`group_symbols_to_left`). Known limitation: `group_rows()` doesn't apply per-row grouping when `constant_memory=True` -- see README's Known Limitations. |
| **v0.2.7** | ✅ Patch release: `ConditionalFormatIconSet` and `ConditionalFormatCustomIcon` -- all 20 icon set styles, `reverse_icons`, `show_icons_only`, per-icon threshold/type/direction overrides. |
| **v0.2.8** | ✅ Patch release: chart secondary axes -- `set_x2_axis_*`/`set_y2_axis_*`, mirroring the existing `x_axis`/`y_axis` setters. Takes effect once a series is routed to the secondary axis via the existing `ChartSeries.set_secondary_axis()`. |
| **v0.3** | 🚧 Date/time data validation rules; chart error bars |
| **v0.4** | 🚧 Chart error bars; cell notes and autofilter criteria; full xlsxwriter API compatibility layer |

**Charts** (`rust_xlsxwriter`'s largest subsystem -- 18k+ lines, 23 chart
types, ~214 public methods across ~39 types) were implemented in phases
so each could be verified rather than shipping a large surface untested.
All three phases have landed:
1. ✅ Core `Chart`/`ChartSeries` + common types (Bar/Column/Line/Pie/
   Scatter + stacked variants) + basic title/legend/axis
2. ✅ Formatting depth: `ChartFormat`/`ChartFont`, line/font/fill styling
3. ✅ Decorations: `ChartMarker`, `ChartTrendline`, `ChartDataLabel`
4. ✅ Secondary axes: `set_x2_axis_*`/`set_y2_axis_*`

Remaining chart work, tracked in [MISSING.md](MISSING.md): error bars,
data tables, manual layout, and the less common chart types (Radar,
Stock, Doughnut, Surface).

### Secondary Axes

Route a series to the secondary axis with the existing
`ChartSeries.set_secondary_axis()`, then style that axis with
`set_x2_axis_*`/`set_y2_axis_*` (same method set as the primary
`set_x_axis_*`/`set_y_axis_*`, minus `date_axis`/`text_axis` on the y
side). The secondary-axis XML is only written once a series actually
uses it — calling the setters alone has no effect.

```python
from rvgsrust_xlsxwriter import Workbook, Chart, ChartSeries

wb = Workbook("secondary_axis.xlsx")
ws = wb.add_worksheet()
ws.write_column(0, 0, [10, 40, 50, 20, 10, 50])
ws.write_column(0, 1, [1, 4, 5, 2, 1, 5])

units = ChartSeries()
units.set_values("Sheet1!$A$1:$A$6")

revenue = ChartSeries()
revenue.set_values("Sheet1!$B$1:$B$6")
revenue.set_secondary_axis(True)

chart = Chart("column")
chart.push_series(units)
chart.push_series(revenue)
chart.set_y_axis_name("Units")
chart.set_y2_axis_name("Revenue ($M)")
ws.insert_chart(0, 3, chart)

wb.close()
```

### Known limitations

These are current, deliberate gaps rather than oversights:

- **`write_dataframe()` column types.** Decimal, list, struct and
  dictionary-encoded Arrow columns are not supported; such a column
  raises `TypeError` and `dataframe.py` falls back to the per-cell path.
- **Timezones.** Excel has no timezone concept. Timezone-aware Arrow
  timestamps are written as UTC wall-clock time and warn once per column.
- **Precision.** Excel cells hold an f64, so integers above 2^53 lose
  precision. This is a format limitation, not an implementation one.
- **`constant_memory=True`** requires rows to be written in
  non-decreasing order, and raises `ValueError` naming the offending row
  on violation rather than silently emitting a corrupt file (which is
  what `rust_xlsxwriter` itself would do) — see
  [`constant_memory` fails loudly, not silently](#constant_memory-fails-loudly-not-silently)
  for why this is deliberate rather than just defensive. Separately,
  `group_rows()`/`group_rows_collapsed()` don't apply per-row grouping
  at all under `constant_memory=True` — see
  [Row and Column Grouping](#row-and-column-grouping).
- **Python 3.8.** `requires-python` still declares `>=3.8`, but CI tests
  3.9 upward; the `pandas>=2.0` extra already requires 3.9+.

---

## API Parity

[MISSING.md](MISSING.md) audits the exposed Python API against
rust_xlsxwriter 0.98.2 (re-checked line-by-line against that source
after this project's 0.96->0.98.2 upgrade -- see MISSING.md's own note)
and lists what is not yet wrapped, with upstream `file:line` references,
a suggested Python API shape, and a priority for each. Sparklines,
cell/row/column/range formats, page setup and print settings, and
per-side border naming are all at full parity now -- an
earlier audit pass had flagged per-side borders as missing entirely,
which was a false positive from a naming mismatch, and separately the
naming itself (`set_top_border` vs upstream's `set_border_top`) has
since been reconciled: both spellings work, the reversed ones kept for
compatibility.

The largest remaining gaps are header/footer text images on
`Worksheet` (`set_header_image`/`set_footer_image` -- still need an
`Image` pyclass first); date/time data validation rules and the
cell-reference formula variants of the numeric/text-length/date/time
rules; and error bars and a handful of formatting options on `Chart`
(secondary axes closed in v0.2.8). `Conditional formats` and `Format`
are both at full parity now (`Format` except `set_font_scheme()`, deliberately
not exposed) -- see MISSING.md's "Suggested order" for the full
ranked list.

## Performance TODO

Measured bottlenecks and untried optimisations are logged in
[PERFORMANCE_TODO.md](PERFORMANCE_TODO.md), separately from the benchmark
results in [PERFORMANCE.md](PERFORMANCE.md). Nothing there is implemented
yet, by design: they change hot paths and each needs its own before/after
benchmark run.

## Contributing

We welcome contributions! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/amazing-feature`
3. Commit your changes: `git commit -m 'Add amazing feature'`
4. Push to the branch: `git push origin feature/amazing-feature`
5. Open a Pull Request

---

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

---

## Acknowledgements

- [John McNamara](https://github.com/jmcnamara) for `rust_xlsxwriter` and `XlsxWriter`
- [PyO3](https://github.com/PyO3/pyo3) for Rust-Python bindings
- [maturin](https://github.com/PyO3/maturin) for building and publishing

---

<p align="center">
  <b>RVGSRust — Excel at Rust speed, with Python ease.</b>
</p>
