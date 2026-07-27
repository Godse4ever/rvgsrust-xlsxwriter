# RVGSRust-XLSXWriter

> **A Rust-powered XLSX library for Python, with a Pythonic API and Polars/Pandas/PyArrow support.**
>
> Built on the official [`rust_xlsxwriter`](https://github.com/jmcnamara/rust_xlsxwriter) crate. Not affiliated with the Python [`XlsxWriter`](https://xlsxwriter.readthedocs.io/) package (a separate, pure-Python project it shares a similar name and purpose with) -- this library's Python-facing API is inspired by it but is not a drop-in replacement; see [Quick Start](#quick-start) below for the real differences.

[![PyPI version](https://badge.fury.io/py/rvgsrust-xlsxwriter.svg)](https://pypi.org/project/rvgsrust-xlsxwriter/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Rust](https://img.shields.io/badge/Rust-1.83%2B-orange.svg)](https://www.rust-lang.org)
[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org)

---

## Why RVGSRust-XLSXWriter?

| Feature | `xlsxwriter` (Python) | `rustpy-xlsxwriter` | `rvgsrust-xlsxwriter` |
|---------|----------------------|---------------------|----------------------|
| Speed (100k rows) | 1× (baseline) | — | **6–8× faster** ⚡ |
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
| **Charts** | ✅ Yes | ❌ No | 🚧 *Coming v0.3* |
| **Conditional Format** | ✅ Yes | ❌ No | 🚧 *Coming v0.3* |

**We win on completeness.** Both `rustpy-xlsxwriter` and `rvgsrust-xlsxwriter` use the same Rust core, but we expose *every* feature so you never have to fall back to Python.

Speed figures measured against `rust_xlsxwriter` 0.96 on Intel i5-6267U, 100k rows × 8 cols — see [Performance & Benchmarks](#performance--benchmarks) for full numbers.

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

# Create workbook
wb = Workbook()
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

# Save
wb.close("report.xlsx")
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

**Current type support:** `int64`, `float64`, `string`/`utf8`, `large_utf8`, `utf8view` (Polars default), `bool`.

---

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

---

## Roadmap

| Version | Features |
|---------|----------|
| **v0.1** | ✅ Core writing, formatting, merging, formulas, dates, images, Polars/Pandas support |
| **v0.2** | ✅ Bulk `write_records()`; Arrow zero-copy `write_dataframe()`; `constant_memory` streaming mode; autofilter; defined names; worksheet tables. 🚧 Charts (in progress, phased -- see below), conditional formatting, extended Arrow types |
| **v0.3** | 🚧 Data validation, sparklines |
| **v0.4** | 🚧 Full xlsxwriter API compatibility layer |

**Charts** (`rust_xlsxwriter`'s largest subsystem -- 18k+ lines, 23 chart
types, ~214 public methods across ~39 types) is being implemented in
phases rather than all at once, so each phase can be properly verified
rather than shipping a large surface untested:
1. Core `Chart`/`ChartSeries` + common types (Bar/Column/Line/Pie/
   Scatter + stacked variants) + basic title/legend/axis
2. Formatting depth: line/font/fill styling
3. Advanced: trendlines, error bars, data tables, layout, remaining
   chart types (Radar, Stock, Doughnut, Surface, etc.)

---

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
