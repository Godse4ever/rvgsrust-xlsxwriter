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
| Speed | Baseline (1x) | ~7-10x faster* | ~5-6x faster* ⚡ |
| **Cell Merging** | ✅ Yes | ❌ No | ✅ **Yes** |
| **Full Format API** | ✅ Yes | ⚠️ Limited | ✅ **Complete** |
| **Borders (all sides)** | ✅ Yes | ⚠️ Basic | ✅ **All sides + colors** |
| **Cell/Font Colors** | ✅ Yes | ✅ Yes | ✅ **Yes** |
| **Formulas** | ✅ Yes | ❌ No | ✅ **Yes** |
| **Hyperlinks** | ✅ Yes | ❌ No | ✅ **Yes** |
| **Date/Time** | ✅ Yes | ✅ Yes | ✅ **Yes** |
| **Images** | ✅ Yes | ❌ No | ✅ **Yes** |
| **Freeze Panes** | ✅ Yes | ✅ Yes | ✅ **Yes** |
| **Sheet Protection** | ✅ Yes | ✅ Yes | ✅ **Yes** |
| **Polars/Pandas** | Manual | ✅ Zero-copy | ✅ **Zero-copy** |
| **Charts** | ✅ Yes | ❌ No | 🚧 *Coming v0.2* |
| **Conditional Format** | ✅ Yes | ❌ No | 🚧 *Coming v0.2* |

**We win on completeness.** Both `rustpy-xlsxwriter` and `rvgsrust-xlsxwriter` use the same Rust core, but we expose *every* feature so you never have to fall back to Python.

*Speed figures above are from measurements against `rust_xlsxwriter 0.75`, before this project's 0.96 upgrade -- see [Performance & Benchmarks](#performance--benchmarks) below for the honest, current status and how to reproduce them on your own machine.

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

**Current type support:** `int64`, `float64`, `string`/`utf8`, `large_utf8`, `bool`.

---

## Performance & Benchmarks

See [PERFORMANCE.md](PERFORMANCE.md) for optimization strategies and details.

**Quick comparison (100,000 rows × 5 columns) -- measured against
`rust_xlsxwriter 0.75`, before this project's upgrade to 0.96
(`zmij` + `constant_memory` feature flags). Not yet re-measured against
0.96 -- run the benchmark yourself (below) for current numbers on your
machine and build.**

| Strategy | Time (0.75-era) | Speedup vs. xlsxwriter |
|----------|------|---------|
| `write_dataframe()` (Polars, zero-copy) | 0.496s | 5.7x |
| `write_records()` (bulk dict) | 0.550s | 5.1x |
| `write()` per-cell | 0.562s | 5.0x |
| `rustpy-xlsxwriter` (native API) | ~0.27-0.38s | ~7-10x |
| Pure Python xlsxwriter | 2.821s | 1x (baseline) |

Honest take at the time of this measurement: `rustpy-xlsxwriter` was
faster than `rvgsrust-xlsxwriter`, tracing to it building against
`rust_xlsxwriter`'s `constant_memory` feature. That feature is now
available here too (`Workbook.add_worksheet(constant_memory=True)`),
but isn't wired into any default path, and its actual effect on this
gap hasn't been measured yet.

**Run the benchmark on your machine:**

```bash
maturin develop --release
python examples/benchmark.py
```

This tests multiple strategies (write_records, write_dataframe with Polars/Pandas, per-cell writes) and compares against rustpy-xlsxwriter and xlsxwriter.

---

## Roadmap

| Version | Features |
|---------|----------|
| **v0.1** | ✅ Core writing, formatting, merging, formulas, dates, images, Polars/Pandas support |
| **v0.2** | ✅ Bulk `write_records()`; Arrow zero-copy `write_dataframe()`; `constant_memory` streaming mode; autofilter; defined names. 🚧 Charts, conditional formatting, extended Arrow types |
| **v0.3** | 🚧 Data validation, tables, sparklines |
| **v0.4** | 🚧 Full xlsxwriter API compatibility layer |

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
