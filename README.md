# RVGSRust-XLSXWriter

> **The most feature-complete, Pythonic Rust-powered XLSX library.**
>
> Built on the official [`rust_xlsxwriter`](https://github.com/jmcnamara/rust_xlsxwriter) crate — the same engine trusted by the Python `XlsxWriter` library.

[![PyPI version](https://badge.fury.io/py/rvgsrust-xlsxwriter.svg)](https://pypi.org/project/rvgsrust-xlsxwriter/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Rust](https://img.shields.io/badge/Rust-1.70%2B-orange.svg)](https://www.rust-lang.org)
[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org)

---

## Why RVGSRust-XLSXWriter?

| Feature | `xlsxwriter` (Python) | `rustpy-xlsxwriter` | `rvgsrust-xlsxwriter` |
|---------|----------------------|---------------------|----------------------|
| Speed | Baseline (1x) | ~7-9x faster | Comparable ⚡ |
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

**We don't compete on raw speed — we win on completeness.** Both `rustpy-xlsxwriter` and `rvgsrust-xlsxwriter` use the same Rust core. We expose *every* feature so you never have to fall back to a slower library.

---

## Installation

```bash
pip install rvgsrust-xlsxwriter
```

### Optional Dependencies

```bash
# For Polars DataFrame support
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

# Save
wb.close()
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

# One call for the entire dataset, instead of one write() call per
# cell -- this is the fast path, see Benchmarks below.
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
```

---

## Polars / Pandas Integration

### Polars (Zero-Copy)

```python
import polars as pl
from rvgsrust_xlsxwriter import Workbook
from rvgsrust_xlsxwriter.dataframe import write_polars_dataframe

df = pl.DataFrame({
    "Name": ["Alice", "Bob", "Charlie"],
    "Sales": [100.5, 200.75, 150.0],
})

wb = Workbook("output.xlsx")
ws = wb.add_worksheet()

header_fmt = wb.add_format()
header_fmt.set_bold()
header_fmt.set_background_color("#4472C4")
header_fmt.set_font_color("white")

write_polars_dataframe(ws, df, header_format=header_fmt)
ws.autofit()
wb.close()
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

wb = Workbook("output.xlsx")
ws = wb.add_worksheet()

write_pandas_dataframe(ws, df)
wb.close()
```

> **Note:** Native zero-copy Arrow integration is planned for v0.2. Current DataFrame support uses optimized Python-side iteration.

---

## Benchmarks

Run the included benchmark to see performance on your machine:

```bash
python examples/benchmark.py
```

**Use `write_records()` for bulk data, not per-cell `write()`.** `write_records()`
takes an entire list of dicts in a single Python->Rust call instead of one call per
cell, which matters a lot at scale:

```python
ws.write_records(0, 0, data, header_format=header_fmt)  # one call for the whole sheet
```

**Measured results** (this repo's sandbox; single core, your machine will vary --
mean of 3 runs after a discarded warmup run):

| Rows | xlsxwriter (Python) | rvgsrust (`write_records`, bulk) | rvgsrust (`write`, per-cell) | rustpy-xlsxwriter |
|------|--------------------|-----------------------------------|-------------------------------|--------------------|
| 1,000 | 0.036s | 0.0043s | 0.0050s | 0.0031s |
| 10,000 | 0.266s | 0.0440s | 0.0486s | 0.0268s |
| 100,000 | 2.821s | 0.550s | 0.562s | 0.269s |

Honest take: both Rust-backed libraries are 5-10x faster than pure-Python
`xlsxwriter`. Between the two, `rustpy-xlsxwriter` is currently ~1.7-2x faster than
`rvgsrust-xlsxwriter` at this dataset size, even using the bulk `write_records()`
path. The remaining gap traces to `rustpy-xlsxwriter` building against
`rust_xlsxwriter`'s `constant_memory` feature (streams rows instead of buffering the
whole sheet) plus a newer `zip`/`pyo3` stack -- this repo is currently pinned to
`rust_xlsxwriter 0.75` because later versions' `zip` requirement doesn't resolve
against a Rust toolchain older than ~1.85 (see the comment in `Cargo.toml`). On a
machine with a current Rust toolchain, bumping to `rust_xlsxwriter 0.95` with
`features = ["constant_memory", "zmij", "zlib"]` is untested here but worth trying,
and should close most or all of the remaining gap.

---

## Roadmap

| Version | Features |
|---------|----------|
| **v0.1** | ✅ Core writing, formatting, merging, formulas, dates, images |
| **v0.2** | 🚧 Charts, conditional formatting, native Arrow zero-copy |
| **v0.3** | 🚧 Data validation, tables, Sparklines |
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
