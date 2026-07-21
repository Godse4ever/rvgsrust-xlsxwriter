"""
RVGSRust-XLSXWriter
===================

The most feature-complete, Pythonic Rust-powered XLSX library.
Built on the official rust_xlsxwriter crate.

Features
--------
- Full formatting: borders, colors, fonts, alignment, patterns
- Cell merging
- Formulas and hyperlinks
- Date/time support
- Image insertion
- Sheet protection and freeze panes
- Polars and Pandas DataFrame support (zero-copy where possible)
- Drop-in replacement for Python xlsxwriter

Installation
------------
    pip install rvgsrust-xlsxwriter

Quick Start
-----------
    from rvgsrust_xlsxwriter import Workbook

    wb = Workbook("report.xlsx")
    ws = wb.add_worksheet("Sheet1")

    fmt = wb.add_format()
    fmt.set_bold()
    fmt.set_background_color("#4472C4")
    fmt.set_font_color("white")
    fmt.set_border("thin")

    ws.write(0, 0, "Hello", fmt)
    ws.merge_range(1, 0, 1, 2, "Merged Cell", fmt)
    ws.autofit()

    wb.close()

"""

__version__ = "0.1.0"
__author__ = "RVGS Team"
__license__ = "MIT"

from rvgsrust_xlsxwriter._core import Workbook, Worksheet, Format

__all__ = ["Workbook", "Worksheet", "Format"]
