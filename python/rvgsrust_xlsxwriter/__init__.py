"""
RVGSRust-XLSXWriter
===================

A Rust-powered XLSX library for Python, built on the official
rust_xlsxwriter crate, with a Pythonic API and Polars/Pandas/PyArrow
support.

Features
--------
- Full formatting: borders, colors, fonts, alignment, patterns
- Cell merging
- Formulas and hyperlinks
- Date/time support
- Image insertion
- Sheet protection and freeze panes
- Polars and Pandas DataFrame support (zero-copy where possible)

Not a drop-in replacement for Python xlsxwriter: the API is inspired by
it but differs in real ways (e.g. Workbook() takes no path argument;
call close(path) instead; Format objects are built via chained setter
methods rather than a single dict). Charts, conditional formatting,
data validation, and tables are not yet implemented -- see the
Roadmap in README.md for what's planned and what's already there.

Installation
------------
    pip install rvgsrust-xlsxwriter

Quick Start
-----------
    from rvgsrust_xlsxwriter import Workbook

    wb = Workbook()
    ws = wb.add_worksheet("Sheet1")

    fmt = wb.add_format()
    fmt.set_bold()
    fmt.set_background_color("#4472C4")
    fmt.set_font_color("white")
    fmt.set_border("thin")

    ws.write(0, 0, "Hello", fmt)
    ws.merge_range(1, 0, 1, 2, "Merged Cell", fmt)
    ws.autofit()

    wb.close("report.xlsx")

"""

__version__ = "0.1.0"
__author__ = "RVGS Team"
__license__ = "MIT"

from rvgsrust_xlsxwriter._core import Workbook, Worksheet, Format

__all__ = ["Workbook", "Worksheet", "Format"]
