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
- Date/time support (write_datetime / write_date)
- Image insertion
- Worksheet tables (Table / TableColumn)
- Autofilter
- Sheet protection and freeze panes
- Rich string support (write_rich_string)
- Polars, Pandas, and PyArrow DataFrame support (zero-copy Arrow path)

Not a drop-in replacement for Python xlsxwriter: the API is inspired by
it but differs in real ways (e.g. Workbook() takes no path argument;
call close(path) instead; Format objects are built via chained setter
methods rather than a single dict). Charts, conditional formatting, and
data validation are not yet implemented -- see the Roadmap in README.md.

Note on memory management: Worksheet objects hold a reference back to
their parent Workbook (needed to access the underlying Rust worksheet).
This creates a reference cycle that Python's cyclic garbage collector
handles correctly, but may delay collection in long-running processes
if many Workbook objects are created without explicitly dropping all
Worksheet references. For best practice, prefer the context manager
form (``with Workbook() as wb``) or ensure Worksheet handles go out of
scope before the Workbook is closed.

Installation
------------
    pip install rvgsrust-xlsxwriter

Quick Start
-----------
    from rvgsrust_xlsxwriter import Workbook

    # Context manager form (recommended):
    with Workbook() as wb:
        ws = wb.add_worksheet("Sheet1")

        fmt = wb.add_format()
        fmt.set_bold()
        fmt.set_background_color("#4472C4")
        fmt.set_font_color("white")
        fmt.set_border("thin")

        ws.write(0, 0, "Hello", fmt)
        ws.merge_range(1, 0, 1, 2, "Merged Cell", fmt)
        ws.autofit()
    # wb.close("report.xlsx") is called automatically on __exit__

    # Or explicit close:
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write(0, 0, "Hello")
    wb.close("report.xlsx")

"""

__version__ = "0.2.0.dev0"
__author__ = "RVGS Team"
__license__ = "MIT"

from rvgsrust_xlsxwriter._core import (
    Workbook as _CoreWorkbook,
    Worksheet,
    Format,
    Table,
    TableColumn,
)


class Workbook(_CoreWorkbook):
    """Workbook with context manager support.

    Wraps the core Rust-backed Workbook to add ``__enter__`` /
    ``__exit__``, so the workbook can be used as::

        with Workbook() as wb:
            ws = wb.add_worksheet()
            ws.write(0, 0, "Hello")
        # wb.close(path) called automatically

    ``__exit__`` requires a ``path`` to save to; set ``wb.path``
    before the ``with`` block, or use the explicit ``close(path)``
    form instead.
    """

    def __init__(self, path: str | None = None) -> None:
        super().__init__()
        # Optional: pre-set the output path so __exit__ knows where to save.
        self.path = path

    def __enter__(self) -> "Workbook":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if exc_type is None:
            if self.path is None:
                raise ValueError(
                    "Workbook.__exit__: no path set. "
                    "Either pass path= to Workbook() or call close(path) manually."
                )
            self.close(self.path)
        # If an exception occurred, don't try to save -- let it propagate.
        return False


__all__ = ["Workbook", "Worksheet", "Format", "Table", "TableColumn"]
