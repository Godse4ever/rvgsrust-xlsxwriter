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
- Charts, conditional formatting and sparklines
- Polars, Pandas, and PyArrow DataFrame support (zero-copy Arrow path)

Not a drop-in replacement for Python xlsxwriter: the API is inspired by
it but differs in real ways (e.g. Workbook() takes no path argument;
call close(path) instead; Format objects are built via chained setter
methods rather than a single dict). Charts, conditional formatting and
sparklines ARE implemented (see the exports below); data validation is
not yet -- see the Roadmap in README.md.

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
    with Workbook("report.xlsx") as wb:
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

# from __future__ import annotations makes all annotations lazily
# evaluated strings, letting us use PEP 604 unions (str | None) even
# under Python 3.8/3.9 which don't support the runtime | operator on
# types. Without this, importing this module on 3.8/3.9 fails with
# TypeError at the annotated line -> pytest collection exits with
# code 2, taking the whole test job with it.
from __future__ import annotations

__version__ = "0.2.2"
__author__ = "RVGS Team"
__license__ = "MIT"

from rvgsrust_xlsxwriter._core import (
    Workbook as _CoreWorkbook,
    Worksheet,
    Format,
    Table,
    TableColumn,
    ConditionalFormatCell,
    ConditionalFormatBlank,
    ConditionalFormatDuplicate,
    ConditionalFormatError,
    ConditionalFormatFormula,
    ConditionalFormatAverage,
    ConditionalFormatTop,
    ConditionalFormatText,
    ConditionalFormatDate,
    ConditionalFormat2ColorScale,
    ConditionalFormat3ColorScale,
    ConditionalFormatDataBar,
    Sparkline,
    Chart,
    ChartSeries,
    ChartFont,
    ChartFormat,
    ChartMarker,
    ChartTrendline,
    ChartDataLabel,
)


class Workbook(_CoreWorkbook):
    """Workbook with context manager support.

    Wraps the core Rust-backed Workbook to add ``__enter__`` /
    ``__exit__``, so the workbook can be used as::

        with Workbook("out.xlsx") as wb:
            ws = wb.add_worksheet()
            ws.write(0, 0, "Hello")
        # wb.close("out.xlsx") called automatically

    ``__exit__`` requires a ``path`` to save to; pass it to the
    constructor or set ``wb.path`` before the ``with`` block, or
    use the explicit ``close(path)`` form instead.
    """

    def __new__(cls, path=None):
        # pyo3's #[new] fn new() takes zero args, so the generated __new__
        # rejects any kwarg. Python calls __new__ BEFORE __init__, so if
        # we let 'path' fall through to super().__new__ we get:
        #   TypeError: Workbook.__new__() got an unexpected keyword argument 'path'
        # and __init__ is never reached. Absorb the kwarg here instead and
        # call the parent __new__ with no args; __init__ below then stores
        # path on the instance for __exit__ to find.
        return super().__new__(cls)

    def __init__(self, path: str | None = None) -> None:
        # _CoreWorkbook is a #[pyclass(subclass)] with #[new] fn new().
        # Python invokes the pyo3-generated __new__ (via our __new__ above)
        # before calling this __init__; the Rust side is already fully
        # constructed. We just tack on the Python-side attribute for __exit__.
        self.path = path

    def __enter__(self) -> "Workbook":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if exc_type is None:
            # Try to close using the stored path if the user omitted it.
            self.close()
        # If an exception occurred, don't try to save -- let it propagate.
        return False

    def close(self, path: str | None = None) -> None:
        """Close and save the workbook.

        If `path` is omitted, the constructor-provided `path` stored on
        the instance is used. Passing an explicit `path` continues to work
        as an override (backwards compatible).
        """
        if path is None:
            path = getattr(self, "path", None)
        if path is None:
            raise ValueError(
                "Workbook.close(): no path specified. Either pass a path to close() or construct the Workbook with a path."
            )
        # Delegate to the underlying Rust-backed close(path) method.
        return super().close(path)


__all__ = [
    "Workbook",
    "Worksheet",
    "Format",
    "Table",
    "TableColumn",
    "ConditionalFormatCell",
    "ConditionalFormatBlank",
    "ConditionalFormatDuplicate",
    "ConditionalFormatError",
    "ConditionalFormatFormula",
    "ConditionalFormatAverage",
    "ConditionalFormatTop",
    "ConditionalFormatText",
    "ConditionalFormatDate",
    "ConditionalFormat2ColorScale",
    "ConditionalFormat3ColorScale",
    "ConditionalFormatDataBar",
    "Sparkline",
    "Chart",
    "ChartSeries",
    "ChartFont",
    "ChartFormat",
    "ChartMarker",
    "ChartTrendline",
    "ChartDataLabel",
]
