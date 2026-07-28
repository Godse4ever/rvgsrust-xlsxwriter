"""
DataFrame support for Polars and Pandas.

Uses Worksheet.write_dataframe() -- a native Rust path that reads
directly from the DataFrame's underlying Arrow buffers via the Arrow
PyCapsule Interface, without extracting individual Python objects per
cell -- whenever possible. That path supports the signed and unsigned
integer widths, float32/float64, string (utf8/large_utf8/utf8view), bool,
date32/date64 and timestamp columns, and a single format applied to the
whole header row. When `column_formats` is given (per-column
formatting isn't supported by write_dataframe() yet) or a column has an
unsupported type, this falls back to the original per-cell write() loop
so every DataFrame this module has ever accepted still works -- just
without the speed benefit in those cases.
"""

try:
    import polars as pl
    HAS_POLARS = True
except ImportError:
    HAS_POLARS = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


def _write_dataframe_per_cell(worksheet, columns, rows, row, col, header_format, column_formats):
    """Slow path: one write() call per cell. Used when column_formats is
    given, or when write_dataframe() rejects a column type it doesn't
    support yet.
    """
    for i, col_name in enumerate(columns):
        worksheet.write(row, col + i, col_name, header_format)
    for r_idx, row_data in enumerate(rows):
        for c_idx, value in enumerate(row_data):
            fmt = None
            if column_formats and columns[c_idx] in column_formats:
                fmt = column_formats[columns[c_idx]]
            worksheet.write(row + 1 + r_idx, col + c_idx, value, fmt)


def write_polars_dataframe(worksheet, df, row=0, col=0, header_format=None, column_formats=None):
    """Write a Polars DataFrame to a worksheet.

    Args:
        worksheet: rvgsrust_xlsxwriter Worksheet object
        df: Polars DataFrame
        row: Starting row (0-indexed)
        col: Starting column (0-indexed)
        header_format: Format object for header row
        column_formats: Dict mapping column names to Format objects.
            When given, uses the slower per-cell write path (see module
            docstring) since the fast Arrow path doesn't support
            per-column formatting yet.
    """
    if not HAS_POLARS:
        raise ImportError("Polars is not installed. Install with: pip install polars")

    if not isinstance(df, pl.DataFrame):
        raise TypeError(f"Expected Polars DataFrame, got {type(df)}")

    if column_formats is None:
        try:
            worksheet.write_dataframe(row, col, df, header_format=header_format)
            return
        except TypeError:
            pass  # unsupported column type in the fast path -- fall back below

    _write_dataframe_per_cell(
        worksheet, df.columns, df.iter_rows(), row, col, header_format, column_formats
    )


def write_pandas_dataframe(worksheet, df, row=0, col=0, header_format=None, column_formats=None):
    """Write a Pandas DataFrame to a worksheet.

    Args:
        worksheet: rvgsrust_xlsxwriter Worksheet object
        df: Pandas DataFrame
        row: Starting row (0-indexed)
        col: Starting column (0-indexed)
        header_format: Format object for header row
        column_formats: Dict mapping column names to Format objects.
            When given, uses the slower per-cell write path (see module
            docstring) since the fast Arrow path doesn't support
            per-column formatting yet.
    """
    if not HAS_PANDAS:
        raise ImportError("Pandas is not installed. Install with: pip install pandas")

    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"Expected Pandas DataFrame, got {type(df)}")

    if column_formats is None:
        try:
            worksheet.write_dataframe(row, col, df, header_format=header_format)
            return
        except TypeError:
            pass  # unsupported column type in the fast path -- fall back below

    columns = [str(c) for c in df.columns]
    rows = (row_data for _, row_data in df.iterrows())
    _write_dataframe_per_cell(worksheet, columns, rows, row, col, header_format, column_formats)
