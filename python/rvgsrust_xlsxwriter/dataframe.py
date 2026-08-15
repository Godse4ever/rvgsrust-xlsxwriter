"""
DataFrame support for Polars and Pandas.

Uses Worksheet.write_dataframe() -- a native Rust path that reads
directly from the DataFrame's underlying Arrow buffers via the Arrow
PyCapsule Interface, without extracting individual Python objects per
cell -- whenever possible. That path supports the signed and unsigned
integer widths, float32/float64, string (utf8/large_utf8/utf8view), bool,
date32/date64 and timestamp columns, and a single format applied to the
whole header row.

`column_formats` is passed straight through to write_dataframe(), which
merges each column's format into every cell of that column as it's
written -- a border on a date column survives alongside the column's
own date number format, because both are applied to the same cellXf
rather than the column format losing to the cell's own format under
OOXML precedence (which is what happens with set_column_format() /
set_column_range_format() applied after the fact). The merge happens
once per column before Rust starts writing rows, not per cell, so
passing column_formats does not touch the fast path's per-cell cost.

When a column has a type write_dataframe() doesn't support yet, this
falls back to the original per-cell write() loop instead, so every
DataFrame this module has ever accepted still works -- just without the
speed benefit in that case. column_formats is applied in the fallback
too, though there it's whatever format was given, with no dtype
merging -- moot in practice, since date/timestamp columns are always
supported by the fast path and never reach this fallback.

The fast path streams Arrow record batches rather than materialising them,
so it also accepts any streaming producer exposing __arrow_c_stream__ (for
example a pyarrow.RecordBatchReader over a Parquet file), and its peak
memory is bounded by one batch rather than by the whole dataset.

Fallback contract: only TypeError triggers the per-cell retry, because
write_dataframe() validates the entire schema before writing anything, so a
TypeError is guaranteed to mean nothing was written. Failures after writing
has started surface as RuntimeError and are not retried. An unknown
column_formats key raises ValueError from write_dataframe() itself before
any row is written, so that isn't caught here either -- it should propagate
straight to the caller, not trigger a fallback retry.
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
    """Slow path: one write() call per cell. Used only when
    write_dataframe() rejects a column type it doesn't support yet.
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
            Merged into each column's cells on the fast Arrow path --
            see module docstring for what "merged" means for
            date/datetime columns.
    """
    if not HAS_POLARS:
        raise ImportError("Polars is not installed. Install with: pip install polars")

    if not isinstance(df, pl.DataFrame):
        raise TypeError(f"Expected Polars DataFrame, got {type(df)}")

    try:
        worksheet.write_dataframe(
            row, col, df, header_format=header_format, column_formats=column_formats
        )
        return
    except TypeError:
        # Only TypeError is caught, and that is deliberate rather than
        # incidental. write_dataframe() validates the whole schema before
        # writing any cell, so a TypeError means "this column type is not
        # supported yet" and *nothing has been written* -- making the
        # per-cell retry below safe. Any failure that occurs once writing
        # has begun is raised as RuntimeError instead, precisely so it
        # propagates here rather than triggering a retry that would
        # duplicate the rows already on the sheet. An unknown
        # column_formats key raises ValueError, also not caught here,
        # for the same reason: nothing was written, but retrying
        # wouldn't fix an unknown column name either.
        pass

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
            Merged into each column's cells on the fast Arrow path --
            see module docstring for what "merged" means for
            date/datetime columns.
    """
    if not HAS_PANDAS:
        raise ImportError("Pandas is not installed. Install with: pip install pandas")

    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"Expected Pandas DataFrame, got {type(df)}")

    columns = [str(c) for c in df.columns]

    try:
        worksheet.write_dataframe(
            row, col, df, header_format=header_format, column_formats=column_formats
        )
        return
    except TypeError:
        # Only TypeError is caught, and that is deliberate rather than
        # incidental. write_dataframe() validates the whole schema before
        # writing any cell, so a TypeError means "this column type is not
        # supported yet" and *nothing has been written* -- making the
        # per-cell retry below safe. Any failure that occurs once writing
        # has begun is raised as RuntimeError instead, precisely so it
        # propagates here rather than triggering a retry that would
        # duplicate the rows already on the sheet. An unknown
        # column_formats key raises ValueError, also not caught here,
        # for the same reason: nothing was written, but retrying
        # wouldn't fix an unknown column name either.
        pass

    rows = (row_data for _, row_data in df.iterrows())
    _write_dataframe_per_cell(worksheet, columns, rows, row, col, header_format, column_formats)
