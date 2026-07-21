"""
DataFrame support for Polars and Pandas.
This is a Python-side helper that will be enhanced with native Rust Arrow
integration in future versions.
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


def write_polars_dataframe(worksheet, df, row=0, col=0, header_format=None, column_formats=None):
    """Write a Polars DataFrame to a worksheet.

    Args:
        worksheet: rvgsrust_xlsxwriter Worksheet object
        df: Polars DataFrame
        row: Starting row (0-indexed)
        col: Starting column (0-indexed)
        header_format: Format object for header row
        column_formats: Dict mapping column names to Format objects
    """
    if not HAS_POLARS:
        raise ImportError("Polars is not installed. Install with: pip install polars")

    if not isinstance(df, pl.DataFrame):
        raise TypeError(f"Expected Polars DataFrame, got {type(df)}")

    # Write headers
    for i, col_name in enumerate(df.columns):
        worksheet.write(row, col + i, col_name, header_format)

    # Write data row by row (will be optimized to native Rust in v0.2)
    for r_idx, row_data in enumerate(df.iter_rows()):
        for c_idx, value in enumerate(row_data):
            fmt = None
            if column_formats and df.columns[c_idx] in column_formats:
                fmt = column_formats[df.columns[c_idx]]
            worksheet.write(row + 1 + r_idx, col + c_idx, value, fmt)


def write_pandas_dataframe(worksheet, df, row=0, col=0, header_format=None, column_formats=None):
    """Write a Pandas DataFrame to a worksheet.

    Args:
        worksheet: rvgsrust_xlsxwriter Worksheet object
        df: Pandas DataFrame
        row: Starting row (0-indexed)
        col: Starting column (0-indexed)
        header_format: Format object for header row
        column_formats: Dict mapping column names to Format objects
    """
    if not HAS_PANDAS:
        raise ImportError("Pandas is not installed. Install with: pip install pandas")

    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"Expected Pandas DataFrame, got {type(df)}")

    # Write headers
    for i, col_name in enumerate(df.columns):
        worksheet.write(row, col + i, str(col_name), header_format)

    # Write data
    for r_idx, (_, row_data) in enumerate(df.iterrows()):
        for c_idx, value in enumerate(row_data):
            fmt = None
            if column_formats and df.columns[c_idx] in column_formats:
                fmt = column_formats[df.columns[c_idx]]
            worksheet.write(row + 1 + r_idx, col + c_idx, value, fmt)
