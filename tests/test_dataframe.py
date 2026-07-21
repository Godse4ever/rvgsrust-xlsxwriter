"""DataFrame support tests."""
import pytest
from rvgsrust_xlsxwriter import Workbook
from rvgsrust_xlsxwriter.dataframe import write_polars_dataframe, write_pandas_dataframe

TEST_FILE = "test_df_output.xlsx"


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


@pytest.mark.skipif(not HAS_POLARS, reason="Polars not installed")
def test_polars_dataframe():
    df = pl.DataFrame({
        "Name": ["Alice", "Bob", "Charlie"],
        "Age": [30, 25, 35],
        "Salary": [50000.0, 60000.0, 70000.0],
    })

    wb = Workbook()
    ws = wb.add_worksheet()

    header_fmt = wb.add_format()
    header_fmt.set_bold()
    header_fmt.set_background_color("#4472C4")
    header_fmt.set_font_color("white")

    write_polars_dataframe(ws, df, row=0, col=0, header_format=header_fmt)
    ws.autofit()
    wb.close(TEST_FILE)


@pytest.mark.skipif(not HAS_PANDAS, reason="Pandas not installed")
def test_pandas_dataframe():
    df = pd.DataFrame({
        "Product": ["Widget", "Gadget", "Thingama"],
        "Price": [9.99, 19.99, 29.99],
        "Stock": [100, 50, 25],
    })

    wb = Workbook()
    ws = wb.add_worksheet()

    header_fmt = wb.add_format()
    header_fmt.set_bold()
    header_fmt.set_background_color("#70AD47")
    header_fmt.set_font_color("white")

    write_pandas_dataframe(ws, df, row=0, col=0, header_format=header_fmt)
    ws.autofit()
    wb.close(TEST_FILE)
