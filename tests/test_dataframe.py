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


try:
    import pyarrow as pa
    HAS_PYARROW = True
except ImportError:
    HAS_PYARROW = False

import openpyxl


def _load(path=TEST_FILE):
    return openpyxl.load_workbook(path)


@pytest.mark.skipif(not HAS_PYARROW, reason="PyArrow not installed")
def test_write_dataframe_pyarrow_table():
    table = pa.table({
        "id": pa.array([1, 2, 3], type=pa.int64()),
        "name": pa.array(["Alice", "Bob", "Carol"], type=pa.string()),
        "score": pa.array([95.5, 88.2, 71.0], type=pa.float64()),
        "passed": pa.array([True, False, True], type=pa.bool_()),
    })
    wb = Workbook()
    ws = wb.add_worksheet()
    fmt = wb.add_format()
    fmt.set_bold()
    ws.write_dataframe(0, 0, table, header_format=fmt)
    wb.close(TEST_FILE)

    sheet = _load().active
    assert [c.value for c in sheet[1]] == ["id", "name", "score", "passed"]
    assert sheet["A1"].font.bold is True
    assert [c.value for c in sheet[2]] == [1, "Alice", 95.5, True]


@pytest.mark.skipif(not HAS_PYARROW, reason="PyArrow not installed")
def test_write_dataframe_nulls():
    table = pa.table({
        "a": pa.array([1, None, 3], type=pa.int64()),
        "b": pa.array(["x", "y", None], type=pa.string()),
    })
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write_dataframe(0, 0, table)
    wb.close(TEST_FILE)

    sheet = _load().active
    assert sheet["A3"].value is None  # the None in column a
    assert sheet["B4"].value is None  # the None in column b


@pytest.mark.skipif(not HAS_PANDAS, reason="Pandas not installed")
def test_write_dataframe_pandas_large_utf8():
    # Pandas commonly exports string columns as Arrow LargeUtf8, not
    # Utf8 -- this must work too, not just pyarrow's default Utf8.
    df = pd.DataFrame({"x": [1, 2, 3], "y": ["a", "b", "c"]})
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write_dataframe(0, 0, df)
    wb.close(TEST_FILE)

    sheet = _load().active
    assert [c.value for c in sheet[1]] == ["x", "y"]
    assert [c.value for c in sheet[2]] == [1, "a"]


@pytest.mark.skipif(not HAS_PYARROW, reason="PyArrow not installed")
def test_write_dataframe_unsupported_type_raises():
    table = pa.table({"d": pa.array([1, 2], type=pa.int32())})
    wb = Workbook()
    ws = wb.add_worksheet()
    with pytest.raises(TypeError):
        ws.write_dataframe(0, 0, table)


def test_write_dataframe_rejects_non_dataframe():
    wb = Workbook()
    ws = wb.add_worksheet()
    with pytest.raises(TypeError):
        ws.write_dataframe(0, 0, {"not": "a dataframe"})
