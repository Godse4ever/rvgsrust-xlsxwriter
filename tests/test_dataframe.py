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


# --- Regression tests for the unsafe Arrow PyCapsule handling ---
#
# write_dataframe() takes ownership of a C-level ArrowArrayStream
# struct out of a PyCapsule (record_batches_from_arrow() in
# src/lib.rs), which involves an unsafe ptr::read plus manually
# nulling the source struct's release field so the capsule's own
# destructor doesn't double-release the same underlying data (see the
# detailed comment there for the full Arrow C Data Interface
# reasoning). That's the highest-risk code in this crate: a mistake
# there is a potential double-free or use-after-free, not just a wrong
# answer. These tests exercise it repeatedly and across varied shapes
# to have some chance of catching a memory-safety regression, though
# a clean run here is not a substitute for testing with a sanitizer
# (ASAN/valgrind) on a real build -- pytest can't detect memory
# corruption that doesn't happen to crash or corrupt something we
# assert on.


@pytest.mark.skipif(not HAS_PYARROW, reason="PyArrow not installed")
def test_write_dataframe_repeated_calls_no_crash():
    table = pa.table({
        "id": pa.array(list(range(500)), type=pa.int64()),
        "value": pa.array([float(i) * 1.5 for i in range(500)], type=pa.float64()),
    })
    # Repeated capsule take-ownership cycles in one process -- a
    # double-free from the release-field bug would tend to show up as
    # a crash somewhere in this loop, not necessarily the first call.
    for _ in range(20):
        wb = Workbook()
        ws = wb.add_worksheet()
        ws.write_dataframe(0, 0, table)
        wb.close(TEST_FILE)
    # If we got here without segfaulting, that's the main signal this
    # test can give. Also check the last write actually has content.
    sheet = _load().active
    assert sheet["A2"].value == 0
    assert sheet["B2"].value == 0.0


@pytest.mark.skipif(not HAS_PYARROW, reason="PyArrow not installed")
def test_write_dataframe_multiple_worksheets_same_workbook():
    # Exercises taking ownership of two separate capsules (two
    # separate __arrow_c_stream__() calls) within the same workbook,
    # writing to different worksheets.
    table_a = pa.table({"x": pa.array([1, 2], type=pa.int64())})
    table_b = pa.table({"y": pa.array(["p", "q"], type=pa.string())})
    wb = Workbook()
    ws_a = wb.add_worksheet("A")
    ws_b = wb.add_worksheet("B")
    ws_a.write_dataframe(0, 0, table_a)
    ws_b.write_dataframe(0, 0, table_b)
    wb.close(TEST_FILE)
    book = _load()
    assert book["A"]["A2"].value == 1
    assert book["B"]["A2"].value == "p"


@pytest.mark.skipif(not HAS_PYARROW, reason="PyArrow not installed")
def test_write_dataframe_empty_table():
    # Zero rows: the capsule/stream is still created and consumed even
    # though there's no data to iterate -- makes sure the ownership
    # handling doesn't assume at least one batch exists.
    table = pa.table({"x": pa.array([], type=pa.int64())})
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write_dataframe(0, 0, table)  # should not raise
    wb.close(TEST_FILE)
    sheet = _load().active
    # No rows to write, but this shouldn't crash and shouldn't produce
    # spurious content either.
    assert sheet["A1"].value is None
