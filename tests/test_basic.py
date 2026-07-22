"""Basic functionality tests.

These tests open the generated .xlsx with openpyxl and assert on the
actual cell contents/formatting rather than only checking that a file
got created. A previous version of this library had a bug where
worksheet writes never reached the saved file at all (the file still
existed and was a valid, but essentially empty, workbook) -- a
file-existence-only test suite does not catch that class of bug.
"""
import os
import pytest
import openpyxl
from rvgsrust_xlsxwriter import Workbook

TEST_FILE = "test_output.xlsx"


def teardown_module(module):
    """Clean up test files."""
    if os.path.exists(TEST_FILE):
        os.remove(TEST_FILE)


def _load(path=TEST_FILE):
    return openpyxl.load_workbook(path)


def test_create_workbook():
    wb = Workbook()
    assert wb is not None


def test_add_worksheet():
    wb = Workbook()
    ws = wb.add_worksheet("TestSheet")
    assert ws is not None


def test_write_string():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write(0, 0, "Hello")
    ws.write(0, 1, "World")
    wb.close(TEST_FILE)
    sheet = _load().active
    assert sheet["A1"].value == "Hello"
    assert sheet["B1"].value == "World"


def test_write_number():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write(0, 0, 42)
    ws.write(0, 1, 3.14)
    wb.close(TEST_FILE)
    sheet = _load().active
    assert sheet["A1"].value == 42
    assert sheet["B1"].value == pytest.approx(3.14)


def test_write_boolean():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write(0, 0, True)
    ws.write(0, 1, False)
    wb.close(TEST_FILE)
    sheet = _load().active
    # Must round-trip as real booleans, not be silently coerced to 1.0/0.0
    # (bool is a subclass of int in Python, so this ordering is easy to
    # get wrong when dispatching on Python value type).
    assert sheet["A1"].value is True
    assert sheet["B1"].value is False


def test_write_row():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write_row(0, 0, ["A", "B", "C"])
    ws.write_row(1, 0, [1, 2, 3])
    wb.close(TEST_FILE)
    sheet = _load().active
    assert [c.value for c in sheet[1]] == ["A", "B", "C"]
    assert [c.value for c in sheet[2]] == [1, 2, 3]


def test_write_column():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write_column(0, 0, ["A", "B", "C"])
    ws.write_column(0, 1, [1, 2, 3])
    wb.close(TEST_FILE)
    sheet = _load().active
    assert [sheet.cell(row=r, column=1).value for r in (1, 2, 3)] == ["A", "B", "C"]
    assert [sheet.cell(row=r, column=2).value for r in (1, 2, 3)] == [1, 2, 3]


def test_format_bold():
    wb = Workbook()
    ws = wb.add_worksheet()
    fmt = wb.add_format()
    fmt.set_bold()
    ws.write(0, 0, "Bold Text", fmt)
    wb.close(TEST_FILE)
    sheet = _load().active
    assert sheet["A1"].value == "Bold Text"
    assert sheet["A1"].font.bold is True


def test_format_colors():
    wb = Workbook()
    ws = wb.add_worksheet()
    fmt = wb.add_format()
    fmt.set_background_color("#FFFF00")
    fmt.set_font_color("#FF0000")
    ws.write(0, 0, "Colored", fmt)
    wb.close(TEST_FILE)
    sheet = _load().active
    assert sheet["A1"].value == "Colored"
    assert sheet["A1"].fill.fgColor.rgb == "FFFFFF00"
    assert sheet["A1"].font.color.rgb == "FFFF0000"


def test_format_border():
    wb = Workbook()
    ws = wb.add_worksheet()
    fmt = wb.add_format()
    fmt.set_border("thin")
    fmt.set_border_color("#000000")
    ws.write(0, 0, "Bordered", fmt)
    wb.close(TEST_FILE)
    sheet = _load().active
    assert sheet["A1"].value == "Bordered"
    assert sheet["A1"].border.top.style == "thin"


def test_merge_range():
    wb = Workbook()
    ws = wb.add_worksheet()
    fmt = wb.add_format()
    fmt.set_bold()
    fmt.set_background_color("#4472C4")
    fmt.set_font_color("white")
    ws.merge_range(0, 0, 0, 2, "Merged Header", fmt)
    wb.close(TEST_FILE)
    book = _load()
    sheet = book.active
    assert sheet["A1"].value == "Merged Header"
    assert "A1:C1" in [str(r) for r in sheet.merged_cells.ranges]


def test_formula():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write(0, 0, 10)
    ws.write(0, 1, 20)
    ws.write_formula(0, 2, "=A1+B1")
    wb.close(TEST_FILE)
    sheet = _load().active
    assert sheet["C1"].value == "=A1+B1"


def test_freeze_panes():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.freeze_panes(1, 0)
    ws.write_row(0, 0, ["Header1", "Header2", "Header3"])
    ws.write_row(1, 0, [1, 2, 3])
    wb.close(TEST_FILE)
    sheet = _load().active
    assert sheet.freeze_panes == "A2"
    assert [c.value for c in sheet[1]] == ["Header1", "Header2", "Header3"]


def test_column_width():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.set_column_width(0, 20.0)
    ws.write(0, 0, "Wide Column")
    wb.close(TEST_FILE)
    sheet = _load().active
    assert sheet["A1"].value == "Wide Column"
    # Excel stores column width in its own padded units, not raw
    # character width, so an exact match isn't meaningful here -- just
    # confirm it was actually widened from the ~8.43 default.
    assert sheet.column_dimensions["A"].width > 15.0


def test_autofit():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write(0, 0, "Short")
    ws.write(0, 1, "This is a much longer text")
    ws.autofit()
    wb.close(TEST_FILE)
    sheet = _load().active
    assert sheet["A1"].value == "Short"
    assert sheet["B1"].value == "This is a much longer text"


def test_multiple_sheets():
    wb = Workbook()
    ws1 = wb.add_worksheet("Sheet1")
    ws2 = wb.add_worksheet("Sheet2")
    ws1.write(0, 0, "Data1")
    ws2.write(0, 0, "Data2")
    wb.close(TEST_FILE)
    book = _load()
    assert book["Sheet1"]["A1"].value == "Data1"
    # This is exactly the case the ownership bug broke: a second
    # worksheet's writes overwriting/vanishing relative to the first.
    assert book["Sheet2"]["A1"].value == "Data2"


def test_duplicate_sheet_name_raises():
    wb = Workbook()
    wb.add_worksheet("Sheet1")
    wb.add_worksheet("Sheet1")
    with pytest.raises(Exception):
        wb.close(TEST_FILE)


def test_write_out_of_bounds_raises():
    wb = Workbook()
    ws = wb.add_worksheet()
    with pytest.raises(Exception):
        ws.write(2_000_000, 0, "too far down")


def test_write_records_basic():
    wb = Workbook()
    ws = wb.add_worksheet()
    records = [
        {"Name": "Alice", "Age": 30, "Active": True},
        {"Name": "Bob", "Age": 25, "Active": False},
    ]
    ws.write_records(0, 0, records)
    wb.close(TEST_FILE)
    sheet = _load().active
    assert [c.value for c in sheet[1]] == ["Name", "Age", "Active"]
    assert [c.value for c in sheet[2]] == ["Alice", 30, True]
    assert [c.value for c in sheet[3]] == ["Bob", 25, False]


def test_write_records_explicit_headers_and_no_header_row():
    wb = Workbook()
    ws = wb.add_worksheet()
    records = [{"a": 1, "b": 2, "c": 3}]
    # Explicit headers control column order/subset; write_header=False
    # skips the header row entirely.
    ws.write_records(0, 0, records, headers=["c", "a"], write_header=False)
    wb.close(TEST_FILE)
    sheet = _load().active
    assert [c.value for c in sheet[1]] == [3, 1]


def test_write_records_with_formats():
    wb = Workbook()
    ws = wb.add_worksheet()
    header_fmt = wb.add_format()
    header_fmt.set_bold()
    ws.write_records(0, 0, [{"Name": "Alice"}], header_format=header_fmt)
    wb.close(TEST_FILE)
    sheet = _load().active
    assert sheet["A1"].value == "Name"
    assert sheet["A1"].font.bold is True
    assert sheet["A2"].value == "Alice"


def test_write_records_empty_list_is_noop():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write_records(0, 0, [])  # should not raise
    wb.close(TEST_FILE)
    sheet = _load().active
    assert sheet["A1"].value is None


def test_write_records_rejects_non_dict_records():
    wb = Workbook()
    ws = wb.add_worksheet()
    with pytest.raises(Exception):
        ws.write_records(0, 0, ["not", "a", "dict"])
