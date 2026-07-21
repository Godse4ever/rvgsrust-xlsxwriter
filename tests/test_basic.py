"""Basic functionality tests."""
import os
import pytest
from rvgsrust_xlsxwriter import Workbook

TEST_FILE = "test_output.xlsx"


def teardown_module(module):
    """Clean up test files."""
    if os.path.exists(TEST_FILE):
        os.remove(TEST_FILE)


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
    assert os.path.exists(TEST_FILE)


def test_write_number():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write(0, 0, 42)
    ws.write(0, 1, 3.14)
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_write_boolean():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write(0, 0, True)
    ws.write(0, 1, False)
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_write_row():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write_row(0, 0, ["A", "B", "C"])
    ws.write_row(1, 0, [1, 2, 3])
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_write_column():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write_column(0, 0, ["A", "B", "C"])
    ws.write_column(0, 1, [1, 2, 3])
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_format_bold():
    wb = Workbook()
    ws = wb.add_worksheet()
    fmt = wb.add_format()
    fmt.set_bold()
    ws.write(0, 0, "Bold Text", fmt)
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_format_colors():
    wb = Workbook()
    ws = wb.add_worksheet()
    fmt = wb.add_format()
    fmt.set_background_color("#FFFF00")
    fmt.set_font_color("#FF0000")
    ws.write(0, 0, "Colored", fmt)
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_format_border():
    wb = Workbook()
    ws = wb.add_worksheet()
    fmt = wb.add_format()
    fmt.set_border("thin")
    fmt.set_border_color("#000000")
    ws.write(0, 0, "Bordered", fmt)
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_merge_range():
    wb = Workbook()
    ws = wb.add_worksheet()
    fmt = wb.add_format()
    fmt.set_bold()
    fmt.set_background_color("#4472C4")
    fmt.set_font_color("white")
    ws.merge_range(0, 0, 0, 2, "Merged Header", fmt)
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_formula():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write(0, 0, 10)
    ws.write(0, 1, 20)
    ws.write_formula(0, 2, "=A1+B1")
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_freeze_panes():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.freeze_panes(1, 0)
    ws.write_row(0, 0, ["Header1", "Header2", "Header3"])
    ws.write_row(1, 0, [1, 2, 3])
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_column_width():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.set_column_width(0, 20.0)
    ws.write(0, 0, "Wide Column")
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_autofit():
    wb = Workbook()
    ws = wb.add_worksheet()
    ws.write(0, 0, "Short")
    ws.write(0, 1, "This is a much longer text")
    ws.autofit()
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)


def test_multiple_sheets():
    wb = Workbook()
    ws1 = wb.add_worksheet("Sheet1")
    ws2 = wb.add_worksheet("Sheet2")
    ws1.write(0, 0, "Data1")
    ws2.write(0, 0, "Data2")
    wb.close(TEST_FILE)
    assert os.path.exists(TEST_FILE)
