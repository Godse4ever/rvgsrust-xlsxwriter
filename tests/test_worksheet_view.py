"""Row/column visibility & sizing, view/zoom/selection, ignored errors,
autofit tuning, and NaN/infinity display strings.

Assertions are against the emitted XML, mirroring test_page_setup.py's
approach (openpyxl doesn't surface most of these conveniently).
"""
import os
import tempfile
import zipfile

import pytest

from rvgsrust_xlsxwriter import Workbook


def _xml(build, n_sheets=1):
    """Run build(ws) against the last of n_sheets worksheets, return
    (sheetN.xml, workbook.xml)."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        sheets = [wb.add_worksheet() for _ in range(n_sheets)]
        for ws in sheets:
            ws.write(0, 0, 1)
        build(*sheets)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            return (
                z.read(f"xl/worksheets/sheet{n_sheets}.xml").decode("utf-8"),
                z.read("xl/workbook.xml").decode("utf-8"),
            )
    finally:
        if os.path.exists(path):
            os.remove(path)


def _shared_strings(build):
    """Like _xml() but also returns xl/sharedStrings.xml, for values
    (e.g. NaN/Infinity) that rust_xlsxwriter stores as shared strings
    rather than inline in the sheet XML."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet()
        ws.write(0, 0, 1)
        build(ws)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            return z.read("xl/sharedStrings.xml").decode("utf-8")
    finally:
        if os.path.exists(path):
            os.remove(path)


# -------------------------- visibility & sizing --------------------------


def test_set_row_hidden():
    sheet, _ = _xml(lambda ws: ws.set_row_hidden(0))
    assert 'hidden="1"' in sheet


def test_set_row_unhidden_is_default_noop():
    # Row 0 was never hidden, so unhiding it is a no-op -- just confirm
    # it doesn't raise.
    sheet, _ = _xml(lambda ws: ws.set_row_unhidden(0))
    assert sheet


def test_set_column_hidden():
    sheet, _ = _xml(lambda ws: ws.set_column_hidden(0))
    assert 'hidden="1"' in sheet


def test_hide_unused_rows():
    sheet, _ = _xml(lambda ws: ws.hide_unused_rows(True))
    assert sheet


def test_set_row_height_pixels():
    sheet, _ = _xml(lambda ws: ws.set_row_height_pixels(0, 30))
    assert "<row " in sheet


def test_set_row_height_pixels_zero_hides_row():
    # Upstream treats height 0 as set_row_hidden().
    sheet, _ = _xml(lambda ws: ws.set_row_height_pixels(0, 0))
    assert 'hidden="1"' in sheet


def test_set_column_width_pixels():
    sheet, _ = _xml(lambda ws: ws.set_column_width_pixels(0, 100))
    assert "<col " in sheet


def test_set_default_row_height():
    sheet, _ = _xml(lambda ws: ws.set_default_row_height(30.0))
    assert 'defaultRowHeight="30"' in sheet


# ---------------------------- view / zoom / tabs ----------------------------


def test_set_zoom():
    sheet, _ = _xml(lambda ws: ws.set_zoom(200))
    assert 'zoomScale="200"' in sheet


def test_set_selection():
    sheet, _ = _xml(lambda ws: ws.set_selection(0, 0, 2, 2))
    assert "<selection" in sheet


def test_set_top_left_cell():
    sheet, _ = _xml(lambda ws: ws.set_top_left_cell(5, 5))
    assert 'topLeftCell="F6"' in sheet


def test_set_active():
    sheet, _ = _xml(lambda ws: ws.set_active(True))
    assert 'tabSelected="1"' in sheet


def test_set_first_tab():
    # Only visible in workbook.xml, and only when the flagged sheet isn't
    # already the first one.
    def build(ws1, ws2):
        ws2.set_first_tab(True)

    _, workbook_xml = _xml(build, n_sheets=2)
    assert 'firstSheet="2"' in workbook_xml


def test_set_right_to_left():
    sheet, _ = _xml(lambda ws: ws.set_right_to_left(True))
    assert 'rightToLeft="1"' in sheet


def test_set_view_page_layout():
    sheet, _ = _xml(lambda ws: ws.set_view_page_layout())
    assert 'view="pageLayout"' in sheet


def test_set_view_page_break_preview():
    sheet, _ = _xml(lambda ws: ws.set_view_page_break_preview())
    assert 'view="pageBreakPreview"' in sheet


def test_set_view_normal_is_default():
    sheet, _ = _xml(lambda ws: ws.set_view_normal())
    assert 'view="pageLayout"' not in sheet
    assert 'view="pageBreakPreview"' not in sheet


# ------------------------------ ignored errors ------------------------------


def test_ignore_error_number_stored_as_text():
    sheet, _ = _xml(lambda ws: ws.ignore_error(0, 0, "number_stored_as_text"))
    assert "<ignoredErrors>" in sheet
    assert 'numberStoredAsText="1"' in sheet


def test_ignore_error_formula_error():
    sheet, _ = _xml(lambda ws: ws.ignore_error(0, 0, "formula_error"))
    assert 'evalError="1"' in sheet


def test_ignore_error_range():
    sheet, _ = _xml(
        lambda ws: ws.ignore_error_range(0, 0, 2, 2, "formula_omits_cells")
    )
    assert 'formulaRange="1"' in sheet
    assert 'sqref="A1:C3"' in sheet


def test_ignore_error_invalid_type_raises():
    with pytest.raises(ValueError):
        _xml(lambda ws: ws.ignore_error(0, 0, "not_a_real_type"))


# ------------------------------ autofit tuning ------------------------------


def test_autofit_tuning_methods_do_not_raise():
    def build(ws):
        ws.set_autofit_max_width(300)
        ws.set_autofit_max_row(50)
        ws.autofit()

    sheet, _ = _xml(build)
    assert sheet


# ---------------------------- NaN / infinity strings ----------------------------


def test_nan_and_infinity_values():
    def build(ws):
        ws.set_nan_value("NaN")
        ws.set_infinity_value("Inf")
        ws.set_neg_infinity_value("-Inf")
        ws.write(1, 0, float("nan"))
        ws.write(2, 0, float("inf"))
        ws.write(3, 0, float("-inf"))

    # NaN/Infinity are stored as shared strings by rust_xlsxwriter (Excel
    # has no NaN/Inf numeric type), not inline in the sheet XML.
    shared = _shared_strings(build)
    assert "<t>NaN</t>" in shared
    assert "<t>Inf</t>" in shared
    assert "<t>-Inf</t>" in shared
