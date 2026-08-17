"""Page setup and print settings.

Assertions are against the emitted XML, since none of these produce
anything openpyxl surfaces conveniently. Print area and repeat rows or
columns become defined names in workbook.xml rather than sheet properties.
"""
import os
import tempfile
import zipfile

import pytest

from rvgsrust_xlsxwriter import Workbook


def _xml(build):
    """Run build(ws), return (sheet1.xml, workbook.xml)."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet()
        ws.write(0, 0, 1)
        build(ws)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            return (
                z.read("xl/worksheets/sheet1.xml").decode("utf-8"),
                z.read("xl/workbook.xml").decode("utf-8"),
            )
    finally:
        if os.path.exists(path):
            os.remove(path)


# ---------------------------- orientation ----------------------------


def test_landscape():
    sheet, _ = _xml(lambda ws: ws.set_landscape())
    assert 'orientation="landscape"' in sheet


def test_portrait():
    sheet, _ = _xml(lambda ws: ws.set_portrait())
    assert 'orientation="portrait"' in sheet


def test_paper_size():
    # 9 is A4 in Excel's numeric paper size codes.
    sheet, _ = _xml(lambda ws: ws.set_paper_size(9))
    assert 'paperSize="9"' in sheet


def test_page_order_over_then_down():
    """True is Excel's default and writes nothing; False emits the
    attribute."""
    sheet, _ = _xml(lambda ws: ws.set_page_order(False))
    assert 'pageOrder="overThenDown"' in sheet


def test_page_order_default_writes_nothing():
    sheet, _ = _xml(lambda ws: ws.set_page_order(True))
    assert "pageOrder" not in sheet


# ------------------------------ margins ------------------------------


def test_margins():
    sheet, _ = _xml(lambda ws: ws.set_margins(0.5, 0.5, 0.75, 0.75, 0.3, 0.3))
    assert "<pageMargins" in sheet
    assert 'left="0.5"' in sheet
    assert 'top="0.75"' in sheet


# --------------------------- area and titles ---------------------------


def test_print_area_becomes_a_defined_name():
    _, book = _xml(lambda ws: ws.set_print_area(0, 0, 9, 4))
    assert "Print_Area" in book


def test_repeat_rows_becomes_print_titles():
    _, book = _xml(lambda ws: ws.set_repeat_rows(0, 1))
    assert "Print_Titles" in book


def test_repeat_columns_becomes_print_titles():
    _, book = _xml(lambda ws: ws.set_repeat_columns(0, 1))
    assert "Print_Titles" in book


def test_print_area_reversed_range_raises():
    with pytest.raises(ValueError):
        _xml(lambda ws: ws.set_print_area(9, 4, 0, 0))


# ------------------------------ scaling ------------------------------


def test_print_fit_to_pages():
    sheet, _ = _xml(lambda ws: ws.set_print_fit_to_pages(2, 1))
    assert 'fitToPage="1"' in sheet
    # fitToHeight is only written when it differs from 1, so only width
    # is asserted here.
    assert 'fitToWidth="2"' in sheet


def test_print_scale():
    sheet, _ = _xml(lambda ws: ws.set_print_scale(80))
    assert 'scale="80"' in sheet


# --------------------------- page breaks ---------------------------


def test_horizontal_page_breaks():
    sheet, _ = _xml(lambda ws: ws.set_page_breaks([5, 10]))
    assert "<rowBreaks" in sheet
    assert "<brk" in sheet


def test_vertical_page_breaks():
    """Upstream takes u32 column numbers here, unlike the u16 used
    elsewhere."""
    sheet, _ = _xml(lambda ws: ws.set_vertical_page_breaks([3, 6]))
    assert "<colBreaks" in sheet


def test_empty_page_breaks_is_accepted():
    sheet, _ = _xml(lambda ws: ws.set_page_breaks([]))
    assert "<worksheet" in sheet


# --------------------------- print options ---------------------------


def test_print_gridlines():
    sheet, _ = _xml(lambda ws: ws.set_print_gridlines(True))
    assert 'gridLines="1"' in sheet


def test_print_headings():
    sheet, _ = _xml(lambda ws: ws.set_print_headings(True))
    assert 'headings="1"' in sheet


def test_print_centered():
    def build(ws):
        ws.set_print_center_horizontally(True)
        ws.set_print_center_vertically(True)

    sheet, _ = _xml(build)
    assert 'horizontalCentered="1"' in sheet
    assert 'verticalCentered="1"' in sheet


def test_print_black_and_white():
    sheet, _ = _xml(lambda ws: ws.set_print_black_and_white(True))
    assert 'blackAndWhite="1"' in sheet


def test_print_draft():
    sheet, _ = _xml(lambda ws: ws.set_print_draft(True))
    assert 'draft="1"' in sheet


def test_print_first_page_number():
    """Pins upstream's actual output, which looks wrong.

    rust_xlsxwriter writes the page number into the useFirstPageNumber
    attribute (worksheet.rs:18697) and never emits a firstPageNumber
    attribute. In ECMA-376 those are distinct: useFirstPageNumber is a
    boolean and firstPageNumber is the uint holding the value. Excel reads
    a nonzero boolean as true, so the feature switches on, but the number
    itself most likely falls back to 1.

    This test asserts what is written, not what ought to be, so it will
    fail loudly if upstream fixes it -- which is the point.
    """
    sheet, _ = _xml(lambda ws: ws.set_print_first_page_number(5))
    assert 'useFirstPageNumber="5"' in sheet
    assert "firstPageNumber=" not in sheet.replace("useFirstPageNumber=", "")


# ---------------------------- combination ----------------------------


def test_a_realistic_report_page_setup():
    def build(ws):
        ws.set_landscape()
        ws.set_paper_size(9)
        ws.set_margins(0.5, 0.5, 0.75, 0.75, 0.3, 0.3)
        ws.set_print_area(0, 0, 49, 7)
        ws.set_repeat_rows(0, 0)
        ws.set_print_fit_to_pages(1, 0)
        ws.set_print_gridlines(True)
        ws.set_print_headings(False)
        ws.set_print_center_horizontally(True)

    sheet, book = _xml(build)
    assert 'orientation="landscape"' in sheet
    assert 'paperSize="9"' in sheet
    assert 'fitToPage="1"' in sheet
    assert 'horizontalCentered="1"' in sheet
    assert "Print_Area" in book
    assert "Print_Titles" in book


def test_page_setup_works_in_constant_memory_mode():
    """These set sheet metadata, not cells, so they are not affected by the
    row-order guard."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet(constant_memory=True)
        ws.write(10, 0, 1)
        ws.set_landscape()
        ws.set_repeat_rows(0, 0)
        ws.set_print_area(0, 0, 20, 4)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert 'orientation="landscape"' in sheet
    finally:
        if os.path.exists(path):
            os.remove(path)


# ------------------------- headers / footers -------------------------


def test_set_header_and_footer_basic():
    def build(ws):
        ws.set_header("&CMy Report")
        ws.set_footer("&LConfidential")

    sheet, _ = _xml(build)
    assert "<oddHeader>" in sheet
    assert "My Report" in sheet
    assert "<oddFooter>" in sheet
    assert "Confidential" in sheet


def test_set_header_placeholders_pass_through_unmodified():
    # &[Page]/&[Pages]/&[File]/&[Tab] are Excel placeholders this binding
    # doesn't interpret -- they should survive verbatim (XML-escaped) so
    # Excel expands them when the file is opened.
    sheet, _ = _xml(
        lambda ws: ws.set_header("&LPage &[Page] of &[Pages]&RSheet: &[Tab]")
    )
    assert "Page &amp;[Page] of &amp;[Pages]" in sheet
    assert "Sheet: &amp;[Tab]" in sheet


def test_set_footer_only_does_not_require_header():
    sheet, _ = _xml(lambda ws: ws.set_footer("&CPage &P"))
    assert "<oddFooter>" in sheet
    assert "<oddHeader>" not in sheet


def test_header_footer_work_in_constant_memory_mode():
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet(constant_memory=True)
        ws.write(0, 0, 1)
        ws.set_header("&CHeader text")
        ws.set_footer("&CFooter text")
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert "Header text" in sheet
        assert "Footer text" in sheet
    finally:
        if os.path.exists(path):
            os.remove(path)

