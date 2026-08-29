"""Worksheet protection: protect_with_options() and unprotect_range().
protect()/protect_with_password() already had coverage elsewhere.
"""
import os
import tempfile
import zipfile

from rvgsrust_xlsxwriter import ProtectionOptions, Workbook


def _sheet_xml(build):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet()
        ws.write(0, 0, 1)
        build(ws)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            return z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    finally:
        if os.path.exists(path):
            os.remove(path)


def test_protect_with_options_defaults():
    sheet = _sheet_xml(lambda ws: ws.protect_with_options(ProtectionOptions()))
    assert "<sheetProtection " in sheet
    assert 'sheet="1"' in sheet
    # select_locked_cells/select_unlocked_cells default to permissive
    # (True), so those two stay absent. edit_objects/edit_scenarios
    # default to restrictive (False), so "objects"/"scenarios" DO appear
    # by default -- covered by test_protect_with_options_allow_edit below.
    assert 'selectLockedCells="1"' not in sheet
    assert 'selectUnlockedCells="1"' not in sheet
    assert 'objects="1"' in sheet
    assert 'scenarios="1"' in sheet


def test_protect_with_options_allow_edit():
    options = ProtectionOptions(edit_objects=True, edit_scenarios=True)
    sheet = _sheet_xml(lambda ws: ws.protect_with_options(options))
    assert 'objects="1"' not in sheet
    assert 'scenarios="1"' not in sheet


def test_protect_with_options_allow_insert():
    options = ProtectionOptions(insert_columns=True, insert_rows=True)
    sheet = _sheet_xml(lambda ws: ws.protect_with_options(options))
    assert 'insertColumns="0"' in sheet
    assert 'insertRows="0"' in sheet


def test_protect_with_options_and_password():
    options = ProtectionOptions()
    sheet = _sheet_xml(
        lambda ws: ws.protect_with_options(options, password="secret")
    )
    assert 'password="' in sheet


def test_protect_with_options_disallow_selection():
    options = ProtectionOptions(select_locked_cells=False, select_unlocked_cells=False)
    sheet = _sheet_xml(lambda ws: ws.protect_with_options(options))
    assert 'selectLockedCells="1"' in sheet
    assert 'selectUnlockedCells="1"' in sheet


def test_unprotect_range_default_name():
    sheet = _sheet_xml(lambda ws: ws.unprotect_range(0, 0, 2, 2))
    assert "<protectedRanges>" in sheet
    assert 'sqref="A1:C3"' in sheet
    assert 'name="Range1"' in sheet


def test_unprotect_range_custom_name():
    sheet = _sheet_xml(lambda ws: ws.unprotect_range(0, 0, 2, 2, name="MyRange"))
    assert 'name="MyRange"' in sheet


def test_unprotect_range_with_password():
    sheet = _sheet_xml(
        lambda ws: ws.unprotect_range(0, 0, 2, 2, password="secret")
    )
    assert '<protectedRange password="' in sheet


def test_unprotect_range_does_not_require_protect():
    # protectedRanges is independent of sheetProtection/protect().
    sheet = _sheet_xml(lambda ws: ws.unprotect_range(0, 0, 1, 1))
    assert "<protectedRanges>" in sheet
    assert "<sheetProtection" not in sheet
