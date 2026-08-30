"""Workbook flat setters: themes, default format, VBA projects,
read_only_recommended, tempdir/large-zip.
"""
import os
import tempfile
import zipfile

import pytest

from rvgsrust_xlsxwriter import Format, Workbook


def _zip_contents(build):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        build(wb)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            return {name: z.read(name) for name in z.namelist()}
    finally:
        if os.path.exists(path):
            os.remove(path)


def _text(files, name):
    return files[name].decode("utf-8")


# --------------------------------- themes ---------------------------------


def test_use_excel_2023_theme():
    def build(wb):
        wb.use_excel_2023_theme()
        ws = wb.add_worksheet()
        ws.write(0, 0, "Hello")

    files = _zip_contents(build)
    theme = _text(files, "xl/theme/theme1.xml")
    assert "Aptos" in theme


def test_use_excel_2023_theme_after_worksheet_raises():
    def build(wb):
        wb.add_worksheet()
        wb.use_excel_2023_theme()

    with pytest.raises(ValueError):
        _zip_contents(build)


def test_use_custom_theme_invalid_file_raises():
    path = os.path.join(tempfile.gettempdir(), "rvgsrust_not_a_theme.xml")
    with open(path, "w") as f:
        f.write("not xml at all")
    try:
        with pytest.raises(ValueError):
            _zip_contents(lambda wb: wb.use_custom_theme(path))
    finally:
        os.remove(path)


# ----------------------------- default format -----------------------------


def test_set_default_format():
    def build(wb):
        fmt = Format()
        fmt.set_font_name("Georgia")
        wb.set_default_format(fmt, 18, 72)
        ws = wb.add_worksheet()
        ws.write(0, 0, "Hi")

    files = _zip_contents(build)
    styles = _text(files, "xl/styles.xml")
    assert '<name val="Georgia"/>' in styles


def test_set_default_format_after_worksheet_raises():
    def build(wb):
        wb.add_worksheet()
        wb.set_default_format(Format(), 18, 72)

    with pytest.raises(ValueError):
        _zip_contents(build)


def test_set_default_format_unsupported_width_raises():
    def build(wb):
        wb.set_default_format(Format(), 15, 999)

    with pytest.raises(ValueError):
        _zip_contents(build)


# --------------------------------- VBA ---------------------------------


def _dummy_vba_path():
    path = os.path.join(tempfile.gettempdir(), "rvgsrust_dummy_vba.bin")
    with open(path, "wb") as f:
        f.write(b"\x00\x01\x02not a real OLE file, upstream doesn't validate")
    return path


def test_add_vba_project():
    path = _dummy_vba_path()
    try:
        def build(wb):
            wb.add_vba_project(path)
            wb.add_worksheet()

        files = _zip_contents(build)
        assert "xl/vbaProject.bin" in files
        sheet = _text(files, "xl/worksheets/sheet1.xml")
        assert 'codeName="{37E998C4-C9E5-D4B9-71C8-EB1FF731991C}"' in sheet
    finally:
        os.remove(path)


def test_add_vba_project_missing_file_raises():
    with pytest.raises(OSError):
        _zip_contents(lambda wb: wb.add_vba_project("/tmp/does_not_exist_xyz.bin"))


def test_set_vba_name():
    def build(wb):
        wb.set_vba_name("MyWorkbook")
        wb.add_worksheet()

    files = _zip_contents(build)
    assert "xl/worksheets/sheet1.xml" in files


def test_worksheet_set_vba_name():
    def build(wb):
        ws = wb.add_worksheet()
        ws.set_vba_name("MySheet")
        ws.write(0, 0, 1)

    files = _zip_contents(build)
    assert "xl/worksheets/sheet1.xml" in files


# ------------------------ read-only, tempdir, large-zip ------------------------


def test_read_only_recommended():
    def build(wb):
        wb.add_worksheet()
        wb.read_only_recommended()

    files = _zip_contents(build)
    workbook_xml = _text(files, "xl/workbook.xml")
    assert '<fileSharing readOnlyRecommended="1"/>' in workbook_xml


def test_set_tempdir_does_not_raise():
    def build(wb):
        wb.set_tempdir(tempfile.gettempdir())
        ws = wb.add_worksheet()
        ws.write(0, 0, "Hi")

    files = _zip_contents(build)
    assert "xl/worksheets/sheet1.xml" in files


def test_use_zip_large_file_does_not_raise():
    def build(wb):
        wb.use_zip_large_file(True)
        ws = wb.add_worksheet()
        ws.write(0, 0, "Hi")

    files = _zip_contents(build)
    assert "xl/worksheets/sheet1.xml" in files
