"""Image placement: insert_image_with_offset, embed_image(_with_format),
insert_image_fit_to_cell(_centered), insert_background_image.
insert_image() itself already had coverage in test_basic.py.
"""
import base64
import os
import tempfile
import zipfile

import pytest

from rvgsrust_xlsxwriter import Format, Workbook

# Smallest possible valid 1x1 transparent PNG.
_MINIMAL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNk"
    "+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


@pytest.fixture
def image_path():
    path = os.path.join(tempfile.gettempdir(), "rvgsrust_test_1x1.png")
    with open(path, "wb") as f:
        f.write(_MINIMAL_PNG)
    yield path
    if os.path.exists(path):
        os.remove(path)


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
            names = z.namelist()
            sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
            return sheet, names
    finally:
        if os.path.exists(path):
            os.remove(path)


def test_insert_image_with_offset(image_path):
    sheet, names = _sheet_xml(
        lambda ws: ws.insert_image_with_offset(1, 1, image_path, 5, 10)
    )
    assert any(n.startswith("xl/media/") for n in names)
    assert "<drawing " in sheet


def test_insert_image_with_offset_missing_file_raises():
    wb = Workbook()
    ws = wb.add_worksheet()
    with pytest.raises(OSError, match="not found"):
        ws.insert_image_with_offset(0, 0, "/tmp/does_not_exist_xyz.png", 0, 0)


def test_embed_image(image_path):
    sheet, names = _sheet_xml(lambda ws: ws.embed_image(0, 0, image_path))
    assert any(n.startswith("xl/media/") for n in names)
    assert "<drawing " in sheet


def test_embed_image_with_format(image_path):
    def build(ws):
        fmt = Format()
        ws.embed_image_with_format(0, 0, image_path, fmt)

    sheet, names = _sheet_xml(build)
    assert any(n.startswith("xl/media/") for n in names)


def test_insert_image_fit_to_cell(image_path):
    sheet, names = _sheet_xml(
        lambda ws: ws.insert_image_fit_to_cell(0, 0, image_path, keep_aspect_ratio=False)
    )
    assert any(n.startswith("xl/media/") for n in names)
    assert "<drawing " in sheet


def test_insert_image_fit_to_cell_centered(image_path):
    sheet, names = _sheet_xml(
        lambda ws: ws.insert_image_fit_to_cell_centered(0, 0, image_path)
    )
    assert any(n.startswith("xl/media/") for n in names)


def test_insert_background_image(image_path):
    sheet, names = _sheet_xml(lambda ws: ws.insert_background_image(image_path))
    assert any(n.startswith("xl/media/") for n in names)
    assert "<picture " in sheet


def test_insert_background_image_missing_file_raises():
    wb = Workbook()
    ws = wb.add_worksheet()
    with pytest.raises(OSError, match="not found"):
        ws.insert_background_image("/tmp/does_not_exist_xyz.png")
