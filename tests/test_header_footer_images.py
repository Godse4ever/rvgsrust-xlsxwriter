"""Header/footer images and align/scale-with-page settings.
set_header(text)/set_footer(text) already had coverage elsewhere.
"""
import base64
import os
import tempfile
import zipfile

import pytest

from rvgsrust_xlsxwriter import Workbook

_MINIMAL_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNk"
    "+A8AAQUBAScY42YAAAAASUVORK5CYII="
)


@pytest.fixture
def image_path():
    path = os.path.join(tempfile.gettempdir(), "rvgsrust_test_hf_image.png")
    with open(path, "wb") as f:
        f.write(_MINIMAL_PNG)
    yield path
    if os.path.exists(path):
        os.remove(path)


def _zip_contents(build):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet()
        ws.write(0, 0, 1)
        build(ws)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            return {name: z.read(name) for name in z.namelist()}
    finally:
        if os.path.exists(path):
            os.remove(path)


def _text(files, name):
    return files[name].decode("utf-8")


def test_set_header_image_left(image_path):
    def build(ws):
        ws.set_header("&L&[Picture]")
        ws.set_header_image(image_path, "left")

    files = _zip_contents(build)
    assert any(n.startswith("xl/media/") for n in files)
    sheet = _text(files, "xl/worksheets/sheet1.xml")
    assert "<legacyDrawingHF " in sheet


def test_set_header_image_without_placeholder_raises(image_path):
    def build(ws):
        ws.set_header("no placeholder here")
        ws.set_header_image(image_path, "left")

    with pytest.raises(ValueError):
        _zip_contents(build)


def test_set_footer_image_right(image_path):
    def build(ws):
        ws.set_footer("&R&[Picture]")
        ws.set_footer_image(image_path, "right")

    files = _zip_contents(build)
    assert any(n.startswith("xl/media/") for n in files)
    sheet = _text(files, "xl/worksheets/sheet1.xml")
    assert "<legacyDrawingHF " in sheet


def test_set_header_image_invalid_position_raises(image_path):
    def build(ws):
        ws.set_header("&C&[Picture]")
        ws.set_header_image(image_path, "sideways")

    with pytest.raises(ValueError):
        _zip_contents(build)


def test_header_footer_scale_and_align_do_not_raise():
    def build(ws):
        ws.set_header_footer_scale_with_doc(False)
        ws.set_header_footer_align_with_page(False)

    files = _zip_contents(build)
    assert "xl/worksheets/sheet1.xml" in files
