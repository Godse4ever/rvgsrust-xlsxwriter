"""Checkboxes, Form Control buttons, and Textbox shapes."""
import os
import tempfile
import zipfile

from rvgsrust_xlsxwriter import Button, Format, Shape, Workbook


def _zip_contents(build):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet()
        build(ws)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            return {name: z.read(name) for name in z.namelist()}
    finally:
        if os.path.exists(path):
            os.remove(path)


def _text(files, name):
    return files[name].decode("utf-8")


# -------------------------------- checkbox --------------------------------


def test_insert_checkbox_true():
    files = _zip_contents(lambda ws: ws.insert_checkbox(0, 0, True))
    sheet = _text(files, "xl/worksheets/sheet1.xml")
    assert '<c r="A1" t="b"><v>1</v></c>' in sheet
    styles = _text(files, "xl/styles.xml")
    assert "xfpb:xfComplement" in styles


def test_insert_checkbox_false():
    files = _zip_contents(lambda ws: ws.insert_checkbox(0, 0, False))
    sheet = _text(files, "xl/worksheets/sheet1.xml")
    assert '<c r="A1" t="b"><v>0</v></c>' in sheet


def test_insert_checkbox_with_format():
    def build(ws):
        fmt = Format()
        fmt.set_background_color("FFC7CE")
        ws.insert_checkbox(0, 0, True, format=fmt)

    files = _zip_contents(build)
    sheet = _text(files, "xl/worksheets/sheet1.xml")
    assert '<c r="A1" t="b"><v>1</v></c>' in sheet
    styles = _text(files, "xl/styles.xml")
    assert "FFC7CE" in styles


# --------------------------------- button ---------------------------------


def test_insert_button_default_caption():
    files = _zip_contents(lambda ws: ws.insert_button(0, 0, Button()))
    sheet = _text(files, "xl/worksheets/sheet1.xml")
    assert "<legacyDrawing " in sheet
    vml_files = [n for n in files if n.startswith("xl/drawings/vmlDrawing")]
    assert vml_files
    assert "Button 1" in _text(files, vml_files[0])


def test_insert_button_custom_caption():
    button = Button()
    button.set_caption("Press Me")
    files = _zip_contents(lambda ws: ws.insert_button(0, 0, button))
    vml_files = [n for n in files if n.startswith("xl/drawings/vmlDrawing")]
    assert "Press Me" in _text(files, vml_files[0])


def test_insert_button_with_offset():
    button = Button()
    files = _zip_contents(lambda ws: ws.insert_button_with_offset(0, 0, button, 5, 10))
    sheet = _text(files, "xl/worksheets/sheet1.xml")
    assert "<legacyDrawing " in sheet


# ---------------------------------- shape ----------------------------------


def test_insert_shape_textbox_text():
    shape = Shape.textbox()
    shape.set_text("Hello from a textbox")
    files = _zip_contents(lambda ws: ws.insert_shape(0, 0, shape))
    drawing_files = [n for n in files if n.startswith("xl/drawings/drawing")]
    assert drawing_files
    assert "<a:t>Hello from a textbox</a:t>" in _text(files, drawing_files[0])


def test_insert_shape_with_offset():
    shape = Shape.textbox()
    shape.set_text("Offset box")
    files = _zip_contents(lambda ws: ws.insert_shape_with_offset(0, 0, shape, 5, 10))
    sheet = _text(files, "xl/worksheets/sheet1.xml")
    assert "<drawing " in sheet


def test_shape_width_and_height_do_not_raise():
    shape = Shape.textbox()
    shape.set_text("Sized box")
    shape.set_width(300)
    shape.set_height(150)
    files = _zip_contents(lambda ws: ws.insert_shape(0, 0, shape))
    assert any(n.startswith("xl/drawings/drawing") for n in files)
