"""Cell notes: insert_note, show_all_notes, set_default_note_author."""
import os
import tempfile
import zipfile

from rvgsrust_xlsxwriter import Note, Workbook


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


def test_insert_note_text():
    files = _zip_contents(lambda ws: ws.insert_note(0, 0, Note("Some text")))
    assert "xl/comments1.xml" in files
    comments = _text(files, "xl/comments1.xml")
    assert "<t>Some text</t>" in comments
    sheet = _text(files, "xl/worksheets/sheet1.xml")
    assert "<legacyDrawing " in sheet


def test_insert_note_default_author():
    files = _zip_contents(lambda ws: ws.insert_note(0, 0, Note("Text")))
    comments = _text(files, "xl/comments1.xml")
    assert "<author>Author</author>" in comments


def test_insert_note_custom_author():
    note = Note("Text")
    note.set_author("Priya")
    files = _zip_contents(lambda ws: ws.insert_note(0, 0, note))
    comments = _text(files, "xl/comments1.xml")
    assert "<author>Priya</author>" in comments


def test_set_default_note_author():
    def build(ws):
        ws.set_default_note_author("Sam")
        ws.insert_note(0, 0, Note("Text"))

    files = _zip_contents(build)
    comments = _text(files, "xl/comments1.xml")
    assert "<author>Sam</author>" in comments


def test_note_visible_by_default_shows_no_hidden_marker():
    # Notes are hidden by default (shown on hover); set_visible(True)
    # marks it always shown.
    note = Note("Text")
    note.set_visible(True)
    files = _zip_contents(lambda ws: ws.insert_note(0, 0, note))
    vml_files = [n for n in files if n.startswith("xl/drawings/vmlDrawing")]
    assert vml_files
    assert 'visible' in _text(files, vml_files[0]).lower()


def test_show_all_notes_does_not_raise():
    def build(ws):
        ws.insert_note(0, 0, Note("Text"))
        ws.show_all_notes(True)

    files = _zip_contents(build)
    assert "xl/comments1.xml" in files


def test_note_width_height_alt_text():
    note = Note("Text")
    note.set_width(200)
    note.set_height(100)
    note.set_alt_text("A note")
    files = _zip_contents(lambda ws: ws.insert_note(0, 0, note))
    assert "xl/comments1.xml" in files


def test_note_font_and_background_color():
    note = Note("Text")
    note.set_font_name("Arial")
    note.set_font_size(12)
    note.set_background_color("#FFC7CE")
    files = _zip_contents(lambda ws: ws.insert_note(0, 0, note))
    vml_files = [n for n in files if n.startswith("xl/drawings/vmlDrawing")]
    assert "Arial" in _text(files, vml_files[0])
