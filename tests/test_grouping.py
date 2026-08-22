"""Row/column outline grouping tests.

XML attribute names (outlineLevel, collapsed, summaryBelow, summaryRight)
verified against the actual rust_xlsxwriter 0.98.2 source
(worksheet.rs write_row/write_col/write_outline_pr) before writing these
assertions, not guessed.
"""
import os
import tempfile
import zipfile

import pytest

from rvgsrust_xlsxwriter import Workbook


def _xml(build):
    """Run build(ws), return sheet1.xml."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet()
        for r in range(10):
            ws.write(r, 0, r)
        build(ws)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            return z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    finally:
        if os.path.exists(path):
            os.remove(path)


# ---------------------------- rows ----------------------------


def test_group_rows_sets_outline_level():
    def build(ws):
        ws.group_rows(1, 5)

    sheet = _xml(build)
    assert 'outlineLevel="1"' in sheet


def test_group_rows_nested_increments_level():
    def build(ws):
        # Outer range first, then a nested inner range -- the normal
        # way multi-level outlines are built. Must succeed in either
        # call order; see check_row_order_readonly()'s comment in
        # lib.rs for why a naive row-order guard would incorrectly
        # reject this specific, common pattern.
        ws.group_rows(0, 9)
        ws.group_rows(1, 5)

    sheet = _xml(build)
    assert 'outlineLevel="2"' in sheet
    assert 'outlineLevel="1"' in sheet


def test_group_rows_collapsed_sets_collapsed_flag():
    def build(ws):
        ws.group_rows_collapsed(1, 5)

    sheet = _xml(build)
    assert 'outlineLevel="1"' in sheet
    assert 'collapsed="1"' in sheet


def test_group_rows_rejects_reversed_range():
    def build(ws):
        ws.group_rows(5, 1)

    with pytest.raises(ValueError):
        _xml(build)


def test_group_rows_exceeding_max_level_raises():
    def build(ws):
        # Excel allows at most 7 nested outline levels.
        for _ in range(8):
            ws.group_rows(1, 5)

    with pytest.raises(ValueError):
        _xml(build)


# ---------------------------- columns ----------------------------


def test_group_columns_sets_outline_level():
    def build(ws):
        ws.group_columns(1, 3)

    sheet = _xml(build)
    assert 'outlineLevel="1"' in sheet


def test_group_columns_collapsed_sets_collapsed_flag():
    def build(ws):
        ws.group_columns_collapsed(1, 3)

    sheet = _xml(build)
    assert 'outlineLevel="1"' in sheet
    assert 'collapsed="1"' in sheet


def test_group_columns_rejects_reversed_range():
    def build(ws):
        ws.group_columns(3, 1)

    with pytest.raises(ValueError):
        _xml(build)


# ---------------------------- symbol position ----------------------------


def test_group_symbols_above_writes_summary_below_zero():
    def build(ws):
        ws.group_rows(1, 5)
        ws.group_symbols_above(True)

    sheet = _xml(build)
    assert 'summaryBelow="0"' in sheet


def test_group_symbols_to_left_writes_summary_right_zero():
    def build(ws):
        ws.group_columns(1, 3)
        ws.group_symbols_to_left(True)

    sheet = _xml(build)
    assert 'summaryRight="0"' in sheet


def test_group_symbols_default_omits_outline_pr():
    def build(ws):
        ws.group_rows(1, 5)
        # group_symbols_above/to_left not called -- default position,
        # <outlinePr> should be entirely absent (write_outline_pr()
        # returns early when both flags are false).

    sheet = _xml(build)
    assert "<outlinePr" not in sheet


# ---------------------------- constant_memory ----------------------------


def test_group_rows_after_rows_already_written_raises_in_constant_memory():
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet(constant_memory=True)
        for r in range(10):
            ws.write(r, 0, r)
        with pytest.raises(ValueError):
            ws.group_rows(1, 5)  # rows 1-5 already flushed
    finally:
        if os.path.exists(path):
            os.remove(path)


def test_group_rows_works_in_constant_memory_mode():
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet(constant_memory=True)
        # Rows must be grouped before being flushed by a write to a
        # later row -- group_rows() touches each row's own XML
        # attributes, same ordering requirement as set_row_format().
        # Nested groups, outer range first, must also work.
        ws.group_rows(0, 9)
        ws.group_rows(1, 5)
        for r in range(10):
            ws.write(r, 0, r)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert 'outlineLevel="1"' in sheet
        assert 'outlineLevel="2"' in sheet
    finally:
        if os.path.exists(path):
            os.remove(path)


def test_group_columns_works_in_constant_memory_mode_after_writes():
    # Columns are a separate <cols> section, unaffected by the
    # constant_memory row buffer -- must work called after writes too,
    # unlike group_rows().
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet(constant_memory=True)
        for r in range(10):
            ws.write(r, 0, r)
        ws.group_columns(1, 3)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert 'outlineLevel="1"' in sheet
    finally:
        if os.path.exists(path):
            os.remove(path)
