"""Autofilter criteria: FilterCondition, Worksheet.filter_column."""
import os
import tempfile
import zipfile

import pytest

from rvgsrust_xlsxwriter import FilterCondition, Workbook


def _sheet_xml(build):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet()
        ws.write_row(0, 0, ["Header", "Value"])
        ws.write_column(1, 0, ["North", "South", "East"])
        ws.autofilter(0, 0, 3, 1)
        build(ws)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            return z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    finally:
        if os.path.exists(path):
            os.remove(path)


# ------------------------------- list filter -------------------------------


def test_list_filter_string():
    fc = FilterCondition()
    fc.add_list_filter("North")
    sheet = _sheet_xml(lambda ws: ws.filter_column(0, fc))
    assert '<filterColumn colId="0">' in sheet
    assert '<filters><filter val="North"/></filters>' in sheet


def test_list_filter_multiple_values():
    fc = FilterCondition()
    fc.add_list_filter("North")
    fc.add_list_filter("South")
    sheet = _sheet_xml(lambda ws: ws.filter_column(0, fc))
    assert '<filter val="North"/>' in sheet
    assert '<filter val="South"/>' in sheet


def test_list_filter_number():
    fc = FilterCondition()
    fc.add_list_filter(42)
    sheet = _sheet_xml(lambda ws: ws.filter_column(1, fc))
    assert '<filter val="42"/>' in sheet


def test_list_blanks_filter():
    fc = FilterCondition()
    fc.add_list_blanks_filter()
    sheet = _sheet_xml(lambda ws: ws.filter_column(0, fc))
    assert '<filters blank="1"/>' in sheet


# ------------------------------ custom filter ------------------------------


def test_custom_filter_equal_to_has_no_operator_attribute():
    fc = FilterCondition()
    fc.add_custom_filter("equal_to", "North")
    sheet = _sheet_xml(lambda ws: ws.filter_column(0, fc))
    assert '<customFilter val="North"/>' in sheet


def test_custom_filter_greater_than():
    fc = FilterCondition()
    fc.add_custom_filter("greater_than", 100)
    sheet = _sheet_xml(lambda ws: ws.filter_column(1, fc))
    assert '<customFilter operator="greaterThan" val="100"/>' in sheet


def test_custom_filter_begins_with_adds_wildcard():
    fc = FilterCondition()
    fc.add_custom_filter("begins_with", "No")
    sheet = _sheet_xml(lambda ws: ws.filter_column(0, fc))
    assert '<customFilter val="No*"/>' in sheet


def test_custom_filter_contains_adds_wildcards_both_sides():
    fc = FilterCondition()
    fc.add_custom_filter("contains", "orth")
    sheet = _sheet_xml(lambda ws: ws.filter_column(0, fc))
    assert '<customFilter val="*orth*"/>' in sheet


def test_two_custom_filters_default_to_and():
    fc = FilterCondition()
    fc.add_custom_filter("greater_than", 10)
    fc.add_custom_filter("less_than", 100)
    sheet = _sheet_xml(lambda ws: ws.filter_column(1, fc))
    assert '<customFilters and="1">' in sheet


def test_two_custom_filters_with_boolean_or():
    fc = FilterCondition()
    fc.add_custom_filter("less_than", 10)
    fc.add_custom_filter("greater_than", 100)
    fc.add_custom_boolean_or()
    sheet = _sheet_xml(lambda ws: ws.filter_column(1, fc))
    assert "<customFilters>" in sheet
    assert '<customFilters and="1">' not in sheet


def test_invalid_filter_criteria_raises():
    fc = FilterCondition()
    with pytest.raises(ValueError):
        fc.add_custom_filter("sideways", "North")


def test_invalid_filter_value_type_raises():
    fc = FilterCondition()
    with pytest.raises(TypeError):
        fc.add_list_filter(["not", "a", "scalar"])
