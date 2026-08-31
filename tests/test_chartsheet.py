"""Workbook.add_chartsheet() -- a worksheet consisting of only a chart."""
import os
import tempfile
import zipfile

from rvgsrust_xlsxwriter import Chart, ChartSeries, Workbook


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


def test_add_chartsheet_file_and_default_name():
    def build(wb):
        ws = wb.add_worksheet()
        ws.write_column(0, 0, [10, 40, 50])
        series = ChartSeries()
        series.set_values("Sheet1!$A$1:$A$3")
        chart = Chart("column")
        chart.push_series(series)
        cs = wb.add_chartsheet()
        cs.insert_chart(0, 0, chart)

    files = _zip_contents(build)
    assert "xl/chartsheets/sheet1.xml" in files
    assert "xl/worksheets/sheet1.xml" in files  # the regular worksheet
    workbook_xml = _text(files, "xl/workbook.xml")
    assert 'name="Chart1"' in workbook_xml


def test_add_chartsheet_rename():
    def build(wb):
        ws = wb.add_worksheet()
        ws.write(0, 0, 1)
        series = ChartSeries()
        series.set_values("Sheet1!$A$1")
        chart = Chart("column")
        chart.push_series(series)
        cs = wb.add_chartsheet()
        cs.set_name("Dashboard")
        cs.insert_chart(0, 0, chart)

    files = _zip_contents(build)
    workbook_xml = _text(files, "xl/workbook.xml")
    assert 'name="Dashboard"' in workbook_xml


def test_multiple_chartsheets_number_independently_of_worksheets():
    def build(wb):
        ws1 = wb.add_worksheet()
        ws1.write(0, 0, 1)
        series = ChartSeries()
        series.set_values("Sheet1!$A$1")

        chart1 = Chart("column")
        chart1.push_series(series)
        cs1 = wb.add_chartsheet()
        cs1.insert_chart(0, 0, chart1)

        ws2 = wb.add_worksheet()
        ws2.write(0, 0, 2)

        chart2 = Chart("bar")
        chart2.push_series(series)
        cs2 = wb.add_chartsheet()
        cs2.insert_chart(0, 0, chart2)

    files = _zip_contents(build)
    # Two chartsheets, numbered 1 and 2 independently of the two regular
    # worksheets (also numbered 1 and 2 in their own sequence).
    assert "xl/chartsheets/sheet1.xml" in files
    assert "xl/chartsheets/sheet2.xml" in files
    assert "xl/worksheets/sheet1.xml" in files
    assert "xl/worksheets/sheet2.xml" in files
