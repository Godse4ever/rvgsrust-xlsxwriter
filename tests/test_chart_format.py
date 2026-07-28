"""ChartFormat and ChartFont tests, asserted against chart1.xml."""
import os
import tempfile
import zipfile

import pytest

from rvgsrust_xlsxwriter import (
    Chart,
    ChartFont,
    ChartFormat,
    ChartSeries,
    Workbook,
)


def _chart_xml(build):
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet()
        for i, value in enumerate([10, 40, 20, 50, 30]):
            ws.write(i, 0, value)
        build(ws)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            return z.read("xl/charts/chart1.xml").decode("utf-8")
    finally:
        if os.path.exists(path):
            os.remove(path)


def _with_series_format(fmt):
    def build(ws):
        series = ChartSeries()
        series.set_values("Sheet1!$A$1:$A$5")
        series.set_format(fmt)
        chart = Chart("column")
        chart.push_series(series)
        ws.insert_chart(0, 3, chart)

    return _chart_xml(build)


def _with_chart_call(method, arg):
    def build(ws):
        series = ChartSeries()
        series.set_values("Sheet1!$A$1:$A$5")
        chart = Chart("column")
        chart.push_series(series)
        chart.set_title_name("Title")
        chart.set_x_axis_name("X")
        chart.set_y_axis_name("Y")
        getattr(chart, method)(arg)
        ws.insert_chart(0, 3, chart)

    return _chart_xml(build)


# ------------------------------ fills ------------------------------


def test_fill_color():
    fmt = ChartFormat()
    fmt.set_fill_color("#FF0000")
    assert "FF0000" in _with_series_format(fmt)


def test_fill_transparency():
    fmt = ChartFormat()
    fmt.set_fill_color("#FF0000")
    fmt.set_fill_transparency(40)
    xml = _with_series_format(fmt)
    assert "FF0000" in xml
    assert "<a:alpha" in xml


def test_no_fill():
    fmt = ChartFormat()
    fmt.set_no_fill()
    assert "<a:noFill/>" in _with_series_format(fmt)


def test_fill_color_invalid_raises():
    fmt = ChartFormat()
    with pytest.raises(ValueError):
        fmt.set_fill_color("almost red")


# ------------------------------ lines ------------------------------


def test_line_color_and_width():
    fmt = ChartFormat()
    fmt.set_line_color("#0070C0")
    fmt.set_line_width(2.5)
    xml = _with_series_format(fmt)
    assert "0070C0" in xml
    # Widths are written in EMU: 2.5pt * 12700.
    assert 'w="31750"' in xml


@pytest.mark.parametrize(
    "name,val",
    [
        ("round_dot", "sysDot"),
        ("square_dot", "sysDash"),
        ("dash", "dash"),
        ("dash_dot", "dashDot"),
        ("long_dash", "lgDash"),
        ("long_dash_dot", "lgDashDot"),
        ("long_dash_dot_dot", "lgDashDotDot"),
    ],
)
def test_line_dash_types(name, val):
    fmt = ChartFormat()
    fmt.set_line_color("#000000")
    fmt.set_line_dash_type(name)
    xml = _with_series_format(fmt)
    assert f'<a:prstDash val="{val}"/>' in xml


def test_solid_dash_type_emits_no_prst_dash():
    """Solid is the default, so upstream omits the element entirely."""
    fmt = ChartFormat()
    fmt.set_line_color("#000000")
    fmt.set_line_dash_type("solid")
    assert "<a:prstDash" not in _with_series_format(fmt)


def test_line_dash_type_invalid_raises():
    fmt = ChartFormat()
    with pytest.raises(ValueError) as exc:
        fmt.set_line_dash_type("squiggly")
    assert "line dash type" in str(exc.value)


def test_line_transparency_and_hidden():
    fmt = ChartFormat()
    fmt.set_line_color("#000000")
    fmt.set_line_transparency(25)
    assert _with_series_format(fmt)

    hidden = ChartFormat()
    hidden.set_line_hidden(True)
    assert _with_series_format(hidden)


def test_no_line():
    fmt = ChartFormat()
    fmt.set_no_line()
    assert _with_series_format(fmt)


def test_line_setters_compose():
    """Line state is kept in the pyclass, so successive calls accumulate."""
    fmt = ChartFormat()
    fmt.set_line_color("#0070C0")
    fmt.set_line_width(3.0)
    fmt.set_line_dash_type("dash")
    xml = _with_series_format(fmt)
    assert "0070C0" in xml
    assert 'w="38100"' in xml
    assert '<a:prstDash val="dash"/>' in xml


# ------------------------------ borders ------------------------------


def test_border_color_and_width():
    fmt = ChartFormat()
    fmt.set_border_color("#00B050")
    fmt.set_border_width(1.5)
    xml = _with_series_format(fmt)
    assert "00B050" in xml


def test_border_dash_and_transparency():
    fmt = ChartFormat()
    fmt.set_border_color("#000000")
    fmt.set_border_dash_type("long_dash")
    fmt.set_border_transparency(10)
    assert '<a:prstDash val="lgDash"/>' in _with_series_format(fmt)


def test_no_border():
    fmt = ChartFormat()
    fmt.set_no_border()
    assert _with_series_format(fmt)


def test_border_hidden():
    fmt = ChartFormat()
    fmt.set_border_hidden(True)
    assert _with_series_format(fmt)


# ------------------------------ fonts ------------------------------


def test_font_bold_and_italic():
    font = ChartFont()
    font.set_bold()
    font.set_italic()
    xml = _with_chart_call("set_title_font", font)
    assert 'b="1"' in xml
    assert 'i="1"' in xml


def test_font_unset_bold_and_default_bold():
    font = ChartFont()
    font.set_bold()
    font.unset_bold()
    font.set_default_bold(False)
    assert _with_chart_call("set_title_font", font)


def test_font_underline_and_strikethrough():
    font = ChartFont()
    font.set_underline()
    font.set_strikethrough()
    assert _with_chart_call("set_title_font", font)


def test_font_size_is_written_in_hundredths():
    font = ChartFont()
    font.set_size(14.0)
    assert 'sz="1400"' in _with_chart_call("set_title_font", font)


def test_font_color_and_name():
    font = ChartFont()
    font.set_color("#7030A0")
    font.set_name("Courier New")
    xml = _with_chart_call("set_title_font", font)
    assert "7030A0" in xml
    assert "Courier New" in xml


def test_font_rotation_and_rtl():
    font = ChartFont()
    font.set_rotation(-45)
    font.set_right_to_left(True)
    assert _with_chart_call("set_title_font", font)


def test_font_pitch_family_and_character_set():
    font = ChartFont()
    font.set_pitch_family(2)
    font.set_character_set(0)
    assert _with_chart_call("set_title_font", font)


def test_font_color_invalid_raises():
    font = ChartFont()
    with pytest.raises(ValueError):
        font.set_color("chartreuse-ish")


# --------------------------- wiring points ---------------------------


@pytest.mark.parametrize(
    "method",
    [
        "set_title_font",
        "set_x_axis_font",
        "set_x_axis_name_font",
        "set_y_axis_font",
        "set_y_axis_name_font",
        "set_legend_font",
    ],
)
def test_every_font_attachment_point(method):
    font = ChartFont()
    font.set_bold()
    font.set_color("#FF0000")
    xml = _with_chart_call(method, font)
    assert "FF0000" in xml


@pytest.mark.parametrize(
    "method",
    [
        "set_title_format",
        "set_x_axis_format",
        "set_x_axis_name_format",
        "set_y_axis_format",
        "set_y_axis_name_format",
        "set_legend_format",
    ],
)
def test_every_format_attachment_point(method):
    fmt = ChartFormat()
    fmt.set_fill_color("#FFC000")
    xml = _with_chart_call(method, fmt)
    assert "FFC000" in xml


def test_same_format_reused_on_two_elements():
    """set_format clones internally, so one object can be attached twice."""

    def build(ws):
        fmt = ChartFormat()
        fmt.set_fill_color("#FFC000")
        series = ChartSeries()
        series.set_values("Sheet1!$A$1:$A$5")
        series.set_format(fmt)
        chart = Chart("column")
        chart.push_series(series)
        chart.set_title_name("Title")
        chart.set_title_format(fmt)
        chart.set_legend_format(fmt)
        ws.insert_chart(0, 3, chart)

    assert _chart_xml(build).count("FFC000") >= 3


def test_same_font_reused_on_two_elements():
    def build(ws):
        font = ChartFont()
        font.set_bold()
        series = ChartSeries()
        series.set_values("Sheet1!$A$1:$A$5")
        chart = Chart("column")
        chart.push_series(series)
        chart.set_title_name("Title")
        chart.set_x_axis_name("X")
        chart.set_title_font(font)
        chart.set_x_axis_font(font)
        ws.insert_chart(0, 3, chart)

    assert 'b="1"' in _chart_xml(build)


def test_format_rejects_non_format_object():
    series = ChartSeries()
    with pytest.raises(TypeError):
        series.set_format("not a format")


def test_font_rejects_non_font_object():
    chart = Chart("column")
    with pytest.raises(TypeError):
        chart.set_title_font("not a font")


def test_series_format_plus_chart_fonts_together():
    def build(ws):
        fmt = ChartFormat()
        fmt.set_fill_color("#4472C4")
        fmt.set_border_color("#000000")
        fmt.set_border_width(1.0)

        title_font = ChartFont()
        title_font.set_bold()
        title_font.set_size(16.0)
        title_font.set_color("#333333")

        series = ChartSeries()
        series.set_values("Sheet1!$A$1:$A$5")
        series.set_format(fmt)

        chart = Chart("column")
        chart.push_series(series)
        chart.set_title_name("Styled")
        chart.set_title_font(title_font)
        chart.set_y_axis_name("Value")
        chart.set_legend_position("bottom")
        ws.insert_chart(0, 3, chart)

    xml = _chart_xml(build)
    assert "4472C4" in xml
    assert 'sz="1600"' in xml
    assert "Styled" in xml
