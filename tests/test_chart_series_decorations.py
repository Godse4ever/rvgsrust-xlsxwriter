"""ChartMarker, ChartTrendline and ChartDataLabel tests.

Asserted against chart1.xml. Several of the OOXML strings are not
guessable from the Rust variant names -- ShortDash writes "dot" and
LongDash writes "dash" -- so they are pinned explicitly.
"""
import os
import tempfile
import zipfile

import pytest

from rvgsrust_xlsxwriter import (
    Chart,
    ChartDataLabel,
    ChartFont,
    ChartFormat,
    ChartMarker,
    ChartSeries,
    ChartTrendline,
    Workbook,
)


def _chart_xml(configure, chart_type="line"):
    """Build a chart with one series, run configure(series), return XML."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet()
        for i, value in enumerate([10, 40, 20, 50, 30]):
            ws.write(i, 0, value)
        series = ChartSeries()
        series.set_values("Sheet1!$A$1:$A$5")
        configure(series)
        chart = Chart(chart_type)
        chart.push_series(series)
        ws.insert_chart(0, 3, chart)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            return z.read("xl/charts/chart1.xml").decode("utf-8")
    finally:
        if os.path.exists(path):
            os.remove(path)


# ------------------------------ marker ------------------------------


@pytest.mark.parametrize(
    "name,val",
    [
        ("square", "square"),
        ("diamond", "diamond"),
        ("triangle", "triangle"),
        ("x", "x"),
        ("star", "star"),
        ("circle", "circle"),
        ("plus_sign", "plus"),
        # These two do not match their Rust variant names.
        ("short_dash", "dot"),
        ("long_dash", "dash"),
    ],
)
def test_marker_types(name, val):
    def configure(series):
        marker = ChartMarker()
        marker.set_type(name)
        series.set_marker(marker)

    assert f'<c:symbol val="{val}"/>' in _chart_xml(configure)


def test_marker_none():
    """None is not a ChartMarkerType variant; it is a separate method."""

    def none(series):
        marker = ChartMarker()
        marker.set_none()
        series.set_marker(marker)

    assert '<c:symbol val="none"/>' in _chart_xml(none)


def test_marker_automatic_emits_no_marker_element():
    """Automatic is not a ChartMarkerType variant either, and it means
    "let Excel choose": upstream guards write_marker with
    `if !marker.automatic`, so no series-level c:marker is written at all.

    The <c:marker val="1"/> that appears at the lineChart level is a
    different element and is not matched by "<c:marker>".
    """

    def automatic(series):
        marker = ChartMarker()
        marker.set_automatic()
        series.set_marker(marker)

    xml = _chart_xml(automatic)
    assert "<c:marker>" not in xml
    assert "<c:chartSpace" in xml


def test_marker_size_and_format():
    def configure(series):
        fmt = ChartFormat()
        fmt.set_fill_color("#FF0000")
        marker = ChartMarker()
        marker.set_type("circle")
        marker.set_size(9)
        marker.set_format(fmt)
        series.set_marker(marker)

    xml = _chart_xml(configure)
    assert '<c:size val="9"/>' in xml
    assert "FF0000" in xml


def test_marker_invalid_type_raises():
    marker = ChartMarker()
    with pytest.raises(ValueError) as exc:
        marker.set_type("hexagon")
    assert "marker type" in str(exc.value)


def test_marker_invalid_type_message_mentions_the_methods():
    """automatic/none are methods, so the error should say so."""
    marker = ChartMarker()
    with pytest.raises(ValueError) as exc:
        marker.set_type("automatic")
    assert "set_automatic()" in str(exc.value)


# ----------------------------- trendline -----------------------------


@pytest.mark.parametrize(
    "name,val",
    [
        ("linear", "linear"),
        ("power", "power"),
        ("exponential", "exp"),
        ("logarithmic", "log"),
        ("polynomial", "poly"),
        ("moving_average", "movingAvg"),
    ],
)
def test_trendline_types(name, val):
    def configure(series):
        trend = ChartTrendline()
        trend.set_type(name)
        series.set_trendline(trend)

    xml = _chart_xml(configure)
    assert f'<c:trendlineType val="{val}"/>' in xml


def test_trendline_logarithmic_spelling():
    """Upstream's variant is Logarithmic, not Logarithm."""

    def configure(series):
        trend = ChartTrendline()
        trend.set_type("logarithmic")
        series.set_trendline(trend)

    assert '<c:trendlineType val="log"/>' in _chart_xml(configure)


@pytest.mark.parametrize("kind", ["polynomial", "moving_average"])
def test_trendline_period_is_accepted(kind):
    """Polynomial and MovingAverage carry a u8 period."""

    def configure(series):
        trend = ChartTrendline()
        trend.set_type(kind, 3)
        series.set_trendline(trend)

    assert "<c:trendline>" in _chart_xml(configure)


def test_trendline_display_equation_and_r_squared():
    """Exposed with a set_ prefix; upstream has none."""

    def configure(series):
        trend = ChartTrendline()
        trend.set_type("linear")
        trend.set_display_equation(True)
        trend.set_display_r_squared(True)
        series.set_trendline(trend)

    xml = _chart_xml(configure)
    assert '<c:dispEq val="1"/>' in xml
    assert '<c:dispRSqr val="1"/>' in xml


def test_trendline_periods_name_and_intercept():
    def configure(series):
        trend = ChartTrendline()
        trend.set_type("linear")
        trend.set_name("Trend")
        trend.set_forward_period(2.0)
        trend.set_backward_period(1.0)
        trend.set_intercept(0.0)
        trend.delete_from_legend(False)
        series.set_trendline(trend)

    xml = _chart_xml(configure)
    assert "Trend" in xml


def test_trendline_format_and_label_styling():
    def configure(series):
        fmt = ChartFormat()
        fmt.set_line_color("#FF0000")
        label_fmt = ChartFormat()
        label_fmt.set_fill_color("#FFFF00")
        font = ChartFont()
        font.set_bold()
        trend = ChartTrendline()
        trend.set_type("linear")
        trend.set_format(fmt)
        trend.set_label_format(label_fmt)
        trend.set_label_font(font)
        series.set_trendline(trend)

    xml = _chart_xml(configure)
    assert "FF0000" in xml


def test_trendline_invalid_type_raises():
    trend = ChartTrendline()
    with pytest.raises(ValueError) as exc:
        trend.set_type("logarithm")
    assert "trendline type" in str(exc.value)


# ---------------------------- data labels ----------------------------


def test_data_label_show_value():
    def configure(series):
        label = ChartDataLabel()
        label.show_value()
        series.set_data_label(label)

    xml = _chart_xml(configure)
    assert "<c:dLbls>" in xml
    assert '<c:showVal val="1"/>' in xml


def test_data_label_all_show_toggles():
    def configure(series):
        label = ChartDataLabel()
        label.show_value()
        label.show_category_name()
        label.show_series_name()
        label.show_leader_lines()
        label.show_legend_key()
        label.show_percentage()
        series.set_data_label(label)

    assert "<c:dLbls>" in _chart_xml(configure)


def test_data_label_x_and_y_value_for_scatter():
    def configure(series):
        # Upstream rejects a scatter chart with no categories range.
        series.set_categories("Sheet1!$A$1:$A$5")
        label = ChartDataLabel()
        label.show_x_value()
        label.show_y_value()
        series.set_data_label(label)

    assert "<c:dLbls>" in _chart_xml(configure, chart_type="scatter")


@pytest.mark.parametrize(
    "name,val",
    [
        ("center", "ctr"),
        ("left", "l"),
        ("above", "t"),
        ("below", "b"),
        ("inside_base", "inBase"),
        ("inside_end", "inEnd"),
        ("outside_end", "outEnd"),
        ("best_fit", "bestFit"),
    ],
)
def test_data_label_positions(name, val):
    def configure(series):
        label = ChartDataLabel()
        label.show_value()
        label.set_position(name)
        series.set_data_label(label)

    assert f'<c:dLblPos val="{val}"/>' in _chart_xml(configure)


def test_data_label_right_is_the_line_chart_default_and_omitted():
    """Line charts set default_label_position to Right, and upstream only
    writes dLblPos when the position differs from that default."""

    def configure(series):
        label = ChartDataLabel()
        label.show_value()
        label.set_position("right")
        series.set_data_label(label)

    xml = _chart_xml(configure)
    assert "<c:dLbls>" in xml
    assert "<c:dLblPos" not in xml


def test_data_label_default_position_is_accepted():
    """Default writes an empty string upstream, so only check it builds."""

    def configure(series):
        label = ChartDataLabel()
        label.show_value()
        label.set_position("default")
        series.set_data_label(label)

    assert "<c:dLbls>" in _chart_xml(configure)


def test_data_label_invalid_position_raises():
    label = ChartDataLabel()
    with pytest.raises(ValueError) as exc:
        label.set_position("diagonally")
    assert "data label position" in str(exc.value)


def test_data_label_num_format_and_separator():
    def configure(series):
        label = ChartDataLabel()
        label.show_value()
        label.set_num_format("#,##0.0")
        label.set_separator(";")
        series.set_data_label(label)

    xml = _chart_xml(configure)
    assert "#,##0.0" in xml
    assert "<c:separator>" in xml


@pytest.mark.parametrize("bad", ["", "ab", "  "])
def test_data_label_separator_must_be_one_character(bad):
    label = ChartDataLabel()
    if bad == "  ":
        with pytest.raises(ValueError):
            label.set_separator(bad)
    else:
        with pytest.raises(ValueError) as exc:
            label.set_separator(bad)
        assert "one character" in str(exc.value)


def test_data_label_hidden_and_font_and_format():
    def configure(series):
        fmt = ChartFormat()
        fmt.set_fill_color("#FFC000")
        font = ChartFont()
        font.set_size(11.0)
        label = ChartDataLabel()
        label.show_value()
        label.set_font(font)
        label.set_format(fmt)
        series.set_data_label(label)

    assert "FFC000" in _chart_xml(configure)


def test_data_label_set_hidden():
    def configure(series):
        label = ChartDataLabel()
        label.set_hidden()
        series.set_data_label(label)

    assert _chart_xml(configure)


def test_custom_data_labels():
    def configure(series):
        labels = []
        for text in ("one", "two", "three", "four", "five"):
            label = ChartDataLabel()
            label.show_value()
            label.set_value(text)
            label.set_custom()
            labels.append(label)
        series.set_custom_data_labels(labels)

    xml = _chart_xml(configure)
    assert "<c:dLbls>" in xml
    assert "three" in xml


def test_custom_data_labels_empty_list_is_accepted():
    def configure(series):
        series.set_custom_data_labels([])

    assert _chart_xml(configure)


# ---------------------------- type checks ----------------------------


def test_set_marker_rejects_non_marker():
    series = ChartSeries()
    with pytest.raises(TypeError):
        series.set_marker("not a marker")


def test_set_trendline_rejects_non_trendline():
    series = ChartSeries()
    with pytest.raises(TypeError):
        series.set_trendline(ChartMarker())


def test_set_data_label_rejects_non_label():
    series = ChartSeries()
    with pytest.raises(TypeError):
        series.set_data_label(ChartFont())


def test_marker_trendline_and_labels_together():
    def configure(series):
        marker = ChartMarker()
        marker.set_type("circle")
        marker.set_size(7)

        trend = ChartTrendline()
        trend.set_type("moving_average", 2)
        trend.set_display_r_squared(True)

        label = ChartDataLabel()
        label.show_value()
        label.set_position("above")

        series.set_marker(marker)
        series.set_trendline(trend)
        series.set_data_label(label)

    xml = _chart_xml(configure)
    assert '<c:symbol val="circle"/>' in xml
    assert '<c:trendlineType val="movingAvg"/>' in xml
    assert '<c:dLblPos val="t"/>' in xml


# ------------------- push_series default ordering -------------------
# Upstream's push_series applies chart-type defaults after copying the
# caller's series, clobbering a marker set beforehand. Chart.push_series
# goes through add_series to restore the right precedence. These pin both
# halves of that: the caller wins, but the default still applies when the
# caller made no choice.


def test_marker_survives_push_series_on_a_line_chart():
    """The regression this ordering fix exists for."""

    def configure(series):
        marker = ChartMarker()
        marker.set_type("diamond")
        marker.set_size(8)
        series.set_marker(marker)

    xml = _chart_xml(configure)
    assert '<c:symbol val="diamond"/>' in xml
    assert '<c:size val="8"/>' in xml


def test_line_chart_without_a_marker_still_defaults_to_none():
    """The default must still apply when the caller sets no marker."""

    def configure(series):
        pass

    assert '<c:symbol val="none"/>' in _chart_xml(configure)


@pytest.mark.parametrize(
    "chart_type", ["line", "line_stacked", "line_percent_stacked", "radar"]
)
def test_marker_survives_on_every_markers_off_chart_type(chart_type):
    def configure(series):
        marker = ChartMarker()
        marker.set_type("square")
        series.set_marker(marker)

    assert '<c:symbol val="square"/>' in _chart_xml(configure, chart_type)


def test_scatter_default_hidden_line_still_applied():
    """Plain scatter gets a hidden 2.25pt line unless a format is set."""

    def configure(series):
        series.set_categories("Sheet1!$A$1:$A$5")

    # 2.25pt in EMU.
    assert 'w="28575"' in _chart_xml(configure, chart_type="scatter")


def test_scatter_caller_format_wins_over_the_default_line():
    def configure(series):
        series.set_categories("Sheet1!$A$1:$A$5")
        fmt = ChartFormat()
        fmt.set_line_color("#FF0000")
        fmt.set_line_width(1.0)
        series.set_format(fmt)

    assert "FF0000" in _chart_xml(configure, chart_type="scatter")
