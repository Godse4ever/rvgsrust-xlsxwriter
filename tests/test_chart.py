"""Chart tests.

Assertions are made against the generated xl/charts/chart1.xml, which is
what Excel actually reads.
"""
import os
import tempfile
import zipfile
from datetime import date

import pytest

from rvgsrust_xlsxwriter import Chart, ChartErrorBars, ChartFormat, ChartSeries, Workbook

ALL_TYPES = [
    "area", "area_stacked", "area_percent_stacked",
    "bar", "bar_stacked", "bar_percent_stacked",
    "column", "column_stacked", "column_percent_stacked",
    "doughnut",
    "line", "line_stacked", "line_percent_stacked",
    "pie",
    "radar", "radar_with_markers", "radar_filled",
    "scatter", "scatter_straight", "scatter_straight_with_markers",
    "scatter_smooth", "scatter_smooth_with_markers",
    "stock",
]


def _build(build):
    """Write a 5x2 grid, run build(ws), return (chart_xml, all_zip_names)."""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tf:
        path = tf.name
    try:
        wb = Workbook()
        ws = wb.add_worksheet()
        for i, (label, value) in enumerate(
            [("a", 10), ("b", 40), ("c", 20), ("d", 50), ("e", 30)]
        ):
            ws.write(i, 0, label)
            ws.write(i, 1, value)
        build(ws)
        wb.close(path)
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            xml = ""
            if "xl/charts/chart1.xml" in names:
                xml = z.read("xl/charts/chart1.xml").decode("utf-8")
            return xml, names
    finally:
        if os.path.exists(path):
            os.remove(path)


def _simple(chart_type="column", **calls):
    """One chart with a single series over B1:B5, configured by calls."""

    def build(ws):
        series = ChartSeries()
        series.set_values("Sheet1!$B$1:$B$5")
        series.set_categories("Sheet1!$A$1:$A$5")
        chart = Chart(chart_type)
        chart.push_series(series)
        for name, arg in calls.items():
            method = getattr(chart, name)
            if arg is None:
                method()
            elif isinstance(arg, tuple):
                method(*arg)
            else:
                method(arg)
        ws.insert_chart(0, 3, chart)

    return _build(build)[0]


# ------------------------------ types ------------------------------


@pytest.mark.parametrize("chart_type", ALL_TYPES)
def test_every_chart_type_produces_a_chart(chart_type):
    """Catches a bad ChartType mapping for any of the 23 variants."""
    xml = _simple(chart_type)
    assert xml, f"no chart1.xml produced for {chart_type}"
    assert "<c:chartSpace" in xml


def test_chart_type_is_case_insensitive():
    assert _simple("COLUMN")


def test_invalid_chart_type_raises():
    with pytest.raises(ValueError) as exc:
        Chart("piechart")
    assert "chart type" in str(exc.value)
    # The message should list the accepted values.
    assert "scatter_smooth_with_markers" in str(exc.value)


@pytest.mark.parametrize(
    "chart_type,element",
    [
        ("bar", "<c:barChart>"),
        ("column", "<c:barChart>"),
        ("line", "<c:lineChart>"),
        ("pie", "<c:pieChart>"),
        ("doughnut", "<c:doughnutChart>"),
        ("scatter", "<c:scatterChart>"),
        ("radar", "<c:radarChart>"),
        ("area", "<c:areaChart>"),
    ],
)
def test_chart_type_maps_to_expected_element(chart_type, element):
    assert element in _simple(chart_type)


# ------------------------------ series ------------------------------


def test_series_values_and_categories_reach_the_xml():
    xml = _simple()
    assert "Sheet1!$B$1:$B$5" in xml
    assert "Sheet1!$A$1:$A$5" in xml


def test_series_name():
    def build(ws):
        series = ChartSeries()
        series.set_values("Sheet1!$B$1:$B$5")
        series.set_name("Revenue")
        chart = Chart("column")
        chart.push_series(series)
        ws.insert_chart(0, 3, chart)

    assert "Revenue" in _build(build)[0]


def test_multiple_series():
    def build(ws):
        chart = Chart("column")
        for col in ("$A$1:$A$5", "$B$1:$B$5"):
            series = ChartSeries()
            series.set_values(f"Sheet1!{col}")
            chart.push_series(series)
        ws.insert_chart(0, 3, chart)

    xml = _build(build)[0]
    assert xml.count("<c:ser>") == 2


def test_series_options_do_not_error():
    def build(ws):
        series = ChartSeries()
        series.set_values("Sheet1!$B$1:$B$5")
        series.set_secondary_axis(False)
        series.set_overlap(-25)
        series.set_gap(150)
        series.set_smooth(True)
        series.set_invert_if_negative()
        series.set_invert_if_negative_color("#FF0000")
        series.delete_from_legend(False)
        chart = Chart("column")
        chart.push_series(series)
        ws.insert_chart(0, 3, chart)

    assert "<c:ser>" in _build(build)[0]


def test_series_point_colors():
    def build(ws):
        series = ChartSeries()
        series.set_values("Sheet1!$B$1:$B$5")
        series.set_point_colors(["#FF0000", "#00B050", "#0070C0"])
        chart = Chart("pie")
        chart.push_series(series)
        ws.insert_chart(0, 3, chart)

    xml = _build(build)[0]
    assert "FF0000" in xml
    assert "00B050" in xml


def test_series_point_colors_rejects_bad_color():
    series = ChartSeries()
    with pytest.raises(ValueError):
        series.set_point_colors(["#FF0000", "chartreuse-ish"])


def test_series_reused_across_two_charts():
    """ChartSeries derives Clone upstream, so reuse must be safe."""

    def build(ws):
        series = ChartSeries()
        series.set_values("Sheet1!$B$1:$B$5")
        first = Chart("column")
        first.push_series(series)
        second = Chart("line")
        second.push_series(series)
        ws.insert_chart(0, 3, first)
        ws.insert_chart(20, 3, second)

    _, names = _build(build)
    assert "xl/charts/chart1.xml" in names
    assert "xl/charts/chart2.xml" in names


# ------------------------------ title ------------------------------


def test_title_name():
    assert "Quarterly" in _simple(set_title_name="Quarterly")


def test_title_hidden_takes_no_argument():
    """Upstream's ChartTitle::set_hidden() has no bool parameter."""
    xml = _simple(set_title_hidden=None)
    assert "<c:chartSpace" in xml


def test_title_overlay():
    assert _simple(set_title_overlay=True)


# ------------------------------- axes -------------------------------


def test_axis_names():
    xml = _simple(set_x_axis_name="Month", set_y_axis_name="Sales")
    assert "Month" in xml
    assert "Sales" in xml


def test_y_axis_min_and_max():
    xml = _simple(set_y_axis_min=0.0, set_y_axis_max=100.0)
    assert '<c:max val="100"/>' in xml
    assert '<c:min val="0"/>' in xml


def test_y_axis_units():
    xml = _simple(set_y_axis_major_unit=20.0, set_y_axis_minor_unit=5.0)
    assert '<c:majorUnit val="20"/>' in xml
    assert '<c:minorUnit val="5"/>' in xml


def test_y_axis_log_base():
    assert '<c:logBase val="10"/>' in _simple(set_y_axis_log_base=10)


def test_y_axis_num_format():
    assert "#,##0.00" in _simple(set_y_axis_num_format="#,##0.00")


def test_axis_hidden_emits_delete():
    assert '<c:delete val="1"/>' in _simple(set_y_axis_hidden=True)


def test_axis_reverse_takes_no_argument():
    """Upstream's ChartAxis::set_reverse() has no bool parameter."""
    assert _simple(set_y_axis_reverse=None)


def test_axis_gridlines():
    xml = _simple(set_x_axis_major_gridlines=True, set_y_axis_minor_gridlines=True)
    assert "<c:majorGridlines/>" in xml or "<c:majorGridlines>" in xml


def test_x_axis_date_and_text_axis():
    assert _simple(set_x_axis_date_axis=True)
    assert _simple(set_x_axis_text_axis=True)


# --------------------------- secondary axes ---------------------------


def _with_secondary(chart_calls=None):
    """Two series (one routed to the secondary axis) + optional chart-level
    x2/y2 axis calls. Upstream only emits secondary-axis XML when a series
    has secondary_axis=True (Chart::check_for_secondary_axis()), so a
    secondary-axis test must route a series there, not just call the
    chart-level setters."""

    def build(ws):
        primary = ChartSeries()
        primary.set_values("Sheet1!$A$1:$A$5")

        secondary = ChartSeries()
        secondary.set_values("Sheet1!$B$1:$B$5")
        secondary.set_secondary_axis(True)

        chart = Chart("column")
        chart.push_series(primary)
        chart.push_series(secondary)
        for name, arg in (chart_calls or {}).items():
            method = getattr(chart, name)
            if arg is None:
                method()
            else:
                method(arg)
        ws.insert_chart(0, 3, chart)

    return _build(build)[0]


def test_secondary_axis_not_emitted_without_a_secondary_series():
    """A naive call to set_x2_axis_* without routing any series to the
    secondary axis must not silently produce secondary-axis XML."""
    xml = _simple(set_x2_axis_name="Secondary")
    assert "valAx" in xml
    assert xml.count("<c:valAx>") == 1


def test_secondary_axis_emitted_with_a_secondary_series():
    xml = _with_secondary()
    assert xml.count("<c:valAx>") == 2


def test_x2_axis_name():
    assert "Units" in _with_secondary({"set_x2_axis_name": "Units"})


def test_y2_axis_name():
    assert "Revenue" in _with_secondary({"set_y2_axis_name": "Revenue"})


def test_y2_axis_min_and_max():
    xml = _with_secondary({"set_y2_axis_min": 0.0})
    xml = _with_secondary(
        {"set_y2_axis_min": 0.0, "set_y2_axis_max": 100.0}
    )
    assert xml.count("<c:valAx>") == 2


def test_y2_axis_units():
    xml = _with_secondary(
        {"set_y2_axis_major_unit": 20.0, "set_y2_axis_minor_unit": 5.0}
    )
    assert "20" in xml and "5" in xml


def test_y2_axis_log_base():
    assert '<c:logBase val="10"/>' in _with_secondary({"set_y2_axis_log_base": 10})


def test_y2_axis_num_format():
    assert "#,##0.00" in _with_secondary({"set_y2_axis_num_format": "#,##0.00"})


def test_x2_axis_hidden_emits_delete():
    xml = _with_secondary({"set_x2_axis_hidden": True})
    assert '<c:delete val="1"/>' in xml


def test_x2_axis_reverse_takes_no_argument():
    assert _with_secondary({"set_x2_axis_reverse": None})


def test_x2_axis_gridlines():
    xml = _with_secondary(
        {"set_x2_axis_major_gridlines": True, "set_y2_axis_minor_gridlines": True}
    )
    assert xml.count("<c:valAx>") == 2


def test_x2_axis_date_and_text_axis():
    assert _with_secondary({"set_x2_axis_date_axis": True})
    assert _with_secondary({"set_x2_axis_text_axis": True})


# ------------------------------ legend ------------------------------


@pytest.mark.parametrize(
    "position,val",
    [("right", "r"), ("left", "l"), ("top", "t"), ("bottom", "b"), ("top_right", "tr")],
)
def test_legend_positions(position, val):
    xml = _simple(set_legend_position=position)
    assert f'<c:legendPos val="{val}"/>' in xml


def test_legend_hidden_removes_legend():
    """set_hidden() takes no argument, and drops the element entirely."""
    xml = _simple(set_legend_hidden=None)
    assert "<c:legend>" not in xml


def test_legend_overlay():
    assert _simple(set_legend_overlay=True)


def test_invalid_legend_position_raises():
    chart = Chart("column")
    with pytest.raises(ValueError) as exc:
        chart.set_legend_position("overlay_right")
    assert "legend position" in str(exc.value)


# --------------------------- chart options ---------------------------


def test_style_and_size():
    assert _simple(set_style=12, set_width=600, set_height=400)


def test_name_and_alt_text():
    assert _simple(set_name="Chart 1", set_alt_text="A column chart")


def test_doughnut_hole_size_and_rotation():
    xml = _simple("doughnut", set_hole_size=30, set_rotation=90)
    assert '<c:holeSize val="30"/>' in xml


def test_show_empty_cells_as():
    for option in ("gaps", "zero", "connected"):
        assert _simple(show_empty_cells_as=option)


def test_show_empty_cells_as_invalid_raises():
    chart = Chart("column")
    with pytest.raises(ValueError):
        chart.show_empty_cells_as("blank")


def test_show_hidden_data_and_na():
    assert _simple(show_hidden_data=None, show_na_as_empty_cell=None)


# ---------------------------- insert_chart ----------------------------


def test_insert_chart_with_offsets():
    def build(ws):
        series = ChartSeries()
        series.set_values("Sheet1!$B$1:$B$5")
        chart = Chart("column")
        chart.push_series(series)
        ws.insert_chart(0, 3, chart, 15, 25)

    _, names = _build(build)
    assert "xl/charts/chart1.xml" in names


def test_insert_chart_rejects_non_chart():
    wb = Workbook()
    ws = wb.add_worksheet()
    with pytest.raises(TypeError):
        ws.insert_chart(0, 3, "not a chart")


def test_chart_without_series_raises():
    """Upstream validation rejects a chart with no series."""

    def build(ws):
        ws.insert_chart(0, 3, Chart("column"))

    with pytest.raises(ValueError):
        _build(build)


def test_two_charts_on_one_worksheet():
    def build(ws):
        for i, kind in enumerate(("column", "line")):
            series = ChartSeries()
            series.set_values("Sheet1!$B$1:$B$5")
            chart = Chart(kind)
            chart.push_series(series)
            ws.insert_chart(i * 20, 3, chart)

    _, names = _build(build)
    assert "xl/charts/chart1.xml" in names
    assert "xl/charts/chart2.xml" in names


def test_full_option_sweep():
    def build(ws):
        series = ChartSeries()
        series.set_values("Sheet1!$B$1:$B$5")
        series.set_categories("Sheet1!$A$1:$A$5")
        series.set_name("Sales")
        chart = Chart("column")
        chart.push_series(series)
        chart.set_style(11)
        chart.set_width(640)
        chart.set_height(480)
        chart.set_title_name("Sales by month")
        chart.set_x_axis_name("Month")
        chart.set_y_axis_name("Revenue")
        chart.set_y_axis_min(0.0)
        chart.set_y_axis_max(60.0)
        chart.set_y_axis_major_unit(10.0)
        chart.set_y_axis_num_format("#,##0")
        chart.set_x_axis_major_gridlines(False)
        chart.set_y_axis_major_gridlines(True)
        chart.set_legend_position("bottom")
        ws.insert_chart(0, 3, chart)

    xml = _build(build)[0]
    assert "Sales by month" in xml
    assert '<c:legendPos val="b"/>' in xml



# ----------------------------- error bars -----------------------------
#
# Real upstream quirk, confirmed against source (chart.rs write_series()
# vs write_scatter_series()): only Scatter charts write real x AND y
# error bars with a <c:errDir> telling Excel which axis they're on. Every
# other chart type goes through the generic write_series() path, which
# special-cases Bar (only x_error_bars is honored, written with an empty/
# absent errDir since a Bar chart only has one value axis) and Column
# (mirror image: only y_error_bars, same empty errDir), and for every
# other type (Line, Pie, Area, Radar, Doughnut, Stock, ...) only
# y_error_bars is honored, written with errDir="y". In every one of
# these non-Scatter cases the "wrong" axis's error bars are silently
# dropped -- not a binding bug, this is upstream's own behavior.


def _with_error_bars(y_error_bars=None, x_error_bars=None, chart_type="scatter"):
    def build(ws):
        series = ChartSeries()
        series.set_values("Sheet1!$B$1:$B$5")
        if chart_type == "scatter":
            series.set_categories("Sheet1!$A$1:$A$5")
        if y_error_bars is not None:
            series.set_y_error_bars(y_error_bars)
        if x_error_bars is not None:
            series.set_x_error_bars(x_error_bars)
        chart = Chart(chart_type)
        chart.push_series(series)
        ws.insert_chart(0, 3, chart)

    return _build(build)[0]


def test_no_error_bars_by_default():
    assert "errBars" not in _simple()


def test_scatter_y_error_bars_standard_error_default():
    eb = ChartErrorBars()
    xml = _with_error_bars(y_error_bars=eb)
    assert "<c:errBars>" in xml
    assert '<c:errDir val="y"/>' in xml
    assert '<c:errBarType val="both"/>' in xml
    assert '<c:errValType val="stdErr"/>' in xml
    # StandardError writes no <c:val>.
    assert "<c:val " not in xml


def test_scatter_x_and_y_error_bars_together():
    y_eb = ChartErrorBars()
    x_eb = ChartErrorBars()
    x_eb.set_type_percentage(5.0)
    xml = _with_error_bars(y_error_bars=y_eb, x_error_bars=x_eb, chart_type="scatter")
    assert xml.count("<c:errBars>") == 2
    assert '<c:errDir val="x"/>' in xml
    assert '<c:errDir val="y"/>' in xml


def test_scatter_error_bars_fixed_value():
    eb = ChartErrorBars()
    eb.set_type_fixed_value(2.5)
    xml = _with_error_bars(y_error_bars=eb)
    assert '<c:errValType val="fixedVal"/>' in xml
    assert '<c:val val="2.5"/>' in xml


def test_scatter_error_bars_percentage():
    eb = ChartErrorBars()
    eb.set_type_percentage(10.0)
    xml = _with_error_bars(y_error_bars=eb)
    assert '<c:errValType val="percentage"/>' in xml
    assert '<c:val val="10"/>' in xml


def test_scatter_error_bars_standard_deviation():
    eb = ChartErrorBars()
    eb.set_type_standard_deviation(1.0)
    xml = _with_error_bars(y_error_bars=eb)
    assert '<c:errValType val="stdDev"/>' in xml


def test_scatter_error_bars_custom_range():
    eb = ChartErrorBars()
    eb.set_type_custom("Sheet1!$C$1:$C$5", "Sheet1!$D$1:$D$5")
    xml = _with_error_bars(y_error_bars=eb)
    assert '<c:errValType val="cust"/>' in xml
    assert "<c:plus>" in xml
    assert "<c:minus>" in xml


def test_scatter_error_bars_direction():
    eb = ChartErrorBars()
    eb.set_direction("minus")
    xml = _with_error_bars(y_error_bars=eb)
    assert '<c:errBarType val="minus"/>' in xml


def test_error_bars_invalid_direction_raises():
    eb = ChartErrorBars()
    with pytest.raises(ValueError):
        eb.set_direction("sideways")


def test_scatter_error_bars_no_end_cap():
    eb = ChartErrorBars()
    eb.set_end_cap(False)
    xml = _with_error_bars(y_error_bars=eb)
    assert '<c:noEndCap val="1"/>' in xml


def test_scatter_error_bars_end_cap_on_by_default():
    eb = ChartErrorBars()
    xml = _with_error_bars(y_error_bars=eb)
    assert "noEndCap" not in xml


def test_bar_chart_only_honors_x_error_bars():
    x_eb = ChartErrorBars()
    y_eb = ChartErrorBars()
    xml = _with_error_bars(x_error_bars=x_eb, y_error_bars=y_eb, chart_type="bar")
    assert xml.count("<c:errBars>") == 1
    assert "<c:errDir" not in xml


def test_column_chart_only_honors_y_error_bars():
    x_eb = ChartErrorBars()
    y_eb = ChartErrorBars()
    xml = _with_error_bars(x_error_bars=x_eb, y_error_bars=y_eb, chart_type="column")
    assert xml.count("<c:errBars>") == 1
    assert "<c:errDir" not in xml


def test_line_chart_only_honors_y_error_bars():
    x_eb = ChartErrorBars()
    y_eb = ChartErrorBars()
    xml = _with_error_bars(x_error_bars=x_eb, y_error_bars=y_eb, chart_type="line")
    assert xml.count("<c:errBars>") == 1
    assert '<c:errDir val="y"/>' in xml


# --------------------- up-down bars / drop / high-low lines ---------------------


def test_up_down_bars_default_empty_tags():
    xml = _simple(chart_type="line", set_up_down_bars=True)
    assert "<c:upDownBars>" in xml
    assert "<c:upBars/>" in xml
    assert "<c:downBars/>" in xml


def test_up_down_bars_only_on_line_charts():
    xml = _simple(chart_type="column", set_up_down_bars=True)
    assert "upDownBars" not in xml


def test_high_low_lines_default_empty_tag():
    xml = _simple(chart_type="line", set_high_low_lines=True)
    assert "<c:hiLowLines/>" in xml


def test_drop_lines_default_empty_tag():
    xml = _simple(chart_type="line", set_drop_lines=True)
    assert "<c:dropLines/>" in xml


def test_drop_lines_not_written_when_disabled():
    xml = _simple(chart_type="line")
    assert "dropLines" not in xml
    assert "hiLowLines" not in xml
    assert "upDownBars" not in xml


def test_up_bar_format_gets_sppr_element():
    fmt = ChartFormat()
    fmt.set_fill_color("#FF0000")
    xml = _simple(chart_type="line", set_up_down_bars=True, set_up_bar_format=fmt)
    assert "<c:upBars><c:spPr>" in xml


# ----------------------- chart area / plot area formatting -----------------------


def test_chart_area_format():
    fmt = ChartFormat()
    fmt.set_fill_color("#FF0000")
    xml = _simple(set_chart_area_format=fmt)
    assert "FF0000" in xml


def test_plot_area_format():
    fmt = ChartFormat()
    fmt.set_fill_color("#00FF00")
    xml = _simple(set_plot_area_format=fmt)
    assert "00FF00" in xml


def test_chart_area_and_plot_area_format_together():
    chart_fmt = ChartFormat()
    chart_fmt.set_fill_color("#FF0000")
    plot_fmt = ChartFormat()
    plot_fmt.set_fill_color("#00FF00")
    xml = _simple(set_chart_area_format=chart_fmt, set_plot_area_format=plot_fmt)
    assert "FF0000" in xml
    assert "00FF00" in xml
    # chart area's spPr is a direct child of chartSpace (after </c:chart>),
    # plot area's is the last child of plotArea -- confirm both fills
    # landed in their own distinct spPr, not merged into one.
    assert xml.count("<c:spPr>") >= 2


# -------------------------------- combined charts --------------------------------


def test_combine_charts_both_types_present():
    def build(ws):
        primary = ChartSeries()
        primary.set_values("Sheet1!$A$1:$A$5")
        primary_chart = Chart("column")
        primary_chart.push_series(primary)

        secondary = ChartSeries()
        secondary.set_values("Sheet1!$B$1:$B$5")
        secondary.set_secondary_axis(True)
        secondary_chart = Chart("line")
        secondary_chart.push_series(secondary)

        primary_chart.combine(secondary_chart)
        ws.insert_chart(0, 3, primary_chart)

    xml = _build(build)[0]
    assert "<c:barChart>" in xml
    assert "<c:lineChart>" in xml


def test_combine_mutating_other_after_call_has_no_effect():
    def build(ws):
        primary = ChartSeries()
        primary.set_values("Sheet1!$A$1:$A$5")
        primary_chart = Chart("column")
        primary_chart.push_series(primary)

        secondary = ChartSeries()
        secondary.set_values("Sheet1!$B$1:$B$5")
        secondary.set_secondary_axis(True)
        secondary_chart = Chart("line")
        secondary_chart.push_series(secondary)

        primary_chart.combine(secondary_chart)
        # Mutating secondary_chart after combine() should not affect the
        # already-combined snapshot (upstream clones on combine()).
        secondary_chart.set_title_name("Should not appear")
        ws.insert_chart(0, 3, primary_chart)

    xml = _build(build)[0]
    assert "Should not appear" not in xml


# ---------------- axis label/tick/date/crossing/display-unit settings ----------------


def test_axis_label_position():
    xml = _simple(set_x_axis_label_position="low")
    assert '<c:tickLblPos val="low"/>' in xml


def test_axis_label_position_none():
    xml = _simple(set_y_axis_label_position="none")
    assert '<c:tickLblPos val="none"/>' in xml


def test_axis_label_alignment():
    xml = _simple(chart_type="bar", set_y_axis_label_alignment="right")
    assert '<c:lblAlgn val="r"/>' in xml


def test_axis_label_interval():
    xml = _simple(set_x_axis_label_interval=3)
    assert '<c:tickLblSkip val="3"/>' in xml


def test_axis_major_tick_type():
    xml = _simple(set_x_axis_major_tick_type="outside")
    assert '<c:majorTickMark val="out"/>' in xml


def test_axis_minor_tick_type():
    xml = _simple(set_y_axis_minor_tick_type="cross")
    assert '<c:minorTickMark val="cross"/>' in xml


def test_axis_tick_interval():
    # tick_interval/label_interval are category-axis-only properties
    # upstream -- only meaningful on x_axis for a standard column/line
    # chart (where x is the category axis).
    xml = _simple(set_x_axis_tick_interval=2)
    assert '<c:tickMarkSkip val="2"/>' in xml


def test_axis_min_max_date_and_major_unit_date_type():
    def build(ws):
        series = ChartSeries()
        series.set_values("Sheet1!$B$1:$B$5")
        chart = Chart("line")
        chart.push_series(series)
        chart.set_x_axis_date_axis(True)
        chart.set_x_axis_min_date(date(2000, 1, 1))
        chart.set_x_axis_max_date(date(2000, 1, 31))
        chart.set_x_axis_major_unit_date_type("months")
        ws.insert_chart(0, 3, chart)

    xml = _build(build)[0]
    assert '<c:min val="36526"/>' in xml
    assert '<c:max val="36556"/>' in xml
    assert '<c:majorTimeUnit val="months"/>' in xml


def test_axis_crossing_keyword():
    xml = _simple(chart_type="bar", set_y_axis_crossing="max")
    assert 'val="max"' in xml


def test_axis_crossing_numeric_value():
    xml = _simple(chart_type="bar", set_y_axis_crossing=25.5)
    assert '<c:crossesAt val="25.5"/>' in xml


def test_axis_crossing_invalid_type_raises():
    with pytest.raises(TypeError):
        _simple(set_x_axis_crossing=["not", "valid"])


def test_axis_position_between_ticks():
    # crossBetween is rendered on the value axis, but the field that
    # controls it lives on the category axis struct upstream (same
    # "crossing comes from the opposite axis" OOXML quirk as crossing
    # itself) -- so this is x_axis even though the XML element ends up
    # inside <c:valAx>.
    xml = _simple(set_x_axis_position_between_ticks=False)
    assert '<c:crossBetween val="midCat"/>' in xml


def test_axis_display_unit_type_and_visible():
    xml = _simple(
        set_y_axis_display_unit_type="thousands",
        set_y_axis_display_units_visible=True,
    )
    assert "<c:dispUnits>" in xml
    assert '<c:builtInUnit val="thousands"/>' in xml
    assert "<c:dispUnitsLbl" in xml


def test_axis_display_units_not_written_when_none():
    xml = _simple(chart_type="column")
    assert "dispUnits" not in xml
