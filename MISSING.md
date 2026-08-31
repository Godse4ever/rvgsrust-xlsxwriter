# API parity gaps vs rust_xlsxwriter 0.98.2

Only open items below. Anything closed has been removed from this file
rather than marked done, so this stays a to-do list, not a changelog —
see [CHANGELOG.md](CHANGELOG.md) for what shipped.

Worksheet and Workbook are both fully closed now (PRs #30-#38) --
Charts is the only section with open gaps. serde serialisation is the
one deliberate exception, left open below rather than closed, because
it's a real feature (new Cargo feature flag + a Python-dict-to-
serde_json::Value bridge) rather than a small gap, and hasn't been
attempted yet.

All `file:line` references below were read from source, re-checked
against
[`v0.98.2`](https://github.com/jmcnamara/rust_xlsxwriter/tree/v0.98.2/src).

Priorities are a judgement call about request frequency, weighted toward
this project's actual use — survey and market-research reporting, where
page setup, print layout and per-column formatting come up far more often
than VBA or chartsheets. They are not upstream's opinion.

---

## Charts — ~58 gaps

Parts 1–3 covered the core (series, formatting, markers/trendlines/data
labels). Secondary axes (`set_x2_axis_*`/`set_y2_axis_*`, PR #28) and
error bars (`ChartErrorBars`, PR #29) have shipped. What remains:

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Up-down bars | `chart.rs:2015` | `Chart.set_up_down_bars()`, `set_up_bar_format(fmt)`, `set_down_bar_format(fmt)` | Medium |
| Drop lines | `chart.rs:2326` | `Chart.set_drop_lines()`, `set_drop_lines_format(fmt)` | Medium |
| High-low lines | `chart.rs:2184` | `Chart.set_high_low_lines()`, `set_high_low_lines_format(fmt)` | Medium |
| Data table | `chart.rs:2470` | `ChartDataTable` pyclass (`chart.rs:17269`), via `Chart.set_data_table(table)` | Medium |
| Combined charts | `chart.rs:1744` | `Chart.combine(other_chart)` | Medium |
| Chart / plot area formatting | `chart.rs:1530`, `1592` | Flatten as `set_chart_area_format(fmt)` / `set_plot_area_format(fmt)`, since `ChartArea::new` is `pub(crate)` like the axes | Medium |
| Per-point formatting | `chart.rs:7750` | `ChartPoint` pyclass, via `ChartSeries.set_points([...])` | Medium |
| Gradient fills | `chart.rs:13980` | `ChartFormat.set_gradient_fill(...)` plus `ChartGradientStop` | Medium |
| Pattern fills | `chart.rs:13915` | `ChartFormat.set_pattern_fill(...)` | Low |
| Axis label placement | `chart.rs:12102` | `set_x_axis_label_position(str)`, `..._label_interval(n)`, `..._label_alignment(str)` | Medium |
| Axis tick marks | `chart.rs:12376` | `set_x_axis_major_tick_type(str)`, `..._minor_tick_type(str)`, `..._tick_interval(n)` | Medium |
| Date axis min/max/units | `chart.rs:11559` | `set_x_axis_min_date(...)`, `set_x_axis_max_date(...)`, `..._major_unit_date_type(str)` | Medium |
| Axis crossing | `chart.rs:11329` | `set_x_axis_crossing(...)`, `set_x_axis_position_between_ticks(bool)` | Low |
| Display units | `chart.rs:11747` | `set_y_axis_display_unit_type(str)`, `..._display_units_visible(bool)` | Low |
| Gridline formatting | — | `set_x_axis_major_gridlines_format(fmt)` — currently gridlines can only be toggled, not styled | Medium |
| Manual layouts | `chart.rs:9147` | `ChartLayout` pyclass for `title`, `legend`, `plot_area` | Low |
| Legend entry deletion | `chart.rs:13314` | `Chart.set_legend_delete_entries([indices])` | Low |
| Object movement, decorative, scaling | `chart.rs:2668`, `2645`, `2583` | `set_object_movement(str)` (note `MoveButDontSizeWithCells`, not `MoveWithCells`), `set_decorative(bool)`, `set_scale_width(n)`, `set_scale_height(n)` | Low |

---

## serde serialisation — deferred, not attempted

`Worksheet.serialize()`/`serialize_headers()` and friends are feature-gated
upstream behind Cargo's `serde` feature. `serde`/`serde_json` are already
present transitively in `Cargo.lock` (pulled in by other dependencies), so
enabling the feature and adding a direct `serde_json` dependency is lower-risk
than it looks at first. The real work is the bridge: upstream's methods are
generic over `T: Serialize`, which only works for actual Rust structs deriving
`Serialize` -- there's no such thing from Python. A binding would need to
accept a Python dict (or list of dicts) and convert it to `serde_json::Value`
(which does implement `Serialize`) before calling through, and would need to
be checked against upstream's serializer to see whether it tolerates a plain
JSON value the way it tolerates a derived struct (field-renaming/skip
attributes on a real struct have no JSON equivalent). Not a small gap --
left as a dedicated follow-up rather than rushed.

---

## Known limitations (not parity gaps)

**`group_rows()` in `constant_memory=True` mode:** individual `<row>`
elements never get an `outlineLevel` attribute -- only the sheet-wide
`outlineLevelRow` maximum is written. `group_rows()` still succeeds and
doesn't raise, but the per-row visual grouping in Excel won't appear.
Appears to be an upstream `constant_memory` streaming limitation, not
fixable from this binding without writing worksheet XML directly.
`group_columns()` is unaffected (columns are a separate `<cols>`
section, outside the row-streaming mechanism).

**Suspected upstream bug:** `Worksheet::set_print_first_page_number`
(`worksheet.rs:18697`) writes the page number into the
`useFirstPageNumber` attribute and never emits a `firstPageNumber`
attribute:

```rust
if self.first_page_number > 0 {
    attributes.push(("useFirstPageNumber", self.first_page_number.to_string()));
}
```

Per ECMA-376, `pageSetup@useFirstPageNumber` is a boolean and
`pageSetup@firstPageNumber` is the uint carrying the value. Excel treats a
nonzero boolean as true, so the feature is enabled but the first page
number probably defaults to 1 rather than the requested value.

Our binding passes the call straight through, so it inherits the
behaviour. `tests/test_page_setup.py::test_print_first_page_number` pins
what is currently written rather than what ought to be, and will fail if
upstream changes it. Not worked around locally: doing so would mean
writing worksheet XML ourselves. Worth reporting upstream.
