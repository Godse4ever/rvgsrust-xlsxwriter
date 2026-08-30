# API parity gaps vs rust_xlsxwriter 0.98.2

Only open items below. Anything closed has been removed from this file
rather than marked done, so this stays a to-do list, not a changelog —
see [CHANGELOG.md](CHANGELOG.md) for what shipped.

All `file:line` references below were read from source. Originally
captured against
[`v0.96.0`](https://github.com/jmcnamara/rust_xlsxwriter/tree/v0.96.0/src);
re-checked line-by-line against
[`v0.98.2`](https://github.com/jmcnamara/rust_xlsxwriter/tree/v0.98.2/src)
after this project's 0.96→0.98.2 upgrade (all 50 references across
`chart.rs`, `conditional_format.rs`, `data_validation.rs`, `format.rs`,
`workbook.rs`, `worksheet.rs`) — every one still points to the exact
same function at the exact same line, unchanged across both version
bumps. The 0.97.0/0.98.0 "Added" changelog entries were dependency
version bumps only (`zip`, `polars`), not new public API, so nothing
in this list closed on its own between 0.96 and 0.98.2 either.

Priorities are a judgement call about request frequency, weighted toward
this project's actual use — survey and market-research reporting, where
page setup, print layout and per-column formatting come up far more often
than VBA or chartsheets. They are not upstream's opinion.

---

## 1. Worksheet

### 1a. Cell/row/column/range formats — closed

`set_range_format_with_border()` and `clear_cell_format()` are
implemented.

### 1b. Headers and footers — text done, images/scale remain

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Header / footer images | `worksheet.rs:13662`, `13700` | `set_header_image(position, image)` — needs an `Image` pyclass first | Medium |
| Align / scale with page | — | `set_header_footer_align_with_page(bool)`, `set_header_footer_scale_with_doc(bool)` | Low |

`set_header(text)` / `set_footer(text)` are implemented.

### 1c. Row and column grouping — closed

`group_rows`/`group_rows_collapsed`/`group_columns`/
`group_columns_collapsed`/`group_symbols_above`/`group_symbols_to_left`
are implemented.

**Known limitation, verified against actual output:** in
`constant_memory=True` mode, individual `<row>` elements never get an
`outlineLevel` attribute -- only the sheet-wide `outlineLevelRow`
maximum is written. `group_rows()` still succeeds and doesn't raise,
but the per-row visual grouping in Excel won't appear when
`constant_memory=True` is used. Appears to be an upstream
`constant_memory` streaming limitation, not fixable from this binding
without writing worksheet XML directly. `group_columns()` is
unaffected (columns are a separate `<cols>` section, outside the
row-streaming mechanism).

### 1d. Data validation — mostly closed, two pieces remain

`DataValidation` pyclass and `Worksheet.add_data_validation()` are
implemented: whole-number/decimal-number/text-length range rules (all
8 comparison types), string dropdown lists (`allow_list_strings`,
`allow_list_formula`), custom formula rules, `allow_any_value`, and
every input/error-message and behaviour setting.

Not implemented:

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Date/time range rules | `data_validation.rs` `allow_date`/`allow_time` | `allow_date(rule_type, value, value2=None)`, `allow_time(...)` | Medium |
| Cell-reference formula variants | `allow_*_formula()` on whole_number/decimal_number/text_length/date/time | `allow_whole_number_formula(rule_type, formula, formula2=None)` etc. | Low |

`allow_date`/`allow_time` need an `ExcelDateTime` construction path
from a Python `date`/`time`/`datetime` object, which none of the
numeric/string/formula rules already implemented require — that's
why they were left for a follow-up rather than blocking the rest of
this feature. The formula variants swap a literal value for a cell
reference in the same 8 comparison rules; lower priority since a
`allow_custom()` formula already covers the same ground less
conveniently.

### 1e. Notes, filters, views, protection, images — closed (except serde)

Row/column visibility, pixel dimensions, zoom/selection/tabs, ignored
errors, autofit tuning, and NaN/infinity display strings shipped in
PR #30 (`autofit_to_max_width` was dropped from the original list --
upstream deprecates it in favor of `set_autofit_max_width()` +
`autofit()`, both already exposed). Password protection
(`ProtectionOptions` pyclass, `protect_with_options()`,
`unprotect_range()`) shipped in PR #31 -- `protect()`/
`protect_with_password()` already existed. Image placement
(`insert_image_with_offset`, `embed_image(_with_format)`,
`insert_image_fit_to_cell(_centered)`, `insert_background_image`)
shipped in PR #32 -- all take a plain `image_path: str` like the
existing `insert_image()`, no `Image` pyclass needed. Checkboxes
(flat bool + optional `Format`), a new `Button` pyclass, and a new
`Shape` pyclass (Textbox only -- the only shape type upstream
implements -- text/sizing only, not fill/line/font) shipped in
PR #33. Cell notes (`Note` pyclass -- text/author/sizing/visible/
alt_text/background_color/font_name/font_size, not `set_format()` or
`set_object_movement()`) shipped in PR #34. Autofilter criteria
(`FilterCondition` pyclass -- `add_list_filter`, `add_list_blanks_filter`,
`add_custom_filter` with all 12 `FilterCriteria` variants,
`add_custom_boolean_or`) shipped in PR #35, via a new
`Worksheet.filter_column()` (previously only the autofilter *range*
was exposed). What remains:

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| serde serialisation | — | `serialize_headers(...)` and friends. Feature-gated upstream; would need the `serde` feature enabled | Low |

---

## 2. Workbook — 6 gaps

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Chartsheets | `workbook.rs:906` | `add_chartsheet()` — needs a `Chartsheet` pyclass | Medium |
| Themes | `workbook.rs:1918` | `use_excel_2023_theme()`, `use_custom_theme(...)` | Low |
| Workbook default format | — | `set_default_format(format)` | Low |
| VBA projects | `workbook.rs:2141` | `add_vba_project(path)`, `set_vba_name(name)` | Low |
| Read-only recommended | `workbook.rs:2345` | `read_only_recommended()` | Low |
| Temp dir / large zip | `workbook.rs:828` | `set_tempdir(path)`, `use_zip_large_file(bool)` | Low |

---

## 3. Charts — ~58 gaps

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

## 4. Conditional formats — closed

`ConditionalFormatIconSet` and `ConditionalFormatCustomIcon` are
implemented: all 20 icon set styles, `reverse_icons`, `show_icons_only`,
per-icon threshold/type/direction overrides via `set_icons()`.

Real upstream gotcha, worked around rather than just documented: a
freshly-constructed `ConditionalFormatIconSet` fails upstream's own
validation unless `set_icon_type()` has been called at least once
(`rust_xlsxwriter`'s `new()` leaves its internal icon-rules list empty,
and `add_conditional_format()` requires it to have exactly 3/4/5
entries matching the icon type). This binding's constructor calls
`set_icon_type()` with upstream's own default internally, so a
freshly-constructed instance is valid on its own without the caller
needing to know about this.

---

## Suspected upstream bug (not a parity gap)

`Worksheet::set_print_first_page_number` (`worksheet.rs:18697`) writes the
page number into the `useFirstPageNumber` attribute and never emits a
`firstPageNumber` attribute:

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

---

## Suggested order

Charts deprioritized to last per explicit direction -- everything
else (Worksheet 1e, Workbook) goes first regardless of the
value-per-effort ranking that guided the list before this:

1. Worksheet 1e -- notes, filters, views, protection, images.
   **Closed** (PRs #30-#35), except serde serialisation (Low
   priority, feature-gated upstream, left for later).
2. Workbook -- 6 gaps. Chartsheets need a `Chartsheet` pyclass; the
   rest (themes, default format, VBA, read-only-recommended,
   tempdir/large-zip) are flat setters. **Next up.**
3. Charts -- ~58 remaining gaps (data tables, combined charts,
   gradient/pattern fills, axis label placement/tick marks/date-min-
   max, manual layout, per-point formatting, up-down/drop/high-low
   bars, legend entry deletion, object movement/decorative/scaling).
   Secondary axes (PR #28) and error bars (PR #29) already shipped.

Headers/footers (text), the Format/range-format gaps, the core of
data validation, row/column grouping, and conditional format icon
sets shipped ahead of this list. Date/time validation rules, the
data-validation formula variants, and the constant_memory grouping
limitation remain -- see 1c and 1d above.
