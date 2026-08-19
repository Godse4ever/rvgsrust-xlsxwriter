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

### 1c. Row and column grouping — Medium-High

`group_rows(first, last)` (`worksheet.rs:6929`),
`group_columns(first, last)` (`7278`), the `*_collapsed` variants, and
`group_symbols_above(bool)` / `group_symbols_to_left(bool)`.

Outline grouping is how collapsible sections in a report are built.

### 1d. Data validation — High

`add_data_validation(r1, c1, r2, c2, validation)` at
`worksheet.rs:9388`, plus the whole `DataValidation` type
(`data_validation.rs:203`).

Suggested shape: a `DataValidation` pyclass following the same pattern as
the conditional formats — string-valued rule kinds validated with a
`ValueError` listing accepted values. Dropdown lists are one of the most
requested Excel-writer features across every language binding
— including openpyxl and pandas.

### 1e. Notes, filters, views, protection, images — Medium

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Cell notes | `worksheet.rs:5858` | `insert_note(row, col, text)`, `show_all_notes()`, `set_default_note_author(name)` — needs a `Note` pyclass | Medium |
| Autofilter criteria | `worksheet.rs:8791` | `filter_column(col, filter)` — needs a `FilterCondition` pyclass. We expose only the autofilter *range* today | Medium |
| Row / column visibility | `worksheet.rs:7717`, `8252` | `set_row_hidden(row)`, `set_column_hidden(col)`, `set_row_unhidden`, `hide_unused_rows` | Medium |
| Pixel dimensions | — | `set_row_height_pixels`, `set_column_width_pixels`, `set_default_row_height` | Low |
| Zoom, selection, tabs | `worksheet.rs:13140`, `10119` | `set_zoom(n)`, `set_selection(...)`, `set_active(bool)`, `set_first_tab(bool)`, `set_top_left_cell(...)`, `set_right_to_left(bool)`, `set_view_normal/page_layout/page_break_preview()` | Medium |
| Password protection | `worksheet.rs:9828` | `protect_with_password(pw)`, `protect_with_options(...)`, `unprotect_range(...)` | Medium |
| Image placement | `worksheet.rs:5197`, `5312` | `insert_image_with_offset(...)`, `insert_image_fit_to_cell(...)`, `embed_image(...)`, `insert_background_image(...)` | Medium |
| Clear a cell | `worksheet.rs:10776` | `clear_cell(row, col)`, `clear_cell_format(row, col)` | Low |
| Ignore-error flags | `worksheet.rs:14940` | `ignore_error(row, col, kind)` | Low |
| Checkboxes, buttons, shapes | `worksheet.rs:6393` | `insert_checkbox(...)`, `insert_button(...)`, `insert_shape(...)` | Low |
| Autofit tuning | `worksheet.rs:14663` | `autofit_to_max_width(w)`, `set_autofit_max_width(w)`, `set_autofit_max_row(r)` | Low |
| NaN / infinity strings | — | `set_nan_value(s)`, `set_infinity_value(s)`, `set_neg_infinity_value(s)` | Low |
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

## 3. Charts — ~60 gaps

Parts 1–3 covered the core (series, formatting, markers/trendlines/data
labels). What remains:

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Error bars | `chart.rs:7966` | `ChartErrorBars` pyclass (`chart.rs:17005`), attached via `ChartSeries.set_x_error_bars` / `set_y_error_bars` | **Medium-High** |
| Up-down bars | `chart.rs:2015` | `Chart.set_up_down_bars()`, `set_up_bar_format(fmt)`, `set_down_bar_format(fmt)` | Medium |
| Drop lines | `chart.rs:2326` | `Chart.set_drop_lines()`, `set_drop_lines_format(fmt)` | Medium |
| High-low lines | `chart.rs:2184` | `Chart.set_high_low_lines()`, `set_high_low_lines_format(fmt)` | Medium |
| Data table | `chart.rs:2470` | `ChartDataTable` pyclass (`chart.rs:17269`), via `Chart.set_data_table(table)` | Medium |
| Combined charts | `chart.rs:1744` | `Chart.combine(other_chart)` | Medium |
| Chart / plot area formatting | `chart.rs:1530`, `1592` | Flatten as `set_chart_area_format(fmt)` / `set_plot_area_format(fmt)`, since `ChartArea::new` is `pub(crate)` like the axes | Medium |
| Secondary axes | `chart.rs:1404` | `set_x2_axis_*` / `set_y2_axis_*`, mirroring the existing flattened axis methods. Pairs with the `ChartSeries.set_secondary_axis` we already expose | **Medium-High** |
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

## 4. Conditional formats — 2 gaps

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Icon sets | `conditional_format.rs:5719` | `ConditionalFormatIconSet` pyclass with `set_icon_type(str)` for the 12 styles, `reverse_icons(bool)`, `show_icons_only(bool)` | **Medium-High** |
| Custom icons | `conditional_format.rs:6351` | `ConditionalFormatCustomIcon` for per-threshold icon and value overrides, via `set_icons([...])` | Medium |

Icon sets are the one conspicuously missing conditional format — they are
a first-class Excel feature and visually obvious by their absence.

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

Ranked by value per unit of effort:

1. **Data validation** — needs a new pyclass, but high demand.
2. **Row and column grouping.**
3. **Conditional format icon sets.**
4. **Chart error bars and secondary axes.**
5. Everything else.

Headers/footers (text) and the Format/range-format gaps shipped ahead of this list.
