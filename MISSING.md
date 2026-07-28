# API parity gaps vs rust_xlsxwriter 0.96

**Listing only — nothing here is implemented, and nothing here should be
implemented in the same session as the audit.**

All `file:line` references are to
[`v0.96.0`](https://github.com/jmcnamara/rust_xlsxwriter/tree/v0.96.0/src)
and were read from the source, not from docs.rs.

## Method

Public inherent methods on each upstream type were extracted and diffed
against the methods exposed by our `#[pymethods]` blocks. Three classes of
apparent gap were excluded as false positives:

- **`*_with_format` variants.** Upstream splits `write_string` /
  `write_string_with_format`; we take an optional `format=None` instead.
  About 20 upstream methods collapse this way.
- **Accessors we deliberately flattened.** `Chart::x_axis()`, `title()`,
  `legend()`, `Workbook::worksheet_from_index()` and similar return `&mut`
  references that cannot cross into Python. Their *contents* are audited
  below; the accessors themselves are not gaps.
- **Constructors and internals.** `Chart::new_pie()` and friends are
  covered by our `Chart("pie")` string form. `validate()`,
  `populate_string_table()`, `format_dxf_index()`, `set_axis_ids()` are
  internal.

Priorities are a judgement call about request frequency, weighted toward
this project's actual use — survey and market-research reporting, where
page setup, print layout and per-column formatting come up far more often
than VBA or chartsheets. They are not upstream's opinion.

## Already at parity

Worth recording so this is not re-audited:

- **Sparklines: 28 of 28 methods exposed. No gaps.** Note the brief listed
  a `custom_ranges` parameter as a parity item — **no such method or
  parameter exists in 0.96**, on either `Sparkline` or the two
  `Worksheet::add_sparkline*` methods.
- **Workbook document properties and workbook-level defined names** are
  already exposed (`set_properties`, `define_name`), also listed in the
  brief as gaps.
- Conditional formats: all 12 rule types except icon sets (below).

---

## 1. Format — ~30 gaps

The single biggest cluster, and the one users hit first.

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Individual border sides | `format.rs:2030` | `set_border_bottom(style)`, `..._top`, `..._left`, `..._right`, plus `set_border_*_color(color)` | **High** |
| Diagonal borders | `format.rs:2175` | `set_border_diagonal(style)`, `set_border_diagonal_color`, `set_border_diagonal_type` | Medium |
| Cell protection | `format.rs:2409`, `2284`, `2299` | `set_locked()`, `set_unlocked()`, `set_hidden()` | **High** |
| Strikethrough | `format.rs:1329` | `set_font_strikethrough()` | Medium |
| Pattern foreground colour | `format.rs:1855` | `set_foreground_color(color)` — pairs with the existing background colour | Medium |
| Quote prefix | `format.rs:2343` | `set_quote_prefix()` | Low |
| Hyperlink style | `format.rs:2217` | `set_hyperlink()` | Low |
| Checkbox format | `format.rs:2355` | `set_checkbox()` | Low |
| Font family / charset / scheme / script | `format.rs:1226` | `set_font_family(n)` etc. | Low |
| Reading direction | `format.rs:1647` | `set_reading_direction(n)` | Low |
| `unset_*` inverses | `format.rs:2364` | `unset_bold()`, `unset_italic()`, `unset_text_wrap()`, `unset_shrink()`, `unset_hidden()`, `unset_quote_prefix()` | Low |

**Note on borders:** we expose only `set_border(style)`, which sets all
four sides. Per-side borders are extremely common in report layout — this
is the highest-value item in the table.

---

## 2. Worksheet — ~120 gaps

### 2a. Cell, row, column and range formats — **CLOSED**

`set_column_format`, `set_column_range_format`, `set_row_format`,
`set_cell_format` and `set_range_format` are now exposed. This was ranked
first because it was the gap hit directly while implementing extended
Arrow types: date columns needed a number format applied and a caller had
no way to fix a wrongly formatted column afterwards, which is why
`write_dataframe` applies the date formats itself.

Still missing in this area:

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Range format with a border | `worksheet.rs:10549` | `set_range_format_with_border(r1, c1, r2, c2, format, border)` | Medium |
| Clear a cell's format | `worksheet.rs:10776` | `clear_cell_format(row, col)` | Low |

### 2b. Page setup and printing — **High as a group**

Anchor: `worksheet.rs:12853` onward.

`set_landscape()` / `set_portrait()` (`12940`), `set_paper_size(n)`
(`12853`), `set_margins(...)` (`13811`), `set_print_area(...)`
(`14278`), `set_print_fit_to_pages(w, h)` (`14017`),
`set_repeat_rows(first, last)` / `set_repeat_columns(...)` (`14367`),
`set_page_breaks(rows)` / `set_vertical_page_breaks(cols)` (`13040`),
plus `set_print_scale`, `set_print_gridlines`, `set_print_headings`,
`set_print_center_horizontally` / `_vertically`, `set_print_black_and_white`,
`set_print_draft`, `set_print_first_page_number`, `set_page_order`.

Suggested shape: mirror upstream names 1:1 on `Worksheet`. Nearly all take
a bool or a small integer, so none need a new class.

Anyone generating a printable report needs most of this. It is also the
cheapest large win in the whole audit: no new pyclasses, no enum wrappers
beyond paper size.

### 2c. Headers and footers — **High**

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Header / footer text | `worksheet.rs:13505`, `13539` | `set_header(text)`, `set_footer(text)` | **High** |
| Header / footer images | `worksheet.rs:13662`, `13700` | `set_header_image(position, image)` — needs an `Image` pyclass first | Medium |
| Align / scale with page | — | `set_header_footer_align_with_page(bool)`, `set_header_footer_scale_with_doc(bool)` | Low |

### 2d. Row and column grouping — **Medium-High**

`group_rows(first, last)` (`worksheet.rs:6929`),
`group_columns(first, last)` (`7278`), the `*_collapsed` variants, and
`group_symbols_above(bool)` / `group_symbols_to_left(bool)`.

Outline grouping is how collapsible sections in a report are built.

### 2e. Data validation — **High**

`add_data_validation(r1, c1, r2, c2, validation)` at
`worksheet.rs:9388`, plus the whole `DataValidation` type
(`data_validation.rs:203`).

Suggested shape: a `DataValidation` pyclass following the same pattern as
the conditional formats — string-valued rule kinds validated with a
`ValueError` listing accepted values. Dropdown lists are one of the most
requested spreadsheet features there is.

### 2f. Formulas — **Medium**

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Array formulas over a range | `worksheet.rs:3448` | `write_array_formula(r1, c1, r2, c2, formula, format=None)` | Medium |
| Dynamic array formulas | `worksheet.rs:3641` | `write_dynamic_array_formula(...)`, `write_dynamic_formula(row, col, ...)` | Medium |
| Cached formula results | `worksheet.rs:10285` | `set_formula_result(row, col, result)`, `set_formula_result_default(...)` | Medium |

Cached results matter for anything read by a tool that doesn't recalculate
— including openpyxl and pandas.

### 2g. Notes, filters, views, protection, images — **Medium**

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

## 3. Workbook — ~15 gaps

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Save to a buffer | `workbook.rs:1371` | `save_to_buffer() -> bytes` | **Medium-High** |
| Chartsheets | `workbook.rs:906` | `add_chartsheet()` — needs a `Chartsheet` pyclass | Medium |
| Themes | `workbook.rs:1918` | `use_excel_2023_theme()`, `use_custom_theme(...)` | Low |
| Workbook default format | — | `set_default_format(format)` | Low |
| VBA projects | `workbook.rs:2141` | `add_vba_project(path)`, `set_vba_name(name)` | Low |
| Read-only recommended | `workbook.rs:2345` | `read_only_recommended()` | Low |
| Temp dir / large zip | `workbook.rs:828` | `set_tempdir(path)`, `use_zip_large_file(bool)` | Low |

`save_to_buffer()` is the notable one: returning bytes instead of writing
a path is what a web service needs, and it avoids a temp-file round trip.

---

## 4. Charts — ~60 gaps

Parts 1–3 covered the core. What remains:

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

## 5. Conditional formats — 2 gaps

| Feature | Upstream | Suggested Python API | Priority |
|---|---|---|---|
| Icon sets | `conditional_format.rs:5719` | `ConditionalFormatIconSet` pyclass with `set_icon_type(str)` for the 12 styles, `reverse_icons(bool)`, `show_icons_only(bool)` | **Medium-High** |
| Custom icons | `conditional_format.rs:6351` | `ConditionalFormatCustomIcon` for per-threshold icon and value overrides, via `set_icons([...])` | Medium |

Icon sets are the one conspicuously missing conditional format — they are
a first-class Excel feature and visually obvious by their absence. The
"custom min/max/mid values" the brief lists as a gap are already covered
by the existing `set_minimum` / `set_midpoint` / `set_maximum`; the custom
values that remain missing are the icon-set thresholds specifically.

---

## Suggested order

Ranked by value per unit of effort, not by section order above:

1. ~~**`set_column_format` / `set_row_format` / `set_range_format`**~~ —
   done, see section 2a.
2. **Page setup and print settings** — large surface, trivial individually,
   no new pyclasses.
3. **Per-side borders and cell protection on `Format`** — same shape as
   existing `Format` setters.
4. **Headers and footers** (text only; images need an `Image` pyclass).
5. **`save_to_buffer()`** — one method, unlocks web-service use.
6. **Data validation** — needs a new pyclass, but high demand.
7. **Row and column grouping.**
8. **Conditional format icon sets.**
9. **Chart error bars and secondary axes.**
10. Everything else.

Items 1–5 are all mechanical and would close the majority of what a
reporting-focused user would notice missing.
