# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0-dev] - Unreleased

**Core build confirmed on real hardware** (macOS, 4-core, rustc 1.83+,
`maturin develop --release`): `rust_xlsxwriter 0.96` compiles cleanly
with no source changes and no sandbox-specific dependency pins, and a
basic Workbook/add_worksheet/write/close smoke test passes. This was
run against an earlier point in this entry's history, before the
`constant_memory` row-order enforcement below and the version bump
itself -- worth re-confirming after pulling latest, but the
fundamental "does this even build" risk this project's development
sandbox couldn't resolve is now answered: yes.

Not yet confirmed on real hardware: the full test suite (`pytest
tests/`), the `constant_memory` row-order enforcement specifically, and
current performance numbers (an early comparison against
`rustpy-xlsxwriter` on that same machine showed `rustpy` still ahead,
1.26-1.53x depending on row count and narrowing at scale -- see
PERFORMANCE.md for what's confirmed vs. still estimated).

### Added
- Charts, part 2: `ChartFormat` and `ChartFont`, attachable to a series,
  a chart title, either axis (both the axis labels and the axis name), and
  the legend.
  `ChartLine` and `ChartSolidFill` are not exposed as separate classes.
  Upstream they exist only to be passed to `ChartFormat`, so they are
  flattened into it as `set_line_*`, `set_border_*` and `set_fill_*`, with
  the line and fill state kept per format object so successive calls
  compose. Pattern and gradient fills are logged for the parity audit.
  `set_format` is generic over `IntoChartFormat`, which upstream implements
  for `&mut ChartFormat`, so each call site passes an owned mutable clone.
  The trait itself never needs importing despite not being re-exported
  from the crate root, since it only ever appears as a bound on a generic
  parameter.
- Charts, part 1: `Chart` and `ChartSeries` classes plus
  `Worksheet.insert_chart(row, col, chart, x_offset=0, y_offset=0)`.
  Covers all 23 chart types, series ranges and options, and the title,
  x/y axis and legend settings.
  Axes, titles and legends are flattened onto `Chart` as `set_x_axis_*`,
  `set_title_*` and `set_legend_*` rather than being separate classes,
  because `ChartAxis::new`, `ChartTitle::new` and `ChartLegend::new` are
  `pub(crate)` upstream and cannot be constructed from a binding.
  Series are attached with `Chart.push_series(series)` rather than being
  passed to `insert_chart`. `Chart` does not derive `Clone` upstream, so
  pushing at insert time would mutate the only copy and silently duplicate
  series if the same chart were inserted twice, with no `remove_series` to
  undo it.
  `ChartFormat`, `ChartFont`, `ChartMarker`, `ChartTrendline` and
  `ChartDataLabel` follow in parts 2 and 3.
- Sparklines: a single `Sparkline` class plus
  `Worksheet.add_sparkline(row, col, sparkline)` and
  `Worksheet.add_sparkline_group(first_row, first_col, last_row, last_col,
  sparkline)`. Covers the three types, all seven point-marker toggles, all
  seven colors, line weight, custom and group min/max, style presets, date
  ranges, right-to-left, column order, and empty-cell handling.
  As with the conditional formats, enum-valued options are strings
  validated with a `ValueError` listing the accepted values. The type
  accepts `win_lose` (upstream spells the variant `WinLose`) and also
  `win_loss`, since that is the spelling most people reach for first; both
  serialize to Excel's `type="stacked"`.
  Grouped sparklines require a 2D data range, one row per sparkline;
  passing a 1D range raises `ValueError`, as does adding a sparkline with
  no range set.
- Conditional formatting: 12 rule types, each a `#[pyclass]` wrapping the
  matching `rust_xlsxwriter` builder -- `ConditionalFormatCell`, `Blank`,
  `Duplicate`, `Error`, `Formula`, `Average`, `Top`, `Text`, `Date`,
  `2ColorScale`, `3ColorScale` and `DataBar`. Applied with
  `Worksheet.add_conditional_format(first_row, first_col, last_row,
  last_col, rule)`. Icon sets are deferred to the parity audit.
  Enum-valued options (average/date/text/top rules, scale value types,
  data bar direction and axis position) are taken as strings and validated
  with a `ValueError` that lists the accepted values, rather than being
  exposed as separate enum classes.
  Note these setters return `None` rather than `self`, so unlike `Format`
  they do not chain; adding a return value later is backwards compatible.
- Extended Arrow type support in `Worksheet.write_dataframe()`. Added
  `int8`/`int16`/`int32`, `uint8`/`uint16`/`uint32`/`uint64`, `float32`,
  `date32`/`date64`, and `timestamp` in all four units (second,
  millisecond, microsecond, nanosecond). `timestamp[ns]` matters most in
  practice: it is what pandas' default `datetime64[ns]` dtype maps to, so
  the most common real-world DataFrame previously raised `TypeError` here
  and silently fell back to the per-cell path in `dataframe.py`, where
  datetimes were written as strings. They are now real Excel dates.
- Date and timestamp columns get a number format applied automatically
  (`yyyy-mm-dd` and `yyyy-mm-dd hh:mm:ss` respectively). Without one Excel
  renders a date serial as a bare number such as `45123`, and this binding
  exposes no `set_column_format()` for the caller to fix it afterwards.
  The two formats are built once per `write_dataframe()` call and resolved
  per column, not per cell, so the inner loop cost is unchanged for
  non-temporal data.
- Timezone-aware `timestamp` columns now emit a `UserWarning` naming the
  column and its timezone, once per column at schema-validation time.
  Values are written as UTC wall-clock time, since Excel has no timezone
  concept. Use `.dt.tz_convert(None)` beforehand to pick the offset
  explicitly.
- Out-of-range dates (before 1900 or after 9999, which Excel cannot
  represent) raise `ValueError` naming the offending column and row,
  rather than surfacing `rust_xlsxwriter`'s bare
  `"Serial datetime: '-18288' outside ..."` message.
- `Workbook.add_worksheet(constant_memory=True)`: streams a worksheet's
  rows to a temp file instead of buffering the whole sheet in memory,
  via `rust_xlsxwriter`'s `constant_memory` feature. Requires rows to
  be written in non-decreasing order -- enforced by this binding layer
  itself (a clear `ValueError` on violation), since `rust_xlsxwriter`
  does not raise an error for this on its own and would otherwise
  silently produce a corrupt or incomplete `.xlsx` file.
- `Worksheet.autofilter(first_row, first_col, last_row, last_col)`:
  adds Excel's autofilter dropdown controls over a range.
- `Workbook.define_name(name, formula)`: defines a workbook-global or
  sheet-scoped (`"Sheet1!Name"`) named range/formula.
- `Table`/`TableColumn` classes and `Worksheet.add_table()`: full
  worksheet table support -- header row, total row (built-in functions
  or a custom formula), banded rows/columns, first/last column
  styling, autofilter toggle, 61 table styles, per-column formats and
  calculated-column formulas. Two methods on `Table`
  (`set_alt_text()`/`set_alt_text_title()`) exist only in
  `rust_xlsxwriter` 0.96+, not the 0.75 version everything else in
  this project has been stand-in-verified against -- see the note in
  Cargo.toml, they haven't been compiled at all yet, only confirmed
  correct by reading 0.96's source.

### Changed
- Upgraded the pinned `rust_xlsxwriter` version to 0.96 (from 0.75),
  enabling the `zmij` (faster numeric writes) and `constant_memory`
  Cargo features.
- `write_records()`/`write_dataframe()`/`merge_range()`/etc. no longer
  clone the caller's `Format` on every call -- pass a reference
  instead, since `rust_xlsxwriter`'s `write_x_with_format()` /
  `merge_range()` take `&Format`, not an owned value.
- I/O failures on `Workbook.close()` (bad path, permissions, disk
  full) now raise `OSError` instead of the generic `ValueError` used
  for parameter/limit errors, so callers can distinguish the two.
- `merge_range()` now preserves numeric and boolean cell types
  (previously stringified every merged value, which broke `SUM()` over
  a merged numeric range).

### Fixed
- Removed `panic = "abort"` from the release profile: for a PyO3
  extension this turns any Rust panic into a hard crash of the whole
  Python process instead of a catchable exception, which is a
  reliability regression, not a pure performance win.
- Several documentation inaccuracies: an unverified "drop-in
  replacement for Python xlsxwriter" claim (false -- the two projects
  are unrelated and the APIs differ in real ways), an unverifiable
  "most feature-complete"/"full feature parity" superlative (charts,
  conditional formatting, data validation, and tables are all still
  unimplemented), a factually incorrect implication that Python's
  `XlsxWriter` package uses this project's `rust_xlsxwriter` crate (it
  doesn't -- they're separate, unrelated projects), and benchmark
  figures that were presented as current without noting they predated
  this release's `rust_xlsxwriter` upgrade.

### Tests
- Regression tests for the unsafe Arrow PyCapsule ownership-transfer
  code (`write_dataframe()`): repeated calls, multiple worksheets in
  one workbook, and the zero-row edge case.
- Tests locking in the `constant_memory` API contract and its
  row-order enforcement, including the write-column-then-write-earlier-
  row edge case (validates against the *last* row a multi-row call
  touched, not just the first).
- Tests for `autofilter()` (correct range, out-of-range rejection) and
  `define_name()`: global and sheet-scoped names, and the real
  validation rules `rust_xlsxwriter` enforces (name must start with a
  letter or underscore, and can't contain certain characters).
  Duplicate names are NOT rejected -- confirmed that's
  `rust_xlsxwriter`'s own behavior, not a gap in this binding.
- Tests for `Table`/`TableColumn`: basic creation, total row with a
  built-in function (verified the exact generated `SUBTOTAL()`
  formula, not just that it didn't crash), the custom-formula escape
  hatch for both total functions and calculated columns, per-column
  formats, banded rows/columns and other boolean options, style
  validation, and that `Table`/`TableColumn` are importable from the
  package root. Does NOT cover `set_alt_text()`/`set_alt_text_title()`
  -- see the Added section above for why those specifically couldn't
  be tested (or even compiled) at all in this environment.

## [0.1.0] - 2026-07-23

### Added
- Initial release of RVGSRust-XLSXWriter
- Core workbook and worksheet functionality
- Complete formatting API: borders, colors, fonts, alignment, patterns
- Cell merging with support for numeric and boolean cell types
- Formulas and hyperlink support
- Date/time writing capabilities
- Image insertion support
- Sheet operations: freeze panes, hide sheets, set tab colors, sheet protection
- Bulk write operations via `write_records()` for list-of-dicts data
- Zero-copy DataFrame support via `write_dataframe()` with Arrow PyCapsule Interface
- Support for int64, float64, string/utf8, large_utf8, and boolean Arrow column types
- Polars DataFrame integration with automatic Arrow conversion
- Pandas DataFrame integration with Arrow support (2.x+)
- PyArrow Table support
- Multi-threaded sheet assembly during workbook save
- Automatic multi-threading across worksheets (no configuration needed)
- Format method chaining for convenient API
- Comprehensive test suite with openpyxl validation

### Features
- **Performance**: 5-10x faster than pure Python xlsxwriter
- **Zero-copy**: Native Arrow integration eliminates per-cell Python object extraction
- **Easy to use**: Pythonic API with optional DataFrame support
- **Well-tested**: Comprehensive test coverage validating actual cell contents and formatting

### Known Limitations (Roadmap)
- Charts support coming in v0.2
- Conditional formatting coming in v0.2
- Extended Arrow type support (unsigned ints, dates/timestamps, decimals) coming in v0.2
- Per-column DataFrame formatting coming in v0.2
- Data validation, tables, and sparklines planned for v0.3
- Full xlsxwriter API compatibility layer planned for v0.4

### Notes
- Built on top of the official [`rust_xlsxwriter`](https://github.com/jmcnamara/rust_xlsxwriter) crate
- Requires Rust 1.83+ (see CONTRIBUTING.md -- `rust_xlsxwriter 0.96`'s own `zip` dependency needs this) and Python 3.8+
- Platform support: Linux (manylinux2014), macOS, Windows (via maturin builds)
