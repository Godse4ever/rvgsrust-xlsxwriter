# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.14] - 2026-08-29

Patch release, no breaking changes. Adds cell notes.

### Added

- **`Note`.** `set_author()`, `set_width()`, `set_height()`,
  `set_visible()` (notes are hidden-until-hover by default), `set_alt_text()`,
  `set_background_color()`, `set_font_name()`, `set_font_size()` --
  not `set_format()` (a full `Format` override) or
  `set_object_movement()`.
- **`Worksheet.insert_note(row, col, note)`**, **`show_all_notes(enable)`**,
  **`set_default_note_author(name)`**.

## [0.2.13] - 2026-08-29

Patch release, no breaking changes. Adds checkboxes, buttons, and
textbox shapes.

### Added

- **`insert_checkbox(row, col, value)`** / **`insert_checkbox_with_format(row,
  col, value, format)`.** Flat bool + optional `Format`, no new pyclass.
- **`Button`.** `set_caption()`, `set_macro()`, `set_width()`,
  `set_height()`, `set_alt_text()`, via `Worksheet.insert_button()`/
  `insert_button_with_offset()`.
- **`Shape`.** Textbox only -- the only shape type upstream implements.
  `Shape.textbox()`, `set_text()`, `set_width()`, `set_height()`,
  `set_alt_text()`, via `Worksheet.insert_shape()`/
  `insert_shape_with_offset()`. Fill/line/font formatting
  (`ShapeFormat`/`ShapeFont`) not yet exposed.

## [0.2.12] - 2026-08-29

Patch release, no breaking changes. Adds Worksheet image placement.

### Added

- **`insert_image_with_offset(row, col, image_path, x_offset, y_offset)`.**
- **`embed_image(row, col, image_path)`** / **`embed_image_with_format(row,
  col, image_path, format)`.** "Place in Cell" -- only renders in Excel
  365 (2023+); older Excel shows a `#VALUE!` fallback that this binding
  also writes, matching upstream. Uses Excel's rich-value cell-metadata
  mechanism rather than a classic drawing anchor.
- **`insert_image_fit_to_cell(row, col, image_path, keep_aspect_ratio=True)`**
  / **`insert_image_fit_to_cell_centered(row, col, image_path)`.** Scales
  the image to the cell but renders in every Excel version, unlike
  `embed_image()`.
- **`insert_background_image(image_path)`.**
- All five take a plain `image_path: str`, matching the existing
  `insert_image()`'s convention, rather than introducing a Python-facing
  `Image` pyclass -- none of them need to pre-configure the image beyond
  what they already take as parameters.

## [0.2.11] - 2026-08-29

Patch release, no breaking changes. Adds Worksheet protection options.

### Added

- **`ProtectionOptions`.** 15 boolean fields matching upstream's own
  `ProtectionOptions::new()` defaults exactly (`select_locked_cells`/
  `select_unlocked_cells` default to permissive; everything else,
  including `edit_objects`/`edit_scenarios`, defaults to restrictive).
- **`Worksheet.protect_with_options(options, password=None)`.** Sets
  which worksheet elements are protected; the optional password is a
  convenience that also calls `protect_with_password()` in the same
  call (upstream keeps these as two independent methods).
- **`Worksheet.unprotect_range(first_row, first_col, last_row, last_col,
  name="", password="")`.** `protect()`/`protect_with_password()`
  already existed as a single flattened `protect(password=None)`.

## [0.2.10] - 2026-08-29

Patch release, no breaking changes. Adds Worksheet view, visibility,
and sizing setters.

### Added

- **Row/column visibility and pixel sizing.** `set_row_hidden()`,
  `set_row_unhidden()`, `set_column_hidden()`, `hide_unused_rows()`,
  `set_row_height_pixels()`, `set_column_width_pixels()`,
  `set_default_row_height()`.
- **View, zoom, selection, tabs.** `set_zoom()`, `set_selection()`,
  `set_top_left_cell()`, `set_active()`, `set_first_tab()`,
  `set_right_to_left()`, `set_view_normal()`, `set_view_page_layout()`,
  `set_view_page_break_preview()`.
- **Ignored errors.** `ignore_error()`/`ignore_error_range()`, covering
  all 9 upstream `IgnoreError` types as a string parameter (e.g.
  `"number_stored_as_text"`, `"formula_error"`) with validation. Uses
  the same non-mutating row-order guard as `group_rows()` in
  constant_memory mode, so flagging an error on an earlier row after
  writing a later one isn't blocked.
- **Autofit tuning.** `set_autofit_max_width()`, `set_autofit_max_row()`
  (`autofit_to_max_width()` intentionally not bound -- upstream
  deprecates it in favor of these two plus the existing `autofit()`).
- **NaN/infinity display strings.** `set_nan_value()`,
  `set_infinity_value()`, `set_neg_infinity_value()`.
- **`clear_cell()`.** Clears both a cell's value and format
  (`clear_cell_format()` already existed for format-only clearing).

## [0.2.9] - 2026-08-29

Patch release, no breaking changes. Adds chart error bars.

### Added

- **`ChartErrorBars`.** All five upstream types (`set_type_standard_error()`
  -- the default, `set_type_fixed_value()`, `set_type_percentage()`,
  `set_type_standard_deviation()`, `set_type_custom(plus_range,
  minus_range)`), `set_direction()` (both/minus/plus), `set_end_cap()`,
  and `set_format()` (Excel only honors `ChartFormat.set_line()` on error
  bars, per upstream's own doc comment). Attached to a series via the new
  `ChartSeries.set_y_error_bars()`/`set_x_error_bars()`; horizontal bars
  only render in Excel for Scatter and Bar charts.

## [0.2.8] - 2026-08-29

Patch release, no breaking changes. Adds chart secondary axes.

### Added

- **Chart secondary axes.** `set_x2_axis_*`/`set_y2_axis_*`, mirroring
  the existing `set_x_axis_*`/`set_y_axis_*` setters (`date_axis`/
  `text_axis` stay x-only, matching the primary axis). Takes effect once
  a series is routed to the secondary axis via the existing
  `ChartSeries.set_secondary_axis()` -- calling the setters alone with
  no series on the secondary axis emits no secondary-axis XML, matching
  upstream's own `check_for_secondary_axis()` gate.

## [0.2.7] - 2026-08-19

Patch release, no breaking changes. Adds chart secondary axes and
conditional format icon sets.

### Added

- **`ConditionalFormatIconSet` and `ConditionalFormatCustomIcon`.** All
  20 icon set styles (`three_arrows` through `five_boxes`),
  `reverse_icons()`, `show_icons_only()`, and per-icon
  threshold/type/direction overrides via `set_icons()`.

  Real upstream gotcha, worked around rather than just documented: a
  freshly-constructed icon set fails `rust_xlsxwriter`'s own validation
  unless `set_icon_type()` has been called at least once --
  `ConditionalFormatIconSet::new()` leaves its internal icon-rules list
  empty, and `add_conditional_format()` requires exactly 3/4/5 entries
  matching the icon type. This binding's constructor calls
  `set_icon_type()` with upstream's own default internally, so a
  freshly-constructed instance is valid on its own -- the way a caller
  would reasonably expect -- without needing to know about this.

  Also fixed while verifying against source: two icon type names are
  reversed from what they'd suggest if guessed --
  `three_symbols`/`three_symbols_circled` map to the OOXML presets
  `3Symbols2`/`3Symbols` respectively, not the other way around.

## [0.2.6] - 2026-08-19

Patch release, no breaking changes. Adds row/column outline grouping.

### Added

- **`Worksheet.group_rows()`/`group_rows_collapsed()`/`group_columns()`/
  `group_columns_collapsed()`/`group_symbols_above()`/
  `group_symbols_to_left()`.** Collapsible outline sections for
  expandable summary reports. Groups nest up to Excel's 7-level limit.

  **Known limitation:** under `constant_memory=True`,
  `group_rows()`/`group_rows_collapsed()` don't apply per-row grouping
  to the output at all -- verified against actual output, not assumed.
  The call succeeds and the sheet records the correct maximum outline
  level in `sheetFormatPr`, but individual `<row>` elements never get
  an `outlineLevel` attribute, so Excel won't show the per-row
  collapse/expand behavior. Appears to be a `constant_memory` streaming
  limitation in `rust_xlsxwriter` itself, not fixable from this binding
  without writing worksheet XML directly. `group_columns()` is
  unaffected -- columns are a separate `<cols>` section, outside the
  row-streaming mechanism.

  Also added: `check_row_order_readonly()`, a new internal ordering
  guard for `group_rows()`/`group_rows_collapsed()` that checks a range
  hasn't already been written without advancing the `constant_memory`
  high-water mark the way `check_row_order_range()` (used by
  `set_range_format()`/`set_cell_format()`/`add_conditional_format()`)
  does -- the latter would incorrectly reject the normal way nested
  outlines are built (group the outer range, then the inner range) and
  block subsequent writes to earlier, ungrouped rows.

## [0.2.5] - 2026-08-19

Patch release, no breaking changes. Adds data validation.

### Added

- **`DataValidation` and `Worksheet.add_data_validation()`.** Dropdown
  lists (`allow_list_strings()` from a fixed string list, capped at
  Excel's 255-character limit and raising a clean `ValueError` if
  exceeded; `allow_list_formula()` from a cell range, no such limit),
  numeric and text-length range rules (`allow_whole_number()`,
  `allow_decimal_number()`, `allow_text_length()`, all 8 comparison
  types: equal/not-equal/greater/greater-or-equal/less/less-or-equal/
  between/not-between), `allow_custom()` for an arbitrary formula rule,
  `allow_any_value()` to clear a rule while keeping messages, and every
  input/error-message and behaviour setting
  (`ignore_blank`/`show_dropdown`/`show_input_message`/
  `set_input_title`/`set_input_message`/`show_error_message`/
  `set_error_title`/`set_error_message`/`set_error_style`/
  `set_multi_range`).

  Not implemented: `allow_date()`/`allow_time()` range rules (need an
  `ExcelDateTime` construction path from a Python date/time that none
  of the other rules require) and the cell-reference formula variants
  of the numeric/text-length/date/time rules. Dropdown lists -- the
  single most-requested feature this closes -- don't depend on either,
  so shipping without them was a reasonable line to draw. Both are
  candidates for a follow-up; see MISSING.md.

  `add_data_validation()` deliberately does not use the
  `constant_memory` row-order guard that `set_cell_format()`/
  `set_range_format()`/`add_conditional_format()` use: a data
  validation rule is stored in its own independent collection and
  written as a standalone XML block built from row/col numbers, not
  looked up from the per-row buffer `constant_memory` streams to disk,
  so there's nothing for that guard to protect. It can be called
  before, after, or interleaved with writes to the same range.

## [0.2.4] - 2026-08-19

Patch release, no breaking changes. One bug fix, found through an external
evaluation against real SPSS survey exports.

### Fixed

- **`write_dataframe(..., column_formats={...})` no longer materializes a
  cell for every null.** A null value in a column with a `column_formats`
  entry was being written as a formatted blank cell (`write_blank()`)
  instead of being skipped, because it reached the same code path used
  by `write()`/`write_row()`/etc for an explicit `None` + format --
  which is a different, deliberate case (a caller asking for an
  intentionally empty but styled cell) from an absent dataframe value.
  On sparse wide data -- the normal shape for survey exports, where
  most variables are not asked of most respondents -- this multiplied
  file size and write time by the sparsity ratio. Measured on a real
  1,261 x 76,480 export: **29x file size (17.0 MB → 495.8 MB), 84x
  write time (~10.1s → 847.4s)**, verified via cell counts (16,000/16,000
  cells "written" in a row with 115 real values).

  Also fixes the same bug for nullable date/datetime columns even
  without `column_formats` at all -- those columns get an automatic
  date-format entry internally, which triggered the identical
  materialization path.

  Nulls are now always skipped in `write_dataframe()`, regardless of
  column format. `write()`/`write_row()`/`write_records()`/etc are
  unaffected: an explicit `None` paired with a format there still
  produces a formatted blank cell, as before.

### Changed

- **Upgraded `rust_xlsxwriter` 0.96 → 0.98.2** (0.97.x was never
  installed as an intermediate step). MSRV raised **1.85 → 1.88**,
  required by 0.97.0's own `zip` dependency bump — not a choice this
  project made, inherited directly from upstream. See the detailed
  upgrade notes in `Cargo.toml` for what was checked at each version
  and why.

  Two upstream correctness bugs, both silently producing an *invalid*
  Excel file (not an error) rather than corrupting on a code path this
  project already tests: `set_tab_color()` + `set_print_fit_to_pages()`
  together (fixed in 0.97.1), and a sparkline plus a conditional-format
  data bar on the same worksheet (fixed in 0.98.1). Neither combination
  was covered by an existing test, so this upgrade adds one for each.

  Two panic-safety fixes in 0.98.2, relevant to this project's exposed
  API: a truncated PNG/JPEG, or a BMP with a height of `i32::MIN`,
  previously panicked while reading image dimensions in
  `insert_image()`. Because this crate builds with `panic = "unwind"`
  (not `"abort"`), this was never a process-crash risk — PyO3 already
  catches it at the FFI boundary and surfaces it as a Python exception
  — but it was an opaque `PanicException` instead of the clean
  `ValueError`/`OSError` every other image error already produces.
  Also fixed: `define_name()` with an empty sheet or variable name
  created an invalid file instead of erroring; `define_name()` passes
  its arguments straight through with no validation of its own, so
  this was a real, previously-unverified gap.

  `ryu` was already superseded by `zmij` (already enabled) as of
  0.95.0 — nothing to reconsider there. `enhanced_autofit` and
  `rust_decimal` were both considered and left disabled: the former is
  a pure quality enhancement with a new transitive dependency and no
  reported issue driving it; the latter would need explicit
  `decimal.Decimal` handling added to this project's own value
  classifier before the Cargo feature would do anything, which is new
  work beyond a version bump. Both are candidates for their own
  follow-up.

### Added

- **`Format` is now at full parity with upstream, except
  `set_font_scheme()`** (deliberately not exposed — upstream's own doc
  comment calls it "rarely, if ever, required," and it describes a
  font's role in the workbook theme rather than a cell property, an
  odd fit for this API). New: `set_quote_prefix()`, `set_hyperlink()`,
  `set_checkbox()`, `set_font_family()`, `set_font_charset()`,
  `set_font_script()`, `set_reading_direction()` (validated against
  0/1/2 with a clean `ValueError` — upstream itself only prints a
  warning to stderr and silently no-ops on an invalid value, which
  would be invisible from Python), and `unset_bold()`/`unset_italic()`/
  `unset_font_strikethrough()`/`unset_text_wrap()`/`unset_shrink()`/
  `unset_hidden()`/`unset_quote_prefix()`/`unset_checkbox()`/
  `unset_hyperlink_style()`.
- **`Worksheet.set_range_format_with_border(r1, c1, r2, c2, cell_format,
  border_format)`.** Applies interior styling and a border around the
  outside of a range in one call, handling the up-to-9 distinct
  per-position format combinations (corners, edges, interior)
  internally instead of the caller tracking them by hand.
- **`Worksheet.clear_cell_format(row, col)`.** Clears a cell's format
  while leaving its value untouched. A no-op on an unformatted or
  nonexistent cell, not an error.

## [0.2.3] - 2026-08-16

Patch release, no breaking changes. One new feature, one perf/correctness
fix, one new API surface.

### Added

- **`Workbook.save_to_buffer() -> bytes`.** Serializes the workbook to
  bytes instead of writing to a path -- useful for a web response or
  in-memory pipeline with no temp-file round trip. Mirrors `close()`
  exactly: same reentrancy guard, GIL released during the save for the
  same reason (identical serialize+deflate cost).
- **`Worksheet.set_header(text)` / `set_footer(text)`.** Thin wrappers
  over `rust_xlsxwriter`'s infallible header/footer setters. Not
  validated here; note that `&[Page]`/`&[Pages]`/`&[Tab]` bracket-style
  placeholders are normalized by the underlying crate to the older
  single-letter codes (`&P`/`&N`/`&A`) before writing -- both spellings
  are valid Excel codes and render identically, the file on disk just
  won't preserve which one was typed. Header/footer *images* still need
  an `Image` pyclass and remain unimplemented.

### Fixed

- **`write_rows()` no longer holds the dataset in memory twice.** It
  previously classified every row into a `Vec<(Vec<CellValue>, bool)>`
  before writing anything, then wrote it in a second pass -- doubling
  peak memory for the whole call, and doing the opposite of what
  `write_records()`'s own comment argues for. Now interleaves classify
  and write in a single pass, matching `write_records()`.

  This is also a small, deliberate behavior change: the old two-pass
  version had an incidental all-or-nothing guarantee on error (nothing
  written if a row failed to classify). The interleaved version doesn't
  have that -- rows before a failure are already on the sheet, same as
  `write_records()` already behaves, and now the same as each other. No
  existing code depended on the old atomicity; the new behavior is
  covered by a test rather than left implicit.

## [0.2.2] - 2026-08-15

Patch release, no breaking changes. Started as a fix for one API
annoyance (`close()` requiring the path twice) and grew to cover a
round of gaps found through real-world testing against a wide (1,261 ×
76,480) production workload.

### Fixed

- **`Workbook.close()` now uses the constructor-provided path when
  called with no argument.** `with Workbook("out.xlsx") as wb:` and
  `wb.close()` after constructing with a path no longer raise
  `TypeError`. `close(path)` still works as an explicit override, even
  when the constructor was also given a path. `close()` with no path
  anywhere raises `ValueError` instead of a bare `TypeError`.
- **`write_dataframe(column_formats=...)`** — new keyword argument, a
  `dict[str, Format]` keyed by column name. Merges the format into
  every cell of that column *as it's written*, once per column before
  the batch loop starts, not per cell. This is the actual fix for a
  correctness bug: applying a format to a whole column afterward (via
  `set_column_format()`/`set_column_range_format()`) loses to a cell's
  own number format under OOXML precedence, so a border on a date or
  datetime column was silently dropped with no error. `column_formats`
  sidesteps that by merging into the same cellXf instead of competing
  for it. Works under `constant_memory=True`. An unknown column name
  raises `ValueError` before any row is written. Omitting the argument
  is byte-identical to 0.2.1's behaviour.
- **`set_column_range_width(first_col, last_col, width)`** — new
  method, mirrors `set_column_range_format()`. Setting a uniform width
  across a wide sheet previously cost one Python→Rust call per column;
  this delegates to `rust_xlsxwriter`'s native range-width call.
- **`annotations` no longer leaks into the module namespace.** It was
  showing up in `dir(rvgsrust_xlsxwriter)` and `from ... import *` as a
  side effect of `from __future__ import annotations`.

### Added

- **`set_border_top()` / `set_border_bottom()` / `set_border_left()` /
  `set_border_right()`** as the canonical `Format` border-style names,
  matching `set_border_color()` / `set_border_diagonal()` /
  `set_border_*_color()`. The original `set_top_border()` /
  `set_bottom_border()` / `set_left_border()` / `set_right_border()`
  spellings (which reversed that word order) are kept working,
  documented as deprecated aliases rather than removed.
- **`.pyi` type stubs** (`_core.pyi`, `__init__.pyi`) and a `py.typed`
  marker. `Literal` types for every accepted enum string (border style,
  align, pattern, chart type, and so on) were read directly out of the
  `parse_*()` match arms in `src/lib.rs`, including the two `Format`
  setters (`set_align`, `set_pattern`) that silently fall back to a
  default on an unrecognised string instead of raising — documented in
  the stub since a type checker won't catch that either.
- **`Cargo.lock`**, committed for reproducible builds.

### Documentation

- README: benchmark claim now states its shape (100k rows × 8 cols vs
  pure-Python `xlsxwriter`) and adds the comparison against another
  Rust-backed writer on a wide workload (~12% faster on time; the
  larger win is GIL release during `save()` — ~20ms worst-case stall
  vs ~1,300ms, now its own row in the feature comparison table).
- README/MISSING.md: corrected an earlier audit pass that had flagged
  per-side border *styles* as entirely missing, which was a false
  positive from a naming mismatch, not an actual gap; MISSING.md
  trimmed to list only genuinely open items rather than tracking closed
  ones inline.

### Tests

- Full coverage added for every item above, including a `styles.xml`
  parity check that a single shared `Format` instance passed across
  columns of different dtypes produces one border definition, not one
  per column.

## [0.2.1] - 2026-07-29

Audit release. No public API changes; every item below is an internal fix or improvement, so upgrading from 0.2.0 requires no code changes.

### Fixed

- **`add_worksheet()` wrote to the wrong worksheet after a rejected name.**
  The worksheet was appended, then `set_name()` was called, then an index
  counter advanced. A rejected name (blank, >31 chars, or containing
  `* ? : [ ] \ /`) returned early between those steps, leaving an orphan
  worksheet while the counter stood still. The next `add_worksheet()`
  handed back that stale index, so writes landed on the orphan and the
  intended sheet saved empty -- with no exception raised. The name is now
  validated before the workbook is touched, and the index comes from the
  workbook's own worksheet vector rather than a parallel counter.
- **`add_table()` bypassed the `constant_memory` row-order guard.** It
  writes header cells at call time, so a table anchored above the current
  high-water mark silently produced a corrupt file.
- **Re-entrant workbook access panicked.** A `__str__` that touched the
  same `Workbook` during a bulk write hit `BorrowMutError`. It is now a
  `RuntimeError` that names the cause.
- **A partially-written `write_dataframe()` could be silently retried and
  duplicated.** A mid-stream schema disagreement now raises `RuntimeError`
  rather than `TypeError`, so `dataframe.py`'s per-cell fallback cannot
  rewrite rows already on the sheet.
- **`__version__` disagreed with the build files** (`0.2.0.dev0` vs
  `0.2.0`), and the module docstring claimed charts and conditional
  formatting were unimplemented while exporting them.
- **`rust-version` was wrong.** It declared 1.83; the real floor is 1.85,
  because `indexmap 2.14` requires `edition2024`, which Cargo 1.83 cannot
  parse. Now declared and enforced by a CI job.

### Performance

- **The GIL is released during `save()`.** Serialisation and deflation
  touch no Python objects but previously blocked every other thread for
  their full duration -- seconds on a large workbook.
- **Arrow string columns no longer allocate per cell.** `CellValue::Str`
  is now `Cow<str>`, so the Arrow path borrows the columnar buffer
  directly instead of allocating and dropping a `String` for every cell
  (one million allocations for a one-million-row string column).
- **`write_dataframe()` streams record batches** instead of collecting
  them. Peak memory is now bounded by one batch rather than the whole
  dataset, which makes `constant_memory=True` genuinely O(1) for
  streaming producers such as a `pyarrow.RecordBatchReader` over Parquet.
  Schema validation still happens fully up front.

### CI / packaging

- Tests run on Linux (3.9, 3.12, 3.13), macOS and Windows. Previously
  Linux/3.11 only, despite shipping wheels for three platforms -- this
  immediately surfaced a Windows-only test-teardown bug.
- `polars` is now installed in CI; its `.to_arrow()` path was never
  exercised before.
- Added an MSRV job, an sdist to the release, and a separate `lint` job.
- Removed `MANIFEST.in`, which has no effect under the maturin backend.
- 17 regression tests added; the suite is 448 passing on all platforms.

## [0.2.0] - 2026-07-28

**Core build confirmed on real hardware** (macOS, 4-core, rustc 1.83+,
`maturin develop --release`): `rust_xlsxwriter 0.96` compiles cleanly
with no source changes and no sandbox-specific dependency pins, and a
basic Workbook/add_worksheet/write/close smoke test passes. This was
run against an earlier point in this entry's history, before the
`constant_memory` row-order enforcement below and the version bump
itself -- worth re-confirming after pulling latest, but the
fundamental "does this even build" risk this project's development
sandbox couldn't resolve is now answered: yes.

### Added
- `Format` completions: per-side border colours
  (`set_border_top_color` and the other three sides), diagonal borders
  (`set_border_diagonal`, `set_border_diagonal_color`,
  `set_border_diagonal_type`), cell protection (`set_locked`,
  `set_unlocked`, `set_hidden`), `set_font_strikethrough` and
  `set_foreground_color`.
  Note the per-side border *styles* were already exposed all along, as
  `set_top_border` and friends, which reverse upstream's `set_border_top`
  word order. `MISSING.md` had wrongly listed them as missing and ranked
  them highest-value; that is corrected, along with the same mistake for
  `freeze_panes`.
- Page setup and print settings: all 19 `Worksheet` methods, mirroring
  upstream names 1:1 -- orientation, paper size, margins, print area,
  repeat rows and columns, fit-to-pages, print scale, horizontal and
  vertical page breaks, gridlines, headings, centring, black and white,
  draft and first page number. No new classes needed.
  These set worksheet metadata rather than writing cells, so none are
  guarded by the constant-memory row-order check even where they take row
  numbers. Closes `MISSING.md` section 2b.
- `Worksheet.set_column_format`, `set_column_range_format`,
  `set_row_format`, `set_cell_format` and `set_range_format`. Column and
  row formats apply to cells without a format of their own, so these are
  the way to reformat data after `write_dataframe` has written it -- the
  gap `MISSING.md` ranked first, and the reason `write_dataframe` applies
  its date formats itself rather than leaving them to the caller.
  `set_row_format` and `set_cell_format` are guarded by the constant-memory
  row-order check; the column variants target no particular row and so are
  not.
- Added `MISSING.md`, a parity audit of the exposed Python API against
  rust_xlsxwriter 0.96, with upstream `file:line` references, a suggested
  Python API shape and a priority for every gap. Listing only; nothing in
  it is implemented.
  Two items the project brief listed as gaps turned out to be false: a
  sparkline `custom_ranges` parameter does not exist in 0.96 at all, and
  workbook document properties and defined names are already exposed.
  Sparklines are at full parity, 28 of 28 methods.
- Charts, part 3: `ChartMarker`, `ChartTrendline` and `ChartDataLabel`,
  attached via `ChartSeries.set_marker`, `set_trendline`, `set_data_label`
  and `set_custom_data_labels`. Note `Chart.push_series()` does not call
  upstream's `push_series()`: that applies the chart-type defaults after
  copying the series, so on a line, radar or scatter-straight/smooth chart
  it silently overwrites a marker set beforehand. It goes through
  `add_series()` instead, which restores upstream's intended
  defaults-first, caller-wins precedence.
  Per-point labels are marked with
  `ChartDataLabel.set_custom()`, named for upstream's `to_custom()` but
  renamed because `clippy::wrong_self_convention` forbids a `to_*` method
  taking `&mut self`. All three accept a `ChartFormat`, and the
  trendline and data label also accept a `ChartFont`.
  Automatic and no-marker are methods (`set_automatic()`, `set_none()`)
  rather than marker types, matching upstream. Trendline `polynomial` and
  `moving_average` take a period, defaulting to 2. `display_equation` and
  `display_r_squared` are exposed with a `set_` prefix for consistency
  with the rest of the binding, though upstream omits it.
  This completes the chart work begun in parts 1 and 2. Still outstanding
  for the parity audit: icon sets, pattern and gradient fills, error bars,
  drop lines, high-low lines, up-down bars, data tables, combined charts,
  and chart/plot area formatting.
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
- Sparklines: a single `Sparkline` class plus
  `Worksheet.add_sparkline(row, col, sparkline)` and
  `Worksheet.add_sparkline_group(first_row, first_col, last_row, last_col,
  sparkline)`. Covers the three types, all seven point-marker toggles, all
  seven colors, line weight, custom and group min/max, style presets, date
  ranges, right-to-left, column order, and empty-cell handling.
  As with the conditional formats, enum-valued options are strings
  validated with a `ValueError` that lists the accepted values. The type
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
  are unrelated, and the APIs differ in real ways), an unverifiable
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
