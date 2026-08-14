1| # Changelog
2| 
3| All notable changes to this project will be documented in this file.
4| 
5| The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
6| and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
7| 
8| ## [0.2.2] - 2026-08-14
9| 
10| Patch release: minor fixes, packaging and documentation updates, and developer convenience improvements. This release is intended to be a quick follow-up to 0.2.1 to fix a user-facing API annoyance and prepare the project for an audited rebuild.
11| 
12| ### Fixed
13| 
14| - **`Workbook.close()` now uses the constructor-provided path when called with no argument.** Calling `with Workbook("out.xlsx") as wb:` or calling `wb.close()` after constructing with a path no longer raises a `TypeError`. The explicit `close(path)` form continues to work as an override.
15| - **Python package version and module version aligned to 0.2.2.** Updated `pyproject.toml` and `python/rvgsrust_xlsxwriter/__init__.py`.
16| 
17| ### Documentation & UX
18| 
19| - Added minimal type stubs (.pyi) for common API shapes so IDEs and type checkers show accepted literal options for Format-like setters (border, align, pattern, underline) and the Worksheet.write_dataframe() convenience signature including `column_formats` as an optional fallback argument.
20| - Clarified README benchmark language and called out the GIL-release behaviour during save() as a desktop-app responsiveness advantage.
21| 
22| ### Packaging
23| 
24| - Bumped package metadata to 0.2.2 in pyproject.toml.
25| - Note: Cargo.lock must be generated and committed from a machine capable of resolving the pinned Rust dependencies (see MSRV notes in Cargo.toml). This lockfile is not yet committed in this patch; generate it with `cargo generate-lockfile` on a compatible toolchain (Rust >= MSRV) and commit the resulting `Cargo.lock` before building release artifacts.
26| 
27| ### Release notes
28| 
29| - This release prepares a tag `v0.2.2` and a GitHub Release draft. Building the native wheels and publishing to PyPI requires a suitable Rust toolchain (see Cargo.toml MSRV notes) and PyPI credentials (API token). See RELEASE_NOTES_0.2.2.md for a full publish checklist.
30| 
31| ### Tests & Build
32| 
33| - Please run the full test suite (pytest) and a `maturin develop --release` build locally or in CI that has the required Rust toolchain before publishing to PyPI. The CI in this repo is configured to validate MSRV and run the test matrix.
34| 
35| ## [0.2.1] - 2026-07-29
36| 
37| Audit release. No public API changes; every item below is an internal fix or improvement, so upgrading from 0.2.0 requires no code changes.
38| 
39| ### Fixed
40| 
40| - **`add_worksheet()` wrote to the wrong worksheet after a rejected name.**
41|   The worksheet was appended, then `set_name()` was called, then an index
42|   counter advanced. A rejected name (blank, >31 chars, or containing
43|   `* ? : [ ] \ /`) returned early between those steps, leaving an orphan
44|   worksheet while the counter stood still. The next `add_worksheet()`
45|   handed back that stale index, so writes landed on the orphan and the
46|   intended sheet saved empty -- with no exception raised. The name is now
47|   validated before the workbook is touched, and the index comes from the
48|   workbook's own worksheet vector rather than a parallel counter.
49| - **`add_table()` bypassed the `constant_memory` row-order guard.** It
50|   writes header cells at call time, so a table anchored above the current
51|   high-water mark silently produced a corrupt file.
52| - **Re-entrant workbook access panicked.** A `__str__` that touched the
53|   same `Workbook` during a bulk write hit `BorrowMutError`. It is now a
54|   `RuntimeError` that names the cause.
55| - **A partially-written `write_dataframe()` could be silently retried and
56|   duplicated.** A mid-stream schema disagreement now raises `RuntimeError`
57|   rather than `TypeError`, so `dataframe.py`'s per-cell fallback cannot
58|   rewrite rows already on the sheet.
59| - **`__version__` disagreed with the build files** (`0.2.0.dev0` vs
60|   `0.2.0`), and the module docstring claimed charts and conditional
61|   formatting were unimplemented while exporting them.
61| - **`rust-version` was wrong.** It declared 1.83; the real floor is 1.85,
62|   because `indexmap 2.14` requires `edition2024`, which Cargo 1.83 cannot
63|   parse. Now declared and enforced by a CI job.
64| 
65| ### Performance
66| 
67| - **The GIL is released during `save()`.** Serialisation and deflation
68|   touch no Python objects but previously blocked every other thread for
69|   their full duration -- seconds on a large workbook.
70| - **Arrow string columns no longer allocate per cell.** `CellValue::Str`
71|   is now `Cow<str>`, so the Arrow path borrows the columnar buffer
72|   directly instead of allocating and dropping a `String` for every cell
73|   (one million allocations for a one-million-row string column).
74| - **`write_dataframe()` streams record batches** instead of collecting
75|   them. Peak memory is now bounded by one batch rather than the whole
76|   dataset, which makes `constant_memory=True` genuinely O(1) for
77|   streaming producers such as a `pyarrow.RecordBatchReader` over Parquet.
78|   Schema validation still happens fully up front.
79| 
80| ### CI / packaging
81| 
82| - Tests run on Linux (3.9, 3.12, 3.13), macOS and Windows. Previously
83|   Linux/3.11 only, despite shipping wheels for three platforms -- this
84|   immediately surfaced a Windows-only test-teardown bug.
85| - `polars` is now installed in CI; its `.to_arrow()` path was never
86|   exercised before.
87| - Added an MSRV job, an sdist to the release, and a separate `lint` job.
88| - Removed `MANIFEST.in`, which has no effect under the maturin backend.
89| - 17 regression tests added; the suite is 448 passing on all platforms.
90| 
91| ## [0.2.0] - 2026-07-28
92| 
93| **Core build confirmed on real hardware** (macOS, 4-core, rustc 1.83+,
94| `maturin develop --release`): `rust_xlsxwriter 0.96` compiles cleanly
95| with no source changes and no sandbox-specific dependency pins, and a
96| basic Workbook/add_worksheet/write/close smoke test passes. This was
97| run against an earlier point in this entry's history, before the
98| `constant_memory` row-order enforcement below and the version bump
99| itself -- worth re-confirming after pulling latest, but the
100| fundamental "does this even build" risk this project's development
101| sandbox couldn't resolve is now answered: yes.
102| 
103| ### Added
104| - `Format` completions: per-side border colours
105|   (`set_border_top_color` and the other three sides), diagonal borders
106|   (`set_border_diagonal`, `set_border_diagonal_color`,
107|   `set_border_diagonal_type`), cell protection (`set_locked`,
108|   `set_unlocked`, `set_hidden`), `set_font_strikethrough` and
109|   `set_foreground_color`.
110|   Note the per-side border *styles* were already exposed all along, as
111|   `set_top_border` and friends, which reverse upstream's `set_border_top`
112|   word order. `MISSING.md` had wrongly listed them as missing and ranked
113|   them highest-value; that is corrected, along with the same mistake for
114|   `freeze_panes`.
115| - Page setup and print settings: all 19 `Worksheet` methods, mirroring
116|   upstream names 1:1 -- orientation, paper size, margins, print area,
117|   repeat rows and columns, fit-to-pages, print scale, horizontal and
118|   vertical page breaks, gridlines, headings, centring, black and white,
119|   draft and first page number. No new classes needed.
120|   These set worksheet metadata rather than writing cells, so none are
121|   guarded by the constant-memory row-order check even where they take row
122|   numbers. Closes `MISSING.md` section 2b.
123| - `Worksheet.set_column_format`, `set_column_range_format`,
124|   `set_row_format`, `set_cell_format` and `set_range_format`. Column and
125|   row formats apply to cells without a format of their own, so these are
126|   the way to reformat data after `write_dataframe` has written it -- the
127|   gap `MISSING.md` ranked first, and the reason `write_dataframe` applies
128|   its date formats itself rather than leaving them to the caller.
129|   `set_row_format` and `set_cell_format` are guarded by the constant-memory
130|   row-order check; the column variants target no particular row and so are
131|   not.
132| - Added `MISSING.md`, a parity audit of the exposed Python API against
133|   rust_xlsxwriter 0.96, with upstream `file:line` references, a suggested
134|   Python API shape and a priority for every gap. Listing only; nothing in
135|   it is implemented.
136|   Two items the project brief listed as gaps turned out to be false: a
137|   sparkline `custom_ranges` parameter does not exist in 0.96 at all, and
138|   workbook document properties and defined names are already exposed.
139|   Sparklines are at full parity, 28 of 28 methods.
140| - Charts, part 3: `ChartMarker`, `ChartTrendline` and `ChartDataLabel`,
141|   attached via `ChartSeries.set_marker`, `set_trendline`, `set_data_label`
142|   and `set_custom_data_labels`. Note `Chart.push_series()` does not call
143|   upstream's `push_series()`: that applies the chart-type defaults after
144|   copying the series, so on a line, radar or scatter-straight/smooth chart
145|   it silently overwrites a marker set beforehand. It goes through
146|   `add_series()` instead, which restores upstream's intended
147|   defaults-first, caller-wins precedence.
148|   Per-point labels are marked with
149|   `ChartDataLabel.set_custom()`, named for upstream's `to_custom()` but
150|   renamed because `clippy::wrong_self_convention` forbids a `to_*` method
151|   taking `&mut self`. All three accept a `ChartFormat`, and the
152|   trendline and data label also accept a `ChartFont`.
153|   Automatic and no-marker are methods (`set_automatic()`, `set_none()`)
154|   rather than marker types, matching upstream. Trendline `polynomial` and
155|   `moving_average` take a period, defaulting to 2. `display_equation` and
156|   `display_r_squared` are exposed with a `set_` prefix for consistency
157|   with the rest of the binding, though upstream omits it.
158|   This completes the chart work begun in parts 1 and 2. Still outstanding
159|   for the parity audit: icon sets, pattern and gradient fills, error bars,
160|   drop lines, high-low lines, up-down bars, data tables, combined charts,
161|   and chart/plot area formatting.
162| - Charts, part 2: `ChartFormat` and `ChartFont`, attachable to a series,
163|   a chart title, either axis (both the axis labels and the axis name), and
164|   the legend.
165|   `ChartLine` and `ChartSolidFill` are not exposed as separate classes.
166|   Upstream they exist only to be passed to `ChartFormat`, so they are
167|   flattened into it as `set_line_*`, `set_border_*` and `set_fill_*`, with
168|   the line and fill state kept per format object so successive calls
169|   compose. Pattern and gradient fills are logged for the parity audit.
170|   `set_format` is generic over `IntoChartFormat`, which upstream implements
171|   for `&mut ChartFormat`, so each call site passes an owned mutable clone.
172|   The trait itself never needs importing despite not being re-exported
173|   from the crate root, since it only ever appears as a bound on a generic
174|   parameter.
175| - Charts, part 1: `Chart` and `ChartSeries` classes plus
176|   `Worksheet.insert_chart(row, col, chart, x_offset=0, y_offset=0)`.
177|   Covers all 23 chart types, series ranges and options, and the title,
178|   x/y axis and legend settings.
179|   Axes, titles and legends are flattened onto `Chart` as `set_x_axis_*`,
180|   `set_title_*` and `set_legend_*` rather than being separate classes,
181|   because `ChartAxis::new`, `ChartTitle::new` and `ChartLegend::new` are
182|   `pub(crate)` upstream and cannot be constructed from a binding.
183|   Series are attached with `Chart.push_series(series)` rather than being
184|   passed to `insert_chart`. `Chart` does not derive `Clone` upstream, so
185|   pushing at insert time would mutate the only copy and silently duplicate
186|   series if the same chart were inserted twice, with no `remove_series` to
187|   undo it.
188| - Sparklines: a single `Sparkline` class plus
189|   `Worksheet.add_sparkline(row, col, sparkline)` and
190|   `Worksheet.add_sparkline_group(first_row, first_col, last_row, last_col,
191|   sparkline)`. Covers the three types, all seven point-marker toggles, all
192|   seven colors, line weight, custom and group min/max, style presets, date
193|   ranges, right-to-left, column order, and empty-cell handling.
194|   As with the conditional formats, enum-valued options are strings
195|   validated with a `ValueError` that lists the accepted values. The type
196|   accepts `win_lose` (upstream spells the variant `WinLose`) and also
197|   `win_loss`, since that is the spelling most people reach for first; both
198|   serialize to Excel's `type="stacked"`.
199|   Grouped sparklines require a 2D data range, one row per sparkline;
200|   passing a 1D range raises `ValueError`, as does adding a sparkline with
201|   no range set.
202| - Conditional formatting: 12 rule types, each a `#[pyclass]` wrapping the
203|   matching `rust_xlsxwriter` builder -- `ConditionalFormatCell`, `Blank`,
204|   `Duplicate`, `Error`, `Formula`, `Average`, `Top`, `Text`, `Date`,
205|   `2ColorScale`, `3ColorScale` and `DataBar`. Applied with
206|   `Worksheet.add_conditional_format(first_row, first_col, last_row,
207|   last_col, rule)`. Icon sets are deferred to the parity audit.
208|   Enum-valued options (average/date/text/top rules, scale value types,
209|   data bar direction and axis position) are taken as strings and validated
210|   with a `ValueError` that lists the accepted values, rather than being
211|   exposed as separate enum classes.
212|   Note these setters return `None` rather than `self`, so unlike `Format`
213|   they do not chain; adding a return value later is backwards compatible.
214| - Extended Arrow type support in `Worksheet.write_dataframe()`. Added
215|   `int8`/`int16`/`int32`, `uint8`/`uint16`/`uint32`/`uint64`, `float32`,
215|   `date32`/`date64`, and `timestamp` in all four units (second,
216|   millisecond, microsecond, nanosecond). `timestamp[ns]` matters most in
217|   practice: it is what pandas' default `datetime64[ns]` dtype maps to, so
218|   the most common real-world DataFrame previously raised `TypeError` here
219|   and silently fell back to the per-cell path in `dataframe.py`, where
220|   datetimes were written as strings. They are now real Excel dates.
221| - Date and timestamp columns get a number format applied automatically
222|   (`yyyy-mm-dd` and `yyyy-mm-dd hh:mm:ss` respectively). Without one Excel
223|   renders a date serial as a bare number such as `45123`, and this binding
224|   exposes no `set_column_format()` for the caller to fix it afterwards.
225|   The two formats are built once per `write_dataframe()` call and resolved
226|   per column, not per cell, so the inner loop cost is unchanged for
227|   non-temporal data.
228| - Timezone-aware `timestamp` columns now emit a `UserWarning` naming the
229|   column and its timezone, once per column at schema-validation time.
230|   Values are written as UTC wall-clock time, since Excel has no timezone
231|   concept. Use `.dt.tz_convert(None)` beforehand to pick the offset
232|   explicitly.
233| - Out-of-range dates (before 1900 or after 9999, which Excel cannot
234|   represent) raise `ValueError` naming the offending column and row,
235|   rather than surfacing `rust_xlsxwriter`'s bare
236|   `"Serial datetime: '-18288' outside ..."` message.
237| - `Workbook.add_worksheet(constant_memory=True)`: streams a worksheet's
238|   rows to a temp file instead of buffering the whole sheet in memory,
239|   via `rust_xlsxwriter`'s `constant_memory` feature. Requires rows to
240|   be written in non-decreasing order -- enforced by this binding layer
241|   itself (a clear `ValueError` on violation), since `rust_xlsxwriter`
242|   does not raise an error for this on its own and would otherwise
243|   silently produce a corrupt or incomplete `.xlsx` file.
244| - `Worksheet.autofilter(first_row, first_col, last_row, last_col)`:
245|   adds Excel's autofilter dropdown controls over a range.
246| - `Workbook.define_name(name, formula)`: defines a workbook-global or
247|   sheet-scoped (`"Sheet1!Name"`) named range/formula.
248| - `Table`/`TableColumn` classes and `Worksheet.add_table()`: full
249|   worksheet table support -- header row, total row (built-in functions
250|   or a custom formula), banded rows/columns, first/last column
251|   styling, autofilter toggle, 61 table styles, per-column formats and
252|   calculated-column formulas. Two methods on `Table`
253|   (`set_alt_text()`/`set_alt_text_title()`) exist only in
254|   `rust_xlsxwriter` 0.96+, not the 0.75 version everything else in
255|   this project has been stand-in-verified against -- see the note in
256|   Cargo.toml, they haven't been compiled at all yet, only confirmed
257|   correct by reading 0.96's source.
258| 
259| ### Changed
260| - Upgraded the pinned `rust_xlsxwriter` version to 0.96 (from 0.75),
261|   enabling the `zmij` (faster numeric writes) and `constant_memory`
262|   Cargo features.
263| - `write_records()`/`write_dataframe()`/`merge_range()`/etc. no longer
264|   clone the caller's `Format` on every call -- pass a reference
265|   instead, since `rust_xlsxwriter`'s `write_x_with_format()` /
266|   `merge_range()` take `&Format`, not an owned value.
267| - I/O failures on `Workbook.close()` (bad path, permissions, disk
268|   full) now raise `OSError` instead of the generic `ValueError` used
269|   for parameter/limit errors, so callers can distinguish the two.
270| - `merge_range()` now preserves numeric and boolean cell types
271|   (previously stringified every merged value, which broke `SUM()` over
272|   a merged numeric range).
273| 
274| ### Fixed
275| - Removed `panic = "abort"` from the release profile: for a PyO3
276|   extension this turns any Rust panic into a hard crash of the whole
277|   Python process instead of a catchable exception, which is a
278|   reliability regression, not a pure performance win.
279| - Several documentation inaccuracies: an unverified "drop-in
280|   replacement for Python xlsxwriter" claim (false -- the two projects
281|   are unrelated, and the APIs differ in real ways), an unverifiable
282|   "most feature-complete"/"full feature parity" superlative (charts,
283|   conditional formatting, data validation, and tables are all still
284|   unimplemented), a factually incorrect implication that Python's
285|   `XlsxWriter` package uses this project's `rust_xlsxwriter` crate (it
286|   doesn't -- they're separate, unrelated projects), and benchmark
287|   figures that were presented as current without noting they predated
288|   this release's `rust_xlsxwriter` upgrade.
289| 
290| ### Tests
291| - Regression tests for the unsafe Arrow PyCapsule ownership-transfer
292|   code (`write_dataframe()`): repeated calls, multiple worksheets in
293|   one workbook, and the zero-row edge case.
294| - Tests locking in the `constant_memory` API contract and its
295|   row-order enforcement, including the write-column-then-write-earlier-
296|   row edge case (validates against the *last* row a multi-row call
297|   touched, not just the first).
298| - Tests for `autofilter()` (correct range, out-of-range rejection) and
299|   `define_name()`: global and sheet-scoped names, and the real
300|   validation rules `rust_xlsxwriter` enforces (name must start with a
301|   letter or underscore, and can't contain certain characters).
302|   Duplicate names are NOT rejected -- confirmed that's
303|   `rust_xlsxwriter`'s own behavior, not a gap in this binding.
304| - Tests for `Table`/`TableColumn`: basic creation, total row with a
305|   built-in function (verified the exact generated `SUBTOTAL()`
306|   formula, not just that it didn't crash), the custom-formula escape
307|   hatch for both total functions and calculated columns, per-column
308|   formats, banded rows/columns and other boolean options, style
309|   validation, and that `Table`/`TableColumn` are importable from the
310|   package root. Does NOT cover `set_alt_text()`/`set_alt_text_title()`
311|   -- see the Added section above for why those specifically couldn't
312|   be tested (or even compiled) at all in this environment.
313| 
314| ## [0.1.0] - 2026-07-23
314| 
315| ### Added
316| - Initial release of RVGSRust-XLSXWriter
317| - Core workbook and worksheet functionality
318| - Complete formatting API: borders, colors, fonts, alignment, patterns
316| - Cell merging with support for numeric and boolean cell types
317| - Formulas and hyperlink support
318| - Date/time writing capabilities
319| - Image insertion support
320| - Sheet operations: freeze panes, hide sheets, set tab colors, sheet protection
321| - Bulk write operations via `write_records()` for list-of-dicts data
322| - Zero-copy DataFrame support via `write_dataframe()` with Arrow PyCapsule Interface
323| - Support for int64, float64, string/utf8, large_utf8, and boolean Arrow column types
324| - Polars DataFrame integration with automatic Arrow conversion
325| - Pandas DataFrame integration with Arrow support (2.x+)
326| - PyArrow Table support
327| - Multi-threaded sheet assembly during workbook save
328| - Automatic multi-threading across worksheets (no configuration needed)
329| - Format method chaining for convenient API
330| - Comprehensive test suite with openpyxl validation
