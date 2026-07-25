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
- `Workbook.add_worksheet(constant_memory=True)`: streams a worksheet's
  rows to a temp file instead of buffering the whole sheet in memory,
  via `rust_xlsxwriter`'s `constant_memory` feature. Requires rows to
  be written in non-decreasing order -- enforced by this binding layer
  itself (a clear `ValueError` on violation), since `rust_xlsxwriter`
  does not raise an error for this on its own and would otherwise
  silently produce a corrupt or incomplete `.xlsx` file.

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
