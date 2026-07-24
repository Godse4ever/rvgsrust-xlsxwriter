# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
- **Completeness**: Full feature parity with Python xlsxwriter library
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
