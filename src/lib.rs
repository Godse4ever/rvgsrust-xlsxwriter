use pyo3::prelude::*;
use pyo3::types::{PyCapsule, PyDict, PyDictMethods, PyList, PyListMethods, PyString};
use pyo3::PyRefMut;
use arrow::array::{Array, BooleanArray, Float64Array, Int64Array, LargeStringArray, RecordBatch, StringArray};
use arrow::datatypes::DataType as ArrowDataType;
use arrow::ffi_stream::{ArrowArrayStreamReader, FFI_ArrowArrayStream};
use rust_xlsxwriter::{
    Color, FormatAlign, FormatBorder, FormatPattern,
    Workbook as RustWorkbook, Worksheet as RustWorksheet, Format as RustFormat,
};
use std::cell::{Cell, RefCell};

// ============================================
// ERROR HELPER
// ============================================
// Converts any rust_xlsxwriter error (or other Display error) into a
// proper Python exception instead of being silently discarded, which
// is what the previous version did everywhere via `let _ = ...`.
fn to_pyerr<E: std::fmt::Display>(e: E) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(e.to_string())
}

// rust_xlsxwriter::XlsxError gets a more specific mapping than the
// generic to_pyerr() above: most of its variants (bad row/col, reused
// sheet name, merge-range overlap, etc.) really are ValueErrors -- a
// caller passed something invalid. But XlsxError::IoError wraps a
// std::io::Error from the underlying `save()`/file write (permissions,
// disk full, path doesn't exist) -- that's a different failure
// category a caller would reasonably want to catch separately (e.g.
// `except OSError` around a save() call vs `except ValueError` around
// a write() call), so it maps to Python's OSError instead. This is the
// path most likely to actually hit IoError in practice: see
// Workbook.close()'s call to save() below.
fn xlsx_err_to_pyerr(e: rust_xlsxwriter::XlsxError) -> PyErr {
    match e {
        rust_xlsxwriter::XlsxError::IoError(io_err) => {
            PyErr::new::<pyo3::exceptions::PyOSError, _>(io_err.to_string())
        }
        other => PyErr::new::<pyo3::exceptions::PyValueError, _>(other.to_string()),
    }
}

// ============================================
// COLOR HELPER
// ============================================
fn parse_color(color: &str) -> Color {
    if color.starts_with('#') && color.len() == 7 {
        if let (Ok(r), Ok(g), Ok(b)) = (
            u8::from_str_radix(&color[1..3], 16),
            u8::from_str_radix(&color[3..5], 16),
            u8::from_str_radix(&color[5..7], 16),
        ) {
            let rgb = ((r as u32) << 16) | ((g as u32) << 8) | (b as u32);
            return Color::RGB(rgb);
        }
    }
    match color.to_lowercase().as_str() {
        "black" => Color::Black,
        "blue" => Color::Blue,
        "brown" => Color::Brown,
        "cyan" => Color::Cyan,
        "gray" | "grey" => Color::Gray,
        "green" => Color::Green,
        "lime" => Color::Lime,
        "magenta" => Color::Magenta,
        "navy" => Color::Navy,
        "orange" => Color::Orange,
        "pink" => Color::Pink,
        "purple" => Color::Purple,
        "red" => Color::Red,
        "silver" => Color::Silver,
        "white" => Color::White,
        "yellow" => Color::Yellow,
        _ => Color::Black,
    }
}

fn parse_border(border: &str) -> FormatBorder {
    match border.to_lowercase().as_str() {
        "thin" => FormatBorder::Thin,
        "medium" => FormatBorder::Medium,
        "thick" => FormatBorder::Thick,
        "dashed" => FormatBorder::Dashed,
        "dotted" => FormatBorder::Dotted,
        "double" => FormatBorder::Double,
        "hair" => FormatBorder::Hair,
        _ => FormatBorder::Thin,
    }
}

// ============================================
// CELL VALUE CLASSIFICATION
// ============================================
// Shared by write / write_row / write_column so the Python-value-to-
// Rust-type dispatch logic exists in exactly one place instead of
// being copy-pasted four times.
//
// IMPORTANT: bool must be checked BEFORE f64/i64. Python's bool is a
// subclass of int, and PyO3's numeric extraction happily converts a
// bool into 1.0/0.0 if given the chance -- so checking numeric types
// first would silently turn True/False into numbers instead of real
// boolean cells (this was a real bug in the original ordering).
enum CellValue {
    Blank,
    Str(String),
    Num(f64),
    Bool(bool),
}

fn classify(value: &Bound<'_, PyAny>) -> PyResult<CellValue> {
    if value.is_none() {
        Ok(CellValue::Blank)
    } else if let Ok(b) = value.extract::<bool>() {
        // Must stay first, before both numeric checks below: Python's
        // bool is an int subclass and PyO3 will happily coerce it to
        // f64/i64 if given the chance, silently turning True/False into
        // numbers instead of real boolean cells.
        Ok(CellValue::Bool(b))
    } else if let Ok(f) = value.extract::<f64>() {
        // Numeric checks before the String check: strings don't have
        // __float__/__index__, so this never misclassifies a string --
        // it just avoids paying for a failed String-extraction attempt
        // on every number, which is the common case for numeric-heavy
        // data.
        Ok(CellValue::Num(f))
    } else if let Ok(i) = value.extract::<i64>() {
        Ok(CellValue::Num(i as f64))
    } else if let Ok(s) = value.extract::<String>() {
        Ok(CellValue::Str(s))
    } else {
        Ok(CellValue::Str(value.str()?.to_string()))
    }
}

// rust_xlsxwriter 0.75 exposes separate `write_x` (unformatted) and
// `write_x_with_format` methods rather than an `Option<&Format>`
// parameter, so we branch once here instead of at every call site.
fn write_value(
    sheet: &mut RustWorksheet,
    row: u32,
    col: u16,
    cv: &CellValue,
    fmt: Option<&RustFormat>,
) -> Result<(), rust_xlsxwriter::XlsxError> {
    match (cv, fmt) {
        (CellValue::Blank, Some(f)) => sheet.write_blank(row, col, f).map(|_| ()),
        (CellValue::Blank, None) => Ok(()), // nothing meaningful to write without a format
        (CellValue::Str(s), Some(f)) => sheet.write_string_with_format(row, col, s.as_str(), f).map(|_| ()),
        (CellValue::Str(s), None) => sheet.write_string(row, col, s.as_str()).map(|_| ()),
        (CellValue::Num(n), Some(f)) => sheet.write_number_with_format(row, col, *n, f).map(|_| ()),
        (CellValue::Num(n), None) => sheet.write_number(row, col, *n).map(|_| ()),
        (CellValue::Bool(b), Some(f)) => sheet.write_boolean_with_format(row, col, *b, f).map(|_| ()),
        (CellValue::Bool(b), None) => sheet.write_boolean(row, col, *b).map(|_| ()),
    }
}

fn merge_value(
    sheet: &mut RustWorksheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    cv: &CellValue,
    fmt: &RustFormat,
) -> Result<(), rust_xlsxwriter::XlsxError> {
    // merge_range() itself only accepts a plain &str -- confirmed this
    // is still true as of rust_xlsxwriter 0.95.0 (checked the actual
    // source directly; it isn't generic over IntoExcelData like
    // write_string/write_number/etc are), so a version bump alone
    // wouldn't fix this. The crate's own documentation demonstrates the
    // correct workaround: establish the merge (and its format) with an
    // empty string, then overwrite the anchor cell with the real typed
    // value via the appropriate write_x_with_format call. This
    // preserves numeric/boolean cell types across the merge, so e.g.
    // SUM() over a merged numeric range still works, instead of every
    // merged value silently becoming text.
    sheet.merge_range(first_row, first_col, last_row, last_col, "", fmt)?;
    write_value(sheet, first_row, first_col, cv, Some(fmt))
}

// ============================================
// ARROW ZERO-COPY DATAFRAME READING
// ============================================
// Reads a Polars/Pandas/PyArrow object into native arrow-rs RecordBatches
// via the Arrow PyCapsule Interface (__arrow_c_stream__), without
// extracting individual Python objects per cell. This is Phase 1: it
// supports the four core types (int64, float64, string, bool) that
// exercise the full read -> classify -> write path end to end; wider
// type coverage (unsigned ints, dates/timestamps, decimals) is tracked
// as follow-up work, not attempted here.

/// Pulls RecordBatches out of any object exposing `__arrow_c_stream__`
/// (pyarrow.Table, pandas.DataFrame 2.x+) or a `.to_arrow()` method
/// (polars.DataFrame, which returns a pyarrow object with the capsule
/// method).
fn record_batches_from_arrow(obj: &Bound<'_, PyAny>) -> PyResult<Vec<RecordBatch>> {
    let stream_source: Bound<'_, PyAny> = if obj.hasattr("__arrow_c_stream__")? {
        obj.clone()
    } else if obj.hasattr("to_arrow")? {
        obj.call_method0("to_arrow")?
    } else {
        return Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(
            "write_dataframe() expects an object exposing __arrow_c_stream__ \
             (pyarrow.Table, pandas.DataFrame) or a .to_arrow() method (polars.DataFrame)",
        ));
    };

    let capsule_obj = stream_source.call_method0("__arrow_c_stream__")?;
    let capsule = capsule_obj.downcast::<PyCapsule>().map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyTypeError, _>(
            "__arrow_c_stream__() did not return a PyCapsule",
        )
    })?;

    let raw_ptr = capsule.pointer() as *mut FFI_ArrowArrayStream;
    if raw_ptr.is_null() {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
            "__arrow_c_stream__() returned a null pointer",
        ));
    }

    // SAFETY: raw_ptr points at a valid, producer-allocated
    // ArrowArrayStream C struct for the lifetime of `capsule_obj` (which
    // we're still holding above). ptr::read takes a bitwise copy of the
    // struct -- including its function-pointer callbacks and
    // private_data pointer -- giving us an owned FFI_ArrowArrayStream
    // whose Drop impl will call the stream's release() callback exactly
    // once. Per the Arrow C Data Interface spec, once a consumer takes
    // ownership this way it MUST null out the *original* struct's
    // release field so the capsule's own destructor doesn't invoke
    // release() a second time on the same private_data (a double-free).
    let stream: FFI_ArrowArrayStream = unsafe {
        let s = std::ptr::read(raw_ptr);
        (*raw_ptr).release = None;
        s
    };

    let reader = ArrowArrayStreamReader::try_new(stream).map_err(to_pyerr)?;
    reader
        .collect::<std::result::Result<Vec<RecordBatch>, arrow::error::ArrowError>>()
        .map_err(to_pyerr)
}

/// A typed reference into one column of a RecordBatch, resolved once per
/// batch rather than re-checked on every cell.
enum ArrowColumn<'a> {
    Int64(&'a Int64Array),
    Float64(&'a Float64Array),
    Utf8(&'a StringArray),
    LargeUtf8(&'a LargeStringArray),
    Bool(&'a BooleanArray),
}

fn resolve_arrow_column(array: &dyn Array) -> PyResult<ArrowColumn<'_>> {
    // Each unwrap() below is guarded by having just matched the exact
    // ArrowDataType that guarantees the downcast succeeds -- arrow-rs's
    // own invariant is that an array's concrete type always agrees
    // with array.data_type(). expect() with a specific message instead
    // of unwrap() means if that invariant is ever violated (an arrow-rs
    // bug, or a version mismatch between the arrow crate version this
    // was compiled against and the one that produced the array), the
    // panic clearly names which type pairing broke rather than just
    // saying "called unwrap on a None value".
    match array.data_type() {
        ArrowDataType::Int64 => Ok(ArrowColumn::Int64(
            array
                .as_any()
                .downcast_ref::<Int64Array>()
                .expect("arrow array reported DataType::Int64 but isn't an Int64Array"),
        )),
        ArrowDataType::Float64 => Ok(ArrowColumn::Float64(
            array
                .as_any()
                .downcast_ref::<Float64Array>()
                .expect("arrow array reported DataType::Float64 but isn't a Float64Array"),
        )),
        ArrowDataType::Utf8 => Ok(ArrowColumn::Utf8(
            array
                .as_any()
                .downcast_ref::<StringArray>()
                .expect("arrow array reported DataType::Utf8 but isn't a StringArray"),
        )),
        ArrowDataType::LargeUtf8 => Ok(ArrowColumn::LargeUtf8(
            array
                .as_any()
                .downcast_ref::<LargeStringArray>()
                .expect("arrow array reported DataType::LargeUtf8 but isn't a LargeStringArray"),
        )),
        ArrowDataType::Boolean => Ok(ArrowColumn::Bool(
            array
                .as_any()
                .downcast_ref::<BooleanArray>()
                .expect("arrow array reported DataType::Boolean but isn't a BooleanArray"),
        )),
        other => Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(format!(
            "write_dataframe(): column type {other:?} isn't supported yet \
             (supported in this initial implementation: int64, float64, \
             string/utf8, large_utf8, bool)"
        ))),
    }
}

fn arrow_cell_value(col: &ArrowColumn<'_>, row: usize) -> CellValue {
    match col {
        ArrowColumn::Int64(a) => {
            if a.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Num(a.value(row) as f64)
            }
        }
        ArrowColumn::Float64(a) => {
            if a.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Num(a.value(row))
            }
        }
        ArrowColumn::Utf8(a) => {
            if a.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Str(a.value(row).to_string())
            }
        }
        ArrowColumn::LargeUtf8(a) => {
            if a.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Str(a.value(row).to_string())
            }
        }
        ArrowColumn::Bool(a) => {
            if a.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Bool(a.value(row))
            }
        }
    }
}


#[pyclass]
struct Format {
    inner: RustFormat,
}

impl Format {
    fn new() -> Self {
        Format {
            inner: RustFormat::new(),
        }
    }

    // Swaps the current format out for a fresh placeholder and applies
    // `op` to the *owned* original value, avoiding a clone on every
    // single setter call (RustFormat's builder methods consume `self`
    // by value, which is why this indirection exists at all).
    fn update(&mut self, op: impl FnOnce(RustFormat) -> RustFormat) {
        let current = std::mem::replace(&mut self.inner, RustFormat::new());
        self.inner = op(current);
    }
}

#[pymethods]
impl Format {
    #[new]
    fn py_new() -> Self {
        Format::new()
    }

    // All setters below return `PyRefMut<Self>` (the same object) so
    // Python code can chain calls, e.g.
    // `wb.add_format().set_bold().set_font_size(12)`. Every example
    // in this repo relies on this chaining, so the setters must
    // return self rather than None.
    fn set_bold(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_bold());
        slf
    }

    fn set_italic(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_italic());
        slf
    }

    fn set_underline(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_underline(rust_xlsxwriter::FormatUnderline::Single));
        slf
    }

    fn set_font_name<'a>(mut slf: PyRefMut<'a, Self>, name: &str) -> PyRefMut<'a, Self> {
        slf.update(|f| f.set_font_name(name));
        slf
    }

    fn set_font_size(mut slf: PyRefMut<'_, Self>, size: f64) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_font_size(size));
        slf
    }

    fn set_font_color<'a>(mut slf: PyRefMut<'a, Self>, color: &str) -> PyRefMut<'a, Self> {
        let c = parse_color(color);
        slf.update(|f| f.set_font_color(c));
        slf
    }

    fn set_background_color<'a>(mut slf: PyRefMut<'a, Self>, color: &str) -> PyRefMut<'a, Self> {
        let c = parse_color(color);
        slf.update(|f| f.set_background_color(c));
        slf
    }

    fn set_border<'a>(mut slf: PyRefMut<'a, Self>, border: &str) -> PyRefMut<'a, Self> {
        let style = parse_border(border);
        slf.update(|f| f.set_border(style));
        slf
    }

    fn set_border_color<'a>(mut slf: PyRefMut<'a, Self>, color: &str) -> PyRefMut<'a, Self> {
        let c = parse_color(color);
        slf.update(|f| {
            f.set_border_color(c)
                .set_border_top_color(c)
                .set_border_bottom_color(c)
                .set_border_left_color(c)
                .set_border_right_color(c)
        });
        slf
    }

    fn set_top_border<'a>(mut slf: PyRefMut<'a, Self>, border: &str) -> PyRefMut<'a, Self> {
        let style = parse_border(border);
        slf.update(|f| f.set_border_top(style));
        slf
    }

    fn set_bottom_border<'a>(mut slf: PyRefMut<'a, Self>, border: &str) -> PyRefMut<'a, Self> {
        let style = parse_border(border);
        slf.update(|f| f.set_border_bottom(style));
        slf
    }

    fn set_left_border<'a>(mut slf: PyRefMut<'a, Self>, border: &str) -> PyRefMut<'a, Self> {
        let style = parse_border(border);
        slf.update(|f| f.set_border_left(style));
        slf
    }

    fn set_right_border<'a>(mut slf: PyRefMut<'a, Self>, border: &str) -> PyRefMut<'a, Self> {
        let style = parse_border(border);
        slf.update(|f| f.set_border_right(style));
        slf
    }

    fn set_num_format<'a>(mut slf: PyRefMut<'a, Self>, format: &str) -> PyRefMut<'a, Self> {
        slf.update(|f| f.set_num_format(format));
        slf
    }

    fn set_align<'a>(mut slf: PyRefMut<'a, Self>, align: &str) -> PyRefMut<'a, Self> {
        let alignment = match align.to_lowercase().as_str() {
            "left" => FormatAlign::Left,
            "center" => FormatAlign::Center,
            "right" => FormatAlign::Right,
            "fill" => FormatAlign::Fill,
            "justify" => FormatAlign::Justify,
            "center_across" => FormatAlign::CenterAcross,
            "distributed" => FormatAlign::Distributed,
            _ => FormatAlign::Left,
        };
        slf.update(|f| f.set_align(alignment));
        slf
    }

    fn set_vertical_align<'a>(mut slf: PyRefMut<'a, Self>, align: &str) -> PyRefMut<'a, Self> {
        let alignment = match align.to_lowercase().as_str() {
            "top" => FormatAlign::Top,
            "vcenter" | "center" => FormatAlign::VerticalCenter,
            "bottom" => FormatAlign::Bottom,
            "vdistributed" | "distributed" => FormatAlign::VerticalDistributed,
            "vjustify" | "justify" => FormatAlign::VerticalJustify,
            _ => FormatAlign::VerticalCenter,
        };
        slf.update(|f| f.set_align(alignment));
        slf
    }

    fn set_text_wrap(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_text_wrap());
        slf
    }

    fn set_shrink(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_shrink());
        slf
    }

    fn set_rotation(mut slf: PyRefMut<'_, Self>, rotation: i16) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_rotation(rotation));
        slf
    }

    fn set_indent(mut slf: PyRefMut<'_, Self>, indent: u8) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_indent(indent));
        slf
    }

    fn set_pattern<'a>(mut slf: PyRefMut<'a, Self>, pattern: &str) -> PyRefMut<'a, Self> {
        let pat = match pattern.to_lowercase().as_str() {
            "none" => FormatPattern::None,
            "solid" => FormatPattern::Solid,
            "medium_gray" => FormatPattern::MediumGray,
            "dark_gray" => FormatPattern::DarkGray,
            "light_gray" => FormatPattern::LightGray,
            "dark_horizontal" => FormatPattern::DarkHorizontal,
            "dark_vertical" => FormatPattern::DarkVertical,
            "dark_down" => FormatPattern::DarkDown,
            "dark_up" => FormatPattern::DarkUp,
            "dark_grid" => FormatPattern::DarkGrid,
            "dark_trellis" => FormatPattern::DarkTrellis,
            "light_horizontal" => FormatPattern::LightHorizontal,
            "light_vertical" => FormatPattern::LightVertical,
            "light_down" => FormatPattern::LightDown,
            "light_up" => FormatPattern::LightUp,
            "light_grid" => FormatPattern::LightGrid,
            "light_trellis" => FormatPattern::LightTrellis,
            "gray_125" => FormatPattern::Gray125,
            "gray_0625" => FormatPattern::Gray0625,
            _ => FormatPattern::None,
        };
        slf.update(|f| f.set_pattern(pat));
        slf
    }
}

// ============================================
// WORKSHEET CLASS
// ============================================
// Worksheet does NOT own a RustWorksheet. Owning a clone was the
// critical bug in the previous version: rust_xlsxwriter's
// `Workbook::add_worksheet()` hands back a `&mut Worksheet` that lives
// *inside* the workbook's own storage. Cloning it detaches the copy
// Python holds from the one the workbook will actually serialize, so
// every write made through Python silently vanished on save.
//
// Instead, each Worksheet just remembers its own index and borrows the
// worksheet fresh from the workbook (via `worksheet_from_index`, which
// rust_xlsxwriter provides for exactly this "held across calls" use
// case) every time a method is called.
#[pyclass]
struct Worksheet {
    workbook: Py<Workbook>,
    index: usize,
    // Both only meaningful when constant_memory is true. rust_xlsxwriter's
    // own constant-memory mode requires rows to be written in
    // non-decreasing order, but -- confirmed directly from its source
    // (check_dimensions/store_string in worksheet.rs) -- does NOT raise an
    // error when that's violated. It silently proceeds, which produces a
    // corrupt or incomplete .xlsx with no signal to the caller. This
    // binding layer enforces the restriction itself instead of leaving
    // callers exposed to that.
    constant_memory: bool,
    min_allowed_row: Cell<u32>,
}

impl Worksheet {
    fn with_sheet<F, R>(&self, py: Python<'_>, f: F) -> PyResult<R>
    where
        F: FnOnce(&mut RustWorksheet) -> Result<R, rust_xlsxwriter::XlsxError>,
    {
        let wb_ref = self.workbook.borrow(py);
        let mut wb = wb_ref.inner.borrow_mut();
        let sheet = wb.worksheet_from_index(self.index).map_err(xlsx_err_to_pyerr)?;
        f(sheet).map_err(xlsx_err_to_pyerr)
    }

    // Call before writing to `row`. No-op unless constant_memory is set.
    // Rejects a write to any row before the highest row already written,
    // with a clear Python exception, instead of letting it through to
    // rust_xlsxwriter's silent-corruption path. Advances the tracked
    // high-water mark to `row` on success.
    fn check_row_order(&self, row: u32) -> PyResult<()> {
        if self.constant_memory {
            let min_row = self.min_allowed_row.get();
            if row < min_row {
                return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                    "Cannot write to row {row}: this worksheet was created with \
                     constant_memory=True, which requires rows to be written in \
                     non-decreasing order, and row {min_row} has already been \
                     written. (rust_xlsxwriter itself does not raise an error for \
                     this -- it would silently produce a corrupt or incomplete \
                     .xlsx file instead -- so this check exists specifically to \
                     catch it here.)"
                )));
            }
            self.min_allowed_row.set(row);
        }
        Ok(())
    }

    // Same as check_row_order(), but for a call that's already known (by
    // the caller) to touch rows up through `last_row_touched` in one go
    // (write_column, merge_range, write_records, write_dataframe) --
    // validates the starting row, then advances the high-water mark to
    // the last row actually written instead of just the first, so a
    // later out-of-order call is still caught correctly.
    fn check_row_order_range(&self, first_row: u32, last_row_touched: u32) -> PyResult<()> {
        self.check_row_order(first_row)?;
        if self.constant_memory {
            self.min_allowed_row.set(last_row_touched);
        }
        Ok(())
    }
}

#[pymethods]
impl Worksheet {
    #[pyo3(signature = (row, col, value, format=None))]
    fn write(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        value: &Bound<'_, PyAny>,
        format: Option<&Format>,
    ) -> PyResult<()> {
        let cv = classify(value)?;
        self.check_row_order(row)?;
        let fmt = format.map(|f| &f.inner);
        self.with_sheet(py, |sheet| write_value(sheet, row, col, &cv, fmt))
    }

    #[pyo3(signature = (row, col, values, format=None))]
    fn write_row(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        values: &Bound<'_, PyList>,
        format: Option<&Format>,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        let fmt = format.map(|f| &f.inner);
        let mut classified = Vec::with_capacity(values.len());
        for v in values.iter() {
            classified.push(classify(&v)?);
        }
        self.with_sheet(py, |sheet| {
            for (i, cv) in classified.iter().enumerate() {
                write_value(sheet, row, col + i as u16, cv, fmt)?;
            }
            Ok(())
        })
    }

    #[pyo3(signature = (row, col, values, format=None))]
    fn write_column(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        values: &Bound<'_, PyList>,
        format: Option<&Format>,
    ) -> PyResult<()> {
        let fmt = format.map(|f| &f.inner);
        let mut classified = Vec::with_capacity(values.len());
        for v in values.iter() {
            classified.push(classify(&v)?);
        }
        if !classified.is_empty() {
            let last_row = row + (classified.len() as u32 - 1);
            self.check_row_order_range(row, last_row)?;
        }
        self.with_sheet(py, |sheet| {
            for (i, cv) in classified.iter().enumerate() {
                write_value(sheet, row + i as u32, col, cv, fmt)?;
            }
            Ok(())
        })
    }

    // Writes an entire list-of-dicts dataset in a single Python->Rust
    // call, instead of one call per row or per cell. This exists
    // specifically to remove the FFI-crossing bottleneck: benchmarking
    // showed per-cell write() making 500,000 individual PyO3 calls for
    // a 100k-row x 5-col sheet, where write_records() makes exactly
    // one. All Python object access here (dict lookups, key
    // extraction) happens natively against the passed-in PyList/PyDict
    // without re-entering Python bytecode, so cost scales with data
    // size, not with how many times Python and Rust hand control back
    // and forth.
    //
    // `headers` controls column order and which keys are pulled from
    // each record; if omitted, it's taken from the first record's
    // keys (insertion order, matching Python dict semantics).
    #[pyo3(signature = (start_row, start_col, records, headers=None, format=None, header_format=None, write_header=true))]
    #[allow(clippy::too_many_arguments)]
    fn write_records(
        &self,
        py: Python<'_>,
        start_row: u32,
        start_col: u16,
        records: &Bound<'_, PyList>,
        headers: Option<Vec<String>>,
        format: Option<&Format>,
        header_format: Option<&Format>,
        write_header: bool,
    ) -> PyResult<()> {
        if records.is_empty() {
            return Ok(());
        }

        // write_records() always writes rows sequentially starting at
        // start_row (optionally +1 for the header row) through
        // start_row + records.len() (+1) - 1, so the whole range can be
        // validated in one check up front instead of once per row --
        // consistent with write_records() being the bulk/low-per-call-
        // overhead path.
        let header_rows = if write_header { 1 } else { 0 };
        let last_row = start_row + header_rows + records.len() as u32 - 1;
        self.check_row_order_range(start_row, last_row)?;

        let headers: Vec<String> = match headers {
            Some(h) => h,
            None => {
                let first = records.get_item(0)?;
                let first_dict = first.downcast::<PyDict>().map_err(|_| {
                    PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                        "write_records() expects a list of dicts, or an explicit `headers` list",
                    )
                })?;
                first_dict
                    .keys()
                    .iter()
                    .map(|k| k.extract::<String>())
                    .collect::<PyResult<Vec<String>>>()?
            }
        };

        // Single pass: borrow the worksheet once, then read each dict
        // value and write it immediately. This avoids materializing a
        // full Vec<Vec<CellValue>> copy of the dataset before writing
        // (100k+ extra heap allocations for a 100k-row sheet) -- read
        // and write are interleaved instead.
        let data_fmt = format.map(|f| &f.inner);
        let head_fmt = header_format.map(|f| &f.inner);

        let wb_ref = self.workbook.borrow(py);
        let mut wb = wb_ref.inner.borrow_mut();
        let sheet = wb.worksheet_from_index(self.index).map_err(xlsx_err_to_pyerr)?;

        // dict.get_item(key) converts `key` to a Python object on
        // every call (K: ToPyObject) -- passing a &str/&String there
        // allocates a brand new PyUnicode object per lookup, which
        // adds up to 500,000 allocations for a 100k-row x 5-col sheet.
        // Building each header's Python string object once up front
        // and reusing it removes that per-cell allocation.
        let header_objs: Vec<Bound<'_, PyString>> =
            headers.iter().map(|h| PyString::new_bound(py, h)).collect();

        let mut row_cursor = start_row;
        if write_header {
            for (i, h) in headers.iter().enumerate() {
                write_value(
                    sheet,
                    row_cursor,
                    start_col + i as u16,
                    &CellValue::Str(h.clone()),
                    head_fmt,
                )
                .map_err(xlsx_err_to_pyerr)?;
            }
            row_cursor += 1;
        }

        for record in records.iter() {
            let dict = record.downcast::<PyDict>().map_err(|_| {
                PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                    "write_records() expects every record to be a dict",
                )
            })?;

            // NOTE: a "read via dict.values() when this record's key
            // order matches headers" fast path was tried here and
            // reverted. A length-only check silently misaligned columns
            // when a same-length record had its keys in a different
            // order (verified this really happens); fixing that
            // required comparing actual key order too (via dict.iter(),
            // since dict.keys()/dict.values() each materialize a new
            // PyList per call and cost more than they saved). Once
            // correct, it measured as a wash against the get_item path
            // below (0.98-0.99x, i.e. not faster) -- the safety check
            // costs about what the hash lookups it would have replaced
            // cost. Not worth the added code complexity for no
            // measured benefit.
            for (i, key_obj) in header_objs.iter().enumerate() {
                let value = dict.get_item(key_obj)?;
                let cv = match value {
                    Some(v) => classify(&v)?,
                    None => CellValue::Blank,
                };
                write_value(sheet, row_cursor, start_col + i as u16, &cv, data_fmt)
                    .map_err(xlsx_err_to_pyerr)?;
            }
            row_cursor += 1;
        }

        Ok(())
    }

    // Zero-copy bulk write from a Polars/Pandas/PyArrow object, via the
    // Arrow PyCapsule Interface. Unlike write_records(), which still
    // does a PyO3 extract() per cell, this reads directly from Arrow's
    // native columnar buffers -- no Python object is touched once the
    // initial __arrow_c_stream__() call hands over the data. Phase 1:
    // supports int64/float64/string/bool columns; see the README's
    // Roadmap for wider type coverage.
    #[pyo3(signature = (start_row, start_col, data, header_format=None, write_header=true))]
    fn write_dataframe(
        &self,
        py: Python<'_>,
        start_row: u32,
        start_col: u16,
        data: &Bound<'_, PyAny>,
        header_format: Option<&Format>,
        write_header: bool,
    ) -> PyResult<()> {
        let batches = record_batches_from_arrow(data)?;
        if batches.is_empty() {
            return Ok(());
        }

        let schema = batches[0].schema();
        let field_names: Vec<String> = schema.fields().iter().map(|f| f.name().clone()).collect();

        // Check every column's type up front so we fail with one clear
        // error instead of partway through a partially-written sheet.
        for field in schema.fields() {
            match field.data_type() {
                ArrowDataType::Int64
                | ArrowDataType::Float64
                | ArrowDataType::Utf8
                | ArrowDataType::LargeUtf8
                | ArrowDataType::Boolean => {}
                other => {
                    return Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(format!(
                        "write_dataframe(): column '{}' has type {other:?}, \
                         which isn't supported yet (supported: int64, \
                         float64, string/utf8, large_utf8, bool)",
                        field.name()
                    )));
                }
            }
        }

        let head_fmt = header_format.map(|f| &f.inner);

        // Same one-check-for-the-whole-call approach as write_records():
        // total rows this call will touch is known upfront (sum of each
        // batch's row count, each an O(1) lookup), so there's no need to
        // check per-row.
        let total_data_rows: u32 = batches.iter().map(|b| b.num_rows() as u32).sum();
        let header_rows = if write_header { 1 } else { 0 };
        if total_data_rows + header_rows > 0 {
            let last_row = start_row + header_rows + total_data_rows - 1;
            self.check_row_order_range(start_row, last_row)?;
        }

        let wb_ref = self.workbook.borrow(py);
        let mut wb = wb_ref.inner.borrow_mut();
        let sheet = wb.worksheet_from_index(self.index).map_err(xlsx_err_to_pyerr)?;

        let mut row_cursor = start_row;
        if write_header {
            for (i, name) in field_names.iter().enumerate() {
                write_value(
                    sheet,
                    row_cursor,
                    start_col + i as u16,
                    &CellValue::Str(name.clone()),
                    head_fmt,
                )
                .map_err(xlsx_err_to_pyerr)?;
            }
            row_cursor += 1;
        }

        for batch in &batches {
            let columns: Vec<ArrowColumn<'_>> = (0..batch.num_columns())
                .map(|c| resolve_arrow_column(batch.column(c).as_ref()))
                .collect::<PyResult<Vec<_>>>()?;

            for r in 0..batch.num_rows() {
                for (c, col) in columns.iter().enumerate() {
                    let cv = arrow_cell_value(col, r);
                    write_value(sheet, row_cursor, start_col + c as u16, &cv, None)
                        .map_err(xlsx_err_to_pyerr)?;
                }
                row_cursor += 1;
            }
        }

        Ok(())
    }

    #[pyo3(signature = (first_row, first_col, last_row, last_col, value, format=None))]
    fn merge_range(
        &self,
        py: Python<'_>,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        value: &Bound<'_, PyAny>,
        format: Option<&Format>,
    ) -> PyResult<()> {
        let cv = classify(value)?;
        self.check_row_order_range(first_row, last_row)?;
        let default_fmt = RustFormat::new();
        let fmt: &RustFormat = format.map(|f| &f.inner).unwrap_or(&default_fmt);
        self.with_sheet(py, |sheet| {
            merge_value(sheet, first_row, first_col, last_row, last_col, &cv, fmt)
        })
    }

    fn set_column_width(&self, py: Python<'_>, col: u16, width: f64) -> PyResult<()> {
        self.with_sheet(py, |sheet| sheet.set_column_width(col, width).map(|_| ()))
    }

    fn set_row_height(&self, py: Python<'_>, row: u32, height: f64) -> PyResult<()> {
        self.with_sheet(py, |sheet| sheet.set_row_height(row, height).map(|_| ()))
    }

    fn freeze_panes(&self, py: Python<'_>, row: u32, col: u16) -> PyResult<()> {
        self.with_sheet(py, |sheet| sheet.set_freeze_panes(row, col).map(|_| ()))
    }

    fn set_tab_color(&self, py: Python<'_>, color: &str) -> PyResult<()> {
        let c = parse_color(color);
        self.with_sheet(py, |sheet| {
            sheet.set_tab_color(c);
            Ok(())
        })
    }

    fn hide(&self, py: Python<'_>) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_hidden(true);
            Ok(())
        })
    }

    #[pyo3(signature = (password=None))]
    fn protect(&self, py: Python<'_>, password: Option<&str>) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            match password {
                Some(p) => {
                    sheet.protect_with_password(p);
                }
                None => {
                    sheet.protect();
                }
            }
            Ok(())
        })
    }

    #[pyo3(signature = (row, col, formula, format=None))]
    fn write_formula(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        formula: &str,
        format: Option<&Format>,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        let fmt = format.map(|f| &f.inner);
        self.with_sheet(py, |sheet| match &fmt {
            Some(f) => sheet.write_formula_with_format(row, col, formula, f).map(|_| ()),
            None => sheet.write_formula(row, col, formula).map(|_| ()),
        })
    }

    #[pyo3(signature = (row, col, url, format=None))]
    fn write_url(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        url: &str,
        format: Option<&Format>,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        let fmt = format.map(|f| &f.inner);
        self.with_sheet(py, |sheet| match &fmt {
            Some(f) => sheet.write_url_with_format(row, col, url, f).map(|_| ()),
            None => sheet.write_url(row, col, url).map(|_| ()),
        })
    }

    #[pyo3(signature = (row, col, year, month, day, hour, min, sec, format=None))]
    #[allow(clippy::too_many_arguments)]
    fn write_datetime(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        year: u16,
        month: u8,
        day: u8,
        hour: u16,
        min: u8,
        sec: f64,
        format: Option<&Format>,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        let fmt = format.map(|f| &f.inner);
        let dt = rust_xlsxwriter::ExcelDateTime::from_ymd(year, month, day)
            .map_err(xlsx_err_to_pyerr)?
            .and_hms(hour, min, sec)
            .map_err(xlsx_err_to_pyerr)?;
        self.with_sheet(py, |sheet| match &fmt {
            Some(f) => sheet.write_datetime_with_format(row, col, &dt, f).map(|_| ()),
            None => sheet.write_datetime(row, col, &dt).map(|_| ()),
        })
    }

    #[pyo3(signature = (row, col, year, month, day, format=None))]
    fn write_date(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        year: u16,
        month: u8,
        day: u8,
        format: Option<&Format>,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        let fmt = format.map(|f| &f.inner);
        let dt = rust_xlsxwriter::ExcelDateTime::from_ymd(year, month, day).map_err(xlsx_err_to_pyerr)?;
        self.with_sheet(py, |sheet| match &fmt {
            Some(f) => sheet.write_datetime_with_format(row, col, &dt, f).map(|_| ()),
            None => sheet.write_datetime(row, col, &dt).map(|_| ()),
        })
    }

    fn insert_image(&self, py: Python<'_>, row: u32, col: u16, image_path: &str) -> PyResult<()> {
        self.check_row_order(row)?;
        let image = rust_xlsxwriter::Image::new(image_path).map_err(xlsx_err_to_pyerr)?;
        self.with_sheet(py, |sheet| sheet.insert_image(row, col, &image).map(|_| ()))
    }

    fn autofit(&self, py: Python<'_>) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.autofit();
            Ok(())
        })
    }

    fn set_name(&self, py: Python<'_>, name: &str) -> PyResult<()> {
        self.with_sheet(py, |sheet| sheet.set_name(name).map(|_| ()))
    }

    // Adds Excel's autofilter dropdown controls to the header row of a
    // range (first_row is typically the header row; data rows follow
    // below it). Does not itself hide/filter any rows -- that's a
    // display-time Excel feature, not something this writes into the
    // file -- it just adds the filter UI and the range it applies to.
    fn autofilter(
        &self,
        py: Python<'_>,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
    ) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet
                .autofilter(first_row, first_col, last_row, last_col)
                .map(|_| ())
        })
    }
}

// ============================================
// WORKBOOK CLASS
// ============================================
#[pyclass]
struct Workbook {
    inner: RefCell<RustWorkbook>,
    // Tracks how many worksheets have been added so each Worksheet
    // handle knows its own stable index into the workbook's storage.
    sheet_count: RefCell<usize>,
}

#[pymethods]
impl Workbook {
    #[new]
    fn new() -> Self {
        Workbook {
            inner: RefCell::new(RustWorkbook::new()),
            sheet_count: RefCell::new(0),
        }
    }

    // constant_memory=True routes to rust_xlsxwriter's
    // add_worksheet_with_constant_memory() instead of add_worksheet(),
    // streaming each row to a temp file as it's written instead of
    // buffering the whole sheet in memory. It comes with a hard
    // restriction, confirmed directly from rust_xlsxwriter 0.96's own
    // source (src/performance.rs's "Restrictions when using constant
    // memory mode" section): rows must be written in non-decreasing
    // order -- once you've written to row n, you can no longer write to
    // any row < n.
    //
    // IMPORTANT: violating that restriction does NOT raise a clean
    // error. Checked this directly too (worksheet.rs's check_dimensions
    // and store_string): the row-order check only affects internal
    // dimension-tracking bookkeeping, not whether the write is allowed
    // to proceed -- an out-of-order write on a constant_memory
    // worksheet silently continues rather than returning an XlsxError.
    // The underlying crate does not protect you here; the caller must
    // guarantee monotonic row order.
    //
    // write_records() and write_dataframe() both already write
    // strictly row-by-row in increasing order (see their
    // implementations below), so they're safe to use with
    // constant_memory=True. Plain write()/write_row()/write_column()/
    // merge_range() calls are NOT restricted from being called out of
    // order by anything in this binding layer -- doing so on a
    // constant_memory worksheet is the caller's responsibility to avoid.
    #[pyo3(signature = (name=None, constant_memory=false))]
    fn add_worksheet(
        slf: Py<Self>,
        py: Python<'_>,
        name: Option<&str>,
        constant_memory: bool,
    ) -> PyResult<Py<Worksheet>> {
        let index = {
            let wb_ref = slf.borrow(py);
            let mut wb = wb_ref.inner.borrow_mut();
            let sheet = if constant_memory {
                wb.add_worksheet_with_constant_memory()
            } else {
                wb.add_worksheet()
            };
            if let Some(n) = name {
                sheet.set_name(n).map_err(xlsx_err_to_pyerr)?;
            }
            let mut count = wb_ref.sheet_count.borrow_mut();
            let idx = *count;
            *count += 1;
            idx
        };
        Py::new(
            py,
            Worksheet {
                workbook: slf.clone_ref(py),
                index,
                constant_memory,
                min_allowed_row: Cell::new(0),
            },
        )
    }

    fn add_format(&self, py: Python<'_>) -> PyResult<Py<Format>> {
        Py::new(py, Format::new())
    }

    fn close(&self, path: &str) -> PyResult<()> {
        self.inner.borrow_mut().save(path).map_err(xlsx_err_to_pyerr)?;
        Ok(())
    }

    // Defines a named range/formula usable in Excel formulas and the
    // Name Box. `name` can be a plain name for a workbook-global
    // definition ("MyRange"), or "SheetName!Name" for a name scoped to
    // one worksheet (matches rust_xlsxwriter's own convention for
    // distinguishing the two -- see its define_name() docs). `formula`
    // is the range/formula the name refers to, e.g. "Sheet1!$A$1:$A$10".
    fn define_name(&self, name: &str, formula: &str) -> PyResult<()> {
        self.inner
            .borrow_mut()
            .define_name(name, formula)
            .map_err(xlsx_err_to_pyerr)?;
        Ok(())
    }

    #[pyo3(signature = (title=None, author=None, subject=None, keywords=None, comments=None))]
    fn set_properties(
        &self,
        title: Option<&str>,
        author: Option<&str>,
        subject: Option<&str>,
        keywords: Option<&str>,
        comments: Option<&str>,
    ) {
        let mut wb = self.inner.borrow_mut();
        let mut props = rust_xlsxwriter::DocProperties::new();
        if let Some(t) = title {
            props = props.set_title(t);
        }
        if let Some(a) = author {
            props = props.set_author(a);
        }
        if let Some(s) = subject {
            props = props.set_subject(s);
        }
        if let Some(k) = keywords {
            props = props.set_keywords(k);
        }
        if let Some(c) = comments {
            props = props.set_comment(c);
        }
        wb.set_properties(&props);
    }
}

// ============================================
// MODULE INITIALIZATION
// ============================================
#[pymodule]
fn _core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Workbook>()?;
    m.add_class::<Worksheet>()?;
    m.add_class::<Format>()?;
    Ok(())
}
