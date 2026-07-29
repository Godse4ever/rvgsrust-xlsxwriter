use arrow::array::{
    Array, BooleanArray, Date32Array, Date64Array, Float32Array, Float64Array, Int16Array,
    Int32Array, Int64Array, Int8Array, LargeStringArray, RecordBatch, StringArray, StringViewArray,
    TimestampMicrosecondArray, TimestampMillisecondArray, TimestampNanosecondArray,
    TimestampSecondArray, UInt16Array, UInt32Array, UInt64Array, UInt8Array,
};
use arrow::datatypes::{DataType as ArrowDataType, TimeUnit};
use arrow::ffi_stream::{ArrowArrayStreamReader, FFI_ArrowArrayStream};
use pyo3::prelude::*;
use pyo3::types::{PyCapsule, PyDict, PyDictMethods, PyList, PyListMethods, PyString};
use pyo3::PyRefMut;
use rust_xlsxwriter::{
    Color, ExcelDateTime, Format as RustFormat, FormatAlign, FormatBorder, FormatPattern,
    Table as RustTable, TableColumn as RustTableColumn, TableFunction, TableStyle,
    Workbook as RustWorkbook, Worksheet as RustWorksheet,
};
use std::borrow::Cow;
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

// Every path that mutates the workbook goes through
// RefCell::borrow_mut(). That panics on a double borrow, and a double
// borrow is genuinely reachable from Python rather than being an
// internal invariant: the bulk write paths hold the borrow across
// per-cell classify() calls, and classify() falls back to value.str()
// for unrecognised types, which runs arbitrary Python. A __str__ (or a
// __del__ triggered mid-loop) that touches the same Workbook re-enters
// borrow_mut() while the first borrow is still live.
//
// A panic there crosses the FFI boundary as pyo3's PanicException,
// which is catchable but reports only "already mutably borrowed:
// BorrowMutError" -- nothing about what the caller did or how to avoid
// it. try_borrow_mut() plus this message turns it into an ordinary,
// explicable Python exception instead.
fn reentrant_workbook_err() -> PyErr {
    PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(
        "Workbook is already being modified by a write in progress on this \
         thread. This usually means a value's __str__/__repr__ (or a \
         __del__) called back into the same Workbook while a bulk write \
         such as write_records() was running. Convert such values to str \
         before passing them in, and do not re-enter the Workbook from a \
         conversion method.",
    )
}

// ============================================
// COLOR HELPER
// ============================================
fn parse_color(color: &str) -> PyResult<Color> {
    if color.starts_with('#') && color.len() == 7 {
        if let (Ok(r), Ok(g), Ok(b)) = (
            u8::from_str_radix(&color[1..3], 16),
            u8::from_str_radix(&color[3..5], 16),
            u8::from_str_radix(&color[5..7], 16),
        ) {
            let rgb = ((r as u32) << 16) | ((g as u32) << 8) | (b as u32);
            return Ok(Color::RGB(rgb));
        }
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
            "Invalid hex color '{color}': expected '#RRGGBB' with valid hex digits"
        )));
    }
    match color.to_lowercase().as_str() {
        "black" => Ok(Color::Black),
        "blue" => Ok(Color::Blue),
        "brown" => Ok(Color::Brown),
        "cyan" => Ok(Color::Cyan),
        "gray" | "grey" => Ok(Color::Gray),
        "green" => Ok(Color::Green),
        "lime" => Ok(Color::Lime),
        "magenta" => Ok(Color::Magenta),
        "navy" => Ok(Color::Navy),
        "orange" => Ok(Color::Orange),
        "pink" => Ok(Color::Pink),
        "purple" => Ok(Color::Purple),
        "red" => Ok(Color::Red),
        "silver" => Ok(Color::Silver),
        "white" => Ok(Color::White),
        "yellow" => Ok(Color::Yellow),
        _ => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
            "Unknown color '{color}'. Expected a '#RRGGBB' hex string or one of:              black, blue, brown, cyan, gray/grey, green, lime, magenta, navy,              orange, pink, purple, red, silver, white, yellow"
        ))),
    }
}

fn parse_diagonal_border(name: &str) -> PyResult<rust_xlsxwriter::FormatDiagonalBorder> {
    use rust_xlsxwriter::FormatDiagonalBorder as D;
    match name.to_ascii_lowercase().as_str() {
        "none" => Ok(D::None),
        "up" => Ok(D::BorderUp),
        "down" => Ok(D::BorderDown),
        "up_down" => Ok(D::BorderUpDown),
        other => Err(cf_type_err(
            "diagonal border type",
            other,
            "none, up, down, up_down",
        )),
    }
}

fn parse_border(border: &str) -> PyResult<FormatBorder> {
    match border.to_lowercase().as_str() {
        "thin" => Ok(FormatBorder::Thin),
        "medium" => Ok(FormatBorder::Medium),
        "thick" => Ok(FormatBorder::Thick),
        "dashed" => Ok(FormatBorder::Dashed),
        "dotted" => Ok(FormatBorder::Dotted),
        "double" => Ok(FormatBorder::Double),
        "hair" => Ok(FormatBorder::Hair),
        _ => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
            "Unknown border style '{border}'. Expected one of:              thin, medium, thick, dashed, dotted, double, hair"
        ))),
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
enum CellValue<'a> {
    Blank,
    // Cow rather than String so the Arrow path can borrow straight out of
    // the columnar string buffer instead of heap-allocating a String per
    // cell that is dropped again immediately after write_string(). The
    // Python paths still produce Cow::Owned, since the PyString they read
    // from is a temporary that does not outlive the classify() call.
    Str(Cow<'a, str>),
    Num(f64),
    Bool(bool),
    // Excel serial date/datetime -- days since 1899-12-30, fractional
    // part being time of day. Deliberately NOT folded into Num: Excel
    // stores dates as plain f64 and decides how to display them purely
    // from the cell's number format, so a serial written as a bare
    // number renders as "45123" rather than "2023-07-14". Keeping these
    // as distinct variants lets write_value() attach a date format.
    // Date is whole-day (yyyy-mm-dd); DateTime carries a time component.
    Date(f64),
    DateTime(f64),
}

// Excel's day-zero is 1899-12-30; the Unix epoch (1970-01-01) is serial
// 25569. Arrow stores all its date/time types as an offset from the Unix
// epoch, so every conversion below is (value / units_per_day) + 25569.
const EXCEL_UNIX_EPOCH_DAYS: f64 = 25569.0;

// Single source of truth for the supported-type list, so the two error
// paths (per-column schema check and resolve_arrow_column's fallback)
// can't drift apart as coverage grows.
const SUPPORTED_ARROW_TYPES: &str = "supported: int8/16/32/64, uint8/16/32/64, \
     float32/64, string/utf8, large_utf8, utf8view, bool, date32, date64, \
     timestamp[s|ms|us|ns]";

fn classify(value: &Bound<'_, PyAny>) -> PyResult<CellValue<'static>> {
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
        Ok(CellValue::Str(Cow::Owned(s)))
    } else {
        Ok(CellValue::Str(Cow::Owned(value.str()?.to_string())))
    }
}

// rust_xlsxwriter 0.75 exposes separate `write_x` (unformatted) and
// `write_x_with_format` methods rather than an `Option<&Format>`
// parameter, so we branch once here instead of at every call site.
fn write_value(
    sheet: &mut RustWorksheet,
    row: u32,
    col: u16,
    cv: &CellValue<'_>,
    fmt: Option<&RustFormat>,
) -> Result<(), rust_xlsxwriter::XlsxError> {
    match (cv, fmt) {
        (CellValue::Blank, Some(f)) => sheet.write_blank(row, col, f).map(|_| ()),
        (CellValue::Blank, None) => Ok(()), // nothing meaningful to write without a format
        (CellValue::Str(s), Some(f)) => sheet
            .write_string_with_format(row, col, s.as_ref(), f)
            .map(|_| ()),
        (CellValue::Str(s), None) => sheet.write_string(row, col, s.as_ref()).map(|_| ()),
        (CellValue::Num(n), Some(f)) => sheet.write_number_with_format(row, col, *n, f).map(|_| ()),
        (CellValue::Num(n), None) => sheet.write_number(row, col, *n).map(|_| ()),
        (CellValue::Bool(b), Some(f)) => {
            sheet.write_boolean_with_format(row, col, *b, f).map(|_| ())
        }
        (CellValue::Bool(b), None) => sheet.write_boolean(row, col, *b).map(|_| ()),
        // from_serial_datetime() range-checks against 0.0..2_958_466.0
        // (Excel years 1900-9999) and returns Err outside it, so dates
        // Excel physically cannot represent surface as a Python
        // exception rather than being silently clamped or wrapped.
        (CellValue::Date(n) | CellValue::DateTime(n), Some(f)) => {
            let edt = ExcelDateTime::from_serial_datetime(*n)?;
            sheet.write_datetime_with_format(row, col, &edt, f)?;
            Ok(())
        }
        // No format supplied: still written as a datetime cell type, but
        // Excel will render the bare serial. Callers inside this crate
        // always pass a format for these variants; this arm exists for
        // exhaustiveness.
        (CellValue::Date(n) | CellValue::DateTime(n), None) => {
            let edt = ExcelDateTime::from_serial_datetime(*n)?;
            sheet.write_datetime(row, col, &edt)?;
            Ok(())
        }
    }
}

fn merge_value(
    sheet: &mut RustWorksheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    cv: &CellValue<'_>,
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
// extracting individual Python objects per cell. Covers the integer and
// float widths, the three string encodings, bool, and the date/timestamp
// types; decimal, list, struct and dictionary-encoded columns are still
// tracked as follow-up work.

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

    // Validate the capsule name per the Arrow C Data Interface spec.
    // The spec mandates the name "arrow_array_stream" for stream capsules.
    // Accepting an unnamed or differently-named capsule from a buggy/malicious
    // object could hand us a pointer to an unrelated struct, causing UB in the
    // unsafe block below. PyCapsule::name() returns None for unnamed capsules.
    let capsule_name = capsule.name().ok().flatten();
    let expected = c"arrow_array_stream";
    if capsule_name != Some(expected) {
        return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
            "__arrow_c_stream__() returned a PyCapsule with unexpected name {:?};              expected \"arrow_array_stream\"",
            capsule_name.map(|c| c.to_string_lossy().into_owned())
        )));
    }

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
    Int8(&'a Int8Array),
    Int16(&'a Int16Array),
    Int32(&'a Int32Array),
    Int64(&'a Int64Array),
    UInt8(&'a UInt8Array),
    UInt16(&'a UInt16Array),
    UInt32(&'a UInt32Array),
    UInt64(&'a UInt64Array),
    Float32(&'a Float32Array),
    Float64(&'a Float64Array),
    Utf8(&'a StringArray),
    LargeUtf8(&'a LargeStringArray),
    // Utf8View = default string type in Polars >= 1.0 (Arrow StringView)
    Utf8View(&'a StringViewArray),
    Bool(&'a BooleanArray),
    // Date32 = days since Unix epoch, Date64 = milliseconds since Unix
    // epoch (Arrow specifies Date64 values should be exact multiples of
    // a whole day). Both render as date-only.
    Date32(&'a Date32Array),
    Date64(&'a Date64Array),
    // Timestamp[ns] is what pandas' default datetime64[ns] dtype maps to,
    // so it's the common case rather than an exotic one.
    TimestampSecond(&'a TimestampSecondArray),
    TimestampMillisecond(&'a TimestampMillisecondArray),
    TimestampMicrosecond(&'a TimestampMicrosecondArray),
    TimestampNanosecond(&'a TimestampNanosecondArray),
}

// Every downcast below is guarded by having just matched the exact
// ArrowDataType that guarantees it succeeds. Factored into a macro
// because doing it longhand across 20 variants is where a copy-paste
// type mismatch would hide.
macro_rules! arrow_col {
    ($array:expr, $variant:ident, $arr_ty:ty) => {{
        let arr = $array.as_any().downcast_ref::<$arr_ty>();
        let arr = arr.expect(concat!("arrow array is not a ", stringify!($arr_ty)));
        Ok(ArrowColumn::$variant(arr))
    }};
}

fn resolve_arrow_column(array: &dyn Array) -> PyResult<ArrowColumn<'_>> {
    // arrow-rs's invariant is that an array's concrete type always agrees
    // with array.data_type(). The expect() inside arrow_col! (rather than
    // unwrap()) means that if that invariant is ever violated -- an
    // arrow-rs bug, or a version mismatch between the arrow crate this was
    // compiled against and the one that produced the array -- the panic
    // names which type pairing broke instead of just saying "called unwrap
    // on a None value".
    match array.data_type() {
        ArrowDataType::Int8 => arrow_col!(array, Int8, Int8Array),
        ArrowDataType::Int16 => arrow_col!(array, Int16, Int16Array),
        ArrowDataType::Int32 => arrow_col!(array, Int32, Int32Array),
        ArrowDataType::Int64 => arrow_col!(array, Int64, Int64Array),
        ArrowDataType::UInt8 => arrow_col!(array, UInt8, UInt8Array),
        ArrowDataType::UInt16 => arrow_col!(array, UInt16, UInt16Array),
        ArrowDataType::UInt32 => arrow_col!(array, UInt32, UInt32Array),
        ArrowDataType::UInt64 => arrow_col!(array, UInt64, UInt64Array),
        ArrowDataType::Float32 => arrow_col!(array, Float32, Float32Array),
        ArrowDataType::Float64 => arrow_col!(array, Float64, Float64Array),
        ArrowDataType::Utf8 => arrow_col!(array, Utf8, StringArray),
        ArrowDataType::LargeUtf8 => arrow_col!(array, LargeUtf8, LargeStringArray),
        ArrowDataType::Utf8View => arrow_col!(array, Utf8View, StringViewArray),
        ArrowDataType::Boolean => arrow_col!(array, Bool, BooleanArray),
        ArrowDataType::Date32 => arrow_col!(array, Date32, Date32Array),
        ArrowDataType::Date64 => arrow_col!(array, Date64, Date64Array),
        // The timezone half of Timestamp(unit, tz) is deliberately ignored
        // here: Arrow stores tz-aware timestamps as UTC instants, so the
        // underlying i64 needs no adjustment. write_dataframe() warns once
        // per tz-aware column that Excel will show UTC wall-clock time.
        ArrowDataType::Timestamp(TimeUnit::Second, _) => {
            arrow_col!(array, TimestampSecond, TimestampSecondArray)
        }
        ArrowDataType::Timestamp(TimeUnit::Millisecond, _) => {
            arrow_col!(array, TimestampMillisecond, TimestampMillisecondArray)
        }
        ArrowDataType::Timestamp(TimeUnit::Microsecond, _) => {
            arrow_col!(array, TimestampMicrosecond, TimestampMicrosecondArray)
        }
        ArrowDataType::Timestamp(TimeUnit::Nanosecond, _) => {
            arrow_col!(array, TimestampNanosecond, TimestampNanosecondArray)
        }
        other => Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(format!(
            "write_dataframe(): column type {other:?} isn't supported yet \
             ({SUPPORTED_ARROW_TYPES})"
        ))),
    }
}

// Numeric widths all collapse to f64 because that is the only numeric
// type an XLSX cell can hold. int64/uint64 magnitudes above 2^53 lose
// precision in that conversion -- unavoidable, and equally true of the
// pre-existing Int64 path and of Excel itself.
macro_rules! num_cell {
    ($a:expr, $row:expr) => {
        if $a.is_null($row) {
            CellValue::Blank
        } else {
            CellValue::Num($a.value($row) as f64)
        }
    };
}

// Arrow temporal value -> Excel serial. $per_day is the number of the
// column's units in one day; dividing before adding the epoch offset
// keeps the magnitudes small enough to stay exact.
macro_rules! datetime_cell {
    ($a:expr, $row:expr, $per_day:expr) => {
        if $a.is_null($row) {
            CellValue::Blank
        } else {
            CellValue::DateTime($a.value($row) as f64 / $per_day + EXCEL_UNIX_EPOCH_DAYS)
        }
    };
}

// Adds column/row context to a failed Arrow cell write. Of the cell types
// write_dataframe() produces, only the temporal ones can fail for a reason
// the caller can act on -- a serial outside Excel's 1900-9999 range -- and
// upstream's message for that ("Serial datetime: '-25567' outside ...")
// doesn't say which column produced it. Everything else keeps the standard
// mapping, so IoError still surfaces as OSError rather than ValueError.
fn arrow_write_err(
    cv: &CellValue<'_>,
    column_name: &str,
    row: usize,
    e: rust_xlsxwriter::XlsxError,
) -> PyErr {
    match cv {
        CellValue::Date(_) | CellValue::DateTime(_) => {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
                "write_dataframe(): column '{column_name}', row {row}: {e}"
            ))
        }
        _ => xlsx_err_to_pyerr(e),
    }
}

fn arrow_cell_value<'a>(col: &ArrowColumn<'a>, row: usize) -> CellValue<'a> {
    match col {
        ArrowColumn::Int8(a) => num_cell!(a, row),
        ArrowColumn::Int16(a) => num_cell!(a, row),
        ArrowColumn::Int32(a) => num_cell!(a, row),
        ArrowColumn::Int64(a) => num_cell!(a, row),
        ArrowColumn::UInt8(a) => num_cell!(a, row),
        ArrowColumn::UInt16(a) => num_cell!(a, row),
        ArrowColumn::UInt32(a) => num_cell!(a, row),
        ArrowColumn::UInt64(a) => num_cell!(a, row),
        // f32 -> f64 widens exactly, but the decimal shown in Excel is
        // the f64 rendering of the f32 value (0.1f32 becomes
        // 0.10000000149011612), which is the same behaviour pandas and
        // pyarrow give when casting.
        ArrowColumn::Float32(a) => num_cell!(a, row),
        // Deliberately not routed through num_cell!: its `as f64` would be
        // a no-op here and clippy::unnecessary_cast fires on local macro
        // expansions, which CI promotes to an error.
        ArrowColumn::Float64(a) => {
            if a.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Num(a.value(row))
            }
        }
        ArrowColumn::Utf8(a) => {
            let arr: &'a StringArray = *a;
            if arr.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Str(Cow::Borrowed(arr.value(row)))
            }
        }
        ArrowColumn::LargeUtf8(a) => {
            let arr: &'a LargeStringArray = *a;
            if arr.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Str(Cow::Borrowed(arr.value(row)))
            }
        }
        ArrowColumn::Utf8View(a) => {
            let arr: &'a StringViewArray = *a;
            if arr.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Str(Cow::Borrowed(arr.value(row)))
            }
        }
        ArrowColumn::Bool(a) => {
            if a.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Bool(a.value(row))
            }
        }
        // Date32 is already in days, so it needs no division -- just the
        // epoch shift.
        ArrowColumn::Date32(a) => {
            if a.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Date(a.value(row) as f64 + EXCEL_UNIX_EPOCH_DAYS)
            }
        }
        ArrowColumn::Date64(a) => {
            if a.is_null(row) {
                CellValue::Blank
            } else {
                CellValue::Date(a.value(row) as f64 / 86_400_000.0 + EXCEL_UNIX_EPOCH_DAYS)
            }
        }
        ArrowColumn::TimestampSecond(a) => datetime_cell!(a, row, 86_400.0),
        ArrowColumn::TimestampMillisecond(a) => datetime_cell!(a, row, 86_400_000.0),
        ArrowColumn::TimestampMicrosecond(a) => datetime_cell!(a, row, 86_400_000_000.0),
        // A nanosecond count for a modern date (~1.7e18) exceeds f64's
        // 2^53 exact-integer range, so this loses sub-microsecond
        // precision. Excel serials cannot represent it either (one f64
        // ulp at ~45000 days is roughly 0.9us), so nothing is lost that
        // could have been stored.
        ArrowColumn::TimestampNanosecond(a) => datetime_cell!(a, row, 86_400_000_000_000.0),
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

    fn set_font_color<'a>(
        mut slf: PyRefMut<'a, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let c = parse_color(color)?;
        slf.update(|f| f.set_font_color(c));
        Ok(slf)
    }

    fn set_background_color<'a>(
        mut slf: PyRefMut<'a, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let c = parse_color(color)?;
        slf.update(|f| f.set_background_color(c));
        Ok(slf)
    }

    fn set_border<'a>(mut slf: PyRefMut<'a, Self>, border: &str) -> PyResult<PyRefMut<'a, Self>> {
        let style = parse_border(border)?;
        slf.update(|f| f.set_border(style));
        Ok(slf)
    }

    fn set_border_color<'a>(
        mut slf: PyRefMut<'a, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let c = parse_color(color)?;
        slf.update(|f| {
            f.set_border_color(c)
                .set_border_top_color(c)
                .set_border_bottom_color(c)
                .set_border_left_color(c)
                .set_border_right_color(c)
        });
        Ok(slf)
    }

    // Per-side border colours. set_border_color above sets all four at
    // once; these target one side each. Note the per-side border *styles*
    // are already exposed as set_top_border / set_bottom_border /
    // set_left_border / set_right_border, which reverse upstream's word
    // order.
    fn set_border_top_color<'a>(
        mut slf: PyRefMut<'a, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let c = parse_color(color)?;
        slf.update(|f| f.set_border_top_color(c));
        Ok(slf)
    }

    fn set_border_bottom_color<'a>(
        mut slf: PyRefMut<'a, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let c = parse_color(color)?;
        slf.update(|f| f.set_border_bottom_color(c));
        Ok(slf)
    }

    fn set_border_left_color<'a>(
        mut slf: PyRefMut<'a, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let c = parse_color(color)?;
        slf.update(|f| f.set_border_left_color(c));
        Ok(slf)
    }

    fn set_border_right_color<'a>(
        mut slf: PyRefMut<'a, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let c = parse_color(color)?;
        slf.update(|f| f.set_border_right_color(c));
        Ok(slf)
    }

    // Diagonal borders. The style and colour work like any other side; the
    // type chooses which diagonal(s) the border is drawn on.
    fn set_border_diagonal<'a>(
        mut slf: PyRefMut<'a, Self>,
        border: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let style = parse_border(border)?;
        slf.update(|f| f.set_border_diagonal(style));
        Ok(slf)
    }

    fn set_border_diagonal_color<'a>(
        mut slf: PyRefMut<'a, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let c = parse_color(color)?;
        slf.update(|f| f.set_border_diagonal_color(c));
        Ok(slf)
    }

    // One of none, up, down, up_down.
    fn set_border_diagonal_type<'a>(
        mut slf: PyRefMut<'a, Self>,
        border_type: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let parsed = parse_diagonal_border(border_type)?;
        slf.update(|f| f.set_border_diagonal_type(parsed));
        Ok(slf)
    }

    // Cell protection. These only take effect once the worksheet itself is
    // protected: everything is locked by default, so set_unlocked is what
    // makes a cell editable on a protected sheet.
    fn set_locked(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_locked());
        slf
    }

    fn set_unlocked(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_unlocked());
        slf
    }

    // Hides the cell's formula in the formula bar on a protected sheet.
    fn set_hidden(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_hidden());
        slf
    }

    fn set_font_strikethrough(mut slf: PyRefMut<'_, Self>) -> PyRefMut<'_, Self> {
        slf.update(|f| f.set_font_strikethrough());
        slf
    }

    // The pattern foreground colour, which pairs with set_pattern and
    // set_background_color.
    fn set_foreground_color<'a>(
        mut slf: PyRefMut<'a, Self>,
        color: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let c = parse_color(color)?;
        slf.update(|f| f.set_foreground_color(c));
        Ok(slf)
    }

    fn set_top_border<'a>(
        mut slf: PyRefMut<'a, Self>,
        border: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let style = parse_border(border)?;
        slf.update(|f| f.set_border_top(style));
        Ok(slf)
    }

    fn set_bottom_border<'a>(
        mut slf: PyRefMut<'a, Self>,
        border: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let style = parse_border(border)?;
        slf.update(|f| f.set_border_bottom(style));
        Ok(slf)
    }

    fn set_left_border<'a>(
        mut slf: PyRefMut<'a, Self>,
        border: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let style = parse_border(border)?;
        slf.update(|f| f.set_border_left(style));
        Ok(slf)
    }

    fn set_right_border<'a>(
        mut slf: PyRefMut<'a, Self>,
        border: &str,
    ) -> PyResult<PyRefMut<'a, Self>> {
        let style = parse_border(border)?;
        slf.update(|f| f.set_border_right(style));
        Ok(slf)
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
// TABLE / TABLECOLUMN CLASSES
// ============================================
// Excel worksheet tables (Insert > Table), via rust_xlsxwriter's Table
// and TableColumn builders. Usage mirrors the crate's own: write the
// data with the normal write_*() methods first, build a Table (with
// TableColumn entries for header/total-row/formula/format options),
// then call Worksheet.add_table() over the range that data occupies.

fn parse_table_style(s: &str) -> PyResult<TableStyle> {
    match s.to_lowercase().as_str() {
        "none" => Ok(TableStyle::None),
        "light1" => Ok(TableStyle::Light1),
        "light2" => Ok(TableStyle::Light2),
        "light3" => Ok(TableStyle::Light3),
        "light4" => Ok(TableStyle::Light4),
        "light5" => Ok(TableStyle::Light5),
        "light6" => Ok(TableStyle::Light6),
        "light7" => Ok(TableStyle::Light7),
        "light8" => Ok(TableStyle::Light8),
        "light9" => Ok(TableStyle::Light9),
        "light10" => Ok(TableStyle::Light10),
        "light11" => Ok(TableStyle::Light11),
        "light12" => Ok(TableStyle::Light12),
        "light13" => Ok(TableStyle::Light13),
        "light14" => Ok(TableStyle::Light14),
        "light15" => Ok(TableStyle::Light15),
        "light16" => Ok(TableStyle::Light16),
        "light17" => Ok(TableStyle::Light17),
        "light18" => Ok(TableStyle::Light18),
        "light19" => Ok(TableStyle::Light19),
        "light20" => Ok(TableStyle::Light20),
        "light21" => Ok(TableStyle::Light21),
        "medium1" => Ok(TableStyle::Medium1),
        "medium2" => Ok(TableStyle::Medium2),
        "medium3" => Ok(TableStyle::Medium3),
        "medium4" => Ok(TableStyle::Medium4),
        "medium5" => Ok(TableStyle::Medium5),
        "medium6" => Ok(TableStyle::Medium6),
        "medium7" => Ok(TableStyle::Medium7),
        "medium8" => Ok(TableStyle::Medium8),
        "medium9" => Ok(TableStyle::Medium9),
        "medium10" => Ok(TableStyle::Medium10),
        "medium11" => Ok(TableStyle::Medium11),
        "medium12" => Ok(TableStyle::Medium12),
        "medium13" => Ok(TableStyle::Medium13),
        "medium14" => Ok(TableStyle::Medium14),
        "medium15" => Ok(TableStyle::Medium15),
        "medium16" => Ok(TableStyle::Medium16),
        "medium17" => Ok(TableStyle::Medium17),
        "medium18" => Ok(TableStyle::Medium18),
        "medium19" => Ok(TableStyle::Medium19),
        "medium20" => Ok(TableStyle::Medium20),
        "medium21" => Ok(TableStyle::Medium21),
        "medium22" => Ok(TableStyle::Medium22),
        "medium23" => Ok(TableStyle::Medium23),
        "medium24" => Ok(TableStyle::Medium24),
        "medium25" => Ok(TableStyle::Medium25),
        "medium26" => Ok(TableStyle::Medium26),
        "medium27" => Ok(TableStyle::Medium27),
        "medium28" => Ok(TableStyle::Medium28),
        "dark1" => Ok(TableStyle::Dark1),
        "dark2" => Ok(TableStyle::Dark2),
        "dark3" => Ok(TableStyle::Dark3),
        "dark4" => Ok(TableStyle::Dark4),
        "dark5" => Ok(TableStyle::Dark5),
        "dark6" => Ok(TableStyle::Dark6),
        "dark7" => Ok(TableStyle::Dark7),
        "dark8" => Ok(TableStyle::Dark8),
        "dark9" => Ok(TableStyle::Dark9),
        "dark10" => Ok(TableStyle::Dark10),
        "dark11" => Ok(TableStyle::Dark11),
        other => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
            "Unknown table style '{other}'. Expected 'none', 'light1'..'light21', \
             'medium1'..'medium28', or 'dark1'..'dark11'."
        ))),
    }
}

// Unlike table style (a fixed set with no escape hatch), any total-row
// function string that isn't one of the built-in names is passed
// through as a custom formula -- matches TableFunction::Custom(Formula)
// in the underlying crate, which exists for exactly this case.
fn parse_table_function(s: &str) -> TableFunction {
    match s.to_lowercase().replace('_', "").as_str() {
        "none" => TableFunction::None,
        "average" => TableFunction::Average,
        "count" => TableFunction::Count,
        "countnumbers" => TableFunction::CountNumbers,
        "max" => TableFunction::Max,
        "min" => TableFunction::Min,
        "sum" => TableFunction::Sum,
        "stddev" => TableFunction::StdDev,
        "var" => TableFunction::Var,
        _ => TableFunction::Custom(s.into()),
    }
}

#[pyclass]
struct TableColumn {
    inner: RustTableColumn,
}

impl TableColumn {
    fn new() -> Self {
        TableColumn {
            inner: RustTableColumn::new(),
        }
    }

    fn update(&mut self, op: impl FnOnce(RustTableColumn) -> RustTableColumn) {
        let current = std::mem::replace(&mut self.inner, RustTableColumn::new());
        self.inner = op(current);
    }
}

#[pymethods]
impl TableColumn {
    #[new]
    fn py_new() -> Self {
        TableColumn::new()
    }

    // All setters return self (chainable), same convention as Format.
    fn set_header<'a>(mut slf: PyRefMut<'a, Self>, caption: &str) -> PyRefMut<'a, Self> {
        slf.update(|c| c.set_header(caption));
        slf
    }

    fn set_total_function<'a>(mut slf: PyRefMut<'a, Self>, function: &str) -> PyRefMut<'a, Self> {
        let f = parse_table_function(function);
        slf.update(|c| c.set_total_function(f));
        slf
    }

    fn set_total_label<'a>(mut slf: PyRefMut<'a, Self>, label: &str) -> PyRefMut<'a, Self> {
        slf.update(|c| c.set_total_label(label));
        slf
    }

    fn set_formula<'a>(mut slf: PyRefMut<'a, Self>, formula: &str) -> PyRefMut<'a, Self> {
        slf.update(|c| c.set_formula(formula));
        slf
    }

    fn set_format<'a>(mut slf: PyRefMut<'a, Self>, format: &Format) -> PyRefMut<'a, Self> {
        let fmt = format.inner.clone();
        slf.update(|c| c.set_format(fmt));
        slf
    }

    fn set_header_format<'a>(mut slf: PyRefMut<'a, Self>, format: &Format) -> PyRefMut<'a, Self> {
        let fmt = format.inner.clone();
        slf.update(|c| c.set_header_format(fmt));
        slf
    }
}

#[pyclass]
struct Table {
    inner: RustTable,
}

impl Table {
    fn new() -> Self {
        Table {
            inner: RustTable::new(),
        }
    }

    fn update(&mut self, op: impl FnOnce(RustTable) -> RustTable) {
        let current = std::mem::replace(&mut self.inner, RustTable::new());
        self.inner = op(current);
    }
}

#[pymethods]
impl Table {
    #[new]
    fn py_new() -> Self {
        Table::new()
    }

    fn set_header_row<'a>(mut slf: PyRefMut<'a, Self>, enable: bool) -> PyRefMut<'a, Self> {
        slf.update(|t| t.set_header_row(enable));
        slf
    }

    fn set_total_row<'a>(mut slf: PyRefMut<'a, Self>, enable: bool) -> PyRefMut<'a, Self> {
        slf.update(|t| t.set_total_row(enable));
        slf
    }

    fn set_banded_rows<'a>(mut slf: PyRefMut<'a, Self>, enable: bool) -> PyRefMut<'a, Self> {
        slf.update(|t| t.set_banded_rows(enable));
        slf
    }

    fn set_banded_columns<'a>(mut slf: PyRefMut<'a, Self>, enable: bool) -> PyRefMut<'a, Self> {
        slf.update(|t| t.set_banded_columns(enable));
        slf
    }

    fn set_first_column<'a>(mut slf: PyRefMut<'a, Self>, enable: bool) -> PyRefMut<'a, Self> {
        slf.update(|t| t.set_first_column(enable));
        slf
    }

    fn set_last_column<'a>(mut slf: PyRefMut<'a, Self>, enable: bool) -> PyRefMut<'a, Self> {
        slf.update(|t| t.set_last_column(enable));
        slf
    }

    fn set_autofilter<'a>(mut slf: PyRefMut<'a, Self>, enable: bool) -> PyRefMut<'a, Self> {
        slf.update(|t| t.set_autofilter(enable));
        slf
    }

    fn set_columns<'a>(
        mut slf: PyRefMut<'a, Self>,
        columns: Vec<PyRef<'_, TableColumn>>,
    ) -> PyRefMut<'a, Self> {
        let cols: Vec<RustTableColumn> = columns.iter().map(|c| c.inner.clone()).collect();
        slf.update(|t| t.set_columns(&cols));
        slf
    }

    fn set_name<'a>(mut slf: PyRefMut<'a, Self>, name: &str) -> PyRefMut<'a, Self> {
        slf.update(|t| t.set_name(name));
        slf
    }

    fn set_style<'a>(mut slf: PyRefMut<'a, Self>, style: &str) -> PyResult<PyRefMut<'a, Self>> {
        let s = parse_table_style(style)?;
        slf.update(|t| t.set_style(s));
        Ok(slf)
    }

    fn set_alt_text<'a>(mut slf: PyRefMut<'a, Self>, alt_text: &str) -> PyRefMut<'a, Self> {
        slf.update(|t| t.set_alt_text(alt_text));
        slf
    }

    fn set_alt_text_title<'a>(mut slf: PyRefMut<'a, Self>, title: &str) -> PyRefMut<'a, Self> {
        slf.update(|t| t.set_alt_text_title(title));
        slf
    }

    fn has_header_row(&self) -> bool {
        self.inner.has_header_row()
    }

    fn has_total_row(&self) -> bool {
        self.inner.has_total_row()
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
        let mut wb = wb_ref
            .inner
            .try_borrow_mut()
            .map_err(|_| reentrant_workbook_err())?;
        let sheet = wb
            .worksheet_from_index(self.index)
            .map_err(xlsx_err_to_pyerr)?;
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

    // write_rows() is a list-of-lists bulk write path. Unlike write_records()
    // which takes list[dict] and does a Python hash lookup per cell, this
    // takes list[list] and accesses values by position — zero hash lookups.
    // For bulk data from CSV readers, DB cursors, or any positional source,
    // this is faster than write_records().
    //
    // Implementation note: we pre-classify ALL cell values (Python→Rust
    // type conversion) in one pass before acquiring the workbook borrow,
    // then write the fully-classified data in a tight Rust loop. This keeps
    // the costly Python API calls separated from the I/O path and lets the
    // Rust write loop run without any Python interaction.
    //
    // write_header=True writes the first row with header_format (matching
    // write_records() API). Default is write_header=False.
    #[pyo3(signature = (start_row, start_col, rows, format=None, header_format=None, write_header=false))]
    fn write_rows(
        &self,
        py: Python<'_>,
        start_row: u32,
        start_col: u16,
        rows: &Bound<'_, PyList>,
        format: Option<&Format>,
        header_format: Option<&Format>,
        write_header: bool,
    ) -> PyResult<()> {
        if rows.is_empty() {
            return Ok(());
        }

        // rows.len() is the total number of rows being written -- when
        // write_header=True, the first element IS the header row (not an
        // extra one). Adding header_rows here would double-count and push
        // min_allowed_row one row past the actual last write, incorrectly
        // rejecting valid subsequent writes in constant_memory mode.
        let last_row = start_row + rows.len() as u32 - 1;
        self.check_row_order_range(start_row, last_row)?;

        let data_fmt = format.map(|f| &f.inner);
        let head_fmt = header_format.map(|f| &f.inner);

        // Pre-classify: convert all Python values to CellValue in one pass
        // before acquiring the workbook borrow. Each row becomes a Vec<CellValue>.
        // Using rows.get_item(r) + row_list.get_item(c) (direct index, O(1))
        // avoids the per-row downcast and iterator allocation of the naive
        // downcast::<PyList>() + .iter() approach.
        let n_rows = rows.len();
        let mut classified_rows: Vec<(Vec<CellValue<'static>>, bool)> = Vec::with_capacity(n_rows);

        for r in 0..n_rows {
            let row_obj = rows.get_item(r)?;
            let row_list = row_obj.downcast::<PyList>().map_err(|_| {
                PyErr::new::<pyo3::exceptions::PyTypeError, _>(
                    "write_rows(): each row must be a list",
                )
            })?;
            let n_cols = row_list.len();
            let mut row_vals = Vec::with_capacity(n_cols);
            for c in 0..n_cols {
                let val = row_list.get_item(c)?;
                row_vals.push(classify(&val)?);
            }
            let is_header = write_header && r == 0;
            classified_rows.push((row_vals, is_header));
        }

        // Write all pre-classified values in a tight Rust loop.
        let wb_ref = self.workbook.borrow(py);
        let mut wb = wb_ref
            .inner
            .try_borrow_mut()
            .map_err(|_| reentrant_workbook_err())?;
        let sheet = wb
            .worksheet_from_index(self.index)
            .map_err(xlsx_err_to_pyerr)?;

        for (row_num, (row_vals, is_header)) in (start_row..).zip(classified_rows.iter()) {
            let fmt = if *is_header { head_fmt } else { data_fmt };
            for (c, cv) in row_vals.iter().enumerate() {
                write_value(sheet, row_num, start_col + c as u16, cv, fmt)
                    .map_err(xlsx_err_to_pyerr)?;
            }
        }

        Ok(())
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
        let mut wb = wb_ref
            .inner
            .try_borrow_mut()
            .map_err(|_| reentrant_workbook_err())?;
        let sheet = wb
            .worksheet_from_index(self.index)
            .map_err(xlsx_err_to_pyerr)?;

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
                    &CellValue::Str(Cow::Borrowed(h.as_str())),
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
    // initial __arrow_c_stream__() call hands over the data. See
    // SUPPORTED_ARROW_TYPES for the column types accepted; date and
    // timestamp columns get a date number format applied automatically
    // so they render as dates rather than as raw serial numbers.
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
                ArrowDataType::Int8
                | ArrowDataType::Int16
                | ArrowDataType::Int32
                | ArrowDataType::Int64
                | ArrowDataType::UInt8
                | ArrowDataType::UInt16
                | ArrowDataType::UInt32
                | ArrowDataType::UInt64
                | ArrowDataType::Float32
                | ArrowDataType::Float64
                | ArrowDataType::Utf8
                | ArrowDataType::LargeUtf8
                | ArrowDataType::Utf8View
                | ArrowDataType::Boolean
                | ArrowDataType::Date32
                | ArrowDataType::Date64 => {}
                // TimeUnit has exactly four variants, all handled by
                // resolve_arrow_column, so the unit needs no check here.
                // The timezone does: Excel has no concept of one, and
                // silently shifting someone's timestamps by their UTC
                // offset is the kind of error that only shows up much
                // later. Warned once per column, at schema-validation
                // time, so it costs nothing on the per-cell hot path.
                ArrowDataType::Timestamp(_, tz) => {
                    if let Some(tz) = tz {
                        let warn_cls = py.get_type_bound::<pyo3::exceptions::PyUserWarning>();
                        PyErr::warn_bound(
                            py,
                            warn_cls.as_any(),
                            &format!(
                                "write_dataframe(): column '{}' is timezone-aware \
                                 ({tz}); Excel has no timezone concept, so values \
                                 are written as UTC wall-clock time. Convert with \
                                 .dt.tz_convert(None) first to choose the offset \
                                 yourself.",
                                field.name()
                            ),
                            1,
                        )?;
                    }
                }
                other => {
                    return Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(format!(
                        "write_dataframe(): column '{}' has type {other:?}, \
                         which isn't supported yet ({SUPPORTED_ARROW_TYPES})",
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
        let mut wb = wb_ref
            .inner
            .try_borrow_mut()
            .map_err(|_| reentrant_workbook_err())?;
        let sheet = wb
            .worksheet_from_index(self.index)
            .map_err(xlsx_err_to_pyerr)?;

        let mut row_cursor = start_row;
        if write_header {
            for (i, name) in field_names.iter().enumerate() {
                write_value(
                    sheet,
                    row_cursor,
                    start_col + i as u16,
                    &CellValue::Str(Cow::Borrowed(name.as_str())),
                    head_fmt,
                )
                .map_err(xlsx_err_to_pyerr)?;
            }
            row_cursor += 1;
        }

        // Temporal columns need a number format or Excel renders the raw
        // serial. Built once for the whole call rather than per cell:
        // rust_xlsxwriter dedupes identical formats in its xf table, so a
        // per-cell Format would be correct but would allocate 100k times.
        let date_fmt = RustFormat::new().set_num_format("yyyy-mm-dd");
        let datetime_fmt = RustFormat::new().set_num_format("yyyy-mm-dd hh:mm:ss");

        for batch in &batches {
            let columns: Vec<ArrowColumn<'_>> = (0..batch.num_columns())
                .map(|c| resolve_arrow_column(batch.column(c).as_ref()))
                .collect::<PyResult<Vec<_>>>()?;

            // A column's type is fixed for the whole batch, so pick its
            // format once here instead of re-matching on every cell --
            // this keeps the inner loop's cost identical to before for
            // non-temporal data.
            let col_fmts: Vec<Option<&RustFormat>> = columns
                .iter()
                .map(|col| match col {
                    ArrowColumn::Date32(_) | ArrowColumn::Date64(_) => Some(&date_fmt),
                    ArrowColumn::TimestampSecond(_)
                    | ArrowColumn::TimestampMillisecond(_)
                    | ArrowColumn::TimestampMicrosecond(_)
                    | ArrowColumn::TimestampNanosecond(_) => Some(&datetime_fmt),
                    _ => None,
                })
                .collect();

            for r in 0..batch.num_rows() {
                for (c, col) in columns.iter().enumerate() {
                    let cv = arrow_cell_value(col, r);
                    let col_num = start_col + c as u16;
                    write_value(sheet, row_cursor, col_num, &cv, col_fmts[c])
                        .map_err(|e| arrow_write_err(&cv, &field_names[c], r, e))?;
                }
                row_cursor += 1;
            }
        }

        Ok(())
    }

    // Applies a conditional formatting rule to a cell range. Accepts any
    // of the ConditionalFormat* objects. Upstream's add_conditional_format
    // is generic over the ConditionalFormat trait, so the concrete type is
    // recovered by extract_cf()'s downcast chain first.
    fn add_conditional_format(
        &self,
        py: Python<'_>,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        cf: &Bound<'_, PyAny>,
    ) -> PyResult<()> {
        self.check_row_order_range(first_row, last_row)?;
        let any = extract_cf(cf)?;
        let (r1, c1, r2, c2) = (first_row, first_col, last_row, last_col);
        self.with_sheet(py, |sheet| match &any {
            AnyCf::Cell(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
            AnyCf::Blank(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
            AnyCf::Duplicate(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
            AnyCf::ErrorCf(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
            AnyCf::Formula(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
            AnyCf::Average(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
            AnyCf::Top(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
            AnyCf::Text(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
            AnyCf::Date(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
            AnyCf::Scale2(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
            AnyCf::Scale3(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
            AnyCf::DataBar(v) => sheet.add_conditional_format(r1, c1, r2, c2, v).map(|_| ()),
        })
    }

    // Adds a single sparkline to one cell.
    fn add_sparkline(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        sparkline: &Sparkline,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        let inner = sparkline.inner.clone();
        // rustfmt's fn_call_width (default 60) applies to the argument
        // list of with_sheet: "py, |sheet| sheet.add_sparkline(row, col,
        // &inner).map(|_| ())" is 61 chars, so the closure body has to go
        // in a block. This is why the pre-existing insert_image call next
        // door fits on one line at exactly 60 and this one doesn't.
        self.with_sheet(py, |sheet| {
            sheet.add_sparkline(row, col, &inner).map(|_| ())
        })
    }

    // Adds a grouped sparkline spanning a range of cells. Grouped
    // sparklines share one set of options and, with set_group_max/min, one
    // scale.
    fn add_sparkline_group(
        &self,
        py: Python<'_>,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        sparkline: &Sparkline,
    ) -> PyResult<()> {
        self.check_row_order_range(first_row, last_row)?;
        let inner = sparkline.inner.clone();
        let (r1, c1, r2, c2) = (first_row, first_col, last_row, last_col);
        // Split across two statements rather than chaining .map(|_| ()):
        // the chained form is 61 chars wide and rustfmt's chain_width
        // default is 60, so it would get reflowed. A two-statement closure
        // body also can't be collapsed back to a single expression.
        self.with_sheet(py, |sheet| {
            sheet.add_sparkline_group(r1, c1, r2, c2, &inner)?;
            Ok(())
        })
    }

    // Inserts a chart at a cell, optionally offset within it by a number
    // of pixels. Series are attached to the chart beforehand with
    // Chart.push_series().
    #[pyo3(signature = (row, col, chart, x_offset=0, y_offset=0))]
    fn insert_chart(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        chart: &Chart,
        x_offset: u32,
        y_offset: u32,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        let inner = &chart.inner;
        self.with_sheet(py, |sheet| {
            sheet.insert_chart_with_offset(row, col, inner, x_offset, y_offset)?;
            Ok(())
        })
    }

    // Applies a format to a whole column. Cells written afterwards inherit
    // it unless they carry a format of their own. This is the escape hatch
    // that was missing when write_dataframe gained date support: without
    // it a caller had no way to reformat a column after the fact.
    fn set_column_format(&self, py: Python<'_>, col: u16, format: &Format) -> PyResult<()> {
        let fmt = &format.inner;
        self.with_sheet(py, |sheet| {
            sheet.set_column_format(col, fmt)?;
            Ok(())
        })
    }

    fn set_column_range_format(
        &self,
        py: Python<'_>,
        first_col: u16,
        last_col: u16,
        format: &Format,
    ) -> PyResult<()> {
        let fmt = &format.inner;
        self.with_sheet(py, |sheet| {
            sheet.set_column_range_format(first_col, last_col, fmt)?;
            Ok(())
        })
    }

    fn set_row_format(&self, py: Python<'_>, row: u32, format: &Format) -> PyResult<()> {
        self.check_row_order(row)?;
        let fmt = &format.inner;
        self.with_sheet(py, |sheet| {
            sheet.set_row_format(row, fmt)?;
            Ok(())
        })
    }

    fn set_cell_format(&self, py: Python<'_>, row: u32, col: u16, format: &Format) -> PyResult<()> {
        self.check_row_order(row)?;
        let fmt = &format.inner;
        self.with_sheet(py, |sheet| {
            sheet.set_cell_format(row, col, fmt)?;
            Ok(())
        })
    }

    fn set_range_format(
        &self,
        py: Python<'_>,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        format: &Format,
    ) -> PyResult<()> {
        self.check_row_order_range(first_row, last_row)?;
        let fmt = &format.inner;
        let (r1, c1, r2, c2) = (first_row, first_col, last_row, last_col);
        self.with_sheet(py, |sheet| {
            sheet.set_range_format(r1, c1, r2, c2, fmt)?;
            Ok(())
        })
    }

    // ---- page setup and printing ----
    // Sheet-level page properties. None of these are guarded by the
    // constant-memory row-order check: they set worksheet metadata rather
    // than writing cells, so they are order-independent even though some
    // take row numbers.

    fn set_landscape(&self, py: Python<'_>) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_landscape();
            Ok(())
        })
    }

    fn set_portrait(&self, py: Python<'_>) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_portrait();
            Ok(())
        })
    }

    // Excel's numeric paper size codes: 1 = Letter, 9 = A4, 8 = A3.
    fn set_paper_size(&self, py: Python<'_>, paper_size: u8) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_paper_size(paper_size);
            Ok(())
        })
    }

    // true prints down then across, false across then down.
    fn set_page_order(&self, py: Python<'_>, enable: bool) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_page_order(enable);
            Ok(())
        })
    }

    // Margins in inches.
    fn set_margins(
        &self,
        py: Python<'_>,
        left: f64,
        right: f64,
        top: f64,
        bottom: f64,
        header: f64,
        footer: f64,
    ) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_margins(left, right, top, bottom, header, footer);
            Ok(())
        })
    }

    fn set_print_area(
        &self,
        py: Python<'_>,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
    ) -> PyResult<()> {
        let (r1, c1, r2, c2) = (first_row, first_col, last_row, last_col);
        self.with_sheet(py, |sheet| {
            sheet.set_print_area(r1, c1, r2, c2)?;
            Ok(())
        })
    }

    // Rows repeated at the top of every printed page.
    fn set_repeat_rows(&self, py: Python<'_>, first_row: u32, last_row: u32) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_repeat_rows(first_row, last_row)?;
            Ok(())
        })
    }

    // Columns repeated at the left of every printed page.
    fn set_repeat_columns(&self, py: Python<'_>, first_col: u16, last_col: u16) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_repeat_columns(first_col, last_col)?;
            Ok(())
        })
    }

    // Scale the sheet to fit a given number of pages. 0 means automatic.
    fn set_print_fit_to_pages(&self, py: Python<'_>, width: u16, height: u16) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_print_fit_to_pages(width, height);
            Ok(())
        })
    }

    // Percentage, 10 to 400. Mutually exclusive with fit-to-pages in Excel.
    fn set_print_scale(&self, py: Python<'_>, scale: u16) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_print_scale(scale);
            Ok(())
        })
    }

    // Horizontal page breaks, given as the row numbers to break above.
    fn set_page_breaks(&self, py: Python<'_>, breaks: Vec<u32>) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_page_breaks(&breaks)?;
            Ok(())
        })
    }

    // Vertical page breaks. Note upstream takes u32 column numbers here,
    // not the u16 used elsewhere.
    fn set_vertical_page_breaks(&self, py: Python<'_>, breaks: Vec<u32>) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_vertical_page_breaks(&breaks)?;
            Ok(())
        })
    }

    fn set_print_gridlines(&self, py: Python<'_>, enable: bool) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_print_gridlines(enable);
            Ok(())
        })
    }

    // Print the row numbers and column letters.
    fn set_print_headings(&self, py: Python<'_>, enable: bool) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_print_headings(enable);
            Ok(())
        })
    }

    fn set_print_center_horizontally(&self, py: Python<'_>, enable: bool) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_print_center_horizontally(enable);
            Ok(())
        })
    }

    fn set_print_center_vertically(&self, py: Python<'_>, enable: bool) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_print_center_vertically(enable);
            Ok(())
        })
    }

    fn set_print_black_and_white(&self, py: Python<'_>, enable: bool) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_print_black_and_white(enable);
            Ok(())
        })
    }

    fn set_print_draft(&self, py: Python<'_>, enable: bool) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_print_draft(enable);
            Ok(())
        })
    }

    fn set_print_first_page_number(&self, py: Python<'_>, page_number: u16) -> PyResult<()> {
        self.with_sheet(py, |sheet| {
            sheet.set_print_first_page_number(page_number);
            Ok(())
        })
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
        let c = parse_color(color)?;
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

    // protect() enables worksheet protection.
    //
    // IMPORTANT: calling protect() with no password (or password=None)
    // still enables Excel sheet protection, but with an empty-string
    // password. In Excel, an empty-string password means the sheet IS
    // protected (editing is blocked) but anyone can unprotect it
    // without entering a password. This is intentional -- it deters
    // casual edits without requiring secret management -- but callers
    // should be aware it is NOT a security mechanism. Pass a non-empty
    // password string if you want password-gated unprotection.
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
        self.with_sheet(py, |sheet| match fmt {
            Some(f) => sheet
                .write_formula_with_format(row, col, formula, f)
                .map(|_| ()),
            None => sheet.write_formula(row, col, formula).map(|_| ()),
        })
    }

    // write_url writes a hyperlink cell. Optional `text` overrides the
    // display label (defaults to the URL itself). Optional `tip` sets the
    // tooltip that appears on hover. Both map to rust_xlsxwriter's Url
    // builder, which was introduced in 0.75.
    #[pyo3(signature = (row, col, url, format=None, text=None, tip=None))]
    fn write_url(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        url: &str,
        format: Option<&Format>,
        text: Option<&str>,
        tip: Option<&str>,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        // Build the Url value via the builder chain. Url::set_text/set_tip
        // consume self and return Url, so we fold the optional fields in.
        // Pass `link` by value (not &link): write_url/write_url_with_format
        // take `impl Into<Url>`, which Url satisfies but &Url does not.
        let link = {
            let mut u = rust_xlsxwriter::Url::new(url);
            if let Some(t) = text {
                u = u.set_text(t);
            }
            if let Some(t) = tip {
                u = u.set_tip(t);
            }
            u
        };
        let fmt = format.map(|f| &f.inner);
        self.with_sheet(py, |sheet| match fmt {
            Some(f) => sheet
                .write_url_with_format(row, col, link.clone(), f)
                .map(|_| ()),
            None => sheet.write_url(row, col, link).map(|_| ()),
        })
    }

    #[pyo3(signature = (row, col, year, month, day, hour, min, sec, format=None))]
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
        let dt = ExcelDateTime::from_ymd(year, month, day)
            .map_err(xlsx_err_to_pyerr)?
            .and_hms(hour, min, sec)
            .map_err(xlsx_err_to_pyerr)?;
        self.with_sheet(py, |sheet| match fmt {
            Some(f) => sheet
                .write_datetime_with_format(row, col, &dt, f)
                .map(|_| ()),
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
        let dt = ExcelDateTime::from_ymd(year, month, day).map_err(xlsx_err_to_pyerr)?;
        self.with_sheet(py, |sheet| match fmt {
            Some(f) => sheet
                .write_datetime_with_format(row, col, &dt, f)
                .map(|_| ()),
            None => sheet.write_datetime(row, col, &dt).map(|_| ()),
        })
    }

    // write_datetime_py / write_date_py accept Python datetime/date objects
    // directly, instead of separate year/month/day/hour/min/sec args.
    // The existing write_datetime()/write_date() with individual components
    // are kept for backwards compatibility.
    //
    // Python's datetime module is accessed via PyO3's bound interface:
    // we pull year/month/day/hour/minute/second out of the Python object
    // with attribute lookups rather than using a special datetime extractor,
    // which avoids a dependency on PyO3's chrono feature.
    #[pyo3(signature = (row, col, dt, format=None))]
    fn write_datetime_py(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        dt: &Bound<'_, PyAny>,
        format: Option<&Format>,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        let year: u16 = dt.getattr("year")?.extract()?;
        let month: u8 = dt.getattr("month")?.extract()?;
        let day: u8 = dt.getattr("day")?.extract()?;
        let hour: u16 = dt.getattr("hour")?.extract()?;
        let minute: u8 = dt.getattr("minute")?.extract()?;
        let second: u8 = dt.getattr("second")?.extract()?;
        let microsecond: u32 = dt.getattr("microsecond")?.extract()?;
        let sec_frac = second as f64 + microsecond as f64 / 1_000_000.0;
        let fmt = format.map(|f| &f.inner);
        let edt = ExcelDateTime::from_ymd(year, month, day)
            .map_err(xlsx_err_to_pyerr)?
            .and_hms(hour, minute, sec_frac)
            .map_err(xlsx_err_to_pyerr)?;
        self.with_sheet(py, |sheet| match fmt {
            Some(f) => sheet
                .write_datetime_with_format(row, col, &edt, f)
                .map(|_| ()),
            None => sheet.write_datetime(row, col, &edt).map(|_| ()),
        })
    }

    #[pyo3(signature = (row, col, date, format=None))]
    fn write_date_py(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        date: &Bound<'_, PyAny>,
        format: Option<&Format>,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        let year: u16 = date.getattr("year")?.extract()?;
        let month: u8 = date.getattr("month")?.extract()?;
        let day: u8 = date.getattr("day")?.extract()?;
        let fmt = format.map(|f| &f.inner);
        let edt = ExcelDateTime::from_ymd(year, month, day).map_err(xlsx_err_to_pyerr)?;
        self.with_sheet(py, |sheet| match fmt {
            Some(f) => sheet
                .write_datetime_with_format(row, col, &edt, f)
                .map(|_| ()),
            None => sheet.write_datetime(row, col, &edt).map(|_| ()),
        })
    }

    // write_rich_string writes a cell containing multiple text fragments
    // with different formats (bold, italic, colors, etc.). The `parts`
    // argument is a list of (text, format_or_none) tuples:
    //
    //   ws.write_rich_string(0, 0, [
    //       ("Hello, ", None),
    //       ("bold part", bold_fmt),
    //       (" and back to normal", None),
    //   ])
    //
    // At least one fragment must carry a non-None format, or Excel will
    // store the result as plain text. rust_xlsxwriter raises an XlsxError
    // if given an empty list or all-None formats, which surfaces here as
    // ValueError.
    //
    // Lifetime strategy: rather than fighting the borrow checker trying to
    // hold &RustFormat refs from PyRef borrows alongside the workbook
    // borrow in with_sheet(), we materialise owned (String, RustFormat)
    // pairs first, then build the &[(&str, &RustFormat)] slice over that
    // owned Vec inside the closure. The clone cost is negligible -- rich
    // strings are short by definition (Excel's own cell limit applies).
    #[pyo3(signature = (row, col, parts, format=None))]
    fn write_rich_string(
        &self,
        py: Python<'_>,
        row: u32,
        col: u16,
        parts: Vec<(String, Option<Py<Format>>)>,
        format: Option<&Format>,
    ) -> PyResult<()> {
        self.check_row_order(row)?;
        if parts.is_empty() {
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "write_rich_string(): parts list must not be empty",
            ));
        }
        // Build owned (text, format_clone) pairs. Cloning RustFormat is
        // cheap (it's a pure value type) and lets us avoid the lifetime
        // entanglement of holding PyRef borrows across the with_sheet call.
        let owned: Vec<(RustFormat, String)> = parts
            .into_iter()
            .map(|(text, fmt_py)| {
                let fmt_owned = match fmt_py {
                    Some(f) => f.borrow(py).inner.clone(),
                    None => RustFormat::new(),
                };
                (fmt_owned, text)
            })
            .collect();

        // Build the &[(&RustFormat, &str)] slice that rust_xlsxwriter expects.
        // Note: upstream API is (&Format, &str) -- format first, text second.
        let rich_parts: Vec<(&RustFormat, &str)> = owned
            .iter()
            .map(|(fmt, text)| (fmt, text.as_str()))
            .collect();

        let cell_fmt = format.map(|f| &f.inner);
        self.with_sheet(py, |sheet| match cell_fmt {
            Some(f) => sheet
                .write_rich_string_with_format(row, col, &rich_parts, f)
                .map(|_| ()),
            None => sheet.write_rich_string(row, col, &rich_parts).map(|_| ()),
        })
    }

    fn insert_image(&self, py: Python<'_>, row: u32, col: u16, image_path: &str) -> PyResult<()> {
        self.check_row_order(row)?;
        // Pre-check existence before handing to rust_xlsxwriter so the
        // error message includes the full path (the crate's own IoError
        // message may not).
        if !std::path::Path::new(image_path).exists() {
            return Err(PyErr::new::<pyo3::exceptions::PyOSError, _>(format!(
                "insert_image(): file not found: '{image_path}'"
            )));
        }
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

    // Adds an Excel worksheet table over first_row..=last_row,
    // first_col..=last_col. Expects the data in that range to already
    // have been written via write()/write_row()/write_records()/etc --
    // add_table() itself only adds the table structure (headers, total
    // row, banding, style), it doesn't write cell values.
    //
    // Deliberately NOT going through check_row_order()/
    // check_row_order_range() (the constant_memory row-order guard):
    // rust_xlsxwriter's own usage pattern for tables is to write data
    // first (advancing the row high-water mark past the table's rows),
    // *then* call add_table() over that already-written range --
    // meaning first_row here is routinely a row at or before ones
    // already written, which is correct and expected, not a backward
    // write. Applying the cell-write guard here would incorrectly
    // reject that normal usage.
    fn add_table(
        &self,
        py: Python<'_>,
        first_row: u32,
        first_col: u16,
        last_row: u32,
        last_col: u16,
        table: &Table,
    ) -> PyResult<()> {
        // add_table() writes real cells immediately, not only at save time:
        // upstream's implementation calls write_string_with_format() for each
        // header caption (worksheet.rs, add_table) and likewise for a total
        // row. It was therefore the one cell-writing method on this class
        // without a constant_memory row-order check, which left exactly the
        // silent-corruption path check_row_order() exists to close -- an
        // add_table() anchored above the current high-water mark would be
        // accepted here and quietly produce a damaged .xlsx.
        self.check_row_order_range(first_row, last_row)?;
        let t = table.inner.clone();
        self.with_sheet(py, |sheet| {
            sheet
                .add_table(first_row, first_col, last_row, last_col, &t)
                .map(|_| ())
        })
    }
}

// ============================================
// WORKBOOK CLASS
// ============================================
// subclass=true so the Python wrapper in python/rvgsrust_xlsxwriter/
// __init__.py (which extends this to add __enter__/__exit__ context
// manager support) can inherit from it. Without this flag PyO3
// generates an unsubclassable native type and the Python `class
// Workbook(_CoreWorkbook):` line raises TypeError at import time.
#[pyclass(subclass)]
struct Workbook {
    inner: RefCell<RustWorkbook>,
}

#[pymethods]
impl Workbook {
    #[new]
    fn new() -> Self {
        Workbook {
            inner: RefCell::new(RustWorkbook::new()),
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
            let mut wb = wb_ref
            .inner
            .try_borrow_mut()
            .map_err(|_| reentrant_workbook_err())?;

            // Validate the name BEFORE touching the workbook.
            //
            // set_name() runs exactly this validator internally
            // (utility::validate_sheetname, reached via the public
            // check_sheet_name wrapper), so the accepted/rejected set and the
            // resulting XlsxError variant are unchanged. What changes is the
            // ordering: previously the worksheet was appended first and
            // set_name()'s error propagated afterwards, which left an
            // unnamed worksheet in the workbook while the index counter had
            // not advanced. The next add_worksheet() then handed back an
            // index pointing at that orphan, so every subsequent write went
            // to the wrong sheet -- silently, and the intended sheet saved
            // empty. Validating first means a rejected name mutates nothing.
            if let Some(n) = name {
                rust_xlsxwriter::utility::check_sheet_name(n).map_err(xlsx_err_to_pyerr)?;
            }

            // Take the index from the workbook's own worksheet vector rather
            // than a counter maintained alongside it. The two cannot drift if
            // there is only one of them; the drift was the bug above.
            let idx = wb.worksheets().len();

            let sheet = if constant_memory {
                wb.add_worksheet_with_constant_memory()
            } else {
                wb.add_worksheet()
            };
            // Pre-validated above, so this cannot fail on name grounds; the
            // mapping is kept rather than unwrapped so a future upstream
            // validation change surfaces as an exception, not a panic.
            if let Some(n) = name {
                sheet.set_name(n).map_err(xlsx_err_to_pyerr)?;
            }
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

    // `py` is injected by pyo3 and is not part of the Python-visible
    // signature, so this remains `wb.close(path)` from Python.
    fn close(&self, py: Python<'_>, path: &str) -> PyResult<()> {
        let mut guard = self
            .inner
            .try_borrow_mut()
            .map_err(|_| reentrant_workbook_err())?;

        // save() is the single longest operation in the library -- it
        // serialises every worksheet and deflates the whole archive -- and it
        // touches no Python objects at all. Holding the GIL across it stalls
        // every other thread in the process for the entire duration, which for
        // a large workbook is seconds.
        //
        // Soundness of releasing it here:
        //  - Ungil is satisfied via Send (pyo3 0.22 marker.rs: `unsafe impl<T:
        //    Send> Ungil for T`). rust_xlsxwriter's types contain no Rc,
        //    RefCell, Cell or raw pointers anywhere in the crate, so Workbook
        //    is auto-Send and `&mut Workbook` is Send with it.
        //  - The RefMut guard is held across the release, which keeps this the
        //    only live mutable borrow. Another thread that acquires the GIL and
        //    calls into the same Workbook hits try_borrow_mut() and gets the
        //    RuntimeError above rather than a second &mut.
        //  - RefCell's borrow flag is not atomic, but every access to it still
        //    happens under the GIL: the guard is created before the release and
        //    dropped after the re-acquire. Nothing touches the flag while the
        //    GIL is released.
        let workbook: &mut RustWorkbook = &mut guard;
        py.allow_threads(move || workbook.save(path))
            .map_err(xlsx_err_to_pyerr)?;
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
            .try_borrow_mut()
            .map_err(|_| reentrant_workbook_err())?
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
    ) -> PyResult<()> {
        let mut wb = self
            .inner
            .try_borrow_mut()
            .map_err(|_| reentrant_workbook_err())?;
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
        Ok(())
    }
}

// ============================================
// MODULE INITIALIZATION
// ============================================
// ============================================
// CONDITIONAL FORMATTING
// ============================================
// The 12 rule types rust_xlsxwriter exposes, minus IconSet, which the
// parity audit tracks separately. Every upstream type here is a consuming
// builder (mut self -> Self) that derives Clone in 0.96, so each setter
// has the same shape: clone the inner value, run the builder method on
// the clone, store the result back.
//
// The whole upstream module is reached through one alias rather than ~20
// individual imports. That also sidesteps the pyclass name-shadowing
// problem: our pyclass is plain `ConditionalFormatCell`, upstream's is
// always written `rcf::ConditionalFormatCell`, so the two never collide
// inside the pymethods-generated submodule.
//
// Note these setters return None rather than self, so unlike Format they
// don't chain. Adding a return value later is backwards compatible, so
// this can be revisited without breaking callers.
use rust_xlsxwriter::conditional_format as rcf;

fn cf_type_err(kind: &str, got: &str, expected: &str) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(format!(
        "Unknown {kind} '{got}'. Expected one of: {expected}"
    ))
}

// Accepts a number or a string. Strings are what the Formula rule type
// needs; numbers cover every other rule type.
fn cf_value(value: &Bound<'_, PyAny>) -> PyResult<rcf::ConditionalFormatValue> {
    if let Ok(f) = value.extract::<f64>() {
        return Ok(f.into());
    }
    if let Ok(s) = value.extract::<String>() {
        return Ok(s.into());
    }
    Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(
        "conditional format value must be a number or a string",
    ))
}

fn parse_cf_type(name: &str) -> PyResult<rcf::ConditionalFormatType> {
    use rcf::ConditionalFormatType as T;
    match name.to_ascii_lowercase().as_str() {
        "automatic" => Ok(T::Automatic),
        "lowest" | "min" => Ok(T::Lowest),
        "highest" | "max" => Ok(T::Highest),
        "number" => Ok(T::Number),
        "percent" => Ok(T::Percent),
        "percentile" => Ok(T::Percentile),
        "formula" => Ok(T::Formula),
        other => Err(cf_type_err(
            "conditional format type",
            other,
            "automatic, lowest/min, highest/max, number, percent, percentile, formula",
        )),
    }
}

fn parse_average_rule(name: &str) -> PyResult<rcf::ConditionalFormatAverageRule> {
    use rcf::ConditionalFormatAverageRule as R;
    match name.to_ascii_lowercase().as_str() {
        "above" => Ok(R::AboveAverage),
        "below" => Ok(R::BelowAverage),
        "equal_or_above" => Ok(R::EqualOrAboveAverage),
        "equal_or_below" => Ok(R::EqualOrBelowAverage),
        "1_std_dev_above" => Ok(R::OneStandardDeviationAbove),
        "1_std_dev_below" => Ok(R::OneStandardDeviationBelow),
        "2_std_dev_above" => Ok(R::TwoStandardDeviationsAbove),
        "2_std_dev_below" => Ok(R::TwoStandardDeviationsBelow),
        "3_std_dev_above" => Ok(R::ThreeStandardDeviationsAbove),
        "3_std_dev_below" => Ok(R::ThreeStandardDeviationsBelow),
        other => Err(cf_type_err(
            "average rule",
            other,
            "above, below, equal_or_above, equal_or_below, \
             {1,2,3}_std_dev_above, {1,2,3}_std_dev_below",
        )),
    }
}

fn parse_date_rule(name: &str) -> PyResult<rcf::ConditionalFormatDateRule> {
    use rcf::ConditionalFormatDateRule as R;
    match name.to_ascii_lowercase().as_str() {
        "yesterday" => Ok(R::Yesterday),
        "today" => Ok(R::Today),
        "tomorrow" => Ok(R::Tomorrow),
        "last_7_days" => Ok(R::Last7Days),
        "last_week" => Ok(R::LastWeek),
        "this_week" => Ok(R::ThisWeek),
        "next_week" => Ok(R::NextWeek),
        "last_month" => Ok(R::LastMonth),
        "this_month" => Ok(R::ThisMonth),
        "next_month" => Ok(R::NextMonth),
        other => Err(cf_type_err(
            "date rule",
            other,
            "yesterday, today, tomorrow, last_7_days, last_week, this_week, \
             next_week, last_month, this_month, next_month",
        )),
    }
}

fn parse_text_rule(kind: &str, text: &str) -> PyResult<rcf::ConditionalFormatTextRule> {
    use rcf::ConditionalFormatTextRule as R;
    let owned = text.to_string();
    match kind.to_ascii_lowercase().as_str() {
        "contains" => Ok(R::Contains(owned)),
        "does_not_contain" => Ok(R::DoesNotContain(owned)),
        "begins_with" => Ok(R::BeginsWith(owned)),
        "ends_with" => Ok(R::EndsWith(owned)),
        other => Err(cf_type_err(
            "text rule",
            other,
            "contains, does_not_contain, begins_with, ends_with",
        )),
    }
}

fn parse_top_rule(kind: &str, value: u16) -> PyResult<rcf::ConditionalFormatTopRule> {
    use rcf::ConditionalFormatTopRule as R;
    match kind.to_ascii_lowercase().as_str() {
        "top" => Ok(R::Top(value)),
        "bottom" => Ok(R::Bottom(value)),
        "top_percent" => Ok(R::TopPercent(value)),
        "bottom_percent" => Ok(R::BottomPercent(value)),
        other => Err(cf_type_err(
            "top rule",
            other,
            "top, bottom, top_percent, bottom_percent",
        )),
    }
}

fn parse_bar_direction(name: &str) -> PyResult<rcf::ConditionalFormatDataBarDirection> {
    use rcf::ConditionalFormatDataBarDirection as D;
    match name.to_ascii_lowercase().as_str() {
        "context" => Ok(D::Context),
        "left_to_right" => Ok(D::LeftToRight),
        "right_to_left" => Ok(D::RightToLeft),
        other => Err(cf_type_err(
            "data bar direction",
            other,
            "context, left_to_right, right_to_left",
        )),
    }
}

fn parse_bar_axis(name: &str) -> PyResult<rcf::ConditionalFormatDataBarAxisPosition> {
    use rcf::ConditionalFormatDataBarAxisPosition as A;
    match name.to_ascii_lowercase().as_str() {
        "automatic" => Ok(A::Automatic),
        "midpoint" => Ok(A::Midpoint),
        "none" => Ok(A::None),
        other => Err(cf_type_err(
            "data bar axis position",
            other,
            "automatic, midpoint, none",
        )),
    }
}

// -------------------- Cell --------------------

#[pyclass]
struct ConditionalFormatCell {
    inner: rcf::ConditionalFormatCell,
}

#[pymethods]
impl ConditionalFormatCell {
    #[new]
    fn new() -> Self {
        ConditionalFormatCell {
            inner: rcf::ConditionalFormatCell::new(),
        }
    }

    fn set_rule_equal_to(&mut self, value: f64) {
        use rcf::ConditionalFormatCellRule as R;
        self.inner = self.inner.clone().set_rule(R::EqualTo(value));
    }

    fn set_rule_not_equal_to(&mut self, value: f64) {
        use rcf::ConditionalFormatCellRule as R;
        self.inner = self.inner.clone().set_rule(R::NotEqualTo(value));
    }

    fn set_rule_greater_than(&mut self, value: f64) {
        use rcf::ConditionalFormatCellRule as R;
        self.inner = self.inner.clone().set_rule(R::GreaterThan(value));
    }

    fn set_rule_greater_than_or_equal_to(&mut self, value: f64) {
        use rcf::ConditionalFormatCellRule as R;
        self.inner = self.inner.clone().set_rule(R::GreaterThanOrEqualTo(value));
    }

    fn set_rule_less_than(&mut self, value: f64) {
        use rcf::ConditionalFormatCellRule as R;
        self.inner = self.inner.clone().set_rule(R::LessThan(value));
    }

    fn set_rule_less_than_or_equal_to(&mut self, value: f64) {
        use rcf::ConditionalFormatCellRule as R;
        self.inner = self.inner.clone().set_rule(R::LessThanOrEqualTo(value));
    }

    fn set_rule_between(&mut self, minimum: f64, maximum: f64) {
        use rcf::ConditionalFormatCellRule as R;
        let rule = R::Between(minimum, maximum);
        self.inner = self.inner.clone().set_rule(rule);
    }

    fn set_rule_not_between(&mut self, minimum: f64, maximum: f64) {
        use rcf::ConditionalFormatCellRule as R;
        let rule = R::NotBetween(minimum, maximum);
        self.inner = self.inner.clone().set_rule(rule);
    }

    fn set_format(&mut self, format: &Format) {
        self.inner = self.inner.clone().set_format(format.inner.clone());
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- Blank --------------------

#[pyclass]
struct ConditionalFormatBlank {
    inner: rcf::ConditionalFormatBlank,
}

#[pymethods]
impl ConditionalFormatBlank {
    #[new]
    fn new() -> Self {
        ConditionalFormatBlank {
            inner: rcf::ConditionalFormatBlank::new(),
        }
    }

    fn invert(&mut self) {
        self.inner = self.inner.clone().invert();
    }

    fn set_format(&mut self, format: &Format) {
        self.inner = self.inner.clone().set_format(format.inner.clone());
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- Duplicate --------------------

#[pyclass]
struct ConditionalFormatDuplicate {
    inner: rcf::ConditionalFormatDuplicate,
}

#[pymethods]
impl ConditionalFormatDuplicate {
    #[new]
    fn new() -> Self {
        ConditionalFormatDuplicate {
            inner: rcf::ConditionalFormatDuplicate::new(),
        }
    }

    // invert() turns "highlight duplicates" into "highlight uniques".
    fn invert(&mut self) {
        self.inner = self.inner.clone().invert();
    }

    fn set_format(&mut self, format: &Format) {
        self.inner = self.inner.clone().set_format(format.inner.clone());
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- Error --------------------

#[pyclass]
struct ConditionalFormatError {
    inner: rcf::ConditionalFormatError,
}

#[pymethods]
impl ConditionalFormatError {
    #[new]
    fn new() -> Self {
        ConditionalFormatError {
            inner: rcf::ConditionalFormatError::new(),
        }
    }

    // invert() turns "highlight errors" into "highlight non-errors".
    fn invert(&mut self) {
        self.inner = self.inner.clone().invert();
    }

    fn set_format(&mut self, format: &Format) {
        self.inner = self.inner.clone().set_format(format.inner.clone());
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- Formula --------------------

#[pyclass]
struct ConditionalFormatFormula {
    inner: rcf::ConditionalFormatFormula,
}

#[pymethods]
impl ConditionalFormatFormula {
    #[new]
    fn new() -> Self {
        ConditionalFormatFormula {
            inner: rcf::ConditionalFormatFormula::new(),
        }
    }

    // Takes an Excel formula string such as "=$A1>50", relative to the
    // top-left cell of the range the format is applied to.
    fn set_rule(&mut self, formula: &str) {
        self.inner = self.inner.clone().set_rule(formula);
    }

    fn set_format(&mut self, format: &Format) {
        self.inner = self.inner.clone().set_format(format.inner.clone());
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- Average --------------------

#[pyclass]
struct ConditionalFormatAverage {
    inner: rcf::ConditionalFormatAverage,
}

#[pymethods]
impl ConditionalFormatAverage {
    #[new]
    fn new() -> Self {
        ConditionalFormatAverage {
            inner: rcf::ConditionalFormatAverage::new(),
        }
    }

    fn set_rule(&mut self, rule: &str) -> PyResult<()> {
        let parsed = parse_average_rule(rule)?;
        self.inner = self.inner.clone().set_rule(parsed);
        Ok(())
    }

    fn set_format(&mut self, format: &Format) {
        self.inner = self.inner.clone().set_format(format.inner.clone());
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- Top --------------------

#[pyclass]
struct ConditionalFormatTop {
    inner: rcf::ConditionalFormatTop,
}

#[pymethods]
impl ConditionalFormatTop {
    #[new]
    fn new() -> Self {
        ConditionalFormatTop {
            inner: rcf::ConditionalFormatTop::new(),
        }
    }

    // kind is one of top, bottom, top_percent, bottom_percent; value is
    // the N in "top N" or "top N percent".
    fn set_rule(&mut self, kind: &str, value: u16) -> PyResult<()> {
        let parsed = parse_top_rule(kind, value)?;
        self.inner = self.inner.clone().set_rule(parsed);
        Ok(())
    }

    fn set_format(&mut self, format: &Format) {
        self.inner = self.inner.clone().set_format(format.inner.clone());
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- Text --------------------

#[pyclass]
struct ConditionalFormatText {
    inner: rcf::ConditionalFormatText,
}

#[pymethods]
impl ConditionalFormatText {
    #[new]
    fn new() -> Self {
        ConditionalFormatText {
            inner: rcf::ConditionalFormatText::new(),
        }
    }

    // kind is one of contains, does_not_contain, begins_with, ends_with.
    fn set_rule(&mut self, kind: &str, text: &str) -> PyResult<()> {
        let parsed = parse_text_rule(kind, text)?;
        self.inner = self.inner.clone().set_rule(parsed);
        Ok(())
    }

    fn set_format(&mut self, format: &Format) {
        self.inner = self.inner.clone().set_format(format.inner.clone());
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- Date --------------------

#[pyclass]
struct ConditionalFormatDate {
    inner: rcf::ConditionalFormatDate,
}

#[pymethods]
impl ConditionalFormatDate {
    #[new]
    fn new() -> Self {
        ConditionalFormatDate {
            inner: rcf::ConditionalFormatDate::new(),
        }
    }

    fn set_rule(&mut self, rule: &str) -> PyResult<()> {
        let parsed = parse_date_rule(rule)?;
        self.inner = self.inner.clone().set_rule(parsed);
        Ok(())
    }

    fn set_format(&mut self, format: &Format) {
        self.inner = self.inner.clone().set_format(format.inner.clone());
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- 2 Color Scale --------------------
// Color scales and data bars have no set_format(): Excel renders them
// from the scale/bar definition itself rather than from a dxf record.

#[pyclass]
struct ConditionalFormat2ColorScale {
    inner: rcf::ConditionalFormat2ColorScale,
}

#[pymethods]
impl ConditionalFormat2ColorScale {
    #[new]
    fn new() -> Self {
        ConditionalFormat2ColorScale {
            inner: rcf::ConditionalFormat2ColorScale::new(),
        }
    }

    fn set_minimum(&mut self, rule_type: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let parsed = parse_cf_type(rule_type)?;
        let val = cf_value(value)?;
        self.inner = self.inner.clone().set_minimum(parsed, val);
        Ok(())
    }

    fn set_maximum(&mut self, rule_type: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let parsed = parse_cf_type(rule_type)?;
        let val = cf_value(value)?;
        self.inner = self.inner.clone().set_maximum(parsed, val);
        Ok(())
    }

    fn set_minimum_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_minimum_color(parsed);
        Ok(())
    }

    fn set_maximum_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_maximum_color(parsed);
        Ok(())
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- 3 Color Scale --------------------

#[pyclass]
struct ConditionalFormat3ColorScale {
    inner: rcf::ConditionalFormat3ColorScale,
}

#[pymethods]
impl ConditionalFormat3ColorScale {
    #[new]
    fn new() -> Self {
        ConditionalFormat3ColorScale {
            inner: rcf::ConditionalFormat3ColorScale::new(),
        }
    }

    fn set_minimum(&mut self, rule_type: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let parsed = parse_cf_type(rule_type)?;
        let val = cf_value(value)?;
        self.inner = self.inner.clone().set_minimum(parsed, val);
        Ok(())
    }

    fn set_midpoint(&mut self, rule_type: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let parsed = parse_cf_type(rule_type)?;
        let val = cf_value(value)?;
        self.inner = self.inner.clone().set_midpoint(parsed, val);
        Ok(())
    }

    fn set_maximum(&mut self, rule_type: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let parsed = parse_cf_type(rule_type)?;
        let val = cf_value(value)?;
        self.inner = self.inner.clone().set_maximum(parsed, val);
        Ok(())
    }

    fn set_minimum_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_minimum_color(parsed);
        Ok(())
    }

    fn set_midpoint_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_midpoint_color(parsed);
        Ok(())
    }

    fn set_maximum_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_maximum_color(parsed);
        Ok(())
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- Data Bar --------------------

#[pyclass]
struct ConditionalFormatDataBar {
    inner: rcf::ConditionalFormatDataBar,
}

#[pymethods]
impl ConditionalFormatDataBar {
    #[new]
    fn new() -> Self {
        ConditionalFormatDataBar {
            inner: rcf::ConditionalFormatDataBar::new(),
        }
    }

    fn set_minimum(&mut self, rule_type: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let parsed = parse_cf_type(rule_type)?;
        let val = cf_value(value)?;
        self.inner = self.inner.clone().set_minimum(parsed, val);
        Ok(())
    }

    fn set_maximum(&mut self, rule_type: &str, value: &Bound<'_, PyAny>) -> PyResult<()> {
        let parsed = parse_cf_type(rule_type)?;
        let val = cf_value(value)?;
        self.inner = self.inner.clone().set_maximum(parsed, val);
        Ok(())
    }

    fn set_fill_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_fill_color(parsed);
        Ok(())
    }

    fn set_border_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_border_color(parsed);
        Ok(())
    }

    fn set_negative_fill_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_negative_fill_color(parsed);
        Ok(())
    }

    fn set_negative_border_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_negative_border_color(parsed);
        Ok(())
    }

    fn set_axis_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_axis_color(parsed);
        Ok(())
    }

    fn set_solid_fill(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_solid_fill(enable);
    }

    fn set_border_off(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_border_off(enable);
    }

    fn set_bar_only(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_bar_only(enable);
    }

    // One of context, left_to_right, right_to_left.
    fn set_direction(&mut self, direction: &str) -> PyResult<()> {
        let parsed = parse_bar_direction(direction)?;
        self.inner = self.inner.clone().set_direction(parsed);
        Ok(())
    }

    // One of automatic, midpoint, none.
    fn set_axis_position(&mut self, position: &str) -> PyResult<()> {
        let parsed = parse_bar_axis(position)?;
        self.inner = self.inner.clone().set_axis_position(parsed);
        Ok(())
    }

    fn use_classic_style(&mut self) {
        self.inner = self.inner.clone().use_classic_style();
    }

    fn set_multi_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_multi_range(range);
    }

    fn set_stop_if_true(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_stop_if_true(enable);
    }
}

// -------------------- dispatch --------------------
// Worksheet::add_conditional_format is generic over the ConditionalFormat
// trait, and a #[pyclass] can't be generic, so the concrete type has to be
// recovered by trying each downcast in turn. Pulling the matched value out
// into this enum first keeps the downcast chain in a plain function (no
// `self`, so a local macro_rules is safe) and leaves the call site as one
// short match.
enum AnyCf {
    Cell(rcf::ConditionalFormatCell),
    Blank(rcf::ConditionalFormatBlank),
    Duplicate(rcf::ConditionalFormatDuplicate),
    ErrorCf(rcf::ConditionalFormatError),
    Formula(rcf::ConditionalFormatFormula),
    Average(rcf::ConditionalFormatAverage),
    Top(rcf::ConditionalFormatTop),
    Text(rcf::ConditionalFormatText),
    Date(rcf::ConditionalFormatDate),
    Scale2(rcf::ConditionalFormat2ColorScale),
    Scale3(rcf::ConditionalFormat3ColorScale),
    DataBar(rcf::ConditionalFormatDataBar),
}

macro_rules! try_downcast_cf {
    ($obj:expr, $py_ty:ty, $variant:ident) => {
        if let Ok(found) = $obj.downcast::<$py_ty>() {
            return Ok(AnyCf::$variant(found.borrow().inner.clone()));
        }
    };
}

fn extract_cf(cf: &Bound<'_, PyAny>) -> PyResult<AnyCf> {
    try_downcast_cf!(cf, ConditionalFormatCell, Cell);
    try_downcast_cf!(cf, ConditionalFormatBlank, Blank);
    try_downcast_cf!(cf, ConditionalFormatDuplicate, Duplicate);
    try_downcast_cf!(cf, ConditionalFormatError, ErrorCf);
    try_downcast_cf!(cf, ConditionalFormatFormula, Formula);
    try_downcast_cf!(cf, ConditionalFormatAverage, Average);
    try_downcast_cf!(cf, ConditionalFormatTop, Top);
    try_downcast_cf!(cf, ConditionalFormatText, Text);
    try_downcast_cf!(cf, ConditionalFormatDate, Date);
    try_downcast_cf!(cf, ConditionalFormat2ColorScale, Scale2);
    try_downcast_cf!(cf, ConditionalFormat3ColorScale, Scale3);
    try_downcast_cf!(cf, ConditionalFormatDataBar, DataBar);
    Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(
        "add_conditional_format() expects a ConditionalFormat* object",
    ))
}

// ============================================
// SPARKLINES
// ============================================
// One pyclass, unlike conditional formatting: upstream's add_sparkline and
// add_sparkline_group both take a concrete &Sparkline rather than being
// generic over a trait, so no downcast chain is needed.
//
// Same aliasing approach as the conditional formats above -- our pyclass is
// plain `Sparkline`, upstream's is always `rsl::Sparkline`, so the two
// cannot collide inside the pymethods-generated submodule. ChartEmptyCells
// lives in the chart module rather than the sparkline one.
use rust_xlsxwriter::chart as rch;
use rust_xlsxwriter::sparkline as rsl;

fn parse_sparkline_type(name: &str) -> PyResult<rsl::SparklineType> {
    use rsl::SparklineType as T;
    match name.to_ascii_lowercase().as_str() {
        "line" => Ok(T::Line),
        "column" => Ok(T::Column),
        // Upstream spells this WinLose, not WinLoss.
        "win_lose" | "win_loss" => Ok(T::WinLose),
        other => Err(cf_type_err(
            "sparkline type",
            other,
            "line, column, win_lose",
        )),
    }
}

fn parse_empty_cells(name: &str) -> PyResult<rch::ChartEmptyCells> {
    use rch::ChartEmptyCells as E;
    match name.to_ascii_lowercase().as_str() {
        // The variant is Gaps, not Gap.
        "gaps" => Ok(E::Gaps),
        "zero" => Ok(E::Zero),
        "connected" => Ok(E::Connected),
        other => Err(cf_type_err(
            "empty cells option",
            other,
            "gaps, zero, connected",
        )),
    }
}

#[pyclass]
struct Sparkline {
    inner: rsl::Sparkline,
}

#[pymethods]
impl Sparkline {
    #[new]
    fn new() -> Self {
        Sparkline {
            inner: rsl::Sparkline::new(),
        }
    }

    // Range of the data the sparkline plots, e.g. "Sheet1!A1:E1".
    fn set_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_range(range);
    }

    // Optional range of dates giving the sparkline a date axis.
    fn set_date_range(&mut self, range: &str) {
        self.inner = self.inner.clone().set_date_range(range);
    }

    // One of line, column, win_lose.
    fn set_type(&mut self, sparkline_type: &str) -> PyResult<()> {
        let parsed = parse_sparkline_type(sparkline_type)?;
        self.inner = self.inner.clone().set_type(parsed);
        Ok(())
    }

    fn show_high_point(&mut self, enable: bool) {
        self.inner = self.inner.clone().show_high_point(enable);
    }

    fn show_low_point(&mut self, enable: bool) {
        self.inner = self.inner.clone().show_low_point(enable);
    }

    fn show_first_point(&mut self, enable: bool) {
        self.inner = self.inner.clone().show_first_point(enable);
    }

    fn show_last_point(&mut self, enable: bool) {
        self.inner = self.inner.clone().show_last_point(enable);
    }

    fn show_negative_points(&mut self, enable: bool) {
        self.inner = self.inner.clone().show_negative_points(enable);
    }

    fn show_markers(&mut self, enable: bool) {
        self.inner = self.inner.clone().show_markers(enable);
    }

    fn show_axis(&mut self, enable: bool) {
        self.inner = self.inner.clone().show_axis(enable);
    }

    fn show_hidden_data(&mut self, enable: bool) {
        self.inner = self.inner.clone().show_hidden_data(enable);
    }

    // One of gaps, zero, connected.
    fn show_empty_cells_as(&mut self, option: &str) -> PyResult<()> {
        let parsed = parse_empty_cells(option)?;
        self.inner = self.inner.clone().show_empty_cells_as(parsed);
        Ok(())
    }

    fn set_right_to_left(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_right_to_left(enable);
    }

    fn set_column_order(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_column_order(enable);
    }

    fn set_sparkline_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_sparkline_color(parsed);
        Ok(())
    }

    fn set_high_point_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_high_point_color(parsed);
        Ok(())
    }

    fn set_low_point_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_low_point_color(parsed);
        Ok(())
    }

    fn set_first_point_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_first_point_color(parsed);
        Ok(())
    }

    fn set_last_point_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_last_point_color(parsed);
        Ok(())
    }

    fn set_negative_points_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_negative_points_color(parsed);
        Ok(())
    }

    fn set_markers_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner = self.inner.clone().set_markers_color(parsed);
        Ok(())
    }

    fn set_line_weight(&mut self, weight: f64) {
        self.inner = self.inner.clone().set_line_weight(weight);
    }

    fn set_custom_max(&mut self, max: f64) {
        self.inner = self.inner.clone().set_custom_max(max);
    }

    fn set_custom_min(&mut self, min: f64) {
        self.inner = self.inner.clone().set_custom_min(min);
    }

    // Scale the group's maximum to the largest value across the whole
    // group rather than per sparkline.
    fn set_group_max(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_group_max(enable);
    }

    fn set_group_min(&mut self, enable: bool) {
        self.inner = self.inner.clone().set_group_min(enable);
    }

    fn set_style(&mut self, style: u8) {
        self.inner = self.inner.clone().set_style(style);
    }
}

// ============================================
// CHARTS (part 1: Chart, ChartSeries, insert_chart)
// ============================================
// Reuses the `rch` alias introduced by the sparkline section above.
//
// Two upstream facts shape this API.
//
// First, ChartAxis, ChartTitle and ChartLegend have pub(crate)
// constructors, so they cannot exist as separate Python objects. They are
// reached through Chart::x_axis(), title() and legend(), which hand back
// &mut references, so the options are flattened onto Chart itself as
// set_x_axis_*, set_title_* and set_legend_* methods.
//
// Second, Chart does NOT derive Clone (ChartSeries does). That rules out
// the alternative of collecting series and pushing them at insert time:
// pushing would mutate the only copy of the chart, so inserting the same
// Chart twice would silently duplicate its series, and there is no
// remove_series to undo it. Series are therefore attached explicitly with
// chart.push_series(series), which maps 1:1 onto upstream and makes any
// duplication the caller's visible choice.
//
// Unlike the conditional formats and sparklines, these upstream setters
// take &mut self and return &mut Self rather than consuming self, so no
// clone-and-reassign dance is needed.

fn parse_chart_type(name: &str) -> PyResult<rch::ChartType> {
    use rch::ChartType as T;
    match name.to_ascii_lowercase().as_str() {
        "area" => Ok(T::Area),
        "area_stacked" => Ok(T::AreaStacked),
        "area_percent_stacked" => Ok(T::AreaPercentStacked),
        "bar" => Ok(T::Bar),
        "bar_stacked" => Ok(T::BarStacked),
        "bar_percent_stacked" => Ok(T::BarPercentStacked),
        "column" => Ok(T::Column),
        "column_stacked" => Ok(T::ColumnStacked),
        "column_percent_stacked" => Ok(T::ColumnPercentStacked),
        "doughnut" => Ok(T::Doughnut),
        "line" => Ok(T::Line),
        "line_stacked" => Ok(T::LineStacked),
        "line_percent_stacked" => Ok(T::LinePercentStacked),
        "pie" => Ok(T::Pie),
        "radar" => Ok(T::Radar),
        "radar_with_markers" => Ok(T::RadarWithMarkers),
        "radar_filled" => Ok(T::RadarFilled),
        "scatter" => Ok(T::Scatter),
        "scatter_straight" => Ok(T::ScatterStraight),
        "scatter_straight_with_markers" => Ok(T::ScatterStraightWithMarkers),
        "scatter_smooth" => Ok(T::ScatterSmooth),
        "scatter_smooth_with_markers" => Ok(T::ScatterSmoothWithMarkers),
        "stock" => Ok(T::Stock),
        other => Err(cf_type_err(
            "chart type",
            other,
            "area, area_stacked, area_percent_stacked, bar, bar_stacked, \
             bar_percent_stacked, column, column_stacked, \
             column_percent_stacked, doughnut, line, line_stacked, \
             line_percent_stacked, pie, radar, radar_with_markers, \
             radar_filled, scatter, scatter_straight, \
             scatter_straight_with_markers, scatter_smooth, \
             scatter_smooth_with_markers, stock",
        )),
    }
}

// Only five positions exist upstream. There is no OverlayRight or
// OverlayLeft: overlaying is a separate set_legend_overlay() toggle.
fn parse_legend_position(name: &str) -> PyResult<rch::ChartLegendPosition> {
    use rch::ChartLegendPosition as P;
    match name.to_ascii_lowercase().as_str() {
        "right" => Ok(P::Right),
        "left" => Ok(P::Left),
        "top" => Ok(P::Top),
        "bottom" => Ok(P::Bottom),
        "top_right" => Ok(P::TopRight),
        other => Err(cf_type_err(
            "legend position",
            other,
            "right, left, top, bottom, top_right",
        )),
    }
}

// -------------------- ChartSeries --------------------

#[pyclass]
struct ChartSeries {
    inner: rch::ChartSeries,
    // Whether the caller set these explicitly. Chart.push_series() needs to
    // know, so it can re-apply a chart-type default only where the caller
    // did not make a choice of their own.
    marker_set: bool,
    format_set: bool,
}

#[pymethods]
impl ChartSeries {
    #[new]
    fn new() -> Self {
        ChartSeries {
            inner: rch::ChartSeries::new(),
            marker_set: false,
            format_set: false,
        }
    }

    // Range holding the series values, e.g. "Sheet1!$B$1:$B$5".
    fn set_values(&mut self, range: &str) {
        self.inner.set_values(range);
    }

    // Range holding the category (x axis) labels.
    fn set_categories(&mut self, range: &str) {
        self.inner.set_categories(range);
    }

    // Either a literal name or a range reference holding one.
    fn set_name(&mut self, name: &str) {
        self.inner.set_name(name);
    }

    fn set_secondary_axis(&mut self, enable: bool) {
        self.inner.set_secondary_axis(enable);
    }

    fn set_overlap(&mut self, overlap: i8) {
        self.inner.set_overlap(overlap);
    }

    fn set_gap(&mut self, gap: u16) {
        self.inner.set_gap(gap);
    }

    fn set_smooth(&mut self, enable: bool) {
        self.inner.set_smooth(enable);
    }

    fn set_invert_if_negative(&mut self) {
        self.inner.set_invert_if_negative();
    }

    fn set_invert_if_negative_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner.set_invert_if_negative_color(parsed);
        Ok(())
    }

    fn delete_from_legend(&mut self, enable: bool) {
        self.inner.delete_from_legend(enable);
    }

    // One color per point, mainly useful for pie and doughnut charts.
    // Vec<String> rather than Vec<&str>: PyO3 0.22 cannot extract the
    // latter.
    fn set_point_colors(&mut self, colors: Vec<String>) -> PyResult<()> {
        let mut parsed = Vec::with_capacity(colors.len());
        for color in &colors {
            parsed.push(parse_color(color)?);
        }
        self.inner.set_point_colors(&parsed);
        Ok(())
    }

    // Generic over IntoChartFormat upstream, which is implemented for
    // &mut ChartFormat, so this needs an owned mutable clone.
    fn set_format(&mut self, format: &ChartFormat) {
        let mut fmt = format.inner.clone();
        self.inner.set_format(&mut fmt);
        self.format_set = true;
    }

    fn set_marker(&mut self, marker: &ChartMarker) {
        self.inner.set_marker(&marker.inner);
        self.marker_set = true;
    }

    fn set_trendline(&mut self, trendline: &ChartTrendline) {
        self.inner.set_trendline(&trendline.inner);
    }

    fn set_data_label(&mut self, data_label: &ChartDataLabel) {
        self.inner.set_data_label(&data_label.inner);
    }

    // One label per point. Labels meant to differ from the series default
    // should have had to_custom() called on them.
    fn set_custom_data_labels(&mut self, py: Python<'_>, labels: Vec<Py<ChartDataLabel>>) {
        let mut owned = Vec::with_capacity(labels.len());
        for label in &labels {
            let borrowed = label.borrow(py);
            owned.push(borrowed.inner.clone());
        }
        self.inner.set_custom_data_labels(&owned);
    }
}

// Chart types that can show point markers but default to none. This
// mirrors upstream's condition exactly: ScatterStraight, ScatterSmooth,
// Radar, and the three types whose chart_group_type is Line. Stock has its
// own group and is deliberately not included.
fn markers_off_by_default(chart_type: rch::ChartType) -> bool {
    use rch::ChartType as T;
    chart_type == T::ScatterStraight
        || chart_type == T::ScatterSmooth
        || chart_type == T::Line
        || chart_type == T::LineStacked
        || chart_type == T::LinePercentStacked
        || chart_type == T::Radar
}

// -------------------- Chart --------------------

#[pyclass]
struct Chart {
    inner: rch::Chart,
    // Needed by push_series to decide which chart-type defaults apply.
    // ChartType is Copy, so this is a cheap duplicate of what the inner
    // chart already knows but keeps private.
    chart_type: rch::ChartType,
}

#[pymethods]
impl Chart {
    #[new]
    fn new(chart_type: &str) -> PyResult<Self> {
        let parsed = parse_chart_type(chart_type)?;
        Ok(Chart {
            inner: rch::Chart::new(parsed),
            chart_type: parsed,
        })
    }

    // Appends a configured ChartSeries. Called once per series; calling it
    // twice with the same series adds it twice.
    //
    // Deliberately NOT upstream's push_series(). That applies the
    // chart-type defaults after copying the caller's series, so on a line,
    // radar or scatter-straight/smooth chart it overwrites any marker set
    // beforehand with a "none" marker -- silently discarding it. Upstream's
    // intended flow is add_series() then configure, where the defaults land
    // first and the caller's setters win.
    //
    // add_series() hands back a mutable slot with those defaults already
    // applied. Overwriting it with the caller's series restores the right
    // precedence; the two defaults are then re-applied only where the
    // caller made no choice, which is what push_series should have done.
    fn push_series(&mut self, series: &ChartSeries) {
        let chart_type = self.chart_type;
        let slot = self.inner.add_series();
        *slot = series.inner.clone();

        if !series.format_set && chart_type == rch::ChartType::Scatter {
            let mut line = rch::ChartLine::new();
            line.set_width(2.25);
            line.set_hidden(true);
            let mut fmt = rch::ChartFormat::new();
            fmt.set_line(&line);
            slot.set_format(&mut fmt);
        }

        if !series.marker_set && markers_off_by_default(chart_type) {
            let mut marker = rch::ChartMarker::new();
            marker.set_none();
            slot.set_marker(&marker);
        }
    }

    fn set_style(&mut self, style: u8) {
        self.inner.set_style(style);
    }

    fn set_width(&mut self, width: u32) {
        self.inner.set_width(width);
    }

    fn set_height(&mut self, height: u32) {
        self.inner.set_height(height);
    }

    fn set_name(&mut self, name: &str) {
        self.inner.set_name(name);
    }

    fn set_alt_text(&mut self, alt_text: &str) {
        self.inner.set_alt_text(alt_text);
    }

    // Doughnut hole size as a percentage, and pie/doughnut start angle.
    fn set_hole_size(&mut self, hole_size: u8) {
        self.inner.set_hole_size(hole_size);
    }

    fn set_rotation(&mut self, rotation: u16) {
        self.inner.set_rotation(rotation);
    }

    fn show_hidden_data(&mut self) {
        self.inner.show_hidden_data();
    }

    fn show_na_as_empty_cell(&mut self) {
        self.inner.show_na_as_empty_cell();
    }

    // One of gaps, zero, connected.
    fn show_empty_cells_as(&mut self, option: &str) -> PyResult<()> {
        let parsed = parse_empty_cells(option)?;
        self.inner.show_empty_cells_as(parsed);
        Ok(())
    }

    // ---- title ----
    // ChartTitle::new() is pub(crate), so these route through title().

    fn set_title_name(&mut self, name: &str) {
        self.inner.title().set_name(name);
    }

    // Upstream's set_hidden() takes no argument.
    fn set_title_hidden(&mut self) {
        self.inner.title().set_hidden();
    }

    fn set_title_overlay(&mut self, enable: bool) {
        self.inner.title().set_overlay(enable);
    }

    // ---- x axis ----

    fn set_x_axis_name(&mut self, name: &str) {
        self.inner.x_axis().set_name(name);
    }

    fn set_x_axis_min(&mut self, min: f64) {
        self.inner.x_axis().set_min(min);
    }

    fn set_x_axis_max(&mut self, max: f64) {
        self.inner.x_axis().set_max(max);
    }

    fn set_x_axis_major_unit(&mut self, value: f64) {
        self.inner.x_axis().set_major_unit(value);
    }

    fn set_x_axis_minor_unit(&mut self, value: f64) {
        self.inner.x_axis().set_minor_unit(value);
    }

    fn set_x_axis_log_base(&mut self, base: u16) {
        self.inner.x_axis().set_log_base(base);
    }

    fn set_x_axis_num_format(&mut self, num_format: &str) {
        self.inner.x_axis().set_num_format(num_format);
    }

    fn set_x_axis_hidden(&mut self, enable: bool) {
        self.inner.x_axis().set_hidden(enable);
    }

    // Upstream's set_reverse() takes no argument.
    fn set_x_axis_reverse(&mut self) {
        self.inner.x_axis().set_reverse();
    }

    fn set_x_axis_major_gridlines(&mut self, enable: bool) {
        self.inner.x_axis().set_major_gridlines(enable);
    }

    fn set_x_axis_minor_gridlines(&mut self, enable: bool) {
        self.inner.x_axis().set_minor_gridlines(enable);
    }

    fn set_x_axis_date_axis(&mut self, enable: bool) {
        self.inner.x_axis().set_date_axis(enable);
    }

    fn set_x_axis_text_axis(&mut self, enable: bool) {
        self.inner.x_axis().set_text_axis(enable);
    }

    // ---- y axis ----

    fn set_y_axis_name(&mut self, name: &str) {
        self.inner.y_axis().set_name(name);
    }

    fn set_y_axis_min(&mut self, min: f64) {
        self.inner.y_axis().set_min(min);
    }

    fn set_y_axis_max(&mut self, max: f64) {
        self.inner.y_axis().set_max(max);
    }

    fn set_y_axis_major_unit(&mut self, value: f64) {
        self.inner.y_axis().set_major_unit(value);
    }

    fn set_y_axis_minor_unit(&mut self, value: f64) {
        self.inner.y_axis().set_minor_unit(value);
    }

    fn set_y_axis_log_base(&mut self, base: u16) {
        self.inner.y_axis().set_log_base(base);
    }

    fn set_y_axis_num_format(&mut self, num_format: &str) {
        self.inner.y_axis().set_num_format(num_format);
    }

    fn set_y_axis_hidden(&mut self, enable: bool) {
        self.inner.y_axis().set_hidden(enable);
    }

    fn set_y_axis_reverse(&mut self) {
        self.inner.y_axis().set_reverse();
    }

    fn set_y_axis_major_gridlines(&mut self, enable: bool) {
        self.inner.y_axis().set_major_gridlines(enable);
    }

    fn set_y_axis_minor_gridlines(&mut self, enable: bool) {
        self.inner.y_axis().set_minor_gridlines(enable);
    }

    // ---- legend ----

    // One of right, left, top, bottom, top_right.
    fn set_legend_position(&mut self, position: &str) -> PyResult<()> {
        let parsed = parse_legend_position(position)?;
        self.inner.legend().set_position(parsed);
        Ok(())
    }

    // Upstream's set_hidden() takes no argument.
    fn set_legend_hidden(&mut self) {
        self.inner.legend().set_hidden();
    }

    fn set_legend_overlay(&mut self, enable: bool) {
        self.inner.legend().set_overlay(enable);
    }

    // ---- fonts and formats ----
    // set_font takes a plain &ChartFont, but set_format is generic over
    // IntoChartFormat, which upstream implements for &mut ChartFormat.
    // Hence the owned mutable clone rather than passing &format.inner.

    fn set_title_font(&mut self, font: &ChartFont) {
        self.inner.title().set_font(&font.inner);
    }

    fn set_title_format(&mut self, format: &ChartFormat) {
        let mut fmt = format.inner.clone();
        self.inner.title().set_format(&mut fmt);
    }

    fn set_x_axis_font(&mut self, font: &ChartFont) {
        self.inner.x_axis().set_font(&font.inner);
    }

    fn set_x_axis_name_font(&mut self, font: &ChartFont) {
        self.inner.x_axis().set_name_font(&font.inner);
    }

    fn set_x_axis_format(&mut self, format: &ChartFormat) {
        let mut fmt = format.inner.clone();
        self.inner.x_axis().set_format(&mut fmt);
    }

    fn set_x_axis_name_format(&mut self, format: &ChartFormat) {
        let mut fmt = format.inner.clone();
        self.inner.x_axis().set_name_format(&mut fmt);
    }

    fn set_y_axis_font(&mut self, font: &ChartFont) {
        self.inner.y_axis().set_font(&font.inner);
    }

    fn set_y_axis_name_font(&mut self, font: &ChartFont) {
        self.inner.y_axis().set_name_font(&font.inner);
    }

    fn set_y_axis_format(&mut self, format: &ChartFormat) {
        let mut fmt = format.inner.clone();
        self.inner.y_axis().set_format(&mut fmt);
    }

    fn set_y_axis_name_format(&mut self, format: &ChartFormat) {
        let mut fmt = format.inner.clone();
        self.inner.y_axis().set_name_format(&mut fmt);
    }

    fn set_legend_font(&mut self, font: &ChartFont) {
        self.inner.legend().set_font(&font.inner);
    }

    fn set_legend_format(&mut self, format: &ChartFormat) {
        let mut fmt = format.inner.clone();
        self.inner.legend().set_format(&mut fmt);
    }
}

// ============================================
// CHARTS (part 2: ChartFormat, ChartFont)
// ============================================
// ChartLine and ChartSolidFill are deliberately NOT exposed as separate
// pyclasses. Upstream they exist only to be handed to
// ChartFormat::set_line/set_border/set_solid_fill, so they are flattened
// into ChartFormat as set_line_*, set_border_* and set_fill_* methods,
// the same way axes were flattened onto Chart. Their state is kept in the
// pyclass and re-applied on every setter, so calls compose.
//
// Pattern and gradient fills are not exposed yet; they are logged for the
// parity audit.
//
// Note IntoChartFormat never needs importing here even though it isn't
// re-exported from the crate root: it appears only as a bound on a
// generic parameter, and Rust requires a trait in scope only to call its
// methods. What matters is that upstream implements it for
// `&mut ChartFormat`, hence the owned mutable clone at each call site.

fn parse_dash_type(name: &str) -> PyResult<rch::ChartLineDashType> {
    use rch::ChartLineDashType as D;
    match name.to_ascii_lowercase().as_str() {
        "solid" => Ok(D::Solid),
        "round_dot" => Ok(D::RoundDot),
        "square_dot" => Ok(D::SquareDot),
        "dash" => Ok(D::Dash),
        "dash_dot" => Ok(D::DashDot),
        "long_dash" => Ok(D::LongDash),
        "long_dash_dot" => Ok(D::LongDashDot),
        "long_dash_dot_dot" => Ok(D::LongDashDotDot),
        other => Err(cf_type_err(
            "line dash type",
            other,
            "solid, round_dot, square_dot, dash, dash_dot, long_dash, \
             long_dash_dot, long_dash_dot_dot",
        )),
    }
}

// -------------------- ChartFont --------------------

#[pyclass]
struct ChartFont {
    inner: rch::ChartFont,
}

#[pymethods]
impl ChartFont {
    #[new]
    fn new() -> Self {
        ChartFont {
            inner: rch::ChartFont::new(),
        }
    }

    // set_bold/italic/underline/strikethrough take no argument upstream.
    // unset_bold() is the way back, and set_default_bold(false) suppresses
    // the bold that some chart elements apply by default.
    fn set_bold(&mut self) {
        self.inner.set_bold();
    }

    fn unset_bold(&mut self) {
        self.inner.unset_bold();
    }

    fn set_default_bold(&mut self, enable: bool) {
        self.inner.set_default_bold(enable);
    }

    fn set_italic(&mut self) {
        self.inner.set_italic();
    }

    fn set_underline(&mut self) {
        self.inner.set_underline();
    }

    fn set_strikethrough(&mut self) {
        self.inner.set_strikethrough();
    }

    fn set_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.inner.set_color(parsed);
        Ok(())
    }

    fn set_name(&mut self, font_name: &str) {
        self.inner.set_name(font_name);
    }

    fn set_size(&mut self, font_size: f64) {
        self.inner.set_size(font_size);
    }

    // Degrees, -90 to 90, or 270 to 360 for stacked text.
    fn set_rotation(&mut self, rotation: i16) {
        self.inner.set_rotation(rotation);
    }

    fn set_right_to_left(&mut self, enable: bool) {
        self.inner.set_right_to_left(enable);
    }

    fn set_pitch_family(&mut self, family: u8) {
        self.inner.set_pitch_family(family);
    }

    fn set_character_set(&mut self, character_set: u8) {
        self.inner.set_character_set(character_set);
    }
}

// -------------------- ChartFormat --------------------

#[pyclass]
struct ChartFormat {
    inner: rch::ChartFormat,
    line: rch::ChartLine,
    border: rch::ChartLine,
    fill: rch::ChartSolidFill,
}

#[pymethods]
impl ChartFormat {
    // ChartLine and ChartSolidFill must be built with ::new(), not
    // ::default().
    #[new]
    fn new() -> Self {
        ChartFormat {
            inner: rch::ChartFormat::new(),
            line: rch::ChartLine::new(),
            border: rch::ChartLine::new(),
            fill: rch::ChartSolidFill::new(),
        }
    }

    // ---- line ----

    fn set_line_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.line.set_color(parsed);
        let line = self.line.clone();
        self.inner.set_line(&line);
        Ok(())
    }

    fn set_line_width(&mut self, width: f64) {
        self.line.set_width(width);
        let line = self.line.clone();
        self.inner.set_line(&line);
    }

    fn set_line_dash_type(&mut self, dash_type: &str) -> PyResult<()> {
        let parsed = parse_dash_type(dash_type)?;
        self.line.set_dash_type(parsed);
        let line = self.line.clone();
        self.inner.set_line(&line);
        Ok(())
    }

    fn set_line_transparency(&mut self, transparency: u8) {
        self.line.set_transparency(transparency);
        let line = self.line.clone();
        self.inner.set_line(&line);
    }

    fn set_line_hidden(&mut self, enable: bool) {
        self.line.set_hidden(enable);
        let line = self.line.clone();
        self.inner.set_line(&line);
    }

    fn set_no_line(&mut self) {
        self.inner.set_no_line();
    }

    // ---- border ----
    // Same underlying ChartLine type as set_line_*, but border is the name
    // Excel uses for the outline of a filled shape.

    fn set_border_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.border.set_color(parsed);
        let border = self.border.clone();
        self.inner.set_border(&border);
        Ok(())
    }

    fn set_border_width(&mut self, width: f64) {
        self.border.set_width(width);
        let border = self.border.clone();
        self.inner.set_border(&border);
    }

    fn set_border_dash_type(&mut self, dash_type: &str) -> PyResult<()> {
        let parsed = parse_dash_type(dash_type)?;
        self.border.set_dash_type(parsed);
        let border = self.border.clone();
        self.inner.set_border(&border);
        Ok(())
    }

    fn set_border_transparency(&mut self, transparency: u8) {
        self.border.set_transparency(transparency);
        let border = self.border.clone();
        self.inner.set_border(&border);
    }

    fn set_border_hidden(&mut self, enable: bool) {
        self.border.set_hidden(enable);
        let border = self.border.clone();
        self.inner.set_border(&border);
    }

    fn set_no_border(&mut self) {
        self.inner.set_no_border();
    }

    // ---- fill ----

    fn set_fill_color(&mut self, color: &str) -> PyResult<()> {
        let parsed = parse_color(color)?;
        self.fill.set_color(parsed);
        let fill = self.fill.clone();
        self.inner.set_solid_fill(&fill);
        Ok(())
    }

    fn set_fill_transparency(&mut self, transparency: u8) {
        self.fill.set_transparency(transparency);
        let fill = self.fill.clone();
        self.inner.set_solid_fill(&fill);
    }

    fn set_no_fill(&mut self) {
        self.inner.set_no_fill();
    }
}

// ============================================
// CHARTS (part 3: ChartMarker, ChartTrendline, ChartDataLabel)
// ============================================
// These three consume part 2: each takes a ChartFormat and/or ChartFont,
// and all three attach to a ChartSeries.
//
// Upstream naming is irregular here and the differences are load-bearing:
//   - ChartMarkerType has no Automatic or None variant. Those are the
//     separate set_automatic() and set_none() methods.
//   - ChartTrendlineType::Logarithmic, not Logarithm. Polynomial and
//     MovingAverage carry a u8 period, so set_type takes one.
//   - ChartTrendline spells its toggles display_equation() and
//     display_r_squared(), with no set_ prefix. They are exposed here as
//     set_display_equation / set_display_r_squared so the Python surface
//     stays internally consistent, but they call the unprefixed names.

fn parse_marker_type(name: &str) -> PyResult<rch::ChartMarkerType> {
    use rch::ChartMarkerType as M;
    match name.to_ascii_lowercase().as_str() {
        "square" => Ok(M::Square),
        "diamond" => Ok(M::Diamond),
        "triangle" => Ok(M::Triangle),
        "x" => Ok(M::X),
        "star" => Ok(M::Star),
        "short_dash" => Ok(M::ShortDash),
        "long_dash" => Ok(M::LongDash),
        "circle" => Ok(M::Circle),
        "plus_sign" => Ok(M::PlusSign),
        other => Err(cf_type_err(
            "marker type",
            other,
            "square, diamond, triangle, x, star, short_dash, long_dash, \
             circle, plus_sign (for automatic or no marker use \
             set_automatic() or set_none())",
        )),
    }
}

// period applies only to polynomial and moving_average; the other kinds
// ignore it.
fn parse_trendline_type(kind: &str, period: u8) -> PyResult<rch::ChartTrendlineType> {
    use rch::ChartTrendlineType as T;
    match kind.to_ascii_lowercase().as_str() {
        "none" => Ok(T::None),
        "linear" => Ok(T::Linear),
        "exponential" => Ok(T::Exponential),
        // Logarithmic, not Logarithm.
        "logarithmic" => Ok(T::Logarithmic),
        "power" => Ok(T::Power),
        "polynomial" => Ok(T::Polynomial(period)),
        "moving_average" => Ok(T::MovingAverage(period)),
        other => Err(cf_type_err(
            "trendline type",
            other,
            "none, linear, exponential, logarithmic, power, polynomial, \
             moving_average",
        )),
    }
}

fn parse_label_position(name: &str) -> PyResult<rch::ChartDataLabelPosition> {
    use rch::ChartDataLabelPosition as P;
    match name.to_ascii_lowercase().as_str() {
        "default" => Ok(P::Default),
        "center" => Ok(P::Center),
        "right" => Ok(P::Right),
        "left" => Ok(P::Left),
        "above" => Ok(P::Above),
        "below" => Ok(P::Below),
        "inside_base" => Ok(P::InsideBase),
        "inside_end" => Ok(P::InsideEnd),
        "outside_end" => Ok(P::OutsideEnd),
        "best_fit" => Ok(P::BestFit),
        other => Err(cf_type_err(
            "data label position",
            other,
            "default, center, right, left, above, below, inside_base, \
             inside_end, outside_end, best_fit",
        )),
    }
}

// -------------------- ChartMarker --------------------

#[pyclass]
struct ChartMarker {
    inner: rch::ChartMarker,
}

#[pymethods]
impl ChartMarker {
    #[new]
    fn new() -> Self {
        ChartMarker {
            inner: rch::ChartMarker::new(),
        }
    }

    // Let Excel pick the marker shape.
    fn set_automatic(&mut self) {
        self.inner.set_automatic();
    }

    // Suppress the marker entirely.
    fn set_none(&mut self) {
        self.inner.set_none();
    }

    fn set_type(&mut self, marker_type: &str) -> PyResult<()> {
        let parsed = parse_marker_type(marker_type)?;
        self.inner.set_type(parsed);
        Ok(())
    }

    fn set_size(&mut self, size: u8) {
        self.inner.set_size(size);
    }

    fn set_format(&mut self, format: &ChartFormat) {
        let mut fmt = format.inner.clone();
        self.inner.set_format(&mut fmt);
    }
}

// -------------------- ChartTrendline --------------------

#[pyclass]
struct ChartTrendline {
    inner: rch::ChartTrendline,
}

#[pymethods]
impl ChartTrendline {
    #[new]
    fn new() -> Self {
        ChartTrendline {
            inner: rch::ChartTrendline::new(),
        }
    }

    // period is the order for polynomial and the window for
    // moving_average; other kinds ignore it.
    #[pyo3(signature = (trend_type, period=2))]
    fn set_type(&mut self, trend_type: &str, period: u8) -> PyResult<()> {
        let parsed = parse_trendline_type(trend_type, period)?;
        self.inner.set_type(parsed);
        Ok(())
    }

    fn set_name(&mut self, name: &str) {
        self.inner.set_name(name);
    }

    fn set_forward_period(&mut self, period: f64) {
        self.inner.set_forward_period(period);
    }

    fn set_backward_period(&mut self, period: f64) {
        self.inner.set_backward_period(period);
    }

    // Upstream spells these without a set_ prefix.
    fn set_display_equation(&mut self, enable: bool) {
        self.inner.display_equation(enable);
    }

    fn set_display_r_squared(&mut self, enable: bool) {
        self.inner.display_r_squared(enable);
    }

    fn set_intercept(&mut self, intercept: f64) {
        self.inner.set_intercept(intercept);
    }

    fn delete_from_legend(&mut self, enable: bool) {
        self.inner.delete_from_legend(enable);
    }

    fn set_format(&mut self, format: &ChartFormat) {
        let mut fmt = format.inner.clone();
        self.inner.set_format(&mut fmt);
    }

    fn set_label_format(&mut self, format: &ChartFormat) {
        let mut fmt = format.inner.clone();
        self.inner.set_label_format(&mut fmt);
    }

    fn set_label_font(&mut self, font: &ChartFont) {
        self.inner.set_label_font(&font.inner);
    }
}

// -------------------- ChartDataLabel --------------------

#[pyclass]
struct ChartDataLabel {
    inner: rch::ChartDataLabel,
}

#[pymethods]
impl ChartDataLabel {
    #[new]
    fn new() -> Self {
        ChartDataLabel {
            inner: rch::ChartDataLabel::new(),
        }
    }

    // The show_* toggles take no argument upstream.
    fn show_value(&mut self) {
        self.inner.show_value();
    }

    fn show_category_name(&mut self) {
        self.inner.show_category_name();
    }

    fn show_series_name(&mut self) {
        self.inner.show_series_name();
    }

    fn show_leader_lines(&mut self) {
        self.inner.show_leader_lines();
    }

    fn show_legend_key(&mut self) {
        self.inner.show_legend_key();
    }

    fn show_percentage(&mut self) {
        self.inner.show_percentage();
    }

    fn show_x_value(&mut self) {
        self.inner.show_x_value();
    }

    fn show_y_value(&mut self) {
        self.inner.show_y_value();
    }

    fn set_hidden(&mut self) {
        self.inner.set_hidden();
    }

    fn set_position(&mut self, position: &str) -> PyResult<()> {
        let parsed = parse_label_position(position)?;
        self.inner.set_position(parsed);
        Ok(())
    }

    fn set_num_format(&mut self, num_format: &str) {
        self.inner.set_num_format(num_format);
    }

    // Upstream takes a char, so anything that isn't exactly one character
    // is rejected. Matching on both nexts at once rather than checking
    // is_none() and then unwrapping, which clippy::unnecessary_unwrap
    // rejects and CI promotes to an error.
    fn set_separator(&mut self, separator: &str) -> PyResult<()> {
        let mut chars = separator.chars();
        match (chars.next(), chars.next()) {
            (Some(first), None) => {
                self.inner.set_separator(first);
                Ok(())
            }
            _ => {
                // Message kept short deliberately: at 16 spaces of indent a
                // longer one pushes the statement past max_width, and which
                // way rustfmt then breaks it is not something this container
                // can verify without cargo fmt.
                let message = format!("separator must be one character, got {separator:?}");
                Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(message))
            }
        }
    }

    // A literal string, or a range reference holding the label text.
    fn set_value(&mut self, value: &str) {
        self.inner.set_value(value);
    }

    // Marks this label as a custom one, for set_custom_data_labels().
    // Named set_custom rather than to_custom, which is upstream's name:
    // clippy::wrong_self_convention requires a to_* method to take &self
    // or self, never &mut self, and it is a default-on style lint that CI
    // promotes to an error. set_custom also matches the rest of this
    // binding's naming.
    fn set_custom(&mut self) {
        self.inner = self.inner.to_custom();
    }

    fn set_format(&mut self, format: &ChartFormat) {
        let mut fmt = format.inner.clone();
        self.inner.set_format(&mut fmt);
    }

    fn set_font(&mut self, font: &ChartFont) {
        self.inner.set_font(&font.inner);
    }
}

#[pymodule]
fn _core(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Workbook>()?;
    m.add_class::<Worksheet>()?;
    m.add_class::<Format>()?;
    m.add_class::<Table>()?;
    m.add_class::<TableColumn>()?;
    m.add_class::<ConditionalFormatCell>()?;
    m.add_class::<ConditionalFormatBlank>()?;
    m.add_class::<ConditionalFormatDuplicate>()?;
    m.add_class::<ConditionalFormatError>()?;
    m.add_class::<ConditionalFormatFormula>()?;
    m.add_class::<ConditionalFormatAverage>()?;
    m.add_class::<ConditionalFormatTop>()?;
    m.add_class::<ConditionalFormatText>()?;
    m.add_class::<ConditionalFormatDate>()?;
    m.add_class::<ConditionalFormat2ColorScale>()?;
    m.add_class::<ConditionalFormat3ColorScale>()?;
    m.add_class::<ConditionalFormatDataBar>()?;
    m.add_class::<Sparkline>()?;
    m.add_class::<Chart>()?;
    m.add_class::<ChartSeries>()?;
    m.add_class::<ChartFont>()?;
    m.add_class::<ChartFormat>()?;
    m.add_class::<ChartMarker>()?;
    m.add_class::<ChartTrendline>()?;
    m.add_class::<ChartDataLabel>()?;
    Ok(())
}
