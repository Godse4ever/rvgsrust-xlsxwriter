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
enum CellValue {
    Blank,
    Str(String),
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
        (CellValue::Str(s), Some(f)) => sheet
            .write_string_with_format(row, col, s.as_str(), f)
            .map(|_| ()),
        (CellValue::Str(s), None) => sheet.write_string(row, col, s.as_str()).map(|_| ()),
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
    cv: &CellValue,
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

fn arrow_cell_value(col: &ArrowColumn<'_>, row: usize) -> CellValue {
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
        ArrowColumn::Utf8View(a) => {
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
        let mut wb = wb_ref.inner.borrow_mut();
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
        let mut classified_rows: Vec<(Vec<CellValue>, bool)> = Vec::with_capacity(n_rows);

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
        let mut wb = wb_ref.inner.borrow_mut();
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
        let mut wb = wb_ref.inner.borrow_mut();
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
        let mut wb = wb_ref.inner.borrow_mut();
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
                    &CellValue::Str(name.clone()),
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
        self.inner
            .borrow_mut()
            .save(path)
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
    m.add_class::<Table>()?;
    m.add_class::<TableColumn>()?;
    Ok(())
}
