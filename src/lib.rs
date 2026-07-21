use pyo3::prelude::*;
use pyo3::types::PyList;
use pyo3::PyRefMut;
use rust_xlsxwriter::{
    Color, FormatAlign, FormatBorder, FormatPattern,
    Workbook as RustWorkbook, Worksheet as RustWorksheet, Format as RustFormat,
};
use std::cell::RefCell;

// ============================================
// ERROR HELPER
// ============================================
// Converts any rust_xlsxwriter error (or other Display error) into a
// proper Python exception instead of being silently discarded, which
// is what the previous version did everywhere via `let _ = ...`.
fn to_pyerr<E: std::fmt::Display>(e: E) -> PyErr {
    PyErr::new::<pyo3::exceptions::PyValueError, _>(e.to_string())
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
        Ok(CellValue::Bool(b))
    } else if let Ok(s) = value.extract::<String>() {
        Ok(CellValue::Str(s))
    } else if let Ok(f) = value.extract::<f64>() {
        Ok(CellValue::Num(f))
    } else if let Ok(i) = value.extract::<i64>() {
        Ok(CellValue::Num(i as f64))
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

// merge_range in rust_xlsxwriter only accepts a `&str`, so non-string
// values are formatted to their Excel-visible text representation.
// (Note: if you later bump to a rust_xlsxwriter version where
// `merge_range` becomes generic over `IntoExcelData`, this can be
// widened to preserve real numeric/boolean cell types.)
fn cell_value_to_string(cv: &CellValue) -> String {
    match cv {
        CellValue::Blank => String::new(),
        CellValue::Str(s) => s.clone(),
        CellValue::Num(n) => {
            if n.fract() == 0.0 {
                format!("{}", *n as i64)
            } else {
                n.to_string()
            }
        }
        CellValue::Bool(b) => b.to_string(),
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
    let text = cell_value_to_string(cv);
    sheet
        .merge_range(first_row, first_col, last_row, last_col, &text, fmt)
        .map(|_| ())
}

// ============================================
// FORMAT CLASS
// ============================================
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
}

impl Worksheet {
    fn with_sheet<F, R>(&self, py: Python<'_>, f: F) -> PyResult<R>
    where
        F: FnOnce(&mut RustWorksheet) -> Result<R, rust_xlsxwriter::XlsxError>,
    {
        let wb_ref = self.workbook.borrow(py);
        let mut wb = wb_ref.inner.borrow_mut();
        let sheet = wb.worksheet_from_index(self.index).map_err(to_pyerr)?;
        f(sheet).map_err(to_pyerr)
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
        let fmt = format.map(|f| f.inner.clone());
        self.with_sheet(py, |sheet| write_value(sheet, row, col, &cv, fmt.as_ref()))
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
        let fmt = format.map(|f| f.inner.clone());
        let mut classified = Vec::with_capacity(values.len());
        for v in values.iter() {
            classified.push(classify(&v)?);
        }
        self.with_sheet(py, |sheet| {
            for (i, cv) in classified.iter().enumerate() {
                write_value(sheet, row, col + i as u16, cv, fmt.as_ref())?;
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
        let fmt = format.map(|f| f.inner.clone());
        let mut classified = Vec::with_capacity(values.len());
        for v in values.iter() {
            classified.push(classify(&v)?);
        }
        self.with_sheet(py, |sheet| {
            for (i, cv) in classified.iter().enumerate() {
                write_value(sheet, row + i as u32, col, cv, fmt.as_ref())?;
            }
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
        let fmt = format.map(|f| f.inner.clone()).unwrap_or_else(RustFormat::new);
        self.with_sheet(py, |sheet| {
            merge_value(sheet, first_row, first_col, last_row, last_col, &cv, &fmt)
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
        let fmt = format.map(|f| f.inner.clone());
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
        let fmt = format.map(|f| f.inner.clone());
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
        let fmt = format.map(|f| f.inner.clone());
        let dt = rust_xlsxwriter::ExcelDateTime::from_ymd(year, month, day)
            .map_err(to_pyerr)?
            .and_hms(hour, min, sec)
            .map_err(to_pyerr)?;
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
        let fmt = format.map(|f| f.inner.clone());
        let dt = rust_xlsxwriter::ExcelDateTime::from_ymd(year, month, day).map_err(to_pyerr)?;
        self.with_sheet(py, |sheet| match &fmt {
            Some(f) => sheet.write_datetime_with_format(row, col, &dt, f).map(|_| ()),
            None => sheet.write_datetime(row, col, &dt).map(|_| ()),
        })
    }

    fn insert_image(&self, py: Python<'_>, row: u32, col: u16, image_path: &str) -> PyResult<()> {
        let image = rust_xlsxwriter::Image::new(image_path).map_err(to_pyerr)?;
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

    #[pyo3(signature = (name=None))]
    fn add_worksheet(slf: Py<Self>, py: Python<'_>, name: Option<&str>) -> PyResult<Py<Worksheet>> {
        let index = {
            let wb_ref = slf.borrow(py);
            let mut wb = wb_ref.inner.borrow_mut();
            let sheet = wb.add_worksheet();
            if let Some(n) = name {
                sheet.set_name(n).map_err(to_pyerr)?;
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
            },
        )
    }

    fn add_format(&self, py: Python<'_>) -> PyResult<Py<Format>> {
        Py::new(py, Format::new())
    }

    fn close(&self, path: &str) -> PyResult<()> {
        self.inner.borrow_mut().save(path).map_err(to_pyerr)?;
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
