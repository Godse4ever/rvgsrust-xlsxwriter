use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList, PyTuple, PyString, PyFloat, PyInt, PyBool};
use rust_xlsxwriter::{Workbook as RustWorkbook, Worksheet as RustWorksheet, Format as RustFormat, FormatBorder, FormatAlign, FormatPattern, Color};
use std::collections::HashMap;
use std::sync::Arc;
use std::cell::RefCell;

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
            return Color::RGB(r, g, b);
        }
    }
    // Named colors fallback
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

// ============================================
// FORMAT CLASS
// ============================================
#[pyclass]
struct Format {
    inner: RustFormat,
}

#[pymethods]
impl Format {
    #[new]
    fn new() -> Self {
        Format {
            inner: RustFormat::new(),
        }
    }

    fn set_bold(&mut self) -> PyResult<()> {
        self.inner = self.inner.clone().set_bold();
        Ok(())
    }

    fn set_italic(&mut self) -> PyResult<()> {
        self.inner = self.inner.clone().set_italic();
        Ok(())
    }

    fn set_underline(&mut self) -> PyResult<()> {
        self.inner = self.inner.clone().set_underline(rust_xlsxwriter::FormatUnderline::Single);
        Ok(())
    }

    fn set_font_name(&mut self, name: &str) -> PyResult<()> {
        self.inner = self.inner.clone().set_font_name(name);
        Ok(())
    }

    fn set_font_size(&mut self, size: f64) -> PyResult<()> {
        self.inner = self.inner.clone().set_font_size(size);
        Ok(())
    }

    fn set_font_color(&mut self, color: &str) -> PyResult<()> {
        self.inner = self.inner.clone().set_font_color(parse_color(color));
        Ok(())
    }

    fn set_background_color(&mut self, color: &str) -> PyResult<()> {
        self.inner = self.inner.clone().set_background_color(parse_color(color));
        Ok(())
    }

    fn set_border(&mut self, border: &str) -> PyResult<()> {
        let border_style = match border.to_lowercase().as_str() {
            "thin" => FormatBorder::Thin,
            "medium" => FormatBorder::Medium,
            "thick" => FormatBorder::Thick,
            "dashed" => FormatBorder::Dashed,
            "dotted" => FormatBorder::Dotted,
            "double" => FormatBorder::Double,
            "hair" => FormatBorder::Hair,
            _ => FormatBorder::Thin,
        };
        self.inner = self.inner.clone().set_border(border_style);
        Ok(())
    }

    fn set_border_color(&mut self, color: &str) -> PyResult<()> {
        let c = parse_color(color);
        self.inner = self.inner.clone()
            .set_border_color(c)
            .set_border_top_color(c)
            .set_border_bottom_color(c)
            .set_border_left_color(c)
            .set_border_right_color(c);
        Ok(())
    }

    fn set_top_border(&mut self, border: &str) -> PyResult<()> {
        let border_style = match border.to_lowercase().as_str() {
            "thin" => FormatBorder::Thin,
            "medium" => FormatBorder::Medium,
            "thick" => FormatBorder::Thick,
            "dashed" => FormatBorder::Dashed,
            "dotted" => FormatBorder::Dotted,
            "double" => FormatBorder::Double,
            _ => FormatBorder::Thin,
        };
        self.inner = self.inner.clone().set_border_top(border_style);
        Ok(())
    }

    fn set_bottom_border(&mut self, border: &str) -> PyResult<()> {
        let border_style = match border.to_lowercase().as_str() {
            "thin" => FormatBorder::Thin,
            "medium" => FormatBorder::Medium,
            "thick" => FormatBorder::Thick,
            "dashed" => FormatBorder::Dashed,
            "dotted" => FormatBorder::Dotted,
            "double" => FormatBorder::Double,
            _ => FormatBorder::Thin,
        };
        self.inner = self.inner.clone().set_border_bottom(border_style);
        Ok(())
    }

    fn set_left_border(&mut self, border: &str) -> PyResult<()> {
        let border_style = match border.to_lowercase().as_str() {
            "thin" => FormatBorder::Thin,
            "medium" => FormatBorder::Medium,
            "thick" => FormatBorder::Thick,
            "dashed" => FormatBorder::Dashed,
            "dotted" => FormatBorder::Dotted,
            "double" => FormatBorder::Double,
            _ => FormatBorder::Thin,
        };
        self.inner = self.inner.clone().set_border_left(border_style);
        Ok(())
    }

    fn set_right_border(&mut self, border: &str) -> PyResult<()> {
        let border_style = match border.to_lowercase().as_str() {
            "thin" => FormatBorder::Thin,
            "medium" => FormatBorder::Medium,
            "thick" => FormatBorder::Thick,
            "dashed" => FormatBorder::Dashed,
            "dotted" => FormatBorder::Dotted,
            "double" => FormatBorder::Double,
            _ => FormatBorder::Thin,
        };
        self.inner = self.inner.clone().set_border_right(border_style);
        Ok(())
    }

    fn set_num_format(&mut self, format: &str) -> PyResult<()> {
        self.inner = self.inner.clone().set_num_format(format);
        Ok(())
    }

    fn set_align(&mut self, align: &str) -> PyResult<()> {
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
        self.inner = self.inner.clone().set_align(alignment);
        Ok(())
    }

    fn set_vertical_align(&mut self, align: &str) -> PyResult<()> {
        let alignment = match align.to_lowercase().as_str() {
            "top" => FormatAlign::VerticalTop,
            "vcenter" | "center" => FormatAlign::VerticalCenter,
            "bottom" => FormatAlign::VerticalBottom,
            "vdistributed" | "distributed" => FormatAlign::VerticalDistributed,
            "vjustify" | "justify" => FormatAlign::VerticalJustify,
            _ => FormatAlign::VerticalCenter,
        };
        self.inner = self.inner.clone().set_align(alignment);
        Ok(())
    }

    fn set_text_wrap(&mut self) -> PyResult<()> {
        self.inner = self.inner.clone().set_text_wrap();
        Ok(())
    }

    fn set_shrink(&mut self) -> PyResult<()> {
        self.inner = self.inner.clone().set_shrink();
        Ok(())
    }

    fn set_rotation(&mut self, rotation: i16) -> PyResult<()> {
        self.inner = self.inner.clone().set_rotation(rotation);
        Ok(())
    }

    fn set_indent(&mut self, indent: u8) -> PyResult<()> {
        self.inner = self.inner.clone().set_indent(indent);
        Ok(())
    }

    fn set_pattern(&mut self, pattern: &str) -> PyResult<()> {
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
        self.inner = self.inner.clone().set_pattern(pat);
        Ok(())
    }
}

// ============================================
// WORKSHEET CLASS
// ============================================
#[pyclass]
struct Worksheet {
    inner: RefCell<RustWorksheet>,
}

#[pymethods]
impl Worksheet {
    fn write(&self, row: u32, col: u16, value: &Bound<'_, PyAny>, format: Option<&Format>) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();

        let fmt = format.map(|f| f.inner.clone());

        if value.is_none() {
            if let Some(f) = fmt {
                let _ = sheet.write_blank(row, col, &f);
            }
        } else if let Ok(s) = value.extract::<String>() {
            if let Some(f) = fmt {
                let _ = sheet.write_string(row, col, &s, &f);
            } else {
                let _ = sheet.write_string(row, col, &s, &RustFormat::new());
            }
        } else if let Ok(f) = value.extract::<f64>() {
            if let Some(fmt) = fmt {
                let _ = sheet.write_number(row, col, f, &fmt);
            } else {
                let _ = sheet.write_number(row, col, f, &RustFormat::new());
            }
        } else if let Ok(i) = value.extract::<i64>() {
            if let Some(fmt) = fmt {
                let _ = sheet.write_number(row, col, i as f64, &fmt);
            } else {
                let _ = sheet.write_number(row, col, i as f64, &RustFormat::new());
            }
        } else if let Ok(b) = value.extract::<bool>() {
            if let Some(fmt) = fmt {
                let _ = sheet.write_boolean(row, col, b, &fmt);
            } else {
                let _ = sheet.write_boolean(row, col, b, &RustFormat::new());
            }
        } else {
            // Fallback: convert to string
            let s = value.str()?.unwrap_or("").to_string();
            if let Some(f) = fmt {
                let _ = sheet.write_string(row, col, &s, &f);
            } else {
                let _ = sheet.write_string(row, col, &s, &RustFormat::new());
            }
        }

        Ok(())
    }

    fn write_row(&self, row: u32, col: u16, values: &Bound<'_, PyList>, format: Option<&Format>) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        let fmt = format.map(|f| f.inner.clone());

        for (i, value) in values.iter().enumerate() {
            let c = col + i as u16;

            if value.is_none() {
                if let Some(ref f) = fmt {
                    let _ = sheet.write_blank(row, c, f);
                }
            } else if let Ok(s) = value.extract::<String>() {
                if let Some(ref f) = fmt {
                    let _ = sheet.write_string(row, c, &s, f);
                } else {
                    let _ = sheet.write_string(row, c, &s, &RustFormat::new());
                }
            } else if let Ok(f) = value.extract::<f64>() {
                if let Some(ref fmt) = fmt {
                    let _ = sheet.write_number(row, c, f, fmt);
                } else {
                    let _ = sheet.write_number(row, c, f, &RustFormat::new());
                }
            } else if let Ok(i) = value.extract::<i64>() {
                if let Some(ref fmt) = fmt {
                    let _ = sheet.write_number(row, c, i as f64, fmt);
                } else {
                    let _ = sheet.write_number(row, c, i as f64, &RustFormat::new());
                }
            } else if let Ok(b) = value.extract::<bool>() {
                if let Some(ref fmt) = fmt {
                    let _ = sheet.write_boolean(row, c, b, fmt);
                } else {
                    let _ = sheet.write_boolean(row, c, b, &RustFormat::new());
                }
            } else {
                let s = value.str()?.unwrap_or("").to_string();
                if let Some(ref f) = fmt {
                    let _ = sheet.write_string(row, c, &s, f);
                } else {
                    let _ = sheet.write_string(row, c, &s, &RustFormat::new());
                }
            }
        }

        Ok(())
    }

    fn write_column(&self, row: u32, col: u16, values: &Bound<'_, PyList>, format: Option<&Format>) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        let fmt = format.map(|f| f.inner.clone());

        for (i, value) in values.iter().enumerate() {
            let r = row + i as u32;

            if value.is_none() {
                if let Some(ref f) = fmt {
                    let _ = sheet.write_blank(r, col, f);
                }
            } else if let Ok(s) = value.extract::<String>() {
                if let Some(ref f) = fmt {
                    let _ = sheet.write_string(r, col, &s, f);
                } else {
                    let _ = sheet.write_string(r, col, &s, &RustFormat::new());
                }
            } else if let Ok(f) = value.extract::<f64>() {
                if let Some(ref fmt) = fmt {
                    let _ = sheet.write_number(r, col, f, fmt);
                } else {
                    let _ = sheet.write_number(r, col, f, &RustFormat::new());
                }
            } else if let Ok(i) = value.extract::<i64>() {
                if let Some(ref fmt) = fmt {
                    let _ = sheet.write_number(r, col, i as f64, fmt);
                } else {
                    let _ = sheet.write_number(r, col, i as f64, &RustFormat::new());
                }
            } else if let Ok(b) = value.extract::<bool>() {
                if let Some(ref fmt) = fmt {
                    let _ = sheet.write_boolean(r, col, b, fmt);
                } else {
                    let _ = sheet.write_boolean(r, col, b, &RustFormat::new());
                }
            } else {
                let s = value.str()?.unwrap_or("").to_string();
                if let Some(ref f) = fmt {
                    let _ = sheet.write_string(r, col, &s, f);
                } else {
                    let _ = sheet.write_string(r, col, &s, &RustFormat::new());
                }
            }
        }

        Ok(())
    }

    fn merge_range(&self, first_row: u32, first_col: u16, last_row: u32, last_col: u16, value: &Bound<'_, PyAny>, format: Option<&Format>) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        let fmt = format.map(|f| f.inner.clone()).unwrap_or_else(RustFormat::new);

        if value.is_none() {
            let _ = sheet.merge_range(first_row, first_col, last_row, last_col, "", &fmt);
        } else if let Ok(s) = value.extract::<String>() {
            let _ = sheet.merge_range(first_row, first_col, last_row, last_col, &s, &fmt);
        } else if let Ok(f) = value.extract::<f64>() {
            let _ = sheet.merge_range(first_row, first_col, last_row, last_col, f, &fmt);
        } else if let Ok(i) = value.extract::<i64>() {
            let _ = sheet.merge_range(first_row, first_col, last_row, last_col, i as f64, &fmt);
        } else if let Ok(b) = value.extract::<bool>() {
            let _ = sheet.merge_range(first_row, first_col, last_row, last_col, b, &fmt);
        } else {
            let s = value.str()?.unwrap_or("").to_string();
            let _ = sheet.merge_range(first_row, first_col, last_row, last_col, &s, &fmt);
        }

        Ok(())
    }

    fn set_column_width(&self, col: u16, width: f64) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        let _ = sheet.set_column_width(col, width);
        Ok(())
    }

    fn set_row_height(&self, row: u32, height: f64) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        let _ = sheet.set_row_height(row, height);
        Ok(())
    }

    fn freeze_panes(&self, row: u32, col: u16) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        sheet.freeze_panes(row, col);
        Ok(())
    }

    fn set_tab_color(&self, color: &str) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        sheet.set_tab_color(parse_color(color));
        Ok(())
    }

    fn hide(&self) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        sheet.hide();
        Ok(())
    }

    fn protect(&self, password: &str) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        sheet.protect(password);
        Ok(())
    }

    fn write_formula(&self, row: u32, col: u16, formula: &str, format: Option<&Format>) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        let fmt = format.map(|f| f.inner.clone()).unwrap_or_else(RustFormat::new);
        let _ = sheet.write_formula(row, col, formula, &fmt);
        Ok(())
    }

    fn write_url(&self, row: u32, col: u16, url: &str, format: Option<&Format>) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        let fmt = format.map(|f| f.inner.clone()).unwrap_or_else(RustFormat::new);
        let _ = sheet.write_url(row, col, url, &fmt);
        Ok(())
    }

    fn write_datetime(&self, row: u32, col: u16, year: i32, month: u32, day: u32, hour: u32, min: u32, sec: u32, format: Option<&Format>) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        let fmt = format.map(|f| f.inner.clone()).unwrap_or_else(RustFormat::new);
        let dt = rust_xlsxwriter::ExcelDateTime::from_ymd(year, month, day)?
            .and_hms(hour, min, sec)?;
        let _ = sheet.write_datetime(row, col, &dt, &fmt);
        Ok(())
    }

    fn write_date(&self, row: u32, col: u16, year: i32, month: u32, day: u32, format: Option<&Format>) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        let fmt = format.map(|f| f.inner.clone()).unwrap_or_else(RustFormat::new);
        let dt = rust_xlsxwriter::ExcelDateTime::from_ymd(year, month, day)?;
        let _ = sheet.write_datetime(row, col, &dt, &fmt);
        Ok(())
    }

    fn insert_image(&self, row: u32, col: u16, image_path: &str) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        let image = rust_xlsxwriter::Image::new(image_path)?;
        let _ = sheet.insert_image(row, col, &image);
        Ok(())
    }

    fn autofit(&self) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        sheet.autofit();
        Ok(())
    }

    fn set_name(&self, name: &str) -> PyResult<()> {
        let mut sheet = self.inner.borrow_mut();
        sheet.set_name(name);
        Ok(())
    }
}

// ============================================
// WORKBOOK CLASS
// ============================================
#[pyclass]
struct Workbook {
    inner: RefCell<RustWorkbook>,
    worksheets: RefCell<Vec<Py<Worksheet>>>,
}

#[pymethods]
impl Workbook {
    #[new]
    fn new() -> Self {
        Workbook {
            inner: RefCell::new(RustWorkbook::new()),
            worksheets: RefCell::new(Vec::new()),
        }
    }

    fn add_worksheet(&self, py: Python, name: Option<&str>) -> PyResult<Py<Worksheet>> {
        let mut wb = self.inner.borrow_mut();
        let sheet = match name {
            Some(n) => wb.add_worksheet().set_name(n),
            None => wb.add_worksheet(),
        };

        let ws = Worksheet {
            inner: RefCell::new(sheet.clone()),
        };

        let py_ws = Py::new(py, ws)?;
        self.worksheets.borrow_mut().push(py_ws.clone_ref(py));

        Ok(py_ws)
    }

    fn add_format(&self, py: Python) -> PyResult<Py<Format>> {
        let fmt = Format::new();
        Py::new(py, fmt)
    }

    fn close(&self, path: &str) -> PyResult<()> {
        let mut wb = self.inner.borrow_mut();
        wb.save(path).map_err(|e| PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to save workbook: {}", e)))?;
        Ok(())
    }

    fn set_properties(&self, title: Option<&str>, author: Option<&str>, subject: Option<&str>, keywords: Option<&str>, comments: Option<&str>) -> PyResult<()> {
        let mut wb = self.inner.borrow_mut();
        let mut props = rust_xlsxwriter::WorkbookProperties::new();
        if let Some(t) = title { props = props.set_title(t); }
        if let Some(a) = author { props = props.set_author(a); }
        if let Some(s) = subject { props = props.set_subject(s); }
        if let Some(k) = keywords { props = props.set_keywords(k); }
        if let Some(c) = comments { props = props.set_comments(c); }
        wb.set_properties(&props);
        Ok(())
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
