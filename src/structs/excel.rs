#[allow(unused_imports)]
use log::{debug, info, warn};
use pyo3::prelude::*;

use crate::utils::excel::{excel_to_index, index_to_excel, index_to_excel_col};
use crate::structs::Mode;
use umya_spreadsheet::structs::Worksheet;

#[derive(Debug, Clone)]
pub enum ExcelCell {
    Tuple((u32, u32)), // (row, col)
    String(String),
}

// Implement FromPyObject for SourceRange
impl<'source> pyo3::FromPyObject<'source> for ExcelCell {
    fn extract(obj: &'source PyAny) -> PyResult<Self> {
        if let Ok(tuple) = obj.extract::<(u32, u32)>() {
            Ok(ExcelCell::Tuple(tuple))
        } else if let Ok(string) = obj.extract::<String>() {
            Ok(ExcelCell::String(string))
        } else {
            Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "Invalid input type. Expected a tuple of 2 tuples or a string.",
            ))
        }
    }
}

// Implement conversion from string to tuple of tuple
// The should define a range as in Excel A1:B2
impl ExcelCell {
    pub fn idx(&self) -> (u32, u32) {
        match self {
            ExcelCell::Tuple(t) => (t.1, t.0),
            ExcelCell::String(s) => excel_to_index(s),
        }
    }
    pub fn range(&self) -> String {
        match self {
            ExcelCell::Tuple(t) => index_to_excel(t.1, t.0),
            ExcelCell::String(s) => s.clone(),
        }
    }
}

// Implement default for ExcelCell
impl Default for ExcelCell {
    fn default() -> Self {
        ExcelCell::Tuple((1, 1))
    }
}


#[derive(Debug, Clone)]
pub enum ExcelRange {
    Range(((u32, u32), (u32, u32))), // (row, col)
    String(String),
}

// Implement FromPyObject for SourceRange
impl<'source> pyo3::FromPyObject<'source> for ExcelRange {
    fn extract(obj: &'source PyAny) -> PyResult<Self> {
        if let Ok(tuple) = obj.extract::<((u32, u32), (u32, u32))>() {
            Ok(ExcelRange::Range(tuple))
        } else if let Ok(string) = obj.extract::<String>() {
            Ok(ExcelRange::String(string))
        } else {
            Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "Invalid input type. Expected a tuple of 2 tuples or a string.",
            ))
        }
    }
}

// Implement conversion from string to tuple of tuple
// The should define a range as in Excel A1:B2
impl ExcelRange {
    pub fn idx(&self) -> ((u32, u32), (u32, u32)) {
        match self {
            ExcelRange::Range(r) => ((r.0.1, r.0.0), (r.1.1, r.1.0)),
            ExcelRange::String(s) => {
                let (col1, row1) = excel_to_index(s.split(':').next().unwrap());
                let (col2, row2) = excel_to_index(s.split(':').nth(1).unwrap());
                ((col1, row1), (col2, row2))
            }
        }
    }
    pub fn range(&self) -> String {
        match self {
            ExcelRange::Range(r) => format!("{}:{}", index_to_excel(r.0.1, r.0.0), index_to_excel(r.1.1, r.1.0)),
            ExcelRange::String(s) => s.clone(),
        }
    }
}


#[derive(Debug,Clone)]
pub enum ExcelHeader {
    First,
    Last,
    ExcelCell(ExcelCell),
    ExcelRange(ExcelRange),
}

// Implement FromPyObject for SourceRange
impl<'source> pyo3::FromPyObject<'source> for ExcelHeader {
    fn extract(obj: &'source PyAny) -> PyResult<Self> {
        if let Ok(range) = obj.extract::<ExcelRange>() {
            if range.idx().0.1 == range.idx().1.1 || range.idx().0.0 == range.idx().1.0 {
                Ok(ExcelHeader::ExcelRange(range))
            } else {
                Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("Invalid range. Expected a row or column range."))
            }
        } 
        else if let Ok(cell) = obj.extract::<ExcelCell>() {
            Ok(ExcelHeader::ExcelCell(cell))
        }
        else if let Ok(string) = obj.extract::<String>() {
            match string.as_str() {
                "first" => Ok(ExcelHeader::First),
                "last" => Ok(ExcelHeader::Last),
                _ => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("Invalid string identifier. Use 'first' or 'last'.")),
            }
        } 
        else {
            Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "Invalid input type. Expected a first, last, or a cell position (e.g. B5 or (row: 5, col: 2)).",
            ))
        }
    }
}

// Implement conversion from string to tuple of tuple
// The should define a range as in Excel A1:B2
impl ExcelHeader {
    pub fn idx(&self, worksheet: &Worksheet, mode: Mode) -> PyResult<(u32, u32)> {
        match self {
            ExcelHeader::ExcelRange(r) => Ok(r.idx().0),
            ExcelHeader::ExcelCell(c) => Ok(c.idx()),
            ExcelHeader::First => Ok((1, 1)),
            ExcelHeader::Last => {
                match mode {
                    Mode::Row => Ok((worksheet.get_highest_column(), 1)),
                    Mode::Column => Ok((1, worksheet.get_highest_row())),
                }
            }
        }
    }
    pub fn range(&self, worksheet: &Worksheet, mode: Mode) -> String {
        match self {
            ExcelHeader::ExcelRange(r) => r.range(),
            ExcelHeader::ExcelCell(c) => c.range(),
            ExcelHeader::First => "A1".to_string(),
            ExcelHeader::Last => {
                match mode {
                    Mode::Row => format!("A{}", worksheet.get_highest_row()),
                    Mode::Column => format!("{}1", index_to_excel_col(worksheet.get_highest_column())),
                }
            }
        }
    }
}