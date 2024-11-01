use pyo3::prelude::*;
use once_cell::sync::Lazy;
use umya_spreadsheet::*;
use std::path::Path;
use std::sync::Arc;
use pyo3::AsPyPointer;

use crate::template::ExcelTemplate;


#[pymethods]
impl ExcelTemplate {
    pub fn goto_sheet(&mut self, sheet_name: &str, current_cell: Option<(u32, u32)>) -> &Self {
        self.current_sheet_name = Some(sheet_name.to_string());
        self.current_cell_in_current_sheet = current_cell;
        self
    }

    pub fn goto_cell(&mut self, current_cell: Option<(u32, u32)>) -> &Self {
        self.current_cell_in_current_sheet = current_cell;
        self
    }

    pub fn set_value(&mut self, value: &str) -> PyResult<()> {
        // Store the current sheet name and cell in variables to avoid mutable borrow conflict
        let sheet_name = self.current_sheet_name
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No sheet specified"))?
            .to_string(); // Clone the string to avoid borrowing self

        let current_cell = self.current_cell_in_current_sheet
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No cell specified"))?;
                
        // Now call write_cell with the extracted values
        self.write_cell(&sheet_name, current_cell, value)
    }
}