#[allow(unused_imports)]
use log::{debug, info, warn};
use pyo3::prelude::*;
use pyo3::types::PyTuple;
use umya_spreadsheet::{self, reader, writer};

mod action;
mod mode;
mod aggregate;
mod paste;

use crate::transform::action::Action;
use crate::transform::mode::Mode;
use crate::transform::aggregate::aggregate_range;
use crate::transform::paste::paste_vector_to_sheet;

// Function to copy a range from one sheet to another
#[pyfunction]
pub fn transform_range_from_to(
    _py: Python,
    source_file_path: &str,
    source_sheet_name: &str,
    source_range: &PyTuple,
    dest_file_path: &str,
    dest_sheet_name: &str,
    dest_start_cell: &PyTuple,
    action: Action,
    mode: Mode,
    
) -> PyResult<()> {
    debug!("Copying range from {} to {}", source_file_path, dest_file_path);
    debug!("  ARGS:");
    debug!("        source_file_path: {:?}", source_file_path);
    debug!("        source_sheet_name: {:?}", source_sheet_name);
    debug!("        source_range: {:?}", source_range);
    debug!("        dest_file_path: {:?}", dest_file_path);
    debug!("        dest_sheet_name: {:?}", dest_sheet_name);
    debug!("        dest_start_cell: {:?}", dest_start_cell);
    debug!("        action: {:?}", action);
    debug!("        mode: {:?}", mode);
   // Extract values from the source_range tuple ((start_row, start_col), (end_row, end_col))
    let start_row = source_range.get_item(0)?.get_item(0)?.extract()?;
    let start_col = source_range.get_item(0)?.get_item(1)?.extract()?;
    let end_row = source_range.get_item(1)?.get_item(0)?.extract()?;
    let end_col = source_range.get_item(1)?.get_item(1)?.extract()?;

    // Read the source workbook or return an error if it doesn't exist  
    let source_workbook = reader::xlsx::read(source_file_path).map_err(|_| {
        let err_msg = format!("Failed to read Excel file: {}. Check if the file exists and is readable.", source_file_path);
        PyErr::new::<pyo3::exceptions::PyFileNotFoundError, _>(err_msg)
    })?;
    let source_sheet = source_workbook.get_sheet_by_name(source_sheet_name)
        .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyValueError, _>("Source sheet not found"))?;

    // Copy the range from the source sheet to the destination sheet
    let results =aggregate_range(source_sheet, start_row, start_col, end_row, end_col, action, mode.clone()).map_err(|e| {
        let err_msg = format!("Failed to aggregate range: {}", e);
        PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg)
    })?;

    debug!("Results: {:?}", results);

    let dest_row: u32 = dest_start_cell.get_item(0)?.extract()?;
    let dest_col: u32 = dest_start_cell.get_item(1)?.extract()?;
    // Read the destination workbook or return an error if it doesn't exist
    let mut dest_workbook = reader::xlsx::read(dest_file_path).map_err(|_| {
        let err_msg = format!("Failed to read Excel file: {}. Check if the file exists and is readable.", dest_file_path);
        PyErr::new::<pyo3::exceptions::PyFileNotFoundError, _>(err_msg)
    })?;
    let dest_sheet = dest_workbook.get_sheet_by_name_mut(dest_sheet_name)
        .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyValueError, _>("Destination sheet not found"))?;

    paste_vector_to_sheet(results, dest_sheet, dest_row, dest_col, mode.clone()).map_err(|e| {
        let err_msg = format!("Failed to paste vector to sheet: {}", e);
        PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg)
    })?;

    // Write the modified destination workbook to disk
    writer::xlsx::write(&dest_workbook, dest_file_path).map_err(|_| {
        let err_msg = format!("Failed to write Excel file: {}.", dest_file_path);
        PyErr::new::<pyo3::exceptions::PyIOError, _>(err_msg)
    })?;

    Ok(())
}
