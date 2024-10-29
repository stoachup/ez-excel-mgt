#[allow(unused_imports)]
use log::{debug, info, warn};
use pyo3::prelude::*;
use pyo3::types::PyTuple;
use umya_spreadsheet::{self, reader, writer};


// Function to copy a range from one sheet to another
#[pyfunction]
pub fn copy_range_between_files(
    _py: Python,
    source_file_path: &str,
    source_sheet_name: &str,
    source_range: &PyTuple,
    dest_file_path: &str,
    dest_sheet_name: &str,
    dest_start_cell: &PyTuple,
    transpose: Option<bool>
) -> PyResult<()> {
    debug!("Copying range from {} to {}", source_file_path, dest_file_path);
    debug!("  ARGS:");
    debug!("        source_file_path: {:?}", source_file_path);
    debug!("        source_sheet_name: {:?}", source_sheet_name);
    debug!("        source_range: {:?}", source_range);
    debug!("        dest_file_path: {:?}", dest_file_path);
    debug!("        dest_sheet_name: {:?}", dest_sheet_name);
    debug!("        dest_start_cell: {:?}", dest_start_cell);
    debug!("        transpose: {:?}", transpose);
   // Extract values from the source_range tuple
    let start_row: u32 = source_range.get_item(0)?.extract()?;
    let start_col: u32 = source_range.get_item(1)?.extract()?;
    let end_row: u32 = source_range.get_item(2)?.extract()?;
    let end_col: u32 = source_range.get_item(3)?.extract()?;

    let transpose = transpose.unwrap_or(false);

    // Extract values from the dest_start_cell tuple
    let dest_row: u32 = dest_start_cell.get_item(0)?.extract()?;
    let dest_col: u32 = dest_start_cell.get_item(1)?.extract()?;

    // Read the source workbook or return an error if it doesn't exist  
    let source_workbook = reader::xlsx::read(source_file_path).map_err(|_| {
        let err_msg = format!("Failed to read Excel file: {}. Check if the file exists and is readable.", source_file_path);
        PyErr::new::<pyo3::exceptions::PyFileNotFoundError, _>(err_msg)
    })?;
    let source_sheet = source_workbook.get_sheet_by_name(source_sheet_name)
        .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyValueError, _>("Source sheet not found"))?;

    // Read the destination workbook or return an error if it doesn't exist
    let mut dest_workbook = reader::xlsx::read(dest_file_path).map_err(|_| {
        let err_msg = format!("Failed to read Excel file: {}. Check if the file exists and is readable.", dest_file_path);
        PyErr::new::<pyo3::exceptions::PyFileNotFoundError, _>(err_msg)
    })?;
    let dest_sheet = dest_workbook.get_sheet_by_name_mut(dest_sheet_name)
        .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyValueError, _>("Destination sheet not found"))?;

    // Copy the range from the source sheet to the destination sheet
    for col in start_col..=end_col {
        for row in start_row..=end_row {
            debug!("        ({}, {})", col, row);
            if let Some(source_cell) = source_sheet.get_cell((col, row)) {
                let value = source_cell.get_value().to_string();
                let d_col = dest_col + col - start_col;
                let d_row = dest_row + row - start_row;
                debug!("        value: {:?} from ({}, {}) to ({}, {})", value, col, row, d_col, d_row);
                
                // Use match to handle transpose logic
                match transpose {
                    false => {
                        dest_sheet
                            .get_cell_mut((d_col, d_row))
                            .set_value(value);
                    },
                    true => {
                        dest_sheet
                            .get_cell_mut((d_row, d_col))
                            .set_value(value);
                    },
                }
            }
        }
    }

    // Write the modified destination workbook to disk
    writer::xlsx::write(&dest_workbook, dest_file_path).map_err(|_| {
        let err_msg = format!("Failed to write Excel file: {}.", dest_file_path);
        PyErr::new::<pyo3::exceptions::PyIOError, _>(err_msg)
    })?;

    Ok(())
}
