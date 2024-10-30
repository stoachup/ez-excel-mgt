#[allow(unused_imports)]
use log::{debug, info, warn};
use pyo3::prelude::*;
use pyo3::types::PyTuple;

pub mod copy;
use copy::copy_range_from_to;


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
    // Extract values from the source_range tuple ((start_row, start_col), (end_row, end_col))
    let start_row = source_range.get_item(0)?.get_item(0)?.extract()?;
    let start_col = source_range.get_item(0)?.get_item(1)?.extract()?;
    let end_row = source_range.get_item(1)?.get_item(0)?.extract()?;
    let end_col = source_range.get_item(1)?.get_item(1)?.extract()?;

    // Extract values from the dest_start_cell tuple (dest_row, dest_col)
    let dest_row = dest_start_cell.get_item(0)?.extract()?;
    let dest_col = dest_start_cell.get_item(1)?.extract()?;
   
    // convert Option<bool> into bool
    let transpose = transpose.unwrap_or(false);

    // Replace the range copying logic with a call to copy_range_from_to
    match copy_range_from_to(
        source_file_path,
        source_sheet_name,
        ((start_row, start_col), (end_row, end_col)),
        dest_file_path,
        dest_sheet_name,
        (dest_row, dest_col),
        transpose,
    ) {
        Ok(_) => Ok(()), // Return Ok if the operation is successful
        Err(e) => Err(PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!("Error copying range: {}", e))), // Convert error to PyErr
    }
}
