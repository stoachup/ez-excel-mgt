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
    // convert PyTuple into Vec<u32>    
    let source_range: Vec<u32> = source_range.iter()
        .map(|x| x.extract::<u32>().map_err(|_| PyErr::new::<pyo3::exceptions::PyValueError, _>("Invalid range value.")))
        .collect::<Result<Vec<u32>, _>>()?; // Handle potential errors
    // convert PyTuple into Vec<u32>
    let dest_start_cell: Vec<u32> = dest_start_cell.iter()
        .map(|x| x.extract::<u32>().map_err(|_| PyErr::new::<pyo3::exceptions::PyValueError, _>("Invalid start cell value.")))
        .collect::<Result<Vec<u32>, _>>()?; // Handle potential errors
    // convert Option<bool> into bool
    let transpose = transpose.unwrap_or(false);

    // Replace the range copying logic with a call to copy_range_from_to
    match copy_range_from_to(
        source_file_path,
        source_sheet_name,
        &source_range,
        dest_file_path,
        dest_sheet_name,
        &dest_start_cell,
        transpose,
    ) {
        Ok(_) => Ok(()), // Return Ok if the operation is successful
        Err(e) => Err(PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!("Error copying range: {}", e))), // Convert error to PyErr
    }
}
