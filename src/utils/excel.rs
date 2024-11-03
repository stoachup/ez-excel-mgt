use log::*;
use pyo3::prelude::*;
use pyo3::types::PyAny;


pub fn excel_col_to_index(col: &str) -> u32 {
    col.chars().rev().enumerate().fold(0, |acc, (i, c)| {
        acc + (c as u32 - 'A' as u32 + 1) * 26_u32.pow(i as u32)
    })
}

pub fn index_to_excel_col(col: u32) -> String {
    let mut col_str = String::new();
    let mut col_num = col;

    while col_num > 0 {
        let remainder = (col_num - 1) % 26;
        col_str.push((b'A' + remainder as u8) as char);
        col_num = (col_num - 1) / 26;
    }

    col_str.chars().rev().collect::<String>()
}

pub fn excel_to_index(cell: &str) -> (u32, u32) {
    let col_str = cell.chars().filter(|c| c.is_alphabetic()).collect::<String>();
    let row_str = cell.chars().filter(|c| c.is_numeric()).collect::<String>();
    
    // Convert column letters to a number
    let col = excel_col_to_index(&col_str);
    
    // Parse the row number
    let row = row_str.parse::<u32>().unwrap();
    
    (col, row)
}

#[allow(dead_code)]
pub fn excel_to_tuple(cell: &str) -> (u32, u32) {
    let (col, row) = excel_to_index(cell);
    (row, col)
}

// Function to convert a tuple (row, column) into an Excel cell (e.g., "B2")
pub fn index_to_excel(col: u32, row: u32) -> String {
    index_to_excel_col(col) + &row.to_string()
}

// Function to convert a tuple (row, column) into an Excel cell (e.g., "B2")
#[allow(dead_code)]
pub fn tuple_to_excel(row: u32, col: u32) -> String {
    index_to_excel(col, row)
}


/// Determine the row.
///
/// :param which_row: The row input from Python.
/// :param rows: (first_row, default_row, last_row).
/// :return: The determined row for column names.
#[allow(dead_code)]
pub fn get_header_row(py: Python, which_row: Option<PyObject>, rows: (u32, u32, u32)) -> PyResult<u32> {
    match which_row {
        Some(row) => { 
            let row_ref = row.as_ref(py); // Extract the reference here
            // check if the row is a string
            if PyAny::is_instance(row_ref, py.get_type::<pyo3::types::PyString>())? {
                let value: &str = row_ref.extract()?;
                debug!("[ROW] value: {} ({:?})", value, rows);
                match value {
                    "last" => Ok(rows.2), // last row
                    "first" => Ok(rows.0), // first row
                    _ => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("Invalid row identifier. Use 'first' or 'last'.")), // Raise error for unrecognized string
                }
            } 
            // check if the row is an integer
            else if PyAny::is_instance(row_ref, py.get_type::<pyo3::types::PyInt>())? {
                let n: u32 = row_ref.extract()?;
                debug!("[ROW] value: {} ({:?})", n, rows);
                match n {
                    // row numbering starts at 1
                    0 => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("Row numbering starts at 1.")),
                    _ => Ok(n) // Use specified row
                }
            } else {
                Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>("Invalid type for row input. Expected string or integer.")) // Raise error for invalid type
            }
        },
        None => {
            debug!("[ROW] value: {} ({:?})", rows.1, rows);
            Ok(rows.1) // default row
        }
    }
}
