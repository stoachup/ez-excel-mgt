#[allow(unused_imports)]
use log::{debug, info, warn};
use pyo3::prelude::*;
use pyo3::types::PyAny;
use polars::prelude::*;
use std::collections::HashMap;
use umya_spreadsheet::{self, reader, writer, structs::Worksheet};

mod py2rs; // Import the new module
use py2rs::{get_datatype, convert};

/// Convert an `AnyValue` from a Polars DataFrame into a string suitable for writing into an Excel cell.
///
/// This function converts Polars `AnyValue` types (e.g., string, int32, float64, etc.)
/// into a string representation to be inserted into an Excel worksheet.
///
/// :param value: The `AnyValue` from a Polars DataFrame.
/// :return: The string representation of the value for Excel.
fn convert_anyvalue_to_string(value: AnyValue) -> String {
    match value {
        AnyValue::Null => "".to_string(),        // For null values, return an empty string
        AnyValue::String(val) => val.to_string(),
        AnyValue::Int32(val) => val.to_string(),
        AnyValue::Int64(val) => val.to_string(),
        AnyValue::Float64(val) => val.to_string(),
        AnyValue::Boolean(val) => val.to_string(),
        _ => value.to_string(),
    }
}


// Function to delete all rows after a certain index
fn delete_rows_as_of(sheet: &mut umya_spreadsheet::Worksheet, row_index: u32) {
    let first_row: i32 = row_index as i32;
    let last_row: i32 = sheet.get_highest_row() as i32;

    let rows_to_delete = last_row - first_row + 1;

    if rows_to_delete > 0 {
        debug!("deleting {} rows as of row: {}", rows_to_delete, row_index);
        sheet.remove_row(&(first_row as u32), &(rows_to_delete as u32));
    }
}


/// Add data from a Polars DataFrame to an Excel worksheet by index.
///
/// This function inserts data from a Polars DataFrame into an Excel worksheet,
/// starting from the first empty row, based on the index position of the columns.
///
/// :param df: The Polars DataFrame.
/// :param sheet: The Excel worksheet to insert data into.
/// :param start_row: The row to start inserting data.
#[allow(dead_code)]
fn add_polars_df_to_sheet(
    df: &DataFrame,
    sheet: &mut Worksheet,
    start_row: u32,
    named: bool,
    strict: bool,
) -> Result<(), PyErr> {
    debug!("add_polars_df_to_sheet called with start_row: {}.", start_row);
    let height: usize = df.height();
    let width: usize = df.width();
    let sheet_width: usize = sheet.get_highest_column() as usize; // Get the number of columns in the sheet

    // Check column count based on strictness
    if strict {
        if width != sheet_width {
            let err_msg = format!("Numbers of columns do not match (DataFrame vs sheet: {} != {}.)", width, sheet_width);
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg)); // Handle the error as needed
        }
    } else {
        if width > sheet_width {
            let err_msg = format!("DataFrame has more columns than the sheet ({} > {}).", width, sheet_width);
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg)); // Handle the error as needed
        }
    }

    for col in 0..width {
        if let Some(series) = df.select_at_idx(col) {
            let mut header_row = 0;
            if named { 
                header_row = 1;
                sheet.get_cell_mut((col as u32 + 1, start_row)).set_value(series.name().to_string());
            }
            for row in 0..height {
                let value = series.get(row).unwrap();
                let cell_value = convert_anyvalue_to_string(value);
                debug!("[COL: {}, ROW: {}] value: {}", col as u32 + 1, header_row + row as u32 + start_row, cell_value);
                sheet.get_cell_mut((col as u32 + 1, header_row + row as u32 + start_row)).set_value(cell_value);
            }
        }
    }
    Ok(())
}


/// Add data from a Polars DataFrame to an Excel worksheet by column name.
///
/// This function inserts data from a Polars DataFrame into an Excel worksheet,
/// starting from the first empty row, matching columns by name.
///
/// :param df: The Polars DataFrame.
/// :param sheet: The Excel worksheet to insert data into.
/// :param start_row: The row to start inserting data.
/// :param template_column_map: A map of column names to Excel column indices.
fn add_polars_df_to_sheet_by_column_name(
    df: &DataFrame,
    sheet: &mut Worksheet,
    start_row: u32,
    column_map: &HashMap<String, u32>,
    strict: bool,
    skip_null: bool,
) -> Result<(), PyErr> {
    debug!("add_polars_df_to_sheet_by_column_name called with start_row: {}.", start_row);
    let df_columns: Vec<String> = df.get_column_names().iter().map(|s| s.to_string()).collect(); // Convert to Vec<String>
    
    if strict {
        // Check for missing columns in DataFrame
        for col_name in column_map.keys() {
            if !df_columns.contains(col_name) {
                debug!("Column '{}' is missing in the DataFrame", col_name);
                let err_msg = format!("Column '{}' in excel is missing in DataFrame.", col_name);
                return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg));
            }
        }
    }

    // Check for missing columns in the sheet
    for df_col in &df_columns {
        if !column_map.contains_key(df_col) {
            debug!("Column '{}' is missing in the template", df_col);
            let err_msg = format!("Column '{}' is missing in the template.", df_col);
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg));
        }
    }

    // Proceed with adding the DataFrame to the sheet if all columns match
    let height = df.height();

    for (col_name, excel_col_idx) in column_map {
        if let Some(series) = df.column(col_name).ok() {
            for row in 0..height {
                let value = series.get(row).unwrap();
                if skip_null && value == AnyValue::Null {
                    continue;
                } else {    
                    let cell_value = convert_anyvalue_to_string(value);
                    debug!("[COL: {} ({}), ROW: {}] value: {}", col_name, excel_col_idx, row as u32 + start_row, cell_value);
                    sheet.get_cell_mut((*excel_col_idx, row as u32 + start_row)).set_value(cell_value);
                }
            }
        }
    }

    delete_rows_as_of(sheet, start_row + height as u32);

    Ok(())
}

/// Extract column names from the last row of the worksheet.
///
/// This function scans the last row of a worksheet and returns a mapping
/// of column names to their corresponding Excel column indices.
///
/// :param sheet: The worksheet to scan.
/// :return: A HashMap where keys are column names and values are column indices.
fn get_column_map(sheet: &Worksheet, header_row: u32) -> HashMap<String, u32> {
    debug!("get_column_map called with header_row: {}", header_row);
    let mut column_map = HashMap::new();
    let last_col = sheet.get_highest_column();

    for col in 1..=last_col {
        let col_name = sheet.get_value((col, header_row)).to_string();
        debug!("[COLUMN MAP] coordinate: {:?}, value: {}", (col, header_row), col_name);
        column_map.insert(col_name, col);
    }

    column_map
}

/// Determine the row.
///
/// :param which_row: The row input from Python.
/// :param rows: (first_row, default_row, last_row).
/// :return: The determined row for column names.
fn get_header_row(py: Python, which_row: Option<PyObject>, rows: (u32, u32, u32)) -> PyResult<u32> {
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

/// Add a Polars or Pandas DataFrame to an Excel sheet.
///
/// This function inserts the content of a DataFrame (either Polars or Pandas) into a specified
/// Excel sheet. It starts inserting at the first empty row. The sheet must exist in the provided
/// Excel file.
///
/// :param df: The Polars or Pandas DataFrame to insert.
/// :param sheet_name: The name of the sheet in the Excel file to insert data into.
/// :param excel_file_path: The path to the Excel file.
/// :raises PyIOError: If the Excel file cannot be read.
/// :raises PyValueError: If the sheet does not exist.
#[pyfunction]
pub fn fill_sheet_with(
    py: Python,
    df: PyObject,
    excel_file_path: &str,
    sheet_name: &str,
    header_row: Option<PyObject>,
    columns: Option<PyObject>,
    strict: Option<bool>,
    skip_null: Option<bool>,
    overwrite: Option<bool>,
) -> PyResult<()> {
    debug!("fill_sheet_with called with");
    debug!("  ARGS:");
    debug!("        df: {:?}", df);
    debug!("        excel_file_path: {:?}", excel_file_path);
    debug!("        sheet_name: {:?}", sheet_name);
    debug!("  OPTS:");
    debug!("        header_row: {:?}", header_row);
    debug!("        columns: {:?}", columns);
    debug!("        strict: {:?}", strict);
    debug!("        skip_null: {:?}", skip_null);
    debug!("        overwrite: {:?}", overwrite);
    
    // Convert the input to a Polars DataFrame
    let skip_null = skip_null.unwrap_or(false);
    let strict = strict.unwrap_or(false);
    let overwrite = overwrite.unwrap_or(false);

    let data_type = get_datatype(py, df.as_ref(py))?;
    debug!("data_type: {}", data_type);

    let df: DataFrame = convert(py, data_type, df.as_ref(py), columns)?;
    debug!("df: {:?}", df);

    // Load the existing Excel workbook
    let mut workbook = reader::xlsx::read(excel_file_path).map_err(|_| {
        let err_msg = format!("Failed to read Excel file: {}. Check if the file exists and is readable.", excel_file_path);
        PyErr::new::<pyo3::exceptions::PyIOError, _>(err_msg)
    })?;

    // Get the sheet for mutability
    let mut sheet = workbook.get_sheet_by_name_mut(sheet_name).ok_or_else(|| {
        let err_msg = format!("Sheet '{}' does not exist.", sheet_name);
        PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg)
    })?;

    let last_row = sheet.get_highest_row();
    let header_row = get_header_row(py, header_row, (1, last_row, last_row))?;
    let column_map = get_column_map(&sheet, header_row);
    let start_row = if overwrite { header_row + 1 } else { last_row + 1 };

    add_polars_df_to_sheet_by_column_name(&df, &mut sheet, start_row, &column_map, strict, skip_null)?;
   
    // Write the modified workbook to disk
    writer::xlsx::write(&workbook, excel_file_path).map_err(|_| {
        let err_msg = format!("Failed to write Excel file: {}.", excel_file_path);
        PyErr::new::<pyo3::exceptions::PyIOError, _>(err_msg)
    })?;

    Ok(())
}