#[allow(unused_imports)]
use log::{debug, info, warn};
use pyo3::prelude::*;
use pyo3::types::PyAny;
use polars::prelude::{DataFrame, AnyValue, SerReader};
use polars::io::ipc::IpcReader;
use std::collections::HashMap;
use std::cmp::max;
use umya_spreadsheet::{self, reader, writer, structs::Worksheet};

/// Convert a Python Polars DataFrame to a Rust Polars DataFrame.
///
/// This function serializes a Python Polars DataFrame into Arrow format using `pyarrow`
/// and deserializes it back into a Rust Polars DataFrame using Polars' `IpcReader`.
///
/// :param py: The Python interpreter instance.
/// :param py_df: The Python Polars DataFrame to convert.
/// :return: A Rust Polars DataFrame.
fn py_polars_df_to_rust_polars_df(py: Python, py_df: &PyAny) -> PyResult<DataFrame> {
    let pyarrow: &PyModule = py.import("pyarrow").map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to import pyarrow module")
    })?;

    let arrow_table: &PyAny = py_df.call_method0("to_arrow").map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to convert DataFrame to Arrow format")
    })?;

    // Create an in-memory output stream
    let buffer: &PyAny = pyarrow.call_method0("BufferOutputStream").map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to create buffer stream")
    })?;

    // Use RecordBatchFileWriter to serialize the Arrow table into the buffer
    let writer: &PyAny = pyarrow
        .call_method1("RecordBatchFileWriter", (buffer, arrow_table.getattr("schema").unwrap()))
        .map_err(|_| {
            PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to create Arrow RecordBatchFileWriter")
        })?;
    writer.call_method1("write_table", (arrow_table,)).map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to write Arrow table")
    })?;
    writer.call_method0("close").map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to close Arrow writer")
    })?;

    // Extract the buffer's contents as bytes
    let buffer_bytes: Vec<u8> = buffer
        .call_method0("getvalue")
        .map_err(|_| PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to extract buffer"))?
        .extract()
        .map_err(|_| PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to extract buffer bytes"))?;

    // Deserialize into Rust Polars DataFrame using IpcReader
    let cursor = std::io::Cursor::new(buffer_bytes);
    IpcReader::new(cursor)
        .finish()
        .map_err(|_| PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to deserialize Arrow data"))
}

/// Convert a Pandas DataFrame to a Polars DataFrame in Rust.
///
/// This function takes a Pandas DataFrame and converts it into a Polars DataFrame
/// via an intermediary Polars DataFrame conversion in Python.
///
/// :param py: The Python interpreter instance.
/// :param df: The Python Pandas DataFrame.
/// :return: A Rust Polars DataFrame.
fn pandas_to_polars(py: Python, df: &PyAny) -> PyResult<DataFrame> {
    let polars: &PyModule = py.import("polars").map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to import polars module")
    })?;
    let df_polars: &PyAny = polars.call_method1("DataFrame", (df,)).map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to convert Pandas DataFrame to Polars")
    })?;
    py_polars_df_to_rust_polars_df(py, df_polars)
}

/// Convert an `AnyValue` into a string suitable for writing into an Excel cell.
///
/// This function converts Polars `AnyValue` types (e.g., string, int32, float64, etc.)
/// into a string representation to be inserted into an Excel worksheet.
///
/// :param value: The `AnyValue` from a Polars DataFrame.
/// :return: The string representation of the value for Excel.
fn convert_anyvalue_to_string(value: AnyValue) -> String {
    match value {
        AnyValue::String(val) => val.to_string(),   // For string values
        AnyValue::Int32(val) => val.to_string(),  // For 32-bit integers
        AnyValue::Int64(val) => val.to_string(),  // For 64-bit integers
        AnyValue::Float64(val) => val.to_string(), // For floating-point values
        AnyValue::Boolean(val) => val.to_string(), // For boolean values
        _ => value.to_string(), // Fallback for other types
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
fn add_polars_df_to_sheet(
    df: &DataFrame,
    sheet: &mut Worksheet,
    start_row: u32,
    named: bool,
    strict: bool,
) -> Result<(), PyErr> {
    debug!("add_polars_df_to_sheet called with start_row: {}", start_row);
    let height: usize = df.height();
    let width: usize = df.width();
    let sheet_width: usize = sheet.get_highest_column() as usize; // Get the number of columns in the sheet

    // Check column count based on strictness
    if strict {
        if width != sheet_width {
            let err_msg = format!("Numbers of columns do not match (DataFrame vs sheet: {} != {})", width, sheet_width);
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg)); // Handle the error as needed
        }
    } else {
        if width > sheet_width {
            let err_msg = format!("DataFrame has more columns than the sheet ({} > {})", width, sheet_width);
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
) -> Result<(), PyErr> {
    debug!("add_polars_df_to_sheet_by_column_name called with start_row: {}", start_row);
    let df_columns: Vec<String> = df.get_column_names().iter().map(|s| s.to_string()).collect(); // Convert to Vec<String>
    
    if strict {
        // Check for missing columns in DataFrame
        for col_name in column_map.keys() {
            if !df_columns.contains(col_name) {
                let err_msg = format!("Column '{}' in excel is missing in DataFrame", col_name);
                return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg));
            }
        }
    }

    // Check for missing columns in the sheet
    for df_col in &df_columns {
        if !column_map.contains_key(df_col) {
            let err_msg = format!("Column '{}' in DataFrame is missing in the template", df_col);
            return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg));
        }
    }

    // Proceed with adding the DataFrame to the sheet if all columns match
    let height = df.height();

    for (col_name, excel_col_idx) in column_map {
        if let Some(series) = df.column(col_name).ok() {
            for row in 0..height {
                let value = series.get(row).unwrap();
                let cell_value = convert_anyvalue_to_string(value);
                debug!("[COL: {} ({}), ROW: {}] value: {}", col_name, excel_col_idx, row as u32 + start_row, cell_value);
                sheet.get_cell_mut((*excel_col_idx, row as u32 + start_row)).set_value(cell_value);
            }
        }
    }

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

/// Determine the row for column names based on the header row input.
///
/// :param which_row: The row input from Python.
/// :param rows: (first_row, default_row, last_row).
/// :return: The determined row for column names.
fn get_row(py: Python, which_row: Option<PyObject>, rows: (u32, u32, u32)) -> PyResult<u32> {
    debug!("get_row called with which_row: {:?}, rows: {:?}", which_row, rows);
    match which_row {
        Some(row) => { 
            let row_ref = row.as_ref(py); // Extract the reference here
            if PyAny::is_instance(row_ref, py.get_type::<pyo3::types::PyString>())? {
                let value: &str = row_ref.extract()?;
                debug!("[ROW] value: {}", value);
                match value {
                    "last" => Ok(rows.2),
                    "first" => Ok(rows.0),
                    _ => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("Invalid row identifier. Use 'first' or 'last'.")), // Raise error for unrecognized string
                }
            } else if PyAny::is_instance(row_ref, py.get_type::<pyo3::types::PyInt>())? {
                let n: u32 = row_ref.extract()?;
                debug!("[ROW] value: {}", n);
                match n {
                    0 => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("Row numbering starts at 1.")),
                    _ => Ok(n) // Use specified row
                }
            } else {
                Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>("Invalid type for row input. Expected string or integer.")) // Raise error for invalid type
            }
        },
        None => Ok(rows.1)
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
    excel_file_path: &str,
    sheet_name: &str,
    df: PyObject,
    strict: Option<bool>,
    named: Option<bool>,
    overwrite: Option<bool>,
    header_row: Option<PyObject>,
    start_row: Option<PyObject>,
) -> PyResult<()> {
    debug!("fill_sheet_with called with");
    debug!("  ARGS:");
    debug!("        excel_file_path: {:?}", excel_file_path);
    debug!("        sheet_name: {:?}", sheet_name);
    debug!("        df: {:?}", df);
    debug!("  OPTS: strict: {:?}, named: {:?}, overwrite: {:?}, header_row: {:?}, start_row: {:?}", strict, named, overwrite, header_row, start_row);
    // Determine if the DataFrame is Pandas or Polars
    let df: DataFrame = if PyAny::is_instance(df.as_ref(py), py.get_type::<PyAny>())? {
        pandas_to_polars(py, df.as_ref(py))?
    } else {
        let polars = py.import("polars").map_err(|_| {
            PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to import polars module")
        })?;
        let py_df = polars.call_method1("DataFrame", (df,)).map_err(|_| {
            PyErr::new::<pyo3::exceptions::PyTypeError, _>("Failed to convert input to Polars DataFrame")
        })?;
        py_polars_df_to_rust_polars_df(py, py_df)?
    };

    // Load the existing Excel workbook
    let mut workbook = reader::xlsx::read(excel_file_path).map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyIOError, _>("Failed to read Excel file")
    })?;

    // Get the sheet for mutability
    let mut sheet = workbook.get_sheet_by_name_mut(sheet_name).ok_or_else(|| {
        PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Sheet '{}' does not exist", sheet_name))
    })?;

    let strict = strict.unwrap_or(false);
    debug!("strict: {:?}", strict);
    
    let last_row = sheet.get_highest_row();
    debug!("last_row: {:?}", last_row);
    let named = named.unwrap_or(false);
    let overwrite = overwrite.unwrap_or(false);
    debug!("named: {:?}, overwrite: {:?}", named, overwrite);
    let header_row = get_row(py, header_row, (1, last_row, last_row))?;
    let default_row = if overwrite { 1 } else { last_row + 1 };
    let start_row = get_row(py, start_row, (1, default_row, last_row + 1))?;
    debug!("start_row: {:?}, default_row: {:?}, header_row: {:?}", start_row, default_row, header_row);
    match (named, overwrite) {
        (false, true) => {
            debug!("1: named is false and overwrite is true");
            add_polars_df_to_sheet(&df, &mut sheet, start_row, true, strict)?;
        },
        (false, false) => {
            debug!("2: named is false and overwrite is false");
            let start_row = max(start_row, last_row + 1);
            add_polars_df_to_sheet(&df, &mut sheet, start_row, false, strict)?;
        },
        (true, false) => {
            debug!("3: named is true and overwrite is false");
            if start_row > header_row {
                let column_map = get_column_map(&sheet, header_row);
                add_polars_df_to_sheet_by_column_name(&df, &mut sheet, last_row + 1, &column_map, strict)?;
            } else {
                return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("start_row must be greater than header_row"));
            }
        },
        (true, true) => {
            debug!("4: named and overwrite are true");
            if start_row > last_row {
                let column_map = get_column_map(&sheet, header_row);
                add_polars_df_to_sheet_by_column_name(&df, &mut sheet, start_row, &column_map, strict)?;
            } else {
                return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>("start_row must be greater than the last row of the sheet"));
            }
        }
    }
   
    // Write the modified workbook to disk
    writer::xlsx::write(&workbook, excel_file_path).map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyIOError, _>("Failed to write Excel file")
    })?;
    Ok(())
}

