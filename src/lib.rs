use env_logger::Builder;
#[allow(unused_imports)]
use log::{debug, info, warn, LevelFilter};
use pyo3::prelude::*;
use pyo3::types::PyAny;
use polars::prelude::{DataFrame, AnyValue, SerReader};
use polars::io::ipc::IpcReader;
use std::collections::HashMap;
use std::env;
use std::io::Write;
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
#[allow(dead_code)]
fn add_polars_df_to_sheet(
    df: &DataFrame,
    sheet: &mut Worksheet,
    start_row: u32,
) {
    let height: usize = df.height();
    let width: usize = df.width();

    for col in 0..width {
        if let Some(series) = df.select_at_idx(col) {
            for row in 0..height {
                let value = series.get(row).unwrap();
                let cell_value = convert_anyvalue_to_string(value);
                debug!("[INDEX] coordinate: {:?}, value: {}", (col as u32 + 1, row as u32 + start_row), cell_value);
                sheet.get_cell_mut((col as u32 + 1, row as u32 + start_row)).set_value(cell_value);
            }
        }
    }
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
    template_column_map: &HashMap<String, u32>,
) {
    let height = df.height();

    for (col_name, excel_col_idx) in template_column_map {
        if let Some(series) = df.column(col_name).ok() {
            for row in 0..height {
                let value = series.get(row).unwrap();
                let cell_value = convert_anyvalue_to_string(value);
                debug!("[COLUMN] coordinate: {:?}, value: {}", (excel_col_idx, row as u32 + start_row), cell_value);
                sheet.get_cell_mut((*excel_col_idx, row as u32 + start_row)).set_value(cell_value);
            }
        }
    }
}

/// Extract column names from the last row of the template worksheet.
///
/// This function scans the last row of a template worksheet and returns a mapping
/// of column names to their corresponding Excel column indices.
///
/// :param template_sheet: The worksheet to scan.
/// :return: A HashMap where keys are column names and values are column indices.
fn get_template_column_map(template_sheet: &Worksheet) -> HashMap<String, u32> {
    let mut column_map = HashMap::new();
    let last_row = template_sheet.get_highest_row();
    let last_col = template_sheet.get_highest_column();

    for col in 1..=last_col {
        let col_name = template_sheet.get_value((col, last_row)).to_string();
        debug!("[COLUMN MAP] coordinate: {:?}, value: {}", (col, last_row), col_name);
        column_map.insert(col_name, col);
    }

    column_map
}

/// Initialize logging for the module.
///
/// This function initializes logging by checking the environment variables `LOGLEVEL` and `RUST_LOG`.
/// If neither is set, it defaults to the "error" log level.
fn init_logging() {
    // First check `LOGLEVEL`, then fallback to `RUST_LOG`, or default to "error"
    let log_env = env::var("LOGLEVEL")
        .or_else(|_| env::var("RUST_LOG"))
        .unwrap_or_else(|_| "error".to_string());

    let log_level = log_env.parse::<LevelFilter>().unwrap_or_else(|_| {
        warn!("Invalid log level: {}. Defaulting to 'error'.", log_env);
        LevelFilter::Error
    });

    // Directly initialize env_logger with the log level we determined
    Builder::new()
        .filter_level(log_level)
        .format(|buf, record| writeln!(buf, "[RUST:{}] - {}", record.level(), record.args()))
        .init();
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
fn add_df_to_sheet(
    py: Python,
    df: PyObject, 
    sheet_name: &str,
    excel_file_path: &str,
) -> PyResult<()> {
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
    
    let last_row = sheet.get_highest_row();

    // Get the column map and insert data
    let template_column_map = get_template_column_map(&sheet);
    add_polars_df_to_sheet_by_column_name(&df, &mut sheet, last_row + 1, &template_column_map);
    
    // Write the modified workbook to disk
    writer::xlsx::write(&workbook, excel_file_path).map_err(|_| {
        PyErr::new::<pyo3::exceptions::PyIOError, _>("Failed to write Excel file")
    })?;
    Ok(())
}

/// Python module initialization function.
#[pymodule]
fn ez_excel_mgt(_py: Python, m: &PyModule) -> PyResult<()> {
    // Initialize logger only once
    init_logging();

    // Add functions to the Python module
    m.add_function(wrap_pyfunction!(add_df_to_sheet, m)?)?;
    Ok(())
}