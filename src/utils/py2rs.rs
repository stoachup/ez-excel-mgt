#[allow(unused_imports)]
use log::{debug, info, warn, error};
use pyo3::prelude::*;
use pyo3::types::PyAny;
use pyo3::exceptions::{PyImportError, PyRuntimeError, PyTypeError, PyValueError};
use pyo3::{PyErr, PyTypeInfo};
use polars::prelude::*;
use std::collections::HashMap;
use std::fmt;

fn py_err<E>(err_msg: String) -> PyErr
where
    E: PyTypeInfo,
{
    error!("{}", err_msg);
    PyErr::new::<E, _>(err_msg)
}



pub fn convert_anyvalue_to_string(value: AnyValue) -> String {
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
        py_err::<PyImportError>(format!("Failed to import pyarrow module."))
    })?;

    let arrow_table: &PyAny = py_df.call_method0("to_arrow").map_err(|_| {
        py_err::<PyRuntimeError>(format!("Failed to convert DataFrame to Arrow format."))
    })?;

    // Create an in-memory output stream
    let buffer: &PyAny = pyarrow.call_method0("BufferOutputStream").map_err(|_| {
        py_err::<PyRuntimeError>(format!("Failed to create buffer stream."))
    })?;

    // Use RecordBatchFileWriter to serialize the Arrow table into the buffer
    let writer: &PyAny = pyarrow
        .call_method1("RecordBatchFileWriter", (buffer, arrow_table.getattr("schema").unwrap()))
        .map_err(|_| {
            py_err::<PyRuntimeError>(format!("Failed to create Arrow RecordBatchFileWriter."))
        })?;
    writer.call_method1("write_table", (arrow_table,)).map_err(|_| {
            py_err::<PyRuntimeError>(format!("Failed to write Arrow table."))
        })?;
    writer.call_method0("close").map_err(|_| {
        py_err::<PyRuntimeError>(format!("Failed to close Arrow writer."))
    })?;

    // Extract the buffer's contents as bytes
    let buffer_bytes: Vec<u8> = buffer
        .call_method0("getvalue")
        .map_err(|_| py_err::<PyRuntimeError>(format!("Failed to extract buffer.")))?
        .extract()
        .map_err(|_| py_err::<PyRuntimeError>(format!("Failed to extract buffer bytes.")))?;

    // Deserialize into Rust Polars DataFrame using IpcReader
    let cursor = std::io::Cursor::new(buffer_bytes);
    IpcReader::new(cursor)
        .finish()
        .map_err(|_| py_err::<PyRuntimeError>(format!("Failed to deserialize Arrow data.")))
}

/// Convert a Pandas DataFrame to a Polars DataFrame in Rust.
///
/// This function takes a Pandas DataFrame and converts it into a Polars DataFrame
/// via an intermediary Polars DataFrame conversion in Python.
///
/// :param py: The Python interpreter instance.
/// :param df: The Python Pandas DataFrame.
/// :return: A Rust Polars DataFrame.
fn py_pandas_df_to_rust_polars_df(py: Python, df: &PyAny) -> PyResult<DataFrame> {
    let polars: &PyModule = py.import("polars").map_err(|_| {
        py_err::<PyImportError>(format!("Failed to import polars module."))
    })?;
    let df_polars: &PyAny = polars.call_method1("DataFrame", (df,)).map_err(|_| {
        py_err::<PyTypeError>(format!("Failed to convert Pandas DataFrame to Polars."))
    })?;
    py_polars_df_to_rust_polars_df(py, df_polars)
}

/// Extracts a Polars Series from a vector of optional Python objects.
///
/// This function takes a vector of optional Python objects and infers the type of the first
/// non-None value to determine the appropriate Rust type for the Series. It handles
/// String, integer, float, and boolean types, returning a Series containing the extracted values.
///
/// :param py: The Python interpreter instance.
/// :param column: A vector of optional PyObject values representing the column data.
/// :param name: The name of the Series to be created.
/// :param max_column_len: The maximum length of the column, used to ensure consistent Series length.
/// :return: A PyResult containing the constructed Series or an error if the type is unsupported.
fn extract_series_from_vec_of_optional_py_objects(py: Python, column: &Vec<Option<PyObject>>, name: &str, max_column_len: usize) -> PyResult<Series> {
    // Find the first non-null value to infer the column type
    let first_non_null = column.iter().flatten().next(); // Find the first non-None value

    if let Some(value) = first_non_null {
        let first_value = value.as_ref(py);
        if first_value.is_instance(py.get_type::<pyo3::types::PyString>())? {
            // Handle String type
            let extracted_values: Vec<Option<String>> = (0..max_column_len)
                .map(|idx| {
                    column.get(idx)
                        .and_then(|val| val.as_ref().and_then(|v| v.extract::<Option<String>>(py).ok()).flatten())
                })
                .collect();
            return Ok(Series::new(name.into(), extracted_values));
        } else if first_value.is_instance(py.get_type::<pyo3::types::PyInt>())? {
            // Handle integer type
            let extracted_values: Vec<Option<i32>> = (0..max_column_len)
                .map(|idx| {
                    column.get(idx)
                        .and_then(|val| val.as_ref().and_then(|v| v.extract::<Option<i32>>(py).ok()).flatten())
                })
                .collect();
            return Ok(Series::new(name.into(), extracted_values));
        } else if first_value.is_instance(py.get_type::<pyo3::types::PyFloat>())? {
            // Handle float type
            let extracted_values: Vec<Option<f64>> = (0..max_column_len)
                .map(|idx| {
                    column.get(idx)
                        .and_then(|val| val.as_ref().and_then(|v| v.extract::<Option<f64>>(py).ok()).flatten())
                })
                .collect();
            return Ok(Series::new(name.into(), extracted_values));
        } else if first_value.is_instance(py.get_type::<pyo3::types::PyBool>())? {
            // Handle boolean type
            let extracted_values: Vec<Option<bool>> = (0..max_column_len)
                .map(|idx| {
                    column.get(idx)
                        .and_then(|val| val.as_ref().and_then(|v| v.extract::<Option<bool>>(py).ok()).flatten())
                })
                .collect();
            return Ok(Series::new(name.into(), extracted_values));
        } else {
            Err(py_err::<PyTypeError>(format!("Unsupported value type in column")))
        }
    } else {
        Err(py_err::<PyTypeError>(format!("Column '{}' contains only None values or is empty", name)))
    }
}


///
/// This function takes a Python dictionary where each key corresponds to a list of values
/// and converts it into a Polars DataFrame, ensuring that each column contains data of the same type.
///
/// :param py: The Python interpreter instance.
/// :param dict: The Python dictionary to convert.
/// :return: A Rust Polars DataFrame.
fn py_dict_of_lists_to_rust_polars_df(py: Python, dict_of_lists: &PyAny) -> PyResult<DataFrame> {
    // Check if df is a HashMap<String, Vec<Option<PyObject>>>
    let dict_of_lists: HashMap<String, Vec<Option<PyObject>>> = dict_of_lists.extract().map_err(|_| {
        py_err::<PyTypeError>(format!("Structure of dictionary of lists is not correct."))
    })?;

    // Create a vector to store the columns
    let mut columns: Vec<Series> = Vec::with_capacity(dict_of_lists.len());

    // Determine the maximum column length (since empty lists may exist)
    let max_column_len = dict_of_lists.iter().map(|(_, values)| values.len()).max().unwrap_or(0);

    // Iterate over each key-value pair in the dictionary
    for (name, values) in dict_of_lists {
        //
        if values.len() != max_column_len {
            return Err(py_err::<PyValueError>(format!("At least one list in the dictionary of lists has a different length than the others.")));
        }

        // Extract the series from the list of optional PyObject
        let series = extract_series_from_vec_of_optional_py_objects(py, &values, name.as_str(), max_column_len)?;
        columns.push(series);
    }

    DataFrame::new(columns).map_err(|e| {
        py_err::<PyTypeError>(format!("Failed to create DataFrame from dictionary of lists: {}.", e))
    })
}


/// Convert a Python list of lists to a Rust Polars DataFrame.
///
/// This function takes a Python list where each inner list represents a column of values
/// and converts it into a Polars DataFrame. It also handles `None` values, treating them as nulls.
///
/// :param py: The Python interpreter instance.
/// :param list_of_lists: The Python list of lists to convert.
/// :param columns: The Python list of column names.
/// :return: A Rust Polars DataFrame.
fn py_list_of_lists_to_rust_polars_df(py: Python, list_of_lists: &PyAny, columns: &PyAny) -> PyResult<DataFrame> {
    // Extract column names from the Python list
    let columns: Vec<String> = columns.extract().map_err(|_| {
        py_err::<PyTypeError>(format!("List of columns is not correct."))
    })?;
    
    // Extract the list of lists from Python
    let list_of_lists: Vec<Vec<Option<PyObject>>> = list_of_lists.extract().map_err(|_| {
        py_err::<PyTypeError>(format!("Structure of list of lists is not correct."))
    })?;
    
    // Check if the number of columns and number of lists match
    if columns.len() != list_of_lists.len() {
        return Err(py_err::<PyValueError>(format!("List of columns and list of lists have different lengths.")))
        }

    // Create a vector to store the columns
    let mut df_columns: Vec<Series> = Vec::with_capacity(columns.len());

    // Determine the maximum column length (since empty lists may exist)
    let max_column_len = list_of_lists.iter().map(|col| col.len()).max().unwrap_or(0);

    // Iterate over the list of lists and columns together
    for (values, col_name) in list_of_lists.iter().zip(columns.iter()) {
        let name = col_name.clone(); // Use the column name from the list

        // Handle the case where the list (column) is empty
        if values.is_empty() {
            df_columns.push(Series::new_null(name.into(), max_column_len));
            continue;
        }

        // Check if the length of the current list matches the max length
        else if values.len() != max_column_len {
            return Err(py_err::<PyValueError>(format!("At least one list in the list of lists has a different length than the others.")));
        }

        // Convert the list of optional PyObject values into a Polars Series
        let series = extract_series_from_vec_of_optional_py_objects(py, &values, name.as_str(), max_column_len)?;
        df_columns.push(series);
    }

    // Create a DataFrame from the series of columns
    DataFrame::new(df_columns).map_err(|e| {
        py_err::<PyRuntimeError>(format!("Failed to create DataFrame from list of lists: {}.", e))
    })
}


/// Get the DataFrame type from the specified module.
///
/// :param py: The Python interpreter instance.
/// :param module_name: The name of the module to import (e.g., "pandas" or "polars").
/// :return: The DataFrame type from the specified module.
fn get_dataframe_type<'py>(py: Python<'py>, module_name: &str) -> PyResult<&'py PyAny> {
    // Import the module
    let module = PyModule::import(py, module_name).map_err(|e| {
        py_err::<PyImportError>(format!("Failed to import {} module: {}.", module_name, e))
    })?;
    
    // Get the DataFrame type from the module
    module.getattr("DataFrame").map_err(|e| {
        py_err::<PyRuntimeError>(format!("Failed to get DataFrame from module {}: {}.", module_name, e))
    })
}


// Add an enum for the different types of dataframes
pub enum OriginalDataType {
    Pandas,
    Polars,
    DictionaryOfLists,
    ListOfLists,
}

impl fmt::Display for OriginalDataType {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            OriginalDataType::Pandas => write!(f, "Pandas DataFrame"),
            OriginalDataType::Polars => write!(f, "Polars DataFrame"),
            OriginalDataType::DictionaryOfLists => write!(f, "Dictionary of Lists"),
            OriginalDataType::ListOfLists => write!(f, "List of Lists"),
        }
    }
}

/// Convert a Python object to a Polars DataFrame.
///
/// This function takes a Python object and converts it into a Polars DataFrame.
/// It checks if the input is a Pandas DataFrame or a dictionary and performs
/// the appropriate conversion.
///
/// :param py: The Python interpreter instance.
/// :param df: The Python object to convert.
/// :return: A Rust Polars DataFrame and the type of the original dataframe.
pub fn get_datatype(py: Python, df: &PyAny) -> PyResult<OriginalDataType> {
    let pandas_type = get_dataframe_type(py, "pandas")?;
    let polars_type = get_dataframe_type(py, "polars")?;

    if df.is_instance(pandas_type)? {
        debug!("Pandas DataFrame found");
        Ok(OriginalDataType::Pandas)
    } else if df.is_instance(polars_type)? {
        debug!("Polars DataFrame found");
        Ok(OriginalDataType::Polars)
    } else if df.is_instance(py.get_type::<pyo3::types::PyDict>())? {
        debug!("Dictionary of lists found");
        Ok(OriginalDataType::DictionaryOfLists)
    } else if df.is_instance(py.get_type::<pyo3::types::PyList>())? {
        debug!("List of lists found");
        Ok(OriginalDataType::ListOfLists)
    } else {
        let err_msg = format!("Input must be a Pandas or Polars DataFrame, dictionary of lists or list of lists with column names.");
        error!("{}", err_msg);
        Err(py_err::<PyTypeError>(err_msg))
    }
}


/// Convert a Python object to a Polars DataFrame.
///
/// This function takes a Python object and converts it into a Polars DataFrame.
/// It checks if the input is a Pandas DataFrame or a dictionary and performs
/// the appropriate conversion.
///
/// :param py: The Python interpreter instance.
/// :param df: The Python object to convert.
/// :return: A Rust Polars DataFrame and the type of the original dataframe.
pub fn convert(py: Python, data_type: OriginalDataType, df: &PyAny, columns: Option<PyObject>) -> PyResult<DataFrame> {
    match (data_type, columns) {
        (OriginalDataType::ListOfLists, Some(columns)) => {
            // Convert Dict[str, List[Any]] to Polars DataFrame
            Ok(py_list_of_lists_to_rust_polars_df(py, df, columns.as_ref(py))?)
        },
        (OriginalDataType::Pandas, None) => {
            // Convert Python Pandas DataFrame to Rust Polars DataFrame
            Ok(py_pandas_df_to_rust_polars_df(py, df)?)
        },
        (OriginalDataType::Polars, None) => {
            // Convert Python Polars DataFrame to Rust Polars DataFrame
            Ok(py_polars_df_to_rust_polars_df(py, df)?)
        },
        (OriginalDataType::DictionaryOfLists, None) => {
            // Convert Dict[str, List[Any]] to Polars DataFrame
            Ok(py_dict_of_lists_to_rust_polars_df(py, df)?)
        },
        (_, Some(_)) => {
            Err(py_err::<PyValueError>(format!("Column names should not be provided for Pandas, Polars and Dict of Lists.")))
        },
        (OriginalDataType::ListOfLists, None) => {
            Err(py_err::<PyValueError>(format!("Column names must be provided for List of Lists.")))
        }
    }
}