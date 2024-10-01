#[allow(unused_imports)]
use log::{debug, info, warn, error};
use pyo3::prelude::*;
use pyo3::types::{PyAny, PyString, PyInt, PyFloat, PyBool};
use pyo3::exceptions::{PyTypeError, PyValueError};
use polars::prelude::*;
use std::collections::HashMap;
use std::fmt;

/// Type alias for Python result types that will return a DataFrame on success.
type PyResultDF = PyResult<DataFrame>;

/// Type alias for Python dictionary of lists.
type DictOfLists = HashMap<String, Vec<Option<PyObject>>>;

/// Type alias for Python list of lists
type ListOfLists = Vec<Vec<Option<PyObject>>>;

/// Generic utility function to create a Python error of type `M` and log the error message.
///
/// This function logs the error message and returns a Python exception of the specified type.
///
/// # Arguments
/// * `msg` - The error message to log and use in the exception.
///
/// # Returns
/// A `PyResult` containing an error of the specified Python exception type `M`.
fn py_err<E, T>(msg: &str) -> PyResult<T>
where
    E: pyo3::type_object::PyTypeInfo,
{
    let msg = msg.to_string();  // Convert borrowed msg to owned String
    error!("{}", msg);  // Log the error message
    Err(PyErr::new::<E, _>(msg))  // Return the Python error
}

/// Convert a Python Polars DataFrame to a Rust Polars DataFrame.
///
/// This function serializes a Python Polars DataFrame into Arrow format using `pyarrow`
/// and deserializes it back into a Rust Polars DataFrame using Polars' `IpcReader`.
///
/// # Arguments
/// * `py` - The Python interpreter instance.
/// * `py_df` - The Python Polars DataFrame to convert.
///
/// # Returns
/// A `PyResultDF` containing a Rust Polars DataFrame if successful, or a Python error if failed.
fn py_polars_df_to_rust_polars_df(py: Python, py_df: &PyAny) -> PyResultDF {
    // Import the pyarrow module
    let pyarrow = py.import("pyarrow")
        .map_err(|_| py_err::<PyTypeError, _>("Failed to import pyarrow module."))?;

    // Convert the Python Polars DataFrame to an Arrow table
    let arrow_table = py_df.call_method0("to_arrow")
        .map_err(|_| py_err::<PyTypeError, _>("Failed to convert DataFrame to Arrow format."))?;

    // Create an in-memory buffer stream for serialization
    let buffer = pyarrow.call_method0("BufferOutputStream")
        .map_err(|_| py_err::<PyTypeError, _>("Failed to create buffer stream."))?;

    // Serialize the Arrow table using a RecordBatchFileWriter
    let writer = pyarrow.call_method1("RecordBatchFileWriter", 
        (buffer, arrow_table.getattr("schema")?))
        .map_err(|_| py_err::<PyTypeError, _>("Failed to create Arrow RecordBatchFileWriter."))?;
    
    // Write the Arrow table into the buffer
    writer.call_method1("write_table", (arrow_table,))
        .map_err(|_| py_err::<PyTypeError, _>("Failed to write Arrow table."))?;
    
    // Close the writer to finalize the buffer
    writer.call_method0("close")
        .map_err(|_| py_err::<PyTypeError, _>("Failed to close Arrow writer."))?;

    // Extract the buffer's content as bytes
    let buffer_bytes: Vec<u8> = buffer.call_method0("getvalue")?
        .extract()
        .map_err(|_| py_err::<PyTypeError, _>("Failed to extract buffer bytes."))?;

    // Deserialize the Arrow data into a Rust Polars DataFrame
    IpcReader::new(std::io::Cursor::new(buffer_bytes))
        .finish()
        .map_err(|_| py_err::<PyTypeError, _>("Failed to deserialize Arrow data."))
}

/// Convert a Python Pandas DataFrame to a Rust Polars DataFrame.
///
/// This function first converts the Pandas DataFrame to a Polars DataFrame in Python,
/// and then converts it to a Rust Polars DataFrame.
///
/// # Arguments
/// * `py` - The Python interpreter instance.
/// * `df` - The Python Pandas DataFrame.
///
/// # Returns
/// A `PyResultDF` containing a Rust Polars DataFrame if successful, or a Python error if failed.
fn py_pandas_df_to_rust_polars_df(py: Python, df: &PyAny) -> PyResultDF {
    // Import the polars module
    let polars = py.import("polars")
        .map_err(|_| py_err::<PyTypeError>("Failed to import polars module."))?;
    
    // Convert the Pandas DataFrame to a Polars DataFrame
    let df_polars = polars.call_method1("DataFrame", (df,))
        .map_err(|_| py_err::<PyTypeError>("Failed to convert Pandas DataFrame to Polars."))?;
    
    // Convert the Polars DataFrame to a Rust Polars DataFrame
    py_polars_df_to_rust_polars_df(py, df_polars)
}

/// Extract a Polars Series from a vector of optional Python objects.
///
/// This function takes a vector of optional Python objects and infers the type of the first
/// non-None value to determine the appropriate Rust type for the Series. It handles
/// String, integer, float, and boolean types.
///
/// # Arguments
/// * `py` - The Python interpreter instance.
/// * `column` - A vector of optional PyObject values representing the column data.
/// * `name` - The name of the Series to be created.
/// * `max_column_len` - The maximum length of the column, used to ensure consistent Series length.
///
/// # Returns
/// A `PyResult` containing the constructed Series or an error if the type is unsupported.
fn extract_series_from_vec_of_optional_py_objects(
    py: Python,
    column: &[Option<PyObject>],
    name: &str,
    max_column_len: usize,
) -> PyResult<Series> {
    // Find the first non-null value to infer the column type
    let first_non_null = column.iter().flatten().next();
    if let Some(value) = first_non_null {
        let name = name.to_string();
        let first_value = value.as_ref(py);
        
        // Check the type of the first non-null value and create the appropriate Series
        let series: Series = if first_value.is_instance(py.get_type::<PyString>())? {
            // Handle String type
            let values: Vec<Option<String>> = column.iter()
                .map(|val| val.as_ref().and_then(|v| v.extract(py).ok()))
                .collect();
            Series::new(name.into(), values)
        } else if first_value.is_instance(py.get_type::<PyInt>())? {
            // Handle integer type
            let values: Vec<Option<i32>> = column.iter()
                .map(|val| val.as_ref().and_then(|v| v.extract(py).ok()))
                .collect();
            Series::new(name.into(), values)
        } else if first_value.is_instance(py.get_type::<PyFloat>())? {
            // Handle float type
            let values: Vec<Option<f64>> = column.iter()
                .map(|val| val.as_ref().and_then(|v| v.extract(py).ok()))
                .collect();
            Series::new(name.into(), values)
        } else if first_value.is_instance(py.get_type::<PyBool>())? {
            // Handle boolean type
            let values: Vec<Option<bool>> = column.iter()
                .map(|val| val.as_ref().and_then(|v| v.extract(py).ok()))
                .collect();
            Series::new(name.into(), values)
        } else {
            // Return an error if the type is unsupported
            return py_err::<PyTypeError>(&format!("Unsupported value type in column '{}'", name.to_string()));
        };

        Ok(series)
    } else {
        // If the column contains only None values or is empty, return an error
        py_err::<PyValueError>(&format!("Column '{}' contains only None values or is empty", name.to_string()))
    }
}

/// Convert a Python dictionary of lists to a Rust Polars DataFrame.
///
/// This function takes a Python dictionary where each key corresponds to a list of values
/// and converts it into a Polars DataFrame. The length of all lists must be the same.
///
/// # Arguments
/// * `py` - The Python interpreter instance.
/// * `dict_of_lists` - The Python dictionary to convert.
///
/// # Returns
/// A `PyResultDF` containing a Rust Polars DataFrame if successful, or a Python error if failed.
fn py_dict_of_lists_to_rust_polars_df(py: Python, dict_of_lists: &PyAny) -> PyResultDF {
    // Extract the dictionary of lists from Python
    let dict_of_lists: DictOfLists = dict_of_lists.extract()
        .map_err(|_| py_err::<PyTypeError>("Structure of dictionary of lists is not correct."))?;

    // Determine the maximum column length (since empty lists may exist)
    let max_column_len = dict_of_lists.values().map(Vec::len).max().unwrap_or(0);

    // Create a vector of Series from the dictionary of lists
    let columns: Vec<Series> = dict_of_lists.into_iter()
        .map(|(name, values)| {
            if values.len() != max_column_len {
                return py_err::<PyValueError>("All lists must have the same length.");
            }
            extract_series_from_vec_of_optional_py_objects(py, &values, &name, max_column_len)
        })
        .collect::<Result<Vec<_>, _>>()?;

    // Create a Polars DataFrame from the Series
    DataFrame::new(columns).map_err(|e| py_err::<PyValueError>(&format!("Failed to create DataFrame: {}.", e)))
}

/// Convert a Python list of lists to a Rust Polars DataFrame.
///
/// This function takes a Python list where each inner list represents a column of values
/// and converts it into a Polars DataFrame. The list of column names must be provided separately.
///
/// # Arguments
/// * `py` - The Python interpreter instance.
/// * `list_of_lists` - The Python list of lists to convert.
/// * `columns` - The Python list of column names.
///
/// # Returns
/// A `PyResultDF` containing a Rust Polars DataFrame if successful, or a Python error if failed.
fn py_list_of_lists_to_rust_polars_df(py: Python, list_of_lists: &PyAny, columns: &PyAny) -> PyResultDF {
    // Extract the list of column names from Python
    let columns: Vec<String> = columns.extract()
        .map_err(|_| py_err::<PyTypeError>("List of columns is not correct."))?;

    // Extract the list of lists from Python
    let list_of_lists: Vec<Vec<Option<PyObject>>> = list_of_lists.extract()
        .map_err(|_| py_err::<PyTypeError>("Structure of list of lists is not correct."))?;

    // Determine the maximum column length
    let max_column_len = list_of_lists.iter().map(Vec::len).max().unwrap_or(0);

    // Ensure the number of columns matches the number of lists
    if columns.len() != list_of_lists.len() {
        return py_err::<PyValueError>("List of columns and list of lists have different lengths.");
    }

    // Create a vector of Series from the list of lists
    let df_columns: Vec<Series> = list_of_lists.into_iter()
        .zip(columns.into_iter())
        .map(|(values, name)| {
            if values.is_empty() {
                // If a column is empty, create a null Series
                Ok(Series::new_null(&name, max_column_len))
            } else if values.len() != max_column_len {
                py_err::<pyo3::exceptions::PyValueError, _>("All lists must have the same length.")
            } else {
                extract_series_from_vec_of_optional_py_objects(py, &values, &name, max_column_len)
            }
        })
        .collect::<Result<Vec<_>, _>>()?;

    // Create a Polars DataFrame from the Series
    DataFrame::new(df_columns).map_err(|e| py_err::<pyo3::exceptions::PyValueError, _>(&format!("Failed to create DataFrame: {}.", e)))
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

/// Get the DataFrame type from the specified module.
///
/// :param py: The Python interpreter instance.
/// :param module_name: The name of the module to import (e.g., "pandas" or "polars").
/// :return: The DataFrame type from the specified module.
fn get_dataframe_type<'py>(py: Python<'py>, module_name: &str) -> PyResult<&'py PyAny> {
    // Import the module
    let module = PyModule::import(py, module_name).map_err(|e| {
        let err_msg = format!("Failed to import {} module: {}.", module_name, e);
        PyErr::new::<pyo3::exceptions::PyImportError, _>(err_msg)
    })?;
    
    // Get the DataFrame type from the module
    module.getattr("DataFrame").map_err(|e| {
        let err_msg = format!("Failed to get DataFrame from module {}: {}.", module_name, e);
        PyErr::new::<pyo3::exceptions::PyAttributeError, _>(err_msg)
    })
}

/// Detect the type of the provided DataFrame (Pandas, Polars, dictionary, or list of lists).
///
/// This function checks if the input is a Pandas DataFrame, Polars DataFrame,
/// dictionary of lists, or list of lists and returns the corresponding enum variant.
///
/// # Arguments
/// * `py` - The Python interpreter instance.
/// * `df` - The Python object to identify.
///
/// # Returns
/// A `PyResult` containing the detected `OriginalDataType` or an error if the type is unsupported.
pub fn get_datatype(py: Python, df: &PyAny) -> PyResult<OriginalDataType> {
    // Get the Pandas and Polars DataFrame types
    let pandas_type = get_dataframe_type(py, "pandas")?;
    let polars_type = get_dataframe_type(py, "polars")?;

    // Check the type of the input object
    if df.is_instance(pandas_type)? {
        Ok(OriginalDataType::Pandas)
    } else if df.is_instance(polars_type)? {
        Ok(OriginalDataType::Polars)
    } else if df.is_instance(py.get_type::<pyo3::types::PyDict>())? {
        Ok(OriginalDataType::DictionaryOfLists)
    } else if df.is_instance(py.get_type::<pyo3::types::PyList>())? {
        Ok(OriginalDataType::ListOfLists)
    } else {
        // Return an error if the type is unsupported
        py_err::<PyTypeError>("Input is neither a Pandas DataFrame, Polars DataFrame, nor a dictionary of lists.")
    }
}

/// Convert a Python object to a Polars DataFrame based on its type.
///
/// This function takes a Python object (Pandas DataFrame, Polars DataFrame,
/// dictionary of lists, or list of lists) and converts it into a Rust Polars DataFrame.
///
/// # Arguments
/// * `py` - The Python interpreter instance.
/// * `data_type` - The detected type of the Python object.
/// * `df` - The Python object to convert.
/// * `columns` - Optional column names (only required for list of lists).
///
/// # Returns
/// A `PyResultDF` containing the Rust Polars DataFrame if successful, or a Python error if failed.
pub fn convert(py: Python, data_type: OriginalDataType, df: &PyAny, columns: Option<PyObject>) -> PyResultDF {
    match (data_type, columns) {
        // Error case: Column names must be provided for list of lists
        (OriginalDataType::ListOfLists, None) => {
            py_err::<pyo3::exceptions::PyValueError, _>("Column names must be provided for List of Lists.")
        },
        // Convert list of lists to Polars DataFrame, requiring column names
        (OriginalDataType::ListOfLists, Some(columns)) => {
            py_list_of_lists_to_rust_polars_df(py, df, columns.as_ref(py))
        },
        // Convert Pandas DataFrame to Polars DataFrame
        (OriginalDataType::Pandas, None) => {
            py_pandas_df_to_rust_polars_df(py, df)
        },
        // Convert Polars DataFrame to Rust Polars DataFrame
        (OriginalDataType::Polars, None) => {
            py_polars_df_to_rust_polars_df(py, df)
        },
        // Convert dictionary of lists to Polars DataFrame
        (OriginalDataType::DictionaryOfLists, None) => {
            py_dict_of_lists_to_rust_polars_df(py, df)
        },
        // Error case: Column names should not be provided for Pandas, Polars, or dictionary of lists
        (_, Some(_)) => {
            py_err::<pyo3::exceptions::PyValueError, _>("Column names should not be provided for Pandas, Polars, and Dict of Lists.")
        }
    }
}