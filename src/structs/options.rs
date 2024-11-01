#[allow(unused_imports)]
use log::{debug, info, warn};
use pyo3::prelude::*;
use std::fmt;

// Define an enum to represent different value types
#[derive(Debug, Clone)]
pub enum Value {
    None,
    Int(i32),
    Boolean(bool),
    Float(f64),
    String(String),
}

// Implement FromPyObject for SourceRange
impl<'source> pyo3::FromPyObject<'source> for Value {
    fn extract(obj: &'source PyAny) -> PyResult<Self> {
        if obj.is_none() {
            return Ok(Value::None); // Handle None case
        }
        if let Ok(int) = obj.extract::<i32>() {
            Ok(Value::Int(int))
        } 
        else if let Ok(boolean) = obj.extract::<bool>() {
            Ok(Value::Boolean(boolean))
        } 
        else if let Ok(float) = obj.extract::<f64>() {
            Ok(Value::Float(float))
        } 
        else if let Ok(string) = obj.extract::<String>() {
            Ok(Value::String(string))
        } 
        else {
            Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "Invalid input type. Expected an integer, boolean, float, string, or None.",
            ))
        }
    }
}

impl IntoPy<PyObject> for Value {
    fn into_py(self, py: Python) -> PyObject {
        match self {
            Value::None => py.None(),
            Value::Int(v) => v.into_py(py),
            Value::Boolean(v) => v.into_py(py),
            Value::Float(v) => v.into_py(py),
            Value::String(v) => v.into_py(py),
        }
    }
}

impl Value {
    pub fn value(&self) -> String {
        match self {
            Value::None => "".to_string(),
            Value::Int(v) => v.to_string(),
            Value::Boolean(v) => {
                if *v {
                    "TRUE".to_string()
                } else {
                    "FALSE".to_string()
                }
            },
            Value::Float(v) => v.to_string(),
            Value::String(v) => v.clone(),
        }
    }
}

// Define the Mode enum
#[derive(Debug, Clone)]
pub enum Coerce {
    None,
    Integer,
    Float,
    String,
}

// Implement conversion from Python to Rust enum
impl FromPyObject<'_> for Coerce {
    fn extract(obj: &PyAny) -> PyResult<Self> {
        let mode_str: String = obj.extract()?;
        match mode_str.as_str() {
            "int" | "integer" => Ok(Coerce::Integer),
            "float" | "double" => Ok(Coerce::Float),
            "str" | "string" => Ok(Coerce::String),
            "none" => Ok(Coerce::None),
            _ => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "Invalid mode. Use 'int/integer' or 'float/double' or 'str/string'.",
            )),
        }
    }
}

// Implement conversion from Rust enum to Python
impl IntoPy<PyObject> for Coerce {
    fn into_py(self, py: Python) -> PyObject {
        match self {
            Coerce::Integer => "integer".into_py(py),
            Coerce::Float => "float".into_py(py),
            Coerce::String => "string".into_py(py),
            Coerce::None => "none".into_py(py),
        }
    }
}

impl Default for Coerce {
    fn default() -> Self {
        Coerce::None
    }
}


// Define the Action enum
#[derive(Debug, Clone)]
pub enum Action {
    Sum,
    Count,
    Average,
}

// Implement conversion from Python to Rust enum
impl FromPyObject<'_> for Action {
    fn extract(obj: &PyAny) -> PyResult<Self> {
        let mode_str: String = obj.extract()?;
        match mode_str.as_str() {
            "sum" => Ok(Action::Sum),
            "count" => Ok(Action::Count),
            "avg" => Ok(Action::Average),
            _ => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "Invalid mode. Use 'row' or 'col'.",
            )),
        }
    }
}

// Implement conversion from Rust enum to Python
impl IntoPy<PyObject> for Action {
    fn into_py(self, py: Python) -> PyObject {
        match self {
            Action::Sum => "sum".into_py(py),
            Action::Count => "count".into_py(py),
            Action::Average => "avg".into_py(py),
        }
    }
}


impl Default for Action {
    fn default() -> Self {
        Action::Sum
    }
}

// Define the Mode enum
#[derive(Debug, Clone)]
pub enum Mode {
    Row,
    Column,
}

// Implement conversion from Python to Rust enum
impl FromPyObject<'_> for Mode {
    fn extract(obj: &PyAny) -> PyResult<Self> {
        let mode_str: String = obj.extract()?;
        match mode_str.as_str() {
            "row" => Ok(Mode::Row),
            "col" => Ok(Mode::Column),
            _ => Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(
                "Invalid mode. Use 'row' or 'col'.",
            )),
        }
    }
}

// Implement conversion from Rust enum to Python
impl IntoPy<PyObject> for Mode {
    fn into_py(self, py: Python) -> PyObject {
        match self {
            Mode::Row => "row".into_py(py),
            Mode::Column => "col".into_py(py),
        }
    }
}

impl Default for Mode {
    fn default() -> Self {
        Mode::Row
    }
}

// Implement Display for Mode
impl fmt::Display for Mode {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        match self {
            Mode::Row => write!(f, "row per row"),
            Mode::Column => write!(f, "column per column"),
        }
    }
}