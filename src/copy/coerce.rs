#[allow(unused_imports)]
use log::{debug, info, warn};
use pyo3::prelude::*;

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
