#[allow(unused_imports)]
use log::{debug, info, warn};
use pyo3::prelude::*;

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
