#[allow(unused_imports)]
use log::{debug, info, warn};
use pyo3::prelude::*;

// Define the Mode enum
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
