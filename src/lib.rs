use env_logger::Builder;
#[allow(unused_imports)]
use log::{debug, info, warn, LevelFilter};
use pyo3::prelude::*;
use std::env;
use std::io::Write;

mod utils;
mod structs;
mod template;


/// Python module initialization function.
#[pymodule]
fn ez_excel_mgt(_py: Python, m: &PyModule) -> PyResult<()> {
    // Initialize logger only once
    init_logging();

    m.add_class::<template::ExcelTemplate>()?;

    Ok(())
}

/// Initialize logging for the module.
///
/// This function initializes logging by checking the environment variables `LOGLEVEL` and `RUST_LOG`.
/// If neither is set, it defaults to the "error" log level.
fn init_logging() {
    // First check `LOGLEVEL`, then fallback to `RUST_LOG`, or default to "error"
    let log_env = env::var("RUST_LOG").unwrap_or_else(|_| "error".to_string());

    let log_level = log_env.parse::<LevelFilter>().unwrap_or_else(|_| {
        warn!("Invalid log level: {}. Defaulting to 'error'.", log_env);
        LevelFilter::Error
    });

    // Directly initialize env_logger with the log level we determined
    Builder::new()
        .filter_level(log_level)
        .format(|buf, record| writeln!(buf, "[RUST:{}] - {}", record.level(), record.args()))
        .init();

    info!("Logging initialized with level: {}", log_level);
}