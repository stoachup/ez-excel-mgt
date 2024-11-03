use pyo3::prelude::*;
use once_cell::sync::Lazy;
use umya_spreadsheet::*;
use polars::prelude::*;
use std::collections::HashMap;
use std::path::Path;
use std::sync::Arc;
use log::*;

use crate::structs::*;
use crate::utils::aggregate::aggregate_range;
use crate::utils::excel::{index_to_excel, index_to_excel_col};
use crate::utils::py2rs::{get_datatype, convert, convert_anyvalue_to_string};
use crate::structs::{ExcelCell, ExcelRange, ExcelHeader};

#[pyclass]
pub struct ExcelTemplate {
    spreadsheet: Arc<Spreadsheet>,
    current_sheet_name: Option<String>,
    current_cell_in_current_sheet: Option<ExcelCell>,
}

impl ExcelTemplate {
    /// Internal function to load an Excel spreadsheet
    fn load_spreadsheet(file_path: &str) -> PyResult<Spreadsheet> {
        // Check if the file exists
        if !Path::new(file_path).exists() {
            return Err(PyErr::new::<pyo3::exceptions::PyFileNotFoundError, _>(format!("File not found: {:?}", file_path)));
        }

        reader::xlsx::read(Path::new(file_path)).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to read file: {:?}", e))
        })
    }
}

#[allow(dead_code)] // Suppress the warning for unused static
static LAZY_TEMPLATE: Lazy<Box<ExcelTemplate>> = Lazy::new(|| {
    Box::new(ExcelTemplate {
        spreadsheet: Arc::new(new_file()),
        current_sheet_name: None,
        current_cell_in_current_sheet: None,
    })
});

#[pymethods]
impl ExcelTemplate {
    /// Initializes a new ExcelTemplate by opening an existing file
    #[new]
    pub fn new(_py: Python, file_path: &str) -> PyResult<Self> {
        let spreadsheet = Arc::new(Self::load_spreadsheet(file_path)?);
        debug!("Spreadsheet loadedfrom {}", file_path);
        Ok(ExcelTemplate { spreadsheet, current_sheet_name: None, current_cell_in_current_sheet: None })
    }

    /// Adds a new sheet to the spreadsheet with a specified name
    pub fn add_sheet(&mut self, sheet_name: &str) -> PyResult<()> {
        let spreadsheet = Arc::get_mut(&mut self.spreadsheet)
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Cannot modify spreadsheet."))?;
        
        // Capture potential errors when adding a new sheet
        spreadsheet.new_sheet(sheet_name).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Failed to add sheet '{}': {:?}.", sheet_name, e))
        })?;
        Ok(())
    }

    pub fn goto_sheet(&mut self, sheet_name: &str, cell: Option<ExcelCell>) -> PyResult<()> {
        self.current_sheet_name = Some(sheet_name.to_string());
        self.current_cell_in_current_sheet = cell.clone();
        debug!("Going to sheet {} in cell {}", sheet_name, cell.map_or("None".to_string(), |c| c.range()));
        Ok(())
    }

    pub fn goto_cell(&mut self, cell: ExcelCell) -> PyResult<()> {
        self.current_cell_in_current_sheet = Some(cell.clone());
        debug!("Going to cell {}", cell.range());
        Ok(())
    }

    pub fn set_header_location(&mut self, header: ExcelHeader, mode: Mode) -> PyResult<()> {
        let spreadsheet = Arc::get_mut(&mut self.spreadsheet)
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Cannot modify spreadsheet."))?;
        
        let sheet_name = self.current_sheet_name
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No sheet specified."))?
            .to_string(); // Clone the string to avoid borrowing self

        let worksheet = spreadsheet.get_sheet_by_name_mut(&sheet_name).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Sheet '{}' not found.", sheet_name))
        })?;

        let header_location = match header {
            ExcelHeader::ExcelRange(r) => {
                let ((start_col, start_row), (_, _)) = r.idx();
                ExcelCell::Tuple((start_row, start_col))
            },
            ExcelHeader::ExcelCell(c) => c,
            ExcelHeader::First => ExcelCell::default(),
            ExcelHeader::Last => {
                let (col, row) = worksheet.get_highest_column_and_row();
                match mode {
                    Mode::Row => ExcelCell::Tuple((row, col)),
                    Mode::Column => ExcelCell::Tuple((col, row)),
                }
            }
        };
        self.current_cell_in_current_sheet = Some(header_location.clone());
        debug!("Headers expected in cell {} of {}", index_to_excel(header_location.idx().0, header_location.idx().1), sheet_name);

        Ok(())
    }

    pub fn set_value(&mut self, value: Value) -> PyResult<()> {
        // Store the current sheet name and cell in variables to avoid mutable borrow conflict
        let sheet_name = self.current_sheet_name
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No sheet specified."))?
            .to_string(); // Clone the string to avoid borrowing self

        let current_cell = self.current_cell_in_current_sheet
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>(format!("No cell specified in {}.", sheet_name)))?;
                
        // Now call write_cell with the extracted values
        self.write_cell(&sheet_name, current_cell.clone(), value)
    }

    /// Writes data to a specified cell in a given sheet
    pub fn write_cell(&mut self, sheet_name: &str, cell: ExcelCell, value: Value) -> PyResult<()> {
        let spreadsheet = Arc::get_mut(&mut self.spreadsheet)
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Cannot modify spreadsheet."))?;
        
        let worksheet = spreadsheet.get_sheet_by_name_mut(sheet_name).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Sheet {} not found.", sheet_name))
        })?;

        let (col, row) = cell.idx();        
        worksheet.get_cell_mut((col, row)).set_value(&value.value());
        debug!("Value {:?} set at {} in {}", value, cell.range(), sheet_name);
        Ok(())
    }

    pub fn remove_row_from(&mut self, sheet_name: &str, row: u32) -> PyResult<()> {
        let spreadsheet = Arc::get_mut(&mut self.spreadsheet)
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Cannot modify spreadsheet."))?;
        
        let worksheet = spreadsheet.get_sheet_by_name_mut(sheet_name).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Sheet {} not found.", sheet_name))
        })?;

        worksheet.remove_row(&row, &1);
        Ok(())
    }

    pub fn remove_rows_from(&mut self, sheet_name: &str, row: u32, num: u32) -> PyResult<()> {
        let spreadsheet = Arc::get_mut(&mut self.spreadsheet)
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Cannot modify spreadsheet."))?;
        
        let worksheet = spreadsheet.get_sheet_by_name_mut(sheet_name).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Sheet {} not found.", sheet_name))
        })?;

        worksheet.remove_row(&row, &num);
        Ok(())
    }

    /// Returns the list of sheet names
    pub fn sheet_names(&self) -> PyResult<Vec<String>> {
        let sheet_count = self.spreadsheet.as_ref().get_sheet_count();
        let mut names = Vec::with_capacity(sheet_count);
        
        for i in 0..sheet_count {
            let sheet = self.spreadsheet.as_ref().get_sheet(&i).ok_or_else(|| {
                PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("No sheet found at index {}.", i))
            })?;
            names.push(sheet.get_name().to_string());
        }
        
        Ok(names)
    }

    /// Saves the spreadsheet to a specified file path
    pub fn save(&self, file_path: &str) -> PyResult<()> {
        writer::xlsx::write(&self.spreadsheet, Path::new(file_path)).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to save file: {:?}.", e))
        })
    }

    pub fn copy_range_from(
        &mut self,
        source_file_path: &str,
        source_sheet_name: &str,
        source_range: ExcelRange,
        transpose: Option<bool>,
        coerce: Option<Coerce>,
    ) -> PyResult<()> {
        let spreadsheet = Arc::get_mut(&mut self.spreadsheet)
        .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Cannot modify spreadsheet."))?;
        
        let current_sheet_name = self.current_sheet_name
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No sheet specified. Use goto_sheet to set the sheet."))?
            .to_string(); // Clone the string to avoid borrowing self

        let worksheet = spreadsheet.get_sheet_by_name_mut(&current_sheet_name).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Sheet '{}' not found", current_sheet_name))
        })?;

        let current_cell = self.current_cell_in_current_sheet
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No cell specified. Use goto_cell to set the cell."))?;

        // Read the source workbook or return an error if it doesn't exist  
        let source_workbook = reader::xlsx::read(source_file_path).map_err(|e| {
            PyErr::new::<pyo3::exceptions::PyIOError, _>(format!("Failed to read Excel file: {}. Error: {:?}", source_file_path, e))
        })?;
        let source_sheet = source_workbook.get_sheet_by_name(source_sheet_name).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Source sheet '{}' not found.", source_sheet_name))
        })?;
        debug!("Source sheet {} found in {}", source_sheet_name, source_file_path);
        // Match on the SourceRange enum to handle both cases
        let ((start_col, start_row), (end_col, end_row)) = source_range.idx();

        let transpose = transpose.unwrap_or(false);
        let coerce = coerce.unwrap_or(Coerce::None);

        let (current_cell_col, current_cell_row) = current_cell.idx();
        // Copy the range from the source sheet to the destination sheet
        debug!("Copying range {} of {} to {} of {}", 
            source_range.range(), source_sheet_name, current_cell.range(), current_sheet_name);
        
        for col in start_col..=end_col {
            for row in start_row..=end_row {
                debug!("Processing cell {}", index_to_excel(col, row));
                if let Some(source_cell) = source_sheet.get_cell((col, row)) {
                    let original_value = source_cell.get_value().to_string();
                    debug!("Original value: {:?}", original_value);
                    let value = match coerce {
                        Coerce::None | Coerce::String => original_value.clone(),
                        _ => match original_value.parse::<f64>() {
                            Ok(value) => {
                                match coerce {
                                    Coerce::Integer => (value as i32).to_string(),
                                    _ => value.to_string(),
                                }
                            },
                            Err(_) => {
                                warn!("Value {:?} at {} is not a number. Ignored", original_value, index_to_excel(col, row));
                                String::new()
                            },
                        },
                    };
                    // Calculate destination cell coordinates, with optional transposing
                    let (d_col, d_row) = if transpose {
                        debug!("Transposing range");
                        (current_cell_col + row - start_row, current_cell_row + col - start_col)
                    } else {
                        (current_cell_col + col - start_col, current_cell_row + row - start_row)
                    };
                    // Attempt to set the value
                    worksheet.get_cell_mut((d_col, d_row)).set_value(&value);
                    debug!("Value {:?} taken from {} and set to {:?} at {}", 
                        original_value, index_to_excel(col, row), value, index_to_excel(d_col, d_row));
                }
                else {
                    debug!("Cell {} is empty", index_to_excel(col, row));
                }
            }
        }
        Ok(())
    }

    pub fn aggregate_range_from(
        &mut self,
        source_file_path: &str,
        source_sheet_name: &str,
        source_range: ExcelRange,
        action: Action,
        mode: Mode,        
    ) -> PyResult<()> {
        let spreadsheet = Arc::get_mut(&mut self.spreadsheet)
        .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Cannot modify spreadsheet"))?;
        
        let current_sheet_name = self.current_sheet_name
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No sheet specified. Use goto_sheet to set the sheet."))?
            .to_string(); // Clone the string to avoid borrowing self

        let worksheet = spreadsheet.get_sheet_by_name_mut(&current_sheet_name).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Sheet '{}' not found", current_sheet_name))
        })?;

        let current_cell = self.current_cell_in_current_sheet
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No cell specified. Use goto_cell to set the cell."))?;

        // Read the source workbook or return an error if it doesn't exist  
        let source_workbook = reader::xlsx::read(source_file_path).map_err(|_| {
            let err_msg = format!("Failed to read Excel file: {}. Check if the file exists and is readable.", source_file_path);
            PyErr::new::<pyo3::exceptions::PyFileNotFoundError, _>(err_msg)
        })?;
        let source_sheet = source_workbook.get_sheet_by_name(source_sheet_name)
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyValueError, _>("Source sheet not found"))?;
        debug!("Source sheet {} found in {}", source_sheet_name, source_file_path);
    
        // Match on the SourceRange enum to handle both cases
        let ((start_col, start_row), (end_col, end_row)) = source_range.idx();

        // Copy the range from the source sheet to the destination sheet
        let results =aggregate_range(source_sheet, start_row, start_col, end_row, end_col, action, mode.clone()).map_err(|e| {
            let err_msg = format!("Failed to aggregate range: {}", e);
            PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg)
        })?;
    
        debug!("Results: {:?}", results);

        let current_cell_idx = current_cell.idx();
        for (i, value) in results.iter().enumerate() {
            debug!("Pasting value {}: {} to sheet", i, value);
            match mode {
                Mode::Row => {
                    worksheet.get_cell_mut((current_cell_idx.0, current_cell_idx.1 + i as u32)).set_value(value.to_string());
                    debug!("Pasted value {} to cell {}", value, index_to_excel(current_cell_idx.0, current_cell_idx.1+ i as u32));
                },
                Mode::Column => {
                    worksheet.get_cell_mut((current_cell_idx.0 + i as u32, current_cell_idx.1)).set_value(value.to_string());
                    debug!("Pasted value {} to cell {}", value, index_to_excel(current_cell_idx.0 + i as u32, current_cell_idx.1));
                },
            }
        }
        Ok(())
    }    

    fn get_header_map(&self, mode: Mode) -> PyResult<HashMap<String, u32>> {
        let spreadsheet = Arc::as_ref(&self.spreadsheet);        
        let current_sheet_name = match self.current_sheet_name.as_ref() {
            Some(sheet_name) => sheet_name.clone(),
            None => return Err(PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No sheet specified. Use goto_sheet to set the sheet.")),
        };

        let worksheet = spreadsheet.get_sheet_by_name(&current_sheet_name).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Sheet '{}' not found", current_sheet_name))
        })?;

        let header_location = match self.current_cell_in_current_sheet.as_ref() {
            Some(cell) => cell,
            None => return Err(PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No cell specified. Use set_header_location to set the startingcell.")),
        };

        let (header_col, header_row) = header_location.idx();
        debug!("Getting headers starting from {} in mode {}", index_to_excel(header_col, header_row), mode);

        let mut header_map = HashMap::new();
        let first = match mode {
            Mode::Row => header_col,
            Mode::Column => header_row,
        };
        let last = match mode {
            Mode::Row => worksheet.get_highest_column(),
            Mode::Column => worksheet.get_highest_row(),
        };
        debug!("From {} to {}", first, last);
    
        for i in first..=last {
            let (col, row) = match mode {
                Mode::Row => (i, header_row),
                Mode::Column => (header_col, i),
            };  
            let col_name = worksheet.get_value((col, row)).to_string();
            debug!("Header {} in {}", col_name, index_to_excel(col, row));
            header_map.insert(col_name, col);
        }
    
        Ok(header_map)
    }

    pub fn fill_with(
        &mut self,
        py: Python,
        df: PyObject,
        columns: Option<PyObject>,
        mode: Option<Mode>,
        strict: Option<bool>,
        skip_null: Option<bool>,
        overwrite: Option<bool>,
    ) -> PyResult<()> {
        let data_type = get_datatype(py, df.as_ref(py))?;

        let df: DataFrame = convert(py, data_type, df.as_ref(py), columns)?;
        debug!("df: {:?}", df);

        // Convert the input to a Polars DataFrame
        let mode = mode.unwrap_or(Mode::Row);
        let skip_null = skip_null.unwrap_or(false);
        let strict = strict.unwrap_or(false);
        let overwrite = overwrite.unwrap_or(false);

        let header_map = self.get_header_map(mode.clone())?;

        let spreadsheet = Arc::get_mut(&mut self.spreadsheet)
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Cannot modify spreadsheet"))?;
        
        let current_sheet_name = self.current_sheet_name
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No sheet specified. Use goto_sheet to set the sheet."))?
            .to_string(); // Clone the string to avoid borrowing self

        let worksheet = spreadsheet.get_sheet_by_name_mut(&current_sheet_name).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Sheet '{}' not found", current_sheet_name))
        })?;

        let header_location = self.current_cell_in_current_sheet
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No cell specified. Use set_header_location to set the starting cell."))?;

        let (last_col, last_row) = worksheet.get_highest_column_and_row();
        let (header_col, header_row) = header_location.idx();
        let (first_col, first_row) = match mode {
            Mode::Row => (header_col.clone(), if overwrite { header_row + 1 } else { last_row + 1 }),
            Mode::Column => (if overwrite { header_col + 1 } else { last_col + 1 }, header_row.clone()),
        };

        debug!("Reading {}", mode);
        match mode {
            Mode::Row => {
                debug!("Header location: {}", index_to_excel(header_col, header_row));
                debug!("Last existing row: {}", last_row);
                debug!("First row to fill: {}", first_row);
            },
            Mode::Column => {
                debug!("Header location: {}", index_to_excel(header_col, header_row));
                debug!("Last existing column: {}", last_col);
                debug!("First column to fill: {}", first_col);
            },
        }
        self.goto_cell(ExcelCell::Tuple((first_row, first_col)))?;

        self.add_df_by_column_name(&df, header_map, mode, strict, skip_null)?;

        Ok(())
    }
    
}

// Methods that are not available in Python
impl ExcelTemplate {
    fn add_df_by_column_name(
        &mut self,
        df: &DataFrame,
        header_map: HashMap<String, u32>,
        mode: Mode,
        strict: bool,
        skip_null: bool,
    ) -> Result<(), PyErr> {
        let mut header_map = header_map.clone();
        let spreadsheet = Arc::get_mut(&mut self.spreadsheet)
        .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("Cannot modify spreadsheet"))?;
        
        let current_sheet_name = self.current_sheet_name
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No sheet specified. Use goto_sheet to set the sheet."))?
            .to_string(); // Clone the string to avoid borrowing self

        let worksheet = spreadsheet.get_sheet_by_name_mut(&current_sheet_name).ok_or_else(|| {
            PyErr::new::<pyo3::exceptions::PyValueError, _>(format!("Sheet '{}' not found", current_sheet_name))
        })?;

        let df_headers: Vec<String> = df.get_column_names().iter().map(|s| s.to_string()).collect(); // Convert to Vec<String>
        
        // Check for missing columns in DataFrame
        for col_name in header_map.keys() {
            if !df_headers.contains(col_name) {
                let err_msg = format!("Header '{}' in {} in the ExcelTemplate is missing in the DataFrame.", col_name, current_sheet_name);
                warn!("{}", err_msg);
                if strict {
                    return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg));
                }
            }
        }
    
        // Check for missing columns in the sheet
        for df_col in &df_headers {
            if !header_map.contains_key(df_col) {
                let err_msg = format!("Header '{}' is missing in {} in the ExcelTemplate.", df_col, current_sheet_name);
                warn!("{}", err_msg);
                if strict {
                    return Err(PyErr::new::<pyo3::exceptions::PyValueError, _>(err_msg));
                }
                else {
                    header_map.insert(df_col.to_string(), worksheet.get_highest_column() + 1);
                }
            }
        }
    
        // Proceed with adding the DataFrame to the sheet for matching columns
        let current_cell = self.current_cell_in_current_sheet
            .as_ref()
            .ok_or_else(|| PyErr::new::<pyo3::exceptions::PyRuntimeError, _>("No cell specified. Use set_header_location to set the startingcell."))?;
    
        let (current_col, current_row) = current_cell.idx();

        let height = df.height();
        for (header_name, idx) in header_map {
            debug!("Header {} in {}", header_name, idx);
            if let Some(series) = df.column(&header_name).ok() {
                for i in 0..height {
                    let value = series.get(i).unwrap();
                    if skip_null && value == AnyValue::Null {
                        continue;
                    } else {    
                        let cell_value = convert_anyvalue_to_string(value);
                        let (col, row) = match mode {
                            Mode::Row => (idx, current_row + i as u32),
                            Mode::Column => (current_col + i as u32, idx),
                        };
                        debug!("Column: {}, Row: {}", col, row);
                        worksheet.get_cell_mut((col, row)).set_value(cell_value.clone());
                        debug!("{}: {} = {}", header_name, index_to_excel(col, row), cell_value);
                    }
                }
            }
        }

        match mode {
            Mode::Row => {
                let last_row = worksheet.get_highest_row();
                let first_row_to_remove = current_row + height as u32;
                if first_row_to_remove <= last_row {
                    let num_rows_to_remove = last_row - first_row_to_remove + 1;
                    worksheet.remove_row(&first_row_to_remove, &num_rows_to_remove);
                }
            },
            Mode::Column => {
                let last_col = worksheet.get_highest_column();
                let first_col_to_remove = current_col + height as u32;
                if first_col_to_remove <= last_col {
                    let num_cols_to_remove = last_col - first_col_to_remove + 1;
                    worksheet.remove_column(&index_to_excel_col(first_col_to_remove), &num_cols_to_remove);
                }
            },
        };
    
        Ok(())
    }
}