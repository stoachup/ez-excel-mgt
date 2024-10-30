#[allow(unused_imports)]
use log::{debug, info, warn};
use umya_spreadsheet::{self, Worksheet};

use crate::transform::action::Action;
use crate::transform::mode::Mode;

// Helper to convert (col, row) into Excel-style coordinates like "A1"
fn to_excel_coords(col: u32, row: u32) -> String {
    let col_letter = (col as u8 + b'A' - 1) as char;
    format!("{}{}", col_letter, row)
}

// Aggregation function to calculate the result based on mode (Row or Column)
pub fn aggregate_range(
    source_sheet: &Worksheet,
    start_row: u32,
    start_col: u32,
    end_row: u32,
    end_col: u32,
    action: Action,
    mode: Mode,
) -> Result<Vec<f64>, Box<dyn std::error::Error>> {
    let mut sum_by_row: Vec<f64> = vec![0.0; (end_row - start_row + 1) as usize];   
    let mut sum_by_col: Vec<f64> = vec![0.0; (end_col - start_col + 1) as usize];

    for row in start_row..=end_row {
        for col in start_col..=end_col {
            if let Some(source_cell) = source_sheet.get_cell((col, row)) {
                let value = source_cell.get_value().to_string();
                match value.parse::<f64>() {
                    Ok(parsed_value) => {
                        debug!("Row: {}, Col: {}, Value: {}", row, col, parsed_value);
                        sum_by_row[row as usize - start_row as usize] += parsed_value;
                        sum_by_col[col as usize - start_col as usize] += parsed_value;
                    }
                    Err(_) => {
                        let cell_coord = to_excel_coords(col, row);
                        let err_msg = format!("Non-numeric value found in cell {}: '{}'", cell_coord, value);
                        return Err(err_msg.into());
                    }
                }
            }
        }
    }

    let sum = match mode {
        Mode::Row => sum_by_row,
        Mode::Column => sum_by_col,
    };
    debug!("Sum: {:?}", sum);
    let count = match mode {
        Mode::Row => end_col - start_col + 1,
        Mode::Column => end_row - start_row + 1,
    };

    debug!("Action: {:?}", action);
    match action {
        Action::Sum => Ok(sum),
        Action::Count => Ok(vec![count as f64; sum.len()]),
        Action::Average => {
            if count > 0 { 
                Ok(sum.iter().map(|&s| s / count as f64).collect()) 
            } 
            else { 
                Err("Division by zero".into()) 
            }
        }
    }
}
