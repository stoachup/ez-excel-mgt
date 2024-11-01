#[allow(unused_imports)]
use log::{debug, info, warn};
use umya_spreadsheet::{self, Worksheet};

use crate::structs::{Action, Mode};

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
    let mut count_not_numeric_by_row: Vec<f64> = vec![0.0; (end_row - start_row + 1) as usize];
    let mut count_not_numeric_by_col: Vec<f64> = vec![0.0; (end_col - start_col + 1) as usize];

    for row in start_row..=end_row {
        for col in start_col..=end_col {
            if let Some(source_cell) = source_sheet.get_cell((col, row)) {
                let value = source_cell.get_value().to_string();
                match value.parse::<f64>() {
                    Ok(parsed_value) => {
                        debug!("Row: {}, Col: {}, Value: {}", row, col, parsed_value);
                        sum_by_row[row as usize - start_row as usize] += parsed_value;
                        sum_by_col[col as usize - start_col as usize] += parsed_value;
                        count_not_numeric_by_row[row as usize - start_row as usize] += 1.0;
                        count_not_numeric_by_col[col as usize - start_col as usize] += 1.0;
                    }
                    Err(_) => {
                        warn!("Non-numeric value found in cell {}: '{}'", to_excel_coords(col, row), value);
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
        Mode::Row => count_not_numeric_by_row,
        Mode::Column => count_not_numeric_by_col,
    };
    debug!("Count: {:?}", count);

    debug!("Action: {:?}", action);
    match action {
        Action::Sum => Ok(sum),
        Action::Count => Ok(count),
        Action::Average => {
            if count.iter().any(|&c| c > 0.0) { 
                Ok(sum.iter().zip(count.iter()).map(|(&s, &c)| s / c).collect()) 
            } 
            else { 
                Err("Division by zero".into()) 
            }
        }
    }
}
