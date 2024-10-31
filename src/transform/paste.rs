use log::debug;
use umya_spreadsheet::{self, Worksheet};

use crate::transform::Mode;

pub fn paste_vector_to_sheet(values: Vec<f64>, sheet: &mut Worksheet, start_row: u32, start_col: u32, mode: Mode) -> Result<(), Box<dyn std::error::Error>> {
    for (i, value) in values.iter().enumerate() {
        debug!("Pasting value {}: {} to sheet", i, value);
        match mode {
            Mode::Row => {
                sheet.get_cell_mut((start_col, start_row + i as u32)).set_value(value.to_string());
                debug!("Pasted value {} to cell ({}, {})", value, start_col - 1 + 1, start_row - 1 + i as u32);
            },
            Mode::Column => {
                sheet.get_cell_mut((start_col + i as u32, start_row)).set_value(value.to_string());
                debug!("Pasted value {} to cell ({}, {})", value, start_col - 1 + i as u32, start_row - 1);
            },
        }
    }
    Ok(())
}