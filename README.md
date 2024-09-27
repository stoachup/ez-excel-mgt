# Easy Excel Management

Rust-powered Python module for Excel manipulation

## Installation

Classic through pip or your favourite package manager:

```shell
pip install ez-excel-mgt
```

## Usage

Create an Excel file test.xlsx with a few rows and columns in a sheet named "Sheet1". Sheet name must match exactly the name you'll use in the code. Last row should contain the headers, i.e. the column names in the dataframe ("Name" and "Age" in this example). Pay attention to the column names: they must match exactly the column names in the dataframe.

```python
import pandas as pd
import polars as pl
import ez_excel_mgt as ezex

# Define the sheet name and Excel file path
sheet_name = "Sheet1"
excel_file_path = "template.xlsx"

# Create a Pandas DataFrame
df_pandas = pd.DataFrame({
    "Name": ["Alice", "Bob", "Charlie"],
    "Age": [25, 30, 35]
})

# Call the function to append the DataFrame to the existing sheet using column names located at row 3
ezex.fill_sheet_with(excel_file_path, sheet_name, df_pandas, named=True, header_row=3)

# Create a Polars DataFrame
df_polars = pl.DataFrame({
    "Name": ["Philippe", "Anatole"],
    "Age": [65, 90]
})

# Call the function to write the DataFrame to the existing sheet and overwrite existing data (but not deleting rows beyond)
ezex.fill_sheet_with(excel_file_path, sheet_name, df_polars, named=True, header_row=3, overwrite=True)

df_dict = {
    "Name": ["Mirlinda", "Emilie"],
    "Age": [35, 40]
}

# Call the function to write the DataFrame to the existing sheet and overwrite existing data as of row 6
ezex.fill_sheet_with(excel_file_path, sheet_name, df_dict, named=True, header_row=3, overwrite=True, start_row=6)
```

### Options

- `named`: if the dataframe columns are named (default: False).
- `header_row`: the row number of the column names (default: 1). Can be set to `last`, `first` or be an integer > 0.
- `overwrite`: if the data should be overwritten if it exists (default: False).
- `start_row`: the row number to start writing the data (must be greater than the header row). Can be set to `last`, `first` or be an integer > 0.
- `strict`: if numbers of columns must match (if not named) or if columns names must match (if named). Otherwise, if not named, the number of columns in the dataframe must be lower or equal to the number of columns in the excel sheet, else all columns of the dataframe must be found in the excel sheet.

```python
import pandas as pd
import ez_excel_mgt as ezex

# Create a Pandas DataFrame
df = pd.DataFrame({
    "Name": ["Alice", "Bob", "Charlie"],
    "Age": [25, 30, 35]
})

# Define the sheet name and Excel file path
sheet_name = "Sheet1"
excel_file_path = "template.xlsx"

# Call the function to write the DataFrame to the existing sheet
ezex.will_sheet_with(excel_file_path, sheet_name, df, named=True, header_row=3, overwrite=True, start_row=4)
```

