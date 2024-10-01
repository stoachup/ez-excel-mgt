# Easy Excel Management

A Rust-powered Python module for efficient Excel file manipulation.

## Installation

Install the module using `pip` or your favorite package manager:

```bash
   pip install ez-excel-mgt
```

## Usage

### Creating and Filling an Excel File

Create an Excel file (`test.xlsx`) with a few rows and columns in a sheet named **"Sheet1"**.  
The sheet name must **exactly match** the name you'll use in the code. The **last row** in the sheet should contain the headers (column names) like `"Name"` and `"Age"`. Make sure the column names match exactly the column names in the DataFrame.

### Options

- **column**: List of column names to use with a list of lists. This is optional if your DataFrame already has columns (Pandas, Polars or Dict).
  
- **header_row**: Specifies the row containing the headers. Options:
  - "first": The first row is used as the header.
  - "last" (default): The last row is used as the header.
  - Integer > 0: The specific row number (1-based index) where the header is located.

- **overwrite**: If `True`, existing data will be overwritten. Default is `False`.

- **skip_nulls**: If `True`, null values in the data will be skipped. Default is `False`, which results in blanking cells for null values.

- **strict**: If `True`, enforces strict matching: all column names must match exactly.

### Example

```python
from typing import List, Dict, Any
from pathlib import Path

import pandas as pd
import polars as pl
import openpyxl
import ez_excel_mgt as ezex

# Create a template.xlsx file with a sheet named "Sheet1" and a header row at row 3
def create_template_excel(tmp_path: Path, 
                          sheet_name: str, 
                          metadata: List[str], 
                          data: Dict[str, List[Any]]) -> int:
    # Create a new workbook and add a sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    
    # Add metadata rows
    for i, row in enumerate(metadata):
        sheet[f"A{i+1}"] = row

    # Header row in row 3 (of the excel file)
    for j, (name, values) in enumerate(data.items()):
        sheet[f"{chr(65 + j)}{len(metadata)+1}"] = name
        for i, v in enumerate(values):
            sheet[f"{chr(65 + j)}{len(metadata)+i+1+1}"] = v

    # Write to the temporary file
    workbook.save(excel_path)
    
    return len(metadata) + 1


if __name__ == "__main__":
    sheet_name = "Example"
    excel_path = Path("example.xlsx")
    header_row = create_template_excel(excel_path,
                                       sheet_name, 
                                       ["A comment the first row", "Another in the second row"], 
                                       {"Name": ["Alice", "Bob", "Charlie"], 
                                        "Age": [25, 30, 35], 
                                        "Gender": ["F", "M", "M"],
                                        "City": ["New York", "London", "Paris"]})

    # Create a Pandas DataFrame
    pandas_df = pd.DataFrame({
        "Name": ["Anatole", "Erica", "Jules"],
        "Age": [85, 15, 95]
    })

    # Call the function to append the Pandas DataFrame to the existing sheet
    ezex.fill_sheet_with(pandas_df, str(excel_path), sheet_name, header_row=header_row)

    # Create a Polars DataFrame
    polars_df = pl.DataFrame({
        "Name": ["Philippe", "Paul"],
        "Age": [45, None],
        "Gender": ["M", "M"]
    })

    # Call the function to append the PolarsDataFrame to the existing sheet
    ezex.fill_sheet_with(polars_df, str(excel_path), sheet_name, header_row=header_row)

    # Create a dictionary with the data to be written to the Excel file
    dict_df = {
        "Name": ["Michel", "Amelie"],
        "Age": [35, 45],
        "Gender": ["M", "F"],
        "City": ["Paris", "London"]
    }

    # Read the Excel file into a Polars DataFrame
    ezex.fill_sheet_with(dict_df, str(excel_path), sheet_name, header_row=header_row)

    # Create a list of dictionaries with the data to be written to the Excel file
    # Column names must be provided separately
    columns = ["Age", "Gender", "City"]
    # skip_null allows to skip None values, which combined with overwrite enables filling gaps or replacing specific values like a mask
    list_df = [
        [None, None, None, None, None, None, None, 55, None, None],
        [None, None, None, "M", "F", "M", None, None, None, None],
        [None, None, None, "Brussels", "Madrid", "Berlin", "Lisbon", "Montreal", None, None]
    ]


    ezex.fill_sheet_with(list_df, str(excel_path), sheet_name, header_row=header_row, columns=columns, skip_null=True, overwrite=True)

    df = pl.read_excel(source=excel_path, sheet_name=sheet_name,
                       engine='xlsx2csv', engine_options={"skip_empty_lines": True, "skip_hidden_rows": False},
                       read_options={"skip_rows": 2, "has_header": True, "infer_schema_length": 0})
    print(df)
```

