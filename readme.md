# Easy Excel Management

A Rust-powered Python module for efficient Excel file manipulation.

- **Filling**: Fill an existing sheet of an existing Excel file with data based on headers that are not necessarily in the first row of the sheet. Can handle Pandas, Polars and Dict of Lists. List of lists is also supported, provided the column names are provided too. Can handle overwriting existing data. Can skip nulls. Can handle strict matching of column names.

- **Copying**: Copy a range of cells from one sheet of an Excel file to a sheet in another Excel file. Can handle transposing the range.

- **Aggregating**: Aggregate a range of cells from one sheet to another. Can handle summing, counting, averaging per row or column.

## Installation

Install the module using `pip` or your favorite package manager:

```bash
   pip install ez-excel-mgt
```

## Usage

### Filling an existing sheet in an Excel file

Let's assume test.xls contains a sheet **"Example"** with with a few rows and columns. Let's assume the column names are contained in the third row. 

Data can be provided in a Pandas, Polars, Dict of Lists or List of Lists (with column names specified in another argument).

### Options

- **column**: List of column names to use with a list of lists. This is optional if your DataFrame already has columns (Pandas, Polars or Dict).
  
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
from ez_excel_mgt import ExcelTemplate

# Assuming a template exists as template.xlsx with a sheet named "Example" and headers in row 3
template = ExcelTemplate("template.xlsx")
template.goto_sheet("Example")
template.set_header_location((3, 1)) # Header in row 3, column A
# template.set_header_location('A3') works as well

# Assuming that headers are in row 3, with 'Name' in column A, 'Age' in column B, 'Gender' in column C
pandas_df = pd.DataFrame({
    "Name": ["Anatole", "Erica", "Jules"],
    "Age": [85, 15, 95]
})

# Fill the sheet with the Pandas DataFrame and get rid of existing data (overwrite=True)
template.fill_with(pandas_df, overwrite=True)

# Create a Polars DataFrame
polars_df = pl.DataFrame({
    "Name": ["Philippe", "Paul"],
    "Age": [45, None],
    "Gender": ["M", "M"]
})

# Fill the sheet with the Polars DataFrame at the end of the sheet
template.fill_with(polars_df, strict=True) # strict=True to enforce strict (one to one)matching of headers (template <-> data)

# Create a dictionary with the data to be written to the Excel file
dict_df = {
    "Name": ["Michel", "Amelie"],
    "Age": [35, 45],
    "Gender": ["M", "F"],
    "City": ["Paris", "London"]
}

template.fill_with(dict_df)

# Create a list of dictionaries with the data to be written to the Excel file
# Column names must be provided separately
columns = ["Age", "Gender", "City"]
# skip_null allows to skip None values, which combined with overwrite enables filling gaps or replacing specific values like a mask
list_df = [
    [None, None, None, None, None, None, None, 55, None, None],
    [None, None, None, "M", "F", "M", None, None, None, None],
    [None, None, None, "Brussels", "Madrid", "Berlin", "Lisbon", "Montreal", None, None]
]

template.fill_with(list_df, columns=columns, skip_null=True, overwrite=True)
```

### Copying a range of cells from one file/sheet to another file/sheet

Let's assume test.xls contains a sheet **"Example"** with with a few rows and columns. Let's assume the column names are contained in the first row. 

### Options

- **transpose**: If `True`, the range is transposed.

### Example

```python
from ez_excel_mgt import ExcelTemplate

# Assuming template.xlsx exists, with a sheet named "Example"
template = ExcelTemplate("template.xlsx")
template.goto_sheet("Example")
template.goto_cell((2, 1))

# Assuming that a file source.xlsx exists, with a sheet named "Source"
template.copy_range_from("source.xlsx", "Source", ((2, 1), (51, 4)))
```

### Aggregating a range of cells from one file/sheet to copying the result intoanother file/sheet

Only works with numbers! Can handle summing, counting, averaging per row or column.

Let's assume test.xls contains a sheet **"Example"** with with a few rows and columns. Let's assume the column names are contained in the first row. 

### Example

```python
from ez_excel_mgt import ExcelTemplate

# Assuming template.xlsx exists, with a sheet named "Example"
template = ExcelTemplate("template.xlsx")
template.goto_sheet("Example")
template.goto_cell((2, 1))

# Assuming that a file source.xlsx exists, with a sheet named "Source"
template.aggregate_range_from("source.xlsx", "Source", ((2, 1), (51, 4)), "sum", "row") # Aggregation can be made by row or column
```
