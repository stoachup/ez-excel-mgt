# Easy Excel Management

A Rust-powered Python module for efficient Excel file manipulation.

- **Filling**: Fill an existing sheet of an existing Excel file with data based on headers that are not necessarily in the first row of the sheet. Can handle Pandas, Polars and Dict of Lists. List of lists is also supported, provided the column names are provided too. Can handle overwriting existing data. Can skip nulls. Can handle strict matching of column names.

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
  
- **header_row**: Specifies the row containing the headers. Options:
  - "first": The first row is used as the header.
  - "last" (default): The last row is used as the header.
  - Integer > 0: The specific row number (1-based index) where the header is located.

- **overwrite**: If `True`, existing data will be overwritten. Default is `False`.

- **skip_nulls**: If `True`, null values in the data will be skipped. Default is `False`, which results in blanking cells for null values.

- **strict**: If `True`, enforces strict matching: all column names must match exactly.

### Example

!md_include "../example.py"

