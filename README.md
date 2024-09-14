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
import ez_excel_mgt

# Create a Pandas DataFrame
df = pd.DataFrame({
    "Name": ["Alice", "Bob", "Charlie"],
    "Age": [25, 30, 35]
})

# Define the sheet name and Excel file path
sheet_name = "Sheet1"
excel_file_path = "test.xlsx"

# Call the function to write the DataFrame to the existing sheet
ez_excel_mgt.add_df_to_sheet(df, sheet_name, excel_file_path)
```

