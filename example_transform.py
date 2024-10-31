import random
import polars as pl
import openpyxl
import ez_excel_mgt as ezex

# Function to convert a number to Excel column name
def excel_column(row: int, col: int) -> str:
    column_name = ""
    while row > 0:
        row, remainder = divmod(row - 1, 26)
        column_name = chr(65 + remainder) + column_name
    return f"{column_name}{col}"

if __name__ == "__main__":
    # Create a polars dataframe with 50 rows and 4 columns name Col 1, Col 2, Col 3, Col 4
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Example"
    for i in range(4):
        sheet[excel_column(i+1, 1)] = f"Col {i+1}"
    workbook.save("example_copy.xlsx")
    
    df = pl.DataFrame({
        "Col 1": [random.random() for _ in range(50)], 
        "Col 2": [random.random() for _ in range(50)], 
        "Col 3": [random.random() for _ in range(50)],
        "Col 4": [random.random() for _ in range(50)]
    })
    ezex.fill_sheet_with(df, "example_copy.xlsx", "Example")
    df = pl.read_excel(source="example_copy.xlsx", sheet_name="Example",
                       engine='xlsx2csv', engine_options={"skip_empty_lines": True, "skip_hidden_rows": False},
                       read_options={"has_header": True, "infer_schema_length": 0})
    print(df)

    # Open the destination excel file using openpyxl and create a new sheet named "Result", containing 50 headers in the first row named Row 1, Row 2, etc.
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Result"

    # Update the loop to use the new function for naming columns
    for i in range(4):
        sheet[excel_column(i+1, 1)] = f"Sum {i+1}"
    workbook.save("result_transform.xlsx")
    
    # Call the function to copy the range of cells from the source file to the destination file (row, col; starting at 1)
    ezex.transform_range_from_to("example_copy.xlsx", "Example", ((2, 1), (51, 4)), 
                                 "result_transform.xlsx", "Result", (2, 1), "sum", "col")

    df = pl.read_excel(source="result_transform.xlsx", sheet_name="Result",
                       engine='xlsx2csv', engine_options={"skip_empty_lines": True, "skip_hidden_rows": False},
                       read_options={"has_header": True, "infer_schema_length": 0})
    print(df)