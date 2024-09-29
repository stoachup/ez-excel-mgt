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
ezex.add_df_to_sheet(df, sheet_name, excel_file_path, header_row=3, overwrite=True)
