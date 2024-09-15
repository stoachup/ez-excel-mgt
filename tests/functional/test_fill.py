import polars as pl
import pytest
import openpyxl
from ez_excel_mgt import fill_sheet_with


def test_fill_sheet_with_unnamed(create_test_excel):
    """Test inserting a Polars DataFrame into an existing sheet with `named` as False."""
    excel_path = create_test_excel
    
    # Create a sample Polars DataFrame
    df = pl.DataFrame({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
        "Gender": ["F", "M"]
    })
    
    # Call the function to add the DataFrame without named columns
    fill_sheet_with(str(excel_path), "Sheet1", df, named=False)

    # Load the modified Excel file and verify the contents
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook["Sheet1"]
    
    # Assert that data is inserted without a header row (unnamed columns)
    assert sheet["A6"].value == "Alice"
    assert sheet["B6"].value == 25
    assert sheet["C6"].value == "F"
    assert sheet["A7"].value == "Bob"
    assert sheet["B7"].value == 30
    assert sheet["C7"].value == "M"


def test_fill_sheet_with_named(create_test_excel):
    """Test inserting a Polars DataFrame with named columns (`named=True`)."""
    excel_path = create_test_excel
    
    # Create a sample Polars DataFrame
    df = pl.DataFrame({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
        "Gender": ["F", "M"]
    })
    
    # Call the function to add the DataFrame with named columns
    fill_sheet_with(str(excel_path), "Sheet1", df, named=True, header_row=3)

    # Load the modified Excel file and verify the contents
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook["Sheet1"]
    
    # Assert that data is inserted with a header row (named columns)
    assert sheet["A3"].value == "Name"
    assert sheet["B3"].value == "Age"
    assert sheet["C3"].value == "Gender"
    assert sheet["A6"].value == "Alice"
    assert sheet["B6"].value == 25
    assert sheet["C6"].value == "F"
    assert sheet["A7"].value == "Bob"
    assert sheet["B7"].value == 30
    assert sheet["C7"].value == "M"


def test_fill_sheet_with_unnamed_start_row(create_test_excel):
    """Test inserting a DataFrame starting at a custom row."""
    excel_path = create_test_excel
    
    # Create a sample Polars DataFrame
    df = pl.DataFrame({
        "Name": ["Charlie", "Diana"],
        "Age": [35, 40],
        "Gender": ["M", "F"]
    })

    # Add the DataFrame starting at row 6 (after the prepopulated data)
    fill_sheet_with(str(excel_path), "Sheet1", df, named=False, start_row=40)

    # Load the modified Excel file and verify the contents
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook["Sheet1"]
    
    # Assert that data is inserted starting from row 6
    assert sheet["A40"].value == "Charlie"
    assert sheet["B40"].value == 35
    assert sheet["C40"].value == "M"
    assert sheet["A41"].value == "Diana"
    assert sheet["B41"].value == 40
    assert sheet["C41"].value == "F"


def test_fill_sheet_with_unnamed_overwrite(create_test_excel):
    """Test inserting a DataFrame with overwrite=True."""
    excel_path = create_test_excel
    
    # Overwrite with a new DataFrame
    df2 = pl.DataFrame({
        "Name": ["Charlie", "Diana"],
        "Age": [35, 40],
        "Gender": ["M", "F"]
    })
    fill_sheet_with(str(excel_path), "Sheet1", df2, named=False, overwrite=True)

    # Load the modified Excel file and verify the new contents
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook["Sheet1"]
    
    # Assert that the data was overwritten correctly
    assert sheet["A1"].value == "Name"
    assert sheet["B1"].value == "Age"
    assert sheet["C1"].value == "Gender"
    assert sheet["A2"].value == "Charlie"
    assert sheet["B2"].value == 35
    assert sheet["C2"].value == "M"
    assert sheet["A3"].value == "Diana"
    assert sheet["B3"].value == 40
    assert sheet["C3"].value == "F"


def test_fill_sheet_with_named_at(create_test_excel_commented):
    """Test inserting a DataFrame with a custom header row."""
    excel_path = create_test_excel_commented
    
    # Create a sample Polars DataFrame
    df = pl.DataFrame({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
        "Gender": ["F", "M"]
    })

    # Call the function to add the DataFrame starting with a custom header row
    fill_sheet_with(str(excel_path), "Sheet1", df, named=True, header_row=3, start_row=5)

    # Load the modified Excel file and verify the contents
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook["Sheet1"]
    
    # Assert that data is inserted with the custom header row (2)
    assert sheet["A3"].value == "Name"
    assert sheet["B3"].value == "Age"
    assert sheet["C3"].value == "Gender"
    assert sheet["A5"].value == "Alice"
    assert sheet["B5"].value == 25
    assert sheet["C5"].value == "F"
    assert sheet["A6"].value == "Bob"
    assert sheet["B6"].value == 30
    assert sheet["C6"].value == "M"


def test_fill_sheet_with_start_before_header(create_test_excel):
    """Test behavior when column names in the DataFrame don't match the ones in the Excel file."""
    excel_path = create_test_excel
    
    # Create a DataFrame with mismatching column names
    df = pl.DataFrame({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
        "Sex": ["F", "M"]
    })

    # Expect an error or handle the mismatch gracefully (depending on the implementation)
    with pytest.raises(ValueError, match="start_row must be greater than header_row"):
        fill_sheet_with(str(excel_path), "Sheet1", df, named=True, header_row=3, start_row=1)


def test_fill_sheet_with_named_unfound_df_column(create_test_excel):
    """Test behavior when a column name in the DataFrame is not found in the Excel file."""
    excel_path = create_test_excel
    
    # Create a DataFrame with mismatching column names
    df = pl.DataFrame({
        "Full Name": ["Alice", "Bob"],
        "Age": [25, 30],
        "Sex": ["F", "M"]
    })

    # Expect an error or handle the mismatch gracefully (depending on the implementation)
    with pytest.raises(ValueError, match="Column 'Full Name' in DataFrame is missing in the template"):
        fill_sheet_with(str(excel_path), "Sheet1", df, named=True)


def test_fill_sheet_with_named_unfound_excel_column(create_test_excel):
    """Test behavior when a column name in excel is not found in the DataFrame."""
    excel_path = create_test_excel
    
    # Create a DataFrame with mismatching column names
    df = pl.DataFrame({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
    })

    try:
        # Call the method and check it doesn't raise any exceptions
        fill_sheet_with(str(excel_path), "Sheet1", df, named=True, header_row=3)
    except Exception as e:
        pytest.fail(f"fill_sheet_with raised an exception: {e}")



def test_fill_sheet_with_named_unfound_excel_column_strict(create_test_excel):
    """Test behavior when a column name in excel is not found in the DataFrame."""
    excel_path = create_test_excel
    
    # Create a DataFrame with mismatching column names
    df = pl.DataFrame({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
    })

    # Expect an error or handle the mismatch gracefully (depending on the implementation)
    with pytest.raises(ValueError, match="Column 'Gender' in excel is missing in DataFrame"):
        fill_sheet_with(str(excel_path), "Sheet1", df, named=True, header_row=3, strict=True)


