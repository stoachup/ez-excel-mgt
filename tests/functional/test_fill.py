import polars as pl
import pandas as pd
import pytest
import openpyxl
from ez_excel_mgt import fill_sheet_with


def generate_test_data(data, data_type="polars"):
    """Utility to generate test data in different formats (Polars, Pandas, or dict of lists)."""
    
    if data_type == "polars":
        return pl.DataFrame(data)
    elif data_type == "pandas":
        return pd.DataFrame(data)
    elif data_type == "dict":
        return data
    else:
        raise ValueError(f"Unsupported data type: {data_type}")
    

@pytest.mark.parametrize("data_type", ["polars", "pandas", "dict"])
def test_fill_sheet_with_unnamed(create_test_excel, data_type):
    """Test inserting a Polars DataFrame into an existing sheet with `named` as False."""
    excel_path = create_test_excel
    
    # Create a sample Polars DataFrame
    df = generate_test_data({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
        "Gender": ["F", "M"]
    }, data_type)
    
    if data_type == "dict":
        with pytest.raises(TypeError, match="As you are using a dictionary, you must specify named as True and the header_row position if not the last row in the sheet"):
            fill_sheet_with(str(excel_path), "Sheet1", df, named=False)

    else:
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


@pytest.mark.parametrize("data_type", ["polars", "pandas", "dict"])
def test_fill_sheet_with_named(create_test_excel, data_type):
    """Test inserting a Polars DataFrame with named columns (`named=True`)."""
    excel_path = create_test_excel
    
    # Create a sample Polars DataFrame
    df = generate_test_data({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
        "Gender": ["F", "M"]
    }, data_type)
    
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


@pytest.mark.parametrize("data_type", ["polars", "pandas", "dict"])
def test_fill_sheet_with_unnamed_start_row(create_test_excel, data_type):
    """Test inserting a DataFrame starting at a custom row."""
    excel_path = create_test_excel
    
    # Create a sample Polars DataFrame
    df = generate_test_data({
        "Name": ["Charlie", "Diana"],
        "Age": [35, 40],
        "Gender": ["M", "F"]
    }, data_type)

    if data_type == "dict":
        with pytest.raises(TypeError, match="As you are using a dictionary, you must specify named as True and the header_row position if not the last row in the sheet"):
            fill_sheet_with(str(excel_path), "Sheet1", df, named=False)

    else:
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


@pytest.mark.parametrize("data_type", ["polars", "pandas", "dict"])
def test_fill_sheet_with_unnamed_overwrite(create_test_excel, data_type):
    """Test inserting a DataFrame with overwrite=True."""
    excel_path = create_test_excel
    
    # Overwrite with a new DataFrame
    df = generate_test_data({
        "Name": ["Charlie", "Diana"],
        "Age": [35, 40],
        "Gender": ["M", "F"]
    }, data_type)

    if data_type == "dict":
        with pytest.raises(TypeError, match="As you are using a dictionary, you must specify named as True and the header_row position if not the last row in the sheet"):
            fill_sheet_with(str(excel_path), "Sheet1", df, named=False, overwrite=True)

    else:
        fill_sheet_with(str(excel_path), "Sheet1", df, named=False, overwrite=True)

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


@pytest.mark.parametrize("data_type", ["polars", "pandas", "dict"])
def test_fill_sheet_with_named_at(create_test_excel_commented, data_type):
    """Test inserting a DataFrame with a custom header row."""
    excel_path = create_test_excel_commented
    
    # Create a sample Polars DataFrame
    df = generate_test_data({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
        "Gender": ["F", "M"]
    }, data_type)

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


@pytest.mark.parametrize("data_type", ["polars", "pandas", "dict"])
def test_fill_sheet_with_start_before_header(create_test_excel, data_type):
    """Test behavior when column names in the DataFrame don't match the ones in the Excel file."""
    excel_path = create_test_excel
    
    # Create a DataFrame with mismatching column names
    df = generate_test_data({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
        "Sex": ["F", "M"]
    }, data_type)

    # Expect an error or handle the mismatch gracefully (depending on the implementation)
    with pytest.raises(ValueError, match="start_row must be greater than header_row"):
        fill_sheet_with(str(excel_path), "Sheet1", df, named=True, header_row=3, start_row=1)


@pytest.mark.parametrize("data_type", ["polars", "pandas", "dict"])
def test_fill_sheet_with_named_unfound_df_column(create_test_excel, data_type):
    """Test behavior when a column name in the DataFrame is not found in the Excel file."""
    excel_path = create_test_excel
    
    # Create a DataFrame with mismatching column names
    df = generate_test_data({
        "Full Name": ["Alice", "Bob"],
        "Age": [25, 30],
        "Sex": ["F", "M"]
    }, data_type)

    # Expect an error or handle the mismatch gracefully (depending on the implementation)
    with pytest.raises(ValueError, match="Column 'Full Name' is missing in the template"):
        fill_sheet_with(str(excel_path), "Sheet1", df, named=True)


@pytest.mark.parametrize("data_type", ["polars", "pandas", "dict"])
def test_fill_sheet_with_named_unfound_excel_column(create_test_excel, data_type):
    """Test behavior when a column name in excel is not found in the DataFrame."""
    excel_path = create_test_excel
    
    # Create a DataFrame with mismatching column names
    df = generate_test_data({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
    }, data_type)

    try:
        # Call the method and check it doesn't raise any exceptions
        fill_sheet_with(str(excel_path), "Sheet1", df, named=True, header_row=3)
    except Exception as e:
        pytest.fail(f"fill_sheet_with raised an exception: {e}")



@pytest.mark.parametrize("data_type", ["polars", "pandas", "dict"])
def test_fill_sheet_with_named_unfound_excel_column_strict(create_test_excel, data_type):
    """Test behavior when a column name in excel is not found in the DataFrame."""
    excel_path = create_test_excel
    
    # Create a DataFrame with mismatching column names
    df = generate_test_data({
        "Name": ["Alice", "Bob"],
        "Age": [25, 30],
    }, data_type)

    # Expect an error or handle the mismatch gracefully (depending on the implementation)
    with pytest.raises(ValueError, match="Column 'Gender' in excel is missing in DataFrame"):
        fill_sheet_with(str(excel_path), "Sheet1", df, named=True, header_row=3, strict=True)


