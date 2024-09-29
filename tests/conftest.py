import pytest
import openpyxl


@pytest.fixture(scope="function")
def create_test_excel(tmp_path):
    """Fixture to create a temporary Excel file for testing."""
    # Create a temporary file path for the Excel file
    excel_path = tmp_path / "test.xlsx"
    
    # Create a new workbook and add a sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    
    # Prepopulate the sheet with some data
    sheet["A1"] = "First row"
    sheet["A2"] = "Second row"
    sheet["A3"] = "Name"
    sheet["B3"] = "Age"
    sheet["C3"] = "Gender"
    sheet["A4"] = "Irène"
    sheet["B4"] = 25
    sheet["A5"] = "Matthieu"
    sheet["B5"] = 26

    # Write to the temporary file
    workbook.save(excel_path)
    
    return excel_path, 3


@pytest.fixture(scope="function")
def create_test_excel_commented(tmp_path):
    """Fixture to create a temporary Excel file for testing with comments after the header row."""
    # Create a temporary file path for the Excel file
    excel_path = tmp_path / "test.xlsx"
    
    # Create a new workbook and add a sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    
    # Prepopulate the sheet with some data
    sheet["A1"] = "First row"
    sheet["A2"] = "Second row"
    sheet["A3"] = "Name"
    sheet["B3"] = "Age"
    sheet["C3"] = "Gender"
    sheet["A4"] = "Commented name"
    sheet["B4"] = "Commented age"
    sheet["C4"] = "Commented gender"

    # Write to the temporary file
    workbook.save(excel_path)
    
    return excel_path
