# This project uses Poetry for dependency management.
from pathlib import Path
import openpyxl
from ez_excel_mgt import copy_range_between_files


def test_copy_range_between_files(create_test_excel, create_empty_test_excel):
    """Test copy."""
    source_file_path, source_sheet_name, _header_row = create_test_excel
    dest_file_path, dest_sheet_name, _header_row = create_empty_test_excel
    copy_range_between_files(source_file_path, source_sheet_name, ((1, 1), (5, 3)), 
                             dest_file_path, dest_sheet_name, (1, 1))

    assert Path(dest_file_path).exists()

    # Load the modified Excel file and verify the contents
    workbook = openpyxl.load_workbook(dest_file_path)
    sheet = workbook[dest_sheet_name]

    # Assert that data is inserted with a header row (named columns)
    assert sheet["A1"].value == "First row"
    assert sheet["A2"].value == "Second row"
    assert sheet["B4"].value == 25
    assert sheet["B5"].value == 26

def test_transpose_range_between_files(create_test_excel, create_empty_test_excel):
    """Test copy."""
    source_file_path, source_sheet_name, _header_row = create_test_excel
    dest_file_path, dest_sheet_name, _header_row = create_empty_test_excel
    copy_range_between_files(source_file_path, source_sheet_name, ((1, 1), (5, 3)), 
                             dest_file_path, dest_sheet_name, (1, 1),
                             transpose=True)

    assert Path(dest_file_path).exists()

    # Load the modified Excel file and verify the contents
    workbook = openpyxl.load_workbook(dest_file_path)
    sheet = workbook[dest_sheet_name]

    # Assert that data is inserted with a header row (named columns)
    assert sheet["A1"].value == "First row"
    assert sheet["B1"].value == "Second row"
    assert sheet["D2"].value == 25
    assert sheet["E2"].value == 26
