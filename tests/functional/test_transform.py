# This project uses Poetry for dependency management.
from pathlib import Path
import openpyxl
from ez_excel_mgt import transform_range_from_to


def test_copy_range_between_files(create_test_excel_with_data_to_aggregate, create_empty_test_excel):
    """Test copy."""
    source_file_path, source_sheet_name, _header_row = create_test_excel_with_data_to_aggregate
    dest_file_path, dest_sheet_name, _header_row = create_empty_test_excel
    transform_range_from_to(source_file_path, source_sheet_name, ((2, 1), (3, 4)), 
                             dest_file_path, dest_sheet_name, (1, 1),
                             action='sum', mode='row')

    assert Path(dest_file_path).exists()

    # Load the modified Excel file and verify the contents
    workbook = openpyxl.load_workbook(dest_file_path)
    sheet = workbook[dest_sheet_name]

    # Assert that data is inserted with a header row (named columns)
    assert sheet["A1"].value == 10
    assert sheet["A2"].value == 26

