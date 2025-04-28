# tests/test_deficiency_summary.py
import pytest
from unittest.mock import MagicMock, patch

from converters.deficiency_summary import DeficiencySummaryConverter

@pytest.fixture
def mock_input_output():
    # Create mocks for input_sheet and output_wb
    input_sheet = MagicMock()
    output_sheet = MagicMock()
    output_wb = MagicMock()

    # Mock data that input_sheet.range().value would return
    input_data = [
        ["Device 1", "Deficiency 1", None, None, None, None, None, None, None, None, None, None],
        ["Device 2", "Deficiency 2", None, None, None, None, None, None, None, None, None, None],
    ]
    input_sheet.range.return_value.value = input_data

    # Mock output workbook's sheets
    output_wb.sheets.__getitem__.return_value = output_sheet
    output_sheet.api.ProtectContents = True  # Pretend the sheet is protected

    return input_sheet, output_wb, output_sheet, input_data

def test_convert_unprotects_sheet_and_reads_data(mock_input_output):
    input_sheet, output_wb, output_sheet, input_data = mock_input_output

    converter = DeficiencySummaryConverter(input_sheet=input_sheet, output_wb=output_wb)

    converter.convert()

    # Assert it tried to unprotect the sheet
    output_sheet.api.Unprotect.assert_called_once()

    # Assert it read from the right range
    input_sheet.range.assert_called_with("A19:L60")

    # (Optionally) you could later check how you write data to output_sheet if your convert() writes output
