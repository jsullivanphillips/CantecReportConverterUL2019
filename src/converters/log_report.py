# src/converters/log_report.py

from .base import BaseSheetConverter

"""
get_from_input_cell(self, cell):

put_to_output_cell(self, sheet_index_or_name, cell, value):
"""

class LogReportConverter(BaseSheetConverter):
    """
    Conversion logic for the "LOG REPORT C3.2- Device Record" sheet.
    For demonstration, copies A1 from input to B6 of the output.
    """
    def convert(self):
        # To 20.1 | Report
        building_name = self.get_from_input_cell('C4')
        self.put_to_output_cell("20.1 | Report", 'F9', building_name)

        address = self.get_from_input_cell('C5')
        self.put_to_output_cell("20.1 | Report", 'F10', address)

        # To 22.1 | CU or Transp Insp
        fap_location = self.get_from_input_cell('H5')
        self.put_to_output_cell("22.1 | CU or Transp Insp", 'H15', fap_location)

        fap_make = self.get_from_input_cell('H3')
        fap_model = self.get_from_input_cell('H4')
        fap_identification = fap_make + " " + fap_model
        self.put_to_output_cell("22.1 | CU or Transp Insp", 'H16', fap_identification)




