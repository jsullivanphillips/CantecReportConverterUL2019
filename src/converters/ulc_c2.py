# src/converters/ulc_c2.py

from .base import BaseSheetConverter
"""
get_from_input_cell(self, cell):

put_to_output_cell(self, sheet_index_or_name, cell, value):
"""
class ULCC2Converter(BaseSheetConverter):
    def convert(self):
        ##-- To 22 | CU or Transp Insp --##
        # Input in relation to connected field devices
        self.transfer_checkbox_rating('A49', "22.1 | CU or Transp Insp", 17)

        # Output circuit designations correctly identified in relation to connected field devicse
        self.transfer_checkbox_rating('A51', "22.1 | CU or Transp Insp", 18)
        
        # Correct designations for common control functions & indicators.
        self.transfer_checkbox_rating('A53', "22.1 | CU or Transp Insp", 19)
        
        # Plug-in components and modules securely in place.
        self.transfer_checkbox_rating('A54', "22.1 | CU or Transp Insp", 20)
        
        # Plug-in cables securely in place.
        self.transfer_checkbox_rating('A55', "22.1 | CU or Transp Insp", 21)
        
        # Record Firmware and Software
        self.put_to_output_cell("22.1 | CU or Transp Insp", "Q22", self.get_from_input_cell('E56')) # Date
        self.put_to_output_cell("22.1 | CU or Transp Insp", "Q23", self.get_from_input_cell('F57')) # Revision
        self.put_to_output_cell("22.1 | CU or Transp Insp", "T23", self.get_from_input_cell('F58')) # Version

        # Clean and free of dust and dirt.
        self.transfer_checkbox_rating('A59', "22.1 | CU or Transp Insp", 26)
        
        # Fuses in accordance with manufacturer's specification. 
        self.transfer_checkbox_rating('A60', "22.1 | CU or Transp Insp", 27)
        
        # Control unit or transponder lock functional.
        self.transfer_checkbox_rating('A61', "22.1 | CU or Transp Insp", 28)
        
        # Termination points from wiring to field devices secure.
        self.transfer_checkbox_rating('A62', "22.1 | CU or Transp Insp", 29)


        ##-- To 22.2 | CU or Transp Test --##
        # Power 'ON' visual indicator operates.
        self.transfer_checkbox_rating('A7', "22.2 | CU or Transp Test", 9, col_yes='L', col_no='N', col_na='P')

        # (8 - 15) : (11 - 18)
        for i, input_row in enumerate(range(8, 16), start=11):
            self.transfer_checkbox_rating(f'A{input_row}', "22.2 | CU or Transp Test", i, col_yes='L', col_no='N', col_na='P')
        
        # Manual transfer from alert signal to alarm signal operates.
        self.transfer_checkbox_rating('A16', "22.2 | CU or Transp Test", 20, col_yes='L', col_no='N', col_na='P')

        # Automatic transfer from alert signal to alarm signal cancel
        self.transfer_checkbox_rating('A17', "22.2 | CU or Transp Test", 21, col_yes='L', col_no='N', col_na='P')

        # (19 - 22) : (22 - 25)
        for i, input_row in enumerate(range(19, 23), start=22):
            self.transfer_checkbox_rating(f'A{input_row}', "22.2 | CU or Transp Test", i, col_yes='L', col_no='N', col_na='P')
        
        # Alarm signal silence automatic cut-out timer. 
        self.transfer_checkbox_rating('A24', "22.2 | CU or Transp Test", 26, col_yes='L', col_no='N', col_na='P')
        self.put_to_output_cell("22.2 | CU or Transp Test", "L26", self.get_from_input_cell('F24')) # Time

        # Input circuit, alarm & supervisory operation, including audible
        self.transfer_checkbox_rating('A29', "22.2 | CU or Transp Test", 28, col_yes='L', col_no='N', col_na='P')

        # (31 - 39) : (29 - 36)
        for i, input_row in enumerate(range(31, 40), start=29):
            self.transfer_checkbox_rating(f'A{input_row}', "22.2 | CU or Transp Test", i, col_yes='L', col_no='N', col_na='P')
        
        # (43 - 45) : (37 - 39)
        for i, input_row in enumerate(range(43, 46), start=37):
            self.transfer_checkbox_rating(f'A{input_row}', "22.2 | CU or Transp Test", i, col_yes='L', col_no='N', col_na='P')
        