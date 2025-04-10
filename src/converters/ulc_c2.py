# src/converters/ulc_c2.py

from .base import BaseSheetConverter
"""
get_from_input_cell(self, cell):

put_to_output_cell(self, sheet_index_or_name, cell, value):
"""
class ULCC2Converter(BaseSheetConverter):
    def convert(self):
        # region 22 | CU or Transp Insp
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
        # endregion


        # region 22.2 | CU or Transp Test
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
        # endregion


        # region 22.3 + 22.4 | Voice & PS
        # VOICE COMMUNICATION TEST
        if self.get_checkbox_value('A119') == "NA":
            self.put_to_output_cell("22.3 + 22.4 | Voice & PS", 'J5', 'True')
        else:
            # Transfer Voice Communication Test Values
            for i, input_row in enumerate(range(119, 125), start=7):
                self.transfer_checkbox_rating(f'A{input_row}', "22.3 + 22.4 | Voice & PS", i, col_yes='L', col_no='N', col_na='P')
            
            self.transfer_checkbox_rating('A126', "22.3 + 22.4 | Voice & PS", 13, col_yes='L', col_no='N', col_na='P')

            self.transfer_checkbox_rating('A128', "22.3 + 22.4 | Voice & PS", 14, col_yes='L', col_no='N', col_na='P')
            self.transfer_checkbox_rating('A129', "22.3 + 22.4 | Voice & PS", 15, col_yes='L', col_no='N', col_na='P')

            self.transfer_checkbox_rating('A131', "22.3 + 22.4 | Voice & PS", 16, col_yes='L', col_no='N', col_na='P')
            self.transfer_checkbox_rating('A132', "22.3 + 22.4 | Voice & PS", 17, col_yes='L', col_no='N', col_na='P')

            self.transfer_checkbox_rating('A134', "22.3 + 22.4 | Voice & PS", 18, col_yes='L', col_no='N', col_na='P')

            self.transfer_checkbox_rating('A136', "22.3 + 22.4 | Voice & PS", 19, col_yes='L', col_no='N', col_na='P')

            self.transfer_checkbox_rating('A138', "22.3 + 22.4 | Voice & PS", 20, col_yes='L', col_no='N', col_na='P')

            for i, input_row in enumerate(range(140, 143), start=21):
                self.transfer_checkbox_rating(f'A{input_row}', "22.3 + 22.4 | Voice & PS", i, col_yes='L', col_no='N', col_na='P')
        

        # POWER SUPPLY INSPECTION
        self.transfer_checkbox_rating('H5', "22.3 + 22.4 | Voice & PS", 33, col_yes='L', col_no='N', col_na='P')

        self.transfer_checkbox_rating('H7', "22.3 + 22.4 | Voice & PS", 35, col_yes='L', col_no='N', col_na='P')
        #endregion


        # region 22.5 | Power Supply(s) 
        for i, input_row in enumerate(range(16, 18), start=15):
            self.transfer_checkbox_rating(f'H{input_row}', "22.5 | Power Supply(s)", i, col_yes='L', col_no='N', col_na='P')
        
        battery_voltage_power_on = self.get_from_input_cell('N19')
        self.put_to_output_cell("22.5 | Power Supply(s)", "M17", battery_voltage_power_on)

        battery_voltage_power_off = self.get_from_input_cell('N20')
        self.put_to_output_cell("22.5 | Power Supply(s)", "M19", battery_voltage_power_off)

        battery_amps_power_off = self.get_from_input_cell('N21')
        self.put_to_output_cell("22.5 | Power Supply(s)", "M20", battery_amps_power_off)

        battery_voltage_power_off_full_load = self.get_from_input_cell('N23')
        self.put_to_output_cell("22.5 | Power Supply(s)", "M21", battery_voltage_power_off_full_load)

        battery_amps_power_off_full_load = self.get_from_input_cell('N24')
        self.put_to_output_cell("22.5 | Power Supply(s)", "M22", battery_amps_power_off_full_load)

        for i, input_row in enumerate(range(27, 34), start=23):
            self.transfer_checkbox_rating(f'H{input_row}', "22.5 | Power Supply(s)", i, col_yes='L', col_no='N', col_na='P')
        
        battery_date_code = self.get_from_input_cell('N34')
        self.put_to_output_cell("22.5 | Power Supply(s)", "M30", battery_date_code)

        self.transfer_checkbox_rating('H35', "22.5 | Power Supply(s)", 31, col_yes='L', col_no='N', col_na='P')

        battery_capacity = self.get_from_input_cell('N47')
        self.put_to_output_cell("22.5 | Power Supply(s)", "K37", battery_capacity)

        battery_terminal_voltage = self.get_from_input_cell('N49')
        self.put_to_output_cell("22.5 | Power Supply(s)", "K38", battery_terminal_voltage)

        self.transfer_checkbox_rating('H50', "22.5 | Power Supply(s)", 38, col_yes='L', col_no='N', col_na='P')

        battery_charging_current = self.get_from_input_cell('M26')
        self.put_to_output_cell("22.5 | Power Supply(s)", "K40", battery_charging_current)

        self.transfer_checkbox_rating('H51', "22.5 | Power Supply(s)", 43, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('H52', "22.5 | Power Supply(s)", 44, col_yes='L', col_no='N', col_na='P')
        # endregion


        # region 22.6 | Annunciator(s)
        self.transfer_checkbox_rating('A67', "22.6 | Annunciator(s)", 9, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A68', "22.6 | Annunciator(s)", 10, col_yes='L', col_no='N', col_na='P')

        self.transfer_checkbox_rating('A70', "22.6 | Annunciator(s)", 11, col_yes='L', col_no='N', col_na='P')

        self.transfer_checkbox_rating('A72', "22.6 | Annunciator(s)", 12, col_yes='L', col_no='N', col_na='P')

        for i, input_row in enumerate(range(74, 79), start=13):
            self.transfer_checkbox_rating(f'A{input_row}', "22.6 | Annunciator(s)", i, col_yes='L', col_no='N', col_na='P')

        for i, input_row in enumerate(range(82, 86), start=18):
            self.transfer_checkbox_rating(f'A{input_row}', "22.6 | Annunciator(s)", i, col_yes='L', col_no='N', col_na='P')
        # endregion


        # region 22.7 | Annun & Seq Disp
        if str(self.get_from_input_cell('E90')).strip().upper() == "NOT APPLICABLE":
            self.put_to_output_cell("22.7 | Annun & Seq Disp", "J6", "True")
        else:
            sequential_display_location = self.get_from_input_cell('E90')
            self.put_to_output_cell("22.7 | Annun & Seq Disp", "G8", sequential_display_location)

            self.transfer_checkbox_rating('A92', "22.7 | Annun & Seq Disp", 10, col_yes='L', col_no='N', col_na='P')

            self.transfer_checkbox_rating('A105', "22.7 | Annun & Seq Disp", 13, col_yes='L', col_no='N', col_na='P')

            for i, input_row in enumerate(range(107, 112), start=14):
                self.transfer_checkbox_rating(f'A{input_row}', "22.7 | Annun & Seq Disp", i, col_yes='L', col_no='N', col_na='P')
            
            for i, input_row in enumerate(range(115, 118), start=19):
                self.transfer_checkbox_rating(f'A{input_row}', "22.7 | Annun & Seq Disp", i, col_yes='L', col_no='N', col_na='P')
        
        if str(self.get_from_input_cell('L119')).strip().upper() == "NOT APPLICABLE":
            self.put_to_output_cell("22.7 | Annun & Seq Disp", "J29", "True")
        else:
            self.put_to_output_cell("22.7 | Annun & Seq Disp", "G31", self.get_from_input_cell('L119'))

            for i, input_row in enumerate(range(121, 125), start=19):
                self.transfer_checkbox_rating(f'H{input_row}', "22.7 | Annun & Seq Disp", i, col_yes='L', col_no='N', col_na='P')
        # endregion


        # region 22.9 + 22.10 | Printer
        # Printer
        if str(self.get_from_input_cell('L127')).strip().upper() == "NOT APPLICABLE":
            self.put_to_output_cell("22.9 + 22.10 | Printer", "J6", "True")
        else:
            self.put_to_output_cell("22.9 + 22.10 | Printer", "D8", self.get_from_input_cell('L127'))
            self.transfer_checkbox_rating("H129", "22.9 + 22.10 | Printer", 10, col_yes='L', col_no='N', col_na='P')
            self.transfer_checkbox_rating("H132", "22.9 + 22.10 | Printer", 11, col_yes='L', col_no='N', col_na='P')
        
        # Ancillary Device Circuit
        for i, input_row in enumerate(range(91, 97), start=19):
            self.put_to_output_cell("22.9 + 22.10 | Printer", f"A{i}", self.get_from_input_cell(f"H{input_row}"))
            if self.is_yes(f"L{input_row}"):
                self.put_to_output_cell("22.9 + 22.10 | Printer", f"J{i}", "True")
            elif self.is_no(f"N{input_row}"):
                self.put_to_output_cell("22.9 + 22.10 | Printer", f"L{i}", "True")
        
        # Monitoring / Interconnection
        if str(self.get_from_input_cell('K83')).strip() == "" or self.get_from_input_cell('K83') == None:
            self.put_to_output_cell("22.9 + 22.10 | Printer", "L34", "True")
        else:
            self.transfer_checkbox_rating("H65", "22.9 + 22.10 | Printer", 36, col_yes='O', col_no='R', col_na='R')
            self.transfer_checkbox_rating("H72", "22.9 + 22.10 | Printer", 37, col_yes='O', col_no='R', col_na='S')
            self.transfer_checkbox_rating("H74", "22.9 + 22.10 | Printer", 39, col_yes='O', col_no='R', col_na='S')
            self.transfer_checkbox_rating("H76", "22.9 + 22.10 | Printer", 38, col_yes='N', col_no='P', col_na='R')
            self.transfer_checkbox_rating("H78", "22.9 + 22.10 | Printer", 44, col_yes='O', col_no='R', col_na='S')
            receiving_centre_name = self.get_from_input_cell('K83')
            self.put_to_output_cell("22.9 + 22.10 | Printer", "O42", receiving_centre_name)
            receiving_centre_number = self.get_from_input_cell('M83')
            self.put_to_output_cell("22.9 + 22.10 | Printer", "O43", receiving_centre_number)
        # endregion
