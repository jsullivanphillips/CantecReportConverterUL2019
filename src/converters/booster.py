# converters/booster.py
from .base import BaseSheetConverter

class BoosterConverter(BaseSheetConverter):

    start_row_of_22_1_section = [
        1,
        44,
        90,
        135,
        180,
        228
    ]

    start_row_of_22_2_section = [
        1,
        40,
        79,
        118,
    ]

    start_row_of_22_5_section = [
        1,
        46,
        91,
        136
    ]

    def __init__(self, input_sheet, output_wb, booster_count):
        self.input_sheet = input_sheet
        self.output_wb = output_wb
        self.booster_count = booster_count

    def convert(self):
        print(f"Converting {self.input_sheet.name} with booster count = {self.booster_count}")

        rel_start_row = self.start_row_of_22_1_section[self.booster_count]
        # region 22.1 CU or Transp Insp
        # 22.1 CU or Transp Insp
        # CU location
        location = self.get_from_input_cell("E5")
        if location is not None:
            self.put_to_output_cell("22.1 | CU or Transp Insp", f"H{rel_start_row + 14}", location)

        # CU identifaction
        identif = self.get_from_input_cell("E6")
        if identif is not None:
            self.put_to_output_cell("22.1 | CU or Transp Insp",f"H{rel_start_row + 15}", identif)

        # Input circuit designations 
        self.transfer_checkbox_rating('A49', "22.1 | CU or Transp Insp", rel_start_row + 16 , col_yes='Q', col_no='S', col_na='U')

        # Output circuit designations
        self.transfer_checkbox_rating('A51', "22.1 | CU or Transp Insp", rel_start_row + 17 , col_yes='Q', col_no='S', col_na='U')

        # Correct designations
        self.transfer_checkbox_rating('A53', "22.1 | CU or Transp Insp", rel_start_row + 18 , col_yes='Q', col_no='S', col_na='U')

        # Plug in components
        self.transfer_checkbox_rating('A54', "22.1 | CU or Transp Insp", rel_start_row + 19 , col_yes='Q', col_no='S', col_na='U')

        # Plug in cables
        self.transfer_checkbox_rating('A55', "22.1 | CU or Transp Insp", rel_start_row + 20 , col_yes='Q', col_no='S', col_na='U')

        # Dates of firmware
        date = self.get_from_input_cell("E56")
        if date is not None:
            self.put_to_output_cell("22.1 | CU or Transp Insp", "Q22", date)
        revision = self.get_from_input_cell("F57")
        if revision is not None:
            self.put_to_output_cell("22.1 | CU or Transp Insp", "Q23", revision)
        version = self.get_from_input_cell("F58")
        if version is not None:
            self.put_to_output_cell("22.1 | CU or Transp Insp", "T23", version)
        
        # Dust and Dirt
        self.transfer_checkbox_rating('A59', "22.1 | CU or Transp Insp", rel_start_row + 25 , col_yes='Q', col_no='S', col_na='U')

        # Fused in accordance
        self.transfer_checkbox_rating('A60', "22.1 | CU or Transp Insp", rel_start_row + 26 , col_yes='Q', col_no='S', col_na='U')

        # Lock functional
        self.transfer_checkbox_rating('A61', "22.1 | CU or Transp Insp", rel_start_row + 27 , col_yes='Q', col_no='S', col_na='U')

        # Termination Points
        self.transfer_checkbox_rating('A62', "22.1 | CU or Transp Insp", rel_start_row + 28 , col_yes='Q', col_no='S', col_na='U')

        output_sheet = self.output_wb.sheets["22.1 | CU or Transp Insp"]

        # 🔓 Unprotect once at the beginning
        if output_sheet.api.ProtectContents:
            output_sheet.api.Unprotect()
        
        print_range = f"A1:U{self.start_row_of_22_1_section[self.booster_count + 1]-1}"
        output_sheet.api.PageSetup.PrintArea = f"${print_range.replace(':', ':$')}"
        # endregion

        # region 22.2 CU or Transp Test
        rel_start_row = self.start_row_of_22_2_section[self.booster_count]
        location = self.get_from_input_cell("E5")
        if location is not None:
            self.put_to_output_cell("22.2 | CU or Transp Test", f"F{rel_start_row + 6}", location)

        # CU identifaction
        identif = self.get_from_input_cell("E6")
        if identif is not None:
            self.put_to_output_cell("22.2 | CU or Transp Test",f"F{rel_start_row + 7}", identif)

        self.transfer_checkbox_rating('A7', "22.2 | CU or Transp Test", rel_start_row + 8, col_yes='L', col_no='N', col_na='Q')
        self.transfer_checkbox_rating('A8', "22.2 | CU or Transp Test", rel_start_row + 10, col_yes='L', col_no='N', col_na='Q')
        self.transfer_checkbox_rating('A9', "22.2 | CU or Transp Test", rel_start_row + 11, col_yes='L', col_no='N', col_na='Q')
        self.transfer_checkbox_rating('A10', "22.2 | CU or Transp Test", rel_start_row + 12, col_yes='L', col_no='N', col_na='Q')
        self.transfer_checkbox_rating('A11', "22.2 | CU or Transp Test", rel_start_row + 13, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A13', "22.2 | CU or Transp Test", rel_start_row + 15, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A14', "22.2 | CU or Transp Test", rel_start_row + 16, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A15', "22.2 | CU or Transp Test", rel_start_row + 17, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A16', "22.2 | CU or Transp Test", rel_start_row + 19, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A17', "22.2 | CU or Transp Test", rel_start_row + 20, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A19', "22.2 | CU or Transp Test", rel_start_row + 21, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A20', "22.2 | CU or Transp Test", rel_start_row + 22, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A21', "22.2 | CU or Transp Test", rel_start_row + 23, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A22', "22.2 | CU or Transp Test", rel_start_row + 24, col_yes='L', col_no='N', col_na='P')

        # Alarm signal silence cut-out timer
        time = self.get_from_input_cell("F24")
        if time is not None:
            self.put_to_output_cell("22.2 | CU or Transp Test", "L26", time)
        else:
            self.transfer_checkbox_rating('A24', "22.2 | CU or Transp Test", rel_start_row + 25, col_yes='L', col_no='N', col_na='P')
        
        self.transfer_checkbox_rating('A25', "22.2 | CU or Transp Test", rel_start_row + 26, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A29', "22.2 | CU or Transp Test", rel_start_row + 27, col_yes='L', col_no='N', col_na='Q')
        self.transfer_checkbox_rating('A31', "22.2 | CU or Transp Test", rel_start_row + 28, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A32', "22.2 | CU or Transp Test", rel_start_row + 29, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A33', "22.2 | CU or Transp Test", rel_start_row + 30, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A34', "22.2 | CU or Transp Test", rel_start_row + 31, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A35', "22.2 | CU or Transp Test", rel_start_row + 32, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A37', "22.2 | CU or Transp Test", rel_start_row + 33, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A38', "22.2 | CU or Transp Test", rel_start_row + 34, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A39', "22.2 | CU or Transp Test", rel_start_row + 35, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A43', "22.2 | CU or Transp Test", rel_start_row + 36, col_yes='L', col_no='N', col_na='Q')
        self.transfer_checkbox_rating('A44', "22.2 | CU or Transp Test", rel_start_row + 37, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('A45', "22.2 | CU or Transp Test", rel_start_row + 38, col_yes='L', col_no='N', col_na='P')

        output_sheet = self.output_wb.sheets["22.2 | CU or Transp Test"]

        # 🔓 Unprotect once at the beginning
        if output_sheet.api.ProtectContents:
            output_sheet.api.Unprotect()
        
        print_range = f"A1:P{self.start_row_of_22_2_section[self.booster_count + 1]-1}"
        output_sheet.api.PageSetup.PrintArea = f"${print_range.replace(':', ':$')}"
        # endregion

        # region 22.5 Power Supply
        rel_start_row = self.start_row_of_22_5_section[self.booster_count]
        location = self.get_from_input_cell("E5")
        if location is not None:
            self.put_to_output_cell("22.5 | Power Supply(s)", f"F{rel_start_row + 5}", location)
        
        id = self.get_from_input_cell("E6")
        if id is not None:
            self.put_to_output_cell("22.5 | Power Supply(s)", f"F{rel_start_row + 6}", id)
        
        self.transfer_checkbox_rating('H16', "22.5 | Power Supply(s)", rel_start_row + 14, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('H17', "22.5 | Power Supply(s)", rel_start_row + 15, col_yes='L', col_no='N', col_na='P')
        
        v = self.get_from_input_cell("N19")
        if v is not None:
            self.put_to_output_cell("22.5 | Power Supply(s)", f"M{rel_start_row + 16}", v)
        
        v = self.get_from_input_cell("N20")
        if v is not None:
            self.put_to_output_cell("22.5 | Power Supply(s)", f"M{rel_start_row + 18}", v)
        
        a = self.get_from_input_cell("N21")
        if a is not None:
            self.put_to_output_cell("22.5 | Power Supply(s)", f"M{rel_start_row + 19}", a)
        
        v = self.get_from_input_cell("N23")
        if v is not None:
            self.put_to_output_cell("22.5 | Power Supply(s)", f"M{rel_start_row + 20}", v)
        
        a = self.get_from_input_cell("N24")
        if a is not None:
            self.put_to_output_cell("22.5 | Power Supply(s)", f"M{rel_start_row + 21}", a)
        
        self.transfer_checkbox_rating('H27', "22.5 | Power Supply(s)", rel_start_row + 22, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('H28', "22.5 | Power Supply(s)", rel_start_row + 23, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('H29', "22.5 | Power Supply(s)", rel_start_row + 24, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('H30', "22.5 | Power Supply(s)", rel_start_row + 25, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('H31', "22.5 | Power Supply(s)", rel_start_row + 26, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('H32', "22.5 | Power Supply(s)", rel_start_row + 27, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('H33', "22.5 | Power Supply(s)", rel_start_row + 28, col_yes='L', col_no='N', col_na='P')

        year = self.get_from_input_cell("N34")
        if year is not None:
            self.put_to_output_cell("22.5 | Power Supply(s)", f"M{rel_start_row + 29}", year)
        
        self.transfer_checkbox_rating('H35', "22.5 | Power Supply(s)", rel_start_row + 30, col_yes='L', col_no='N', col_na='P')
        
        capacity = self.get_from_input_cell('N47')
        if capacity is not None:
            self.put_to_output_cell("22.5 | Power Supply(s)", f"K{rel_start_row + 36}", capacity)
        
        v = self.get_from_input_cell("N49")
        if v is not None:
            self.put_to_output_cell("22.5 | Power Supply(s)", f"K{rel_start_row + 37}", v)
        
        self.transfer_checkbox_rating('H50', "22.5 | Power Supply(s)", rel_start_row + 38, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('H51', "22.5 | Power Supply(s)", rel_start_row + 42, col_yes='L', col_no='N', col_na='P')
        self.transfer_checkbox_rating('H52', "22.5 | Power Supply(s)", rel_start_row + 43, col_yes='L', col_no='N', col_na='P')

        a = self.get_from_input_cell("M26")
        if a is not None:
            self.put_to_output_cell("22.5 | Power Supply(s)", f"K{rel_start_row + 39}", a)
        
        output_sheet = self.output_wb.sheets["22.5 | Power Supply(s)"]

        # 🔓 Unprotect once at the beginning
        if output_sheet.api.ProtectContents:
            output_sheet.api.Unprotect()
        
        print_range = f"A1:P{self.start_row_of_22_5_section[self.booster_count + 1]-1}"
        output_sheet.api.PageSetup.PrintArea = f"${print_range.replace(':', ':$')}"
        # endregion


