from .base import BaseSheetConverter

class ExtraBatteryBoxConverter(BaseSheetConverter):

    def convert(self):

        # === Battery Box 1 ===
        location_1 = self.get_from_input_cell("G5")
        if location_1:
            self.put_to_output_cell("22.5 | Power Supply(s)", "F51", location_1)
            self.put_to_output_cell("22.5 | Power Supply(s)", "F52", "Battery Box")

            self.transfer_checkbox_rating("A6", "22.5 | Power Supply(s)", 60, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("A7", "22.5 | Power Supply(s)", 61, col_yes="L", col_no="N", col_na="P")

            self.put_to_output_cell("22.5 | Power Supply(s)", "M62", self.get_from_input_cell("G9"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "M64", self.get_from_input_cell("G10"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "M65", self.get_from_input_cell("G11"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "M66", self.get_from_input_cell("G13"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "M67", self.get_from_input_cell("G14"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "K85", self.get_from_input_cell("F16"))

            self.transfer_checkbox_rating("A17", "22.5 | Power Supply(s)", 68, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("A18", "22.5 | Power Supply(s)", 69, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("A19", "22.5 | Power Supply(s)", 70, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("A20", "22.5 | Power Supply(s)", 71, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("A21", "22.5 | Power Supply(s)", 72, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("A22", "22.5 | Power Supply(s)", 73, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("A23", "22.5 | Power Supply(s)", 74, col_yes="L", col_no="N", col_na="P")

            self.put_to_output_cell("22.5 | Power Supply(s)", "M75", self.get_from_input_cell("G24"))

            self.transfer_checkbox_rating("A25", "22.5 | Power Supply(s)", 76, col_yes="L", col_no="N", col_na="P")

            self.put_to_output_cell("22.5 | Power Supply(s)", "K82", self.get_from_input_cell("G37"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "K83", self.get_from_input_cell("G39"))

            self.transfer_checkbox_rating("A40", "22.5 | Power Supply(s)", 84, col_yes="L", col_no="N", col_na="P")

        # === Battery Box 2 ===
        location_2 = self.get_from_input_cell("P5")
        if location_2:
            self.put_to_output_cell("22.5 | Power Supply(s)", "F96", location_2)
            self.put_to_output_cell("22.5 | Power Supply(s)", "F97", "Battery Box")

            self.transfer_checkbox_rating("J6", "22.5 | Power Supply(s)", 105, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("J7", "22.5 | Power Supply(s)", 106, col_yes="L", col_no="N", col_na="P")

            self.put_to_output_cell("22.5 | Power Supply(s)", "M107", self.get_from_input_cell("P9"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "M109", self.get_from_input_cell("P10"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "M110", self.get_from_input_cell("P11"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "M111", self.get_from_input_cell("P13"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "M112", self.get_from_input_cell("P14"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "K130", self.get_from_input_cell("O16"))

            self.transfer_checkbox_rating("J17", "22.5 | Power Supply(s)", 113, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("J18", "22.5 | Power Supply(s)", 114, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("J19", "22.5 | Power Supply(s)", 115, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("J20", "22.5 | Power Supply(s)", 116, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("J21", "22.5 | Power Supply(s)", 117, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("J22", "22.5 | Power Supply(s)", 118, col_yes="L", col_no="N", col_na="P")
            self.transfer_checkbox_rating("J23", "22.5 | Power Supply(s)", 119, col_yes="L", col_no="N", col_na="P")

            self.put_to_output_cell("22.5 | Power Supply(s)", "M120", self.get_from_input_cell("P24"))

            self.transfer_checkbox_rating("J25", "22.5 | Power Supply(s)", 121, col_yes="L", col_no="N", col_na="P")

            self.put_to_output_cell("22.5 | Power Supply(s)", "K127", self.get_from_input_cell("P37"))
            self.put_to_output_cell("22.5 | Power Supply(s)", "K128", self.get_from_input_cell("P39"))

            self.transfer_checkbox_rating("J40", "22.5 | Power Supply(s)", 129, col_yes="L", col_no="N", col_na="P")

        # === Adjust Print Area Based on Which Boxes Were Found ===
        output_sheet = self.output_wb.sheets["22.5 | Power Supply(s)"]
        if output_sheet.api.ProtectContents:
            output_sheet.api.Unprotect()

        if location_1 and location_2:
            print_area = "A1:P138"
        elif location_1:
            print_area = "A1:P90"
        elif location_2:
            print_area = "A1:P135"
        else:
            print_area = None

        if print_area:
            output_sheet.api.PageSetup.PrintArea = f"${print_area.replace(':', ':$')}"

