from .base import BaseSheetConverter

class ExtraAnnunciatorConverter(BaseSheetConverter):

    start_row_of_22_6_section = [
        7,
        38,
        78,
        118,
        154
    ]

    def convert(self):
        print("Converting extra annunciators")
        num_extra_annuns = 0

        # region Extra Annun 1
        start_row = self.start_row_of_22_6_section[num_extra_annuns]

        annun_location = self.get_from_input_cell("E5")
        if annun_location is not None:
            num_extra_annuns += 1
            start_row = self.start_row_of_22_6_section[num_extra_annuns]
            self.put_to_output_cell("22.6 | Annunciator(s)", f"G{start_row}", annun_location)
            
        
        annun_id = self.get_from_input_cell("E6")
        if annun_id is not None:
            self.put_to_output_cell("22.6 | Annunciator(s)", f"G{start_row + 1}", annun_id)
        
        self.transfer_checkbox_rating("A7", "22.6 | Annunciator(s)", start_row + 2, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A8", "22.6 | Annunciator(s)", start_row + 3, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("A18", "22.6 | Annunciator(s)", start_row + 4, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("A20", "22.6 | Annunciator(s)", start_row + 5, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("A22", "22.6 | Annunciator(s)", start_row + 6, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A23", "22.6 | Annunciator(s)", start_row + 7, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A24", "22.6 | Annunciator(s)", start_row + 8, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A25", "22.6 | Annunciator(s)", start_row + 9, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A26", "22.6 | Annunciator(s)", start_row + 10, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("A30", "22.6 | Annunciator(s)", start_row + 11, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A31", "22.6 | Annunciator(s)", start_row + 12, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A32", "22.6 | Annunciator(s)", start_row + 13, col_yes= "L", col_no="N", col_na="P")
        # endregion

        # region Extra Annun 2
        start_row = self.start_row_of_22_6_section[num_extra_annuns]

        annun_location = self.get_from_input_cell("L5")
        if annun_location is not None:
            num_extra_annuns += 1
            start_row = self.start_row_of_22_6_section[num_extra_annuns]
            self.put_to_output_cell("22.6 | Annunciator(s)", f"G{start_row}", annun_location)
            
        
        annun_id = self.get_from_input_cell("L6")
        if annun_id is not None:
            self.put_to_output_cell("22.6 | Annunciator(s)", f"G{start_row + 1}", annun_id)
        
        self.transfer_checkbox_rating("H7", "22.6 | Annunciator(s)", start_row + 2, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H8", "22.6 | Annunciator(s)", start_row + 3, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("H18", "22.6 | Annunciator(s)", start_row + 4, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("H20", "22.6 | Annunciator(s)", start_row + 5, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("H22", "22.6 | Annunciator(s)", start_row + 6, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H23", "22.6 | Annunciator(s)", start_row + 7, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H24", "22.6 | Annunciator(s)", start_row + 8, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H25", "22.6 | Annunciator(s)", start_row + 9, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H26", "22.6 | Annunciator(s)", start_row + 10, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("H30", "22.6 | Annunciator(s)", start_row + 11, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H31", "22.6 | Annunciator(s)", start_row + 12, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H32", "22.6 | Annunciator(s)", start_row + 13, col_yes= "L", col_no="N", col_na="P")
        # endregion

        # region Extra Annun 3
        start_row = self.start_row_of_22_6_section[num_extra_annuns]

        annun_location = self.get_from_input_cell("E36")
        if annun_location is not None:
            num_extra_annuns += 1
            start_row = self.start_row_of_22_6_section[num_extra_annuns]
            self.put_to_output_cell("22.6 | Annunciator(s)", f"G{start_row}", annun_location)
            
        
        annun_id = self.get_from_input_cell("E37")
        if annun_id is not None:
            self.put_to_output_cell("22.6 | Annunciator(s)", f"G{start_row + 1}", annun_id)
        
        self.transfer_checkbox_rating("A38", "22.6 | Annunciator(s)", start_row + 2, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A39", "22.6 | Annunciator(s)", start_row + 3, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("A42", "22.6 | Annunciator(s)", start_row + 4, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("A47", "22.6 | Annunciator(s)", start_row + 5, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("A53", "22.6 | Annunciator(s)", start_row + 6, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A54", "22.6 | Annunciator(s)", start_row + 7, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A55", "22.6 | Annunciator(s)", start_row + 8, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A56", "22.6 | Annunciator(s)", start_row + 9, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A57", "22.6 | Annunciator(s)", start_row + 10, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("A61", "22.6 | Annunciator(s)", start_row + 11, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A62", "22.6 | Annunciator(s)", start_row + 12, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("A63", "22.6 | Annunciator(s)", start_row + 13, col_yes= "L", col_no="N", col_na="P")
        # endregion

        # region Extra Annun 4
        start_row = self.start_row_of_22_6_section[num_extra_annuns]

        annun_location = self.get_from_input_cell("L36")
        if annun_location is not None:
            num_extra_annuns += 1
            start_row = self.start_row_of_22_6_section[num_extra_annuns]
            self.put_to_output_cell("22.6 | Annunciator(s)", f"G{start_row}", annun_location)
            
        
        annun_id = self.get_from_input_cell("L37")
        if annun_id is not None:
            self.put_to_output_cell("22.6 | Annunciator(s)", f"G{start_row + 1}", annun_id)
        
        self.transfer_checkbox_rating("H38", "22.6 | Annunciator(s)", start_row + 2, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H39", "22.6 | Annunciator(s)", start_row + 3, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("H42", "22.6 | Annunciator(s)", start_row + 4, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("H47", "22.6 | Annunciator(s)", start_row + 5, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("H53", "22.6 | Annunciator(s)", start_row + 6, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H54", "22.6 | Annunciator(s)", start_row + 7, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H55", "22.6 | Annunciator(s)", start_row + 8, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H56", "22.6 | Annunciator(s)", start_row + 9, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H57", "22.6 | Annunciator(s)", start_row + 10, col_yes= "L", col_no="N", col_na="P")

        self.transfer_checkbox_rating("H61", "22.6 | Annunciator(s)", start_row + 11, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H62", "22.6 | Annunciator(s)", start_row + 12, col_yes= "L", col_no="N", col_na="P")
        self.transfer_checkbox_rating("H63", "22.6 | Annunciator(s)", start_row + 13, col_yes= "L", col_no="N", col_na="P")
        # endregion
        

