# src/converters/appendix.py

from .base import BaseSheetConverter

class AppendixConverter(BaseSheetConverter):
    def convert(self):
        ## --  To 20.1 | Report -- ##
        # Copy to company
        company = self.get_from_input_cell('B19')
        self.put_to_output_cell('20.1 | Report', 'F12', company)

        # System is single-stage
        if self.is_yes('K11'):
            self.put_to_output_cell('20.1 | Report', 'D15', "True")
        elif self.is_no('L11'):
            self.put_to_output_cell('20.1 | Report', 'D15', "False")
        
        # System is two-stage
        if self.is_yes('K12'):
            self.put_to_output_cell('20.1 | Report', 'G15', "True")
        elif self.is_no('L12'):
            self.put_to_output_cell('20.1 | Report', 'G15', "False")
        
        # Tested according to ULC
        if self.is_yes('K13'):
            self.put_to_output_cell('20.1 | Report', 'N22', "True")
        elif self.is_no('L13'):
            self.put_to_output_cell('20.1 | Report', 'R22', "True")

        # Fire alarm system fully functional
        if self.is_yes('K16'):
            self.put_to_output_cell('20.1 | Report', 'N23', "True")
        elif self.is_no('L16'):
            self.put_to_output_cell('20.1 | Report', 'R23', "True")

        # Fire alarm deficiencies
        if self.is_yes('K15'):
            self.put_to_output_cell('20.1 | Report', 'N24', "True")
        elif self.is_no('L15'):
            self.put_to_output_cell('20.1 | Report', 'R24', "True")
        

        ## --  To 20.3 | Recommendations -- ##
        recommendation_1 = self.get_from_input_cell('B30')
        self.put_to_output_cell("20.3 | Recommendations", 'A6', recommendation_1)


