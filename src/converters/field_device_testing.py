# src/converters/field_device_testing.py

from .base import BaseSheetConverter
"""
get_from_input_cell(self, cell):

put_to_output_cell(self, sheet_index_or_name, cell, value):
"""
class FieldDeviceTestingConverter(BaseSheetConverter):
    def convert(self):
        # region 23.1 Field Device Legend
        # M, RHT -> M, RHT
        for i, input_row in enumerate(range(5, 7), start=8):
            type = self.get_from_input_cell(f"I{input_row}")
            model = self.get_from_input_cell(f"K{input_row}")
            if type is not None:
                self.put_to_output_cell("23.1 Field Device Legend", f"L{i}", str(type).upper())
            if model is not None:
                self.put_to_output_cell("23.1 Field Device Legend", f"N{i}", str(model).upper())
        
        # HHT -> HT
        type = self.get_from_input_cell(f"I8")
        model = self.get_from_input_cell(f"K8")
        if type is not None:
            self.put_to_output_cell("23.1 Field Device Legend", f"L10", str(type).upper())
        if model is not None:
            self.put_to_output_cell("23.1 Field Device Legend", f"N10", str(model).upper())

        # S -> S
        type = self.get_from_input_cell(f"I9")
        model = self.get_from_input_cell(f"K9")
        if type is not None:
            self.put_to_output_cell("23.1 Field Device Legend", f"L11", str(type).upper())
        if model is not None:
            self.put_to_output_cell("23.1 Field Device Legend", f"N11", str(model).upper())
    
        # RI, DS, OTHER TYPE, SFD, FS, SS,  OTHER SUPERVISORY, 
        # ISO -> EM FAULT ISOLATOR, B, BZ, H, V, SP, HSP,
        for i, input_row in enumerate(range(13, 26), start=20):
            type = self.get_from_input_cell(f"I{input_row}")
            model = self.get_from_input_cell(f"K{input_row}")
            if type is not None:
                self.put_to_output_cell("23.1 Field Device Legend", f"L{i}", str(type).upper())
            if model is not None:
                self.put_to_output_cell("23.1 Field Device Legend", f"N{i}", str(model).upper())

        # AD, ET, EOL
        for i, input_row in enumerate(range(27, 30), start=36):
            type = self.get_from_input_cell(f"I{input_row}")
            model = self.get_from_input_cell(f"K{input_row}")
            if type is not None:
                self.put_to_output_cell("23.1 Field Device Legend", f"L{i}", str(type).upper())
            if model is not None:
                self.put_to_output_cell("23.1 Field Device Legend", f"N{i}", str(model).upper())
        # endregion

