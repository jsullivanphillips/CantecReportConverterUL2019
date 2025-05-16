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
        types = ""
        models = ""
        i = 0
        for row in range(8, 12):
            cell_I = f"I{row}"
            cell_K = f"K{row}"
            type_val = self.get_from_input_cell(cell_I)
            model_val = self.get_from_input_cell(cell_K)

            print(f"Row {row} - I: {type_val}, K: {model_val}")

            if i != 0:
                types += " / "
                models += " / "

            if type_val is not None:
                types += type_val
            if model_val is not None:
                models += model_val

            i += 1

        # Remove the trailing " / " if at least one row was added
        if i != 0:
            types = types.rstrip(" /")
            models = models.rstrip(" /")

        print(f"Final types: '{types}'")
        print(f"Final models: '{models}'")

        if types:
            print("Writing to F13:", types.upper())
            self.put_to_output_cell("23.1 Field Device Legend", "F13", str(types).upper())
        if models:
            print("Writing to H15:", models.upper())
            self.put_to_output_cell("23.1 Field Device Legend", "H15", str(models).upper())

    
        # RI, DS, OTHER TYPE, SFD, FS, SS,  OTHER SUPERVISORY, 
        # ISO -> EM FAULT ISOLATOR, B, BZ, H, V, SP, HSP,
        for i, input_row in enumerate(range(12, 26), start=20):
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

