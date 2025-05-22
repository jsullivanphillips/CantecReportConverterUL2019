# src/converters/field_device_testing.py
from .base import BaseSheetConverter
"""
get_from_input_cell(self, cell):

put_to_output_cell(self, sheet_index_or_name, cell, value):
"""
class FieldDeviceTestingConverter(BaseSheetConverter):
    def convert(self):
        input_range = range(5, 50)
        output_range = range(5, 50)
        output_sheet_name = "23.1 Field Device Legend"

        input_data_by_key = {}
        s_device_rows = []

        # Step 1: Detect merged "S" device region using xlwings
        for row in input_range:
            cell = self.input_sheet.range(f"A{row}")
            val = cell.value
            if val and str(val).strip().upper() == "S":
                merged_area = cell.merge_area
                if merged_area:
                    s_device_rows = list(range(merged_area.row, merged_area.row + merged_area.rows.count))
                else:
                    s_device_rows = [row]
                break

        # Step 2: Collect data for all non-"S" device rows
        for row in input_range:
            if row in s_device_rows:
                continue

            key = self.get_from_input_cell(f"A{row}")
            type_val = self.get_from_input_cell(f"I{row}")
            model_val = self.get_from_input_cell(f"K{row}")
            optional_val = self.get_from_input_cell(f"O{row}")

            if key and (model_val or optional_val):
                formatted_type = self._format_device_value(type_val)
                formatted_model = self._format_device_value(model_val)
                input_data_by_key[str(key).strip().upper()] = {
                    "type": formatted_type,
                    "model": formatted_model
                }

        # Step 3: Match keys and insert formatted type/model into output
        for row in output_range:
            out_key_val = self.get_from_output_cell(output_sheet_name, f"A{row}")
            if out_key_val:
                out_key_str = str(out_key_val).strip().upper()
                if out_key_str in input_data_by_key:
                    data = input_data_by_key[out_key_str]
                    if data["type"]:
                        self.put_to_output_cell(output_sheet_name, f"L{row}", data["type"])
                    if data["model"]:
                        self.put_to_output_cell(output_sheet_name, f"N{row}", data["model"])

        # Step 4: Handle "S" device rows (no comma formatting)
        s_types = []
        s_models = []

        for row in s_device_rows:
            type_val = self.get_from_input_cell(f"I{row}")
            model_val = self.get_from_input_cell(f"K{row}")
            if type_val:
                s_types.append(str(type_val).strip().upper())
            if model_val:
                s_models.append(str(model_val).strip().upper())

        if s_types:
            self.put_to_output_cell(output_sheet_name, "F13", " / ".join(s_types))
        if s_models:
            self.put_to_output_cell(output_sheet_name, "H15", " / ".join(s_models))

    def _format_device_value(self, value):
        """
        If the value contains commas, normalize to ' / ' separators and strip whitespace.
        Otherwise, return the stripped uppercase value.
        """
        if value is None:
            return ""
        value = str(value).strip().upper()
        if "," in value:
            parts = [part.strip().upper() for part in value.split(",")]
            return " / ".join(parts)
        return value


