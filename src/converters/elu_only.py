# src/converters/elu_only.py

from .base import BaseSheetConverter
"""
get_from_input_cell(self, cell):

put_to_output_cell(self, sheet_index_or_name, cell, value):
"""
class EluOnlyConverter(BaseSheetConverter):
    """
    Conversion logic for the "ELU only" sheet.
    For demonstration, copies A1 from input to B8 of the output.
    """
    def convert(self):
        pass
