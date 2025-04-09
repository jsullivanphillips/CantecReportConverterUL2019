# src/converters/ext_only.py

from .base import BaseSheetConverter
"""
get_from_input_cell(self, cell):

put_to_output_cell(self, sheet_index_or_name, cell, value):
"""
class ExtOnlyConverter(BaseSheetConverter):
    """
    Conversion logic for the "EXT only" sheet.
    For demonstration, copies A1 from input to B7 of the output.
    """
    def convert(self):
        pass