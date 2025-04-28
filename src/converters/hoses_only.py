# src/converters/hoses_only.py

from .base import BaseSheetConverter

class HosesOnlyConverter(BaseSheetConverter):
    def convert(self):
        self.put_to_output_cell('Deficiency Summary', 'A10', self.get_from_input_cell('A19'))
