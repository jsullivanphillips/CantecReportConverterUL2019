# src/converters/base.py

class BaseSheetConverter:
    """
    Base class for sheet conversion.
    Subclasses receive the input sheet and output workbook and should override convert().
    Provides helper methods to read from input and write to output.
    """

    def __init__(self, input_sheet, output_wb):
        self.input_sheet = input_sheet    # Sheet from the input workbook
        self.output_wb = output_wb        # Entire output (template) workbook

    def convert(self):
        raise NotImplementedError("Subclasses must implement the convert() method.")

    def get_from_input_cell(self, cell):
        """Read a value from a cell in the input sheet."""
        return self.input_sheet[cell].value

    def is_yes(self, cell_ref):
        val = self.get_from_input_cell(cell_ref)
        return str(val).strip().upper() == "YES"

    def is_no(self, cell_ref):
        val = self.get_from_input_cell(cell_ref)
        return str(val).strip().upper() == "NO"
    
    def get_checkbox_value(self, cell_ref):
        """
        Returns the normalized checkbox value from the input cell.

        Possible return values:
        - "CHECK" → if value is 3 or equivalent
        - "X"     → if value is 5 or "X"
        - "NA"    → if value is "M"
        - None    → if unrecognized or empty
        """
        val = self.get_from_input_cell(cell_ref)
        if val is None:
            return None

        val_str = str(val).strip().upper()

        try:
            numeric = float(val)
            if numeric == 3:
                return "CHECK"
            elif numeric == 5:
                return "X"
        except (ValueError, TypeError):
            pass

        if val_str == "X":
            return "X"
        if val_str == "M":
            return "NA"

        return None


    def transfer_checkbox_rating(self, input_cell: str, output_sheet: str, row: int,
                             col_yes: str = 'Q', col_no: str = 'S', col_na: str = 'U'):
        """
        Writes 'True' to the appropriate column based on the checkbox value in the input cell.
        Accepts custom column letters for each type.
        """
        value = self.get_checkbox_value(input_cell)
        if value == "CHECK":
            self.put_to_output_cell(output_sheet, f'{col_yes}{row}', 'True')
        elif value == "X":
            self.put_to_output_cell(output_sheet, f'{col_no}{row}', 'True')
        elif value == "NA":
            self.put_to_output_cell(output_sheet, f'{col_na}{row}', 'True')


    def put_to_output_cell(self, sheet_index_or_name, cell, value, password=None, suppress_none=False):
        """
        Safely writes a value to a cell in the output workbook, with support for unprotecting sheets.
        
        Parameters:
        - sheet_index_or_name: int or str — the index (e.g., 0) or name (e.g., 'Summary') of the sheet
        - cell: str — the Excel cell reference (e.g., 'H11')
        - value: any — the value to write (will be skipped if None and suppress_none=True)
        - password: str or None — optional password to unprotect/re-protect the sheet
        - suppress_none: bool — if True, do not write anything if value is None
        """
        try:
            # Validate and access the sheet
            try:
                sheet = self.output_wb.sheets[sheet_index_or_name]
            except Exception as e:
                raise ValueError(f"Invalid sheet reference '{sheet_index_or_name}': {e}")

            # Skip if value is None and suppression is on
            if value is None and suppress_none:
                print(f"[INFO] Skipping write to {cell} on '{sheet.name}': value is None and suppress_none is True")
                return

            # Attempt to unprotect the sheet
            if sheet.api.ProtectContents:
                try:
                    if password:
                        sheet.api.Unprotect(Password=password)
                    else:
                        sheet.api.Unprotect()
                except Exception as e:
                    raise PermissionError(f"Could not unprotect sheet '{sheet.name}': {e}")

            # Try writing to the specified cell
            try:
                sheet[cell].value = value
            except Exception as e:
                raise ValueError(f"Invalid cell reference '{cell}' on sheet '{sheet.name}': {e}")

            # Re-protect the sheet
            if password:
                sheet.api.Protect(Password=password)
            else:
                sheet.api.Protect()

        except Exception as final_error:
            # You could log or re-raise here depending on how you want to handle it
            raise Exception(f"[put_to_output_cell ERROR] {final_error}")



class DefaultConverter(BaseSheetConverter):
    """
    Default converter for unhandled sheets.
    Writes the sheet name and cell A1 value to B4 of the first output sheet.
    """

    def convert(self):
        value = self.get_from_input_cell('A1')
        message = f"{self.input_sheet.name}: {value}"
        self.put_to_output_cell(0, 'B4', message)
