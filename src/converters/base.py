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

    def insert_formatted_row_below(self, output_sheet, row_to_clone, log_prefix="", password=None):
        """
        Inserts a new row below row_to_clone in the specified output sheet, and copies
        formatting (borders, font bold, merged cells, and row height).

        Parameters:
        - output_sheet: xlwings Sheet object
        - row_to_clone: int — the row number to clone formatting from
        - log_prefix: str — optional tag to prepend to logs
        - password: str or None — optional password for unprotecting/re-protecting the sheet
        """
        try:
            row_to_insert = row_to_clone + 1

            if row_to_clone < 1:
                return

            # Step 0: Unprotect the sheet if needed
            if output_sheet.api.ProtectContents:
                try:
                    if password:
                        output_sheet.api.Unprotect(Password=password)
                    else:
                        output_sheet.api.Unprotect()
                except Exception as e:
                    print(f"{log_prefix}[ERROR] Could not unprotect sheet: {e}")
                    return

            # Step 1: Insert the new row
            try:
                output_sheet.range(f"{row_to_insert}:{row_to_insert}").insert(shift="down")
            except Exception as e:
                print(f"{log_prefix}[ERROR] Failed to insert new row at {row_to_insert}: {e}")
                return

            original_row = output_sheet.range(f"A{row_to_clone}:L{row_to_clone}")
            new_row = output_sheet.range(f"A{row_to_insert}:L{row_to_insert}")

            # Step 2: Copy formatting
            for source_cell, target_cell in zip(original_row, new_row):
                try:
                    target_cell.value = None  # Clear values
                    for side in (1, 2, 3, 4):  # Left, Right, Top, Bottom
                        target_border = target_cell.api.Borders(side)
                        target_border.LineStyle = 1
                        target_border.Weight = 2
                        target_border.Color = 0.0
                        

                        source_border = source_cell.api.Borders(side)
                        source_border.LineStyle = 1
                        source_border.Weight = 2
                        source_border.Color = 0.0

                        if target_cell.column == 1 and side == 1:
                            target_border.Weight = 3
                        
                        if source_cell.column == 1 and side == 1:
                            source_border.Weight = 3
                        
                        if target_cell.column == 12 and side == 2:
                            target_border.Weight = 3
                        
                        if source_cell.column == 12 and side == 2:
                            source_border.Weight = 3


                except Exception as e:
                    print(f"{log_prefix}[ERROR] Formatting cell {target_cell.address}: {e}")

            # Step 3: Reapply merged ranges
            already_merged = set()
            for source_cell in original_row:
                try:
                    merge_address = source_cell.api.MergeArea.Address
                    if merge_address and merge_address not in already_merged:
                        new_merge_address = merge_address.replace(str(row_to_clone), str(row_to_insert))
                        output_sheet.range(new_merge_address).merge()
                        already_merged.add(merge_address)
                except Exception as e:
                    print(f"{log_prefix}[ERROR] Merging cells in new row from {source_cell.address}: {e}")

            # Step 4: Set row height (defaulting to 15 if uncertain)
            try:
                input_height = 15
                output_sheet.range(f"{row_to_insert}:{row_to_insert}").row_height = input_height
            except Exception as e:
                print(f"{log_prefix}[WARNING] Could not copy row height: {e}")
        


        except Exception as final_error:
            print(f"{log_prefix}[FATAL] insert_formatted_row_below failed: {final_error}")



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


    def apply_wrap_text_to_output_cell(self, sheet_index_or_name, cell, wrap=True, password=None):
        """
        Applies or removes text wrapping in a specified output cell.

        Parameters:
        - sheet_index_or_name: int or str — index or name of the output sheet
        - cell: str — Excel cell reference (e.g., "B12")
        - wrap: bool — whether to enable (True) or disable (False) wrap text
        - password: str or None — optional password to unprotect and re-protect the sheet
        """
        try:
            # Validate and access the sheet
            try:
                sheet = self.output_wb.sheets[sheet_index_or_name]
            except Exception as e:
                raise ValueError(f"Invalid sheet reference '{sheet_index_or_name}': {e}")

            # Unprotect the sheet if necessary
            if sheet.api.ProtectContents:
                try:
                    if password:
                        sheet.api.Unprotect(Password=password)
                    else:
                        sheet.api.Unprotect()
                except Exception as e:
                    raise PermissionError(f"Could not unprotect sheet '{sheet.name}': {e}")

            # Apply wrap text setting
            try:
                sheet[cell].api.WrapText = wrap
            except Exception as e:
                raise ValueError(f"Failed to set WrapText on cell '{cell}' in sheet '{sheet.name}': {e}")

        except Exception as final_error:
            raise Exception(f"[apply_wrap_text_to_output_cell ERROR] {final_error}")

    def get_from_output_cell(self, sheet_index_or_name, cell, password=None):
        """
        Safely retrieves a value from a cell in the output workbook, with support for unprotecting sheets if needed.

        Parameters:
        - sheet_index_or_name: int or str — the index (e.g., 0) or name (e.g., 'Summary') of the sheet
        - cell: str — the Excel cell reference (e.g., 'A12')
        - password: str or None — optional password to unprotect the sheet if needed

        Returns:
        - The value in the specified cell, or raises an exception with a descriptive error.
        """
        try:
            # Validate and access the sheet
            try:
                sheet = self.output_wb.sheets[sheet_index_or_name]
            except Exception as e:
                raise ValueError(f"Invalid sheet reference '{sheet_index_or_name}': {e}")

            # Attempt to unprotect the sheet (usually not needed just for reading, but included for parity)
            if sheet.api.ProtectContents:
                try:
                    if password:
                        sheet.api.Unprotect(Password=password)
                    else:
                        sheet.api.Unprotect()
                except Exception as e:
                    raise PermissionError(f"Could not unprotect sheet '{sheet.name}': {e}")

            # Try reading from the specified cell
            try:
                return sheet[cell].value
            except Exception as e:
                raise ValueError(f"Invalid cell reference '{cell}' on sheet '{sheet.name}': {e}")

        except Exception as final_error:
            raise Exception(f"[get_from_output_cell ERROR] {final_error}")




    def copy_row_height_to_output(self, input_row_number, output_row_number, sheet_index_or_name, password=None):
        """
        Copies the row height from a row on the input sheet to the specified row on the output sheet.

        Parameters:
        - input_row_number: int — row number from the input sheet (1-based)
        - output_row_number: int — row number on the output sheet (1-based)
        - sheet_index_or_name: int or str — the index or name of the output sheet
        - password: str or None — optional password to unprotect/re-protect the output sheet
        """
        try:
            # Access output sheet
            try:
                sheet = self.output_wb.sheets[sheet_index_or_name]
            except Exception as e:
                raise ValueError(f"Invalid output sheet reference '{sheet_index_or_name}': {e}")

            # Unprotect sheet if needed
            if sheet.api.ProtectContents:
                try:
                    if password:
                        sheet.api.Unprotect(Password=password)
                    else:
                        sheet.api.Unprotect()
                except Exception as e:
                    raise PermissionError(f"Could not unprotect sheet '{sheet.name}': {e}")

            # Get row height from input sheet
            try:
                input_height = self.input_sheet.range(f"{input_row_number}:{input_row_number}").row_height
            except Exception as e:
                raise ValueError(f"Could not read row height from input row {input_row_number}: {e}")

            # Apply height to output sheet row
            try:
                sheet.range(f"{output_row_number}:{output_row_number}").row_height = input_height
            except Exception as e:
                raise ValueError(f"Could not set row height on output row {output_row_number}: {e}")

        except Exception as final_error:
            raise Exception(f"[copy_row_height_to_output ERROR] {final_error}")


    def put_to_output_cell(self, sheet_index_or_name, cell, value, password=None, suppress_none=False, wrap_text=None):
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

                if wrap_text is not None:
                    sheet[cell].api.WrapText = wrap_text
            except Exception as e:
                raise ValueError(f"Invalid cell reference '{cell}' on sheet '{sheet.name}': {e}")

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
