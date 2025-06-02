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
        output_sheet = self.output_wb.sheets["EXT only"]

        # 🔓 Unprotect once at the beginning
        if output_sheet.api.ProtectContents:
            output_sheet.api.Unprotect()

        input_data = self.input_sheet.range("A10:H2000").value
        start_output_row = 15
        max_input_rows = len(input_data)

        col_A, col_B, col_C, col_D, col_E, col_F, col_G, bold_mask = [], [], [], [], [], [], [], []
        number_of_consecutive_empty_rows = 0
        last_written_row = start_output_row - 1

        cabinet_line = None  # Store Glass Cabinet Dimensions if found

        for rel_row, row_data in enumerate(input_data):
            input_row = 10 + rel_row
            output_row = start_output_row + rel_row

            device_location = row_data[1]  # column B

            # Look for "Glass Cabinet Dimensions" in column B
            if isinstance(device_location, str) and "glass cabinet dimensions" in device_location.lower():
                cabinet_line = device_location.strip()
                continue  # skip writing this row into the report

            # Track empty rows
            if not device_location or str(device_location).strip() == "":
                number_of_consecutive_empty_rows += 1
            else:
                number_of_consecutive_empty_rows = 0
                last_written_row = output_row

            if number_of_consecutive_empty_rows >= 50:
                break

            # Always record row (even if blank) to preserve spacing
            if row_data[0] == 3:
                row_data[0] = "✔"
            col_A.append(row_data[0])
            col_B.append(device_location)
            col_C.append(row_data[2])
            col_D.append(row_data[3])
            col_E.append(row_data[4])
            col_F.append(row_data[5])
            col_G.append(row_data[6])

            # Track bold (True/False/None)
            is_bold = self.input_sheet.range(f"B{input_row}").font.bold
            bold_mask.append(is_bold is True)

        end_row = start_output_row + len(col_A) - 1

        # === Write data to output ===
        output_sheet.range(f"A{start_output_row}:A{end_row}").value = [[v] for v in col_A]
        output_sheet.range(f"B{start_output_row}:B{end_row}").value = [[v] for v in col_B]
        output_sheet.range(f"C{start_output_row}:C{end_row}").value = [[v] for v in col_C]
        output_sheet.range(f"D{start_output_row}:D{end_row}").value = [[v] for v in col_D]
        output_sheet.range(f"E{start_output_row}:E{end_row}").value = [[v] for v in col_E]
        output_sheet.range(f"F{start_output_row}:F{end_row}").value = [[v] for v in col_F]
        output_sheet.range(f"G{start_output_row}:G{end_row}").value = [[v] for v in col_G]

        output_sheet.range(f"A{start_output_row}:A{end_row}").font.name = "Calibri"

        for i, is_bold in enumerate(bold_mask):
            if is_bold:
                row = start_output_row + i
                output_sheet.range(f"B{row}").font.bold = True

        # === Page break calculation ===
        total_used_rows = last_written_row - start_output_row + 1
        rows_per_page = 49
        total_pages = (total_used_rows + rows_per_page - 1) // rows_per_page
        last_page_row = start_output_row + (total_pages * rows_per_page) - 1

        # === Place cabinet line if found ===
        if cabinet_line:
            target_cell = output_sheet.range(f"B{last_page_row}")
            if not target_cell.value or str(target_cell.value).strip() == "":
                target_cell.value = cabinet_line
                target_cell.font.bold = True

        # === Set print area ===
        print_range = f"A1:H{last_page_row}"
        output_sheet.api.PageSetup.PrintArea = f"${print_range.replace(':', ':$')}"