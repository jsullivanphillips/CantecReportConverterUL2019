# src/converters/elu_only.py

from .base import BaseSheetConverter
"""
get_from_input_cell(self, cell):

put_to_output_cell(self, sheet_index_or_name, cell, value):
"""
class EluOnlyConverter(BaseSheetConverter):
    def convert(self):
        output_sheet = self.output_wb.sheets["ELU only"]

        # 🔓 Unprotect once at the beginning
        if output_sheet.api.ProtectContents:
            output_sheet.api.Unprotect()

        input_data = self.input_sheet.range("A10:H2000").value
        start_output_row = 16

        col_A, col_B, col_C, col_D, col_E, col_F, col_G, bold_mask = [], [], [], [], [], [], [], []
        number_of_consecutive_empty_rows = 0
        last_written_row = start_output_row - 1

        for rel_row, row_data in enumerate(input_data):
            input_row = 10 + rel_row
            output_row = start_output_row + rel_row

            device_location = row_data[1]  # column B

            # Track empty rows
            if not device_location or str(device_location).strip() == "":
                number_of_consecutive_empty_rows += 1
            else:
                number_of_consecutive_empty_rows = 0
                last_written_row = output_row

            if number_of_consecutive_empty_rows >= 100:
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

            is_bold = self.input_sheet.range(f"B{input_row}").font.bold
            bold_mask.append(is_bold is True)

        end_row = start_output_row + len(col_A) - 1

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

        # === Approximate page break handling ===
        total_used_rows = last_written_row - start_output_row + 1
        rows_per_page = 49
        total_pages = (total_used_rows + rows_per_page - 1) // rows_per_page
        last_page_row = start_output_row + (total_pages * rows_per_page) - 1

        # === Set print area ===
        print_range = f"A1:H{last_page_row}"
        output_sheet.api.PageSetup.PrintArea = f"${print_range.replace(':', ':$')}"

        output_sheet.api.Protect(DrawingObjects=True, Contents=True, Scenarios=True)

