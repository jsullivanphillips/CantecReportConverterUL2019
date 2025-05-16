# src/converters/hoses_only.py

from .base import BaseSheetConverter

class HosesOnlyConverter(BaseSheetConverter):
    def convert(self):
        output_sheet = self.output_wb.sheets["HOSES only"]

        # 🔓 Unprotect once at the beginning
        if output_sheet.api.ProtectContents:
            output_sheet.api.Unprotect()
        
        # Step 1: Read input data
        start_input_row = 9
        input_data = self.input_sheet.range(f"A{start_input_row}:H2000").value
        start_output_row = 15
        

        col_A, col_B, col_C, col_D, col_E, col_F, col_G, col_H, bold_mask = [], [], [], [], [], [], [], [], []
        
        number_of_consecutive_empty_rows = 0
        last_written_row = start_output_row - 1

        non_empty_row_exists = False

        for rel_row, row_data in enumerate(input_data):
            input_row = start_input_row + rel_row
            output_row = start_output_row + rel_row
            last_written_row = output_row

            device_location = row_data[1]  # column B

            # Track empty rows
            if not device_location or str(device_location).strip() == "":
                number_of_consecutive_empty_rows += 1
            else:
                number_of_consecutive_empty_rows = 0
                non_empty_row_exists = True
                print(f"non empty row exists. Location: {device_location}")

            if number_of_consecutive_empty_rows >= 50:
                break
            
            # Always record row (even if blank) to preserve spacing
            if row_data[0] == 3:
                row_data[0] = "✔"
            col_A.append(row_data[0])              # Output column A (from column A)
            col_B.append(device_location)          # Output column B (from column B)
            col_C.append(row_data[2])              # Output column C (from column C)
            col_D.append(row_data[3])              # Output column D (from column D)
            col_E.append(row_data[4])              # Output column E (from column E)
            col_F.append(row_data[5])              # Output column F (from column F)
            col_G.append(row_data[6])              # Output column G (from column G)
            col_H.append(row_data[7])              # Output column H (from column H)

            # Track bold (True/False/None)
            is_bold = self.input_sheet.range(f"B{input_row}").font.bold
            bold_mask.append(is_bold is True)
        


        if not non_empty_row_exists:
            # this don't work. Not really important however.
            output_sheet.sheet_state = 'hidden'
        else:
            # Step 2: Determine end row
            end_row = start_output_row + len(col_A) - 1

            # Step 3: Bulk write to output sheet
            output_sheet.range(f"A{start_output_row}:A{end_row}").value = [[v] for v in col_A]
            output_sheet.range(f"B{start_output_row}:B{end_row}").value = [[v] for v in col_B]
            output_sheet.range(f"C{start_output_row}:C{end_row}").value = [[v] for v in col_C]
            output_sheet.range(f"D{start_output_row}:D{end_row}").value = [[v] for v in col_D]
            output_sheet.range(f"E{start_output_row}:E{end_row}").value = [[v] for v in col_E]
            output_sheet.range(f"F{start_output_row}:F{end_row}").value = [[v] for v in col_F]
            output_sheet.range(f"G{start_output_row}:G{end_row}").value = [[v] for v in col_G]
            output_sheet.range(f"H{start_output_row}:H{end_row}").value = [[v] for v in col_H]

            # Step 3.5: Set font style to Calibri for column A
            output_sheet.range(f"A{start_output_row}:A{end_row}").font.name = "Calibri"

            # Step 4: Apply bold formatting only where needed
            for i, is_bold in enumerate(bold_mask):
                if is_bold:
                    row = start_output_row + i
                    output_sheet.range(f"B{row}").font.bold = True

            # Step 5: Set print area
            print_range = f"A1:H{last_written_row}"
            output_sheet.api.PageSetup.PrintArea = f"${print_range.replace(':', ':$')}"
        
