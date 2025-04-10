# src/converters/log_report.py

from .base import BaseSheetConverter

"""
get_from_input_cell(self, cell):

put_to_output_cell(self, sheet_index_or_name, cell, value):
"""

class LogReportConverter(BaseSheetConverter):
    """
    Conversion logic for the "LOG REPORT C3.2- Device Record" sheet.
    For demonstration, copies A1 from input to B6 of the output.
    """
    def copy_bold_if_set(self, input_cell_ref, output_cell_ref, output_sheet_name):
        try:
            input_cell = self.input_sheet.range(input_cell_ref)
            output_cell = self.output_wb.sheets[output_sheet_name].range(output_cell_ref)

            is_bold = input_cell.font.bold

            if is_bold is True:
                output_cell.font.bold = True

        except Exception as e:
            print(f"[ERROR] Could not copy bold from {input_cell_ref} to {output_cell_ref}: {e}")


    def clean_column(self, col):
        return [[v if v != "" else None] for v in col]

    def convert(self):
        # region To 20.1 | Report
        building_name = self.get_from_input_cell('C4')
        self.put_to_output_cell("20.1 | Report", 'F9', building_name)

        address = self.get_from_input_cell('C5')
        self.put_to_output_cell("20.1 | Report", 'F10', address)
        # endregion


        # region To 22.1 | CU or Transp Insp
        fap_location = self.get_from_input_cell('H5')
        self.put_to_output_cell("22.1 | CU or Transp Insp", 'H15', fap_location)

        fap_make = self.get_from_input_cell('H3')
        fap_model = self.get_from_input_cell('H4')
        fap_identification = fap_make + " " + fap_model
        self.put_to_output_cell("22.1 | CU or Transp Insp", 'H16', fap_identification)
        # endregion


        # region To 22.5 | Power Supply(s)
        battery_info = self.get_from_input_cell('P4')
        try:
            parts = battery_info.strip().split()
            voltage = parts[0]            # "12V"
            amps = parts[1]               # "7.2AH"
            count = int(parts[2][1:])     # "X2" → 2
        except (AttributeError, IndexError, ValueError):
            voltage = amps = count = None  # Fallback values if parsing fails
        
        self.put_to_output_cell("22.5 | Power Supply(s)", 'E12', count)
        self.put_to_output_cell("22.5 | Power Supply(s)", 'G12', voltage)
        self.put_to_output_cell("22.5 | Power Supply(s)", 'I12', amps)
        # endregion


        # region 22.6 | Annunciator(s)
        annunciator_location = self.get_from_input_cell('P5')
        self.put_to_output_cell("22.6 | Annunciator(s)", "G7", annunciator_location)
        # endregion

        
        # region 23.2 Device Record
        output_sheet = self.output_wb.sheets["23.2 Device Record"]

        # 🔓 Unprotect once at the beginning
        if output_sheet.api.ProtectContents:
            output_sheet.api.Unprotect()

        # Step 1: Read input data
        input_data = self.input_sheet.range("B22:P2000").value
        start_output_row = 14
        max_input_rows = len(input_data)

        col_A, col_C, col_D, col_E, col_G, col_I, col_J, col_M, bold_mask = [], [], [], [], [], [], [], [], []

        number_of_consecutive_empty_rows = 0
        last_written_row = start_output_row - 1

        for rel_row, row_data in enumerate(input_data):
            input_row = 22 + rel_row
            output_row = start_output_row + rel_row
            last_written_row = output_row

            device_location = row_data[0]  # column B

            # Track empty rows
            if not device_location or str(device_location).strip() == "":
                number_of_consecutive_empty_rows += 1
            else:
                number_of_consecutive_empty_rows = 0

            if number_of_consecutive_empty_rows >= 100:
                break

            # Always record row (even if blank) to preserve spacing
            col_M_data = ""
            sa_replacement_year = row_data[11]
            if sa_replacement_year is not None:
                col_M_data = f"Due to be replaced in {int(sa_replacement_year)}."
                if row_data[14] is not None:
                    col_M_data += " " + str(row_data[14])
            else:
                col_M_data = row_data[14]
            
            # check for failures
            operation_confirmed = "✖" if row_data[3] == 5 else ""
            annunciation_confirmed = "✖" if row_data[4] == 5 else ""
            installed_correctly = "✖" if row_data[12] == 5 else ""


            col_A.append(device_location)
            col_C.append(row_data[2])   # From D
            col_D.append(row_data[13])  # From O
            col_E.append(row_data[6])   # From H
            col_G.append(installed_correctly)
            col_I.append(operation_confirmed)
            col_J.append(annunciation_confirmed)
            col_M.append(col_M_data)  # From P

            # Track bold (True/False/None)
            is_bold = self.input_sheet.range(f"B{input_row}").font.bold
            bold_mask.append(is_bold is True)

        # Step 2: Determine end row
        end_row = start_output_row + len(col_A) - 1

        # Step 3: Bulk write to output sheet
        output_sheet.range(f"A{start_output_row}:A{end_row}").value = [[v] for v in col_A]
        output_sheet.range(f"C{start_output_row}:C{end_row}").value = [[v] for v in col_C]
        output_sheet.range(f"E{start_output_row}:E{end_row}").value = [[v] for v in col_E]
        output_sheet.range(f"D{start_output_row}:D{end_row}").value = [[v] for v in col_D]
        output_sheet.range(f"M{start_output_row}:M{end_row}").value = [[v] for v in col_M]

        output_sheet.range(f"G{start_output_row}:G{end_row}").value = self.clean_column(col_G)
        output_sheet.range(f"I{start_output_row}:I{end_row}").value = self.clean_column(col_I)
        output_sheet.range(f"J{start_output_row}:J{end_row}").value = self.clean_column(col_J)

        # Step 4: Apply bold formatting only where needed
        for i, is_bold in enumerate(bold_mask):
            if is_bold:
                row = start_output_row + i
                output_sheet.range(f"A{row}").font.bold = True

        # Step 5: Set print area
        print_range = f"A1:M{last_written_row + 5}"
        output_sheet.api.PageSetup.PrintArea = f"${print_range.replace(':', ':$')}"

        # 🔒 Re-protect
        output_sheet.api.Protect()
        # endregion









