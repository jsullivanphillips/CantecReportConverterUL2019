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
        fap_make = self.get_from_input_cell('H3')
        fap_model = self.get_from_input_cell('H4')

        fap_make = str(fap_make) if fap_make is not None else ""
        fap_model = str(fap_model) if fap_model is not None else ""

        fap_identification = fap_make + " " + fap_model
        self.put_to_output_cell("22.1 | CU or Transp Insp", 'H16', fap_identification)

        self.put_to_output_cell("20.1 | Report", "D13", fap_make)
        self.put_to_output_cell("20.1 | Report", "D14", fap_model)
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

        col_A, col_C, col_D, col_E, col_G, col_I, col_J, col_F, col_H, col_M, bold_mask = [], [], [], [], [], [], [], [], [], [], []

        number_of_consecutive_empty_rows = 0
        last_written_row = start_output_row - 1

        for rel_row, row_data in enumerate(input_data):
            input_row = 22 + rel_row
            output_row = start_output_row + rel_row
            

            device_location = row_data[0]  # column B

            # Track empty rows
            if not device_location or str(device_location).strip() == "":
                number_of_consecutive_empty_rows += 1
            else:
                number_of_consecutive_empty_rows = 0
                last_written_row = output_row

            if number_of_consecutive_empty_rows >= 60:
                break

            sa_replacement_year = row_data[11]
            skip_testing_data = (str(sa_replacement_year).strip().lower() == 'm')

            if skip_testing_data:
                # Only retain location, device type, and remarks
                col_A.append(device_location)
                col_C.append(row_data[2])
                col_D.append("-")      
                col_E.append("-")
                col_F.append("-")
                col_G.append("-")
                col_I.append("-")
                col_J.append("-")
                col_M.append(str(row_data[14]).lstrip() if row_data[14] is not None else "")
            else:
                col_H_data = ""
                if sa_replacement_year is not None:
                    try:
                        replacement_year = int(sa_replacement_year)
                        col_H_data = f"{replacement_year}"
                    except (ValueError, TypeError):
                        col_H_data = f"{sa_replacement_year}"
                

                operation_confirmed = "✖" if row_data[3] == 5 else ""
                annunciation_confirmed = "✖" if row_data[4] == 5 else ""
                installed_correctly = "N" if row_data[12] == 5 else ""

                loop = row_data[7]
                device = row_data[8]

                if loop is not None:
                    try:
                        loop_str = str(int(loop))
                    except (ValueError, TypeError):
                        loop_str = str(loop)
                else:
                    loop_str = ""

                device_str = str(device).zfill(3) if device is not None else ""

                if loop_str and device_str:
                    device_address_and_loop = f"Loop {loop_str}, {device_str}"
                else:
                    device_address_and_loop = ""

                col_A.append(device_location)
                col_C.append(row_data[2])   # From D
                col_D.append(row_data[13])  # From O
                col_E.append(device_address_and_loop)
                col_F.append(row_data[6])   # From H
                col_G.append(installed_correctly)
                col_I.append(operation_confirmed)
                col_J.append(annunciation_confirmed)
                col_H.append(col_H_data)
                col_M.append(str(row_data[14]).lstrip() if row_data[14] is not None else "")




            # Track bold (True/False/None)
            is_bold = self.input_sheet.range(f"B{input_row}").font.bold
            bold_mask.append(is_bold is True)

        # Step 2: Determine end row
        end_row = start_output_row + len(col_A) - 1

        # Step 3: Bulk write to output sheet
        output_sheet.range(f"A{start_output_row}:A{end_row}").value = [[v] for v in col_A]
        output_sheet.range(f"C{start_output_row}:C{end_row}").value = [[v] for v in col_C]
        output_sheet.range(f"E{start_output_row}:E{end_row}").value = [[v] for v in col_E]
        output_sheet.range(f"F{start_output_row}:F{end_row}").value = [[v] for v in col_F]
        output_sheet.range(f"H{start_output_row}:H{end_row}").value = [[v] for v in col_H]
        output_sheet.range(f"D{start_output_row}:D{end_row}").value = [[v] for v in col_D]
        output_sheet.range(f"M{start_output_row}:M{end_row}").value = [[v] for v in col_M]
        
        for i, (op_confirmed, ann_confirmed, install_corr) in enumerate(zip(col_I, col_J, col_G)):
            row = start_output_row + i
            if op_confirmed:  # only write if not empty
                output_sheet.range(f"I{row}").value = op_confirmed
            if ann_confirmed:
                output_sheet.range(f"J{row}").value = ann_confirmed
            if install_corr:
                output_sheet.range(f"G{row}").value = install_corr

        # Set row height based on length of remarks in col_M
        for i, (remark, location) in enumerate(zip(col_M, col_A)):
            if not remark and not location:
                continue

            def estimate_lines(text):
                text_str = str(text).strip()
                if not text_str:
                    return 1
                manual_lines = text_str.count('\n') + 1
                char_count = len(text_str) + 3  # small buffer for wrapping
                wrapped_lines = (char_count - 1) // 33 + 1
                return min(max(manual_lines, wrapped_lines), 6)

            lines_needed_remark = estimate_lines(remark)
            lines_needed_location = estimate_lines(location)

            lines_needed = max(lines_needed_remark, lines_needed_location)

            height = 15 if lines_needed == 1 else lines_needed * 12

            row = start_output_row + i
            output_sheet.range(f"{row}:{row}").row_height = height


        # Step 4: Apply bold formatting only where needed
        for i, is_bold in enumerate(bold_mask):
            if is_bold:
                row = start_output_row + i
                output_sheet.range(f"A{row}").font.bold = True

        # Step 5: Set print area
        print_range = f"A1:M{last_written_row + 5}"
        output_sheet.api.PageSetup.PrintArea = f"${print_range.replace(':', ':$')}"

        # endregion









