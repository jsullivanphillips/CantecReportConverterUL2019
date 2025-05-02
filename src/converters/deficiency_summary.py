# src/converters/deficiency_summary.py

from .base import BaseSheetConverter

class DeficiencySummaryConverter(BaseSheetConverter):
    NO_BORDER = -4142
    BORDER = 1

    NUM_ROWS_PER_SECTION = [
            1, # owner_responsible_01
            2, # owner_responsible_02
            3, # Fire Alarm
            3, # Emergency Lights
            3, # Extinguishers
            2, # Hoses
            2  # Other
        ]

    OUTPUT_SECTION_ROW_START = [
        10,
        13,
        20,
        26,
        32,
        38,
        43
    ]

    def does_cell_have_border(self, cell):
        if cell is None:
            print("[Border Check] Cell is None → returning False")
            return False

        try:
            for side in (1, 2, 3, 4):  # Left, Right, Top, Bottom
                border = cell.api.Borders(side)
                print(f"[Border Check] {cell.address} side {side} → LineStyle = {border.LineStyle}")

                if border.LineStyle != self.BORDER:
                    print(f"[Border Check] {cell.address} → side {side} does not match BORDER ({self.BORDER}) → returning False")
                    return False

            print(f"[Border Check] {cell.address} → All sides match BORDER ({self.BORDER}) → returning True")
            return True

        except AttributeError as e:
            print(f"[Border Check] AttributeError on {cell.address if cell else 'Unknown'}: {e}")
            return False


    def is_cell_empty(self, cell):
        if cell is None:
            return True

        try:
            value = cell.value
            # Consider empty if None or empty string after stripping
            if value is None:
                return True
            if isinstance(value, str) and value.strip() == "":
                return True
            return False
        except AttributeError:
            # If the cell object is broken somehow
            return True

    def put_content_to_output(self, real_input_row, section_index, content_col, rel_row_to_place_content, rows_added_to_output, output_sheet):
        row_to_place_content = self.OUTPUT_SECTION_ROW_START[section_index] + rows_added_to_output + rel_row_to_place_content
        output_sheet.range(f"A{row_to_place_content}").value = self.get_from_input_cell(f"A{real_input_row}")
        if content_col == 1:
            output_sheet.range(f"B{row_to_place_content}").value = self.get_from_input_cell(f"B{real_input_row}")
            self.apply_wrap_text_to_output_cell(
                sheet_index_or_name="Deficiency Summary",
                cell=f"B{row_to_place_content}",
                wrap=True
            )

        self.copy_row_height_to_output(input_row_number=real_input_row, output_row_number=row_to_place_content, sheet_index_or_name="Deficiency Summary")
        

    def apply_thick_bottom_border(self, row_number, output_sheet):
        """
        Applies a thick bottom border (LineStyle=1, Weight=4, Color=0.0) across merged regions
        of the specified row. Only applies to the top-left cell of each merged area.
        """
        if output_sheet.api.ProtectContents:
            output_sheet.api.Unprotect()
        
        row_range = output_sheet.range(f"A{row_number}:L{row_number}")

        for cell in row_range:
            try:
                border = cell.api.Borders(4)  # Bottom border

                border.LineStyle = 1   # xlContinuous
                border.Weight = 3      # xlThick
                border.Color = 0.0     # Black

                print(f"[Thick Bottom Border] Applied to merged range: {cell.api.Address} (row {row_number})")

            except Exception as e:
                print(f"[Thick Bottom Border] Failed at cell {cell.address}: {e}")

        
    ## TODO:
        # - Deficiency summary: still has a thick border in the middle of the section
        # - EXT and ELU pages: need the ACT column font colour needs to be transfer - Set to black if no input value
        # - Device Log: Would be great if new device log could auto size row height based on content. Look into this.
        # - Device Log: smoke alarms don't have any check marks.

    def convert(self):
        output_sheet = self.output_wb.sheets["Deficiency Summary"]
        if output_sheet.api.ProtectContents:
            output_sheet.api.Unprotect()
        
        start_input_row = 20
        end_input_row = 60

        input_range = self.input_sheet.range(f"A{start_input_row}:L{end_input_row}")
        input_rows = input_range.rows


        # Sections to check
        section_headers = [
            "owner_responsible_01",
            "owner_responsible_02",
            "Fire Alarm",
            "Emergency Lights",
            "Extinguisher",
            "Hoses",
            "Other"
        ]

        
        rows_added_to_output = 0

        # Set when we hit a row with no borders and the previous row had full borders
        inbetween_sections = False 
        previous_row_had_border = False
        section_index = 0
        num_input_rows_for_section = 0

        print(f"Starting conversion...")
    
        for rel_row_idx, row_cells in enumerate(input_rows):
            real_input_row = start_input_row + rel_row_idx # Input row on sheet
            col_a_cell = row_cells[0]  # Column A Cell
            col_a_value = col_a_cell.value  # Column A Value

            has_border = self.does_cell_have_border(col_a_cell)

            print(f"Row {real_input_row}: has_border={has_border}, inbetween_sections={inbetween_sections}, section_index={section_index}, col_a_value={col_a_value}")
            ## TODO: 
            #   - Test with wild wacky inputs
            #   - Copy row height (15px, 48px, 96px) and formatting (bold if posisble) to output sheet
            if has_border:
                if inbetween_sections:
                    if not self.is_cell_empty(col_a_cell) and str(col_a_value).strip().lower() == "quantity":
                        print(f"Found new section header 'Quantity' at row {real_input_row}")
                        inbetween_sections = False 
                        section_index += 1
                        print(f"Incremented section_index to {section_index}")
                elif not inbetween_sections:
                    content_col = 0 if section_index == 0 else 1
                    
                    if section_index == 0 and num_input_rows_for_section > 0:
                        print(f"Skipping extra MBT rows in owner_responsible_01 at row {real_input_row}")
                        continue

                    if not self.is_cell_empty(row_cells[content_col]):
                        print(f"Found content at row {real_input_row} (content_col={content_col})")

                        if num_input_rows_for_section >= self.NUM_ROWS_PER_SECTION[section_index]:
                            print(f"Exceeded predefined number of rows ({self.NUM_ROWS_PER_SECTION[section_index]}) for section {section_headers[section_index]}")
                            row_to_clone = self.OUTPUT_SECTION_ROW_START[section_index] + rows_added_to_output + num_input_rows_for_section - 1
                            self.insert_formatted_row_below(output_sheet, row_to_clone, log_prefix="[Section Clone]")
                            print(f"Inserted new row, rows_added_to_output now {rows_added_to_output}")
                            self.put_content_to_output(real_input_row, section_index, content_col, num_input_rows_for_section, rows_added_to_output, output_sheet)
                            rows_added_to_output += 1
                        else:
                            self.put_content_to_output(real_input_row, section_index, content_col, num_input_rows_for_section, rows_added_to_output, output_sheet)
                        
                        print(f"Transferred content from input row {real_input_row} to output section {section_headers[section_index]} row {num_input_rows_for_section}")
                        num_input_rows_for_section += 1
                        
            if not has_border and previous_row_had_border:
                if not inbetween_sections:
                    print(f"Entering inbetween_sections after row {real_input_row}")
                    inbetween_sections = True
                    if section_index == 1:
                        previous_output_row = self.OUTPUT_SECTION_ROW_START[section_index] + rows_added_to_output + num_input_rows_for_section - 2
                        self.apply_thick_bottom_border(previous_output_row, output_sheet)
                    num_input_rows_for_section = 0

            
            previous_row_had_border = has_border

        print("Finished conversion.")


            