# src/main.py
import os
import sys
import xlwings as xw
import tkinter as tk
from tkinter import filedialog, messagebox
import shutil
import uuid

# Import the converter classes.
from converters.ulc_coverpage import ULCCoverpageConverter
from converters.deficiency_summary import DeficiencySummaryConverter
from converters.appendix import AppendixConverter
from converters.log_report import LogReportConverter
from converters.ext_only import ExtOnlyConverter
from converters.elu_only import EluOnlyConverter
from converters.ulc_c2 import ULCC2Converter
from converters.field_device_testing import FieldDeviceTestingConverter
from converters.base import DefaultConverter

# List of expected sheet names exactly as they appear.
EXPECTED_SHEETS = [
    "ULC Coverpage",
    "DEFICIENCY SUMMARY",
    "APPENDIX C1+C2.13 2.14 2.15",
    "LOG REPORT C3.2- Device Record",
    "EXT only",
    "ELU only",
    "ULC - C2.1-2.12",
    "C3.1FieldDeviceTesting-Legend"
]

def resource_path(relative_path):
    """
    Get absolute path to a resource, compatible with PyInstaller's --add-data.
    """
    try:
        # PyInstaller stores data files in a temp folder referenced by _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # Normal execution (not bundled)
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Path to your template file.
TEMPLATE_FILENAME = "Annual ULC Template - CAN,ULC-S536-19 v1.0.xlsx"
TEMPLATE_PATH = resource_path(os.path.join("report_templates", TEMPLATE_FILENAME))

# Dispatcher: mapping input sheet names to converter classes.
CONVERTER_MAPPING = {
    "ULC Coverpage": ULCCoverpageConverter,
    "DEFICIENCY SUMMARY": DeficiencySummaryConverter,
    "APPENDIX C1+C2.13 2.14 2.15": AppendixConverter,
    "LOG REPORT C3.2- Device Record": LogReportConverter,
    "EXT only": ExtOnlyConverter,
    "ELU only": EluOnlyConverter,
    "ULC - C2.1-2.12": ULCC2Converter,
    "C3.1FieldDeviceTesting-Legend": FieldDeviceTestingConverter,
}

def detect_expected_sheets(input_filepath):
    """Open the input workbook and return a list of expected sheet names found."""
    try:
        app = xw.App(visible=False)
        wb = app.books.open(input_filepath)
        found = [sheet.name for sheet in wb.sheets if sheet.name in EXPECTED_SHEETS]
        wb.close()
        app.quit()
        return found
    except Exception as e:
        messagebox.showerror("Error", f"Error detecting sheets:\n{e}")
        return []

def confirm_sheets_dialog(found_sheets, parent):
    """Displays a modal dialog with checkboxes for each detected sheet.
    Returns the list of sheets the user confirms."""
    dialog = tk.Toplevel(parent)
    dialog.title("Confirm Sheets to Convert")
    dialog.geometry("300x300")
    dialog.grab_set()  # Modal dialog

    tk.Label(dialog, text="Select the sheets to process:").pack(pady=10)

    sheet_vars = {}
    for sheet in found_sheets:
        var = tk.IntVar(value=1)  # Pre-selected by default
        sheet_vars[sheet] = var
        tk.Checkbutton(dialog, text=sheet, variable=var).pack(anchor="w", padx=20)

    confirmed_sheets = []

    def on_confirm():
        for sheet, var in sheet_vars.items():
            if var.get() == 1:
                confirmed_sheets.append(sheet)
        dialog.destroy()

    tk.Button(dialog, text="Confirm", command=on_confirm).pack(pady=10)
    parent.wait_window(dialog)
    return confirmed_sheets

def convert_report(input_filepath, sheets_to_convert):
    try:
        # Create a temporary copy of the template to work with
        temp_template_filename = f"temp_template_{uuid.uuid4().hex}.xlsx"
        temp_template_path = os.path.join(os.getcwd(), temp_template_filename)
        shutil.copy(TEMPLATE_PATH, temp_template_path)

        # Launch Excel
        app = xw.App(visible=True)
        input_wb = app.books.open(input_filepath)
        template_wb = app.books.open(temp_template_path)

        # Process each confirmed sheet
        for sheet_name in sheets_to_convert:
            input_sheet = input_wb.sheets[sheet_name]
            converter_class = CONVERTER_MAPPING.get(sheet_name, DefaultConverter)
            converter = converter_class(input_sheet, template_wb)
            converter.convert()
            print(f"Converted sheet: {sheet_name}")

        # Save the modified copy as the final output
        output_filepath = os.path.join(os.getcwd(), "Converted_Report.xlsx")
        template_wb.save(output_filepath)

        # Clean up
        input_wb.close()
        template_wb.close()
        app.quit()

        # Optionally delete the temporary copy now that it's saved
        os.remove(temp_template_path)

        return output_filepath

    except Exception as e:
        messagebox.showerror("Conversion Error", f"An error occurred:\n{e}")
        return None

def select_file_and_convert():
    filepath = filedialog.askopenfilename(
        title="Select an Excel file to convert",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if filepath:
        found_sheets = detect_expected_sheets(filepath)
        if not found_sheets:
            messagebox.showerror("Error", "No expected sheets found in the selected file.")
            return
        confirmed_sheets = confirm_sheets_dialog(found_sheets, root)
        if not confirmed_sheets:
            messagebox.showinfo("Info", "No sheets selected for conversion.")
            return
        print("Confirmed sheets:", confirmed_sheets)
        output_file = convert_report(filepath, confirmed_sheets)
        if output_file:
            messagebox.showinfo("Conversion Complete", f"Converted file saved at:\n{output_file}")
            os.startfile(output_file)

def main():
    global root
    root = tk.Tk()
    root.title("Report Converter")
    root.geometry("400x150")

    label = tk.Label(root, text="Select an Excel file to convert:")
    label.pack(pady=10)

    convert_button = tk.Button(root, text="Browse and Convert", command=select_file_and_convert, padx=10, pady=5)
    convert_button.pack()

    root.mainloop()

if __name__ == "__main__":
    main()
