# src/main.py
import os
import sys
import xlwings as xw
import tkinter as tk
import re
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from pathlib import Path
import shutil
import uuid
from datetime import datetime
import time


# Import the converter classes.
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
        found = [sheet.name for sheet in wb.sheets if sheet.name.strip() in EXPECTED_SHEETS]
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

def convert_report(input_filepath, sheets_to_convert, progress_callback=None, save_to_input_dir=False):
    try:
        # Create a temporary copy of the template to work with
        temp_template_filename = f"temp_template_{uuid.uuid4().hex}.xlsx"
        temp_template_path = os.path.join(os.getcwd(), temp_template_filename)
        shutil.copy(TEMPLATE_PATH, temp_template_path)

        # Launch Excel
        app = xw.App(visible=False)
        print(f"xwings has opened")
        time.sleep(0.5)
        app.api.ScreenUpdating = False
        app.api.DisplayAlerts = False

        input_wb = xw.Book(input_filepath, update_links=False)
        print(f"{input_filepath} has been opened with xlwings")
        template_wb = xw.Book(temp_template_path, update_links=False)
        print(f"{template_wb} has been opened with xlwings")

        # Process each confirmed sheet
        total = len(sheets_to_convert)
        print(f"number of sheets to convert: {total}")
        for idx, sheet_name in enumerate(sheets_to_convert, start=1):
            print(f"idx: {idx}")        
            input_sheet = input_wb.sheets[sheet_name]
            print(f"{sheet_name} sheet retrieved")
            converter_class = CONVERTER_MAPPING.get(sheet_name, DefaultConverter)
            print(f"{converter_class} class has been grabbed from mapping")
            converter = converter_class(input_sheet, template_wb)
            print(f"{converter} has been instantiated with template_wb")
            converter.convert()
            print(f"Converted sheet: {sheet_name}")

            if progress_callback:
                progress_callback(sheet_name, idx, total)

        input_path = Path(input_filepath)
        input_name = input_path.stem
        input_dir = input_path.parent

        # Replace "V7" (case-insensitive) while preserving case
        match = re.search(r"v7", input_name, re.IGNORECASE)
        if match:
            start, end = match.span()
            output_name = input_name[:start] + "S536-19 v1.0" + input_name[end:]
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_name = f"Converted_Report_{timestamp}"

        output_filename = f"{output_name}.xlsx"
        output_dir = os.path.dirname(input_filepath) if save_to_input_dir else os.getcwd()
        output_filepath = os.path.join(output_dir, output_filename)

        if os.path.exists(output_filepath):
            os.remove(output_filepath)
        
        time.sleep(0.5)
        try:
            template_wb.save(output_filepath)
        except Exception as save_error:
            messagebox.showerror("Save Error", f"Failed to save output file:\n{save_error}")
            raise

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

def update_progress(sheet_name, idx, total):
    progress_percent = int((idx / total) * 100)
    progress_bar['value'] = progress_percent

    padded_text = f"✅ Converted: {sheet_name} ({idx}/{total})".ljust(60)  # <- pad to clear previous
    progress_label.config(text=padded_text)
    root.update_idletasks()


def select_file_and_convert():
    progress_bar['value'] = 0
    progress_label.config(text="Starting conversion...")
    root.update_idletasks()

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
        output_file = convert_report(
            filepath,
            confirmed_sheets,
            progress_callback=update_progress,
            save_to_input_dir=bool(save_in_same_dir_var.get())
        )
        if output_file:
            messagebox.showinfo("Conversion Complete", f"Converted file saved at:\n{output_file}")
            root.quit()  # Close the program


def handle_drop(event):
    progress_bar['value'] = 0
    progress_label.config(text="Starting conversion...")
    root.update_idletasks()

    filepath = event.data.strip("{}")  # Remove {} from file path if present
    if filepath.lower().endswith(".xlsx"):
        drop_zone.config(bg="#d4edda", fg="#155724", text="✅ Valid file received, processing...")
        convert_button.config(state="disabled")  # 👈 Disable the button here
        root.update_idletasks()

        found_sheets = detect_expected_sheets(filepath)
        if not found_sheets:
            drop_zone.config(bg="#fff3cd", fg="#856404", text="⚠️ No expected sheets found")
            messagebox.showerror("Error", "No expected sheets found in the selected file.")
            convert_button.config(state="normal")  # 👈 Re-enable on error
            return

        confirmed_sheets = confirm_sheets_dialog(found_sheets, root)
        if not confirmed_sheets:
            drop_zone.config(bg="#f0f0f0", fg="#333", text="⬇️ Drop Excel file here ⬇️")
            messagebox.showinfo("Info", "No sheets selected for conversion.")
            convert_button.config(state="normal")  # 👈 Re-enable on cancel
            return

        output_file = convert_report(
            filepath,
            confirmed_sheets,
            progress_callback=update_progress,
            save_to_input_dir=bool(save_in_same_dir_var.get())
        )
        if output_file:
            drop_zone.config(bg="#d1ecf1", fg="#0c5460", text="✅ Conversion complete!")
            messagebox.showinfo("Conversion Complete", f"Converted file saved at:\n{output_file}")
            root.destroy()
    else:
        drop_zone.config(bg="#f8d7da", fg="#721c24", text="❌ Invalid file type. Drop a .xlsx file.")
        messagebox.showwarning("Invalid File", "Please drop a valid .xlsx file.")



def main():
    global root, drop_zone, progress_bar, progress_label, convert_button
    root = TkinterDnD.Tk()
    root.title("Report Converter")
    root.geometry("450x280")  # Slightly taller for progress bar

    tk.Label(root, text="Select or drop an Excel file to convert:").pack(pady=10)

    drop_zone = tk.Label(root, text="⬇️ Drop Excel file here ⬇️", relief="ridge", borderwidth=2,
                         width=40, height=4, bg="#f0f0f0", fg="#333")
    drop_zone.pack(pady=10)
    drop_zone.drop_target_register(DND_FILES)
    drop_zone.dnd_bind("<<Drop>>", handle_drop)

    convert_button = tk.Button(root, text="Browse and Convert", command=select_file_and_convert, padx=10, pady=5)
    convert_button.pack(pady=5)

    # Progress Label and Bar
    progress_label = tk.Label(root, text="", fg="gray")
    progress_label.pack(pady=(5, 0))
    
    progress_bar = ttk.Progressbar(root, length=300, mode='determinate')
    progress_bar.pack(pady=(0, 10))

    global save_in_same_dir_var
    save_in_same_dir_var = tk.IntVar(value=1)  # default to unchecked

    save_dir_checkbox = tk.Checkbutton(
        root, 
        text="Save in original file location", 
        variable=save_in_same_dir_var
    )
    save_dir_checkbox.pack()

    root.mainloop()



if __name__ == "__main__":
    main()
