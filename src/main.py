# src/main.py

import os
import xlwings as xw
import tkinter as tk
from tkinter import filedialog, messagebox

# Path to your pre-formatted template
TEMPLATE_FILENAME = "Annual ULC Template - CAN,ULC-S536-19 v1.0.xlsx"
TEMPLATE_PATH = os.path.join(os.getcwd(), "report_templates", TEMPLATE_FILENAME)

def convert_report(input_filepath):
    try:
        # Launch Excel invisibly
        app = xw.App(visible=False)
        
        # Open the input workbook (for reading data if needed)
        input_wb = app.books.open(input_filepath)
        
        # Open the template workbook that has your formatting
        template_wb = app.books.open(TEMPLATE_PATH)
        
        # Example logic: Copy a value from input sheet A1 to template sheet B2
        source_value = input_wb.sheets[0]['A1'].value
        template_wb.sheets[0]['B2'].value = source_value
        
        # Define an output file path; here we simply save it in the project root
        output_filepath = os.path.join(os.getcwd(), "Converted_Report.xlsx")
        template_wb.save(output_filepath)
        
        # Close workbooks and quit Excel
        input_wb.close()
        template_wb.close()
        app.quit()

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
        output_file = convert_report(filepath)
        if output_file:
            messagebox.showinfo("Conversion Complete", f"Converted file saved at:\n{output_file}")
            # Optionally, open the file:
            os.startfile(output_file)

def main():
    # Create the main Tkinter window
    root = tk.Tk()
    root.title("Report Converter")
    root.geometry("400x150")
    
    # Create and pack widgets
    label = tk.Label(root, text="Select an Excel file to convert:")
    label.pack(pady=10)
    
    convert_button = tk.Button(root, text="Browse and Convert", command=select_file_and_convert, padx=10, pady=5)
    convert_button.pack()
    
    # Run the Tkinter event loop
    root.mainloop()

if __name__ == "__main__":
    main()
