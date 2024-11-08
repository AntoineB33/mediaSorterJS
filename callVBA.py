import os
import win32com.client as win32
import pythoncom
from win32com.client import VARIANT


def call_vba_function(macro_name, wbSheet, order):
    file_path = r"C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\excel_prog\mediaSorter\dance.xlsm"
    excel_app = win32.DispatchEx("Excel.Application")
    excel_app.Visible = False

    # Open the workbook
    workbook = excel_app.Workbooks.Open(file_path)

    # Convert Python lists to VARIANT arrays if necessary
    converted_args = [wbSheet]
    variant_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, order)
    converted_args.append(variant_array)
    print(f"'{workbook.Name}'!{macro_name}")
    print(converted_args)
    # Call the VBA macro with parameters
    result = excel_app.Application.Run(f"'{workbook.Name}'!{macro_name}", *converted_args)

    # Close the workbook and quit Excel
    workbook.Close(SaveChanges=False)
    excel_app.Quit()

    # Explicitly release COM objects
    workbook = None
    excel_app = None

    return result
