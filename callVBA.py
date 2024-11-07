import os
import win32com.client as win32
import pythoncom
import gc
from win32com.client import VARIANT


def call_vba_function(file_path, macro_name, *args):
    excel_app = win32.DispatchEx("Excel.Application")
    excel_app.Visible = False

    # Open the workbook
    workbook = excel_app.Workbooks.Open(file_path)

    # Convert Python lists to VARIANT arrays if necessary
    converted_args = []
    for arg in args:
        if isinstance(arg, list):
            # Convert list to VARIANT array
            variant_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, arg)
            converted_args.append(variant_array)
        else:
            converted_args.append(arg)

    # Call the VBA macro with parameters
    result = excel_app.Application.Run(f"'{workbook.Name}'!{macro_name}", *converted_args)

    # Close the workbook and quit Excel
    workbook.Close(SaveChanges=False)
    excel_app.Quit()

    # Explicitly release COM objects
    workbook = None
    excel_app = None

    # Force garbage collection to clean up COM objects
    gc.collect()

    return result

if  __name__ == '__main__':
    file_path = r"C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\excel_prog\mediaSorter\dance.xlsm"
    macro_name = "YourMacroName"
    param1 = "Parameter1"
    param2 = 42
    result = call_vba_function(file_path, macro_name, param1, param2)
    print("VBA Function Result:", result)
