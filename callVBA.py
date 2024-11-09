import os
import win32com.client as win32
import pythoncom
from win32com.client import VARIANT


def call_vba_function(macro_name, wbSheet, order):
    file_path = r"C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\excel_prog\mediaSorter\dance.xlsm"
    
    # Try to get an existing Excel instance
    try:
        excel_app = win32.GetObject(None, "Excel.Application")
        is_new_instance = False
    except Exception:
        # If Excel is not open, create a new instance
        excel_app = win32.DispatchEx("Excel.Application")
        is_new_instance = True
    
    # Preserve the initial visibility state
    initial_visibility = excel_app.Visible
    
    # Only change visibility if it's a new instance
    if is_new_instance:
        excel_app.Visible = False  # Set to True if you want to see Excel during execution

    # Open the workbook or get the already open workbook
    workbook = None
    for wb in excel_app.Workbooks:
        if wb.FullName == file_path:
            workbook = wb
            break
    if not workbook:
        workbook = excel_app.Workbooks.Open(file_path)

    # Convert Python lists to VARIANT arrays if necessary
    converted_args = [wbSheet]
    if order:
        variant_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, order)
        converted_args.append(variant_array)
    # macro_name = "test3"
    print(f"'{workbook.Name}'!{macro_name}")
    print(converted_args)
    # Call the VBA macro with parameters
    result = excel_app.Application.Run(f"'{workbook.Name}'!{macro_name}", *converted_args)
    # result = excel_app.Application.Run(f"'{workbook.Name}'!test")

    # Close the workbook if it was opened by this function
    if workbook and workbook.FullName == file_path and is_new_instance:
        workbook.Close(SaveChanges=False)

    # Quit Excel only if it was a new instance, otherwise restore initial visibility
    if is_new_instance:
        excel_app.Quit()
    else:
        excel_app.Visible = initial_visibility

    # Explicitly release COM objects
    workbook = None
    excel_app = None

    return result

if __name__ == "__main__":
    call_vba_function("SwapRows", "Feuil3", [2, 4, 3])