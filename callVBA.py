import os
import win32com.client as win32
import pythoncom
from win32com.client import VARIANT
import requests
import os
from datetime import datetime


def retrieve_excel_file_by_id(directory: str, id: str):
    # Convert the ID back to a datetime object
    file_date = datetime.strptime(id, "%Y%m%d_%H%M%S")
    
    # Loop through each file in the directory
    for filename in os.listdir(directory):
        # Construct the full file path
        file_path = os.path.join(directory, filename)
        
        # Check if it is an Excel file
        if filename.endswith(('.xlsx', '.xls')) and os.path.isfile(file_path):
            # Get the file's creation time
            creation_time = datetime.fromtimestamp(os.path.getctime(file_path))
            
            # Compare the file's creation time with the ID datetime
            if creation_time == file_date:
                print(f"Found matching file: {file_path}")
                return file_path
    
    # If no matching file is found
    print("No file found with the matching creation date.")
    return None

def call_node_endpoint(url: str, message: str):
    # Add the string parameter as a query parameter
    params = {'message': message}
    
    try:
        # Send GET request with the query parameter
        response = requests.get(url, params=params)
        
        # Check if the request was successful
        if response.status_code == 200:
            print("Response from Node.js:", response.json())
        else:
            print(f"Error: {response.status_code} - {response.text}")
    
    except requests.exceptions.RequestException as e:
        print(f"Request failed: {e}")

def call_vba_function(wbSheet: str):
    file_path = retrieve_excel_file_by_id(r"C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\excel_prog\mediaSorter",
                                          wbSheet)
    
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

    # Call the public VBA function and get the result
    result = workbook.Application.Run(f"'{workbook.Name}'!WhatEndPoint")
    call_node_endpoint(result, wbSheet)

    # Restore initial visibility and close workbook if it's a new instance
    if is_new_instance:
        workbook.Close(False)  # Close without saving
        excel_app.Visible = initial_visibility
        excel_app.Quit()

    return result

if __name__ == "__main__":
    call_vba_function("SwapRows", "Feuil3", [2, 4, 3])