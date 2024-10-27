' Feuil1.cls
' Class module for the worksheet Feuil1






Private Sub Worksheet_SelectionChange(ByVal Target As range)
    ' Call the sub from ThisWorkbook
    editRow = activeCell.Row
    Dim selectedRow As Long
    selectedRow = editRow
    editColumn = activeCell.Column
    CallJavaScriptFunctionAsync "selectionChange", True, "editRow", editRow, "editCol", editColumn
    
    If thbnlDisplayed Then
        RemoveAllControls ImageListForm
        DisplayImagesInUserForm
    End If
End Sub

