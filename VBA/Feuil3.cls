' Feuil1.cls
' Class module for the worksheet Feuil1






Private Sub Worksheet_SelectionChange(ByVal Target As range)
    Dim result As Variant
    result = GetRowAndColumn(activeCell.Address(False, False))
    editRow = result(0)
    editColumn = result(1)
    CallJavaScriptFunctionAsync "selectionChange", True, "editRow", editRow, "editCol", editColumn
    
    If thbnlDisplayed Then
        RemoveAllControls ImageListForm
        DisplayImagesInUserForm
    End If
End Sub

