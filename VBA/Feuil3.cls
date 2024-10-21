' Feuil1.cls
' Class module for the worksheet Feuil1






Private Sub Worksheet_SelectionChange(ByVal Target As range)
    ' Call the sub from ThisWorkbook
    UserForm1.editRow = activeCell.Row
    Dim selectedRow As Long
    selectedRow = UserForm1.editRow
    UserForm1.editColumn = activeCell.Column
    CallJavaScriptFunctionAsync "selectionChange", True, "editRow", UserForm1.editRow, "editCol", UserForm1.editColumn
    
    Dim imgControl As MSForms.Image
    Dim lblControl As MSForms.Label
    Dim frm As Object
    Dim controlName As String
    Dim i As Long

    ' Assuming you have already populated the UserForm with images
    Set frm = ImageListForm ' Replace with your actual UserForm name

    ' Check if the selected row corresponds to an image (assuming images start from row 2)
    If selectedRow >= 2 Then
        ' Loop through all controls on the UserForm and reset borders
        For i = 0 To frm.Controls.Count - 1
            If TypeOf frm.Controls(i) Is MSForms.Image Then
                Set imgControl = frm.Controls(i)
                ' Reset border color for all images
                imgControl.BorderColor = RGB(255, 255, 255) ' Set to white or transparent
            End If
        Next i

        ' If the selected row corresponds to an image, highlight it
        If selectedRow - 2 < frm.Controls.Count Then
            Set imgControl = frm.Controls(selectedRow - 2)
            
            ' Set border color to black for the selected image
            imgControl.BorderColor = RGB(0, 0, 0) ' Black frame
            imgControl.BorderStyle = fmBorderStyleSingle
        End If
    End If
End Sub

