














Private Sub ImgButton_Click()
    StoreImageFromClipboard
    DisplayStoredImage
End Sub

Private Sub WatchButton_Click()
    ThisWorkbook.CallJavaScriptFunctionAsync "show", False, "orderedVideos", 0
End Sub

Private Sub watchFromButton_Click()
    ThisWorkbook.CallJavaScriptFunctionAsync "show", False, "orderedVideos", editRow - 2
End Sub

Private Sub SortButton_Click()
    sheetVBA.Protect UserInterfaceOnly:=sorting(sheetCodeName)
    If sorting(sheetCodeName) Then
        SortButton.Caption = "Sort"
        updateButton.Visible = False
    Else
        SortButton.Caption = "Stop sorting"
        ThisWorkbook.CallJavaScriptFunctionAsync "dataGeneratorSub", True
        updateButton.Visible = True
    End If
    sorting(sheetCodeName) = Not sorting(sheetCodeName)
End Sub

Private Sub ThumbnailsButton_Click()
    DisplayImagesInUserForm
End Sub

Private Sub ctrlZButton_Click()
    ThisWorkbook.CallJavaScriptFunctionAsync "ctrlZ", True
End Sub

Private Sub ctrlYButton_Click()
    ThisWorkbook.CallJavaScriptFunctionAsync "ctrlY", True
End Sub

Private Sub PowerPointButton_Click()

End Sub

Private Sub oldNameInput_Change()
    ThisWorkbook.CallJavaScriptFunctionAsync "handleoldNameInputClick", True
End Sub

Private Sub oldNameInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Me.newNameInput.SetFocus
    End If
End Sub

Private Sub newNameInput_Change()
    newNameInput.Text = oldNameInput.Text
End Sub

Private Sub newNameInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ThisWorkbook.CallJavaScriptFunctionAsync "renameSymbol", True, "oldValue", oldNameInput.Text, "newValue", newNameInput.Text
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Set UserForm properties
    Me.BackColor = RGB(255, 255, 255) ' Set background color to white (same as Excel)

    
    ' Get the width of the Excel application window
    Dim appWidth As Long
    appWidth = Application.Width

    ' Get the left position of the Excel application window
    Dim appLeft As Long
    appLeft = Application.Left
    
    ' Get the width of the active window (workbook window)
    Dim sheetWidth As Long
    sheetWidth = ActiveWindow.Width

    ' Set the Left property of the UserForm to position it more to the right
    ' We'll position the UserForm near the right edge of the Excel application window
    Me.Left = appLeft + (sheetWidth * 0.75) ' Adjust multiplier to shift more or less to the right

    ' Optionally, set the Top property if you want to adjust the vertical position
    Me.Top = Application.Top + 100 ' Example, adjust as needed
End Sub

Private Sub UserForm_Click()

End Sub