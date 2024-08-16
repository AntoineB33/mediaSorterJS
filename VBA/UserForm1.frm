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
    sorting(sheetCodeName) = not sorting(sheetCodeName)
End Sub

Private Sub updateButton_Click()
    
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

Sub LabelClicked()
    Dim clickedLabel As MSForms.Label
    
    ' Identify which label was clicked
    Set clickedLabel = Me.Controls(Application.Caller)
    
    ' Show message based on the clicked label
    MsgBox "You clicked: " & clickedLabel.Caption, vbInformation
End Sub

Private Sub UserForm_Initialize()
    ' Set UserForm properties
    Me.BackColor = RGB(255, 255, 255) ' Set background color to white (same as Excel)
End Sub


Private Sub UserForm_Click()

End Sub

