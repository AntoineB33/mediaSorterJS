
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

Private Sub LinkToCell_Click()
    Dim selectedAddress As String
    Dim index As Integer
    
    ' Get the index of the selected item
    index = LinkToCell.ListIndex

    Dim item As Variant
    Dim acc As Integer
    Dim acc2 As Integer
    acc = 0
    For Each item In ThisWorkbook.listBoxList
        acc2 = acc + Ubound(item(0))
        acc = acc2 + Ubound(item(1))
        If acc2 < index and index < acc Then
            Dim subItem As Variant
            For Each subItem In item(1)(index - acc2)
                Select Case subItem.Keys
                Case "select"
                    subItem.Value.Select
                Case "newVal"
                    subItem.Value(0).Value = subItem.Value(1)
                Case "newBgCol"
                    subItem.Value(0).Interior.Color = subItem.Value(1)
                End Select
            Exit For
        End If
    Next item
End Sub

Private Sub suggestionList_Click()
    suggestionList.Visible = False
    suggestionList.Clear
    If ThisWorkbook.chgBckCol Then
        ThisWorkbook.cellWithError.Interior.Color = suggestionList.Value
    Else
        ThisWorkbook.cellWithError.Value = suggestionList.Value
    End If
    linkToCell.Visible = False
    linkToCell.Clear
End Sub

Private Sub RelativesList_Click()
    Dim selectedAddress As String
    Dim index As Integer
    
    ' Get the index of the selected item
    index = RelativesList.ListIndex + 1 ' ListIndex is zero-based, array is one-based
    
    ' Get the corresponding cell address from the RelativesListAddresses array
    If index > 0 And index <= UBound(ThisWorkbook.RelativesListAddresses) Then
        If ThisWorkbook.relativesListCells(index) Then
            ThisWorkbook.relativesListCells(index).Select
        Else
            selectedAddress = ThisWorkbook.RelativesListAddresses(index)
            ThisWorkbook.relativesListCells(index) = sheetVBA.Range(selectedAddress)
            ThisWorkbook.relativesListCells(index).Select
        End If
    Else
        MsgBox "Invalid selection."
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Set UserForm properties
    Me.BackColor = RGB(255, 255, 255) ' Set background color to white (same as Excel)
End Sub


Private Sub UserForm_Click()

End Sub

