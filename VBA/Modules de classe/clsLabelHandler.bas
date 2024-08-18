Public WithEvents LabelControl As MSForms.Label


Enum Actions
    SelectAction = 1
    NewValAction = 2
    NewBgColAction = 3
End Enum


Private Sub LabelControl_Click()
    With ThisWorkbook.listBoxList(ThisWorkbook.sheetCodeName)
        Dim LBInd As Integer
        LBInd = (CInt(Mid(LabelControl.Name, Len(ThisWorkbook.ListSuggSt) + 1)))
        If LBInd > .ListCount - 1 Then
            Exit Sub
        End If
        For Each Action In .item(LBInd)
            Select Case Action("action")
            Case Actions.SelectAction
                Action("address").Select
            Case Actions.NewValAction
                Action("address").Value = Action("newVal")
            Case Actions.NewBgColAction
                Action("address").Interior.color = Action("newVal")
                ThisWorkbook.give_headerColors
            End Select
        Next Action
    End With
End Sub

