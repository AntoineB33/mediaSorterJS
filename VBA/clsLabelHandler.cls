








Public WithEvents LabelControl As MSForms.Label


Enum Actions
    SelectAction = 1
    NewValAction = 2
    NewBgColAction = 3
End Enum


Private Sub LabelControl_Click()
    Dim LBInd As Integer
    LBInd = (CInt(Mid(LabelControl.Name, Len(ListSuggSt) + 1))) + 1
    For Each Action In listBoxList(sheetCodeName).item(LBInd)
        Select Case Action("action")
        Case Actions.SelectAction
            Action("address").Select
        Case Actions.NewValAction
            Action("address").Value = Action("newVal")
        Case Actions.NewBgColAction
            Action("address").Interior.color = Action("newVal")
            give_headerColors
        End Select
    Next Action
End Sub



