








Public WithEvents ScriptControl As MSHTML.HTMLWindow2


Private Sub ScriptControl_OnFunctionComplete(result As String)
    Debug.Print "Function completed with result: " & result
    ' Handle the result as needed
End Sub

