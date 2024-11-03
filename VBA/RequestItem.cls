








' Inside the RequestItem class module
Public http As Object
Public funcName As String
Public sheetCodeName As String
Public oneAnswer As Boolean
Private params() As Variant

Public Sub setParams(newParams() As Variant)
    params = newParams
End Sub
