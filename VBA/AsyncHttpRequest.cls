








Private Sub Class_Initialize()
    Set http = New WinHttp.WinHttpRequest
End Sub

' Public Sub SendAsyncRequest(URL As String)
'     With http
'         .Open "GET", URL, True
'         .send
'     End With
' End Sub

Public Sub SendAsyncRequest(url As String, jsonData As String)
    With http
        .Open "POST", url, True
        .setRequestHeader "Content-Type", "application/json"
        .send jsonData
    End With
End Sub

Private Sub http_OnResponseFinished()
    MsgBox http.responseText
    response = http.responseText
    Set responseJson = JsonConverter.ParseJson(response)
    
    ' Iterate through the list of functions to call
    For Each functionCall In responseJson
        functionName = functionCall("functionName")
        Parameters = functionCall("parameters")
        
        ' Call the VBA function with the parameters
        Select Case UBound(Parameters)
            Case -1
                Application.Run functionName
            Case 0
                Application.Run functionName, Parameters(0)
            Case 1
                Application.Run functionName, Parameters(0), Parameters(1)
            Case 2
                Application.Run functionName, Parameters(0), Parameters(1), Parameters(2)
            ' Add more cases as needed for more parameters
            Case Else
                MsgBox "Too many parameters for function: " & functionName
        End Select
    Next functionCall
End Sub

Private Sub http_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    Application.OnTime Now + TimeValue("0:00:01"), "ThisWorkbook.TestAsyncHTTPRequest"
    'MsgBox "Error: " & ErrorNumber & " - " & ErrorDescription
End Sub



