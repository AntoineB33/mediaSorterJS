Dim sheetList As Object
Dim requests As Collection
Dim timerStarted As Boolean
Dim lastIsChg As Boolean
Dim sheetVBA As Worksheet
Dim sheet As Worksheet
Dim sorting As Object
Dim sheetCodeName As String
Dim RelativesListAddresses As Object
Dim relativesListCells As Object
Dim targetAddress As Object

Public cellWithError As range
Public chgBckCol As Boolean



Private Sub Workbook_Open()
    Call Shell("cmd.exe /C C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\excel_prog\mediaSorter\my-js-server\startNodeJS.bat", vbHide)
    timerStarted = False
    lastIsChg = False
    Set requests = New Collection


    
    Dim ref As Object
    On Error Resume Next
    ' Check if the reference is already added
    Set ref = ThisWorkbook.VBProject.References("Scripting")
    On Error GoTo 0
    
    ' If the reference is not added, add it
    If ref Is Nothing Then
        ThisWorkbook.VBProject.References.AddFromGuid _
            GUID:="{420B2830-E718-11CF-893D-00A0C9054228}", Major:=1, Minor:=0
    End If




    UserForm1.SortButton.Caption = "Sort"
    UserForm1.updateButton.Caption = "update"
    UserForm1.ctrlZButton.Caption = "ctrlZ"
    UserForm1.ctrlYButton.Caption = "ctrlY"
    UserForm1.PowerPointButton.Caption = "PowerPoint"
    UserForm1.renamings.Caption = ""
    UserForm1.Show vbModeless ' Show the UserForm without blocking Excel

    Set sheetList = CreateObject("Scripting.Dictionary")
    Set sorting = CreateObject("Scripting.Dictionary")
    Set RelativesListAddresses = CreateObject("Scripting.Dictionary")
    Set relativesListCells = CreateObject("Scripting.Dictionary")
    Set targetAddress = CreateObject("Scripting.Dictionary")
    Workbook_SheetActivate activeSheet
End Sub

Public Sub CheckHttpResponse()
    Dim i As Long
    Dim http As Object
    Dim range As range
    Dim rangeCol As range
    Dim arr() As Variant

    For j = requests.Count To 1 Step -1
        Set http = requests(j).http
        If http.readyState = 4 Then
            If http.Status = 200 Then
                sheetCodeName = requests(j).sheetCodeName
                Set sheet = sheetList(sheetCodeName)
                Dim result As String
                result = http.responseText
                Dim json As Object
                Set json = JsonConverter.ParseJson(http.responseText)
                VBARequests = json("result")
                Dim VBARequestsArray() As String
                VBARequestsArray = Split(VBARequests, ";")
                For i = LBound(VBARequestsArray) To UBound(VBARequestsArray)
                    Dim VBARequestParts() As String
                    VBARequestParts = Split(VBARequestsArray(i), ":")
                    Select Case VBARequestParts(0)
                    Case "styleBorders"
                        styleBorders CInt(VBARequestParts(1)), CInt(VBARequestParts(2)), CInt(VBARequestParts(3)), CBool(VBARequestParts(5))
                    Case "chgValue"
                        Set cell = sheet.Cells(CInt(VBARequestParts(2)), CInt(VBARequestParts(1)))
                        cell.Value = VBARequestParts(3)
                    Case "range_updateRegularity"
                        Set range = sheet.range(sheet.Cells(VBARequestParts(1), 1), sheet.Cells(VBARequestParts(1), VBARequestParts(2)))
                    Case "color_updateRegularity"
                        range.Interior.Color = VBARequestParts(1)
                    Case "clear_updateRegularity"
                        range.Interior.Color = xlNone
                    Case "font_updateRegularity"
                        range.Font.Color = VBARequestParts(1)
                    Case "select"
                        Set range = sheet.range(sheet.Cells(CInt(VBARequestParts(1)) + 1, 1), sheet.Cells(CInt(VBARequestParts(1)) + 1, 1))
                        range.Select
                    Case "linkToCell"
                        chgBckCol = CBool(VBARequestParts(5))
                        updateSuggestions(CInt(VBARequestParts(1)), CInt(VBARequestParts(2)), CInt(VBARequestParts(4)))
                        Set cellWithError = sheet.Cells(CInt(VBARequestParts(2)), CInt(VBARequestParts(1)))
                        UserForm1.linkToCell.Visible = True
                        With UserForm1.suggestionList
                            .AddItem VBARequestParts(3)
                            .Visible = True
                        End With
                    Case "sort"
                        Dim fullName As String
                        fullName = ThisWorkbook.Name
                        Dim suffix As String
                        suffix = Left(fullName, InStrRev(fullName, ".") - 1) & "\" & sheet.Name
                        Dim filePath As String
                        filePath = dataFolderPath & suffix & ".txt"
                        Dim fileNumber As Integer

                        
                        ' Check if the directory exists, if not, create it
                        If Dir(dataFolderName, vbDirectory) = "" Then
                            MkDir dataFolderName
                        End If

                        ' Open the file for writing
                        fileNumber = FreeFile
                        Open filePath For Output As fileNumber

                        Dim linesToWriteInFile() As String
                        linesToWriteInFile = Split(VBARequestParts(1), vbLf)
                        For Each Line In linesToWriteInFile
                            Print #fileNumber, Line
                        Next Line

                        ' Close the file
                        Close fileNumber
                        
                        Dim cFilePath As String
                        cFilePath = Root & "programs\c_prog\Project2\x64\Debug\Project2.exe " & suffix
                        'ExecuteCFile cFilePath
                    Case "sorting"
                        sorting(sheetCodeName) = Not sorting(sheetCodeName)
                    Case "clearRelativesList"
                        UserForm1.RelativesList.Clear
                    Case "addItem"
                        ' Initialize the RelativesListAddresses array to the size of the data range
                        arr = relativesListCells(sheetCodeName)
                        ReDim arr(1 To UBound(RelativesListAddresses(sheetCodeName)))
                        arr = RelativesListAddresses(sheetCodeName)
                        ReDim arr(1 To UBound(RelativesListAddresses(sheetCodeName)) + 1)
                        UserForm1.RelativesList.AddItem VBARequestParts(1)
                        arr(UBound(arr)) = VBARequestParts(2)
                    Case "Renamings"
                        With UserForm1.renamings
                            .Caption = VBARequestParts(1)
                            .Visible = True
                        End With
                    Case "oldNameInput"
                        UserForm1.oldNameInput.Text = VBARequestParts(1)
                    Case "newNameInput"
                        UserForm1.newNameInput.Text = VBARequestParts(1)
                    End Select
                Next i
                requests.Remove j
            Else
                MsgBox "Error: " & http.statusText
                requests.Remove j
            End If
        End If
    Next j

    If requests.Count > 0 Then
        ' Reschedule the check if there are still pending requests
        Application.OnTime Now + TimeValue("00:00:01"), "ThisWorkbook.CheckHttpResponse"
    Else
        timerStarted = False
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As range)
    CallJavaScriptFunctionAsync "selectionChange", True, "selectedCells", Target.Address
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    sheetCodeName = Sh.codeName
    If Not sheetList.Exists(sheetCodeName) Then
        ' Get the used range
        Dim usedRange As range
        Dim myArray As Variant
        Set usedRange = Sh.usedRange
        myArray = usedRange.Value

        ' get the color of each header
        Dim headerColors As Variant
        Dim newColCount As Long
        Dim j As Long
        ReDim headerColors(1 To newColCount)
        For j = 1 To newColCount
            headerColors(j) = sheet.Cells(1, j).Interior.Color
        Next j

        CallJavaScriptFunctionAsync "newSheet", False, "values", myArray
        
        ' Add the data to the dictionary
        sheetList.Add sheetCodeName, Sh
        sorting.Add sheetCodeName, False
        Dim emptyList() As Variant
        RelativesListAddresses.Add sheetCodeName, emptyList
        relativesListCells.Add sheetCodeName, emptyList
        targetAddress.Add sheetCodeName, Application.Selection.Address
    End If
    CallJavaScriptFunctionAsync "chgSheet", False, "sheetCodeName", sheetCodeName
    CallJavaScriptFunctionAsync "selectionChange", True, "selection", targetAddress(sheetCodeName), "headerColors", headerColors
    Set sheetVBA = Sh
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim changes() As Variant
    Dim cell As Range
    Dim i As Long

    ' Initialize the dynamic array with the size of the Target range
    ReDim changes(1 To Target.Cells.Count, 1 To 3)

    ' Iterate through each cell in the Target range
    i = 1
    For Each cell In Target
        changes(i, 1) = cell.Row
        changes(i, 2) = cell.Column
        changes(i, 3) = cell.Value
        i = i + 1
    Next cell

    CallJavaScriptFunctionAsync "handleChange", True, "changes", changes
End Sub

Sub styleBorders(i As Integer, jL As Integer, jU As Integer, draw As Boolean)
    Dim cell As range
    Dim j As Integer
    For j = jL To jU
        Set cell = sheet.Cells(i + 1, j + 1)
        If draw Then
            cell.Borders(xlEdgeTop).LineStyle = xlContinuous
            cell.Borders(xlEdgeTop).Color = RGB(0, 0, 0)
            cell.Borders(xlEdgeBottom).LineStyle = xlContinuous
            cell.Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
            cell.Borders(xlEdgeLeft).LineStyle = xlContinuous
            cell.Borders(xlEdgeLeft).Color = RGB(0, 0, 0)
            cell.Borders(xlEdgeRight).LineStyle = xlContinuous
            cell.Borders(xlEdgeRight).Color = RGB(0, 0, 0)
        Else
            cell.Borders(xlEdgeTop).LineStyle = xlNone
            cell.Borders(xlEdgeBottom).LineStyle = xlNone
            cell.Borders(xlEdgeLeft).LineStyle = xlNone
            cell.Borders(xlEdgeRight).LineStyle = xlNone
        End If
    Next j
End Sub

Sub updateSuggestions(i As Integer, j As Integer, suggestions As Variant)
    Dim suggestionList As Object
    Set suggestionList = UserForm1.suggestionList

    Dim suggestionText As Variant
    For Each suggestionText In suggestions
        Dim listItem As Object
        Set listItem = UserForm1.Controls.Add("Forms.Listbox.1", , True)
        listItem.AddItem suggestionText
        listItem.Tag = suggestionText
        listItem.Width = 200
        listItem.Height = 20
        listItem.Left = 10
        listItem.Top = 10 + 20 * (suggestionList.ListCount - 1)
        listItem.Visible = True
        suggestionList.AddItem listItem
    Next suggestionText

    suggestionList.Visible = (suggestionList.ListCount <> 0)
    UserForm1.linkToCell.Visible = (suggestionList.ListCount <> 0)
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Save
    ' Stop the Spring Boot application
    Call Shell("cmd.exe /C C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\excel_prog\mediaSorter\my-js-server\stopNodeJS.bat", vbHide)
End Sub

Public Sub CallJavaScriptFunctionAsync(ByVal funcName As String, ParamArray params() As Variant)
    Dim i As Integer, j As Integer
    On Error GoTo ErrorHandler
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Create a new RequestItem
    Dim requestItem As requestItem
    Set requestItem = New requestItem
    Set requestItem.http = http
    requestItem.funcName = funcName
    requestItem.sheetCodeName = sheetCodeName
    
    ' Convert ParamArray to regular array
    If UBound(params) > -1 Then
        Dim paramsArray() As Variant
        ReDim paramsArray(LBound(params) To UBound(params) - 1)
        For i = LBound(params) + 1 To UBound(params)
            paramsArray(i - 1) = params(i)
        Next i
    End If
        
    
    

    
    ' Create JSON request body
    Dim jsonRequest As String
    If UBound(params) > 0 Then
        Dim json As Object
        Set json = CreateObject("Scripting.Dictionary")
        json.Add "functionName", funcName
        For i = LBound(paramsArray) To UBound(paramsArray) Step 2
            json.Add paramsArray(i), paramsArray(i + 1)
        Next i
        jsonRequest = JsonConverter.ConvertToJson(json)
    Else
        jsonRequest = "{""functionName"": """ & funcName & """}"
    End If
    
    If params(LBound(params)) Then
        requests.Add requestItem
    End If

    Dim url As String
    url = "http://localhost:3000/execute"
    ' Open the request
    http.Open "POST", url, True
    
    ' Set the request headers
    http.setRequestHeader "Content-Type", "application/json"
    
    ' Send the request with the function name and parameters
    http.send jsonRequest
    Debug.Print "Invoke-RestMethod -Uri http://localhost:3000/execute -Method Post -Headers @{ ""Content-Type"" = ""application/json"" } -Body '" & jsonRequest & "'"
    
    ' Start the timer to check the response periodically
    If params(LBound(params)) And Not timerStarted Then
        CheckHttpResponse
        timerStarted = True
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

Private Sub sendHTTP()

End Sub











