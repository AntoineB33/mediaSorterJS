Dim msgTypeColors(3) As Integer
msgTypeColors(0) = RGB(255, 0, 0)
msgTypeColors(1) = RGB(255, 137, 0)
msgTypeColors(2) = RGB(0, 145, 255)
Dim labelTopPos As Integer
labelTopPos = 150
Dim labelLeftPos As Integer
labelLeftPos = 0
Dim labelHeight As Integer
labelHeight = 18
Dim labelSpacing As Integer
labelSpacing = 0
Dim labelWidth As Integer
labelWidth = 18
Dim avgCharWidth As Integer
avgCharWidth = 6 ' Adjust this value based on your font size and type

Dim sheetList As Object
Dim requests As Collection
Dim timerStarted As Boolean
Dim lastIsChg As Boolean
Dim sheetVBA As Worksheet
Dim sheet As Worksheet
Dim sorting As Object
Dim sheetCodeName As String
Dim targetAddress As Object

Public listBoxList As Object


Enum Actions
    Select = 1
    NewVal = 2
    NewBgCol = 3
End Enum

Enum msgType
    ERROR = 0
    WARNING = 1
    SUGGESTION = 2
    RELATIVES = 3
End Enum


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
    Set targetAddress = CreateObject("Scripting.Dictionary")
    Set listBoxList = CreateObject("Scripting.Dictionary")
    Workbook_SheetActivate activeSheet
End Sub

Public Sub CheckHttpResponse()
    Dim http As Object
    Dim jsonResponse As String
    Dim jsonArray As Object
    Dim item As Variant
    Dim key As Variant
    Dim range As range
    Dim rangeCol As range
    Dim arr() As Variant
    Dim currentMsgType As Integer
    Dim color As Long

    For j = requests.Count To 1 Step -1
        Set http = requests(j).http
        If http.readyState = 4 Then
            If http.Status = 200 Then
                sheetCodeName = requests(j).sheetCodeName
                Set sheet = sheetList(sheetCodeName)
                jsonResponse = http.responseText
                Set jsonArray = JsonConverter.ParseJson(jsonResponse)
                For Each item In jsonArray
                    For Each key In item.Keys
                        Select Case key
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
                        Case "listBoxList"
                            For Each subItem In item(key)
                                If item(key)(0) <> msgType.SUGGESTION Then
                                    Dim r As Integer
                                    Dim g As Integer
                                    Dim b As Integer
                                    Dim newR As Integer
                                    Dim newG As Integer
                                    Dim newB As Integer
                                    r = color Mod 256
                                    g = (color \ 256) Mod 256
                                    b = (color \ 65536) Mod 256
                                    newR = (r + 255) \ 2
                                    newG = (g + 255) \ 2
                                    newB = (b + 255) \ 2
                                    color = RGB(newR, newG, newB)
                                Else
                                    currentMsgType = item(key)(0)
                                    color = msgTypeColors(currentMsgType)
                                End If

                                Dim acc As Integer
                                acc = 0
                                Dim i As Integer
                                For i = 0 To currentMsgType - 1 Then
                                    acc = acc + UBound(suggItem)
                                Next suggItem
                                
                                ' Create a new label
                                Set newLabel = UserForm1.Controls.Add("Forms.Label.1")
                                With newLabel
                                    .Caption = item(key)(1)
                                    .Left = labelLeftPos
                                    .Width = labelWidth
                                    .WordWrap = True ' Enable word wrap

                                    ' Calculate the number of lines in the label
                                    Dim labelLines As Integer
                                    Dim textLength As Integer
                                    Dim lineWidth As Integer
                                    
                                    ' Estimate the average character width (this may need adjustment based on font)
                                    lineWidth = labelWidth / avgCharWidth

                                    ' Calculate the number of lines in the label
                                    Dim pos As Integer
                                    Dim pos0 As Integer
                                    pos0 = 1
                                    pos = InStr(1, item(key)(1), vbLf)
                                    If pos = 0 Then
                                        pos = Len(item(key)(1)) + 1
                                    End If
                                    labelLines = 0
                                    Do While pos > 0
                                        ' Calculate the number of lines based on text length and line width
                                        labelLines = labelLines + Int((pos - pos0) / lineWidth) + 1

                                        pos = InStr(pos + 1, item(key)(1), vbLf)
                                    Loop
                                    
                                    ' Adjust the height based on the number of lines
                                    .Height = labelHeight * labelLines
                                    
                                    ' Assign a unique name or tag
                                    .Name = "ListSugg" & currentMsgType & "-" & UBound(listBoxList(sheetCodeName)(currentMsgType))
                                    .Tag = "ListSugg" & currentMsgType
                                    
                                    ' Set colors for each item
                                    .BackColor = color
                                End With
                                
                                For i = UserForm1.Controls.Count - 1 To 0 Step -1
                                    Set ctrl = UserForm1.Controls(i)
                                    If TypeName(ctrl) = "Label" And ctrl.Tag = "Dynamic" Then
                                        Dim x As Integer
                                        Dim xStr As String
                                        Dim hyphenPos As Integer
                                        ' Check if the name starts with the prefix
                                        If Left(ctrl.Name, Len("ListSugg")) = "ListSugg" Then
                                            ' Extract the part after the prefix
                                            Dim suffix As String
                                            suffix = Mid(ctrl.Name, Len("ListSugg") + 1)
                                            ' Find the position of the hyphen
                                            hyphenPos = InStr(suffix, "-")
                                            ' Check if hyphen exists
                                            If hyphenPos > 0 Then
                                                ' Extract the numeric part (x) and the remaining part (y)
                                                xStr = Left(suffix, hyphenPos - 1)
                                                ' Check if the numeric part is actually a number
                                                If IsNumeric(xStr) Then
                                                    x = CInt(xStr)
                                                    ' Compare the numeric part with currentMsgType
                                                    If x > currentMsgType Then
                                                        ctrl.Top = ctrl.Top + labelHeight + labelSpacing
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next i
                            Next subItem
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
                    Next key
                Next item
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
        Dim newCollection As Collection
        Collection = New Collection
        listBoxList.Add sheetCodeName, Collection
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











