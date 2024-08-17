Dim labelTopPos As Integer
Dim labelLeftPos As Integer
Dim labelHeight As Integer
Dim labelSpacing As Integer
Dim labelWidth As Integer
Dim avgCharWidth As Integer

Dim sheetList As Object
Dim requests As Collection
Dim timerStarted As Boolean
Dim lastIsChg As Boolean
Dim sheetVBA As Worksheet
Dim sheet As Worksheet
Dim sorting As Object
Dim sheetCodeName As String
Dim targetAddress As Object
Dim listBoxList As Object
Dim newCollection As Collection


Enum Actions
    SelectAction = 1
    NewValAction = 2
    NewBgColAction = 3
End Enum


Sub LabelClickHandler(labelIndex As Integer)
    If listBoxList(sheetCodeName)(labelIndex)("sheetCodeName") <> sheetCodeName Then
        Exit Sub
    End If
    For Each Action In listBoxList(sheetCodeName)(labelIndex)
        Select Case Action("action")
        Case Actions.SelectAction
            Action("address").Select
        Case Actions.NewValAction
            Action("address").Value = Action("newVal")
        Case Actions.NewBgColAction
            Action("address").Interior.color = Action("newVal")
            CallJavaScriptFunctionAsync "chgSheet", False, "sheetCodeName", sheetCodeName
        End Select
    Next Action
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
    Dim j As Integer
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
                        ' Case "styleBorders"
                        '     styleBorders CInt(VBARequestParts(1)), CInt(VBARequestParts(2)), CInt(VBARequestParts(3)), CBool(VBARequestParts(5))
                        ' Case "chgValue"
                        '     Set cell = sheet.Cells(CInt(VBARequestParts(2)), CInt(VBARequestParts(1)))
                        '     cell.Value = VBARequestParts(3)
                        ' Case "range_updateRegularity"
                        '     Set range = sheet.range(sheet.Cells(VBARequestParts(1), 1), sheet.Cells(VBARequestParts(1), VBARequestParts(2)))
                        ' Case "color_updateRegularity"
                        '     range.Interior.color = VBARequestParts(1)
                        ' Case "clear_updateRegularity"
                        '     range.Interior.color = xlNone
                        ' Case "font_updateRegularity"
                        '     range.Font.color = VBARequestParts(1)
                        ' Case "select"
                        '     Set range = sheet.range(sheet.Cells(CInt(VBARequestParts(1)) + 1, 1), sheet.Cells(CInt(VBARequestParts(1)) + 1, 1))
                        '     range.Select
                        Case "listBoxList"
                            Dim ctrl As Control
                            Dim i As Integer
                            
                            ' Loop through all controls in the UserForm
                            For i = UserForm1.Controls.Count - 1 To 0 Step -1
                                Set ctrl = UserForm1.Controls(i)
                                
                                ' Check if the control is a label and was dynamically added
                                If TypeName(ctrl) = "Label" And ctrl.Tag = "ListSugg" Then
                                    UserForm1.Controls.Remove ctrl.Name
                                End If
                            Next i
                            
                            Dim acc As Integer
                            acc = 0
                            ' Set newCollection = New Collection
                            Set listBoxList(sheetCodeName) = New Collection
                            Dim subItem As Variant
                            For Each subItem In item(key)
                                Dim callAddress As Variant
                                Dim oneListItem As Variant
                                For Each oneListItem In subItem
                                    ' Create a new label
                                    Set ctrl = UserForm1.Controls.Add("Forms.Label.1")
                                    With ctrl
                                        .Caption = oneListItem("msg")
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
                                        .Name = "ListSugg" & acc
                                        .Tag = "ListSugg"
                                        
                                        ' Set colors for each item
                                        .BackColor = oneListItem("color")

                                        oneListItem("actions").Add "sheetCodeName", sheetCodeName

                                        ' Assign the address to the action
                                        For Each Action In oneListItem("actions")
                                            If Action("address") = 0 Then
                                                Action("address") = callAddress
                                            Else
                                                callAddress = sheet.Cells(Action("address")(0), Action("address")(1))
                                                Action("address") = callAddress
                                            End If
                                        Next Action

                                        listBoxList(sheetCodeName).Add oneListItem("actions")

                                        acc = acc + 1

                                        ' Assign the macro to the label's Click event
                                        .Object.OnAction = "'LabelClickHandler " & i & "'"
                                    End With
                                Next oneListItem
                            Next subItem
                        ' Case "sort"
                        '     Dim fullName As String
                        '     fullName = ThisWorkbook.Name
                        '     Dim suffix As String
                        '     suffix = Left(fullName, InStrRev(fullName, ".") - 1) & "\" & sheet.Name
                        '     Dim filePath As String
                        '     filePath = dataFolderPath & suffix & ".txt"
                        '     Dim fileNumber As Integer

                            
                        '     ' Check if the directory exists, if not, create it
                        '     If Dir(dataFolderName, vbDirectory) = "" Then
                        '         MkDir dataFolderName
                        '     End If

                        '     ' Open the file for writing
                        '     fileNumber = FreeFile
                        '     Open filePath For Output As fileNumber

                        '     Dim linesToWriteInFile() As String
                        '     linesToWriteInFile = Split(VBARequestParts(1), vbLf)
                        '     For Each Line In linesToWriteInFile
                        '         Print #fileNumber, Line
                        '     Next Line

                        '     ' Close the file
                        '     Close fileNumber
                            
                        '     Dim cFilePath As String
                        '     cFilePath = Root & "programs\c_prog\Project2\x64\Debug\Project2.exe " & suffix
                        '     'ExecuteCFile cFilePath
                        ' Case "sorting"
                        '     sorting(sheetCodeName) = Not sorting(sheetCodeName)
                        ' Case "Renamings"
                        '     With UserForm1.Renamings
                        '         .Caption = VBARequestParts(1)
                        '         .Visible = True
                        '     End With
                        ' Case "oldNameInput"
                        '     UserForm1.oldNameInput.Text = VBARequestParts(1)
                        ' Case "newNameInput"
                        '     UserForm1.newNameInput.Text = VBARequestParts(1)
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

' Define a custom type
Type TwoLists
    headerColors As Variant
    myArray As Variant
End Type

Private Sub give_headerColors(Optional newColCount As Long) As TwoLists
    If newColCount Is Nothing Then
        ' Determine the last column by finding the last non-empty cell in the header row
        newColCount = sheetVBA.Cells(headerRow, sheetVBA.Columns.Count).End(xlToLeft).Column
    End If
    ' get the color of each header
    Dim headerColors(1 To newColCount) As Variant
    Dim j As Long
    For j = 1 To newColCount
        headerColors(j) = sheetVBA.Cells(1, j).Interior.color
    Next j
    
    Dim result As TwoLists
    Set result.headerColors = headerColors
    Set result.myArray = myArray
    CallJavaScriptFunctionAsync "chgSheet", False, "sheetCodeName", sheetCodeName, "headerColors", headerColors
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Set sheetVBA = Sh
    sheetCodeName = Sh.codeName
    If Not sheetList.Exists(sheetCodeName) Then
        Dim result As TwoLists
        Set result.headerColors = headerColors
        Set result.myArray = myArray
        give_headerColors = result

        CallJavaScriptFunctionAsync "newSheet", False, "sheetCodeName", sheetCodeName, "values", result.myArray
        
        ' Add the data to the dictionary
        sheetList.Add sheetCodeName, Sh
        sorting.Add sheetCodeName, False
        targetAddress.Add sheetCodeName, Application.Selection.Address
        Set newCollection = New Collection
        listBoxList.Add sheetCodeName, newCollection
    End If
    CallJavaScriptFunctionAsync "selectionChange", True, "selection", targetAddress(sheetCodeName)
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As range)
    Dim changes() As Variant
    Dim cell As range
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
            cell.Borders(xlEdgeTop).color = RGB(0, 0, 0)
            cell.Borders(xlEdgeBottom).LineStyle = xlContinuous
            cell.Borders(xlEdgeBottom).color = RGB(0, 0, 0)
            cell.Borders(xlEdgeLeft).LineStyle = xlContinuous
            cell.Borders(xlEdgeLeft).color = RGB(0, 0, 0)
            cell.Borders(xlEdgeRight).LineStyle = xlContinuous
            cell.Borders(xlEdgeRight).color = RGB(0, 0, 0)
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

Private Sub Workbook_Open()
    labelTopPos = 150
    labelLeftPos = 0
    labelHeight = 18
    labelSpacing = 0
    labelWidth = 18
    avgCharWidth = 6 ' Adjust this value based on your font size and type

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
    UserForm1.Renamings.Caption = ""
    UserForm1.Show vbModeless ' Show the UserForm without blocking Excel

    Set sheetList = CreateObject("Scripting.Dictionary")
    Set sorting = CreateObject("Scripting.Dictionary")
    Set targetAddress = CreateObject("Scripting.Dictionary")
    Set listBoxList = CreateObject("Scripting.Dictionary")
    Workbook_SheetActivate activeSheet
End Sub