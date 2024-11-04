'xlwings vba edit


Private Sub Workbook_Open()
    Dim result As Variant
    result = GetRowAndColumn(activeCell.Address(False, False))
    editRow = result(0)
    spacing = 10
    leftSpacing = 10
    thbnlDisplayed = False
    labelTopPos = 150
    labelLeftPos = 0
    labelHeight = 18
    labelSpacing = 0
    labelWidth = 114
    avgCharWidth = 6 ' Adjust this value based on your font size and type


    RunCmdAndWait
    ' Call Shell("cmd.exe /C C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\excel_prog\mediaSorter\startNodeJS.bat", vbHide)
    Call Shell("cmd.exe /C C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\excel_prog\mediaSorter\startNodeJS.bat")
    
    Dim check As String
    check = SendRequestToNodeJSEndpoint
    If check = "Request timed out." Then
        MsgBox "The server is not responding. Please make sure the Node.js server is running."
        Exit Sub
    End If
    check = SendRequestToNodeJSEndpoint2
    If check = "Request timed out." Then
        MsgBox "The server is not responding. Please make sure the Node.js server is running."
        Exit Sub
    End If
    
    ' Dim isServerUp As Boolean
    ' Dim serverStarted As Boolean
    ' Dim waitTime As Long
    
    ' ' Set your server's address and port here
    ' waitTime = 1000 ' 1 second
    ' ' Wait for the server to start by checking if the port is open
    ' serverStarted = False
    
    ' Do While Not serverStarted
    '     isServerUp = CheckIfServerIsUp
        
    '     If isServerUp Then
    '         serverStarted = True
    '     Else
    '         Debug.Print "issues when starting the server..."
    '         Sleep waitTime ' Wait for 1 second before trying again
    '     End If
        
    '     DoEvents ' Keep the VBA application responsive during the wait
    ' Loop


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

    UserForm1.Show vbModeless ' Show the UserForm without blocking Excel

    Set sheetList = CreateObject("Scripting.Dictionary")
    Set sorting = CreateObject("Scripting.Dictionary")
    Set fstSort = CreateObject("Scripting.Dictionary")
    Set originalDatas = CreateObject("Scripting.Dictionary")
    Set targetAddress = CreateObject("Scripting.Dictionary")
    Set listBoxList = CreateObject("Scripting.Dictionary")
    ListSuggSt = "ListSugg"
    logFile = "C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\excel_prog\mediaSorter\log.txt"
    ClearTextFile logFile
    Workbook_SheetActivate activeSheet
End Sub

Sub RunCmdAndWait()
    Dim wsh As Object
    Dim exitCode As Integer
    Dim cmd As String

    ' Create WScript.Shell object
    Set wsh = CreateObject("WScript.Shell")

    ' Specify the command (path to your batch file)
    cmd = "C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\excel_prog\mediaSorter\stopNodeJS.bat"

    ' Run the command and wait for it to complete
    exitCode = wsh.Run("cmd.exe /C """ & cmd & """", 0, True) ' vbHide is 0
End Sub

Public Sub give_headerColors(Optional newColCount As Long)
    If newColCount = 0 Then
        ' Determine the last column by finding the last non-empty cell in the header row
        newColCount = sheetVBA.Cells(1, sheetVBA.Columns.Count).End(xlToLeft).Column
    End If
    ' get the color of each header
    Dim headerColors() As Variant
    ReDim headerColors(1 To newColCount)
    Dim j As Long
    For j = 1 To newColCount
        headerColors(j) = sheetVBA.Cells(1, j).Interior.color
    Next j
    CallJavaScriptFunctionAsync "chgSheet", True, "sheetCodeName", sheetCodeName, "headerColors", headerColors
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Set sheetVBA = Sh
    sheetCodeName = ThisWorkbook.CodeName + "/" + Sh.codeName
    If Not sheetList.Exists(sheetCodeName) Then
        Dim usedRange As range
        Dim myArray As Variant
        Set usedRange = Sh.usedRange
        myArray = usedRange.Value

        CallJavaScriptFunctionAsync "newSheet", False, "sheetCodeName", sheetCodeName, "values", myArray
        
        ' Add the data to the dictionary
        sheetList.Add sheetCodeName, Sh
        sorting.Add sheetCodeName, False
        fstSort.Add sheetCodeName, True
        originalDatas.Add sheetCodeName, Null
        targetAddress.Add sheetCodeName, Application.Selection.Address
        listBoxList.Add sheetCodeName, Null
        give_headerColors UBound(myArray, 2)
    Else
        give_headerColors
    End If
    Dim result As Variant
    result = GetRowAndColumn(activeCell.Address(False, False))
    editRow = result(0)
    editColumn = result(1)
    CallJavaScriptFunctionAsync "selectionChange", True, "editRow", editRow, "editCol", editCol
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
        Dim result As Variant
        result = GetRowAndColumn(cell.Address(False, False))
        changes(i, 1) = result(0)
        changes(i, 2) = result(1)
        changes(i, 3) = CStr(cell.Value)
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

Sub ClearTextFile(filePath As String)
    ' Open the file for output, which will clear its contents
    Open filePath For Output As #1
    
    ' Immediately close the file, leaving it empty
    Close #1
End Sub

' Helper function to check if a file exists
Function fileExists(filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileExists = fso.fileExists(filePath)
End Function

' Helper function to check if a line exists in a file
Function LineExistsInFile(filePath As String, searchString As String) As Boolean
    Dim fso As Object
    Dim file As Object
    Dim line As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.fileExists(filePath) Then
        Set file = fso.OpenTextFile(filePath, 1)
        
        Do While Not file.AtEndOfStream
            line = file.ReadLine
            If InStr(line, searchString) > 0 Then
                LineExistsInFile = True
                file.Close
                Exit Function
            End If
        Loop
        
        file.Close
    End If
    
    LineExistsInFile = False
End Function

Sub RunAndWait(command As String)
    Dim wsh As Object
    Dim process As Object
    
    ' Create WScript.Shell object
    Set wsh = VBA.CreateObject("WScript.Shell")
    
    ' Run the command and wait for it to complete
    Set process = wsh.Exec(command)
    
    ' Wait for process to complete
    Do While process.Status = 0
        Sleep 1
    Loop
End Sub

Function SendRequestToNodeJSEndpoint()
    Dim xmlhttp As Object
    Dim url As String
    Dim startTime As Double
    Dim timeOutSeconds As Double
    Dim responseReceived As Boolean
    Dim res As String
    
    ' Define your endpoint URL
    url = "http://localhost:3000/health"
    
    ' Define the timeout duration (in seconds)
    timeOutSeconds = 5 ' 30 seconds timeout
    
    ' Record the start time
    startTime = Timer
    
    ' Initialize the response flag
    responseReceived = False
    
    Do
        ' Create an XMLHTTP object
        Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
        
        ' Send the HTTP GET request to the Node.js server
        xmlhttp.Open "GET", url, False
        xmlhttp.Send
        
        ' Check if the request was successful
        If xmlhttp.Status = 200 Then
            responseReceived = True
            res = "Response received: " & xmlhttp.responseText
            Exit Do
        End If
        
        ' Check if the timeout has been reached
        If Timer - startTime > timeOutSeconds Then
            res = "Request timed out."
            Exit Do
        End If
        
        ' Wait for a short interval before retrying
        Application.Wait (Now + TimeValue("00:00:01"))
        
    Loop Until responseReceived
    
    ' Clean up
    Set xmlhttp = Nothing
    SendRequestToNodeJSEndpoint = res
End Function

Function SendRequestToNodeJSEndpoint2()
    Dim xmlhttp As Object
    Dim url As String
    Dim startTime As Double
    Dim timeOutSeconds As Double
    Dim responseReceived As Boolean
    Dim res As String
    
    ' Define your endpoint URL
    url = "http://localhost:3000/test"
    
    ' Define the timeout duration (in seconds)
    timeOutSeconds = 5 ' 30 seconds timeout
    
    ' Record the start time
    startTime = Timer
    
    ' Initialize the response flag
    responseReceived = False
    
    Do
        ' Create an XMLHTTP object
        Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
        
        ' Send the HTTP GET request to the Node.js server
        xmlhttp.Open "POST", url, False
        
        ' Set the request headers
        xmlhttp.setRequestHeader "Content-Type", "application/json"
        
        xmlhttp.Send
        
        ' Check if the request was successful
        If xmlhttp.Status = 200 Then
            responseReceived = True
            res = "Response received: " & xmlhttp.responseText
            Exit Do
        End If
        
        ' Check if the timeout has been reached
        If Timer - startTime > timeOutSeconds Then
            res = "Request timed out."
            Exit Do
        End If
        
        ' Wait for a short interval before retrying
        Application.Wait (Now + TimeValue("00:00:01"))
        
    Loop Until responseReceived
    
    ' Clean up
    Set xmlhttp = Nothing
    SendRequestToNodeJSEndpoint2 = res
End Function