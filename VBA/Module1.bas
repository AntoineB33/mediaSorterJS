Attribute VB_Name = "Module1"




Function GetRowAndColumn(cellAddress As String) As Variant
    Dim columnLetters As String
    Dim rowNumber As Long
    Dim i As Long

    ' Separate column letters and row numbers from cellAddress
    For i = 1 To Len(cellAddress)
        If IsNumeric(Mid(cellAddress, i, 1)) Then
            rowNumber = CInt(Mid(cellAddress, i)) ' Extract row number directly
            Exit For
        Else
            columnLetters = columnLetters & Mid(cellAddress, i, 1)
        End If
    Next i

    ' Convert column letters to column number
    Dim colNumber As Long
    colNumber = ColumnNumberFromLetters(columnLetters)
    
    ' Return both row and column numbers as an array
    GetRowAndColumn = Array(rowNumber, colNumber)
End Function

' Helper function to convert column letters to column number
Function ColumnNumberFromLetters(columnLetters As String) As Long
    Dim i As Long, colNum As Long
    colNum = 0
    For i = 1 To Len(columnLetters)
        colNum = colNum * 26 + (Asc(UCase(Mid(columnLetters, i, 1))) - Asc("A") + 1)
    Next i
    ColumnNumberFromLetters = colNum
End Function

Sub DisplayImagesInUserForm()
    'DEBUG
    ' editRow = activeCell.Row
    ' spacing = 10
    ' leftSpacing = 10

    thbnlDisplayed = True
    Dim ws As Worksheet
    Dim imgControl As MSForms.Image
    Dim lblControl As MSForms.Label
    Dim FileNameWithoutExt As String
    Dim FilePath As String
    Dim foundImage As Boolean
    Dim FolderPath As String

    ' Assuming the worksheet is active
    Set ws = ActiveSheet

    ' Specify the folder path to search
    FolderPath = "C:\Users\abarb\Documents\health\news_underground\mediaSorter\media"

    ' Reference the existing UserForm named MyForm
    Dim frm As Object
    Set frm = ImageListForm ' Use the name of your pre-created UserForm

    ' Set UserForm size to fit all images
    Dim thumbnailHeight As Single
    thumbnailHeight = frm.Width - 10 - 2*leftSpacing
    frm.Show vbModeless
    
    Dim startNum As Integer
    Dim minInt As Integer

    Dim selectedThumbnail As Boolean
    selectedThumbnail = True
    
    ' Initialize values
    startNum = editRow
    minInt = 2
    Dim maxInt As Integer
    maxInt = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If startNum < minInt Then
        startNum = minInt
    ElseIf startNum > maxInt Then
        startNum = maxInt
    End If
    

    ' Variables for controlling the pattern
    Dim decrease As Boolean
    decrease = True
    Dim formHeight As Single
    formHeight = frm.Height
    Dim maxCount As Integer
    maxCount = WorksheetFunction.Ceiling((formHeight - thumbnailHeight - 2 * spacing) / ((thumbnailHeight + spacing) * 2), 1)
    Dim selectThbnailTop As Single
    Dim currentRow As Integer
    currentRow = startNum
    Dim dist As Integer
    dist = 1
    Dim descFactor As Integer
    descFactor = -1
    Do While True
        Dim topVal As Single
        Dim shape As Object
        If selectedThumbnail Then
            If currentRow - 2 < maxCount Then
                topVal = (currentRow - 2) * (thumbnailHeight + spacing)
            ElseIf maxInt - currentRow < maxCount Then
                topVal = formHeight - (maxInt - currentRow) * (thumbnailHeight + spacing) - thumbnailHeight
            Else
                topVal = formHeight / 2 - thumbnailHeight / 2
            End If
            Set shape = frm.Controls.Add("Forms.Image.1")
            With shape
                .Width = leftSpacing
                .Height = thumbnailHeight
                .Left = leftSpacing + thumbnailHeight
                .Top = topVal
                .BackColor = RGB(0, 0, 0)
            End With
            selectThbnailTop = topVal
            selectedThumbnail = False
        Else
            currentRow = startNum + descFactor * dist
            If currentRow < minInt Then
                descFactor = descFactor * -1
                currentRow = startNum + descFactor * dist
            ElseIf currentRow > maxInt Then
                descFactor = descFactor * -1
                dist = dist + 1
                currentRow = startNum + descFactor * dist
            End If
            topVal = selectThbnailTop + descFactor * dist * (thumbnailHeight + spacing)
            If topVal + thumbnailHeight < 0 Or topVal > formHeight Then
                Exit Do
            End If
            descFactor = descFactor * -1
            If descFactor = -1 Then
                dist = dist + 1
            End If
        End If
        ' Get the filename without extension from the 6th column (e.g., column F)
        FileNameWithoutExt = ws.Cells(currentRow, 6).Value
        ' Check if the file exists
        FilePath = SearchFileInFolder(FolderPath, FileNameWithoutExt)


        If FilePath <> "" Then
            Set shape = frm.Controls.Add("Forms.Image.1")
            With shape
                .Width = leftSpacing
                .Height = thumbnailHeight
                .Left = 0
                .Top = topVal
            End With
            Set imgControl = frm.Controls.Add("Forms.Image.1")
            With imgControl
                ' Check if the file is an image
                foundImage = False
                If IsImageFile(FilePath) Then
                    shape.BackColor = RGB(255, 0, 0)
                    Set .Picture = LoadPicture(FilePath)
                ElseIf IsVideoFile(FilePath) Then
                    shape.BackColor = RGB(0, 0, 255)
                    ' If it's a video, extract a frame
                    Set .Picture = ExtractFrameFromVideo(FilePath, FileNameWithoutExt)
                End If
                .Left = leftSpacing
                .Top = topVal
                .PictureSizeMode = fmPictureSizeModeZoom

                Dim imgWidth As Long
                Dim imgHeight As Long
                Dim aspectRatio As Double

                ' Get the original image dimensions
                imgWidth = .Picture.Width / 26.458 ' Convert from twips to points
                imgHeight = .Picture.Height / 26.458 ' Convert from twips to points

                ' Calculate aspect ratio
                aspectRatio = imgWidth / imgHeight

                ' Resize the image proportionally
                If imgWidth > imgHeight Then
                    .Width = thumbnailHeight
                    .Height = thumbnailHeight / aspectRatio
                Else
                    .Height = thumbnailHeight
                    .Width = thumbnailHeight * aspectRatio
                End If
                ' .Width = thumbnailHeight - leftSpacing
            End With
        Else
            ' If the image is not found, add a Label control instead
            Set lblControl = frm.Controls.Add("Forms.Label.1")
            With lblControl
                .Caption = "Image not found: " & FileNameWithoutExt
                .Width = thumbnailHeight
                .Height = thumbnailHeight
                .Left = leftSpacing
                .Top = topVal
                .ForeColor = RGB(0, 255, 0)
            End With
        End If
    Loop
End Sub

Sub RemoveAllControls(userForm As Object)
    Dim ctrl As Control
    ' Loop backwards to safely remove all controls
    For i = userForm.Controls.Count - 1 To 0 Step -1
        userForm.Controls.Remove userForm.Controls(i).Name
    Next i
End Sub

' Function to extract a frame from a video at 10% of its duration and return as StdPicture
Function ExtractFrameFromVideo(videoPath As String, FileNameWithoutExt As String) As StdPicture
    Dim ShellApp As Object
    Dim FFmpegCommand As String
    Dim stdPicture As StdPicture
    Dim tempImagePath As String
    Dim ffmpegPath As String
    Dim cmd As String
    Dim shellObj As Object
    Dim duration As String
    Dim totalSeconds As Double
    Dim time10Percent As Double
    Dim time10PercentStr As String
    
    ' Set the path to your ffmpeg executable
    ffmpegPath = "C:\ProgramData\chocolatey\bin\ffmpeg.exe" ' Replace this with the actual path to ffmpeg.exe

    ' Set the path for the temporary image
    tempImagePath = "C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\data\video_thumbnails\" & FileNameWithoutExt & ".jpg"
    
    ' check if tempFilePath exists
    If Dir(tempImagePath) <> "" Then
        ' If the image already exists, load it and return
        Set ExtractFrameFromVideo = LoadPicture(tempImagePath)
        Exit Function
    End If
    
    ' Step 1: Create a Shell object
    Set shellObj = CreateObject("WScript.Shell")




    Dim tempFile As String
    Dim fileNum As Integer
    
    ' Define a temporary file to store FFmpeg output
    tempFile = Environ("Temp") & "\ffmpeg_output.txt"

    cmd = """" & ffmpegPath & """ -i """ & videoPath & """ ""C:\Users\abarb\Documents\health\news_underground\mediaSorter\programs\output.txt"""


    Dim execObj As Object
    Dim outputLine As String
    Dim pos As Integer

    ' Create the shell object
    Set shellObj = CreateObject("WScript.Shell")

    cmd = """" & ffmpegPath & """ -i """ & videoPath & """ 2>&1"

    ' Run the command and capture the Exec object
    Set execObj = shellObj.Exec(cmd)



    ' Loop through the StdErr stream to find the "Duration" line
    Do While Not execObj.StdErr.AtEndOfStream
        outputLine = execObj.StdErr.ReadLine
        If InStr(outputLine, "Duration") > 0 Then
            ' Convert the extracted duration to seconds
            totalSeconds = ExtractDurationInSeconds(outputLine)
            Exit Do
        End If
    Loop
    
    ' Step 3: Calculate 10% of the duration
    time10Percent = totalSeconds * 0.1
    
    ' Convert 10% time to hh:mm:ss.mmm format
    time10PercentStr = Replace(ConvertSecondsToTimeFormat(time10Percent), ",", ".")
    
    ' Step 4: Construct the FFmpeg command to extract a frame at 10% of the duration
    FFmpegCommand = """" & ffmpegPath & """ -ss " & time10PercentStr & " -i """ & videoPath & """ -vframes 1 """ & tempImagePath & """"
    


    ' Print FFmpeg command for debugging
    Debug.Print FFmpegCommand

    ' Step 5: Run the FFmpeg command to extract the frame
    ' Set execObj = shellObj.Exec(FFmpegCommand)
    
    
    ' ' Step 5: Run the FFmpeg command
    shellObj.Run FFmpegCommand, 0, True
    
    ' Step 6: Load the image into a StdPicture object
    Set stdPicture = LoadPicture(tempImagePath)
    
    ' Return the StdPicture object
    Set ExtractFrameFromVideo = stdPicture
End Function

' Function to get the file name without extension
Function GetFileNameWithoutExtension(FilePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameWithoutExtension = fso.GetBaseName(FilePath)
End Function

' Function to recursively search for a file in a folder and its subfolders
Function SearchFileInFolder(FolderPath As String, FileNameWithoutExt As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim BaseName As String
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(FolderPath)
    
    ' Search in the current folder
    For Each file In folder.Files
        BaseName = GetFileNameWithoutExtension(file.Path) ' Get the base name without extension
        If StrComp(BaseName, FileNameWithoutExt, vbTextCompare) = 0 Then
            SearchFileInFolder = file.Path
            Exit Function
        End If
    Next file
    
    ' Search in subfolders recursively
    For Each subFolder In folder.SubFolders
        SearchFileInFolder = SearchFileInFolder(subFolder.Path, FileNameWithoutExt)
        If Len(SearchFileInFolder) > 0 Then Exit Function
    Next subFolder
    
    ' If not found, return empty string
    SearchFileInFolder = ""
End Function

' Function to check if a file is an image
Function IsImageFile(FilePath As String) As Boolean
    Dim ext As String
    ext = LCase(Right(FilePath, Len(FilePath) - InStrRev(FilePath, ".")))
    If ext = "jpg" Or ext = "jpeg" Or ext = "png" Or ext = "bmp" Or ext = "gif" Then
        IsImageFile = True
    Else
        IsImageFile = False
    End If
End Function

' Function to check if a file is a video
Function IsVideoFile(FilePath As String) As Boolean
    Dim ext As String
    ext = LCase(Right(FilePath, Len(FilePath) - InStrRev(FilePath, ".")))
    If ext = "mp4" Or ext = "avi" Or ext = "mov" Or ext = "mkv" Then
        IsVideoFile = True
    Else
        IsVideoFile = False
    End If
End Function

' Function to extract duration in seconds from the ffmpeg output string
Function ExtractDurationInSeconds(durationStr As String) As Double
    Dim hours As Double
    Dim minutes As Double
    Dim seconds As Double
    Dim secondsParts() As String
    Dim timeParts() As String
    Dim timeStr As String

    ' Extract time part from the string
    timeStr = Mid(durationStr, InStr(durationStr, "Duration:") + 10, 11)
    
    ' Split the time string into hours, minutes, and seconds
    timeParts = Split(timeStr, ":")
    
    hours = CDbl(timeParts(0))
    minutes = CDbl(timeParts(1))

    ' Now, split the seconds part by the decimal point
    secondsParts = Split(timeParts(2), ".")
    seconds = CDbl(secondsParts(0)) ' Get the whole seconds part
    
    ' Convert to total seconds
    ExtractDurationInSeconds = (hours * 3600) + (minutes * 60) + seconds
End Function

' Function to convert seconds to hh:mm:ss.mmm format for ffmpeg
Function ConvertSecondsToTimeFormat(seconds As Double) As String
    Dim hh As Long
    Dim mm As Long
    Dim ss As Double
    Dim timeStr As String

    ' Convert to hours, minutes, and seconds
    hh = Int(seconds / 3600)
    mm = Int((seconds - (hh * 3600)) / 60)
    ss = seconds - (hh * 3600) - (mm * 60)
    
    ' Construct the time string in hh:mm:ss.mmm format
    timeStr = Format(hh, "00") & ":" & Format(mm, "00") & ":" & Format(ss, "00.00")
    ConvertSecondsToTimeFormat = timeStr
End Function

Public Sub CallJavaScriptFunctionAsync(ByVal funcName As String, ParamArray params() As Variant)
    Dim i As Integer, j As Integer
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Create a new RequestItem
    Dim requestItem As requestItem
    Set requestItem = New requestItem
    Set requestItem.http = http
    requestItem.funcName = funcName
    requestItem.sheetCodeName = ThisWorkbook.CodeName + "/" + sheetCodeName
        
    ' Generate the timestamp
    Dim timestamp As String
    timestamp = funcName & " " & GetPreciseLocalTimestamp
    
    ' Create JSON request body
    Dim json As Object
    Set json = CreateObject("Scripting.Dictionary")
    json.Add "functionName", funcName
    json.Add "timestamp", timestamp
    For i = LBound(params) + 1 To UBound(params) Step 2
        json.Add params(i), params(i + 1)
    Next i
    Dim jsonRequest As String
    jsonRequest = JsonConverter.ConvertToJson(json)
    
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

    http.oneAnswer = funcName <> "dataGeneratorSub"

    ' Add lines of text to the file
    AddLineToTextFile logFile, "Invoke-RestMethod -Uri http://localhost:3000/execute -Method Post -Headers @{ ""Content-Type"" = ""application/json"" } -Body '" & jsonRequest & "'"
    
    ' Start the timer to check the response periodically
    If params(LBound(params)) And Not timerStarted Then
        CheckHttpResponse
        timerStarted = True
    End If
End Sub

Function GetPreciseLocalTimestamp() As String
    Dim st As SYSTEMTIME
    GetLocalTime st
    GetPreciseLocalTimestamp = _
        Format(st.wYear, "0000") & "-" & _
        Format(st.wMonth, "00") & "-" & _
        Format(st.wDay, "00") & " " & _
        Format(st.wHour, "00") & ":" & _
        Format(st.wMinute, "00") & ":" & _
        Format(st.wSecond, "00") & "." & _
        Format(st.wMilliseconds, "000")
End Function

Function AddLineToTextFile(filePath As String, textToAdd As String)
    ' Open the file for appending (this will add text without erasing existing content)
    Open filePath For Append As #1
    
    ' Write the new line of text to the file
    Print #1, textToAdd
    
    ' Close the file to save changes
    Close #1
End Function

Sub StoreInitialRows(sheetName As String, originalData As Dictionary)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    Dim rowCount As Long
    rowCount = ws.UsedRange.Rows.Count
    
    Dim i As Long
    ' Loop through each row and store its values in the dictionary
    For i = 1 To rowCount
        originalData(i) = ws.Rows(i).Value
    Next i
End Sub

Sub SwapRowsBasedOnList(sheetName As String, originalData As Dictionary, swapList As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    
    Dim i As Integer
    Dim row1 As Long, row2 As Long
    Dim tempRow As Variant

    ' Reset all rows to their original values before performing swaps
    Dim key As Variant
    For Each key In originalData.Keys
        ws.Rows(key).Value = originalData(key)
    Next key
    
    ' Perform row swaps based on swapList
    For i = LBound(swapList) To UBound(swapList) Step 2
        row1 = swapList(i)
        row2 = swapList(i + 1)
        
        ' Store the values temporarily for swapping
        tempRow = ws.Rows(row1).Value
        ws.Rows(row1).Value = ws.Rows(row2).Value
        ws.Rows(row2).Value = tempRow
    Next i
End Sub

Function GetAfterSlash(inputString As String) As String
    Dim position As Integer
    position = InStr(inputString, "/")
    
    If position > 0 Then
        GetAfterSlash = Mid(inputString, position + 1)
    Else
        GetAfterSlash = "" ' Return empty if "/" not found
    End If
End Function

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
                            ' the message to display
                            Dim nextTopLabel As Integer
                            nextTopLabel = labelTopPos
                            Dim ctrl As Control
                            Dim i As Integer
                            
                            ' Loop through all controls in the UserForm
                            For i = UserForm1.Controls.Count - 1 To 0 Step -1
                                Set ctrl = UserForm1.Controls(i)
                                
                                ' Check if the control is a label and was dynamically added
                                If TypeName(ctrl) = "Label" And ctrl.Tag = ListSuggSt Then
                                    UserForm1.Controls.Remove ctrl.Name
                                End If
                            Next i
                            
                            Dim acc As Integer
                            acc = 0
                            Set LabelHandlers = New Collection
                            Dim labelHandler As clsLabelHandler
                            Set listBoxList(sheetCodeName) = New Collection
                            Dim subItem As Variant
                            ' subItem for each 
                            For Each subItem In item(key)
                                Dim callAddress As range
                                Dim oneListItem As Variant
                                For Each oneListItem In subItem
                                    ' Create a new label
                                    Set ctrl = UserForm1.Controls.Add("Forms.Label.1")
                                    With ctrl
                                        .Caption = oneListItem("msg")
                                        .Top = nextTopLabel
                                        .Left = labelLeftPos
                                        .Width = labelWidth
                                        .WordWrap = True ' Enable word wrap
                                        


                                        Dim labelLines As Integer
                                        Dim lines() As String
                                        Dim line As Variant
                                        Dim lineWidth As Single
                                        ' ' Split the text into lines
                                        ' lines = Split(oneListItem("msg"), vbCrLf)
                                        
                                        ' ' Loop through each line and measure its width
                                        ' For Each line In lines
                                        '     lineWidth = UserForm1.TextWidth(line)
                                        '     If lineWidth > labelWidth Then
                                        '         labelLines = labelLines + 1
                                        '     End If
                                        '     labelLines = labelLines + 1
                                        ' Next line





                                        ' Calculate the number of lines in the label
                                        Dim textLength As Integer
                                        
                                        ' Estimate the average character width (this may need adjustment based on font)
                                        lineWidth = labelWidth / avgCharWidth

                                        ' Calculate the number of lines in the label
                                        Dim pos As Integer
                                        Dim pos0 As Integer
                                        pos0 = 1
                                        pos = InStr(1, oneListItem("msg"), vbLf)
                                        If pos = 0 Then
                                            pos = Len(oneListItem("msg")) + 1
                                        End If
                                        labelLines = 0
                                        Do While pos > 0
                                            ' Calculate the number of lines based on text length and line width
                                            ' labelLines = labelLines + Int((pos - pos0) / lineWidth) + 1
                                            labelLines = labelLines + Int((pos - pos0) / lineWidth) + 1

                                            pos = InStr(pos + 1, oneListItem("msg"), vbLf)
                                        Loop
                                        
                                        ' Adjust the height based on the number of lines
                                        .Height = labelHeight * labelLines
                                        nextTopLabel = nextTopLabel + .Height
                                        
                                        ' Assign a unique name or tag
                                        .Name = ListSuggSt & acc
                                        .Tag = ListSuggSt
                                        
                                        ' Set colors for each item
                                        Dim colorRGB As Long
                                        colorRGB = RGB(oneListItem("color")("r"), _
                                                        oneListItem("color")("g"), _
                                                        oneListItem("color")("b"))
                                        .BorderStyle = fmBorderStyleSingle
                                        .BorderColor = colorRGB
                                        .ForeColor = colorRGB




                                        Dim modifiedActions As Collection
                                        Dim oneModifiedAction As Dictionary

                                        ' Create a new Collection to store the modified actions
                                        Set modifiedActions = New Collection

                                        For Each Action In oneListItem("actions")
                                            ' Create a new Dictionary for each action
                                            Set oneModifiedAction = New Dictionary
                                            
                                            ' Copy the existing Action into the new Dictionary
                                            For Each key2 In Action.Keys
                                                If key2 = "address" Then
                                                    ' Modify the "address" field
                                                    If IsNumeric(Action("address")) Then
                                                        ' Store the Range reference directly in the new Dictionary
                                                        Set oneModifiedAction("address") = callAddress
                                                    Else
                                                        ' Calculate the cell reference and store it
                                                        Set callAddress = sheet.Cells(Action("address")(1) + 1, Action("address")(2) + 1)
                                                        Set oneModifiedAction("address") = callAddress
                                                    End If
                                                Else
                                                    ' Handle other object types as needed
                                                    oneModifiedAction(key2) = Action(key2)
                                                End If
                                            Next key2
                                            
                                            ' Add the modified action to the collection
                                            modifiedActions.Add oneModifiedAction
                                        Next Action

                                        ' Add the modified actions collection to the listBoxList (or wherever you need to store them)
                                        listBoxList(sheetCodeName).Add modifiedActions






                                        ' ' Assign the address to the action
                                        ' For Each Action In oneListItem("actions")
                                        '     If IsNumeric(Action("address")) Then
                                        '         Action("address") = callAddress
                                        '     Else
                                        '         Set callAddress = sheet.Cells(Action("address")(1) + 1, Action("address")(2) + 1)
                                        '         Action("address") = callAddress
                                        '     End If
                                        ' Next Action

                                        ' listBoxList(sheetCodeName).Add oneListItem("actions")

                                        acc = acc + 1
                                    End With
                                        
                                    ' Create a new instance of the label handler
                                    Set labelHandler = New clsLabelHandler
                                    Set labelHandler.LabelControl = ctrl
                                    
                                    ' Add the handler to the collection
                                    LabelHandlers.Add labelHandler
                                Next oneListItem
                            Next subItem
                        Case "sort"
                            Dim givenSheet As String
                            givenSheet = GetAfterSlash(item(key)("sheetCodeName"))
                            If fstSort(givenSheet) Then
                                Dim originalData As Object
                                Set originalData = CreateObject("Scripting.Dictionary")
                                
                                ' Store the initial state of the sheet's rows
                                Call StoreInitialRows(givenSheet, originalData)
                                originalDatas(givenSheet) = originalData
                                fstSort(givenSheet) = False
                            End If
                            Call SwapRowsBasedOnList(givenSheet, originalData, item(key)())
                        Case "stop sorting"
                            Dim givenSheet As String
                            givenSheet = GetAfterSlash(item(key)("sheetCodeName"))
                            sheetVBA.Unprotect
                            requests.Remove j
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
                If requests(j).oneAnswer Then
                    requests.Remove j
                End If
            Else
                MsgBox "Error: " & http.statusText
                requests.Remove j
                Dim check As String
                check = SendRequestToNodeJSEndpoint2
                If check = "Request timed out." Then
                    MsgBox "The server is not responding. Please make sure the Node.js server is running."
                    Exit Sub
                End If
            End If
        End If
    Next j

    If requests.Count > 0 Then
        ' Reschedule the check if there are still pending requests
        Application.OnTime Now + TimeValue("00:00:10"), "CheckHttpResponse"
    Else
        timerStarted = False
    End If
End Sub