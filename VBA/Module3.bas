Attribute VB_Name = "Module3"

Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Declare PtrSafe Function CopyImage Lib "user32" (ByVal hImage As LongPtr, ByVal uType As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As LongPtr
Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)


Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long




' Add reference: "Microsoft Scripting Runtime" for FileSystemObject
' Add reference: "Microsoft Shell Controls and Automation" for Shell


Const CF_BITMAP = 2 ' Bitmap format

Public storedImages As Collection
Public sheetCodeName As String
Public requests As Collection
Public logFile As String
Public timerStarted As Boolean
Public sheetList As Object

Public labelTopPos As Integer
Public labelLeftPos As Integer
Public labelHeight As Integer
Public labelSpacing As Integer
Public labelWidth As Integer
Public avgCharWidth As Integer
Public lastIsChg As Boolean
Public sheetVBA As Worksheet
Public sheet As Worksheet
Public sorting As Object
Public targetAddress As Object
Public newCollection As Collection
Public LabelHandlers As Collection
Public listBoxList As Object
Public ListSuggSt As String

Public editRow As Long
Public editColumn As Long
Public thumbnailHeight As Single
Public formHeight As Single
Public spacing As Single
Public leftSpacing As Single


Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type


Sub DisplayImagesInUserForm()
    Dim ws As Worksheet
    Dim imgControl As MSForms.Image
    Dim lblControl As MSForms.Label
    Dim i As Long
    Dim totalHeight As Long
    Dim fso As Object
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

    ' Initialize totalHeight and spacing
    totalHeight = spacing

    ' Set UserForm size to fit all images
    frm.Height = formHeight
    frm.Width = thumbnailHeight + 10
    frm.Show vbModeless

    ' Create FileSystemObject to check if file exists
    Set fso = CreateObject("Scripting.FileSystemObject")

    
    

    


    Dim resultList As Collection
    Set resultList = New Collection
    
    Dim startNum As Integer
    Dim minInt As Integer
    Dim maxInt As Integer
    Dim maxCount As Integer

    Dim selectedThumbnail As Boolean
    selectedThumbnail = False
    
    ' Initialize values
    startNum = editRow
    minInt = 2
    maxInt = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If startNum < minInt Then
        startNum = minInt
    ElseIf startNum > maxInt Then
        startNum = maxInt
    Else
        selectedThumbnail = True
    End If
    
    maxCount = WorksheetFunction.Ceiling(formHeight / (thumbnailHeight + spacing), 1)

    ' Variables for controlling the pattern
    Dim decrease As Boolean
    Dim count As Integer
    decrease = True
    count = 0

    ' Loop until we reach both the minInt and maxInt, or reach maxCount
    Do While count < maxCount And Not (startNum < minInt And Not decrease) And Not (startNum > maxInt And decrease)
        ' Add the current number to the list
        resultList.Add startNum
        count = count + 1
        
        ' Alternate between decreasing and increasing
        If decrease Then
            startNum = startNum - 1
            If startNum <= minInt Then
                decrease = False
            End If
        Else
            startNum = startNum + 1
            If startNum >= maxInt Then
                decrease = True
            End If
        End If
    Loop



    Dim selectThbnailTop As Single
    Dim distToSlctThbnailTop As Single
    distToSlctThbnailTop = 0
    For i = 1 To count
        Dim mediaInd As Integer
        mediaInd = resultList(i) - 2
        ' Get the filename without extension from the 6th column (e.g., column F)
        FileNameWithoutExt = ws.Cells(resultList(i), 6).Value
        ' Check if the file exists
        FilePath = SearchFileInFolder(FolderPath, FileNameWithoutExt)

        Dim leftVal As Single
        Dim shape As Shape
        If selectedThumbnail Then
            If mediaInd < maxCount Then
                leftVal = mediaInd * thumbnailHeight * spacing
            ElseIf maxInt - mediaInd < maxInt Then
                leftVal = formHeight - ((maxInt - mediaInd) * thumbnailHeight * spacing + thumbnailHeight)
            Else
                leftVal = formHeight / 2 - thumbnailHeight / 2
            End If
            Set shape = ws.Shapes.AddShape(msoShapeRectangle, leftSpacing + 200, leftVal, leftSpacing, 200)
            shape.Line.ForeColor.RGB = RGB(0, 0, 0)
            selectedThumbnail = False
        Else
            If resultList(i) < resultList(i - 1) Then
                If distToSlctThbnailTop > 0 Then
                    distToSlctThbnailTop = -distToSlctThbnailTop
                End If
                distToSlctThbnailTop = distToSlctThbnailTop - thumbnailHeight + spacing
            Else
                If distToSlctThbnailTop < 0 Then
                    distToSlctThbnailTop = -distToSlctThbnailTop
                End If
                distToSlctThbnailTop = distToSlctThbnailTop + thumbnailHeight + spacing
            End If
            leftVal = selectThbnailTop + distToSlctThbnailTop
        End If
        If FilePath <> "" Then
            ' Add Image control
            Set imgControl = frm.Controls.Add("Forms.Image.1")
            With imgControl
                ' Check if the file is an image
                foundImage = False
                Set shape = ws.Shapes.AddShape(msoShapeRectangle, 0, leftVal, leftSpacing, 200)
                If IsImageFile(FilePath) Then
                    shape.Line.ForeColor.RGB = RGB(255, 0, 0)
                    Set .Picture = LoadPicture(FilePath)
                ElseIf IsVideoFile(FilePath) Then
                    shape.Line.ForeColor.RGB = RGB(0, 0, 255)
                    ' If it's a video, extract a frame
                    Set .Picture = ExtractFrameFromVideo(FilePath, FileNameWithoutExt)
                End If
                .PictureSizeMode = fmPictureSizeModeZoom
                .Left = leftSpacing
                .Top = leftVal
            End With
            Call ResizeImage(imgControl, 200, 200)
        Else
            ' If the image is not found, add a Label control instead
            Set lblControl = frm.Controls.Add("Forms.Label.1")
            With lblControl
                .Caption = "Image not found: " & FileNameWithoutExt
                .Width = 200
                .Height = 30
                .Left = leftSpacing
                .Top = leftVal
                .ForeColor = RGB(255, 0, 0)
            End With
        End If
    Next i
End Sub

' Helper function to maintain aspect ratio and resize image
Sub ResizeImage(imgControl As MSForms.Image, maxWidth As Long, maxHeight As Long)
    Dim imgWidth As Long
    Dim imgHeight As Long
    Dim aspectRatio As Double

    ' Get the original image dimensions
    imgWidth = imgControl.Picture.Width / 26.458 ' Convert from twips to points
    imgHeight = imgControl.Picture.Height / 26.458 ' Convert from twips to points

    ' Calculate aspect ratio
    aspectRatio = imgWidth / imgHeight

    ' Resize the image proportionally
    If imgWidth > imgHeight Then
        imgControl.Width = maxWidth
        imgControl.Height = maxWidth / aspectRatio
    Else
        imgControl.Height = maxHeight
        imgControl.Width = maxHeight * aspectRatio
    End If
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
    requestItem.sheetCodeName = sheetCodeName
        
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

    ' Add lines of text to the file
    AddLineToTextFile logFile, "Invoke-RestMethod -Uri http://localhost:3000/execute -Method Post -Headers @{ ""Content-Type"" = ""application/json"" } -Body '" & jsonRequest & "'"
    
    ' Start the timer to check the response periodically
    If params(LBound(params)) And Not timerStarted Then
        Debug.Print "ttttttt"
        CheckHttpResponse0
        CheckHttpResponse
        timerStarted = True
    End If
End Sub


Public Sub CheckHttpResponse0()
    Debug.Print "ddddddddddddd"
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
        Application.OnTime Now + TimeValue("00:00:01"), "CheckHttpResponse"
    Else
        timerStarted = False
    End If
End Sub
