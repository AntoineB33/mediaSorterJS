Attribute VB_Name = "Module1"
' API declarations for 64-bit Office
Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Declare PtrSafe Function CopyImage Lib "user32" (ByVal hImage As LongPtr, ByVal uType As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As LongPtr

Const CF_BITMAP = 2 ' Bitmap format

Public storedImages As Collection


Option Explicit

' Add reference: "Microsoft Scripting Runtime" for FileSystemObject
' Add reference: "Microsoft Shell Controls and Automation" for Shell


Sub DisplayImagesInUserForm()
    Dim ws As Worksheet
    Dim imgControl As MSForms.Image
    Dim lblControl As MSForms.Label
    Dim i As Long
    Dim totalHeight As Long
    Dim spacing As Long
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
    totalHeight = 10
    spacing = 10

    ' Create FileSystemObject to check if file exists
    Set fso = CreateObject("Scripting.FileSystemObject")

    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ' Get the filename without extension from the 6th column (e.g., column F)
        FileNameWithoutExt = ws.Cells(i, 6).Value
        ' Check if the file exists
        FilePath = SearchFileInFolder(FolderPath, FileNameWithoutExt)

        If FilePath <> "" Then
            ' Add Image control
            Set imgControl = frm.Controls.Add("Forms.Image.1")

            ' Check if the file is an image
            foundImage = False
            If IsImageFile(FilePath) Then
                imgControl.Picture = LoadPicture(FilePath)
            ElseIf IsVideoFile(FilePath) Then
                ' If it's a video, extract a frame
                imgControl.Picture = ExtractFrameFromVideo(FilePath)
            End If

            ' Set the size and position of the Image control
            imgControl.Width = 200
            imgControl.Height = 200
            imgControl.Top = totalHeight
            imgControl.Left = 10

            ' Update totalHeight for the next image
            totalHeight = totalHeight + imgControl.Height + spacing
        Else
            ' If the image is not found, add a Label control instead
            Set lblControl = frm.Controls.Add("Forms.Label.1")
            lblControl.Caption = "Image not found: " & FileNameWithoutExt
            lblControl.Width = 200
            lblControl.Height = 30
            lblControl.Top = totalHeight
            lblControl.Left = 10
            lblControl.ForeColor = RGB(255, 0, 0) ' Set text color to red

            ' Update totalHeight for the next image or message
            totalHeight = totalHeight + lblControl.Height + spacing
        End If
    Next i

    ' Set UserForm size to fit all images
    frm.Height = totalHeight + 30
    frm.Width = 230

    ' Show the UserForm
    frm.Show
End Sub

' Function to extract a frame from a video at 10% of its duration and return as StdPicture
Function ExtractFrameFromVideo(videoPath As String) As StdPicture
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
    tempImagePath = "C:\Users\abarb\AppData\Local\Temp\frame_temp.jpg"
    
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
    
    ' Step 7: Delete the temporary image file
    Kill tempImagePath
    
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