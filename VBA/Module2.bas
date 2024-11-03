Attribute VB_Name = "Module2"
' API declarations for 64-bit Office




' Add reference: "Microsoft Scripting Runtime" for FileSystemObject
' Add reference: "Microsoft Shell Controls and Automation" for Shell


Const CF_BITMAP = 2 ' Bitmap format

Public storedImages As Collection
Public sheetCodeName As String
Public requests As Collection
Public logFile As String
Public timerStarted As Boolean
Public sheetList As Object
Public values As Object

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
Public spacing As Single
Public leftSpacing As Single
Public thbnlDisplayed As Boolean


Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type


Function Min(a As Integer, b As Integer) As Integer
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function