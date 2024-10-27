VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ImageListForm 
   Caption         =   "UserForm2"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "ImageListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ImageListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Place these declarations at the top of the UserForm code module

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long

Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000


Private Sub MakeUserFormResizable()
    Dim hWnd As LongPtr
    Dim currentStyle As Long

    ' Get the handle of the UserForm window
    hWnd = FindWindow("ThunderDFrame", Me.Caption)
    
    ' Check if the handle is found
    If hWnd <> 0 Then
        ' Get the current window style
        currentStyle = GetWindowLong(hWnd, GWL_STYLE)
        
        ' Add the WS_THICKFRAME style to make the form resizable
        SetWindowLong hWnd, GWL_STYLE, currentStyle Or WS_THICKFRAME
    End If
End Sub


Private Sub UserForm_Initialize()
    ' Call the procedure to make the form resizable
    MakeUserFormResizable
End Sub

Private Sub UserForm_Resize()
   RemoveAllControls ImageListForm
   DisplayImagesInUserForm
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Check if the form is being closed by the user (e.g., by clicking the "X" button)
    If CloseMode = vbFormControlMenu Then
        thbnlDisplayed = False
    End If
End Sub