Attribute VB_Name = "Module3"

Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Declare PtrSafe Function CopyImage Lib "user32" (ByVal hImage As LongPtr, ByVal uType As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As LongPtr
Declare PtrSafe Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)


Declare PtrSafe Function Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) As Long