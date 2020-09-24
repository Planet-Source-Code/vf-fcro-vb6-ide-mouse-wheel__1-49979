Attribute VB_Name = "SysMenu"
Declare Function GetSystemMenu Lib "User32" _
(ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function DeleteMenu Lib "User32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DrawMenuBar Lib "User32" (ByVal hWnd As Long) As Long
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Const SC_SIZE = &HF000
Public Sub RemoveResize(ByVal hWnd As Long)
Dim hMenu As Long
hMenu = GetSystemMenu(hWnd, 0&)
If hMenu Then
Call DeleteMenu(hMenu, SC_SIZE, 0)
DrawMenuBar (hWnd)
End If
End Sub
Public Sub TopMost(ByVal hWnd As Long)
SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
