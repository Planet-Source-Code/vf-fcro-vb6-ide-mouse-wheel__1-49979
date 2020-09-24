Attribute VB_Name = "ShellN"
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long


Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205


Public Type WNDCLASS
    Style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Const CS_CLASSDC = &H40
Public Const CS_OWNDC = &H20
Public Const CS_GLOBALCLASS = &H4000
Public Const CS_HREDRAW = &H2
Public Const CS_PARENTDC = &H80
Public Const CS_VREDRAW = &H1

Public SH As NOTIFYICONDATA
Public SHNwnd As Long

Public Sub RegClass(ByVal Classname As String, ByVal hInstance As Long, ByVal AddressProc As Long)
Dim CLASSX As WNDCLASS
CLASSX.Style = CS_GLOBALCLASS
CLASSX.lpfnwndproc = AddressProc
CLASSX.hInstance = hInstance
CLASSX.lpszClassName = Classname
Call RegisterClass(CLASSX)
End Sub

Public Sub ChangeText(ByRef NewText As String)
SH.cbSize = Len(SH)
SH.szTip = NewText
SH.uID = 1
SH.uFlags = NIF_TIP
SH.uCallbackMessage = &H6002&
Shell_NotifyIcon NIM_MODIFY, SH
End Sub

Public Sub CreateSHN(ByVal Icon As Long)
SHNwnd = CreateWindowEx(0, "SHNOTIFYCW", vbNullString, 0, 0, 0, 0, 0, 0, 0, App.hInstance, 0)
SH.cbSize = Len(SH)
SH.hIcon = Icon
SH.hwnd = SHNwnd
SH.szTip = "Server Stopped,Lines per scroll:" & LN & Chr(0)
SH.uID = 1
SH.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
SH.uCallbackMessage = &H6002&
Shell_NotifyIcon NIM_ADD, SH
End Sub


Public Function ShNotifyProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim ShMsg As Long
Dim ShId As Long
Dim MM As String
If wMsg = &H6002& Then

    ShMsg = lParam And &HFFFF&
    ShId = wParam And &HFFFF&

    If ShMsg = WM_LBUTTONUP Then
        TopMost Form1.hwnd
        Form1.Visible = True
        Form1.WindowState = 0
    End If

End If

ShNotifyProc = DefWindowProc(hwnd, wMsg, wParam, lParam)
End Function

