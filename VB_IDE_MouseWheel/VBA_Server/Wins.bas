Attribute VB_Name = "Wins"


Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Function ClassNameEx(ByVal hwnd As Long) As String
Dim ClLen As Long
ClassNameEx = Space(260)
ClLen = GetClassName(hwnd, ClassNameEx, 260)
ClassNameEx = Left(ClassNameEx, ClLen)
End Function


Public Function EnumTops(ByVal hwnd As Long, ByVal UserData As Long) As Long
EnumChildWindows hwnd, AddressOf EnumChilds, UserData
EnumTops = 1
End Function


Public Function EnumChilds(ByVal hwnd As Long, ByVal UserData As Long) As Long
If StrComp(ClassNameEx(hwnd), "VbaWindow", vbTextCompare) = 0 Then
SendMessage hwnd, &H5005, UserData, ByVal 0&
End If
EnumChilds = 1
End Function



