Attribute VB_Name = "Registry"
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY As Long) As Long
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal HKEY As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String) As Long
Public Const HKEY_CURRENT_USER     As Long = &H80000001
Public Const HKEY_CLASSES_ROOT As Long = &H80000000
Public Const REG_DWORD = 4

Public LN As Long
Public HKEY As Long

Public Sub InitREGISTRY()
If RegOpenKey(HKEY_CLASSES_ROOT, "VB6IDEMOUSEWHEEL", HKEY) <> 0 Then
RegCreateKey HKEY_CLASSES_ROOT, "VB6IDEMOUSEWHEEL", HKEY
SetLines 1
End If

LN = GetLines
End Sub
Public Sub CloseREGISTRY()
RegCloseKey HKEY
End Sub

Public Function GetLines() As Long
Dim RType As Long
RegQueryValueEx HKEY, "Lines", 0, RType, GetLines, 4
End Function

Public Sub SetLines(ByVal newlines As Long)
RegSetValueEx HKEY, "Lines", 0, REG_DWORD, newlines, 4
End Sub
