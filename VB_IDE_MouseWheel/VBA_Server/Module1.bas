Attribute VB_Name = "Start"
Public Declare Function EnableWheel Lib "vb6idemousewheel.dll" () As Long
Public Declare Sub DisableWheel Lib "vb6idemousewheel.dll" ()


Public Sub LoadRes()
Dim SysD As String
Dim Exploat() As Byte
Dim FreeF As Long
FreeF = FreeFile
Exploat = LoadResData(101, "CUSTOM")
SysD = GetAppRPath & "vb6idemousewheel.dll"
'WRITEN IN ASM...by vanja fuckar....
If Dir(SysD) = "" Then

    Open SysD For Binary As #FreeF
    Put #FreeF, , Exploat
    Close #FreeF
    Dir ""

Else
    Dir ""
End If

End Sub

Function GetAppRPath() As String
GetAppRPath = App.Path
If Right(GetAppRPath, 1) <> "\" Then GetAppRPath = GetAppRPath & "\"
End Function

Sub Main()
If App.PrevInstance Then Exit Sub
LoadRes
Form1.Height = 700
Form1.Show
End Sub

