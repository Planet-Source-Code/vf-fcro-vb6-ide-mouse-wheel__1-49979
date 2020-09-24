VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB6 IDE Mouse Wheel (Server) By Vanja Fuckar v1.1"
   ClientHeight    =   60
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5205
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   60
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu ACT 
      Caption         =   "Action"
      Begin VB.Menu Command1 
         Caption         =   "Start Server"
      End
      Begin VB.Menu Command2 
         Caption         =   "Stop Server"
      End
      Begin VB.Menu Command5 
         Caption         =   "Change scroll line number"
      End
   End
   Begin VB.Menu Command4 
      Caption         =   "Minimize"
   End
   Begin VB.Menu Command3 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private EnabledSVR As Boolean

Private Sub Command1_Click()
Dim ret As Long
If Not EnabledSVR Then
    ret = EnableWheel
    If ret = 0 Then MsgBox "Cannot establish connection!", vbCritical, "Error!": Exit Sub
    EnabledSVR = True
    WriteLn
    Command4_Click
End If

End Sub

Private Sub Command2_Click()
If EnabledSVR Then
    DisableWheel
    EnabledSVR = False
     WriteLn
    Command4_Click
End If
       
End Sub

Private Sub WriteLn()
If EnabledSVR Then
ChangeText "Server Running,Lines per scroll:" & LN & Chr(0)
    Else
ChangeText "Server Stopped,Lines per scroll:" & LN & Chr(0)
End If
End Sub


Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
WindowState = 1
Visible = False
End Sub

Private Sub Command5_Click()
On Error GoTo Dalje
Dim LSS As String
Dim LS As Long
LSS = InputBox("New line number:", "Change Scroll line number (1-100)")
If Len(LSS) = 0 Then Exit Sub
LS = CLng(LSS)
If LS < 1 Or LS > 100 Then MsgBox "Out of range!", vbCritical, "Information": Exit Sub
EnumWindows AddressOf EnumTops, LS 'Broadcast value!
SetLines LS
LN = LS
WriteLn
Exit Sub
Dalje:
On Error GoTo 0
MsgBox "Error in value!", vbExclamation, "Information"
End Sub



Private Sub Form_Load()
RemoveResize hwnd
RegClass "SHNOTIFYCW", 0, AddressOf ShNotifyProc
InitREGISTRY
CreateSHN Me.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
CloseREGISTRY
Command2_Click
Shell_NotifyIcon NIM_DELETE, SH
DestroyWindow SHNwnd
UnregisterClass "SHNOTIFYCW", 0
End Sub


