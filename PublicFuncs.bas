Attribute VB_Name = "PublicFuncs"
Declare Function SetWindowPos Lib "user32" _
         (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
          ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
          ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Sub SetPos(frmName As Form)

frmName.Left = (Screen.Width / 2) - (frmName.Width / 2)
frmName.Top = (Screen.Height / 2) - (frmName.Height / 2)

End Sub

Public Sub FlshWindow()

On Error Resume Next

    Dim nReturnValue As Integer
    nReturnValue = FlashWindow(frmMain.hwnd, True)

End Sub

Public Function KeyFilter(KeyAscii As Integer)

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    KeyAscii = 0
Else
    KeyFilter = KeyAscii
End If

End Function

Public Function Wait(ByVal TimeToWait As Long)
Dim EndTime As Long
EndTime = GetTickCount + TimeToWait * 1000
Do Until GetTickCount > EndTime
DoEvents
Loop
End Function
