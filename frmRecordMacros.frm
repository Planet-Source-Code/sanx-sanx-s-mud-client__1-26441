VERSION 5.00
Begin VB.Form frmRecordMacros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Record New Macro"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtMacroCommand 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "/"
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtCmdSeq 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtMacroName 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton butStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      Picture         =   "frmRecordMacros.frx":0000
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton butRecord 
      Caption         =   "Record"
      Height          =   375
      Left            =   120
      Picture         =   "frmRecordMacros.frx":08CA
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Macro Command"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Command Sequence:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "New Macro Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   1320
      Y1              =   2520
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2520
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frmRecordMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butRecord_Click()

If Not CheckValues() Then Exit Sub
butRecord.Enabled = False
butStop.Enabled = True
txtMacroName.Enabled = False
txtMacroCommand.Enabled = False
frmMain.SetRecording (True)
frmMain.txtEntry.SetFocus

End Sub

Private Sub butStop_Click()

frmMain.SetRecording (False)
Me.Visible = False
response = MsgBox("Would you like to add this macro?", vbYesNo, "Record Macro")

If response = vbYes Then
    frmMain.lstMacroName.AddItem txtMacroName.Text
    frmMain.lstMacro.AddItem txtMacroCommand.Text
    frmMain.lstMacroCommand.AddItem txtCmdSeq.Text
    frmMacro.PopulateGrid
End If

txtMacroCommand.Text = "/"
txtMacroCommand.Enabled = True
txtMacroName.Enabled = True
txtMacroName.Text = ""
txtCmdSeq.Text = ""
butRecord.Enabled = True
butStop.Enabled = False

End Sub

Private Sub Form_Load()

Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)

End Sub


Private Sub txtMacroCommand_GotFocus()

If txtMacroCommand.Text = "/" Then txtMacroCommand.SelStart = Len(txtMacroCommand.Text)

End Sub

Private Function CheckValues()

If txtMacroName.Text = "" Then
    MsgBox "You need to specify a Macro name.", vbOKOnly + vbCritical, "Record New Macro"
    txtMacroName.SetFocus
    CheckValues = False
    Exit Function
End If

If txtMacroCommand.Text = "" Then
    MsgBox "You need to set a Macro command string.", vbOKOnly + vbCritical, "Record New Macro"
    txtMacroCommand.SetFocus
    CheckValues = False
    Exit Function
End If

CheckValues = True

End Function
