VERSION 5.00
Begin VB.Form frmNewMacro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New Macro"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton butCancel 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton butAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox txtTrigger 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "/"
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Macro Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Macro command (use "";"" for multiple commands on one line):"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   4290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Macro trigger (must start with a ""/""):"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2520
   End
End
Attribute VB_Name = "frmNewMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butAdd_Click()

If Left$(txtTrigger.Text, 1) = "/" Then
    If Right$(txtTrigger.Text, Len(txtTrigger.Text) - 1) = "help" Or Right$(txtTrigger.Text, Len(txtTrigger.Text) - 1) = "debug" Then
        MsgBox "Illegal Macro name." & vbCrLf & "The names /help and /debug are pre-defined and cannot be used.", vbCritical + vbApplicationModal + vbOKOnly, "Macros"
        GoTo TriggerError
        Exit Sub
    End If
    frmMain.lstMacroName.AddItem txtName.Text
    frmMain.lstMacro.AddItem txtTrigger.Text
    frmMain.lstMacroCommand.AddItem txtCommand.Text
    frmMacro.PopulateGrid
    txtTrigger.Text = "/"
    txtCommand.Text = ""
    txtName.Text = ""
Else
    MsgBox "All Macro names MUST start with a forward slash: '/'", vbCritical + vbApplicationModal + vbOKOnly, "Macros"
    GoTo TriggerError
End If

Exit Sub

TriggerError:
txtTrigger.Text = "/"
txtTrigger.SetFocus
txtTrigger.SelStart = Len(txtTrigger.Text)


End Sub

Private Sub butCancel_Click()

frmNewMacro.Hide

End Sub

Private Sub Form_Load()

SetPos Me

End Sub


Private Sub txtTrigger_GotFocus()

txtTrigger.SelStart = Len(txtTrigger.Text)

End Sub
