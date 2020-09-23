VERSION 5.00
Begin VB.Form frmEditMacro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Macro"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMacroName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox txtTrigger 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "\"
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4335
   End
   Begin VB.CommandButton butOK 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton butCancel 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Macro Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Macro trigger (must start with a ""\""):"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2520
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Macro command (use "";"" for multiple commands on one line):"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4290
   End
End
Attribute VB_Name = "frmEditMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butCancel_Click()

frmEditMacro.Hide

End Sub

Private Sub butOK_Click()

Dim currIndex As Integer

currIndex = frmMacro.grdMacro.Row - 1

frmMain.lstMacroName.List(currIndex) = txtMacroName.Text
frmMain.lstMacro.List(currIndex) = txtTrigger.Text
frmMain.lstMacroCommand.List(currIndex) = txtCommand.Text
frmMacro.PopulateGrid

frmEditMacro.Hide

End Sub

Private Sub Form_Load()

SetPos Me

End Sub

Private Sub txtTrigger_GotFocus()

txtTrigger.SelStart = Len(txtTrigger.Text)

End Sub
