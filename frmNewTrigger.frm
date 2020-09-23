VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmNewTrigger 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New Trigger"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butSoundPicker 
      Caption         =   "..."
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtTrigger 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CommandButton butAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton butCancel 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgSound 
      Left            =   4080
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Choose sound file"
      Filter          =   "*.wav;*.mp3;*.mid;*.mpe"
      InitDir         =   "\"
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      Caption         =   "Play sound when trigger is activated"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Trigger text:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Trigger command (use "";"" for multiple commands on one line):"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4335
   End
End
Attribute VB_Name = "frmNewTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butAdd_Click()

frmMain.lstTrigger.AddItem txtTrigger.Text
frmMain.lstTriggerCommand.AddItem txtCommand.Text
frmTrigger.PopulateGrid
txtTrigger.Text = ""
txtCommand.Text = ""

End Sub

Private Sub butCancel_Click()

frmNewTrigger.Hide

End Sub

Private Sub butSoundPicker_Click()

dlgSound.ShowOpen
If dlgSound.FileName <> "" Then
    txtCommand.Text = txtCommand.Text + "**PLAYSND**" + dlgSound.FileName + "+++"
End If

End Sub

Private Sub Form_Load()

SetPos Me

End Sub


Private Sub txtTrigger_GotFocus()

txtTrigger.SelStart = Len(txtTrigger.Text)

End Sub

