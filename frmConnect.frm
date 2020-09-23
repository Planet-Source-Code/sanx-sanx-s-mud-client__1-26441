VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Connect..."
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butDeleteWorld 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton butUpdateWorld 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton butNewWorld 
      Caption         =   "&New"
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox lstWorlds 
      Height          =   315
      ItemData        =   "frmConnect.frx":0000
      Left            =   120
      List            =   "frmConnect.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton butCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton butConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "World:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3360
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label5 
      Caption         =   "Password:"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   3480
      X2              =   3480
      Y1              =   120
      Y2              =   2040
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butCancel_Click()

frmConnect.Hide

End Sub

Private Sub butConnect_Click()

frmMain.SocketConnect txtHost.Text, txtPort.Text, txtUsername.Text, txtPassword.Text
frmConnect.Hide

End Sub

Private Sub butNewWorld_Click()

CreateNewWorld

End Sub

Public Sub CreateNewWorld()

frmNewWorld.Show vbModal

End Sub

Private Sub butUpdateWorld_Click()

If lstWorlds.List(lstWorlds.ListIndex) <> "" Then
    frmMain.SaveSettings lstWorlds.List(lstWorlds.ListIndex)
End If

End Sub

Private Sub Form_Load()

SetPos Me

End Sub
Public Sub GetWorlds()

Dim numWorlds As Integer
Dim count As Integer

lstWorlds.Clear

numWorlds = GetSetting(App.Title, "Worlds", "NumWorlds", 0)

If numWorlds > 0 Then
    For count = 1 To numWorlds
        lstWorlds.AddItem GetSetting(App.Title, "Worlds", "World" + Format(count), "--blank--")
    Next count
End If

End Sub

Private Sub lstWorlds_Click()

Dim selworld As String

selworld = lstWorlds.List(lstWorlds.ListIndex)

If selworld <> "" Then
    frmMain.LoadGameSettings selworld
    butDeleteWorld.Enabled = True
    butUpdateWorld.Enabled = True
Else
    butDeleteWorld.Enabled = False
    butUpdateWorld.Enabled = False
End If

frmMain.SetConnValues

End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)

KeyAscii = KeyFilter(KeyAscii)

End Sub

Private Sub txtPort_LostFocus()

If Val(txtPort.Text) < 23 Or Val(txtPort.Text) > 65000 Then
    MsgBox "Value must be an integer between 25 and 65000", vbOKOnly + vbApplicationModal + vbCritical, "Sanx's MUD Client"
    txtPort.SetFocus
    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort.Text)
End If

End Sub
