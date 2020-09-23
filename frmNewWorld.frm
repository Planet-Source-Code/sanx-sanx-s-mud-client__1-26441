VERSION 5.00
Begin VB.Form frmNewWorld 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create New World"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butCreate 
      Caption         =   "Create"
      Default         =   -1  'True
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton butCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame fraCopyWorld 
      Height          =   1695
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   4095
      Begin VB.TextBox txtCopyWorldName 
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   3615
      End
      Begin VB.ComboBox lstWorlds 
         Height          =   315
         ItemData        =   "frmNewWorld.frx":0000
         Left            =   240
         List            =   "frmNewWorld.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label lblCopyWorldName 
         Caption         =   "World Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblCopyWorld 
         Caption         =   "World to copy from:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraBlankWorld 
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   4095
      Begin VB.TextBox txtBlankWorldName 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label lblBlankWorldName 
         Caption         =   "World Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.OptionButton optCopyWorld 
      Caption         =   "Create from an existing world"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.OptionButton optNewWorld 
      Caption         =   "Create blank world"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2055
   End
End
Attribute VB_Name = "frmNewWorld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butCancel_Click()

frmNewWorld.Hide

End Sub

Private Sub butCreate_Click()

Dim NewWorldNumber As Integer
Dim NewWorldName As String
Dim selworld As String

If optNewWorld.value = True Then
    NewWorldName = txtBlankWorldName.Text
Else
    NewWorldName = txtCopyWorldName.Text
End If

If optNewWorld.value = True Then
    frmMain.LoadGameSettings "--blank--"
    CreateWorld NewWorldName
Else
    selworld = lstWorlds.List(lstWorlds.ListIndex)
    If selworld <> "" Then
        frmMain.LoadGameSettings selworld
        CreateWorld NewWorldName
    Else
        MsgBox "You have selected to create a world based on an existing world." + vbCrLf + "Therefore, you must select an existing world to copy from.", vbOKOnly, "No existing world selected."
    End If
End If

End Sub

Private Sub CreateWorld(NewWorldName As String)

NewWorldNumber = GetWorldNumber()
SaveSetting App.Title, "Worlds", "NumWorlds", NewWorldNumber
SaveSetting App.Title, "Worlds", "World" + Format(NewWorldNumber), NewWorldName
frmMain.SetFilePath "macro", App.Path & "\" & NewWorldName & ".mac"
frmMain.SetFilePath "trigger", App.Path & "\" & NewWorldName & ".tri"
frmMain.SetFilePath "colour", App.Path & "\" & NewWorldName & ".col"
frmMain.SaveSettings NewWorldName

frmConnect.GetWorlds
frmNewWorld.Hide

End Sub

Private Sub Form_Load()

SetPos Me
SetVisibleOptions

End Sub

Private Sub optCopyWorld_Click()

SetVisibleOptions

End Sub

Private Sub optNewWorld_Click()

SetVisibleOptions

End Sub

Private Sub SetVisibleOptions()

If optNewWorld.value = True Then
    fraBlankWorld.Enabled = True
    txtBlankWorldName.Enabled = True
    lblBlankWorldName.Enabled = True
    fraCopyWorld.Enabled = False
    lstWorlds.Enabled = False
    txtCopyWorldName.Enabled = False
    lblCopyWorld.Enabled = False
    lblCopyWorldName.Enabled = False
Else
    fraBlankWorld.Enabled = False
    txtBlankWorldName.Enabled = False
    lblBlankWorldName.Enabled = False
    fraCopyWorld.Enabled = True
    lstWorlds.Enabled = True
    txtCopyWorldName.Enabled = True
    lblCopyWorld.Enabled = True
    lblCopyWorldName.Enabled = True
    GetWorlds
End If

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

Private Function GetWorldNumber()

Dim numWorlds As Integer

numWorlds = GetSetting(App.Title, "Worlds", "NumWorlds", 0)

GetWorldNumber = numWorlds + 1

End Function
