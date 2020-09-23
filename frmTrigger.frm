VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTrigger 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Triggers"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgTriggers 
      Left            =   5160
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Triggers"
      FileName        =   "*.tri"
      Filter          =   "*.tri"
   End
   Begin VB.CommandButton butSaveTriggers 
      Caption         =   "&Save..."
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton butLoadTriggers 
      Caption         =   "&Load..."
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton butNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton butClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton butEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton butRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grdTrigger 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6165
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub butClose_Click()

frmTrigger.Hide

End Sub

Private Sub butEdit_Click()

EditTrigger grdTrigger.Row

End Sub

Private Sub butLoadTriggers_Click()

Dim filepath As String

On Error GoTo ErrHandler

dlgTriggers.InitDir = frmMain.GetFilePath("trigger")
dlgTriggers.ShowOpen
filepath = dlgTriggers.FileName
frmMain.lstTrigger.Clear
frmMain.lstTriggerCommand.Clear
frmMain.LoadTriggers filepath
PopulateGrid

ErrHandler:

End Sub

Private Sub butNew_Click()

frmNewTrigger.Show vbModal

End Sub

Private Sub butRemove_Click()

Dim currIndex As Integer

currIndex = grdTrigger.Row - 1

response = MsgBox("Are you sure you wish to remove this Trigger?", vbYesNo + vbApplicationModal + vbQuestion, "Sanx's MUD Client")

If response = vbYes Then
    frmMain.lstTrigger.RemoveItem currIndex
    frmMain.lstTriggerCommand.RemoveItem currIndex
    PopulateGrid
End If

End Sub

Private Sub butSaveTriggers_Click()

On Error GoTo ErrHandler

dlgTriggers.InitDir = frmMain.GetFilePath("trigger")
dlgTriggers.ShowSave
frmMain.SaveTriggers dlgTriggers.FileName

ErrHandler:

End Sub

Private Sub Form_Load()

SetPos Me

End Sub

Public Sub PopulateGrid()

Dim count As Integer

With grdTrigger
    .Col = 0
    .Row = 0
    .Text = "Trigger"
    .Col = 1
    .Text = "Command"
    .Rows = frmMain.lstTrigger.ListCount + 1
End With

For count = 0 To frmMain.lstTrigger.ListCount - 1
    With grdTrigger
    .Row = count + 1
    .Col = 0
    .Text = frmMain.lstTrigger.List(count)
    .Col = 1
    .Text = frmMain.lstTriggerCommand.List(count)
    End With
Next count

If (grdTrigger.Rows - 1) Then
    butEdit.Enabled = True
    butRemove.Enabled = True
Else
    butEdit.Enabled = False
    butRemove.Enabled = False
End If

End Sub

Private Sub grdTrigger_DblClick()

EditTrigger grdTrigger.Row

End Sub

Private Sub EditTrigger(TriggerNum As Integer)

With frmEditTrigger
    .txtTrigger.Text = frmMain.lstTrigger.List(TriggerNum - 1)
    .txtCommand.Text = frmMain.lstTriggerCommand.List(TriggerNum - 1)
End With

frmEditTrigger.Show vbModal

End Sub
