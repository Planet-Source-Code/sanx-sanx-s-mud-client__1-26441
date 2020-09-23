VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmMacro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Macros"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgMacros 
      Left            =   5040
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Macro"
      FileName        =   "*.mac"
      Filter          =   "*.mac"
   End
   Begin VB.CommandButton butSaveMacros 
      Caption         =   "&Save..."
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton butLoadMacros 
      Caption         =   "&Load..."
      Height          =   375
      Left            =   2640
      TabIndex        =   3
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
   Begin VB.CommandButton butEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   960
      TabIndex        =   2
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
   Begin VB.CommandButton butNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grdMacro 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butClose_Click()

frmMacro.Hide

End Sub

Private Sub butEdit_Click()

EditMacro grdMacro.Row

End Sub

Private Sub butLoadMacros_Click()

Dim filepath As String

On Error GoTo ErrHandler

dlgMacros.InitDir = frmMain.GetFilePath("macro")
dlgMacros.ShowOpen
filepath = dlgMacros.FileName
frmMain.lstMacro.Clear
frmMain.lstMacroCommand.Clear
frmMain.LoadMacros filepath
PopulateGrid

ErrHandler:

End Sub

Private Sub butNew_Click()

NewMacro

End Sub

Public Sub NewMacro()

frmNewMacro.Show vbModal

End Sub

Private Sub butRemove_Click()

RemoveMacro

End Sub

Public Sub RemoveMacro()

Dim currIndex As Integer

currIndex = grdMacro.Row - 1

response = MsgBox("Are you sure you wish to remove this macro?", vbYesNo + vbApplicationModal + vbQuestion, "Sanx's MUD Client")

If response = vbYes Then
    frmMain.ClearShortcut currIndex
    frmMain.lstMacroName.RemoveItem currIndex
    frmMain.lstMacro.RemoveItem currIndex
    frmMain.lstMacroCommand.RemoveItem currIndex
    PopulateGrid
End If

End Sub

Private Sub butSaveMacros_Click()

On Error GoTo ErrHandler

dlgMacros.InitDir = frmMain.GetFilePath("macro")
dlgMacros.ShowSave
frmMain.SaveMacros dlgMacros.FileName

ErrHandler:

End Sub

Private Sub Form_Load()

SetPos Me

End Sub

Public Sub PopulateGrid()

Dim count As Integer

With grdMacro
    .Col = 0
    .Row = 0
    .Text = "Name"
    .Col = 1
    .Text = "Trigger"
    .Col = 2
    .Text = "Command"
    .Rows = frmMain.lstMacro.ListCount + 1
End With

For count = 0 To frmMain.lstMacro.ListCount - 1
    With grdMacro
    .Row = count + 1
    .Col = 0
    .Text = frmMain.lstMacroName.List(count)
    .Col = 1
    .Text = frmMain.lstMacro.List(count)
    .Col = 2
    .Text = frmMain.lstMacroCommand.List(count)
    End With
Next count

If (grdMacro.Rows - 1) Then
    butEdit.Enabled = True
    butRemove.Enabled = True
Else
    butEdit.Enabled = False
    butRemove.Enabled = False
End If

frmMain.MacroMenu

End Sub

Private Sub grdMacro_DblClick()

EditMacro grdMacro.Row

End Sub

Public Sub EditMacro(macroNum As Integer)

With frmEditMacro
    .txtMacroName.Text = frmMain.lstMacroName.List(macroNum - 1)
    .txtTrigger.Text = frmMain.lstMacro.List(macroNum - 1)
    .txtCommand.Text = frmMain.lstMacroCommand.List(macroNum - 1)
End With

frmEditMacro.Show vbModal

End Sub
