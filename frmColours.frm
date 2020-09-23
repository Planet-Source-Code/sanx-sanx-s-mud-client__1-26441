VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmColours 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Highlighting"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton butAppMsgColor 
      Caption         =   "Change"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin MSComDlg.CommonDialog dlgColour 
      Left            =   5160
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   ".col"
      DialogTitle     =   "Colours"
      FileName        =   "*.col"
      Filter          =   "*.col"
   End
   Begin VB.CommandButton butNew 
      Caption         =   "&New"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton butEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton butRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton butLoadColours 
      Caption         =   "&Load..."
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton butSaveColours 
      Caption         =   "&Save..."
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton butClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grdColours 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Shape shaAppMsg 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2280
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Application Message Colour:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5640
      Y1              =   4320
      Y2              =   4320
   End
End
Attribute VB_Name = "frmColours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butAppMsgColor_Click()

dlgColour.ShowColor
frmMain.UpdateAppMsgColour (dlgColour.Color)
shaAppMsg.FillColor = dlgColour.Color

End Sub

Private Sub butClose_Click()

frmColours.Hide

End Sub

Private Sub butEdit_Click()

frmEditColour.ShowForm grdColours.Row

End Sub

Private Sub butLoadColours_Click()

Dim filepath As String

On Error GoTo ErrHandler

dlgColour.InitDir = frmMain.GetFilePath("colour")
dlgColour.ShowOpen
filepath = dlgColour.FileName
frmMain.lstColours.Clear
frmMain.lstColourTrig.Clear
frmMain.LoadColours filepath
PopulateGrid

ErrHandler:

End Sub

Private Sub butNew_Click()

frmNewColour.Show vbModal, Me

End Sub

Private Sub butRemove_Click()

Dim currIndex As Integer

currIndex = grdColours.Row - 1

response = MsgBox("Are you sure you wish to remove this highlight?", vbYesNo + vbApplicationModal + vbQuestion, "Sanx's MUD Client")

If response = vbYes Then
    frmMain.lstColours.RemoveItem currIndex
    frmMain.lstColourTrig.RemoveItem currIndex
    PopulateGrid
End If

End Sub

Private Sub butSaveColours_Click()

On Error GoTo ErrHandler

dlgColours.InitDir = frmMain.GetFilePath("colour")
dlgColours.ShowSave
frmMain.SaveColours dlgColour.FileName

ErrHandler:

End Sub

Private Sub Form_Load()

With grdColours
    .Row = 0
    .Col = 0
    .Text = "Highlighted Text"
    .Col = 1
    .Text = "Colour"
End With

PopulateGrid

End Sub

Public Sub PopulateGrid()

Dim count As Integer

grdColours.Rows = frmMain.lstColours.ListCount + 1

If frmMain.lstColours.ListCount > 0 Then
    For count = 1 To frmMain.lstColours.ListCount
        With grdColours
            .Row = count
            .Col = 0
            .ColWidth(0) = .Width - (.ColWidth(1) + 100)
            .Text = frmMain.lstColourTrig.List(count - 1)
            .Col = 1
            .FillStyle = flexFillSingle
            .CellBackColor = frmMain.lstColours.List(count - 1)
        End With
    Next count
End If

End Sub

Private Sub grdColours_DblClick()

frmEditColour.ShowForm grdColours.Row

End Sub
