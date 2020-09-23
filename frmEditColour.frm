VERSION 5.00
Begin VB.Form frmEditColour 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Colour Highlight"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton butCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton butUpdate 
      Caption         =   "Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton butChange 
      Caption         =   "Change"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtTextHilite 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Shape shaColour 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   840
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Colour:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Text to highlight:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmEditColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim colIndex As Integer

Private Sub butUpdate_Click()

If txtTextHilite.Text <> "" Then
    frmMain.lstColourTrig.List(colIndex) = txtTextHilite.Text
    frmMain.lstColours.List(colIndex) = shaColour.FillColor
    frmEditColour.Hide
    txtTextHilite.Text = ""
    shaColour.FillColor = RGB(0, 0, 0)
    frmColours.PopulateGrid
End If

End Sub

Private Sub butChange_Click()

frmColours.dlgColour.ShowColor
shaColour.FillColor = frmColours.dlgColour.Color

End Sub

Public Sub ShowForm(index As Integer)

colIndex = index - 1
txtTextHilite.Text = frmMain.lstColourTrig.List(colIndex)
shaColour.FillColor = frmMain.lstColours.List(colIndex)
frmEditColour.Show vbModal, frmColours

End Sub

