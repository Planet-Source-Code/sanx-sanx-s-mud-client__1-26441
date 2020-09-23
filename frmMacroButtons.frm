VERSION 5.00
Begin VB.Form frmMacroButtons 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Macro Shortcuts"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton butCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox lstMacros 
      Enabled         =   0   'False
      Height          =   315
      Index           =   9
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   3720
      Width           =   1935
   End
   Begin VB.ComboBox lstMacros 
      Enabled         =   0   'False
      Height          =   315
      Index           =   8
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3360
      Width           =   1935
   End
   Begin VB.ComboBox lstMacros 
      Enabled         =   0   'False
      Height          =   315
      Index           =   7
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3000
      Width           =   1935
   End
   Begin VB.ComboBox lstMacros 
      Enabled         =   0   'False
      Height          =   315
      Index           =   6
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2640
      Width           =   1935
   End
   Begin VB.ComboBox lstMacros 
      Enabled         =   0   'False
      Height          =   315
      Index           =   5
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.ComboBox lstMacros 
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1920
      Width           =   1935
   End
   Begin VB.ComboBox lstMacros 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1560
      Width           =   1935
   End
   Begin VB.ComboBox lstMacros 
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox lstMacros 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   840
      Width           =   1935
   End
   Begin VB.ComboBox lstMacros 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   480
      Width           =   1935
   End
   Begin VB.CheckBox chkButton 
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CheckBox chkButton 
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox chkButton 
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CheckBox chkButton 
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox chkButton 
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CheckBox chkButton 
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox chkButton 
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CheckBox chkButton 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox chkButton 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox chkButton 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3480
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label2 
      Caption         =   "Assigned Macro"
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enable Shortcut"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3480
      Y1              =   4200
      Y2              =   4200
   End
End
Attribute VB_Name = "frmMacroButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butCancel_Click()

Me.Hide

End Sub

Private Sub butOK_Click()

SetButtons

frmMacroButtons.Hide

End Sub

Public Sub SetButtons()

Dim count As Integer

For count = 0 To 9
    If chkButton(count).value = 1 And lstMacros(count).ListIndex >= 0 Then
        frmMain.ConfigMacroButton count, lstMacros(count).ListIndex
    Else
        frmMain.ConfigMacroButton count, -1
    End If
Next count

End Sub

Private Sub chkButton_Click(index As Integer)

If chkButton(index).value = 1 Then
    lstMacros(index).Enabled = True
    PopulateList index
Else
    lstMacros(index).Enabled = False
    lstMacros(index).Clear
End If

End Sub

Public Sub PopulateList(index As Integer)

Dim count As Integer
Dim numMacros As Integer

numMacros = frmMain.lstMacroName.ListCount

lstMacros(index).Clear
For count = 0 To (numMacros - 1)
    lstMacros(index).AddItem frmMain.lstMacroName.List(count)
Next count

End Sub

Private Sub Form_Load()

SetPos Me

Dim count As Integer

For count = 0 To 9
    chkButton(count).Caption = "F" + Str$(count + 1)
Next count

End Sub
