VERSION 5.00
Begin VB.Form frmNote 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Note Writer"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optNoteMode 
      Caption         =   "Note mode"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton optMailMode 
      Caption         =   "Mail mode"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox txtCCTo 
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txtMailTo 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtNoteTitle 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   6495
   End
   Begin VB.CommandButton butClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton butPreview 
      Caption         =   "P&review"
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton butPost 
      Caption         =   "&Post"
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton butUndo 
      Caption         =   "&Undo Formatting"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtWrap 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   7
      Text            =   "78"
      Top             =   5025
      Width           =   375
   End
   Begin VB.TextBox txtNote 
      Height          =   3495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1440
      Width           =   8775
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8880
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblCCTo 
      Caption         =   "CC:"
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblMailTo 
      Caption         =   "Mail to:"
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Note Contents:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblNoteTitle 
      Caption         =   "Note Title:"
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   600
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8880
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "characters. (note: Most games work on an 80 character-wide screen. Set to 78 if this is the case)"
      Height          =   195
      Left            =   1560
      TabIndex        =   12
      Top             =   5040
      Width           =   6825
   End
   Begin VB.Label Label1 
      Caption         =   "Wrap text at"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   975
   End
End
Attribute VB_Name = "frmNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastStr As String
Dim endOfNote As String
Dim newNoteCmd As String
Dim newMailCmd As String

Private Sub butClose_Click()

Me.Hide

End Sub

Private Sub butPost_Click()

If optMailMode.value Then
    PostMail
Else
    PostNote
End If

End Sub
Private Sub PostMail()

If txtMailTo.Text = "" Then
    MsgBox "You need to specify at least one recipient.", vbApplicationModal + vbCritical + vbOKOnly, "Note Writer"
    txtMailTo.SetFocus
    Exit Sub
End If

If ChkNoteValues() = True Then
    FormatNote
    frmMain.SocketSend newMailCmd & " " & txtMailTo.Text & vbCrLf & txtNote.Text & vbCrLf & endOfNote & vbCrLf & txtCCTo.Text & vbCrLf
Else
    MsgBox "You need to set the Note Writer commands" & vbCrLf & "in the Preferences section.", vbApplicationModal + vbCritical + vbOKOnly, "Note Writer"
    frmOptions.Show
End If

End Sub

Private Sub PostNote()

If txtNoteTitle.Text = "" Then
    MsgBox "You need to set a Note Title.", vbApplicationModal + vbCritical + vbOKOnly, "Note Writer"
    txtNoteTitle.SetFocus
    Exit Sub
End If

If ChkNoteValues() = True Then
    FormatNote
    frmMain.SocketSend newNoteCmd & " " & txtNoteTitle.Text & vbCrLf & txtNote.Text + vbCrLf + endOfNote + vbCrLf
Else
    MsgBox "You need to set the Note Writer commands" & vbCrLf & "in the Preferences section.", vbApplicationModal + vbCritical + vbOKOnly, "Note Writer"
    frmOptions.Show
End If

End Sub

Private Function ChkNoteValues()

If endOfNote <> "" And newNoteCmd <> "" And newMailCmd <> "" Then
    ChkNoteValues = True
Else
    ChkNoteValues = False
End If

End Function

Private Sub FormatNote()

Dim tempStr As String
Dim lenText As Integer
Dim count As Integer
Dim lastPos As Integer
Dim wrapAt As Integer
Dim countTwo As Integer

lenText = Len(txtNote.Text)
lastPos = 0
wrapAt = Val(txtWrap.Text)
lastStr = txtNote.Text

For count = 1 To lenText
    tempStr = Mid$(txtNote.Text, count, 3)
    If tempStr = ".  " Then
        txtNote.SelStart = count - 1
        txtNote.SelLength = 3
        txtNote.SelText = ". "
    End If
    tempStr = Mid$(txtNote.Text, count, 1)
    If tempStr = vbLf Then lastPos = count
    If count > lastPos + (wrapAt - 15) And count < lastPos + wrapAt And tempStr = " " Then
        txtNote.SelStart = count - 1
        txtNote.SelLength = 1
        txtNote.SelText = vbCrLf
        lastPos = count
        End If
    If count >= lastPos + wrapAt Then
        countTwo = lastPos
        Do While (countTwo > lastPos And countTwo > 0)
            countTwo = countTwo - 1
            If Mid$(txtNote.Text, countTwo, 1) = " " Then
                txtNote.SelStart = count - 1
                txtNote.SelLength = 1
                txtNote.SelText = vbCrLf
                lastPos = countTwo
                countTwo = -1
            End If
        Loop
        If countTwo <> -1 Then
            lastPos = count
        End If
    End If
Next count

End Sub

Private Sub butPreview_Click()

FormatNote

End Sub

Private Sub butUndo_Click()

If lastStr <> "" Then
    txtNote.Text = lastStr
Else
    MsgBox "Note not yet formatted. No undo information exists.", vbApplicationModal + vbOKCancel + vbExclamation, "Note Writer"
End If

End Sub

Private Sub Form_Load()

SetPos Me
endOfNote = frmOptions.txtEndOfNote.Text
newNoteCmd = frmOptions.txtNewNoteCmd.Text
newMailCmd = frmOptions.txtNewMailCommand.Text
txtNote.BackColor = frmMain.txtDisplay.BackColor
txtNote.ForeColor = frmMain.txtDisplay.SelColor
txtNote.Font = frmMain.txtDisplay.Font
lastStr = ""
SetMode (optMailMode.value)

End Sub

Private Sub SetMode(MailMode As Boolean)

txtNoteTitle.Enabled = Not (MailMode)
txtMailTo.Enabled = MailMode
txtCCTo.Enabled = MailMode
lblMailTo.Enabled = MailMode
lblCCTo.Enabled = MailMode
lblNoteTitle.Enabled = Not (MailMode)

End Sub


Private Sub optMailMode_Click()

SetMode (optMailMode.value)

End Sub

Private Sub optNoteMode_Click()

SetMode (optMailMode.value)

End Sub

Private Sub txtWrap_LostFocus()

If txtWrap.Text <> Format(Int(Val(txtWrap.Text))) Or Val(txtWrap.Text) < 25 Then
    MsgBox "Value must be an integer between 25 and 99", vbOKOnly + vbApplicationModal + vbCritical, "Note Writer"
    txtWrap.SetFocus
    txtWrap.SelStart = 0
    txtWrap.SelLength = Len(txtWrap.Text)
End If

End Sub
