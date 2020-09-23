VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preferences..."
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton butOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame fraDisplay 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   4935
      Begin VB.CheckBox chkDoubleCrLf 
         Caption         =   "Hide double line-breaks"
         Height          =   255
         Left            =   2760
         TabIndex        =   40
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtScrollback 
         Height          =   285
         Left            =   2400
         TabIndex        =   30
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CheckBox chkLocalEcho 
         Caption         =   "Echo sent commands"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CommandButton butSetFont 
         Caption         =   "Change"
         Height          =   255
         Left            =   3840
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton butSetFore 
         Caption         =   "Change"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton butSetBack 
         Caption         =   "Change"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin MSComDlg.CommonDialog dlgFont 
         Left            =   120
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Caption         =   "Scrollback size (characters):"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Display font:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblDisplayFont 
         Caption         =   " "
         Height          =   195
         Left            =   1200
         TabIndex        =   21
         Top             =   720
         Width           =   2445
      End
      Begin VB.Label Label4 
         Caption         =   "Display foreground:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Shape boxForegrnd 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2640
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Display Background:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Shape boxBackgrnd 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2640
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame fraAutomation 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   5055
      Begin VB.CheckBox chkAutoLogin 
         Caption         =   "Enable Automatic Login"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox txtPasswordPrompt 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtUsernamePrompt 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblPass 
         Caption         =   "Password Prompt:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblUser 
         Caption         =   "Username Prompt:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame fraMisc 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   5175
      Begin VB.CheckBox chkFlashWindow 
         Caption         =   "Flash taskbar icon when minimised"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox chkSplash 
         Caption         =   "Show Splash Screen at Start Up"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtQuitCommand 
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   """Quit"" command:"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   375
         Width           =   1335
      End
   End
   Begin VB.Frame fraNoteWriter 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   5055
      Begin VB.TextBox txtNewMailCommand 
         Height          =   285
         Left            =   2160
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtNewNoteCmd 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtEndOfNote 
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "New Mail command:"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   735
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Note Creation command:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   375
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "End-Of-Note string:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1095
         Width           =   1455
      End
   End
   Begin VB.Frame fraKeepAlive 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      TabIndex        =   32
      Top             =   480
      Width           =   5175
      Begin VB.TextBox txtKeepAliveTime 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   38
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox txtKeepAliveStr 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox chkKeepAlive 
         Caption         =   "Send Keep Alive"
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblKAM 
         Caption         =   "minute(s)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3840
         TabIndex        =   39
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblKAI 
         Caption         =   "Keep Alive interval:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblKAC 
         Caption         =   "Keep Alive command:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   $"frmOptions.frx":0000
         Height          =   855
         Left            =   0
         TabIndex        =   34
         Top             =   120
         Width           =   4935
      End
   End
   Begin MSComctlLib.TabStrip tabPreferences 
      Height          =   2655
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4683
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Display"
            Object.ToolTipText     =   "Change screen colour, default text colour and font and scrollback size"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Auto-Login"
            Object.ToolTipText     =   "Set the text triggers for the auto-login feature"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Note Writer"
            Object.ToolTipText     =   "Set the commands used by the game server's note-writing and mail systems"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Miscellaneous"
            Object.ToolTipText     =   "Set application behaviour options and define the Quit command"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Keep Alive"
            Object.ToolTipText     =   "Configure the Keep-Alive function"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5280
      Y1              =   2880
      Y2              =   2880
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currentFrame As Integer

Private Sub butCancel_Click()

frmOptions.Hide

End Sub

Private Sub butOK_Click()

If dlgFont.FontName <> "" Then
    frmMain.txtDisplay.SelFontName = dlgFont.FontName
    frmMain.txtEntry.Font.Name = dlgFont.FontName
    frmMain.txtDisplay.SelFontSize = dlgFont.FontSize
    frmMain.txtEntry.Font.Size = dlgFont.FontSize
    frmMain.txtDisplay.SelBold = dlgFont.FontBold
    frmMain.txtEntry.Font.Bold = dlgFont.FontBold
    frmMain.txtDisplay.SelItalic = dlgFont.FontItalic
    frmMain.txtEntry.Font.Italic = dlgFont.FontItalic
End If
frmMain.SetPrefs
frmOptions.Hide

SetKeepAliveTimer

End Sub

Private Sub butSetBack_Click()

On Error GoTo ErrHandler
dlgFont.CancelError = True
dlgFont.ShowColor
boxBackgrnd.FillColor = dlgFont.Color

ErrHandler:
    Exit Sub


End Sub

Private Sub butSetFont_Click()

On Error GoTo ErrHandler

dlgFont.CancelError = True
dlgFont.Flags = cdlCFScreenFonts Or cdlCFANSIOnly
dlgFont.FontName = frmMain.txtDisplay.SelFontName
dlgFont.FontBold = frmMain.txtDisplay.SelBold
dlgFont.FontItalic = frmMain.txtDisplay.SelItalic
dlgFont.FontSize = frmMain.txtDisplay.SelFontSize
dlgFont.ShowFont
lblDisplayFont.Caption = dlgFont.FontName + " " + Format(dlgFont.FontSize) + " point"

ErrHandler:
  Exit Sub

End Sub

Private Sub butSetFore_Click()

On Error GoTo ErrHandler
dlgFont.CancelError = True
dlgFont.ShowColor
boxForegrnd.FillColor = dlgFont.Color

ErrHandler:
    Exit Sub

End Sub

Public Sub SetDisplays()

SetAutomation
SetKeepAlive

End Sub

Private Sub SetAutomation()

Dim chkState As Boolean

If chkAutoLogin.value = 1 Then
    chkState = True
Else
    chkState = False
End If

lblUser.Enabled = chkState
lblPass.Enabled = chkState
txtPasswordPrompt.Enabled = chkState
txtUsernamePrompt.Enabled = chkState

End Sub

Public Sub SetKeepAliveTimer()

If chkKeepAlive.value = 1 Then
    frmMain.SetTick Val(txtKeepAliveTime.Text) * 6
    frmMain.tmrKeepAlive.Enabled = True
Else
    frmMain.tmrKeepAlive.Enabled = False
End If

End Sub

Private Sub chkAutoLogin_Click()

SetAutomation

End Sub

Private Sub chkKeepAlive_Click()

SetKeepAlive

End Sub
Private Sub SetKeepAlive()

Dim KeepAliveEnable As Boolean

KeepAliveEnable = False
If chkKeepAlive.value = 1 Then KeepAliveEnable = True

txtKeepAliveStr.Enabled = KeepAliveEnable
txtKeepAliveTime.Enabled = KeepAliveEnable
lblKAC.Enabled = KeepAliveEnable
lblKAI.Enabled = KeepAliveEnable
lblKAM.Enabled = KeepAliveEnable

End Sub
Private Sub Form_Load()

SetPos Me
currentFrame = 0
ChangeTab 1

End Sub

Private Sub tabPreferences_Click()

ChangeTab tabPreferences.SelectedItem.index

End Sub

Private Sub ChangeTab(newFrame As Integer)

If newFrame = currentFrame Then Exit Sub

fraDisplay.Visible = False
fraAutomation.Visible = False
fraNoteWriter.Visible = False
fraMisc.Visible = False
fraKeepAlive.Visible = False

Select Case newFrame
    Case 1
        fraDisplay.Visible = True
    Case 2
        fraAutomation.Visible = True
    Case 3
        fraNoteWriter.Visible = True
    Case 4
        fraMisc.Visible = True
    Case 5
        fraKeepAlive.Visible = True
End Select

currentFrame = newFrame

End Sub

Private Sub txtKeepAliveTime_KeyPress(KeyAscii As Integer)

KeyAscii = KeyFilter(KeyAscii)

End Sub

Private Sub txtScrollback_KeyPress(KeyAscii As Integer)

KeyAscii = KeyFilter(KeyAscii)

End Sub
