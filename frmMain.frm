VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Sanx's MUD Client"
   ClientHeight    =   4395
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9840
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstColours 
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox lstColourTrig 
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.ImageList imgGrey 
      Left            =   7680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":350C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5874
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":614E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   8280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7302
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":84B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":966A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A81E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B0F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B9D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C2AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolMain 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imgGrey"
      HotImageList    =   "imgMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "connect"
            Description     =   "Connect"
            Object.ToolTipText     =   "Connect"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "disconnect"
            Description     =   "Disconnect"
            Object.ToolTipText     =   "Disconnect"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "quit"
            Description     =   "Send Quit"
            Object.ToolTipText     =   "Send Quit"
            Object.Tag             =   "quit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copypaste"
            Description     =   "Copy and Paste"
            Object.ToolTipText     =   "Copy and Paste"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copyclip"
            Description     =   "Copy Clipboard to Game"
            Object.ToolTipText     =   "Copy Clipboard to Game"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "note"
            Description     =   "Note Writer"
            Object.ToolTipText     =   "Note Writer"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "macros"
            Description     =   "Edit Macros"
            Object.ToolTipText     =   "Edit Macros"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "colours"
            Description     =   "Colour Highlighting"
            Object.ToolTipText     =   "Colour Highlighting"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "history"
            Description     =   "History"
            Object.ToolTipText     =   "History"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "preferences"
            Description     =   "Preferences"
            Object.ToolTipText     =   "Preferences"
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin VB.Timer tmrKeepAlive 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   7080
         Top             =   120
      End
   End
   Begin VB.ListBox lstMacroName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      ItemData        =   "frmMain.frx":CB86
      Left            =   60
      List            =   "frmMain.frx":CB88
      TabIndex        =   9
      Top             =   600
      Width           =   1765
   End
   Begin VB.CommandButton butShortcut 
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton butShortcut 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton butShortcut 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton butShortcut 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton butShortcut 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton butShortcut 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton butShortcut 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton butShortcut 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton butShortcut 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton butShortcut 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtErrorLog 
      Height          =   285
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmMain.frx":CB8A
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   7560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Sanx's MUD Client - History"
      Filter          =   "*.txt"
   End
   Begin VB.ListBox lstTriggerCommand 
      Height          =   255
      ItemData        =   "frmMain.frx":CBA4
      Left            =   2040
      List            =   "frmMain.frx":CBA6
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox lstTrigger 
      Height          =   255
      ItemData        =   "frmMain.frx":CBA8
      Left            =   2040
      List            =   "frmMain.frx":CBAA
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtTemp 
      Height          =   285
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox lstMacroCommand 
      Height          =   255
      ItemData        =   "frmMain.frx":CBAC
      Left            =   2040
      List            =   "frmMain.frx":CBAE
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox lstMacro 
      Height          =   255
      ItemData        =   "frmMain.frx":CBB0
      Left            =   2040
      List            =   "frmMain.frx":CBB2
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox lstEntry 
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   9240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1860
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3390
      Width           =   7935
   End
   Begin MSComctlLib.StatusBar barStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4140
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   450
      SimpleText      =   "Sanx's MUD Client"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1138
            MinWidth        =   1147
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1138
            MinWidth        =   1147
            TextSave        =   "23:16"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Key             =   "trigger"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Key             =   "colours"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   6191
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "Sanx's MUD Client"
            TextSave        =   "Sanx's MUD Client"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtDisplay 
      Height          =   2655
      Left            =   1860
      TabIndex        =   21
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4683
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":CBB4
   End
   Begin VB.CommandButton butUseless 
      Caption         =   "Bollocks"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   24
      Top             =   720
      Width           =   855
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connection"
      Begin VB.Menu mnuConnectMud 
         Caption         =   "Co&nnect"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuSendQuit 
         Caption         =   "&Send Quit"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "&Edit"
      Begin VB.Menu mnuPopCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPopCopyPaste 
         Caption         =   "Copy and &Paste"
      End
      Begin VB.Menu mnuEditClipGame 
         Caption         =   "Copy Clipboard to &Game"
      End
      Begin VB.Menu mnuPopQuote 
         Caption         =   "Copy and Paste &Quoted"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBlank6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintBuffer 
         Caption         =   "Print &Buffer"
      End
      Begin VB.Menu mnuClearBuffer 
         Caption         =   "C&lear Buffer"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuBlank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoteWriter 
         Caption         =   "&Note Writer"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuPreferences 
         Caption         =   "&Preferences"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuMacroSub 
         Caption         =   "&Macros..."
         Begin VB.Menu mnuNewMacro 
            Caption         =   "&New Macro..."
         End
         Begin VB.Menu mnuMacro 
            Caption         =   "&Edit Macros"
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuRecordMacro 
            Caption         =   "&Record New Macro"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuMacroButton 
            Caption         =   "&Macro Shortcuts"
         End
         Begin VB.Menu mnuMacroSep 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuMacroList 
            Caption         =   ""
            Index           =   0
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuTrigger 
         Caption         =   "&Triggers"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuColours 
         Caption         =   "&Colours"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuEnableTriggers 
         Caption         =   "&Enable Trigger Processing"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEnableColours 
         Caption         =   "E&nable Colour Processing"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowSndControls 
         Caption         =   "Show Sound Co&ntrols"
      End
      Begin VB.Menu mnuStopSound 
         Caption         =   "Stop So&und"
      End
      Begin VB.Menu mnuBlank4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveSettings 
         Caption         =   "&Save Settings"
      End
      Begin VB.Menu mnuEnableDebug 
         Caption         =   "Enable Error &Logging"
      End
   End
   Begin VB.Menu mnuHistory 
      Caption         =   "&History"
      Begin VB.Menu mnuSaveHistory 
         Caption         =   "&Save History"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuViewHistory 
         Caption         =   "&View Saved History"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuHelp 
         Caption         =   "&Help"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAboutShow 
         Caption         =   "&About Sanx's MUD Client"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Display Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuPopUpCopyPaste 
         Caption         =   "Copy and Paste"
      End
      Begin VB.Menu mnuPopSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuPopSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopNewTrigger 
         Caption         =   "New Trigger"
      End
      Begin VB.Menu mnuPopNewColourHilite 
         Caption         =   "New Colour Highlight"
      End
   End
   Begin VB.Menu mnuMacroPopUp 
      Caption         =   "Macro Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopEditMacro 
         Caption         =   "Edit Macro"
      End
      Begin VB.Menu mnuPopRemoveMacro 
         Caption         =   "Remove Macro"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This application was written and is copyright Sanx, 2000-2001
'This application is freeware and may be freely modified, copied and distributed
'provided that this copyright notice remains.
'http://www.sanx.org/

Dim LocalEcho As Boolean, AutoLogin As Boolean, DebugMode As Boolean
Dim ProcessTriggers As Boolean, ProcessColours As Boolean, IsRecording As Boolean
Dim FlashWindowStatus As Boolean, DoneAutoLogin As Boolean
Dim CacheCount As Integer, MacroButtonCommand(9) As Integer, siTick As Integer, intTick
Dim RemHost As String, RemPort As String, RemUser As String, RemPass As String
Dim RunPath As String, QuitCommand As String, MacroFile As String, TriggerFile As String
Dim ColourFile As String, LastWorld As String, htxt As String
Dim AppMsgColour As Long, Scrollback As Long

Private MydsEncrypt As dsEncrypt

Private Sub txtDisplay_Click()
    If txtDisplay.SelText = "" Then txtEntry.SetFocus
End Sub

Public Sub UpdateAppMsgColour(sColour As Long)

AppMsgColour = sColour

End Sub

Public Sub SetTick(ticktime As Integer)

intTick = ticktime

End Sub

Public Sub ClearShortcut(index As Integer)

Dim count As Integer

For count = 0 To 9
    If MacroButtonCommand(count) = index Then
        butShortcut(count).Caption = Format(count + 1)
        frmMacroButtons.chkButton(count).value = 0
        frmMacroButtons.lstMacros(count).Enabled = False
        frmMacroButtons.lstMacros(count).Clear
        ConfigMacroButton count, -1
    End If
Next
    
End Sub

Public Sub ConfigMacroButton(Button As Integer, macroIndex As Integer)

MacroButtonCommand(Button) = macroIndex

End Sub

Public Sub SetRecording(value As Boolean)

IsRecording = value

End Sub

Private Sub barStatus_PanelDblClick(ByVal Panel As MSComctlLib.Panel)

Select Case Left$(Panel.Text, 4)
    Case "Trig"
        TriggerProcessChange
    Case "Colo"
        ColourProcessChange
    Case "Sanx"
        frmAbout.Show vbModal, Me
End Select

End Sub

Private Sub butShortcut_Click(index As Integer)

Dim butIndex As Integer

butIndex = MacroButtonCommand(index)

If butIndex >= 0 Then
    SendText lstMacro.List(butIndex)
End If

txtEntry.SetFocus

End Sub

Public Sub LogError(ErrText As String)

txtErrorLog.Text = txtErrorLog.Text & ErrText & vbCrLf

End Sub

Private Sub CacheEntry(sendStr As String)

Dim count As Integer

If lstEntry.ListCount < 100 Then
    lstEntry.AddItem sendStr
Else
    For count = 0 To 9
        lstEntry.RemoveItem count
    Next count
    lstEntry.AddItem sendStr
End If

End Sub

Public Sub SendText(sendStr As String)

Dim soundfile As String

If IsRecording Then
    frmRecordMacros.txtCmdSeq.Text = frmRecordMacros.txtCmdSeq.Text + sendStr + ";"
End If

If Left$(sendStr, 1) = "/" Then
    SendMacro (sendStr)
    CacheEntry sendStr
    txtEntry.Text = ""
    Exit Sub
End If

If sckMain.State = sckConnected Then
    SocketSend sendStr + vbCrLf
    If sendStr <> "" Then
        CacheEntry sendStr
    End If
    If LocalEcho Then
        DisplayText sendStr + vbCrLf, frmOptions.boxForegrnd.FillColor, frmOptions.boxForegrnd.FillColor
    End If
    CacheCount = lstEntry.ListCount
ElseIf sckMain.State <> sckConnected Then
    DisplayInfo "Not Connected"
End If

End Sub

Public Sub SocketSend(sendStr As String)

If sckMain.State = sckConnected Then
    sckMain.SendData sendStr
ElseIf sckMain.State <> sckConnected Then
    DisplayInfo "Not Connected"
End If

End Sub

Private Function GetLeft(inputText As String) As String

Dim num As Integer, leftStr As String

If Len(inputText) <= 1 Then Exit Function

    For num = 1 To Len(inputText)
        leftStr = Left$(inputText, Len(inputText) - num)


        If Right$(leftStr, 1) = Chr$(32) Then
            GetLeft = Left$(leftStr, Len(leftStr) - 1)
            Exit Function
        End If
    Next num
    GetLeft = inputText

End Function

Private Function GetRight(inputText As String) As String

Dim num As Integer, rightStr As String

If Len(inputText) <= 1 Then Exit Function

    For num = 1 To Len(inputText)
        rightStr = Right$(inputText, Len(inputText) - num)


        If Left$(rightStr, 1) = Chr$(32) Then
            GetRight = Right$(rightStr, Len(rightStr) - 1)
            Exit Function
        End If
    Next num
    GetRight = ""

End Function

Private Sub SendMacro(sendStr As String)

Dim count As Integer
Dim foundFlag As Boolean
Dim macroStr As String

Select Case sendStr
    Case "/debug"
        If DebugMode Then
            DisplayText txtErrorLog.Text, AppMsgColour, frmOptions.boxForegrnd.FillColor
        Else
            DisplayInfo "Error Logging Mode Not Enabled"
        End If
        
    Case "/clear"
        txtDisplay.TextRTF = ""
        
    Case Else
        foundFlag = False

        For count = 0 To lstMacro.ListCount
            If GetLeft(sendStr) = lstMacro.List(count) Then
                foundFlag = True
                Exit For
            End If
        Next

        If foundFlag Then
            macroStr = lstMacroCommand.List(count)
            txtTemp.Text = ReplaceStr(macroStr, ";", vbCrLf)
            If GetRight(sendStr) <> "" Then txtTemp.Text = ReplaceStr(txtTemp.Text, "%1", GetRight(sendStr))
            SendText txtTemp.Text
            CacheEntry sendStr
            CacheCount = lstEntry.ListCount
            txtEntry.Text = ""
        Else
            DisplayInfo "No Such Macro Defined"
        End If
        
End Select


End Sub

Private Function ReplaceStr(tempStr As String, findStr As String, repStr As String)

txtTemp.Text = ExtractSoundCommand(tempStr)
ReplaceStr = Replace(txtTemp.Text, findStr, repStr)

End Function

Private Function ExtractSoundCommand(tempStr As String)

Dim where As Integer, startPoint As Integer, lenstr As Integer
Dim sndFile As String, findStr As String, endstr As String
Dim allDone As Boolean

txtTemp.Text = tempStr
findStr = "**PLAYSND**"
endstr = "+++"
lenstr = Len(findStr)
allDone = False
startPoint = 1

Do
    where = InStr(startPoint, txtTemp.Text, findStr, vbTextCompare)
    If where Then
        txtTemp.SelStart = where - 1
        txtTemp.SelLength = lenstr
        txtTemp.SelText = ""
        startPoint = where
        where = InStr(startPoint, txtTemp.Text, endstr, vbTextCompare)
        txtTemp.SelStart = where
        txtTemp.SelLength = 3
        txtTemp.SelText = ""
        sndFile = Mid$(txtTemp.Text, startPoint, where - startPoint)
        frmSndControl.sndMain.FileName = sndFile
        frmSndControl.sndMain.Play
        txtTemp.SelStart = startPoint - 1
        txtTemp.SelLength = (where - startPoint) + 1
        txtTemp.SelText = ""
    Else
        allDone = True
    End If
Loop Until allDone = True

ExtractSoundCommand = txtTemp.Text

End Function

Public Sub SetConnValues()

With frmConnect
    .txtHost.Text = RemHost
    .txtPort.Text = RemPort
    .txtUsername.Text = RemUser
    .txtPassword.Text = RemPass
End With

End Sub

Private Sub butUseless_Click()

SendText txtEntry.Text
txtEntry.Text = ""
txtEntry.SetFocus
DoEvents

End Sub

Private Sub Form_Load()

DebugMode = False
IsRecording = False
Set MydsEncrypt = New dsEncrypt
MydsEncrypt.KeyString = ("I love Katie")

RunPath = App.Path

frmOptions.chkSplash.value = GetSetting(App.Title, "settings", "ShowSplash", 1)
If frmOptions.chkSplash.value > 0 Then frmSplash.Show vbModal

Dim count As Integer

For count = 0 To 9
    MacroButtonCommand(count) = -1
Next count

Scrollback = GetSetting(App.Title, "Settings", "Scrollback", 100000)
frmOptions.txtScrollback.Text = Str$(Scrollback)

With Me
    .Width = GetSetting(App.Title, "Settings", "MainWidth", 9000)
    .Height = GetSetting(App.Title, "Settings", "MainHeight", 5000)
    .Left = GetSetting(App.Title, "Settings", "MainLeft", ((Screen.Width / 2) - (Me.Width / 2)))
    .Top = GetSetting(App.Title, "Settings", "MainTop", ((Screen.Height / 2) - (Me.Height / 2)))
    .WindowState = GetSetting(App.Title, "Settings", "WindowState", vbNormal)
End With

AppMsgColour = GetSetting(App.Title, "settings", "AppMsgColour", RGB(255, 0, 0))
frmColours.shaAppMsg.FillColor = AppMsgColour

frmMacro.grdMacro.ColWidth(0) = GetSetting(App.Title, "Settings", "MacroNameWidth", 1000)
frmMacro.grdMacro.ColWidth(1) = GetSetting(App.Title, "Settings", "MacroTriggerWidth", 1000)
frmMacro.grdMacro.ColWidth(2) = GetSetting(App.Title, "Settings", "MacroResponseWidth", 1000)
frmTrigger.grdTrigger.ColWidth(0) = GetSetting(App.Title, "Settings", "TriggerTriggerWidth", 1000)
frmTrigger.grdTrigger.ColWidth(1) = frmTrigger.grdTrigger.Width - (frmTrigger.grdTrigger.ColWidth(0) + 100)

frmOptions.SetDisplays

End Sub

Public Sub LoadGameSettings(worldName As String)

Dim count As Integer

If LastWorld <> "" Then SaveSettings LastWorld

LastWorld = worldName

LocalEcho = GetSetting(App.Title, worldName, "LocalEcho", True)
If LocalEcho Then
    frmOptions.chkLocalEcho.value = 1
Else
    frmOptions.chkLocalEcho.value = 0
End If

AutoLogin = GetSetting(App.Title, worldName, "AutoLogin", True)
If AutoLogin Then
    frmOptions.chkAutoLogin.value = 1
Else
    frmOptions.chkAutoLogin.value = 0
End If

FlashWindowStatus = GetSetting(App.Title, worldName, "FlashWindow", True)
If FlashWindowStatus Then
    frmOptions.chkFlashWindow.value = 1
Else
    frmOptions.chkFlashWindow.value = 0
End If

RemHost = GetSetting(App.Title, worldName, "Host", "")
RemPort = GetSetting(App.Title, worldName, "Port", "")
RemUser = GetSetting(App.Title, worldName, "Username", "")
RemPass = MydsEncrypt.Encrypt(GetSetting(App.Title, worldName, "Password", ""))
QuitCommand = GetSetting(App.Title, worldName, "QuitCommand", "quit")
MacroFile = GetSetting(App.Title, worldName, "MacroFile", RunPath & "\" & worldName & ".mac")
TriggerFile = GetSetting(App.Title, worldName, "TriggerFile", RunPath & "\" & worldName & ".tri")
ColourFile = GetSetting(App.Title, worldName, "ColourFile", RunPath & "\" & worldName & ".col")

With txtDisplay
    .SelColor = GetSetting(App.Title, worldName, "ForeColour", &H0&)
    .BackColor = GetSetting(App.Title, worldName, "BackColour", &HFFFFFF)
    .SelFontName = GetSetting(App.Title, worldName, "FontName", "Courier New")
    .SelFontSize = GetSetting(App.Title, worldName, "FontSize", 8.25)
End With

With txtEntry
    .FontName = GetSetting(App.Title, worldName, "FontName", "Courier New")
    .FontSize = GetSetting(App.Title, worldName, "FontSize", 8.25)
    .BackColor = GetSetting(App.Title, worldName, "Amanda", &HFFFFFF)
End With

frmOptions.SetDisplays
lstMacroName.BackColor = txtEntry.BackColor

With frmOptions
    .boxForegrnd.FillColor = GetSetting(App.Title, worldName, "ForeColour", &H0&)
    .boxBackgrnd.FillColor = GetSetting(App.Title, worldName, "BackColour", &HFFFFFF)
    .dlgFont.FontName = GetSetting(App.Title, worldName, "FontName", "Courier New")
    .dlgFont.FontSize = GetSetting(App.Title, worldName, "FontSize", 8.25)
    .lblDisplayFont.Caption = .dlgFont.FontName + " " + Format(.dlgFont.FontSize) + " point"
    .txtUsernamePrompt.Text = GetSetting(App.Title, worldName, "UserPrompt", "is your name:")
    .txtPasswordPrompt.Text = GetSetting(App.Title, worldName, "PassPrompt", "assword:")
    .txtEndOfNote.Text = GetSetting(App.Title, worldName, "EndOfNote", "**")
    .txtNewNoteCmd.Text = GetSetting(App.Title, worldName, "NewNoteCmd", "note")
    .txtNewMailCommand = GetSetting(App.Title, worldName, "NewMailCmd", "mail")
    .txtQuitCommand.Text = QuitCommand
    .txtKeepAliveStr.Text = GetSetting(App.Title, worldName, "KeepAliveString", "")
    .txtKeepAliveTime.Text = GetSetting(App.Title, worldName, "KeepAliveInterval", "")
    .chkKeepAlive.value = GetSetting(App.Title, worldName, "KeepAliveEnable", 0)
    .chkDoubleCrLf.value = GetSetting(App.Title, worldName, "DoubleCrLf", 0)
    .SetKeepAliveTimer
End With

With frmNote
    .optMailMode.value = GetSetting(App.Title, worldName, "NoteWriterMode", True)
    .optNoteMode.value = Not (frmNote.optMailMode.value)
    .txtWrap.Text = GetSetting(App.Title, worldName, "NoteWriterWidth", "78")
End With

LoadMacros MacroFile
LoadTriggers TriggerFile
LoadColours ColourFile

ProcessTriggers = GetSetting(App.Title, worldName, "ProcessTriggers", True)
mnuEnableTriggers.Checked = ProcessTriggers
If ProcessTriggers Then
    barStatus.Panels(3).Text = "Triggers: ON"
Else
    barStatus.Panels(3).Text = "Triggers: OFF"
End If

ProcessColours = GetSetting(App.Title, worldName, "ProcessColours", True)
mnuEnableColours.Checked = ProcessColours
If ProcessColours Then
    barStatus.Panels(4).Text = "Colours: ON"
Else
    barStatus.Panels(4).Text = "Colours: OFF"
End If

For count = 0 To 9
    MacroButtonCommand(count) = GetSetting(App.Title, worldName, "Button" + Format(count) + "Command", -1)
    If MacroButtonCommand(count) >= 0 And MacroButtonCommand(count) < lstMacro.ListCount Then
        frmMacroButtons.chkButton(count).value = 1
        frmMacroButtons.lstMacros(count).Enabled = True
        frmMacroButtons.PopulateList count
        frmMacroButtons.lstMacros(count).ListIndex = MacroButtonCommand(count)
    End If
Next count

frmMacroButtons.SetButtons

End Sub

Public Sub LoadMacros(filepath As String)

Dim tempStr As String

On Error GoTo ErrHandler

lstMacroName.Clear
lstMacro.Clear
lstMacroCommand.Clear

MacroFile = filepath

Open filepath For Input As #1

Do While Not EOF(1)
    Line Input #1, tempStr
    lstMacroName.AddItem tempStr
    Line Input #1, tempStr
    lstMacro.AddItem tempStr
    Line Input #1, tempStr
    lstMacroCommand.AddItem tempStr
Loop

Close #1

DisplayInfo "Macro File Loaded: " & filepath

Exit Sub

ErrHandler:
    DisplayInfo "Macro File Not Loaded: " & filepath
    Close #1

End Sub

Public Sub LoadColours(filepath As String)

Dim tempStr As String

On Error GoTo ErrHandler

lstColourTrig.Clear
lstColours.Clear

ColourFile = filepath

Open filepath For Input As #1

Do While Not EOF(1)
    Line Input #1, tempStr
    lstColourTrig.AddItem tempStr
    Line Input #1, tempStr
    lstColours.AddItem tempStr
Loop

Close #1

DisplayInfo "Colour File Loaded: " & filepath

Exit Sub

ErrHandler:
    DisplayInfo "Colour File Not Loaded: " & filepath
    Close #1

End Sub

Public Sub MacroMenu()

Dim count As Integer

If lstMacro.ListCount > 0 Then
    mnuMacroSep.Visible = False
    For count = (mnuMacroList.count - 1) To 1 Step -1
        Unload mnuMacroList(count)
    Next count
    For count = 0 To (lstMacro.ListCount - 1)
        If count > 0 Then Load mnuMacroList(count)
        mnuMacroList(count).Caption = lstMacroName.List(count)
        mnuMacroList(count).Visible = False
    Next count
Else
    mnuMacroSep.Visible = False
    mnuMacroList(0).Visible = False
    For count = (mnuMacroList.count - 1) To 1 Step -1
        Unload mnuMacroList(count)
    Next count
End If
        
End Sub

Public Sub LoadTriggers(filepath As String)

Dim tempStr As String

On Error GoTo ErrHandler

lstTrigger.Clear
lstTriggerCommand.Clear

TriggerFile = filepath

Open filepath For Input As #1

Do While Not EOF(1)
    Line Input #1, tempStr
    lstTrigger.AddItem tempStr
    Line Input #1, tempStr
    lstTriggerCommand.AddItem tempStr
Loop

Close #1

DisplayInfo "Trigger File Loaded: " & filepath

Exit Sub

ErrHandler:
    DisplayInfo "Trigger File Not Loaded: " & filepath
    Close #1

End Sub

Public Sub SaveMacros(filepath As String)

Dim count As Integer

On Error GoTo ErrHandler

MacroFile = filepath

Open filepath For Output As #1

If lstMacro.ListCount > 0 Then
    For count = 0 To (lstMacro.ListCount - 1)
        Print #1, lstMacroName.List(count)
        Print #1, lstMacro.List(count)
        Print #1, lstMacroCommand.List(count)
    Next
End If

Close #1

DisplayInfo "Saved Macro File: " & filepath

Exit Sub
ErrHandler:
    If DebugMode Then LogError "Error in SaveMacros:" & vbCrLf & Err.Description & vbCrLf & filepath
    DisplayInfo "Unable To Save Macro File"
    Close #1

End Sub

Public Sub SaveColours(filepath As String)

Dim count As Integer

On Error GoTo ErrHandler

ColourFile = filepath

Open filepath For Output As #1

If lstColours.ListCount > 0 Then
    For count = 0 To (lstColours.ListCount - 1)
        Print #1, lstColourTrig.List(count)
        Print #1, lstColours.List(count)
    Next
End If

Close #1

DisplayInfo "Saved Colour File: " & filepath

Exit Sub
ErrHandler:
    If DebugMode Then LogError "Error in SaveColours:" & vbCrLf & Err.Description & vbCrLf & filepath
    DisplayInfo "Unable To Save Colour File"
    Close #1

End Sub
Public Sub SaveTriggers(filepath As String)

Dim count As Integer

On Error GoTo ErrHandler

TriggerFile = filepath

Open filepath For Output As #1

If lstTrigger.ListCount > 0 Then
    For count = 0 To (lstTrigger.ListCount - 1)
        Print #1, lstTrigger.List(count)
        Print #1, lstTriggerCommand.List(count)
    Next
End If

Close #1

DisplayInfo "Saved Trigger File: " & filepath

Exit Sub
ErrHandler:
    If DebugMode Then LogError "Error in SaveTriggers:" & vbCrLf & Err.Description & vbCrLf & filepath
    DisplayInfo "Unable To Save Trigger File"
    Close #1

End Sub

Public Function GetFilePath(filetype As String)

Select Case filetype
    Case "macro"
        GetFilePath = MacroFile
    Case "trigger"
        GetFilePath = TriggerFile
    Case "colour"
        GetFilePath = ColourFile
End Select

End Function

Public Sub SetFilePath(filetype As String, filepath As String)

Select Case filetype
    Case "macro"
        MacroFile = filepath
    Case "trigger"
        TriggerFile = filepath
    Case "colour"
        ColourFile = filepath
End Select

End Sub

Public Sub SetPrefs()

If frmOptions.chkFlashWindow.value <> 0 Then
    FlashWindowStatus = True
Else
    FlashWindowStatus = False
End If

If frmOptions.chkAutoLogin.value <> 0 Then
    AutoLogin = True
Else
    AutoLogin = False
End If

If frmOptions.chkLocalEcho.value <> 0 Then
    LocalEcho = True
Else
    LocalEcho = False
End If

txtDisplay.BackColor = frmOptions.boxBackgrnd.FillColor
txtDisplay.SelColor = frmOptions.boxForegrnd.FillColor
QuitCommand = frmOptions.txtQuitCommand.Text

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

ExitApp

End Sub

Public Sub SaveSettings(worldName As String)

Dim count As Integer

On Error Resume Next

If worldName = "--blank--" Then Exit Sub

SaveSetting App.Title, "Settings", "AppMsgColour", AppMsgColour
SaveSetting App.Title, "Settings", "MacroNameWidth", frmMacro.grdMacro.ColWidth(0)
SaveSetting App.Title, "Settings", "MacroTriggerWidth", frmMacro.grdMacro.ColWidth(1)
SaveSetting App.Title, "Settings", "MacroResponseWidth", frmMacro.grdMacro.ColWidth(2)
SaveSetting App.Title, "Settings", "TriggerTriggerWidth", frmTrigger.grdTrigger.ColWidth(0)
SaveSetting App.Title, "Settings", "TriggerResponseWidth", frmTrigger.grdTrigger.ColWidth(1)
SaveSetting App.Title, "Settings", "ShowSplash", frmOptions.chkSplash.value
If Me.WindowState <> vbMinimized Then
    SaveSetting App.Title, "Settings", "WindowState", frmMain.WindowState
    frmMain.WindowState = vbNormal
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting App.Title, "Settings", "MainTop", Me.Top
    SaveSetting App.Title, "Settings", "MainWidth", Me.Width
    SaveSetting App.Title, "Settings", "MainHeight", Me.Height
End If

SaveSetting App.Title, worldName, "AutoLogin", AutoLogin
SaveSetting App.Title, worldName, "LocalEcho", LocalEcho
SaveSetting App.Title, worldName, "Host", RemHost
SaveSetting App.Title, worldName, "Port", RemPort
SaveSetting App.Title, worldName, "Username", RemUser
SaveSetting App.Title, worldName, "Password", MydsEncrypt.Encrypt(RemPass)
SaveSetting App.Title, worldName, "QuitCommand", QuitCommand
SaveSetting App.Title, worldName, "UserPrompt", frmOptions.txtUsernamePrompt.Text
SaveSetting App.Title, worldName, "PassPrompt", frmOptions.txtPasswordPrompt.Text
SaveSetting App.Title, worldName, "EndOfNote", frmOptions.txtEndOfNote.Text
SaveSetting App.Title, worldName, "NewNoteCmd", frmOptions.txtNewNoteCmd.Text
SaveSetting App.Title, worldName, "NewMailCmd", frmOptions.txtNewMailCommand.Text
If frmOptions.dlgFont.FontName <> "" Then
    SaveSetting App.Title, worldName, "FontName", frmOptions.dlgFont.FontName
End If
SaveSetting App.Title, worldName, "FontSize", txtDisplay.SelFontSize
SaveSetting App.Title, worldName, "ForeColour", frmOptions.boxForegrnd.FillColor
SaveSetting App.Title, worldName, "BackColour", txtDisplay.BackColor
SaveSetting App.Title, worldName, "MacroFile", MacroFile
SaveSetting App.Title, worldName, "TriggerFile", TriggerFile
SaveSetting App.Title, worldName, "ColourFile", ColourFile
SaveSetting App.Title, worldName, "ProcessTriggers", ProcessTriggers
SaveSetting App.Title, worldName, "ProcessColours", ProcessColours
SaveSetting App.Title, worldName, "NoteWriterMode", frmNote.optMailMode.value
SaveSetting App.Title, worldName, "NoteWriterWidth", frmNote.txtWrap.Text
SaveSetting App.Title, worldName, "FlashWindow", FlashWindowStatus
SaveSetting App.Title, worldName, "KeepAliveInterval", frmOptions.txtKeepAliveTime.Text
SaveSetting App.Title, worldName, "KeepAliveString", frmOptions.txtKeepAliveStr.Text
SaveSetting App.Title, worldName, "KeepAliveEnable", frmOptions.chkKeepAlive.value
SaveSetting App.Title, worldName, "DoubleCrLf", frmOptions.chkDoubleCrLf.value

SaveMacros MacroFile
SaveTriggers TriggerFile
SaveColours ColourFile

For count = 0 To 9
    SaveSetting App.Title, worldName, "Button" + Format(count) + "Command", MacroButtonCommand(count)
Next count

End Sub

Private Sub Form_Resize()

If (frmMain.Width < 7275 Or frmMain.Height < 4845) And frmMain.WindowState = vbNormal Then
    If frmMain.Width < 7275 Then
        frmMain.Width = 7275
    End If
    If frmMain.Height < 4845 Then
        frmMain.Height = 4845
    End If
    Exit Sub
End If

If frmMain.WindowState = vbNormal Or frmMain.WindowState = vbMaximized Then
    txtDisplay.Width = frmMain.Width - 2070
    txtEntry.Width = frmMain.Width - 2070
    txtDisplay.Height = frmMain.Height - 2430
    txtEntry.Top = txtDisplay.Height + txtDisplay.Top + 60
    lstMacroName.Height = (txtEntry.Top + txtEntry.Height) - lstMacroName.Top
End If

End Sub
Public Sub SocketConnect(RemoHost As String, RemoPort As String, RemoUser As String, RemoPass As String)

RemHost = RemoHost
RemPort = RemoPort
RemUser = RemoUser
RemPass = RemoPass

sckMain.RemoteHost = RemHost
sckMain.RemotePort = RemPort
sckMain.Connect
DisplayInfo "Connecting to: " + RemHost + " " + RemPort
frmMain.Caption = "Sanx's MUD Client - " & RemHost & " " & RemPort

End Sub

Private Sub lstMacroName_DblClick()

SendText lstMacro.List(lstMacroName.ListIndex)
txtEntry.SetFocus

End Sub

Private Sub lstMacroName_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 And lstMacroName.ListIndex <> -1 Then
    frmMacro.PopulateGrid
    frmMacro.grdMacro.Row = lstMacroName.ListIndex + 1
    PopupMenu mnuMacroPopUp
End If

End Sub

Private Sub mnuClearBuffer_Click()

txtDisplay.TextRTF = ""

End Sub

Private Sub mnuColours_Click()

frmColours.PopulateGrid
frmColours.Show vbModal, Me

End Sub

Private Sub mnuEditClipGame_Click()

Dim tempStr As String

tempStr = Clipboard.GetText
SendText tempStr

End Sub

Private Sub mnuEnableDebug_Click()

Dim TempFlag As Boolean

TempFlag = DebugMode

DebugMode = Not TempFlag
mnuEnableDebug.Checked = DebugMode

End Sub

Private Sub mnuEnableTriggers_Click()

TriggerProcessChange

End Sub

Private Sub TriggerProcessChange()

mnuEnableTriggers.Checked = Not (mnuEnableTriggers.Checked)
ProcessTriggers = mnuEnableTriggers.Checked
If ProcessTriggers Then
    barStatus.Panels(3).Text = "Triggers: ON"
Else
    barStatus.Panels(3).Text = "Triggers: OFF"
End If

End Sub

Private Sub mnuEnableColours_Click()

ColourProcessChange

End Sub

Private Sub ColourProcessChange()

mnuEnableColours.Checked = Not (mnuEnableColours.Checked)
ProcessColours = mnuEnableColours.Checked
If ProcessColours Then
    barStatus.Panels(4).Text = "Colours: ON"
Else
    barStatus.Panels(4).Text = "Colours: OFF"
End If

End Sub

Private Sub mnuMacro_Click()

frmMacro.PopulateGrid
frmMacro.Show vbModal

End Sub

Private Sub mnuMacroButton_Click()

Dim count As Integer

For count = 0 To 9
    If frmMacroButtons.chkButton(count).value = 1 Then
        frmMacroButtons.PopulateList count
        frmMacroButtons.lstMacros(count).ListIndex = MacroButtonCommand(count)
    End If
Next count

frmMacroButtons.Show

End Sub

Private Sub mnuMacroList_Click(index As Integer)

SendText lstMacro.List(index)

End Sub

Private Sub mnuNewMacro_Click()

frmMacro.PopulateGrid
frmMacro.Show
frmNewMacro.Show vbModal

End Sub


Private Sub mnuNoteWriter_Click()

frmNote.Show

End Sub

Private Sub mnuPopCopy_Click()

Clipboard.Clear
Clipboard.SetText txtDisplay.SelText

End Sub

Private Sub mnuPopCopyPaste_Click()

txtEntry.SelText = txtDisplay.SelText

End Sub

Private Sub mnuPopEditMacro_Click()

frmMacro.EditMacro (lstMacroName.ListIndex + 1)

End Sub

Private Sub mnuPopNewColourHilite_Click()

If txtDisplay.SelLength > 0 Then
    frmNewColour.txtTextHilite = txtDisplay.SelText
    frmNewColour.Show vbModal
End If

End Sub

Private Sub mnuPopNewTrigger_Click()

If txtDisplay.SelLength > 0 Then
    frmNewTrigger.txtTrigger = txtDisplay.SelText
    frmNewTrigger.Show vbModal
End If

End Sub

Private Sub mnuPopQuote_Click()

txtEntry.SelText = ReplaceStr(txtDisplay.SelText, vbCrLf, vbCrLf & "> ")

End Sub


Private Sub mnuPopRemoveMacro_Click()

frmMacro.RemoveMacro

End Sub

Private Sub mnuPopSelectAll_Click()

txtDisplay.SelStart = 0
txtDisplay.SelLength = Len(txtDisplay.Text)

End Sub

Private Sub mnuPopUpCopy_Click()

Clipboard.Clear
Clipboard.SetText txtDisplay.SelText

End Sub

Private Sub mnuPopUpCopyPaste_Click()

txtEntry.SelText = txtDisplay.SelText

End Sub

Private Sub mnuPrintBuffer_Click()

dlgFile.CancelError = True
On Error GoTo ErrHandler
dlgFile.ShowPrinter
txtDisplay.SelPrint Printer.hDC

ErrHandler:
Exit Sub

End Sub

Private Sub mnuRecordMacro_Click()

frmRecordMacros.txtCmdSeq.Text = ""
frmRecordMacros.txtMacroCommand.Text = "/"
frmRecordMacros.txtMacroName.Text = ""
frmRecordMacros.Visible = True
SetPos frmRecordMacros
frmRecordMacros.txtMacroName.SetFocus

End Sub

Private Sub mnuSaveHistory_Click()

SaveHistory

End Sub

Private Sub SaveHistory()

Dim saveFile As String

CheckDirExists

On Error GoTo ErrHandler

dlgFile.DefaultExt = "txt"
dlgFile.InitDir = RunPath & "\histories"
dlgFile.Filter = "*.txt|Text files / MUD histories"
dlgFile.ShowSave

saveFile = dlgFile.FileName

If saveFile <> "" Then
    Open saveFile For Output As #1
    Print #1, txtDisplay.Text
    Close #1
    DisplayInfo "Saved History File: " & saveFile
Else
    DisplayInfo "Error Saving History File"
End If

dlgFile.FileName = ""
ChDir RunPath

Exit Sub

ErrHandler:

If DebugMode Then LogError "Error in SaveHistory:" & vbCrLf & Err.Description

End Sub

Private Sub CheckDirExists()

Dim response

response = Dir(RunPath + "\histories", vbDirectory)

If response = "" Then
    MkDir RunPath + "\histories"
End If

End Sub

Private Sub mnuSaveSettings_Click()

SaveSettings LastWorld

End Sub

Private Sub mnuShowSndControls_Click()

frmSndControl.Show

End Sub

Private Sub mnuStopSound_Click()

frmSndControl.sndMain.Stop

End Sub

Private Sub mnuTrigger_Click()

frmTrigger.PopulateGrid
If ProcessTriggers Then
    frmTrigger.Caption = "Triggers are currently enabled"
Else
    frmTrigger.Caption = "Triggers are currently disabled"
End If
frmTrigger.Show vbModal

End Sub

Private Sub mnuViewHistory_Click()

ViewHistory

End Sub

Private Sub ViewHistory()

Dim historyName As String, tempStr As String

On Error GoTo ErrHandler

dlgFile.DefaultExt = "txt"
dlgFile.FileName = "*.txt"
dlgFile.InitDir = RunPath & "\histories"
dlgFile.Filter = "*.txt|Text files / MUD histories"
dlgFile.ShowOpen

historyName = dlgFile.FileName

If historyName <> "" Then
    Open historyName For Input As #1
    Do
        Line Input #1, tempStr
        frmHistory.txtDisplay.Text = frmHistory.txtDisplay.Text & tempStr & vbCrLf
    Loop Until EOF(1)
    Close #1
    frmHistory.barStatus.SimpleText = historyName
    frmHistory.Show vbModal
Else
    DisplayInfo "Unable To Load History File"
End If

ChDir RunPath

Exit Sub

ErrHandler:
LogError "Error Loading History file: " + historyName + vbCrLf + Err.Description
ChDir RunPath

End Sub

Private Sub tmrKeepAlive_Timer()

Static siTick As Integer

    If siTick > intTick Then
        If sckMain.State = sckConnected Then
            SendText frmOptions.txtKeepAliveStr.Text
        End If
        siTick = 0
    Else
        siTick = siTick + 1
    End If

End Sub

Private Sub toolMain_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
    Case "connect"
        ConnectMud
    Case "disconnect"
        SocketClose
    Case "quit"
        SendText QuitCommand
    Case "copy"
        Clipboard.Clear
        Clipboard.SetText txtDisplay.SelText
    Case "copypaste"
        txtEntry.SelText = txtDisplay.SelText
    Case "copyclip"
        Dim tempStr As String
        tempStr = Clipboard.GetText
        SendText tempStr
    Case "note"
        frmNote.Show
    Case "macros"
        frmMacro.PopulateGrid
        frmMacro.Show vbModal
    Case "history"
        SaveHistory
    Case "preferences"
        frmOptions.Show vbModal, Me
    Case "colours"
        frmColours.PopulateGrid
        frmColours.Show vbModal, Me
End Select

End Sub


Private Sub txtEntry_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 38
        If CacheCount > 0 Then
            CacheCount = CacheCount - 1
            txtEntry.Text = lstEntry.List(CacheCount)
            txtEntry.SelStart = 0
            txtEntry.SelLength = Len(txtEntry.Text) + 1
        End If
    Case 40
        If CacheCount < (lstEntry.ListCount - 1) Then
            CacheCount = CacheCount + 1
            txtEntry.Text = lstEntry.List(CacheCount)
            txtEntry.SelStart = 0
            txtEntry.SelLength = Len(txtEntry.Text) + 1
        End If
    Case 96
        SendText "d"
    Case 97
        SendText "sw"
    Case 98
        SendText "s"
    Case 99
        SendText "se"
    Case 100
        SendText "w"
    Case 101
        SendText "u"
    Case 102
        SendText "e"
    Case 103
        SendText "nw"
    Case 104
        SendText "n"
    Case 105
        SendText "ne"
    Case 112 To 121
        If sckMain.State = sckConnected Then
            SendText lstMacro.List(MacroButtonCommand(KeyCode - 112))
        End If
End Select

End Sub

Private Sub mnuAboutShow_Click()

frmAbout.Show vbModal, Me

End Sub

Private Sub mnuConnectMud_Click()

ConnectMud

End Sub

Private Sub ConnectMud()

SocketClose
DoneAutoLogin = False
frmConnect.GetWorlds
frmConnect.Show vbModal, Me

End Sub

Public Sub DisplayInfo(INFO As String)

DisplayText "== " + INFO + " ==" + vbCrLf, AppMsgColour, frmOptions.boxForegrnd.FillColor

End Sub

Private Sub mnuDisconnect_Click()

SocketClose

End Sub

Private Sub mnuExit_Click()

ExitApp

End Sub

Public Sub ExitApp()

SaveSettings LastWorld

End

End Sub

Private Sub mnuPreferences_Click()

frmOptions.Show vbModal, Me

End Sub

Private Sub mnuSendQuit_Click()

SendText QuitCommand

End Sub

Private Sub sckMain_Close()

frmMain.Caption = "Sanx's MUD Client"
DisplayInfo "Socket Closed"

End Sub

Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)

Dim recData As String

sckMain.GetData recData
CheckText recData

End Sub
Private Sub CheckText(recData As String)

On Error Resume Next

Dim userPrompt As String, passPrompt As String, shownText As String, SearchStr As String
Dim localEchoFlag As Boolean
Dim where As Integer, count As Integer, tempCount As Integer, checksize As Integer
Dim whereStr As Long

recData = Replace(recData, "ÿ", "")
recData = Replace(recData, "", "")
recData = Replace(recData, "û", "")
recData = Replace(recData, "ü", "")
recData = Replace(recData, "þ", "")
recData = Replace(recData, "ý", "")
recData = Replace(recData, "> ", "")
If frmOptions.chkDoubleCrLf.value <> 0 Then recData = Replace(recData, vbLf + vbCr, vbNewLine)

checksize = Len(recData)

DisplayText recData, frmOptions.boxForegrnd.FillColor, frmOptions.boxForegrnd.FillColor
localEchoFlag = LocalEcho

If AutoLogin = True And DoneAutoLogin = False Then
    shownText = txtDisplay.Text
    userPrompt = frmOptions.txtUsernamePrompt.Text
    passPrompt = frmOptions.txtPasswordPrompt.Text
    If InStr(SearchLen(recData), shownText, userPrompt) And Len(RemUser) > 0 Then
        SendText RemUser
    End If
    If InStr(SearchLen(recData), shownText, passPrompt) And Len(RemPass) > 0 Then
        LocalEcho = False
        SendText RemPass
        LocalEcho = localEchoFlag
        DoneAutoLogin = True
    End If
End If

SearchStr = Right$(txtDisplay.Text, checksize)

'Trigger Highlighting Routine
If ProcessTriggers Then
    For count = 0 To lstTrigger.ListCount
        If InStr(1, SearchStr, ReplaceStr(lstTrigger.List(count), "@@", vbCrLf)) > 0 And Len(lstTrigger.List(count)) > 0 Then
            txtTemp.Text = ReplaceStr(lstTriggerCommand.List(count), ";", vbCrLf)
            SendText txtTemp.Text
        End If
    Next count
End If

'Colour Highlighting Routine
If ProcessColours Then
    For count = 0 To lstColours.ListCount
        whereStr = InStr(1, SearchStr, lstColourTrig.List(count))
        If whereStr > 0 And Len(lstColourTrig.List(count)) > 0 Then
            While whereStr > 0
                If whereStr > 0 And Len(lstColourTrig.List(count)) > 0 Then
                    txtDisplay.SelStart = (whereStr + (Len(txtDisplay.Text) - checksize)) - 1
                    txtDisplay.SelLength = Len(lstColourTrig.List(count))
                    txtDisplay.SelColor = lstColours.List(count)
                End If
                whereStr = InStr(whereStr + 1, SearchStr, lstColourTrig.List(count))
            Wend
            tempCount = 0
            whereStr = 0
            txtDisplay.SelStart = Len(txtDisplay.Text)
        End If
    Next count
End If

If FlashWindowStatus = True And frmMain.WindowState = vbMinimized Then FlshWindow

End Sub

Private Function SearchLen(recData As String)

SearchLen = Len(txtDisplay.Text) - Len(recData)

End Function

Private Sub DisplayText(recData As String, pColor As Long, tColor As Long)

Dim displayLen As Long

displayLen = Len(txtDisplay.Text)

If displayLen + Len(recData) > Scrollback Then
    txtDisplay.SelStart = 0
    txtDisplay.SelLength = (Scrollback / 8)
    txtDisplay.SelRTF = " "
    LogError "Display buffer at" + Str(displayLen) + " characters. Trimmed to" + Str(Len(txtDisplay.Text))
End If

txtDisplay.SelStart = Len(txtDisplay.Text)
txtDisplay.SelColor = pColor
OutputText txtDisplay, recData
txtDisplay.SelColor = tColor
txtDisplay.SelStart = Len(txtDisplay.Text)

End Sub

Private Sub OutputText(Rich As RichTextBox, ansicode As String)

Dim isEscape As Boolean
Dim sBuffer As String, i As Integer
Dim curChr As String
Dim j As Integer, modes() As String
Dim INFO() As String, RawDATA As String
Dim Esc As String

Esc = Chr(27)
isEscape = False
For i = 1 To Len(ansicode)
    curChr = Mid(ansicode, i, 1)
    If curChr = "[" And isEscape Then GoTo nex
    If curChr = Esc Then
        Rich.SelText = sBuffer
        Rich.SelStart = Len(Rich.Text)
        Rich.SelFontName = frmOptions.dlgFont.FontName
        Rich.SelFontSize = frmOptions.dlgFont.FontSize
        sBuffer = ""
        isEscape = True
    ElseIf curChr = "m" And isEscape Then
        If sBuffer = "" Then isEscape = False: GoTo nex
        modes = Split(sBuffer, ";")
        For j = LBound(modes) To UBound(modes)
            If IsNumeric(modes(j)) = False Then GoTo nex
            Select Case modes(j)
                Case 0:     If Rich.SelBold Then Rich.SelBold = False
                Case 4:     If Rich.SelUnderline = False Then Rich.SelUnderline = True
                Case 30:    Rich.SelColor = vbBlack
                Case 31:    Rich.SelColor = vbRed
                Case 32:    Rich.SelColor = vbGreen
                Case 33:    Rich.SelColor = vbYellow
                Case 34:    Rich.SelColor = vbBlue
                Case 35:    Rich.SelColor = vbMagenta
                Case 36:    Rich.SelColor = vbCyan
                Case 37:    Rich.SelColor = vbWhite
            End Select
        Next j
        sBuffer = ""
        isEscape = False
    Else
        sBuffer = sBuffer & curChr
        If i = Len(ansicode) Then
            Rich.SelFontName = frmOptions.dlgFont.FontName
            Rich.SelFontSize = frmOptions.dlgFont.FontSize
            Rich.SelText = sBuffer
            Rich.SelFontName = frmOptions.dlgFont.FontName
            Rich.SelFontSize = frmOptions.dlgFont.FontSize
        End If
    End If
nex:
Next i

End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

DisplayInfo "Socket Error: " + Description
SocketClose

End Sub

Private Sub SocketClose()

If sckMain.State <> sckClosed Then
    sckMain.Close
End If

frmMain.Caption = "Sanx's MUD Client"
DisplayInfo "Closing Socket"

End Sub

Private Sub txtDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If txtDisplay.SelLength > 0 Then
    mnuPopUpCopy.Enabled = True
    mnuPopUpCopyPaste.Enabled = True
    mnuPopNewTrigger.Enabled = True
    mnuPopNewColourHilite.Enabled = True
Else
    mnuPopUpCopy.Enabled = False
    mnuPopUpCopyPaste.Enabled = False
    mnuPopNewTrigger.Enabled = False
    mnuPopNewColourHilite.Enabled = False
End If

If Button = 2 Then
    PopupMenu mnuPop
End If

End Sub


Private Sub txtEntry_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 96 To 105
        txtEntry.Text = ""
End Select

End Sub
