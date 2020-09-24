VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MEmail 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   ClientHeight    =   7920
   ClientLeft      =   15
   ClientTop       =   -105
   ClientWidth     =   6795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "MEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MEmail.frx":08CA
   ScaleHeight     =   7920
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   1170
      Left            =   8610
      ScaleHeight     =   1110
      ScaleWidth      =   675
      TabIndex        =   32
      Top             =   2040
      Width           =   735
   End
   Begin Project1.CandyButton CandyButton7 
      Height          =   1005
      Left            =   6120
      TabIndex        =   31
      ToolTipText     =   "Outlook Contacts Display"
      Top             =   1485
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1773
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Contacts"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16711680
      Picture         =   "MEmail.frx":1DB3
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.ctxSysTray ctxSysTray1 
      Left            =   6435
      Top             =   4650
      _ExtentX        =   450
      _ExtentY        =   450
      TrayIcon        =   "MEmail.frx":2205
   End
   Begin VB.PictureBox picTest 
      Height          =   285
      Left            =   8835
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   24
      Top             =   1695
      Visible         =   0   'False
      Width           =   300
   End
   Begin Project1.ocxFormShape ocxFormShape1 
      Left            =   5250
      Top             =   7755
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin Project1.CandyButton CandyButton2 
      Height          =   795
      Left            =   105
      TabIndex        =   16
      Top             =   420
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1402
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Scan/Cap to PDF"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "MEmail.frx":2ADF
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   8295
      Pattern         =   "*.bmp"
      TabIndex        =   15
      Top             =   30
      Width           =   1185
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   8340
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   60
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   270
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txt_email_to 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1470
      TabIndex        =   3
      Top             =   2160
      Width           =   4605
   End
   Begin VB.TextBox txt_subject 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1470
      TabIndex        =   4
      Top             =   2520
      Width           =   4605
   End
   Begin VB.TextBox txt_attach 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1470
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   4605
   End
   Begin VB.TextBox txt_email_from 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1470
      TabIndex        =   2
      Top             =   1800
      Width           =   4605
   End
   Begin VB.TextBox txt_smtp_server 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1470
      TabIndex        =   1
      Top             =   1470
      Width           =   4605
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      Height          =   735
      Left            =   390
      TabIndex        =   7
      Top             =   6960
      Width           =   5895
      Begin VB.TextBox txt_status 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   " Cover Letter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3570
      Left            =   390
      TabIndex        =   0
      Top             =   3405
      Width           =   5895
      Begin VB.TextBox txt_message_text 
         Appearance      =   0  'Flat
         Height          =   3255
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   5655
      End
   End
   Begin Project1.CandyButton CandyButton3 
      Height          =   795
      Left            =   1785
      TabIndex        =   17
      Top             =   420
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1402
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Send Fax"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "MEmail.frx":33B9
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.CandyButton CandyButton1 
      Height          =   795
      Left            =   3465
      TabIndex        =   18
      Top             =   420
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1402
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Servers"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "MEmail.frx":3C93
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.CandyButton CandyButton4 
      Height          =   795
      Left            =   5130
      TabIndex        =   19
      Top             =   420
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   1402
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Saved Faxes"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "MEmail.frx":456D
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.CandyButton CandyButton5 
      Height          =   360
      Left            =   5640
      TabIndex        =   20
      Top             =   30
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "-"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16777215
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   65280
      ColorButtonUp   =   255
      ColorButtonDown =   255
      BorderBrightness=   0
      ColorBright     =   255
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin Project1.CandyButton CandyButton6 
      Height          =   360
      Left            =   6105
      TabIndex        =   21
      Top             =   30
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "X"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16777215
      Style           =   4
      Checked         =   0   'False
      ColorButtonHover=   8454016
      ColorButtonUp   =   255
      ColorButtonDown =   255
      BorderBrightness=   0
      ColorBright     =   255
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   6195
      Top             =   6300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Fax Files|*.pdf;*.txt"
      Flags           =   4
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H8000000B&
      Height          =   195
      Left            =   6180
      Top             =   2940
      Width           =   435
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "žŸ "
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6195
      TabIndex        =   25
      ToolTipText     =   "Open PDF File to Fax"
      Top             =   2895
      Width           =   450
   End
   Begin VB.Image Image1a 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   6165
      Picture         =   "MEmail.frx":4E47
      ToolTipText     =   "Not OnTop"
      Top             =   2505
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "MEmail.frx":5711
      Top             =   -30
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "PDF File Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   390
      TabIndex        =   11
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Email to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   390
      TabIndex        =   10
      Top             =   2160
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Email From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pdf-Email-Fax"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   330
      Index           =   1
      Left            =   585
      TabIndex        =   23
      Top             =   45
      Width           =   1785
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pdf-Email-Fax"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   585
      TabIndex        =   22
      Top             =   60
      Width           =   1785
   End
   Begin VB.Image Image1a 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   6210
      Picture         =   "MEmail.frx":5FDB
      ToolTipText     =   "OnTop"
      Top             =   2505
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   30
      Top             =   1455
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Email From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   29
      Top             =   1815
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Email to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   390
      TabIndex        =   28
      Top             =   2175
      Width           =   1005
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   390
      TabIndex        =   27
      Top             =   2535
      Width           =   1005
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "PDF File Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   26
      Top             =   2895
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuViewFax 
         Caption         =   "View Fax or Log"
      End
      Begin VB.Menu mnuOpenSaved 
         Caption         =   "Open Saved Faxes Directory"
      End
   End
   Begin VB.Menu mnuSccaan 
      Caption         =   "Sccc"
      Visible         =   0   'False
      Begin VB.Menu mnuSelSource 
         Caption         =   "Select Source"
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Scan Documents to PDF"
      End
      Begin VB.Menu jgtjgjy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCaptureUrl 
         Caption         =   "Capture Screen"
      End
   End
   Begin VB.Menu mnuSendFaxx 
      Caption         =   "SendFaxx"
      Visible         =   0   'False
      Begin VB.Menu mnu_send 
         Caption         =   "Send Present Fax"
      End
      Begin VB.Menu mnuSendSaved 
         Caption         =   "Send Saved Fax"
      End
      Begin VB.Menu mnuSendPDF 
         Caption         =   "Send a PDF File"
      End
   End
   Begin VB.Menu mnuServers 
      Caption         =   "Servers"
      Visible         =   0   'False
      Begin VB.Menu mnuBright 
         Caption         =   "Brighthouse"
      End
      Begin VB.Menu mnuEarth 
         Caption         =   "Earthlink"
      End
      Begin VB.Menu mnuCom 
         Caption         =   "Comcast"
      End
      Begin VB.Menu mnuNet 
         Caption         =   "Netzero"
      End
      Begin VB.Menu mnuMSN 
         Caption         =   "MSN"
      End
      Begin VB.Menu mnuJuno 
         Caption         =   "Juno"
      End
      Begin VB.Menu mnuProdigy 
         Caption         =   "Prodigy"
      End
      Begin VB.Menu mnuOther 
         Caption         =   "Other"
      End
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore "
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu hrhtht 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "MEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private shlShell As shell32.Shell
Private shlFolder As shell32.Folder
Const err_SMTP = "No SMTP server"
Const err_FROM = "No Email from"
Const err_TO = "No Email to"
Const err_SUBJECT = "No subject"

Dim response As String, i As Long
Dim Changed As Boolean
Dim ClippedText As String
Dim intSave As Integer
Sub wait_for(winsock_answare As String)
    Do While Left(response, 3) <> winsock_answare
        DoEvents
    Loop
    response = ""
End Sub

Function find_date() As String
    Dim temp As String
    Dim fd_day As String
    Dim fd_month As String
    Dim fd_time As String
    
    fd_day = Format(Date, "Dddd")
    Select Case fd_day
        Case "éåí øàùåï": fd_day = "Sun, "
        Case "éåí ùðé": fd_day = "Mon, "
        Case "éåí ùìéùé": fd_day = "Tue, "
        Case "éåí øáéòé": fd_day = "Wed, "
        Case "éåí çîéùé": fd_day = "Thu, "
        Case "éåí ùéùé": fd_day = "Fri, "
        Case "éåí ùáú": fd_day = "Sat, "
    End Select
    fd_month = Month(Date)
    Select Case fd_month
        Case 1: fd_month = "Jan "
        Case 2: fd_month = "Feb "
        Case 3: fd_month = "Mar "
        Case 4: fd_month = "Apr "
        Case 5: fd_month = "May "
        Case 6: fd_month = "Jun "
        Case 7: fd_month = "Jul "
        Case 8: fd_month = "Aug "
        Case 9: fd_month = "Sep "
        Case 10: fd_month = "Oct "
        Case 11: fd_month = "Nov "
        Case 12: fd_month = "Dec "
    End Select
    fd_time = Format(Time) & " +0200"
    temp = fd_day & Day(Format(Date)) & " " & fd_month & Year(Format(Date, "dd/mm/yyyy")) & " " & fd_time
    find_date = temp
End Function

Function attach_file(attach_str As String) As String
    Dim s As Integer
    Dim temp As String
    
    s = InStr(1, attach_str, "\")
    temp = attach_str
    Do While s > 0
        temp = Mid(temp, s + 1, Len(temp))
        s = InStr(1, temp, "\")
    Loop
    attach_file = temp
End Function

Function encode_the_file(attach_str As String) As String
    Dim Blocksize As Long
    Dim Buffer As String
    Dim s As String
    Dim i As Long
    Dim temp As String
    
    Open App.Path & "\test.pdf" For Binary Access Read As #1
        Blocksize = 3
        Do While Not EOF(1)
            Buffer = Space(Blocksize)
            Get 1, , Buffer
            s = s & base64_encode_string(Buffer)
            DoEvents
        Loop
    Close #1
    For i = 1 To Len(s) Step 76
        temp = temp & Mid(s, i, 76) & vbCrLf
    Next i
    temp = Mid(temp, 1, Len(temp) - 2)
    encode_the_file = temp
End Function

Sub send_email(email_to As String, email_from As String, subject As String, message_text As String, attach As String)
    Const boundary = "Hapoel_Tel_Aviv"
    
    Dim se_body As String
    Dim se_date As String
    Dim se_from As String
    Dim se_to As String
    Dim se_mime As String
    Dim se_content_type As String
    Dim se_content_type_message As String
    Dim se_content_type_attach As String
    Dim x_mailer As String
    Dim x_oem As String
     
    se_date = "Date: " & find_date
    se_from = "From: " & email_from
    se_to = "To: " & email_to
    subject = "Subject: " & subject
    
    se_mime = "MIME-Version: 1.0"
    se_content_type = "Content-Type: multipart/mixed;" & vbCrLf _
        & vbTab & "boundary = " & """" & boundary & """"
    
    x_oem = "X-OEM: zubin"
    x_mailer = "X-Mailer: " & """" & "Faxer" & """"
    
    se_content_type_message = "This is a multi-part message in MIME format." & vbCrLf _
        & "--" & boundary & vbCrLf _
        & "Content-Type: text/plain;" & vbCrLf _
        & vbTab & "charset=" & """" & "iso-8859-1" & """" & vbCrLf _
        & "Content-Transfer-Encoding: 7bit"
        
    If Len(txt_attach.Text) > 0 Then
        se_content_type_attach = "--" & boundary & vbCrLf _
            & "Content-Type: application/octet-stream;" & vbCrLf _
            & vbTab & "name=" & attach_file(txt_attach.Text) & vbCrLf _
            & "Content-Transfer-Encoding: base64" & vbCrLf _
            & "Content-Disposition: attachment;" & vbCrLf _
            & vbTab & "filename=" & attach_file(txt_attach.Text) & vbCrLf _
            & vbCrLf _
            & encode_the_file(txt_attach.Text)
    End If
    
    se_body = se_from & vbCrLf _
        & se_to & vbCrLf _
        & subject & vbCrLf _
        & se_date & vbCrLf _
        & se_mime & vbCrLf _
        & x_oem & vbCrLf _
        & x_mailer & vbCrLf _
        & se_content_type & vbCrLf _
        & vbCrLf _
        & se_content_type_message & vbCrLf _
        & vbCrLf _
        & message_text & vbCrLf _
        & vbCrLf _
        & se_content_type_attach & vbCrLf _
        & "." & vbCrLf
    
    txt_status.Text = "Sending Fax..." & vbCrLf & txt_status.Text & vbCrLf
    Winsock1.SendData "HELO " & Left(email_from, InStr(1, email_from, "@") - 1) & vbCrLf
    wait_for "250"
    Winsock1.SendData "MAIL FROM: " & email_from & vbCrLf
    wait_for "250"
    Winsock1.SendData "RCPT TO: " & email_to & vbCrLf
    wait_for "250"
    Winsock1.SendData "DATA" & vbCrLf
    wait_for "354"
    Winsock1.SendData se_body
    wait_for "250"
    Winsock1.SendData "QUIT" & vbCrLf
    wait_for "221"
    txt_status.Text = "Message sent." & vbCrLf & txt_status.Text & vbCrLf
    Winsock1.Close
    DoEvents
End Sub

Sub connect_to_smtp_server(smtp_server As String)
    Winsock1.LocalPort = 0
    Winsock1.RemoteHost = txt_smtp_server
    Winsock1.RemotePort = 25
    Winsock1.Connect
End Sub

Sub init_me()
    txt_status.Text = "Ready." & vbCrLf
    response = ""
End Sub

Function form_errors() As Boolean
    Dim temp As Boolean
    
    temp = False
    If Len(txt_subject.Text) = 0 Then
        txt_status.Text = "Error: " & err_SUBJECT & "." & vbCrLf & txt_status.Text & vbCrLf
        temp = True
    End If
    If Len(txt_email_to.Text) = 0 Then
        txt_status.Text = "Error: " & err_TO & "." & vbCrLf & txt_status.Text & vbCrLf
        temp = True
    End If
    If Len(txt_email_from.Text) = 0 Then
        txt_status.Text = "Error: " & err_FROM & "." & vbCrLf & txt_status.Text & vbCrLf
        temp = True
    End If
    If Len(txt_smtp_server.Text) = 0 Then
        txt_status.Text = "Error: " & err_SMTP & "." & vbCrLf & txt_status.Text & vbCrLf
        temp = True
    End If
    form_errors = temp
End Function

Private Sub CandyButton1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu mnuServers
End Sub

Private Sub CandyButton2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu mnuSccaan
End Sub

Private Sub CandyButton3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu mnuSendFaxx
End Sub

Private Sub CandyButton4_Click()
PopupMenu mnuFile

End Sub

Private Sub CandyButton5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Me.WindowState = 1
Me.Visible = False
End Sub

Private Sub CandyButton6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub ctxSysTray1_DblClick(Button As Integer)
    Me.Visible = True
    SetTopMostWindow Me.hwnd, True
    SetTopMostWindow Me.hwnd, False
End Sub

Private Sub ctxSysTray1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuSystray
End If
'Me.Visible = True
End Sub
Private Function OutlookExists() As Integer
    On Error Resume Next
    
    Dim objOutlook As Object
    Dim intExists As Integer
    
    'try to create a new instance of MS Outlook
    Set objOutlook = CreateObject("Outlook.Application")
    
    'if the instance of MS Outlook does not exist then MS Outlook is not installed
    If objOutlook Is Nothing Then
        intExists = 0
        
    'else, MS Outlook is installed
    Else
        intExists = 1
    End If
    
    'distroy the object
    Set objOutlook = Nothing
    
    'return the status of MS Outlook being installed
    OutlookExists = intExists
End Function

Private Sub Form_Activate()
    SetTopMostWindow Me.hwnd, True
    SetTopMostWindow Me.hwnd, False

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim f As String
ctxSysTray1.AddIconToSystray "PDF-Email-Fax"
Changed = False
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
If OutlookExists = 1 Then CandyButton7.Visible = True
If Dir(App.Path & "\Sent Faxes") = "" Then MkDir App.Path & "\Sent Faxes"
If Dir(App.Path & "\Parameters") = "" Then
    Open App.Path & "\Parameters" For Output As #1
        Print #1, "Your SMTP Server"
        Print #1, "whoever@whoever.com"
    Close #1

End If
    Open App.Path & "\Parameters" For Input As #1
        Line Input #1, f
        txt_smtp_server = f
        Line Input #1, f
        txt_email_from = f
    Close #1
    init_me
    txt_email_to = "whoever@whoever.com"
    txt_attach = txt_email_to & Format(Now, "ddmmyyhhmmss") & ".pdf"
    txt_subject = "Fax: from " & txt_email_from & " " & Now
    txt_message_text = "                                          Cover Letter" & vbCrLf _
                       & "-------------------------------------------------------------------------------------------------------------------" & vbCrLf _
                       & txt_email_from & vbCrLf _
                       & txt_subject

End Sub

Function StartDoc(DocName As String) As Long
On Error Resume Next

Dim Scr_hDC As Long
Scr_hDC = GetDesktopWindow()
StartDoc = ShellExecute(Scr_hDC, "", DocName, _
"", "C:\", -1)
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If isVaildEmail(Trim(txt_email_from)) = False And txt_email_from <> "" Then
    MsgBox "Your Email From address is invalid! Please correct this."
    txt_email_from = ""
End If
If isVaildEmail(Trim(txt_email_to)) = False And txt_email_to <> "" Then
    MsgBox "Your Email To address is invalid! Please correct this."
    txt_email_to = ""
End If

If Changed = True Then
    Changed = False
    txt_attach = txt_email_to & Format(Now, "ddmmyyhhmmss") & ".pdf"
    Close #1
    Open App.Path & "\Parameters" For Output As #1
        Print #1, txt_smtp_server.Text
        Print #1, txt_email_from.Text
    Close #1
    txt_message_text = "                                          Cover Letter" & vbCrLf _
                       & "-------------------------------------------------------------------------------------------------------------------" & vbCrLf _
                       & txt_email_from & vbCrLf _
                       & txt_subject
End If
End Sub
Function isVaildEmail(EmailName As String) As Boolean
Dim ipart As Integer, lpart As Integer, Length As Integer
Dim isVaild As Boolean
Dim sEmail As String
    If Len(Trim(EmailName) <= 0) Then isVaildEmail = False
    sEmail = Trim(EmailName)
    ipart = InStr(sEmail, "@")
    lpart = InStr(ipart + 1, sEmail, ".")

        Length = Len(Trim(Mid(sEmail, lpart + 1, 3)))
        If ipart <= 0 Or lpart <= 0 Then
            isVaild = False
        ElseIf Length < 3 Then
            isVaild = False
        ElseIf ipart = 1 Then
            isVaild = False
        ElseIf lpart = Len(sEmail) Then
            isVaild = False
        Else
            isVaild = True
        End If
        isVaildEmail = isVaild
End Function
Private Sub Form_Unload(Cancel As Integer)
    Open App.Path & "\Parameters" For Output As #1
        Print #1, txt_smtp_server.Text
        Print #1, txt_email_from.Text
    Close #1
    Winsock1.Close
    Unload Me
    Set MEmail = Nothing
    CloseAll
End Sub

Sub CloseAll()
    On Error Resume Next
    Dim intFrmNum As Integer
    intFrmNum = Forms.count


    Do Until intFrmNum = 0
        Unload Forms(intFrmNum - 1)
        Set Forms(intFrmNum - 1) = Nothing
        intFrmNum = intFrmNum - 1
    Loop
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Image1a_Click(Index As Integer)
If Index = 1 Then
    Image1a(1).Visible = False
    Image1a(0).Visible = True
    SetTopMostWindow Me.hwnd, True
Else
    Image1a(0).Visible = False
    Image1a(1).Visible = True
    SetTopMostWindow Me.hwnd, False
End If

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label6_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mnuSendPDF_Click
End Sub

Private Sub mnu_send_Click()
    On Error Resume Next
    txt_status.Text = ""
    FileCopy App.Path & "\test.pdf", App.Path & "\Sent Faxes\" & txt_attach.Text
    If form_errors = False Then
        connect_to_smtp_server txt_smtp_server
    End If
End Sub
Public Sub Command2_Click()
On Error Resume Next
  Dim filebmp, Holder As String
  Dim filejpg As String
  Dim si As String
  Dim c As New cDIBSection
  Dim qual1
File1.Pattern = "*.bmp"
File1.Refresh
List1.Clear
For i = 0 To File1.ListCount - 1
    List1.AddItem File1.List(i)
Next
  
    For i = 0 To List1.ListCount - 1
        filebmp = List1.List(i)
        Picture2.Picture = LoadPicture(filebmp)
        filejpg = filebmp & ".jpg"
        SetAttr filejpg, vbNormal
        si = filejpg 'fileToSave
        c.CreateFromPicture Picture2.Picture
        qual1 = "50"
        SaveJPG c, si, qual1
    Next i
Kill App.Path & "\*.bmp"
File1.Pattern = "*.jpg"
File1.Refresh
List1.Clear
For i = 0 To File1.ListCount - 1
    List1.AddItem File1.List(i)
Next
Pdfit

End Sub

Private Sub mnuBright_Click()
txt_smtp_server = "smtp-server"
End Sub

Public Sub mnuCaptureUrl_Click()
On Error Resume Next
Kill App.Path & "\*.bmp"
Me.Hide
Delayz 1
Load frmCapture
frmCapture.Show
End Sub

Private Sub mnuCom_Click()
txt_smtp_server = "smtp.comcast.net"

End Sub

Private Sub mnuEarth_Click()
txt_smtp_server = "smtp.earthlink.net"
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHide_Click()
    Me.Visible = False
End Sub

Private Sub mnuJuno_Click()
txt_smtp_server = "smtp.juno.com"

End Sub

Private Sub mnuNet_Click()
txt_smtp_server = "smtp.netzero.com"

End Sub
Private Sub mnuMSN_Click()
txt_smtp_server = "smtp.email.msn.com"
End Sub

Private Sub mnuOpenSaved_Click()
      If shlShell Is Nothing Then
          Set shlShell = New shell32.Shell
      End If
    shlShell.Explore (App.Path & "\Sent Faxes")

End Sub

Private Sub mnuOther_Click()
'On Error Resume Next
StartDoc App.Path & "\POP3 Incoming SMTP Outgoing Mail Servers.pdf"
End Sub

Private Sub mnuProdigy_Click()
txt_smtp_server = "smtp.prodigy.net"
End Sub

Private Sub mnuRestore_Click()
    Me.Visible = True
End Sub

Private Sub mnuScan_Click()
On Error Resume Next
Dim intSave As Integer, nPixTypes As Long
Screen.MousePointer = 11
If Dir(App.Path & "\test.pdf") <> "" Then
    Kill App.Path & "\test.pdf"
End If
If Dir(App.Path & "\*.jpg") <> "" Then
    Kill App.Path & "\*.jpg"
End If
On Error GoTo BadScan
'GoTo here1
here:
If Clipboard.GetText() <> "" Then
    ClippedText = Clipboard.GetText()
End If
If Clipboard.GetData(vbCFBitmap) <> 0 Then
    picTest.Picture = Clipboard.GetData(vbCFBitmap)
End If
Clipboard.Clear
txt_attach.Text = Format(Now, "ddmmyyhhmmss") & ".pdf"
'eztwain will not save directly to a file for some reason. Thus I have to use the clipboard
'to a picture box to save a bmp. I save the Clipboard and restore it after

    If TWAIN_AcquireToClipboard(Me.hwnd, nPixTypes) = 0 Then
        MsgBox "No image was acquired or transfer to the clipboard failed.", vbInformation, ""
    Else
        Picture2.Picture = Clipboard.GetData
        SavePicture Picture2.Image, App.Path & "\" & Format(Now, "ddmmyyhhmmss") & ".bmp"
    End If


    intSave = MsgBox("Do you want to Scan another page?", _
                     vbYesNoCancel + vbExclamation)
    Select Case intSave
        Case vbYes
            GoTo here
        Case vbNo
            Command2_Click
        Case vbCancel
          Exit Sub
    End Select
Clipboard.Clear
If Trim(ClippedText) <> "" Then
    Clipboard.SetText ClippedText
Else
    Clipboard.SetData picTest.Picture
End If

Screen.MousePointer = 0
Exit Sub

BadScan:
Clipboard.Clear
If Trim(ClippedText) <> "" Then
    Clipboard.SetText ClippedText
Else
    Clipboard.SetData picTest.Picture
End If

MsgBox "Scan has been aborted", vbInformation, ""
Screen.MousePointer = 0

End Sub

Private Sub Pdfit()
On Error Resume Next
    Kill App.Path & "\test.pdf"
    ' Create a simple PDF file using the mjwPDF class
    Dim objPDF As New mjwPDF
    
    ' Set the PDF title and filename
    objPDF.PDFTitle = txt_subject.Text
    objPDF.PDFFileName = App.Path & "\test.pdf"
    
    ' We must tell the class where the PDF fonts are located
    objPDF.PDFLoadAfm = App.Path & "\Fonts"
    ' Set the file properties
    objPDF.PDFSetLayoutMode = LAYOUT_DEFAULT
    objPDF.PDFFormatPage = FORMAT_A4
    objPDF.PDFOrientation = ORIENT_PORTRAIT
    objPDF.PDFSetUnit = UNIT_PT
    
    ' Lets us set see the bookmark pane when we view the PDF
    objPDF.PDFUseOutlines = True
    
    ' View the PDF file after we create it
    objPDF.PDFView = False
    
    ' Begin our PDF document
    objPDF.PDFBeginDoc
        ' Lets add a heading
        objPDF.PDFSetFont FONT_ARIAL, 15, FONT_BOLD
        objPDF.PDFSetDrawColor = vbWhite
        objPDF.PDFSetTextColor = vbWhite
        objPDF.PDFSetAlignement = ALIGN_CENTER
        objPDF.PDFSetBorder = BORDER_ALL
        objPDF.PDFSetFill = True
        objPDF.PDFCell "  ", 1, 1, _
            1, 1
        
    

For i = 0 To List1.ListCount - 1
        objPDF.PDFImage App.Path & "\" & List1.List(i), 0, 0, 576, 720, "http://www.vb6.us"
        objPDF.PDFEndPage
        If i <> List1.ListCount - 1 Then objPDF.PDFNewPage
Next
    objPDF.PDFFileName = ""
                             
    objPDF.PDFEndDoc
i = 0
Do While Dir(App.Path & "\test.pdf") = ""
    Delayz 1
    If i = 3 Then MsgBox "PDF file creation ERROR!!": Exit Do: Exit Sub
    i = i + 1
Loop
intSave = MsgBox("Do you want to Preview this fax?", _
                 vbYesNoCancel + vbExclamation)
If intSave = vbYes Then
   StartDoc App.Path & "\test.pdf"
End If

End Sub

Private Sub mnuSelSource_Click()
TWAIN_SelectImageSource (Me.hwnd)

End Sub

Private Sub mnuSendPDF_Click()
On Error Resume Next
    txt_status.Text = ""
    Dim f As String, intSave As Integer
    With cdlOpen
      .InitDir = App.Path
      .Filter = "Fax Files (*.pdf)|*.pdf"
      .CancelError = True
      .ShowOpen
    End With
    If cdlOpen.FileName <> "" Then
        Kill App.Path & "\test.pdf"
        FileCopy cdlOpen.FileName, App.Path & "\test.pdf"
        txt_attach.Text = Format(Now, "ddmmyyhhmmss") & ".pdf"
    i = 0
    Do While Dir(App.Path & "\test.pdf") = ""
        Delayz 1
        If i = 3 Then MsgBox "PDF file creation ERROR!!": Exit Do: Exit Sub
        i = i + 1
    Loop
        intSave = MsgBox("Do you want to Preview this fax?", _
                         vbYesNoCancel + vbExclamation)
        If intSave = vbYes Then
           StartDoc App.Path & "\test.pdf"
        End If
        
        intSave = MsgBox("Do you want to Send this Fax?", _
                         vbYesNoCancel + vbExclamation)
        If intSave = vbYes Then
            If form_errors = False Then
                connect_to_smtp_server txt_smtp_server
            End If
        End If
    End If

End Sub

Private Sub mnuSendSaved_Click()
On Error Resume Next
    txt_status.Text = ""
    Dim f As String, intSave As Integer
    With cdlOpen
      .InitDir = App.Path & "\Sent Faxes"
      .Filter = "Fax Files (*.pdf)|*.pdf"
      .CancelError = True
      .ShowOpen
    End With
    'cdlOpen.ShowOpen
    If cdlOpen.FileName <> "" Then
        FileCopy cdlOpen.FileName, App.Path & "\test.pdf"
        Open Replace(cdlOpen.FileName, "pdf", "txt") For Input As #1
            Line Input #1, f
            txt_email_to.Text = f
            Line Input #1, f
            txt_subject.Text = f
        Close #1
        txt_attach.Text = Format(Now, "ddmmyyhhmmss") & ".pdf"
        intSave = MsgBox("Do you want to Preview this fax?", _
                         vbYesNoCancel + vbExclamation)
        If intSave = vbYes Then
           StartDoc App.Path & "\test.pdf"
        End If
        
        intSave = MsgBox("Do you want to re-Send this Fax?", _
                         vbYesNoCancel + vbExclamation)
        If intSave = vbYes Then
            If form_errors = False Then
                connect_to_smtp_server txt_smtp_server
            End If
        End If
    End If

End Sub

Private Sub mnuViewFax_Click()
On Error Resume Next
    With cdlOpen
      .InitDir = App.Path & "\Sent Faxes"
      .Filter = "Fax Files (*.pdf;*.txt)|*.pdf;*.txt"
      .CancelError = True
      .ShowOpen
    End With
    'cdlOpen.ShowOpen
    If cdlOpen.FileName <> "" Then
        StartDoc cdlOpen.FileName
    End If

End Sub
Private Sub SelText(t As TextBox)
On Error Resume Next

With t
    .SelStart = 0
    .SelLength = Len(t.Text)
End With
End Sub

Private Sub Picture1_Change()

End Sub

Private Sub txt_email_from_Change()
Changed = True
End Sub

Private Sub txt_email_from_Click()
SelText txt_email_from

End Sub

Private Sub txt_email_to_Change()
Changed = True
End Sub

Private Sub txt_email_to_Click()
SelText txt_email_to
End Sub

Private Sub txt_smtp_server_Click()
SelText txt_smtp_server

End Sub

Private Sub txt_status_Change()
    On Error Resume Next
    Dim Fille As String
    Fille = Replace(txt_attach.Text, "pdf", "txt")
    Open App.Path & "\Sent Faxes\" & Fille For Output As #1
        Print #1, txt_email_to & vbCrLf & txt_subject
        Print #1, vbCrLf & txt_message_text.Text & vbCrLf
        Print #1, txt_status.Text
    Close #1
End Sub

Private Sub txt_subject_Change()
Changed = True
End Sub

Private Sub txt_subject_Click()
SelText txt_subject

End Sub

Private Sub Winsock1_Connect()
    txt_status.Text = "Connected to: " & txt_smtp_server & "." & vbCrLf & txt_status.Text & vbCrLf
    send_email txt_email_to, txt_email_from, txt_subject, txt_message_text, txt_attach
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData response
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    txt_status.Text = "Error: " & Description & "." & vbCrLf & txt_status.Text & vbCrLf
End Sub

Private Sub CandyButton7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Load OutlookCnt
'OutlookCnt.Show
Shell App.Path & "\CntrlOutlk.exe", vbNormalFocus
End Sub


