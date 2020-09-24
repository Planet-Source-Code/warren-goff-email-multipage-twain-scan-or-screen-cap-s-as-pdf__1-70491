VERSION 5.00
Begin VB.Form OutlookCnt 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Outlook Contacts"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3420
   Icon            =   "OutlookCnt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   75
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   645
      Width           =   3195
   End
   Begin Project1.CandyButton CandyButton1 
      Height          =   915
      Left            =   600
      TabIndex        =   9
      Top             =   1305
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "View Contact"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "OutlookCnt.frx":0442
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
   Begin VB.ComboBox cboContact 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   105
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   630
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   4320
      TabIndex        =   0
      Top             =   570
      Width           =   375
   End
   Begin Project1.CandyButton CandyButton6 
      Height          =   360
      Left            =   2850
      TabIndex        =   8
      Top             =   75
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
   Begin Project1.ocxFormShape ocxFormShape1 
      Left            =   3165
      Top             =   2340
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin VB.ComboBox cboEntryID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   645
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   165
      Picture         =   "OutlookCnt.frx":0D1C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblCount 
      Height          =   195
      Index           =   1
      Left            =   4845
      TabIndex        =   6
      Top             =   1605
      Width           =   1995
   End
   Begin VB.Label lblLocation 
      Caption         =   "Location:"
      Height          =   195
      Left            =   4410
      TabIndex        =   5
      Top             =   765
      Width           =   3435
   End
   Begin VB.Label lblContacts 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   660
      TabIndex        =   4
      Top             =   180
      Width           =   2040
   End
   Begin VB.Label lblCount 
      Caption         =   "Contacts in folder:"
      Height          =   195
      Index           =   0
      Left            =   4755
      TabIndex        =   3
      Top             =   1050
      Width           =   1335
   End
   Begin VB.Label lblPath 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4860
      TabIndex        =   1
      Top             =   1215
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.Label lblContacts 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contacts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   645
      TabIndex        =   10
      Top             =   180
      Width           =   2040
   End
End
Attribute VB_Name = "OutlookCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BEHIND A FORM
'REQUIREMENTS:
'REFERENCE TO MS OUTLOOK 10.0 OR ABOVE.
'3 COMMAND BUTTONS - CMDCLOSE, CandyButton1, AND CMDBROWSE.
'5 LABELS - LBLCONTACTS, LBLLOCATION, LBLPATH, LBLCOUNT (2).
'1 COMBOBOX - CBOCONTACT - LOADS WITH THE SELECTED FOLDERS CONTACTS
Option Explicit

Private moApp As Outlook.Application
Private moNS As Outlook.NameSpace
Private moFolder As Outlook.MAPIFolder
Private moContactItem As Outlook.ContactItem
Private moDistributionList As Outlook.DistListItem
Private mbClose As Boolean
Dim FirstTime As Boolean

Private Sub CandyButton1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim sContact As String
    
    'SYNC CONTACTS TO ENTRYIDS
    cboEntryID.ListIndex = cboContact.ListIndex
    
    If cboContact.ListIndex <> -1 Then
        sContact = Left$(cboContact.Text, InStr(1, cboContact.Text, "(") - 2)
        If moFolder.Items.Item(sContact).Class = 69 Then
            Set moDistributionList = moFolder.Items.Item(cboEntryID.Text)
            'MAKE SURE THERE ARE NOT TWO CONTACTS WITH THE SAME NAME
            If moDistributionList.EntryID = cboEntryID.Text Then
                moDistributionList.Display True
            Else
                Set moDistributionList = moNS.GetItemFromID(cboEntryID.Text)
                moDistributionList.Display True
            End If
            Set moDistributionList = Nothing
        Else
            Set moContactItem = moFolder.Items.Item(sContact)
            'MAKE SURE THERE ARE NOT TWO CONTACTS WITH THE SAME NAME (DIFFERENT COMPANIES)
            If moContactItem.EntryID = cboEntryID.Text Then
                moContactItem.Display True
            Else
                Set moContactItem = moNS.GetItemFromID(cboEntryID.Text)
                moContactItem.Display True
            End If
            Set moContactItem = Nothing
        End If
    End If
End Sub

Private Sub CandyButton6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePointer = vbHourglass
    Unload Me

End Sub

Private Sub cmdClose_Click()
End Sub

Private Sub cmdBrowse_Click()

    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    'SELECT A CONTACT FOLDER TO LIST (PRIVATE OR PUBLIC)
    Set moFolder = moNS.PickFolder
    If TypeName(moFolder) = "Nothing" Then
        Screen.MousePointer = vbNormal
        Me.SetFocus
        Exit Sub
    ElseIf moFolder.DefaultItemType <> olContactItem Then
        'MsgBox "'Contact' type folders only!", vbOKOnly + vbExclamation, App.ProductName
        Screen.MousePointer = vbNormal
        Me.SetFocus
        Unload Me
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Me.SetFocus
    lblPath.Caption = moFolder.FolderPath
    lblCount(1).Caption = moFolder.Items.Count
    cboContact.Clear
    cboEntryID.Clear
    If moFolder.Items.Count > 0 Then
        For i = 1 To moFolder.Items.Count
            DoEvents
            If moFolder.Items.Item(i).Class = 69 Then
                Set moDistributionList = moFolder.Items.Item(i)
                cboContact.AddItem moDistributionList.DLName & " (Distributiion List)"
                cboContact.ItemData(i - 1) = 0
                cboEntryID.AddItem moDistributionList.EntryID
                Set moDistributionList = Nothing
            Else
                Set moContactItem = moFolder.Items.Item(i)
                cboContact.AddItem moContactItem.FullName & " (" & moContactItem.CompanyName & ")"
                Combo1.AddItem moContactItem.FullName & " (" & moContactItem.CompanyName & ")"
                cboContact.ItemData(i - 1) = 1
                cboEntryID.AddItem moContactItem.EntryID
                Set moContactItem = Nothing
            End If
        Next
        cboContact.ListIndex = 0
        Combo1.ListIndex = 0
    Else
        CandyButton1.Enabled = False
    End If
    SetTopMostWindow Me.hwnd, True
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Combo1_Click()
                cboContact.Text = Combo1.Text
                cboContact.Refresh
                cboEntryID.ListIndex = cboContact.ListIndex

End Sub

Private Sub Form_Activate()
    SetTopMostWindow Me.hwnd, True
    SetTopMostWindow Me.hwnd, False
If FirstTime = True Then
    FirstTime = False
    cmdBrowse_Click
End If

End Sub

Private Sub Form_Load()
    
    On Error GoTo No_Bugs
    mbClose = False
    FirstTime = True
    'ATTACH TO OUTLOOK IF RUNNING - IF NOT CREATE NEW
    Set moApp = GetObject(, "Outlook.Application")
    If TypeName(moApp) = "Nothing" Then
        Set moApp = New Outlook.Application
        mbClose = True
    End If
    'GET NAMESPACE AND OUTLOOK VERSION
    Set moNS = moApp.GetNamespace("MAPI")
    If InStr(1, moApp.Version, "10.") > 1 Then
        MsgBox "Unsupported Outlook version!", vbOKOnly + vbExclamation, App.ProductName
        Set moNS = Nothing
        Set moApp = Nothing
    End If
    'OutlookCnt.Caption = App.ProductName
    Exit Sub
    
No_Bugs:
    If Err.Number = 429 Then
        MsgBox Err.Number & " - " & Err.Description, vbOKOnly + vbExclamation, App.ProductName
        Resume Next
    Else
        MsgBox Err.Number & " - " & Err.Description, vbOKOnly + vbExclamation, App.ProductName
    End If
   
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set moContactItem = Nothing
    Set moFolder = Nothing
    Set moNS = Nothing
    If mbClose = True And TypeName(moApp) <> "Nothing" Then
        moApp.Quit
    End If
    Set moApp = Nothing
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub lblContacts_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub
