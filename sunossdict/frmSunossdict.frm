VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSunossdict 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SunOSSDict - [ English-Indonesian Dictionary ]"
   ClientHeight    =   6180
   ClientLeft      =   7035
   ClientTop       =   3240
   ClientWidth     =   6000
   DrawStyle       =   2  'Dot
   FillColor       =   &H8000000A&
   FillStyle       =   7  'Diagonal Cross
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   170
   Icon            =   "frmSunossdict.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6180
   ScaleMode       =   0  'User
   ScaleWidth      =   6000
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer Timer1 
      Left            =   5280
      Top             =   7680
   End
   Begin VB.CommandButton CmdTerminate 
      Caption         =   "Terminate"
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4560
      Picture         =   "frmSunossdict.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Text To Speech (TTS)"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   375
      Left            =   3720
      TabIndex        =   19
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtLanguage 
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Text            =   "English"
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   1320
      ScaleHeight     =   825
      ScaleWidth      =   2865
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   2895
      Begin VB.Image imgCancelSearch 
         Height          =   300
         Left            =   1440
         Picture         =   "frmSunossdict.frx":73D4
         ToolTipText     =   "Cancel"
         Top             =   360
         Width           =   330
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "No result. Search All databases?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   3975
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgDoSearch 
         Height          =   330
         Left            =   720
         Picture         =   "frmSunossdict.frx":C98E
         ToolTipText     =   "Yes"
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.ListBox list2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmSunossdict.frx":12098
      Left            =   960
      List            =   "frmSunossdict.frx":1209F
      TabIndex        =   15
      Top             =   7560
      Width           =   510
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Search Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   3135
      Begin VB.OptionButton optAny 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Any Word Part"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         ToolTipText     =   "Any Word Part"
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optFirst 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "First Word"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "First  Word"
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optLatest 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Latest Word"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1440
         TabIndex        =   12
         ToolTipText     =   "Latest Word"
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optExact 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Identity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Identity"
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin SunOSSDict.vbButton cmdOutput 
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   1560
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Indonesian :"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSunossdict.frx":120AA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SunOSSDict.vbButton cmdInput 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "List Word :"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSunossdict.frx":120C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmdHelp 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      Picture         =   "frmSunossdict.frx":120E2
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Help"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3960
      Picture         =   "frmSunossdict.frx":176B8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete Entry"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdEdit 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4320
      Picture         =   "frmSunossdict.frx":1CC72
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Edit Entry"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdAddNew 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      Picture         =   "frmSunossdict.frx":22291
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Add New Entry"
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdSpeech 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4920
      Picture         =   "frmSunossdict.frx":27868
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Text To Speech (TTS)"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2040
      Width           =   3495
   End
   Begin VB.ListBox lstResult 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      ItemData        =   "frmSunossdict.frx":2CFB5
      Left            =   120
      List            =   "frmSunossdict.frx":2CFB7
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   5400
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid MSHFlexGrid1 
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      Top             =   7560
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1440
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View Dictionary"
      Begin VB.Menu mnuEnglish 
         Caption         =   "&English"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuBorder 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndonesia 
         Caption         =   "&Indonesia"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnuBorder1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnEnglishGrammatical 
      Caption         =   "English Grammatical"
      Begin VB.Menu mnuIrregularVerb 
         Caption         =   "Irregular Verb"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuBorder2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWordProcessing 
         Caption         =   "Word Processing"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuUpdate 
      Caption         =   "Update Kamus"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Entry"
      End
      Begin VB.Menu mnuBorder4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Entry"
      End
      Begin VB.Menu mnuBorder5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Entry"
      End
   End
   Begin VB.Menu mnuSetting 
      Caption         =   "Setting"
      Begin VB.Menu mnuLayoutForm 
         Caption         =   "Layout"
      End
      Begin VB.Menu mnuFontEnglish 
         Caption         =   "All Text"
      End
      Begin VB.Menu mnuWordList 
         Caption         =   "List Text"
      End
      Begin VB.Menu mnuTextIndonesian 
         Caption         =   "Result Text"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuBorder7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuClose 
      Caption         =   "&Close"
      Begin VB.Menu mnuMinimizeTray 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuMaximize 
         Caption         =   "Maximize"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitTray 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSunossdict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_CLOSE = &H10
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Dim englishForm As Boolean 'set form : value if value = true then form =English, Else form = Indonesia
'Dim Indonesiaform As Boolean
Public Sub AddInTray()
    With nid
    ' Set nilai member UDT
      .hWnd = Me.hWnd
      .uID = Me.Icon
      .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
      .uCallbackMessage = WM_MOUSEMOVE
      .hIcon = Me.Icon.Handle
      .szTip = "SunOSSDict" & vbNullChar
      .cbSize = Len(nid)
    End With
' Tampilkan ikon di System Tray
    Call Shell_NotifyIcon(NIM_ADD, nid)
End Sub

Private Sub RemoveFromTray()
' Hapus ikon dari System Tray.
Call Shell_NotifyIcon(NIM_DELETE, nid)
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    Set MSHFlexGrid1.DataSource = Nothing
    If mnuEnglish.Checked = True Then
        txtLanguage.Text = "English"
    Else
        txtLanguage.Text = "Indonesia"
    End If
    
    If optExact.Value = True Then
        Adodc1.RecordSource = "Select * From " + txtLanguage.Text + " Where " + txtLanguage.Text + " = '" & txtSearch.Text & "'"
    End If
    If optLatest.Value = True Then
        Adodc1.RecordSource = "Select * From " + txtLanguage.Text + " Where " + txtLanguage.Text + " LIKE '%" & txtSearch.Text & "';"
    End If
    If optFirst.Value = True Then
        Adodc1.RecordSource = "Select * From " + txtLanguage.Text + " Where " + txtLanguage.Text + " LIKE '" & txtSearch.Text & "%';"
    End If
    If optAny.Value = True Then
        Adodc1.RecordSource = "Select * From " + txtLanguage.Text + " Where " + txtLanguage.Text + " LIKE '%" & txtSearch.Text & "%';"
    End If
    
    
    Adodc1.Refresh
    Set MSHFlexGrid1.DataSource = Adodc1
    
    Dim IrecI As Long
    For IrecI = 0 To MSHFlexGrid1.ApproxCount - 1
        lstResult.AddItem LCase(MSHFlexGrid1.Columns(0).CellValue(MSHFlexGrid1.GetBookmark(IrecI)))
        List2.AddItem MSHFlexGrid1.Columns(1).CellValue(MSHFlexGrid1.GetBookmark(IrecI))
        Next
    If englishForm = True Then
        cmdInput.Caption = "List Word : " + Format(lstResult.ListCount) + " Results."
    Else
        cmdInput.Caption = "Daftar Kata  : " + Format(lstResult.ListCount) + " hasil."
    End If
End Sub



Private Sub cmdSearch_Click()
    txtSearch_KeyPress (13)
End Sub

Private Sub CmdTerminate_Click()
   ' PostMessage FindWindow(vbNullString, "Confirm"), WM_CLOSE, CLng(0), CLng(0)
End Sub

Private Sub Form_Load()
    'Get the application path and verify it ends with "\"
    s = App.Path
    If (Right$(s, 1) <> "\") Then s = s + "\"
    
    'Set the programs help file path
    App.HelpFile = s + "sunossdict.chm"
    
    mnuEnglish.Checked = True
    Call formEnglish
    txtLanguage.Text = "English"
    
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "sunossdict.b" + ";Persist Security Info=False"
    DoEvents
    Adodc1.RecordSource = "SELECT * FROM " + txtLanguage.Text
    DoEvents
    Adodc1.Refresh
    DoEvents
    List2.Height = lstResult.Height

End Sub

Private Sub imgDoSearch_Click()
    Picture1.Visible = False
    DoEvents
    optAny.Value = True
    DoEvents
    Picture1.Visible = True
    optExact.Value = True
    Picture1.Visible = False
    If lstResult.ListCount = 0 Then
        Picture1.Visible = False
    End If
End Sub



Private Sub imgCancelSearch_Click()
        Picture1.Visible = False
End Sub

Private Sub lstResult_Click()
    txtResult.Text = List2.List(lstResult.ListIndex)
    txtSearch.Text = lstResult.Text
End Sub
Private Sub List2_Click()
    txtResult.Text = lstResult.List(List2.ListIndex)
End Sub

Private Sub mnuEnglish_Click()
    mnuEnglish.Checked = True
    Call formEnglish
End Sub
Private Sub formEnglish()
    mnuIndonesia.Checked = False
    mnuEnglish.Checked = True
    englishForm = True
    Me.Caption = "SunOSSDict - [ English-Indonesian Dictionary ]"
    mnuView.Caption = "View Dictionary"
    mnuAdd.Caption = "Add Entry"
    mnuDelete.Caption = "Delete Entry"
    mnuEdit.Caption = "Edit Entry"
    mnuClose.Caption = "Close"
    mnuHelp.Caption = "Help"
    

    cmdSpeech.Visible = True
    cmdSpeech.ToolTipText = "Text To Speech (TTS)"
    cmdAddNew.ToolTipText = "Add New Entry"
    cmdEdit.ToolTipText = "Edit Entry"
    cmdDelete.ToolTipText = "Delete Entry"
    cmdHelp.ToolTipText = "Help"
    cmdInput.Caption = "List Word :"
    cmdOutput.Caption = "Indonesian :"
    txtSearch.Text = ""
End Sub

Private Sub mnuHelpContents_Click()
    HTMLHELP_Contents Me
End Sub

Public Sub mnuIndonesia_Click()
    mnuIndonesia.Checked = True
    Call formIndonesia
    On Error Resume Next
        
End Sub
Private Sub formIndonesia()
    englishForm = False
    mnuEnglish.Checked = False
    Me.Caption = "SunOSSDict - [ Kamus Indonesia - Inggris ]"
    
    mnuView.Caption = "Lihat Kamus"
    mnuAdd.Caption = "Tambah Entry"
    mnuDelete.Caption = "Hapus Entry"
    mnuEdit.Caption = "Perbaharui Entry"
    mnuClose.Caption = "Tutup"
    mnuHelp.Caption = "Bantuan"

    cmdAddNew.ToolTipText = "Tambah Entry"
    cmdEdit.ToolTipText = "Perbaharui Entry"
    cmdDelete.ToolTipText = "Hapus Entry"
    cmdHelp.ToolTipText = "Bantuan"
    cmdInput.Caption = "Daftar Kata :"
    cmdOutput.Caption = "Inggris :"
    txtSearch.Text = ""
    
End Sub
Private Sub mnuDeleteEnglish_Click()
    Call cmdDelete_Click
End Sub


Public Sub mnuFontEnglish_Click()
    On Error Resume Next
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.Flags = cdlCFEffects Or cdlCFBoth
    dlgCommonDialog.ShowFont
'    Call frmSunossdictFormat
    Call lstResultFormat
    Call txtOutputFormat
    Call txtSearchFormat
End Sub

Private Sub mnuMinimizeTray_Click()
    Call minimize
End Sub


Private Sub mnuTextIndonesia_Click()
    On Error Resume Next
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.Flags = cdlCFEffects Or cdlCFBoth
    dlgCommonDialog.ShowFont
    Call txtOutputFormat
End Sub

Private Sub mnuTextIndonesian_Click()
    On Error Resume Next
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.Flags = cdlCFEffects Or cdlCFBoth
    dlgCommonDialog.ShowFont
    Call txtOutputFormat
End Sub

Private Sub mnuWordList_Click()
    On Error Resume Next
    dlgCommonDialog.CancelError = True
    dlgCommonDialog.Flags = cdlCFEffects Or cdlCFBoth
    dlgCommonDialog.ShowFont
    Call lstResultFormat
End Sub

Private Sub mnuWordProcessing_Click()
    frmMain.Show
End Sub

Private Sub cmdEdit_Click()
    'If txtSearch.Text = "" Or txtResult.Text = "" Then
    If lstResult.ListCount = 0 Then
        'do nothing
    Else
        frmUpdate.Show vbModal, Me
    End If
End Sub
Private Sub mnuExit_Click()
    End
End Sub

Public Sub mnuExitTray_Click()
    End
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuIrregularVerb_Click()
    frmIrregularVerb.Show vbModal, Me
End Sub

Private Sub mnuLayoutForm_Click()
    Me.dlgCommonDialog.CancelError = True
   On Error Resume Next
   With Me.dlgCommonDialog
      .ShowColor
      ' Ubah background form
      warna = .Color

      Me.BackColor = warna
   End With
   Exit Sub

End Sub

Public Sub mnuMaximize_Click()
    frmSunossdict.Show
End Sub

Public Sub minimize()
    Call AddInTray
''    For Each Control In Form
'
     Me.Hide
'     frmUpdate.Hide
'     frmCheckGrammar.Hide
'     frmAbout.Hide
End Sub
Private Sub cmdSpeech_Click()
' SetTimer hWnd, NV_CLOSEMSGBOX, 1, AddressOf TimerProc
    If englishForm = True Then
        Call TextToSpeech(txtSearch.Text)
    ElseIf englishForm = False Then
    'PostMessage FindWindow(vbNullString, “Confirm”), WM_CLOSE, CLng(0), CLng(0)
   
        IndoTTS_Say txtSearch.Text
        'cmdSpeech.ToolTipText = "You're in Indonesia - Inggris  mode. Change it into English-Indonesian mode"
    End If
    
    'If Indonesiaform = True Then
        'IndoTTS_Say txtSearch.Text
    'Else
        'cmdSpeech.ToolTipText = "Kamu dalam mode English-Indonesian. Ganti menjadi mode Indonesia - Inggris"
    'End If
End Sub
Private Sub cmdAddNew_Click()
    frmUpdate.Show vbModal, Me
End Sub

Private Sub cmdDelete_Click()
    If englishForm = True Then
        Call deleteEnglish
    Else
        Call hapusIndonesia
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim sngMsg As Single
  sngMsg = X / Screen.TwipsPerPixelX
  ' Jika diklik kanan, tampilkan popup
  If sngMsg = WM_RBUTTONUP Then
     Me.PopupMenu Me.mnu
  End If
End Sub

Private Sub mnuHide_Click()
Me.Hide
End Sub

Private Sub mnuAdd_Click()
     frmUpdate.Show vbModal, Me
   '  mnuAdd.Checked = True
End Sub

Private Sub mnuDelete_Click()
    Call deleteEnglish
End Sub

Private Sub mnuEdit_Click()
    frmUpdate.Show vbModal, Me
    'mnuEdit.Checked = True
End Sub
Private Sub cmdHelp_Click()
    HTMLHELP_Contents Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Call RemoveFromTray
    End
End Sub
Private Sub optExact_Click()
    If Not Picture1.Visible = True Then
        txtSearch.SetFocus
        txtSearch_KeyPress (13)
    End If
End Sub

Private Sub optLatest_Click()
    txtSearch.SetFocus
    txtSearch_KeyPress (13)
End Sub

Private Sub optFirst_Click()
    txtSearch.SetFocus
    txtSearch_KeyPress (13)
End Sub

Private Sub optAny_Click()
    txtSearch.SetFocus
    txtSearch_KeyPress (13)
End Sub







Private Sub Timer1_Timer()
'    CmdTerminate_Click
End Sub

Private Sub txtSearch_Change()
    If txtSearch.Text = "" Then
        lstResult.clear
        List2.clear
        txtResult.Text = ""
    End If
    If englishForm = True Then
        cmdInput.Caption = "List Word : "
    Else
        cmdInput.Caption = "Daftar Kata : "
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If englishForm = True Then
        txtLanguage.Text = "English"
    Else
        txtLanguage.Text = "Indonesia"
    End If
    
    If KeyAscii = 13 And Not txtSearch.Text = "" Then
        txtResult.Text = ""
        Picture1.Visible = False
        DoEvents
        lstResult.clear
        List2.clear
        'txtLanguage.Text = "English"
        cmdFind_Click
        txtSearch.Text = txtSearch.Text + "."
        cmdFind_Click
        txtSearch.Text = Left(txtSearch.Text, Len(txtSearch.Text) - 1)
        'txtLanguage.Text = "Indonesia"
        'cmdFind_Click
            If lstResult.ListCount > 0 Then
                lstResult.ListIndex = 0
                lstResult_Click
                Picture1.Visible = False
            Else
                Picture1.Visible = True
            End If
        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch.Text)
    End If
    
    
End Sub



Private Sub txtOutputFormat()
    txtOutput.FontBold = dlgCommonDialog.FontBold
    txtOutput.FontItalic = dlgCommonDialog.FontItalic
    txtOutput.FontName = dlgCommonDialog.FontName
    txtOutput.FontSize = dlgCommonDialog.FontSize
    txtOutput.FontStrikethru = dlgCommonDialog.FontStrikethru
    txtOutput.FontUnderline = dlgCommonDialog.FontUnderline
    txtOutput.ForeColor = dlgCommonDialog.Color
End Sub
Public Sub txtSearchFormat()
    txtSearch.FontBold = dlgCommonDialog.FontBold
    txtSearch.FontItalic = dlgCommonDialog.FontItalic
    txtSearch.FontName = dlgCommonDialog.FontName
    txtSearch.FontSize = dlgCommonDialog.FontSize
    txtSearch.FontStrikethru = dlgCommonDialog.FontStrikethru
    txtSearch.FontUnderline = dlgCommonDialog.FontUnderline
    txtSearch.ForeColor = dlgCommonDialog.Color
End Sub
Public Sub lstResultFormat()
    lstResult.FontBold = dlgCommonDialog.FontBold
    lstResult.FontItalic = dlgCommonDialog.FontItalic
    lstResult.FontName = dlgCommonDialog.FontName
    lstResult.FontSize = dlgCommonDialog.FontSize
    lstResult.FontStrikethru = dlgCommonDialog.FontStrikethru
    lstResult.FontUnderline = dlgCommonDialog.FontUnderline
    lstResult.ForeColor = dlgCommonDialog.Color
End Sub

Private Sub deleteEnglish()
    If lstResult.ListCount = 0 Then
        'do nothing
    Else
    Dim indeks As Integer
    indeks = lstResult.ListIndex
        If (MsgBox("Are You sure to delete this entry?", vbYesNo + vbQuestion, "Confirmation") = vbYes) Then
           ' Adodc1.RecordSource = "delete * From " + txtLanguage.Text + " Where " + txtLanguage.Text + " = '" & lstResult.Selected & "'"
                Adodc1.Recordset.Delete
                Adodc1.Refresh
                lstResult.RemoveItem (indeks)
                MsgBox "You've deleted the entry", vbInformation, "Information"
        End If
    End If
End Sub

Private Sub hapusIndonesia()
    If lstResult.ListCount = 0 Then
        'do nothing
    Else
    Dim I As Integer
    I = lstResult.ListIndex
        If (MsgBox("Yakin akan menghapus data ini?", vbYesNo + vbQuestion, "Konfirmasi") = vbYes) Then
                Adodc1.Recordset.Delete
                Adodc1.Refresh
                lstResult.RemoveItem (I)
                MsgBox "Penghapusan data berhasil", vbInformation, "Informasi"
        End If
    End If
End Sub

