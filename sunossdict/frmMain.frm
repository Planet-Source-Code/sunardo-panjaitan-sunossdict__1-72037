VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SunOSSDict Simple Word Processing"
   ClientHeight    =   3690
   ClientLeft      =   5715
   ClientTop       =   4500
   ClientWidth     =   6525
   HelpContextID   =   510
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Speech"
            Object.ToolTipText     =   "English Text To Speech"
            ImageKey        =   "Speech"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "IndoSpeech"
            Object.ToolTipText     =   "Indonesian Text To Speech"
            ImageKey        =   "Speech"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Spell Check"
            Object.ToolTipText     =   "Spell Check"
            ImageKey        =   "Spell Check"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "Align Left"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "Center"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "Align Right"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   3420
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5847
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "25/12/2011"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "21:05"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1740
      Top             =   1305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1740
      Top             =   1305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CCA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DDC
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EEE
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2000
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2112
            Key             =   "Speech"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27B0
            Key             =   "Spell Check"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28C2
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29D4
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AE6
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BF8
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D0A
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E1C
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F2E
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3040
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3152
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New (File Baru)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open...(Buka File)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close (Tutup File)"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save (Simpan)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As...(Simpan Sebagai)"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up...(Pengaturan Halaman)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print...(Cetak)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit (Tutup Aplikasi)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo (Batalkan)"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t (Potong)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy (Salin)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste (Tempel)"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuEnglish 
      Caption         =   "English Grammatical"
      Begin VB.Menu mnuTTS 
         Caption         =   "English Text To Speech [TTS] "
         Begin VB.Menu mnuSelText 
            Caption         =   "Selected Text"
         End
      End
      Begin VB.Menu mnuEngbar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIndoTTS 
         Caption         =   "Indonesian Text To Speech"
         Begin VB.Menu mnuTerpilihText 
            Caption         =   "Text Terpilih"
         End
      End
      Begin VB.Menu mnInBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpellCheck 
         Caption         =   "English Spelling Check"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuEngbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEngInd 
         Caption         =   "English-Indonesian Dictionary"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuIndIng 
         Caption         =   "Kamus Indonesia-Inggris"
         Shortcut        =   +{F2}
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "Format"
      Begin VB.Menu mnuTextFormat 
         Caption         =   "Text Format"
         Begin VB.Menu mnuBold 
            Caption         =   "Bold"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuItalic 
            Caption         =   "Italic"
         End
         Begin VB.Menu mnuUnderline 
            Caption         =   "Underline"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuStrikeThru 
            Caption         =   "Strike Through"
         End
      End
      Begin VB.Menu mnuCase 
         Caption         =   "Change Case"
         Begin VB.Menu mnuLowerCase 
            Caption         =   "lower case"
         End
         Begin VB.Menu mnuUpperCase 
            Caption         =   "UPPER CASE"
         End
      End
      Begin VB.Menu mnuFontItem 
         Caption         =   "Font"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuIsi 
         Caption         =   "Content (Isi)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About (Tentang)"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
'    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
'    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    LoadNewDoc
End Sub

Private Sub LoadNewDoc()
    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
    frmD.Caption = "Document " & lDocumentCount
    frmD.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuFilePrintPreview_Click()

End Sub


Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExitTray_Click()
    End
End Sub

Private Sub mnuMaximize_Click()
    frmEnglish.Show
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
    
End Sub

Private Sub mnuBold_Click()
    ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
'    Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
End Sub

Private Sub mnuEditUndo_Click()
    Call SendMessage(Text1.hWnd, EM_UNDO, 0&, 0&)
End Sub

Private Sub mnuEngInd_Click()
    frmSunossdict.Show
End Sub

Private Sub mnuIndIng_Click()
    frmSunossdict.Show
    frmSunossdict.mnuIndonesia_Click
End Sub

Private Sub mnuIsi_Click()
    HTMLHELP_Contents Me
End Sub

Private Sub mnuLowerCase_Click()
    ActiveForm.rtfText.SelText = LCase(ActiveForm.rtfText.SelText)
End Sub

Private Sub mnuSelText_Click()
    Call TextToSpeech(ActiveForm.rtfText.SelText)
End Sub

Private Sub mnuSpellCheck_Click()
    Set frmCheckGrammar.txtSource = ActiveForm.rtfText
    Call frmCheckGrammar.Load_SpellChek(ActiveForm.rtfText)
End Sub

Private Sub mnuItalic_Click()
    ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
'    Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)

End Sub

Private Sub mnuStrikeThru_Click()
'    Case "Strike Through"
    ActiveForm.rtfText.SelStrikeThru = Not ActiveForm.rtfText.SelStrikeThru
'    Button.Value = IIf(ActiveForm.rtfText.SelStrikeThru, tbrPressed, tbrUnpressed)

End Sub

Private Sub mnuTerpilihText_Click()

    'Call TextToSpeech(ActiveForm.rtfText.SelText)
    IndoTTS_Say ActiveForm.rtfText.SelText
End Sub

Private Sub mnuUnderline_Click()
    ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
'    Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)

End Sub

Private Sub mnuUpperCase_Click()
    ActiveForm.rtfText.SelText = UCase(ActiveForm.rtfText.SelText)
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Speech"
           ' Call TextToSpeech(ActiveForm.rtfText.SelText)
           mnuSelText_Click
        Case "IndoSpeech"
            IndoTTS_Say ActiveForm.rtfText.SelText
        Case "Spell Check"
            mnuSpellCheck_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Bold"
            ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
            Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
            Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
            Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Align Left"
            ActiveForm.rtfText.SelAlignment = rtfLeft
        Case "Center"
            ActiveForm.rtfText.SelAlignment = rtfCenter
        Case "Align Right"
            ActiveForm.rtfText.SelAlignment = rtfRight
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub
Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
End Sub
Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub
Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub
Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub
Private Sub mnuFontItem_Click()
    'Force an error if the user clicks Cancel
    dlgCommonDialog.CancelError = True
    On Error GoTo Errhandler:
    'Set flags for special effects and all available fonts
    dlgCommonDialog.Flags = cdlCFEffects Or cdlCFBoth
    'Display font dialog box
    dlgCommonDialog.ShowFont
    'Set formatting properties with user selections:
    ActiveForm.rtfText.SelFontName = dlgCommonDialog.FontName
    ActiveForm.rtfText.SelFontSize = dlgCommonDialog.FontSize
    ActiveForm.rtfText.SelColor = dlgCommonDialog.Color
    ActiveForm.rtfText.SelBold = dlgCommonDialog.FontBold
    ActiveForm.rtfText.SelItalic = dlgCommonDialog.FontItalic
    ActiveForm.rtfText.SelUnderline = dlgCommonDialog.FontUnderline
    ActiveForm.rtfText.SelStrikeThru = dlgCommonDialog.FontStrikethru

    
Errhandler:
    'exit procedure if the user clicks Cancel
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtfText.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF
    ActiveForm.rtfText.SelText = vbNullString

End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    
    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtfText.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hdc
        End If
    End With

End Sub

'Private Sub mnuFilePrintPreview_Click()
'    On Error Resume Next
'    With dlgCommonDialog
'        .DialogTitle = "Preview"
''        .CancelError = True
'        .Parent
'    End With
'End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileSaveAs_Click()
    Dim sFile As String
    

    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Save As"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.Caption = sFile
    ActiveForm.rtfText.SaveFile sFile

End Sub

Private Sub mnuFileSave_Click()
    Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Document" Then
        With dlgCommonDialog
            .DialogTitle = "Save"
            .CancelError = False
            .Filter = "Document(*.doc)|* .DOC"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfText.SaveFile sFile
    Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub

Private Sub mnuFileClose_Click()
    Unload Screen.ActiveForm
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    If ActiveForm Is Nothing Then LoadNewDoc
    

    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.rtfText.LoadFile sFile
    ActiveForm.Caption = sFile

End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub



