VERSION 5.00
Begin VB.Form frmUpdate 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update English Word Entry"
   ClientHeight    =   3015
   ClientLeft      =   4530
   ClientTop       =   3525
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   90
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4635
   Begin SunOSSDict.vbButton cmdCancel 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmUpdate.frx":1CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtInput 
      Height          =   525
      Left            =   960
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.TextBox txtResult 
      Height          =   1335
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin SunOSSDict.vbButton cmdSave 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmUpdate.frx":1CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblInput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "English :"
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblOutput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Indonesia"
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   870
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit As Boolean
Private Sub cmdCancel_Click()
    If frmSunossdict.mnuEnglish.Checked = True Then
        Rem jika user memilih Edit maka langsung di-close
        If edit = True Then
    '    If (MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Confirmation")) = vbYes Then
            frmSunossdict.mnuEdit.Checked = False
            recordEng.Close
            Unload Me
        Rem jika user tidak memilih Edit (memilih Add) maka user ditanya apakah ingin langsung di-close
        ElseIf edit = False Then
            If (MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Confirmation")) = vbYes Then
                frmSunossdict.mnuAdd.Checked = False
                Unload Me
            Else
                Call clearText
    '            cmdSave.Visible = True
            End If
        End If
    ElseIf frmSunossdict.mnuIndonesia.Checked = True Then
        Rem jika user memilih Edit maka langsung di-close
        If edit = True Then
    '    If (MsgBox("Do you want to close?", vbQuestion + vbYesNo, "Confirmation")) = vbYes Then
            frmSunossdict.mnuEdit.Checked = False
            Unload Me
        Rem jika user tidak memilih Edit (memilih Add) maka user ditanya apakah ingin langsung di-close
        ElseIf edit = False Then
            If (MsgBox("Keluar?", vbQuestion + vbYesNo, "Konfirmasi")) = vbYes Then
                frmSunossdict.mnuAdd.Checked = False
                Unload Me
            Else
                Call clearText
    '            cmdSave.Visible = True
            End If
        End If
    End If
End Sub


Private Sub cmdSave_Click()
On Error Resume Next
    If frmSunossdict.mnuEnglish.Checked = True Then
        If edit = True Then
            Call saveEdit
        Else
            Call saveAdd
        End If
    ElseIf frmSunossdict.mnuIndonesia.Checked = True Then
        If edit = True Then
            Call simpanEdit
        Else
            Call simpanTambah
        End If
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    lstResult.Text = ""
    If frmSunossdict.mnuEnglish.Checked = True Then
        lblInput.Caption = "English"
        lblOutput.Caption = "Indonesian"
        If frmSunossdict.cmdEdit.Value = True Or frmSunossdict.mnuEdit.Checked = True Then
            edit = True
            Me.Caption = "Edit Word Entry"
            txtInput.Text = frmSunossdict.txtSearch.Text
            txtResult.Text = frmSunossdict.txtResult.Text
        ElseIf frmSunossdict.cmdAddNew.Value = True Or frmSunossdict.mnuAdd.Checked = True Then
            edit = False
            Me.Caption = "Add New Word Entry"
        End If
    ElseIf frmSunossdict.mnuIndonesia.Checked = True Then 'If frmSunossdict.mnuEnglish.Checked = false Then
        lblInput.Caption = "Indonesia"
        lblOutput.Caption = "Inggris"
        If frmSunossdict.cmdEdit.Value = True Or frmSunossdict.mnuEdit.Checked = True Then
            edit = True
            Me.Caption = "Perbaharui  Data"
            txtInput.Text = frmSunossdict.txtSearch.Text
            txtResult.Text = frmSunossdict.txtResult.Text
        ElseIf frmSunossdict.cmdAddNew.Value = True Or frmSunossdict.mnuAdd.Checked = True Then
            edit = False
            Me.Caption = "Tambah Data Baru"
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub
Private Sub saveEdit()
    On Error Resume Next
    Rem periksa apakah field kosong (manatau ada user yang usil atau tester:-D)
    If txtInput.Text = "" Or txtInput.Text = "" Then
    Rem or you can do it by using this one :
    Rem If Not IsNull(txtInput.Text) Or Not IsNull(txtDeskripsi.Text) Then
    Rem or:
    Rem If Len(txtInput.Text) = 0 Or Len(txtResult.Text) = 0 Then
            MsgBox "Please fill the Text Box", vbCritical, "Warning"
    Else

            frmSunossdict.Adodc1!English = txtInput.Text
            frmSunossdict.Adodc1!English = txtResult.Text
            frmSunossdict.Adodc1.Recordset.Update
            frmSunossdict.Adodc1.Refresh
            Rem perlu ga bikin message box memberitahu user bahwa update berhasil??
            Rem sepertinya bikin repot donk mengeklik button Ok atau menekan tombol enter
            MsgBox "You have been update your record!", vbInformation, "Information"
            frmSunossdict.txtResult.Text = txtResult.Text
            frmSunossdict.txtSearch.Text = txtInput.Text
            frmSunossdict.lstResult.Refresh
            Unload Me
    End If
End Sub
Private Sub simpanEdit()
'    On Error Resume Next

    If txtInput.Text = "" Or txtInput.Text = "" Then
            MsgBox "Data tidak boleh kosong", vbCritical, "Peringatan"
    Else
            frmSunossdict.Adodc1!English = Trim(txtInput.Text)
            frmSunossdict.Adodc1!Indonesia = Trim(txtResult.Text)
            frmSunossdict.Adodc1.Recordset.Update
            frmSunossdict.Adodc1.Refresh
'            Rem perlu ga bikin message box memberitahu user bahwa update berhasil??
'            Rem sepertinya bikin repot donk mengeklik button Ok atau menekan tombol enter
            MsgBox "Update data berhasil!", vbInformation, "Pemberitahuan"
            frmSunossdict.txtResult.Text = txtResult.Text
            frmSunossdict.txtSearch.Text = txtInput.Text
            frmSunossdict.lstResult.Refresh
            Unload Me
    End If
End Sub
Private Sub saveAdd()
Dim I As Integer
        If Len(txtInput.Text) = 0 Or Len(txtResult.Text) = 0 Then
            MsgBox "Data can't be empty", vbCritical, "Warning!!!"
        Else
            Rem check apakah entry English (sebagai primary key) sudah ada sebelumnnya tersimpan di database
            Rem tapi kayanya kita membuat primary key aja?
            'If frmSunossdict.Adodc1.Recordset.EOF = False And frmSunossdict.Adodc1.Recordset.BOF = False Then
             '   MsgBox "Data 've already exist in database, operation fail!!", vbCritical, "Warning"
            'Else
            I = frmSunossdict.lstResult.ListIndex
                frmSunossdict.Adodc1.Recordset.addNew
                frmSunossdict.Adodc1.Recordset!English = Trim(txtInput.Text)
                frmSunossdict.Adodc1.Recordset!Indonesia = Trim(txtResult.Text)
                frmSunossdict.Adodc1.Recordset.Update
                frmSunossdict.Adodc1.Refresh
                
                frmSunossdict.lstResult.AddItem (txtInput.Text)
'                frmSunossdict.lstResult.SetFocus
                frmSunossdict.lstResult.Refresh
                frmSunossdict.txtResult.Refresh

                MsgBox "Data insertion success!", vbInformation, "Information"
                Call clearText
            'End If
        End If


End Sub
Private Sub simpanTambah()
Dim I As Integer
        If Len(txtInput.Text) = 0 Or Len(txtResult.Text) = 0 Then
            MsgBox "Data tidak boleh kosong", vbCritical, "Warning!!!"
        Else
            Rem check apakah entry English (sebagai primary key) sudah ada sebelumnnya tersimpan di database
            Rem tapi kayanya kita membuat primary key aja?
            'If frmSunossdict.Adodc1.Recordset.EOF = False And frmSunossdict.Adodc1.Recordset.BOF = False Then
            '    MsgBox "Kata yang Anda masukkan sudah ada dalam enrty database", vbCritical, "Peringatan"
            'Else
            I = frmSunossdict.lstResult.ListIndex
                frmSunossdict.Adodc1.Recordset.addNew
                frmSunossdict.Adodc1.Recordset!Indonesia = Trim(txtInput.Text)
                frmSunossdict.Adodc1.Recordset!English = Trim(txtResult.Text)
                frmSunossdict.Adodc1.Recordset.Update
                frmSunossdict.Adodc1.Refresh
                
                frmSunossdict.lstResult.AddItem (txtInput.Text)
'                frmSunossdict.lstResult.SetFocus
                frmSunossdict.lstResult.Refresh
                frmSunossdict.txtResult.Refresh

                MsgBox "Penambahan data berhasil!", vbInformation, "Pemberitahuan"
                Call clearText
            'End If
        End If
End Sub
Public Sub clearText()
    txtInput.Text = ""
    txtResult.Text = ""
End Sub
Private Sub addNew()
    If (MsgBox("Add New?", vbQuestion + vbYesNo, "Confirmation")) = vbYes Then
        cmdSave.Visible = True
        Call clearText
    Else
        frmSunossdict.Show
        Unload Me
    End If
End Sub
