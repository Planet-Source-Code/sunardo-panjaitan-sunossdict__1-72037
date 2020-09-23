VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmIrregularVerb 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SunOSSDict - Irregular Verb"
   ClientHeight    =   5265
   ClientLeft      =   6450
   ClientTop       =   3045
   ClientWidth     =   6270
   HelpContextID   =   930
   Icon            =   "frmIrregularVerb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6270
   Begin SunOSSDict.vbButton cmdLast 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Last  >>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "frmIrregularVerb.frx":1CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   5400
      Top             =   2040
   End
   Begin MSDataGridLib.DataGrid DataGridIrregularVerb 
      Height          =   1725
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3043
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   0   'False
      ForeColor       =   4210752
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
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
   Begin VB.TextBox txtSingular 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtVing 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   1920
      Width           =   2415
   End
   Begin VB.ComboBox cmbSearch 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Text            =   "Present"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtBentuk3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtBentuk2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtBentuk1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin SunOSSDict.vbButton cmdFirst 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "<< First"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "frmIrregularVerb.frx":1CE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SunOSSDict.vbButton cmdPrev 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "< Prev"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "frmIrregularVerb.frx":1D02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SunOSSDict.vbButton cmdNext 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Next >"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "frmIrregularVerb.frx":1D1E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5640
      TabIndex        =   20
      Top             =   4800
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List Of Irreguler Verb"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   19
      Top             =   2520
      Width           =   2715
   End
   Begin VB.Label label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total words :"
      Height          =   195
      Left            =   4680
      TabIndex        =   18
      Top             =   4920
      Width           =   915
   End
   Begin VB.Label lblSing 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3rd Person Singular"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1965
   End
   Begin VB.Label lblIng 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Present Participle"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search By :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   0
      Width           =   1035
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Past Participle"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1275
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Past Simple"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Present/ Base Form"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1770
   End
   Begin VB.Menu mnu 
      Caption         =   ""
      NegotiatePosition=   2  'Middle
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmIrregularVerb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset
Dim rec As ADODB.Recordset
Dim query As String

Private Sub cmdFirst_Click()
    rst.MoveFirst
    Call showData
End Sub

Private Sub cmdLast_Click()
    rst.MoveLast
    Call showData
End Sub

Private Sub cmdNext_Click()
    rst.MoveNext
    If rst.EOF = False Then
        Call showData
    Else
        rst.MoveLast
    End If
End Sub

Private Sub cmdPrev_Click()
    rst.MovePrevious
    If rst.BOF = False Then
        Call showData
    Else
        rst.MoveFirst
    End If
End Sub

Private Sub Form_Load()
'    lbl1.Caption = "Present" & vbCrLf & "(Bentuk 1)"
'    lbl2.Caption = "Past" & vbCrLf & "(Bentuk 2)"
'    lbl3.Caption = "Past Participle" & vbCrLf & "(Bentuk 3)"
'    lblSing.Caption = "3rd Person" & vbCrLf & " Singular"
'    lblIng.Caption = "Present Participle/" & vbCrLf & "Gerund"
'    frmIrregularVerb.BackColor = vbRed
'    DataGridIrregularVerb.Caption = "Present" & vbTab & "Past Simple" & vbTab & "Past Participle" & vbTab & "3rd Person Singular" & vbTab & "Present Participle/ Gerund"
    Timer1.Enabled = True
    cmbSearch.AddItem "Present"
    cmbSearch.AddItem "Past Simple"
    cmbSearch.AddItem "Past Participle"


'    lblNote.Caption = "Note :" & vbCrLf & "Present = Bentuk 1" & vbCrLf & "Past = Bentuk 2" & vbCrLf & "Past Participle = Bentuk 3"
    Call Connect
    Set rst = New ADODB.Recordset
    Set rec = New ADODB.Recordset
    rec.Open "select count(*)As Total From IrregularVerb", con, adOpenStatic, adLockOptimistic
'    rst.Open " select Bentuk1 as [Base Form], Bentuk2 as [Past Simple], Bentuk3 as [Past Participle], BentukSing as [3rd Person Singular], BentukVing as [Present Participle/ Gerund] from IrregularVerb", con, adOpenStatic, adLockOptimistic
     rst.Open " select * from IrregularVerb", con, adOpenStatic, adLockOptimistic
    Set DataGridIrregularVerb.DataSource = rst
    DataGridIrregularVerb.Refresh
    
    Call showData
    lblTotal.Caption = rec("Total")
End Sub
Public Sub showData()
    txtBentuk1.Text = rst("Bentuk1")
    txtBentuk2.Text = rst("Bentuk2")
    txtBentuk3.Text = rst("Bentuk3")
    txtSingular.Text = rst("BentukSing")
    txtVing.Text = rst("BentukVing")
    
End Sub






Private Sub Timer1_Timer()
    If lblTotal.ForeColor = vbRed Then
       lblTotal.ForeColor = vbGreen
    ElseIf lblTotal.ForeColor = vbGreen Then
      lblTotal.ForeColor = vbBlue
    ElseIf lblTotal.ForeColor = vbBlue Then
      lblTotal.ForeColor = vbBlack
    Else
       lblTotal.ForeColor = vbRed
    End If
End Sub

Private Sub txtSearch_Change()
    Set rst = New ADODB.Recordset
    If cmbSearch.Text = "Present" Then
        query = "select * from IrregularVerb where Bentuk1 like " & "'%" & txtSearch.Text & "%'"
        rst.Open query, con, adOpenStatic, adLockOptimistic
        Set DataGridIrregularVerb.DataSource = rst
        If rst.EOF = False And rst.BOF = False Then
            Call showData
            DataGridIrregularVerb.Refresh
        Else
            Call clear
            MsgBox "Data tidak ditemukan", vbInformation, "Konfirmasi"
        End If

    ElseIf cmbSearch.Text = "Past Simple" Then
        query = "select * from IrregularVerb where Bentuk2 like " & "'%" & txtSearch.Text & "%'"
        rst.Open query, con, adOpenStatic, adLockOptimistic
        Set DataGridIrregularVerb.DataSource = rst
        If rst.EOF = False And rst.BOF = False Then
            Call showData
            DataGridIrregularVerb.Refresh
        Else
            Call clear
            MsgBox "Data tidak ditemukan", vbInformation, "Konfirmasi"
        End If

    ElseIf cmbSearch.Text = "Past Participle" Then
        query = "select * from IrregularVerb where Bentuk3 like " & "'%" & txtSearch.Text & "%'"
        rst.Open query, con, adOpenStatic, adLockOptimistic
        Set DataGridIrregularVerb.DataSource = rst
        If rst.EOF = False And rst.BOF = False Then
            Call showData
            DataGridIrregularVerb.Refresh
            rst.MoveNext
        Else
            Call clear
            MsgBox "Data tidak ditemukan", vbInformation, "Konfirmasi"
        End If
    End If

End Sub

Private Sub clear()
    txtBentuk1.Text = ""
    txtBentuk2.Text = ""
    txtBentuk3.Text = ""
    txtSingular.Text = ""
    txtVing.Text = ""
End Sub
