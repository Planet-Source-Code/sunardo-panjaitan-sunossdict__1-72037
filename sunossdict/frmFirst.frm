VERSION 5.00
Begin VB.Form frmFirst 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   2985
   ClientLeft      =   5445
   ClientTop       =   4050
   ClientWidth     =   3975
   HelpContextID   =   430
   Icon            =   "frmFirst.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   3120
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3720
      Top             =   480
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   780
      Left            =   120
      Picture         =   "frmFirst.frx":1CCA
      ScaleHeight     =   720
      ScaleMode       =   0  'User
      ScaleWidth      =   720
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   780
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   15
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Label lblContact 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E-mail/YM : sunardo_panjaitan@yahoo.com"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "©2009 by Sunardo Panjaitan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label lblApp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SunOSSDict v2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   2340
   End
   Begin VB.Label lblSunossdict 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ind"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   3735
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const LWA_BOTH = 3
Const LWA_ALPHA = 2
Const LWA_COLORKEY = 1
Const GWL_EXSTYLE = -20
Const WS_EX_LAYERED = &H80000

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal Color As Long, ByVal X As Byte, _
ByVal alpha As Long) As Boolean
Dim intTrans As Integer
Dim objPic As Picture
Dim status As Boolean
Dim hitung As Integer

Sub Transparan(hWndTrans As Long, Transp As Integer)
On Error Resume Next

Dim OK As Long
OK = GetWindowLong(hWndTrans, GWL_EXSTYLE)

SetWindowLong hWndTrans, GWL_EXSTYLE, OK Or WS_EX_LAYERED
SetLayeredWindowAttributes hWndTrans, RGB(255, 255, 0), Transp, LWA_ALPHA
Exit Sub

End Sub

Private Sub Form_Load()

Timer1.Enabled = True
    'Set objPic = LoadPicture(App.Path & "\image\Bliss.jpg")
    'Me.Picture = objPic
    lblSunossdict.Caption = "Inggris - Indonesia, Indonesia - Inggris," & vbNewLine & "Text To Speech (TTS), Spelling Check," & vbNewLine & "Word Processing, Open Source "
    lblCopyright.Caption = "©2009 by Sunardo Panjaitan"
    lblContact.Caption = "E-mail/YM  : sunardo_panjaitan@yahoo.com" & vbNewLine & _
    "WordPress : http://sunardo.wordpress.com"
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    intTrans = intTrans + 5
    'If intTrans > 255 Then intTrans = 255: frmSunossdict.Show: Timer1.Enabled = False
    If intTrans > 255 Then intTrans = 255:  Timer1.Enabled = False: 'Timer2.Enabled = True
    Transparan Me.hWnd, intTrans

    Me.Show
    hitung = hitung + 1
    If hitung = 40 Then
        Timer2.Enabled = True
    End If

End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    intTrans = intTrans - 5
    If intTrans < 10 Then intTrans = 10: 'Timer2.Enabled = False: End
    Transparan Me.hWnd, intTrans
    If intTrans < 5 Then intTrans = 5: 'Timer2.Enabled = False: End
    frmSunossdict.Show
    'Unload Me
    'Load frmSunossdict
    If intTrans < 0 Then intTrans = 0: Timer2.Enabled = False: End
    'frmSunossdict.Show
    Unload Me
End Sub



