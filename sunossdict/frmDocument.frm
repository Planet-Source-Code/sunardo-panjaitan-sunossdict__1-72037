VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2430
   ClientLeft      =   6030
   ClientTop       =   4710
   ClientWidth     =   3930
   HelpContextID   =   770
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2430
   ScaleWidth      =   3930
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   2235
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   3942
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      MaxLength       =   1000
      TextRTF         =   $"frmDocument.frx":1CCA
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub rtfText_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim sFile As String: sFile = Data.Files(1)
'   Me.Caption = sFile
   On Error GoTo fError
   Open sFile For Input As #1
    rtfText.Text = Input$(LOF(1), #1)
   Close #1
   ' Jika berhasil membuka file, keluar dari Sub
   Exit Sub
' Jika error kosongkan TextBox
fError: Close #1
   MsgBox "Could not read file"
  rtfText = vbNullString
End Sub

Private Sub rtfText_SelChange()
    frmMain.tbToolBar.Buttons("Bold").Value = IIf(rtfText.SelBold, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Italic").Value = IIf(rtfText.SelItalic, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Underline").Value = IIf(rtfText.SelUnderline, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Align Left").Value = IIf(rtfText.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Center").Value = IIf(rtfText.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Align Right").Value = IIf(rtfText.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
End Sub

Private Sub Form_Load()
    Form_Resize
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    rtfText.RightMargin = rtfText.Width - 400
End Sub

'    rtfText.SelText = UCase(rtfText.SelText)
'    rtfText.SelText = LCase(rtfText.SelText)
'    rtfText.SelText = UCase(Left(rtfText.SelText))
    
