VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCheckGrammar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spelling - English U.S."
   ClientHeight    =   3765
   ClientLeft      =   6660
   ClientTop       =   4170
   ClientWidth     =   5400
   HelpContextID   =   390
   Icon            =   "frmCheckGrammar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5400
   Begin SunOSSDict.vbButton cmdIgnore 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Ignore"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "frmCheckGrammar.frx":1CCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstSuggestions 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   3615
   End
   Begin RichTextLib.RichTextBox rtfActual 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2355
      _Version        =   393217
      ScrollBars      =   1
      TextRTF         =   $"frmCheckGrammar.frx":1CE6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SunOSSDict.vbButton cmdChangeAll 
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Change A&ll"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "frmCheckGrammar.frx":1D6A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SunOSSDict.vbButton cmdIgnoreAll 
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "I&gnore All"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "frmCheckGrammar.frx":1D86
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SunOSSDict.vbButton cmdAdd 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "frmCheckGrammar.frx":1DA2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SunOSSDict.vbButton cmdChange 
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Change"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "frmCheckGrammar.frx":1DBE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin SunOSSDict.vbButton cmdCancel 
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "C&ancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      MICON           =   "frmCheckGrammar.frx":1DDA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Suggestions :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Not in Dictionary :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmCheckGrammar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Spell Check
'Thanks to Rachit K.

'Dim word As word.Application
'Dim doc As Document
'Dim error
'Dim errors
'Dim isError As Boolean
'Dim start As Integer
'Public source As Object
'Dim countWordMove As Integer

Dim objWord As word.Application
Dim doc As Document
Dim SpellErrors                 As SpellingSuggestions
Dim SpellError                  As SpellingSuggestion
Dim curWordCount As Integer     'Global Variable used to Increment the word counter to move to the next WORD..
Dim start As Integer            'Global counter to store the current cursor position over the Character..
Dim isError As Boolean          'Boolean variable to store if any word has spelling mistake, else unloading b4 loading.
Public txtSource As Object      'Variable to store the TextBox object or the SybaTextboxes

Private Sub cmdAdd_Click()
'objWord.
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Public Function SpellCheck()

' Purpose   : Function to chek the Spelling of the word in the string. It loops thru all the word present in the string.
'             Start from the Previous position of the word where it had left after pressing Ignore or Change..
'             curWordCount variable stores the Word positions, start variable stores the position of each character
'---------------------------------------------------------------------------------------
'
Dim eachWrd() As String
Dim cnt As Integer

If InStr(1, rtfActual.Text, " ") > 0 Then
    'If a string has multiple words then store each word in an Array..
    eachWrd = Split(rtfActual.Text, " ")
ElseIf rtfActual.Text <> "" Then
    'If a string contain only of a single word then just store it in 0 th array.
    ReDim eachWrd(0)        'Simply creating a one element array..
    eachWrd(0) = rtfActual.Text
End If

If curWordCount > UBound(eachWrd) Then
'If the word count has reached to the last word then displaying the message & coming out..
    Call ResetRTF
    MsgBox "End of the text..!!", vbExclamation
    Unload Me
    Exit Function
End If

Call ResetRTF

For cnt = curWordCount To UBound(eachWrd)
    'Getting the Spelling suggestion ..
    Set SpellErrors = objWord.GetSpellingSuggestions(word:=eachWrd(cnt))

    'Checking if the Suggestion has returned some Spelling Mistakes..
    If SpellErrors.Count > 0 Then
        isError = True
        curWordCount = cnt                          'Storing the current word position
        Call Highlight_Text(eachWrd(cnt), start)    'Highlighting wrong word with Red Color..
        Exit For
    End If
    start = start + Len(eachWrd(cnt)) + 1           'Incrementing the start with the Length of the current word & 1 more for space..
Next

lstSuggestions.clear
If SpellErrors.Count > 0 Then
    'Adding Spelling suggestion given from SpellErrors object to the list object..
    For Each SpellError In SpellErrors
        lstSuggestions.AddItem SpellError
    Next
    'By default selecting the First Item..
    If lstSuggestions.ListCount > 0 Then lstSuggestions.Selected(0) = True
End If

End Function

Private Sub cmdChange_Click()
Dim str1 As String

'The wrong word may have replace with a right word but may having different length so incrementing the Start variable
'with the length of the new word since it is having the correct length else position will be disturbed..
start = start + Len(lstSuggestions.Text) + 1
curWordCount = curWordCount + 1
'If the ListBox is empty or No Suggestion is there then do nothing.
If lstSuggestions.Text <> "(No Suggestions)" And Trim(lstSuggestions.Text) <> "" Then
    'Replacing the Wrong text occuring first with the Selected Suggestion..

    'Replacing concept is to replace the text only from where the error is detected, not the previous text..
    str1 = rtfActual.Text   'Storing temporary in variable
    'Replacing the text with the Suggested one but from where the current cursor is..
    str1 = Replace(rtfActual.Text, GetHiglighted_Text, lstSuggestions.Text, start - Len(lstSuggestions.Text), 1)
    'Appending the left text of the start to the start..
    rtfActual.Text = VBA.Left$(rtfActual.Text, start - Len(lstSuggestions.Text) - 1) & str1

    'Changing the original text with the changed one..
    txtSource.Text = rtfActual.Text
    Call SpellCheck
End If

End Sub

Private Sub cmdChangeAll_Click()
'Will replace all the Spelling error with the Selected suggestion.. (NOT IN USE presently)
start = start + Len(lstSuggestions.Text) + 1
curWordCount = curWordCount + 1
If lstSuggestions.Text <> "(No Suggestions)" And Trim(lstSuggestions.Text) <> "" Then
    rtfActual.Text = Replace(rtfActual.Text, GetHiglighted_Text, lstSuggestions.Text)
    Call SpellCheck
End If
End Sub

Private Sub cmdIgnore_Click()
'Incrementing the word count without doing anything, also incrementing the start position
curWordCount = curWordCount + 1
start = start + Len(GetHiglighted_Text) + 1
Call SpellCheck

End Sub

Private Sub cmdIgnoreAll_Click()
'Just simply ignoring...
Call cmdCancel_Click
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Closing & Disposing the Word Object while Unloading..
'doc.Close
objWord.Quit
Set objWord = Nothing
End Sub

Public Sub Load_SpellChek(ctlSource As Object)
'Procedure to load the SpellCheck & has to be called from the Calling form..
On Error GoTo LoadSpellError

curWordCount = 0
start = 0
isError = False
rtfActual.Text = ctlSource.Text                 'Assigning the Actual text to the RTF textbox..
Set objWord = CreateObject("Word.Application")  'Create an instance of Word Object
Set doc = objWord.Documents.Add                 'Creating a word document
Call SpellCheck

'If isError is False means no Spell Error where detected in the entire sentence
If isError = False Then
    MsgBox "No spelling errors found.. or No Suggestions..!", vbExclamation
    Unload Me
    Exit Sub
End If

Me.Show vbModal

LoadSpellError:
    If Err.Number <> 0 Then
        MsgBox "Terjadi kesalahan saat me-load Spell Check..!!" & vbCrLf & _
            "Cek apakah ada penulisan kata yang akan di cek..!!" & vbCrLf & "Error : " & Err.Description, vbCritical
        Unload Me
    End If
End Sub

Private Sub Highlight_Text(Txt As String, startPos As Integer)
Call ResetRTF                   'Resetting the Format of the Text

rtfActual.SelStart = startPos   'Start position
rtfActual.SelLength = Len(Txt)  'End position
rtfActual.SelColor = vbRed      'Making red the selected text
rtfActual.SelLength = 0         'Making the selection length back to zero
End Sub

Private Sub ResetRTF()
'Reseting the previous Fore Color to Black...
rtfActual.SelStart = 0
rtfActual.SelLength = Len(rtfActual.Text)
rtfActual.SelColor = vbBlack
rtfActual.SelLength = 0
'---------------------------------------------
End Sub

Private Function GetHiglighted_Text() As String
'To find the Text which has a Fore Color as RED
'Finding it thru the RTF Tags generated in the rtfActual Rich TextBox..

Dim strTmp As String
Dim xStart As Integer, xEnd As Integer

'If its red the word & not the first word will be starting by " \cf2 " & ending with "\cf"
'in the RTF text of the Formatted Text..
If curWordCount <> 1 Then
    xStart = InStr(1, rtfActual.TextRTF, " \cf2 ") + Len(" \cf2 ")
    xEnd = InStr(xStart, rtfActual.TextRTF, "\")
Else
'Else it will be starting with "\fs20 " & ending with "\cf"
    xStart = InStr(1, rtfActual.TextRTF, "\fs20 ") + Len("\fs20 ")
    xEnd = InStr(xStart, rtfActual.TextRTF, "\")
End If

If xEnd > xStart Then
    strTmp = Mid(rtfActual.TextRTF, xStart, xEnd - xStart)
End If

'Removing the Line feed if present at the end i.e. Chr$(10)
strTmp = Replace(strTmp, vbCrLf, "")
GetHiglighted_Text = Trim(strTmp)
End Function


