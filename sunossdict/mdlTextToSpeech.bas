Attribute VB_Name = "mdlTextToSpeech"
'untuk autoselect text
Sub SelectAllText(tb As TextBox)
   tb.SelStart = 0
   tb.SelLength = LenB(Trim$(tb.Text))
End Sub
Rem module for Text To Speech (TTS)
'modul untuk Text to Speech
Public Sub TextToSpeech(ByVal TextString As String)
    Set voice = CreateObject("SAPI.SpVoice")
    Call voice.Speak(TextString, SPF_DEFAULT)
End Sub


