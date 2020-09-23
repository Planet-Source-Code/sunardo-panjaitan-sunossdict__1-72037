Attribute VB_Name = "mdlIndoTextToSpeech"
Option Explicit

'VB Interface for ITTS_DLL.dll (IndoTTS-1)
'By Arry Akhmad Arman

'Call this procedure to say something
'String can be contained one or multiple sentence
Declare Sub IndoTTS_Say Lib "ITTS_DLL.dll" Alias "IndoTTS_Say@4" (ByVal TextToSay As String)

'Call this procedure to STOP speaking
'Sistem will response after buffer empty
Declare Sub IndoTTS_Stop Lib "ITTS_DLL.dll" Alias "IndoTTS_Stop@0" ()

'Turn ON or turn OFF prosody, default value is ON
Declare Sub IndoTTS_ProsodyON Lib "ITTS_DLL.dll" Alias "IndoTTS_ProsodyON@0" ()
Declare Sub IndoTTS_ProsodyOFF Lib "ITTS_DLL.dll" Alias "IndoTTS_ProsodyOFF@0" ()

Declare Sub IndoTTS_SpeakON Lib "ITTS_DLL.dll" Alias "IndoTTS_SpeakON@0" ()
Declare Sub IndoTTS_SpeakOFF Lib "ITTS_DLL.dll" Alias "IndoTTS_SpeakOFF@0" ()

Declare Sub IndoTTS_SetPitchRatio Lib "ITTS_DLL.dll" Alias "IndoTTS_SetPitchRatio@4" (ByVal PRatio As Single)
Declare Function IndoTTS_isPlaying Lib "ITTS_DLL.dll" Alias "IndoTTS_IsPlaying@4" () As Boolean

