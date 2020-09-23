Attribute VB_Name = "mdlAll"
Option Explicit

Rem module for System Tray (Systray)
'module for System Tray (Systray)
Public Const NIF_ICON = &H2      ' Ikon ditampilkan
Public Const NIF_MESSAGE = &H1   ' Pesan yang dikirim
Public Const NIF_TIP = &H4       ' Ketersediaan tootip
Public Const NIM_ADD = &H0       ' Masukkan ikon
Public Const NIM_DELETE = &H2    ' Hapus ikon

Public Const WM_MOUSEMOVE = &H200  ' Perpindahan kursor
Public Const WM_RBUTTONUP = &H205  ' Klik kanan mouse

' UDT notifyicon
Public Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

' handling icon in system tray
' Untuk menangani ikon di system tray
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" _
  (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public nid As NOTIFYICONDATA

'untuk undo
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const EM_UNDO = &HC7&

'theme untuk xp
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

'Note: to use these functions the projects help file must be set
'up correctly (project-properties-help file name). This must point to
'a valid HTML help file (*.chm)
'There may be some incompatibilities on systems using IE version less than 5.5


' UDT for accessing the Search tab
Private Type t_HH_Search
  lSzStruct          As Long
  lUnicodeStrings   As Long
  sSearchQuery      As String
  lProximity        As Long
  lStemmedSearch    As Long
  lTitleOnly        As Long
  lExecute          As Long
  sWindow         As String
End Type

' HTML Help Constants
Private Const HH_DISPLAY_TOPIC = &H0            ' WinHelp equivalent
Private Const HH_DISPLAY_TOC = &H1              ' WinHelp equivalent
Private Const HH_DISPLAY_INDEX = &H2            ' WinHelp equivalent
Private Const HH_DISPLAY_SEARCH = &H3           ' WinHelp equivalent
Private Const HH_KEYWORD_LOOKUP = &HD           ' WinHelp equivalent
Private Const HH_HELP_CONTEXT = &HF             ' WinHelp equivalent
Private Const HH_CLOSE_ALL = &H12               ' WinHelp equivalent

' HTML Help API declarations
Private Declare Function HTMLHelp Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hWnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByVal dwData As Long) As Long
    
Private Declare Function HTMLHelpCallSearch Lib "hhctrl.ocx" _
    Alias "HtmlHelpA" (ByVal hWnd As Long, _
    ByVal lpHelpFile As String, _
    ByVal wCommand As Long, _
    ByRef dwData As t_HH_Search) As Long
    
    
'/=============================================================================
' Name:     HTMLHelp_Contents
' Purpose:  Displays HTML Help contents
'\=============================================================================
Public Sub HTMLHELP_Contents(F As Form)

    Dim hwndHelp As Long
    
    hwndHelp = HTMLHelp(F.hWnd, App.HelpFile, HH_DISPLAY_TOC, 0)
    
End Sub



