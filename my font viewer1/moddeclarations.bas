Attribute VB_Name = "moddeclarations"
'================================================================
'Font viewer 1.0
'preview installed/uninstalled .ttf fonts
'This projcet is done in visual basic 6.0.
'email: alexkcherian@rediffmail.com
'july 2004
'================================================================
Option Explicit

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOP = 0

'font add/delete api declarations
Public Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long

'test
'Public Declare Function GetFontData Lib "gdi32" Alias "GetFontDataA" (ByVal hdc As Long, ByVal dwTable As Long, ByVal dwOffset As Long, lpvBuffer As Any, ByVal cbData As Long) As Long
'Public Declare Function GetFontLanguageInfo Lib "gdi32" (ByVal hdc As Long) As Long
'Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function CreateScalableFontResource Lib "gdi32" Alias "CreateScalableFontResourceA" (ByVal fHidden As Long, ByVal lpszResourceFile As String, ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long

'api for help file integration
Public Const HELP_CONTEXT = &H1

Public Const HELP_QUIT = &H2

Public Const HELP_INDEX = &H3

Public Const HELP_CONTENTS = &H3&

Public Const HELP_HELPONHELP = &H4

Public Const HELP_SETINDEX = &H5

Public Const HELP_SETCONTENTS = &H5&

Public Const HELP_CONTEXTPOPUP = &H8&

Public Const HELP_FORCEFILE = &H9&

Public Const HELP_KEY = &H101

Public Const HELP_COMMAND = &H102&

Public Const HELP_PARTIALKEY = &H105&

Public Const HELP_MULTIKEY = &H201&

Public Const HELP_SETWINPOS = &H203&
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long

'declarations for using shell function to open website
Public Const SW_SHOWDEFAULT = 10

Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub OpenURL(strURL As String, lngHwnd As Long)
    ShellExecute lngHwnd, vbNullString, strURL, vbNullString, _
    "c:\", SW_SHOWDEFAULT
End Sub


