Attribute VB_Name = "startup"
'================================================================
'Font viewer 1.0
'preview installed/uninstalled .ttf fonts
'This projcet is done in visual basic 6.0.
'email: alexkcherian@rediffmail.com
'july 2004
'================================================================
Option Explicit
Public instfontcount As Integer    'used to store the number of installed fonts
Public othfont As String        'used to store the other font that is not installed
Public othfontname As String    'stores the name of the otherfont
Public tmpothfont As String 'used to store othfont temporarily
Public alphastr As String 'string to store the text string to display font samples
Public curgridrow As Integer 'stores the active row of the grids
Public selfnttype As String 'stores the current selected font type ie installed or other
Public curselfont As String 'stores the currently selected installed font

Public gridcelnum As Integer    'stores the number of cells in a single row in charmap grid

Public msg As Integer   'used for mesagebox display purpose
Public myapperr As String   'used to store the error messages

Public curres As String 'used to store the current screen resolution width for charmap creation
'Public cmdlg As CommonDialog

'Public instfontlist As ListBox 'temporary list box to store all the installed fonts

'alphastr = AaBbCcDdEeFf12345

'/----------------------------------------------------
'/api and routine for checking previnstance
'Option Explicit
Public Const GW_HWNDPREV = 3


Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long


Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long


Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Sub ActivatePrevInstance()
    Dim OldTitle As String
    Dim PrevHndl As Long
    Dim result As Long
    'Save the title of the application.
    OldTitle = App.Title
    'Rename the title of this application so
    '     FindWindow
    'will not find this application instance
    '     .
    App.Title = "unwanted instance"
    'Attempt to get window handle using VB4
    '     class name.
    PrevHndl = FindWindow("ThunderRTMain", OldTitle)
    'Check for no success.


    If PrevHndl = 0 Then
        'Attempt to get window handle using VB5
        '     class name.
        PrevHndl = FindWindow("ThunderRT5Main", OldTitle)
    End If
    'Check if found


    If PrevHndl = 0 Then
        'Attempt to get window handle using VB6
        '     class name
        PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
    End If
    'Check if found


    If PrevHndl = 0 Then
        'No previous instance found.
        Exit Sub
    End If
    'Get handle to previous window.
    PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
    'Restore the program.
    result = OpenIcon(PrevHndl)
    'Activate the application.
    result = SetForegroundWindow(PrevHndl)
    'End the application.
    End
End Sub



'end of routine---------------------------------------
'-----------------------------------------------------/

Sub readfonttolist()
    instfontcount = Screen.FontCount
    'Dim lnum, lnum1, lnum2 As Integer
    Dim lnum As Integer
    'lnum1 = Screen.FontCount
    For lnum = 0 To instfontcount - 1 'lnum1 - 1
    'List1.List(lnum) = Screen.Fonts(lnum)
    'List1.AddItem (Screen.Fonts(lnum))
    'charmap.Combo1.AddItem (Screen.Fonts(lnum))
    'instfontlist.Enabled = True
    'instfontlist.AddItem "a" '(Screen.Fonts(lnum))
    mainfrm.instfontlist.AddItem (Screen.Fonts(lnum))
    Next
    Exit Sub
End Sub
