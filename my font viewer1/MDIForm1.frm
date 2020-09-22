VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.MDIForm mainfrm 
   BackColor       =   &H8000000C&
   Caption         =   "My Font Viewer 1.0"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picbkgrnd 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   2820
      Left            =   0
      ScaleHeight     =   2820
      ScaleWidth      =   11880
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   11880
      Begin MSComDlg.CommonDialog cmdlg 
         Left            =   1560
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ListBox instfontlist 
         Height          =   1035
         Left            =   480
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2820
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2892
            MinWidth        =   2892
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu sbmnuinstalledfonts 
         Caption         =   "&Installed Fonts"
         Shortcut        =   ^I
      End
      Begin VB.Menu sbmnuotherfonts 
         Caption         =   "&Other Fonts"
         Shortcut        =   ^N
      End
      Begin VB.Menu sbmnufsep 
         Caption         =   "-"
      End
      Begin VB.Menu sbmuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Begin VB.Menu sbmnucut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu sbmnucopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu sbmnuselectall 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu sbmnuclear 
         Caption         =   "C&lear"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuadvanced 
      Caption         =   "&Advanced"
      Begin VB.Menu sbmnufontcompare 
         Caption         =   "&Compare"
      End
      Begin VB.Menu sbmnufreesize 
         Caption         =   "Free &Size"
      End
      Begin VB.Menu sbmnuesep 
         Caption         =   "-"
      End
      Begin VB.Menu sbmnufontsize 
         Caption         =   "Font Size"
         Enabled         =   0   'False
         Begin VB.Menu sbmnufntsize10 
            Caption         =   "10"
         End
         Begin VB.Menu sbmnufntsize15 
            Caption         =   "15"
         End
         Begin VB.Menu sbmnufntsize30 
            Caption         =   "30"
            Checked         =   -1  'True
         End
         Begin VB.Menu sbmnufntsize60 
            Caption         =   "60"
         End
         Begin VB.Menu sbmnufntsize120 
            Caption         =   "120"
         End
      End
      Begin VB.Menu sbmnufontcolor 
         Caption         =   "&Font colour"
         Enabled         =   0   'False
      End
      Begin VB.Menu sbmnubkgrndcolor 
         Caption         =   "&Background Colour"
         Enabled         =   0   'False
      End
      Begin VB.Menu sbmnuadvsep1 
         Caption         =   "-"
      End
      Begin VB.Menu sbmnuautorefresh 
         Caption         =   "Auto &Refresh"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "&View"
      Begin VB.Menu sbmnufontlist 
         Caption         =   "Font &List"
      End
      Begin VB.Menu sbmnucharmap 
         Caption         =   "Character&Map"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu sbmnuhelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sbmnuonlinehelp 
         Caption         =   "&Online Help"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu sbmnuweb 
         Caption         =   "&Web Site"
      End
      Begin VB.Menu sbmnuhelpsep 
         Caption         =   "-"
      End
      Begin VB.Menu sbmnuabout 
         Caption         =   "&About Rainbow's Font Viewer 1.0"
      End
   End
   Begin VB.Menu mnupopups 
      Caption         =   "popupmenus"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnushowcharmap 
         Caption         =   "ShowCharmap"
      End
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================================
'Font viewer 1.0
'preview installed/uninstalled .ttf fonts
'This projcet is done in visual basic 6.0.
'email: alexkcherian@rediffmail.com
'july 2004
'================================================================
Option Explicit

Private Sub MDIForm_Load()
    On Error GoTo mfld
    If App.PrevInstance = True Then
        ActivatePrevInstance
        'myapperr = "another instance running"
        'App.Title = "UnWanted  Font Viewer 1.0"
        'AppActivate " Font Viewer 1.0" 'App.Title
        'myapperr = "unable to switch app"
        Unload Me
        Exit Sub
    Else
        'myapperr = "loading splah ...."
        frmSplash.Show vbModal, Me
        'Me.Visible = True
    End If
    
    'On Error GoTo mfld
    
    'mainfrm.ZOrder 1
    'frmSplash.ZOrder 0
    'ifontlist As ListBox
    ''charmap.Hide
    ''charmap.Top = 12000 'mainfrm.Height
    'charmap.Hide
    'mainfrm.Picture = LoadPicture(App.Path & "\resources\fontviewer.gif")
    'picbkgrnd.Top = 0
    'picbkgrnd.Left = 0
    'picbkgrnd.Width = mainfrm.ScaleWidth
    'picbkgrnd.Height = mainfrm.ScaleHeight
    'picbkgrnd.BackColor = vbBlue
    'picbkgrnd.Picture = LoadPicture(App.Path & "\resources\fontviewer.gif")
    'picbkgrnd.Refresh
    'frmSplash.laactivity.Caption = "Reading Installed Font Details"
    'frmSplash.Refresh
    'Load fontmap
    'frmSplash.laactivity.Caption = "Creating Character Map Display"
    'frmSplash.Refresh
    'Load charmap
    myapperr = "mdfld"
    readfonttolist
    'fontmap.ZOrder 0
    'frmSplash.laactivity.Caption = "Making Main window visible"
    'frmSplash.Refresh
    StatusBar1.Panels(1).Text = "Installed Fonts: " & (instfontcount)
    
    mainfrm.Visible = True
    'frmSplash.Hide
    DoEvents
    'frmSplash.Show vbModal, mainfrm
    'frmSplash.ZOrder 0
        ''frmSplash.laactivity.Caption = "Making Font display visible"
            ''frmSplash.Refresh
    ''fontmap.Visible = True
    
        'frmSplash.laactivity.Caption = "Unloading Splash screen"
        'frmSplash.Refresh
        'Debug.Print "unloading frmsplash"
    'Unload frmSplash
        'frmsplash.u
    selfnttype = "nil"
    
    'mainfrm.Caption = App.Title
    
    Exit Sub
mfld:
    msg = MsgBox(myapperr, vbCritical)
    myapperr = ""
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If othfont <> "" Then
        'RemoveFontResource (othfont)
        RemoveFont (othfont)
    End If
    Unload Me
End Sub

Private Sub mnushowcharmap_Click()
    On Error GoTo mnushcharmp
    If sbmnucharmap.Checked = True Then
        If selfnttype = "inst" Then
            myapperr = "mnshcmp1"
            charmap.Combo1.Text = curselfont
            
        ElseIf selfnttype = "oth" Then
            myapperr = "mnshcmp2"
            'charmap.Option2.Value = True
            charmap.Label2.Font = othfontname
            charmap.Text1.Font = othfontname
            charmap.Text1.ToolTipText = othfontname
            charmap.MSFlexGrid1.Font = othfontname
            charmap.MSFlexGrid1.ToolTipText = othfontname
            charmap.Combo1.Enabled = False
            charmap.Option2.Value = True
            DoEvents
            'cmaprfrsh
            charmap.MSFlexGrid1.Refresh
            'If mainfrm.sbmnuinstalledfonts.Checked = True Then
                'fontmaprfrsh
            'End If
        End If
        charmap.Visible = True
        'charmap.ZOrder 0
        fontmap.Visible = False
        'charmap.ZOrder 0
        mnuadvanced.Enabled = False
        mnuedit.Enabled = True
        Exit Sub
    Else
        If selfnttype = "inst" Then
            myapperr = "mnshcmp3"
            charmap.Combo1.Text = curselfont
        ElseIf selfnttype = "oth" Then
            myapperr = "mnshcmp4"
            'charmap.Option2.Value = True
            charmap.Label2.Font = othfontname
            charmap.Text1.Font = othfontname
            charmap.Text1.ToolTipText = othfontname
            charmap.MSFlexGrid1.Font = othfontname
            charmap.MSFlexGrid1.ToolTipText = othfontname
            charmap.Combo1.Enabled = False
            charmap.Option2.Value = True
            DoEvents
            'cmaprfrsh
            charmap.MSFlexGrid1.Refresh
            'If mainfrm.sbmnuinstalledfonts.Checked = True Then
                'fontmaprfrsh
            'End If
        End If
        charmap.ZOrder 0
        'charmap.Visible = True
        sbmnucharmap.Checked = True
        'charmap.Option2.Value = True
        'charmap.Label2.Font = othfontname
        'charmap.Text1.Font = othfontname
        'charmap.Text1.ToolTipText = othfontname
        'charmap.MSFlexGrid1.Font = othfontname
        'charmap.MSFlexGrid1.ToolTipText = othfontname
        'charmap.Combo1.Enabled = False
        'DoEvents
        'cmaprfrsh
        'charmap.MSFlexGrid1.Refresh
        ''If mainfrm.sbmnuinstalledfonts.Checked = True Then
            ''fontmaprfrsh
        ''End If
        If sbmnufontlist.Checked = True Then
            fontmap.Visible = False
        End If
        charmap.Visible = True
        'charmap.ZOrder 0
        
        mnuadvanced.Enabled = False
        mnuedit.Enabled = True
    End If
    Exit Sub
mnushcharmp:
    msg = MsgBox(myapperr, vbCritical)
    myapperr = ""
End Sub

Private Sub sbmnuabout_Click()
    'frmAbout.Visible = True
    frmAbout.Show vbModal, Me
End Sub

Private Sub sbmnuautorefresh_Click()
    If sbmnuautorefresh.Checked = True Then
        sbmnuautorefresh.Checked = False
    Else
        sbmnuautorefresh.Checked = True
    End If
End Sub

Private Sub sbmnubkgrndcolor_Click()
    cmdlg.CancelError = True
    On Error GoTo er1handler 'exitsub
    cmdlg.ShowColor
    'if cmdlg.
    fontmap.Text1.BackColor = cmdlg.Color
    
er1handler:
    Exit Sub
End Sub

Private Sub sbmnucharmap_Click()
    On Error GoTo smncmp
    'If charmap.Visible = True Then
        'If charmap.WindowState = 1 Then
            'charmap.WindowState = 2 '0
        'Else
            'charmap.WindowState = 2
        'End If
        'charmap.ZOrder 0
    'Else
        If mainfrm.sbmnucharmap.Checked = True Then
            If charmap.Visible = True Then
                myapperr = "smncmp1"
                If charmap.WindowState = 1 Then
                    charmap.WindowState = 2 '0
                End If
            Else
                charmap.Visible = True
            End If
            charmap.ZOrder 0
            'fontmap.Visible = False
        Else
            myapperr = "smncmp2"
            'charmap.Hide
            charmap.Visible = True
            charmap.ZOrder 0
            'If mainfrm.sbmnuinstalledfonts.Checked = True Then
                'fontmap.Visible = False
            'End If
            mainfrm.sbmnucharmap.Checked = True
            'Exit Sub
        End If
    'End If
    'sbmnufontlist.Checked = False
    'sbmnucharmap.Checked = True
    mnuadvanced.Enabled = False
    mnuedit.Enabled = True
    myapperr = "smncmp3"
    If charmap.Option2.Value = True Then
        If selfnttype = "nil" Then
            charmap.Option1.Value = True
        End If
    End If
    'sbmnufontcompare.Checked = True
    'fontmap.Visible = tr
    Exit Sub
smncmp:
    msg = MsgBox(myapperr, vbCritical)
    myapperr = ""
End Sub

Private Sub sbmnuclear_Click()
    If sbmnucharmap.Checked = True Then
        charmap.Text1.Text = ""
    End If
End Sub

Private Sub sbmnucopy_Click()
    If sbmnucharmap.Checked = True Then
        Clipboard.SetText (charmap.Text1.SelText) 'charmap.Text1.Text = ""
    End If
End Sub

Private Sub sbmnucut_Click()
    If sbmnucharmap.Checked = True Then
        Clipboard.SetText (charmap.Text1.SelText) 'charmap.Text1.Text = ""
        charmap.Text1.Text = ""
    End If
End Sub

Private Sub sbmnufntsize10_Click()
    fontmap.Text1.FontSize = 10
    sbmnufntsize10.Checked = True
    sbmnufntsize15.Checked = False
    sbmnufntsize30.Checked = False
    sbmnufntsize60.Checked = False
    sbmnufntsize120.Checked = False
End Sub

Private Sub sbmnufntsize120_Click()
    fontmap.Text1.FontSize = 120
    sbmnufntsize10.Checked = False
    sbmnufntsize15.Checked = False
    sbmnufntsize30.Checked = False
    sbmnufntsize60.Checked = False
    sbmnufntsize120.Checked = True
End Sub

Private Sub sbmnufntsize15_Click()
    fontmap.Text1.FontSize = 15
    sbmnufntsize10.Checked = False
    sbmnufntsize15.Checked = True
    sbmnufntsize30.Checked = False
    sbmnufntsize60.Checked = False
    sbmnufntsize120.Checked = False
End Sub

Private Sub sbmnufntsize30_Click()
    fontmap.Text1.FontSize = 30
    sbmnufntsize10.Checked = False
    sbmnufntsize15.Checked = False
    sbmnufntsize30.Checked = True
    sbmnufntsize60.Checked = False
    sbmnufntsize120.Checked = False
End Sub

Private Sub sbmnufntsize60_Click()
    fontmap.Text1.FontSize = 60
    sbmnufntsize10.Checked = False
    sbmnufntsize15.Checked = False
    sbmnufntsize30.Checked = False
    sbmnufntsize60.Checked = True
    sbmnufntsize120.Checked = False
End Sub

Private Sub sbmnufontcolor_Click()
    cmdlg.CancelError = True
    On Error GoTo er2handler 'exitsub
    cmdlg.ShowColor
    fontmap.Text1.ForeColor = cmdlg.Color
    
er2handler:
    Exit Sub
End Sub

Private Sub sbmnufontcompare_Click()
If mainfrm.sbmnuinstalledfonts.Checked = True Then
    If fontmap.Visible = True Then
        sbmnufontcompare.Checked = True
        sbmnufreesize.Checked = False
        fontmap.MSFlexGrid2.Visible = True
        fontmap.Text1.Visible = False
        sbmnufontsize.Enabled = False
        sbmnufontcolor.Enabled = False
        sbmnubkgrndcolor.Enabled = False
    End If
Else
    msg = MsgBox("You must open the installed fonts using File/Installed Fonts menu!", vbCritical)
End If
End Sub

Private Sub sbmnufontlist_Click()
    If mainfrm.sbmnuinstalledfonts.Checked = True Then
        If fontmap.Visible = True Then
            If fontmap.WindowState = 1 Then
                fontmap.WindowState = 2 '0
            'Else
                'charmap.WindowState = 2
            End If
            fontmap.ZOrder 0
            'sbmnufontcompare.Checked = True
        Else
        'If mainfrm.sbmnuinstalledfonts.Checked = True Then
            ''If fontmap.Visible = True Then
            'fontmap.Hide
            fontmap.Visible = True
            'fontmap.ZOrder 0
            'charmap.Visible = False
        'Else
            'Exit Sub
        End If
    Else
        msg = MsgBox("You must open the installed fonts using File/Installed Fonts menu!", vbCritical)
        Exit Sub
    End If
    sbmnufontlist.Checked = True
    ''sbmnucharmap.Checked = False
    mnuadvanced.Enabled = True
    mnuedit.Enabled = False
    ''sbmnufontcompare.Checked = True
    'charmap.Visible = False
    mainfrm.StatusBar1.Panels(3).Text = ""
End Sub

Private Sub sbmnufreesize_Click()
If mainfrm.sbmnuinstalledfonts.Checked = True Then
    If fontmap.Visible = True Then
        sbmnufontcompare.Checked = False
        sbmnufreesize.Checked = True
        fontmap.MSFlexGrid2.Visible = False
        fontmap.Text1.Visible = True
        sbmnufontsize.Enabled = True
        sbmnufontcolor.Enabled = True
        sbmnubkgrndcolor.Enabled = True
    End If
Else
    msg = MsgBox("You must open the installed fonts using File/Installed Fonts menu!", vbCritical)
End If
End Sub

Private Sub sbmnuHelp_Click()
    Dim retVal As Long
    retVal = WinHelp(mainfrm.hwnd, App.Path & "\help\ FONT VIEWER HELP.HLP", HELP_INDEX, CLng(0))
End Sub

Private Sub sbmnuinstalledfonts_Click()
    If sbmnuinstalledfonts.Checked = True Then
        fontmap.Visible = True
        sbmnufontlist.Checked = True
        'sbmnucharmap.Checked = False
        sbmnufontcompare.Checked = True
        selfnttype = "inst"
    Else
        StatusBar1.Panels(2).Text = "Loading font map"
        selfnttype = "inst"
        'fontmap.Hide
        fontmap.Visible = True
        sbmnuinstalledfonts.Checked = True
        sbmnufontlist.Checked = True
        'sbmnucharmap.Checked = False
        sbmnufontcompare.Checked = True
        StatusBar1.Panels(2).Text = ""
    End If
End Sub

Private Sub sbmnuonlinehelp_Click()
    On Error Resume Next
    'OpenURL "http://cedit.sourceforge.net/doc/index.html", Me.hwnd
    OpenURL "http://www.rainsys.com/doc/index.html", Me.hwnd
End Sub

Private Sub sbmnuotherfonts_Click()
    cmdlg.CancelError = True
    On Error GoTo er3handle
    If sbmnuotherfonts.Checked = True Then
        tmpothfont = othfont
    End If
    With cmdlg
        .Filter = "TrueTypeFonts(*.ttf)|*.ttf"
        .FilterIndex = 1
        '.DefaultExt = ".ttf"
        .ShowOpen
    End With
    othfont = cmdlg.FileName
    'cmdlg.ShowOpen
    If sbmnuotherfonts.Checked = True Then
        RemoveFont (tmpothfont)
        'fontmap.MSFlexGrid1.Height = (fontmap.Height - 1000) - fontmap.MSFlexGrid3.Height
        'AddFontResource (othfont)
        othfontname = UseFont(othfont) '(fntFileName01)
        selfnttype = "oth"
        fontmap.MSFlexGrid3.TextMatrix(0, 0) = "1"
        fontmap.MSFlexGrid3.TextMatrix(0, 1) = othfontname
        fontmap.MSFlexGrid3.ToolTipText = othfontname
        fontmap.MSFlexGrid3.Col = 1
        fontmap.MSFlexGrid3.Row = 0 'alnum
        fontmap.MSFlexGrid3.CellForeColor = vbBlue
        fontmap.MSFlexGrid3.CellAlignment = 1
        
        fontmap.MSFlexGrid4.Col = 0
        fontmap.MSFlexGrid4.Row = 0 'alnum
        fontmap.MSFlexGrid4.CellAlignment = 1
        
        fontmap.MSFlexGrid4.CellFontName = othfontname '(fontmap.MSFlexGrid3.TextMatrix(0, 1)) 'fonname
        fontmap.MSFlexGrid4.TextMatrix(0, 0) = alphastr
        fontmap.MSFlexGrid4.CellForeColor = vbBlue
        fontmap.MSFlexGrid4.ToolTipText = othfontname
        fontmap.Text1.Font = othfontname
        fontmap.Text1.ToolTipText = othfontname
        fontmap.MSFlexGrid2.Visible = False
        fontmap.Text1.Visible = True
        'sbmnucharmap.Checked = False
        sbmnufreesize.Checked = True
        sbmnufontcompare.Checked = False
        If sbmnucharmap.Checked = True Then
            If sbmnuautorefresh.Checked = True Then
                If charmap.Option2.Value = True Then
                    charmap.Label2.Font = othfontname
                    charmap.Text1.Font = othfontname
                    charmap.MSFlexGrid1.Font = othfontname
                    'charmap.Combo1.Enabled = False
                    DoEvents
                    cmaprfrsh
                End If
            End If
        End If
        Exit Sub
    Else
        sbmnuotherfonts.Checked = True
        sbmnuinstalledfonts.Checked = True
        sbmnufontlist.Checked = True
        
        'sbmnufreesize.Checked = True
        'sbmnufontcompare.Checked = False
        sbmnufontsize.Enabled = True
        sbmnufontcolor.Enabled = True
        sbmnubkgrndcolor.Enabled = True
        
        selfnttype = "oth"
        
        fontmap.Visible = True
        
        fontmap.MSFlexGrid1.Height = fontmap.Height - 1550 '1350 ') - fontmap.MSFlexGrid3.Height
        fontmap.MSFlexGrid3.Top = fontmap.MSFlexGrid1.Height + 150
        'fontmap.MSFlexGrid4.Top = fontmap.MSFlexGrid3.Top
        'fontmap.MSFlexGrid4.Width = fontmap.MSFlexGrid2.Width
        fontmap.MSFlexGrid4.Top = fontmap.MSFlexGrid3.Top
        fontmap.MSFlexGrid4.Width = fontmap.MSFlexGrid2.Width
        fontmap.MSFlexGrid2.Height = fontmap.MSFlexGrid1.Height
        fontmap.Text1.Height = fontmap.MSFlexGrid1.Height
        'MSFlexGrid1.Height = frmhiht - 1350
        'MSFlexGrid2.Height = frmhiht - 1350
        'Text1.Height = frmhiht - 1350
        
        'MSFlexGrid3.Top = MSFlexGrid1.Height + 150 '0
        'MSFlexGrid4.Top = MSFlexGrid3.Top
        
        fontmap.MSFlexGrid3.Visible = True
        'fontmap.MSFlexGrid4.Visible = True
        fontmap.MSFlexGrid4.Visible = True
        'AddFontResource (othfont)
        othfontname = UseFont(othfont) '(fntFileName01)
        fontmap.MSFlexGrid3.TextMatrix(0, 0) = "1"
        fontmap.MSFlexGrid3.TextMatrix(0, 1) = othfontname
        fontmap.MSFlexGrid3.ToolTipText = othfontname
        fontmap.MSFlexGrid3.Col = 1
        fontmap.MSFlexGrid3.Row = 0 'alnum
        fontmap.MSFlexGrid3.CellForeColor = vbBlue
        fontmap.MSFlexGrid3.CellAlignment = 1
        
        'fontmap.MSFlexGrid4.Col = 0
        'fontmap.MSFlexGrid4.Row = 0 'alnum
        fontmap.MSFlexGrid4.Col = 0
        fontmap.MSFlexGrid4.Row = 0 'alnum
        fontmap.MSFlexGrid4.CellAlignment = 1
        'fontmap.MSFlexGrid4.CellFontName = othfontname '(fontmap.MSFlexGrid3.TextMatrix(0, 1)) 'fonname
        'fontmap.MSFlexGrid4.Font = othfontname  'testing whether this will work
        'fontmap.MSFlexGrid4.TextMatrix(0, 0) = alphastr
        'fontmap.MSFlexGrid4.CellForeColor = vbBlue
        'fontmap.MSFlexGrid4.ToolTipText = othfontname
        fontmap.MSFlexGrid4.CellFontName = othfontname '(fontmap.MSFlexGrid3.TextMatrix(0, 1)) 'fonname
        fontmap.MSFlexGrid4.Font = othfontname  'testing whether this will work
        fontmap.MSFlexGrid4.TextMatrix(0, 0) = alphastr
        fontmap.MSFlexGrid4.CellForeColor = vbBlue
        fontmap.MSFlexGrid4.ToolTipText = othfontname
        fontmap.Text1.Font = othfontname
        fontmap.Text1.ToolTipText = othfontname
        fontmap.MSFlexGrid2.Visible = False
        fontmap.Text1.Visible = True
        sbmnufreesize.Checked = True
        sbmnufontcompare.Checked = False
        If sbmnucharmap.Checked = True Then
            'Option2.Enabled = False
        'Else
            charmap.Option2.Enabled = True
        End If
        Exit Sub
    End If
    Exit Sub
er3handle:
    sbmnuotherfonts.Checked = False
    fontmap.MSFlexGrid1.Height = (fontmap.Height - 750)
    fontmap.MSFlexGrid2.Height = fontmap.MSFlexGrid1.Height
    fontmap.Text1.Height = fontmap.MSFlexGrid1.Height
    fontmap.MSFlexGrid3.Visible = False
    fontmap.MSFlexGrid4.Visible = False
    If othfont <> "" Then
        'RemoveFontResource (othfont)
        RemoveFont (othfont) '(fntFileName01)
        selfnttype = "nil"
    End If
    If sbmnucharmap.Checked = True Then
            'Option2.Enabled = False
        'Else
            charmap.Option2.Enabled = False
            charmap.Option1.Enabled = True
            charmap.Combo1.Text = ""
            charmap.Text1.Font = ""
            charmap.MSFlexGrid1.Font = ""
        End If
    Exit Sub
End Sub

Private Sub sbmnuselectall_Click()
    If sbmnucharmap.Checked = True Then
        charmap.Text1.SelStart = 0
        charmap.Text1.SelLength = Len(charmap.Text1.Text)
    End If
End Sub

Private Sub sbmnuweb_Click()
    On Error Resume Next
  'OpenURL "http://cedit.sourceforge.net", Me.hwnd
  OpenURL "http://www.rainsys.com", Me.hwnd
End Sub

Private Sub sbmuexit_Click()
    Unload Me
End Sub
