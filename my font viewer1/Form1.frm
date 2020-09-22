VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fontmap 
   Caption         =   "Font Map"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid4 
      Height          =   760
      Left            =   2880
      TabIndex        =   14
      Top             =   7680
      Visible         =   0   'False
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
      Height          =   760
      Left            =   120
      TabIndex        =   13
      Top             =   7680
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1349
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
   End
   Begin VB.CommandButton cmndbkgrndcolor 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   9720
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmndfontcolor 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   8400
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   11640
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   5535
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form1.frx":07CA
      Top             =   480
      Visible         =   0   'False
      Width           =   8775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7320
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      Caption         =   "FreeSize"
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Character Map"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Compare"
      Height          =   195
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7095
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   12515
      _Version        =   393216
      FixedRows       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   6255
      Left            =   2880
      TabIndex        =   7
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label labkgrndcolor 
      Caption         =   "Background Colour"
      Height          =   255
      Left            =   10080
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lacolorfont 
      Caption         =   "Font Colour"
      Height          =   255
      Left            =   8760
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "fontmap"
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
Private Sub Combo1_Change()

'MSFlexGrid2.CellFontSize = Val(Combo1.Text)
'MSFlexGrid2.CellBackColor = vbYellow

End Sub

Private Sub Combo1_Click()

Text1.FontSize = Combo1.Text
'MSFlexGrid2.RowHeightMin = Val(Val(Combo1.Text) + 300)
'MSFlexGrid2.Redraw
'MSFlexGrid2.CellFontSize = Val(Combo1.Text)

'MSFlexGrid2.CellBackColor = vbYellow
End Sub

Private Sub Form_Load()
    On Error GoTo fntmper
    myapperr = "fntmper1"
'Dim scrfonno As Integer
    alphastr = "AaBbCcDdEeFf12345"
    'fontmap.Option1.Value = True
    
    MSFlexGrid1.Top = 100
    MSFlexGrid2.Top = 100
    Text1.Top = 100
    'MSFlexGrid1.Height = 8000 '7815
    'MSFlexGrid2.Height = 7815
    'Text1.Height = 7815
    
    fontmap.MSFlexGrid1.ColWidth(0) = 500
    fontmap.MSFlexGrid1.ColWidth(1) = 2200
    fontmap.MSFlexGrid1.RowHeightMin = 760 '500
    'MSFlexGrid1.Font.Size = 30

    MSFlexGrid2.ColWidth(0) = 15000 '10000
    MSFlexGrid2.RowHeightMin = 500
    MSFlexGrid2.Font.Size = 30
    
    MSFlexGrid3.ColWidth(0) = 500
    MSFlexGrid3.ColWidth(1) = 2200
    MSFlexGrid3.RowHeightMin = 760
    'MSFlexGrid3.Font.Size = 30
    
    MSFlexGrid4.ColWidth(0) = 10000
    MSFlexGrid4.RowHeightMin = 500
    MSFlexGrid4.Font.Size = 30
    'testing
    MSFlexGrid4.ColWidth(0) = 15000 '10000
    MSFlexGrid4.RowHeightMin = 500
    MSFlexGrid4.Font.Size = 30
    
    MSFlexGrid1.Rows = instfontcount + 2 'Screen.FontCount + 2
    MSFlexGrid2.Rows = instfontcount + 2 'Screen.FontCount + 2

    'scrfonno = Screen.FontCount
    fontmap.Label1.Caption = "Screen Fonts" & "      " & "(" & instfontcount & ")"

    Combo1.Text = MSFlexGrid2.CellFontSize
    Combo1.AddItem "10"
    Combo1.AddItem "15"
    Combo1.AddItem "30"
    Combo1.AddItem "60"
    Combo1.AddItem "120"

    'listdataload
    flxg1coldata
    myapperr = "fntmper2"
    flxg1col1data
    myapperr = "fntmper3"
    
    'If Option1.Value = True Then
    'If mainfrm.sbmnufontcompare.Checked = True Then
        flxg2data
        myapperr = "fntmper4"
    'End If
    
    MSFlexGrid1.Row = 0
    Text1.FontSize = 30
    Text1.Font = MSFlexGrid1.TextMatrix(0, 1)
    Text1.Text = alphastr '"AaBbCcDdEeFf12345"
    'charmap.Combo1.Text = MSFlexGrid1.TextMatrix(0, 1)
    'charmap.Text1.Font = charmap.Combo1.Text
    ''Unload frmSplash
    Exit Sub
fntmper:
    msg = MsgBox(myapperr, vbCritical)
    myapperr = ""
    Exit Sub
End Sub
'Sub listdataload()
'Dim lnum, lnum1, lnum2 As Integer
'lnum1 = Screen.FontCount
'For lnum = 0 To lnum1 - 1

'    'List1.List(lnum) = Screen.Fonts(lnum)
'    List1.AddItem (Screen.Fonts(lnum))
'    charmap.Combo1.AddItem (Screen.Fonts(lnum))
'Next
'Exit Sub
'End Sub
Sub flxg1coldata()
    Dim num, num1, num2 As Integer
    'num1 = Screen.FontCount

    For num = 0 To instfontcount - 1 'num1 - 1
        num2 = num + 1
        MSFlexGrid1.TextMatrix(num, 0) = num2
    Next
    Exit Sub
End Sub
Sub flxg1col1data()
    Dim fnum, fnum1, fnum2 As Integer
    'fnum1 = mainfrm.instfontlist.ListCount  'List1.ListCount 'Screen.FontCount
        MSFlexGrid1.Col = 1
    For fnum = 0 To instfontcount - 1 'fnum1 - 1
        MSFlexGrid1.Row = fnum
        MSFlexGrid1.CellAlignment = 1
        MSFlexGrid1.TextMatrix(fnum, 1) = mainfrm.instfontlist.List(fnum)   'List1.List(fnum)   'Screen.Fonts(fnum)
    Next
    Exit Sub
End Sub
Sub flxg2data()
    'Dim alphastr, fonname As String
    Dim alnum, alnum1 As Integer

    'alphastr = "AaBbCcDdEeFf12345"
    'alnum = mainfrm.instfontlist.ListCount '.ListCount
    'mainfrm.StatusBar1.Panels(2).Text = (alnum)
    For alnum = 0 To instfontcount - 1  'alnum - 1
        MSFlexGrid2.Col = 0
        MSFlexGrid2.Row = alnum
        MSFlexGrid2.CellAlignment = 1
        'fonname = MSFlexGrid1.TextMatrix(alnum, 1)
        MSFlexGrid2.CellFontName = (MSFlexGrid1.TextMatrix(alnum, 1)) 'fonname
        MSFlexGrid2.TextMatrix(alnum, 0) = alphastr
    Next
    Exit Sub
End Sub

Private Sub mnuprint_Click()
'CommonDialog1.ShowPrinter

End Sub

Private Sub Form_Resize()
    On Error GoTo fontmapresizeerror
    Dim frmhiht As Integer
    Dim frmwid As Integer
    Dim gridhiht As Integer
    Dim windstate As Integer
    
    windstate = fontmap.WindowState
    frmhiht = fontmap.Height
    frmwid = fontmap.Width
    
    'mainfrm.StatusBar1.Panels(2).Text = (frmhiht)
    'mainfrm.StatusBar1.Panels(3).Text = (frmwid)
    'If windstate = 1 Then
        'If frmhiht < 4500 Then
            'frmhiht = 5000
            'fontmap.Height = frmhiht
        'End If
        'If frmwid < 8565 Then   '4500 Then
            'frmwid = 8565   '5000
            'fontmap.Width = frmwid
        'End If
    'ElseIf windstate = 0 Then
    If windstate = 0 Then
        If frmhiht < 4500 Then
            frmhiht = 5000
            fontmap.Height = frmhiht
        End If
        If frmwid < 8565 Then   '4500 Then
            frmwid = 8565   '5000
            fontmap.Width = frmwid
        End If
    End If
    
    gridhiht = MSFlexGrid1.Height
    If mainfrm.sbmnuotherfonts.Checked = True Then
        MSFlexGrid1.Height = frmhiht - 1550 '1350
        MSFlexGrid2.Height = frmhiht - 1550 '1350
        Text1.Height = frmhiht - 1350
        
        MSFlexGrid3.Top = MSFlexGrid1.Height + 150 '0
        MSFlexGrid4.Top = MSFlexGrid3.Top
        MSFlexGrid4.Top = MSFlexGrid3.Top
    Else
        MSFlexGrid1.Height = frmhiht - 750 '1000
        MSFlexGrid2.Height = frmhiht - 750 '1000
        Text1.Height = frmhiht - 750 '1000
    End If
    'msflexgrid1.Height=
    
    'gridhiht = MSFlexGrid1.Height
    'MSFlexGrid1.Width = frmwid - 1000
    MSFlexGrid2.Width = frmwid - 3150
    Text1.Width = frmwid - 3150
    MSFlexGrid4.Width = MSFlexGrid2.Width
    MSFlexGrid4.Width = MSFlexGrid2.Width
fontmapresizeerror:
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mainfrm.sbmnuinstalledfonts.Checked = False
    mainfrm.sbmnuotherfonts.Checked = False
    mainfrm.sbmnufontlist.Checked = False
    mainfrm.sbmnufreesize.Checked = False
    mainfrm.sbmnufontcompare.Checked = False
End Sub

Private Sub MSFlexGrid1_Click()
'Dim r1 As Integer
    ''MSFlexGrid2.TopRow = MSFlexGrid1.TopRow

    'r1 = MSFlexGrid1.MouseRow
    curgridrow = MSFlexGrid1.MouseRow
    MSFlexGrid1.ToolTipText = MSFlexGrid1.TextMatrix(curgridrow, 1)
    Text1.Font = MSFlexGrid1.TextMatrix(curgridrow, 1)
    curselfont = Text1.Font
    'charmap.Combo1.Text = MSFlexGrid1.TextMatrix(r1, 1)
    'charmap.Text1.Font = charmap.Combo1.Text
    
    MSFlexGrid2.Row = MSFlexGrid1.Row
    ''MSFlexGrid2.RowSel = r1
    MSFlexGrid2.CellForeColor = vbBlue
    MSFlexGrid1.CellForeColor = vbBlue
    MSFlexGrid3.CellForeColor = vbBlack
    MSFlexGrid4.CellForeColor = vbBlack
    
    selfnttype = "inst"
    'if charmap is loaded refresh it
    If mainfrm.sbmnucharmap.Checked = True Then
        If mainfrm.sbmnuautorefresh.Checked = True Then
            If charmap.Combo1.Enabled = True Then
                charmap.Combo1.Text = Text1.Font
                charmap.Label2.Font = Text1.Font 'othfontname
                charmap.Text1.Font = Text1.Font 'othfontname
                charmap.MSFlexGrid1.Font = Text1.Font 'othfontname
                'Combo1.Enabled = False
                DoEvents
                'cmaprfrsh
            End If
        End If
    End If
End Sub

Private Sub MSFlexGrid1_LeaveCell()
    MSFlexGrid1.CellForeColor = vbBlack
    MSFlexGrid2.CellForeColor = vbBlack
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        curgridrow = MSFlexGrid1.MouseRow
        MSFlexGrid1.ToolTipText = MSFlexGrid1.TextMatrix(curgridrow, 1)
        Text1.Font = MSFlexGrid1.TextMatrix(curgridrow, 1)
        curselfont = Text1.Font
        
        MSFlexGrid1.Row = curgridrow
        
        MSFlexGrid2.Row = MSFlexGrid1.Row
        ''MSFlexGrid2.RowSel = r1
        MSFlexGrid2.CellForeColor = vbBlue
        MSFlexGrid1.CellForeColor = vbBlue
        MSFlexGrid3.CellForeColor = vbBlack
        MSFlexGrid4.CellForeColor = vbBlack
    
        selfnttype = "inst"
        
        PopupMenu mainfrm.mnupopups 'showcharmap ', , , ,)
    End If
End Sub

Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim r As Integer
    r = MSFlexGrid1.MouseRow
    'urgridrow = MSFlexGrid1.MouseRow
    MSFlexGrid1.ToolTipText = MSFlexGrid1.TextMatrix(r, 1)
End Sub

Private Sub MSFlexGrid1_RowColChange()
'Dim n As Integer
'n = MSFlexGrid1.Row
'MSFlexGrid2.Row = MSFlexGrid1.Row
'MSFlexGrid2.CellForeColor = vbBlue
'MSFlexGrid1.CellForeColor = vbBlue

End Sub

Private Sub MSFlexGrid1_Scroll()
    MSFlexGrid2.TopRow = MSFlexGrid1.TopRow
End Sub

Private Sub MSFlexGrid2_Click()
'Dim r2 As Integer
'r2 = MSFlexGrid2.MouseRow
'MSFlexGrid1.TopRow = MSFlexGrid2.TopRow
'Text1.Font = MSFlexGrid1.TextMatrix(r2, 1)
'charmap.Combo1.Text = MSFlexGrid1.TextMatrix(r2, 1)
'charmap.Text1.Font = charmap.Combo1.Text

'MSFlexGrid1.Row = MSFlexGrid2.Row
'MSFlexGrid2.CellForeColor = vbBlue
'MSFlexGrid1.CellForeColor = vbBlue
    'MSFlexGrid3.CellForeColor = vbBlack
    'MSFlexGrid4.CellForeColor = vbBlack
    curgridrow = MSFlexGrid2.MouseRow
    MSFlexGrid2.ToolTipText = MSFlexGrid1.TextMatrix(curgridrow, 1)
    Text1.Font = MSFlexGrid1.TextMatrix(curgridrow, 1)
    curselfont = Text1.Font
    'charmap.Combo1.Text = MSFlexGrid1.TextMatrix(r1, 1)
    'charmap.Text1.Font = charmap.Combo1.Text
    
    MSFlexGrid1.Row = MSFlexGrid2.Row
    ''MSFlexGrid2.RowSel = r1
    MSFlexGrid2.CellForeColor = vbBlue
    MSFlexGrid1.CellForeColor = vbBlue
    MSFlexGrid3.CellForeColor = vbBlack
    MSFlexGrid4.CellForeColor = vbBlack
    
    selfnttype = "inst"
    
    If mainfrm.sbmnucharmap.Checked = True Then
        If mainfrm.sbmnuautorefresh.Checked = True Then
            If charmap.Combo1.Enabled = True Then
                charmap.Combo1.Text = Text1.Font
                charmap.Label2.Font = Text1.Font 'othfontname
                charmap.Text1.Font = Text1.Font 'othfontname
                charmap.MSFlexGrid1.Font = Text1.Font 'othfontname
                'Combo1.Enabled = False
                DoEvents
                'cmaprfrsh
            End If
        End If
    End If
End Sub

Private Sub MSFlexGrid2_LeaveCell()
'MSFlexGrid2.CellForeColor = vbBlack
    MSFlexGrid1.CellForeColor = vbBlack
    MSFlexGrid2.CellForeColor = vbBlack
End Sub

Private Sub MSFlexGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        curgridrow = MSFlexGrid2.MouseRow
        MSFlexGrid2.ToolTipText = MSFlexGrid2.TextMatrix(curgridrow, 0)
        Text1.Font = MSFlexGrid2.TextMatrix(curgridrow, 0)
        curselfont = Text1.Font
        
        MSFlexGrid2.Row = curgridrow
        MSFlexGrid1.Row = MSFlexGrid2.Row
        ''MSFlexGrid2.RowSel = r1
        MSFlexGrid2.CellForeColor = vbBlue
        MSFlexGrid1.CellForeColor = vbBlue
        MSFlexGrid3.CellForeColor = vbBlack
        MSFlexGrid4.CellForeColor = vbBlack
    
        selfnttype = "inst"
        
        PopupMenu mainfrm.mnupopups 'showcharmap ', , , ,)
    End If
End Sub

Private Sub MSFlexGrid2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim r As Integer
    r = MSFlexGrid2.MouseRow
    MSFlexGrid2.ToolTipText = MSFlexGrid1.TextMatrix(r, 1)
End Sub

Private Sub MSFlexGrid2_RowColChange()
'MSFlexGrid1.Row = MSFlexGrid2.Row

End Sub

Private Sub MSFlexGrid2_Scroll()
    MSFlexGrid1.TopRow = MSFlexGrid2.TopRow
End Sub

Private Sub MSFlexGrid3_Click()
    Text1.Font = othfontname 'MSFlexGrid3.TextMatrix(0, 1)
    'Text1.ToolTipText = othfont
    MSFlexGrid4.Row = MSFlexGrid3.Row
    ''MSFlexGrid2.RowSel = r1
    MSFlexGrid3.CellForeColor = vbBlue
    MSFlexGrid4.CellForeColor = vbBlue
    
    MSFlexGrid1.CellForeColor = vbBlack
    MSFlexGrid2.CellForeColor = vbBlack
    
    selfnttype = "oth"
    
    If mainfrm.sbmnucharmap.Checked = True Then
        If mainfrm.sbmnuautorefresh.Checked = True Then
            If charmap.Combo1.Enabled = True Then
                charmap.Combo1.Text = Text1.Font
                charmap.Label2.Font = Text1.Font 'othfontname
                charmap.Text1.Font = Text1.Font 'othfontname
                charmap.MSFlexGrid1.Font = Text1.Font 'othfontname
                'Combo1.Enabled = False
                DoEvents
                'cmaprfrsh
            End If
        End If
    End If
End Sub

Private Sub MSFlexGrid3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Text1.Font = othfontname 'MSFlexGrid3.TextMatrix(0, 1)
        'Text1.ToolTipText = othfont
        MSFlexGrid4.Row = MSFlexGrid3.Row
        ''MSFlexGrid2.RowSel = r1
        MSFlexGrid3.CellForeColor = vbBlue
        MSFlexGrid4.CellForeColor = vbBlue
    
        MSFlexGrid1.CellForeColor = vbBlack
        MSFlexGrid2.CellForeColor = vbBlack
    
        selfnttype = "oth"
        
        PopupMenu mainfrm.mnupopups 'showcharmap ', , , ,)
    End If
End Sub

Private Sub MSFlexGrid3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MSFlexGrid3.ToolTipText = "Font to Compare" & (MSFlexGrid3.TextMatrix(0, 1))
End Sub

Private Sub MSFlexGrid4_Click()
    Text1.Font = othfontname 'MSFlexGrid3.TextMatrix(0, 1)
    'Text1.ToolTipText = othfont
    'MSFlexGrid3.Row = MSFlexGrid4.Row
    ''MSFlexGrid2.RowSel = r1
    MSFlexGrid3.CellForeColor = vbBlue
    MSFlexGrid4.CellForeColor = vbBlue
    
    MSFlexGrid1.CellForeColor = vbBlack
    MSFlexGrid2.CellForeColor = vbBlack
    
    selfnttype = "oth"
    
    If mainfrm.sbmnucharmap.Checked = True Then
        If mainfrm.sbmnuautorefresh.Checked = True Then
            If charmap.Combo1.Enabled = True Then
                charmap.Combo1.Text = Text1.Font
                charmap.Label2.Font = Text1.Font 'othfontname
                charmap.Text1.Font = Text1.Font 'othfontname
                charmap.MSFlexGrid1.Font = Text1.Font 'othfontname
                'Combo1.Enabled = False
                DoEvents
                'cmaprfrsh
            End If
        End If
    End If
    'for testing only
    'mainfrm.StatusBar1.Panels(2).Text = MSFlexGrid4.MouseCol
    'mainfrm.StatusBar1.Panels(3).Text = "row" & (MSFlexGrid4.MouseRow)
End Sub

Private Sub MSFlexGrid4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Text1.Font = othfontname 'MSFlexGrid3.TextMatrix(0, 1)
        'Text1.ToolTipText = othfont
        'MSFlexGrid3.Row = MSFlexGrid4.Row
        ''MSFlexGrid2.RowSel = r1
        MSFlexGrid3.CellForeColor = vbBlue
        MSFlexGrid4.CellForeColor = vbBlue
    
        MSFlexGrid1.CellForeColor = vbBlack
        MSFlexGrid2.CellForeColor = vbBlack
    
        selfnttype = "oth"
        
        PopupMenu mainfrm.mnupopups 'showcharmap ', , , ,)
    End If
End Sub

Private Sub MSFlexGrid4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MSFlexGrid4.ToolTipText = MSFlexGrid4.CellFontName  '"Font to Compare" & (MSFlexGrid3.TextMatrix(0, 1))
End Sub

'Private Sub Option1_Click()
'Text1.Visible = False
'MSFlexGrid2.Visible = True
'Combo1.Enabled = False
'End Sub

'Private Sub Option2_Click()
    
    

    'charmap.Visible = True
    'charmap.Text1.SetFocus
    'fontmap.Option2.Value = False
    'cmaprfrsh
'End Sub

'Private Sub Option3_Click()
'MSFlexGrid2.Rows = 1
'MSFlexGrid2.RowHeightMin = 10000
'MSFlexGrid2.Visible = False
'Text1.Visible = True
'Combo1.Enabled = True
'End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        Text1.Enabled = False
        Text1.Enabled = True
        'Text1.SetFocus
        PopupMenu mainfrm.mnupopups 'showcharmap ', , , ,)
        'Text1.Enabled = True
        'Text1.Refresh
        DoEvents
        Exit Sub
    End If
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Text1.ToolTipText = Text1.FontName
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Text1.Enabled = True
    'Text1.Refresh
    'If Button = 2 Then
        'Text1.Enabled = False
        'PopupMenu mainfrm.mnupopups 'showcharmap ', , , ,)
        'Text1.Enabled = True
        'Exit Sub
    'End If
End Sub
