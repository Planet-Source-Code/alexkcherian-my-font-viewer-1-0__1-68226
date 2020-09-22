VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form charmap 
   Caption         =   "Character Map"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11033
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollBars      =   2
   End
   Begin VB.Frame Frame2 
      Caption         =   "Text"
      Height          =   1815
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   375
         Left            =   4920
         TabIndex        =   5
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Copy"
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Select All"
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Character Details"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.Frame Frame4 
         Caption         =   "Keyboard key"
         Height          =   1455
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   1695
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "Shows the corresponding Keyboard Key"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Shows the Key combination needed"
            Top             =   1080
            Width           =   1455
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Other Fonts"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Installed Fonts"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Character"
         Height          =   1455
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   1695
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   39
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "The selected character of the current font."
            Top             =   240
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "charmap"
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
    If Combo1.Text <> "" Then
        MSFlexGrid1.Font = Combo1.Text
        Text1.Font = Combo1.Text
        Label2.Font = Combo1.Text
    Else
        Combo1.Text = Text1.Font
        Label2.Font = Combo1.Text
        MSFlexGrid1.Font = Combo1.Text
    End If
    DoEvents
    cmaprfrsh
    If mainfrm.sbmnuinstalledfonts.Checked = True Then
        fontmaprfrsh
    End If
End Sub

Private Sub Combo1_Click()
    MSFlexGrid1.Font = Combo1.Text
    Text1.Font = Combo1.Text
    Label2.Font = Combo1.Text
    DoEvents
    cmaprfrsh
    MSFlexGrid1.Refresh
    If mainfrm.sbmnuinstalledfonts.Checked = True Then
        fontmaprfrsh
    End If
End Sub

Private Sub Command1_Click()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    'Text1.SelText = Len(Text1.Text)  'Text1.SelLength
    'Text1.SelText = Text1.SelText '(Text1.SelLength)
    'Text1.Text = Text1.SelText
End Sub

Private Sub Command2_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
End Sub

Private Sub Command3_Click()
    Text1.Text = ""
End Sub

Private Sub Form_Load()
    Me.Visible = False
    Me.WindowState = 2
    DoEvents
    charcombload    'load font list to combobox
    findcurres      'find current resolution
    findandsetgrid  'setthe charmap grid row/col/cell
    
    '/replaced by the findandsetgrid routine
    'MSFlexGrid1.Cols = 12
    'MSFlexGrid1.Rows = 22 '20
    'MSFlexGrid1.ColWidth(0) = 935
    'MSFlexGrid1.ColWidth(1) = 935
    'MSFlexGrid1.ColWidth(2) = 935
    'MSFlexGrid1.ColWidth(3) = 935
    'MSFlexGrid1.ColWidth(4) = 935
    'MSFlexGrid1.ColWidth(5) = 935
    'MSFlexGrid1.ColWidth(6) = 935
    'MSFlexGrid1.ColWidth(7) = 935
    'MSFlexGrid1.ColWidth(8) = 935
    'MSFlexGrid1.ColWidth(9) = 935
    'MSFlexGrid1.ColWidth(10) = 935
    'MSFlexGrid1.ColWidth(11) = 935
    'MSFlexGrid1.RowHeightMin = 500
    'MSFlexGrid1.CellAlignment = 4 'flexAlignCenterCenter 'flexAlignCenterCenter
    'end of replacement/
    
    'MSFlexGrid1.ColWidth(0) = 900
    'MSFlexGrid1.ColWidth(1) = 900
    'MSFlexGrid1.ColWidth(2) = 900
    'MSFlexGrid1.ColWidth(3) = 900
    'MSFlexGrid1.ColWidth(4) = 900
    'MSFlexGrid1.ColWidth(5) = 900
    'MSFlexGrid1.ColWidth(6) = 900
    'MSFlexGrid1.ColWidth(7) = 900
    'MSFlexGrid1.ColWidth(8) = 900
    'MSFlexGrid1.ColWidth(9) = 900
    'MSFlexGrid1.ColWidth(10) = 900
    'MSFlexGrid1.ColWidth(11) = 900
    'Option1.Value = True
    If mainfrm.sbmnuotherfonts.Checked = False Then
        Option2.Enabled = False
    Else
        Option2.Enabled = True
    End If
    If selfnttype = "inst" Then
        Option1.Value = True
        Combo1.Enabled = True
        'Combo1.Text = MSFlexGrid1.Font  'this will activate the combo change event
    ElseIf selfnttype = "oth" Then
        Combo1.Enabled = False
        Option2.Value = True
    ElseIf selfnttype = "nil" Then
        Option1.Value = True
        Combo1.Enabled = True
    End If
    Me.Visible = True
    DoEvents
    'cmaprfrsh
End Sub

Private Sub Form_Resize()
    On Error GoTo charmapresizeerror
    Dim cmaphiht As Integer
    Dim cmapwid As Integer
    Dim cgridhiht As Integer
    Dim cwindstate As Integer
    
    cwindstate = charmap.WindowState
    cmaphiht = charmap.Height
    cmapwid = charmap.Width
    
    
    'If charmap.Height >= 10000 Then
        'MSFlexGrid1.Height = charmap.Height - 2700 '3000
    'Else
        'charmap.Height = 10000
        'MSFlexGrid1.Height = charmap.Height - 2700
    'End If
    If cwindstate = 0 Then
        If cmaphiht < 5000 Then '4500 Then
            cmaphiht = 5000
            charmap.Height = cmaphiht
        End If
        If cmapwid < 8565 Then   '4500 Then
            cmapwid = 8565   '5000
            charmap.Width = cmapwid
        End If
    End If
    
    'cgridhiht = MSFlexGrid1.Height
    'gridcelnum = MSFlexGrid1.Width / 935
    
    'If mainfrm.sbmnuotherfonts.Checked = True Then
        'MSFlexGrid1.Height = frmhiht - 1350
        'MSFlexGrid2.Height = frmhiht - 1350
        'Text1.Height = frmhiht - 1350
        
        'MSFlexGrid3.Top = MSFlexGrid1.Height + 150 '0
        'MSFlexGrid4.Top = MSFlexGrid3.Top
        'MSFlexGrid4.Top = MSFlexGrid3.Top
    'Else
        MSFlexGrid1.Height = cmaphiht - 2700 '50 '1000
        'MSFlexGrid2.Height = frmhiht - 750 '1000
        'Text1.Height = frmhiht - 750 '1000
    'End If
    'msflexgrid1.Height=
    
    'gridhiht = MSFlexGrid1.Height
    MSFlexGrid1.Width = cmapwid - 465 '535'1000
    Frame2.Width = cmapwid - 6105 '2
    Text1.Width = cmapwid - 7305
    Command1.Left = Text1.Left + Text1.Width + 130 '205
    Command2.Left = Command1.Left
    Command3.Left = Command1.Left
    If curres = 800 Then
        If MSFlexGrid1.Width < 11535 Then
            MSFlexGrid1.ScrollBars = flexScrollBarBoth
        Else
            MSFlexGrid1.ScrollBars = flexScrollBarVertical
        End If
    ElseIf curres = 1024 Then
        If MSFlexGrid1.Width < 14955 Then
            MSFlexGrid1.ScrollBars = flexScrollBarBoth
        Else
            MSFlexGrid1.ScrollBars = flexScrollBarVertical
        End If
    End If
    'Debug.Print MSFlexGrid1.Width
    '14955
    'MSFlexGrid2.Width = frmwid - 3150
    'Text1.Width = frmwid - 3150
    'MSFlexGrid4.Width = MSFlexGrid2.Width
    'MSFlexGrid4.Width = MSFlexGrid2.Width
charmapresizeerror:
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mainfrm.sbmnucharmap.Checked = False
    If mainfrm.sbmnuinstalledfonts.Checked = False Then
        mainfrm.sbmnuotherfonts.Checked = False
        selfnttype = "nil"
    End If
    mainfrm.mnuedit.Enabled = False
    mainfrm.mnuadvanced.Enabled = True
    mainfrm.StatusBar1.Panels(3).Text = ""
End Sub

Private Sub MSFlexGrid1_Click()
    Text1.SelText = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
End Sub

Private Sub MSFlexGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo msmoverr
    
    Dim mscurow As Integer, mscurcol As Integer
    'charmap.Label1.Caption = charmap.MSFlexGrid1.Text
    MSFlexGrid1.ToolTipText = MSFlexGrid1.Font
    mscurow = MSFlexGrid1.MouseRow
    mscurcol = MSFlexGrid1.MouseCol
    Label1.Caption = MSFlexGrid1.TextMatrix(mscurow, mscurcol)
    'Label2.Font = Combo1.Text
    'Label2.Caption = MSFlexGrid1.TextMatrix(mscurow, mscurcol)
    Label3.Caption = "Alt + 0" & (Asc(MSFlexGrid1.TextMatrix(mscurow, mscurcol)))
    mainfrm.StatusBar1.Panels(3).Text = Asc(MSFlexGrid1.TextMatrix(mscurow, mscurcol))
    If Asc(MSFlexGrid1.TextMatrix(mscurow, mscurcol)) = 38 Then
        Label2.Caption = "&" & Chr(Asc(MSFlexGrid1.TextMatrix(mscurow, mscurcol)))
        Label1.Caption = "&&"
    Else
        Label2.Caption = Chr(Asc(MSFlexGrid1.TextMatrix(mscurow, mscurcol)))
    End If
    Exit Sub
msmoverr:
    mainfrm.StatusBar1.Panels(3).Text = "Out of range"
    Exit Sub
End Sub

Private Sub Option1_Click()
    'fontmap.Visible = True
    'charmap.Visible = False
    'charmap.Option1.Value = False
    'fontmap.Option1.SetFocus
    'fontmap.Option1.Value = True
    Option1.Value = True
    'Label2.Font = othfontname
    'Text1.Font = othfontname
    'MSFlexGrid1.Font = othfontname
    Combo1.Enabled = True
    If Combo1.Text = "" Then
        If curselfont <> "" Then
            Combo1.Text = curselfont
        Else
            Combo1.Text = MSFlexGrid1.Font
        End If
        Label2.Font = Combo1.Text
        Text1.Font = Combo1.Text
    Else
        Label2.Font = Combo1.Text
        Text1.Font = Combo1.Text
        MSFlexGrid1.Font = Combo1.Text
        cmaprfrsh
    End If
    DoEvents
    'cmaprfrsh
    MSFlexGrid1.Refresh
    If mainfrm.sbmnuinstalledfonts.Checked = True Then
        fontmaprfrsh
    End If
End Sub
Sub charcombload()
    'instfontcount = Screen.FontCount
    Dim lnum1 As Integer
    'lnum1 = Screen.FontCount
    For lnum1 = 0 To instfontcount - 1 'lnum1 - 1
        'mainfrm.instfontlist.AddItem (Screen.Fonts(lnum))
        Combo1.AddItem (mainfrm.instfontlist.List(lnum1))
    Next
    Exit Sub
End Sub

Private Sub Option2_Click()
    Option2.Value = True
    Label2.Font = othfontname
    Text1.Font = othfontname
    MSFlexGrid1.Font = othfontname
    Combo1.Enabled = False
    DoEvents
    cmaprfrsh
    'MSFlexGrid1.Refresh
    'If mainfrm.sbmnuinstalledfonts.Checked = True Then
        'fontmaprfrsh
    'End If
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Text1.ToolTipText = Text1.Font
End Sub
Sub findandsetgrid()
'find the flexgridwidth and set the cells accordingly
'12x935
'800 x600 = 11535
'1024x768=14955 1e 15
'Debug.Print MSFlexGrid1.Width
    'gridcelnum = MSFlexGrid1.Width / 935
    'Debug.Print gridcelnum
    Select Case curres
        Case 800
            cellres800
        Case 1024
            cellres1024
        Case Else
            If curres < 800 Then
                cellres800
                MSFlexGrid1.ScrollBars = flexScrollBarBoth
            ElseIf curres > 1024 Then
                cellres1024
                'MSFlexGrid1.ScrollBars = flexScrollBarNone
            End If
    End Select
End Sub
Sub findcurres()
    curres = (Screen.Width / Screen.TwipsPerPixelX)
End Sub
Sub cellres800()
    MSFlexGrid1.Cols = 12
    MSFlexGrid1.Rows = 22 '20
    MSFlexGrid1.ColWidth(0) = 935
    MSFlexGrid1.ColWidth(1) = 935
    MSFlexGrid1.ColWidth(2) = 935
    MSFlexGrid1.ColWidth(3) = 935
    MSFlexGrid1.ColWidth(4) = 935
    MSFlexGrid1.ColWidth(5) = 935
    MSFlexGrid1.ColWidth(6) = 935
    MSFlexGrid1.ColWidth(7) = 935
    MSFlexGrid1.ColWidth(8) = 935
    MSFlexGrid1.ColWidth(9) = 935
    MSFlexGrid1.ColWidth(10) = 935
    MSFlexGrid1.ColWidth(11) = 935
    MSFlexGrid1.RowHeightMin = 500
    MSFlexGrid1.CellAlignment = 4 'flexAlignCenterCenter 'flexAlignCenterCenter
End Sub
Sub cellres1024()
    MSFlexGrid1.Cols = 15 '2 '16 '12
    MSFlexGrid1.Rows = 18 '22 '16 '9 '22 '20
    MSFlexGrid1.ColWidth(0) = 980 '5
    MSFlexGrid1.ColWidth(1) = 980 '5
    MSFlexGrid1.ColWidth(2) = 980 '5
    MSFlexGrid1.ColWidth(3) = 980 '5
    MSFlexGrid1.ColWidth(4) = 980 '5
    MSFlexGrid1.ColWidth(5) = 980 '5
    MSFlexGrid1.ColWidth(6) = 980 '5
    MSFlexGrid1.ColWidth(7) = 980 '5
    MSFlexGrid1.ColWidth(8) = 980 '5
    MSFlexGrid1.ColWidth(9) = 980 '5
    MSFlexGrid1.ColWidth(10) = 980 '5
    MSFlexGrid1.ColWidth(11) = 980 '5
    MSFlexGrid1.ColWidth(12) = 980 '5
    MSFlexGrid1.ColWidth(13) = 980 '5
    MSFlexGrid1.ColWidth(14) = 980 '5
    MSFlexGrid1.RowHeightMin = 600 '500
    
    MSFlexGrid1.CellAlignment = 4 'flexAlignCenterCenter 'flexAlignCenterCenter
End Sub
