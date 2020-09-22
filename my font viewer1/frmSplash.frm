VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FF0000&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5160
         Top             =   360
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font Viewer 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2280
         TabIndex        =   4
         Top             =   2160
         Width           =   3825
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "My"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   2280
         TabIndex        =   3
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label laactivity 
         BackStyle       =   0  'Transparent
         Caption         =   "Activity"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   3600
         Width           =   6735
      End
      Begin VB.Image imgLogo 
         Height          =   1650
         Left            =   720
         Picture         =   "frmSplash.frx":000C
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright: 2000-2004"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         Top             =   3270
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmSplash"
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
Dim t As Integer ', t1 As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    myapperr = "start splash"
    On Error GoTo frmsplashld
    ' See if there is already and instance.
    'If App.PrevInstance Then
        'myapperr = "another instance running"
        '/-
        ' Activate the previous instance
        'AppActivate App.Title
        
        ' Send a key (here SHIFT-key) to set the
        ' form from the previous instance to the
        ' top of the screen.
        'SendKeys "+", True
        
        ' Terminate the new instance
        '-/
        'Unload Me
    '/-
    'Else
        'frmSplash.Show vbModal ', Me
    '-/
    'End If
    '/-
    'If App.PrevInstance = True Then
        'Unload Me
        'exit sub
    'End If
    '-/
    mainfrm.Visible = True
    t = 0
    '/-
    'Me.Show vbModal'
    'SetWindowPos frmSplash.hwnd, HWND_TOPMOST, 100, 100, 500, 290, SWP_SHOWWINDOW
    '-/
    'frmSplash.Visible = True
    'DoEvents
    '/-
    'Me.Refresh
    'Me.ZOrder 0
    ''lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    ''lblProductName.Caption = App.Title
    '-/
    laactivity.Caption = "Starting to Load  Font Viewer 1.0"
    DoEvents
    '/-
    'Me.Refresh
    '-/
    myapperr = "mldstrt"
    'Load mainfrm
    '/-
    'mainfrm.Hide
        'frmSplash.Show vbModal, mainfrm
        '-/
    laactivity.Caption = "Loading  Font Viewer 1.0"
    DoEvents
    '/-
    'Me.Refresh
    'mainfrm.Visible = True
    'frmSplash.ZOrder 0
    'Unload Me
    '-/
    Exit Sub
frmsplashld:
    msg = MsgBox(myapperr, vbCritical)
    myapperr = ""
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    t = t + 1
    If t >= 10 Then Unload frmSplash
    't1
End Sub
