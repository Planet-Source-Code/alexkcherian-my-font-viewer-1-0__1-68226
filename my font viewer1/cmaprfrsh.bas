Attribute VB_Name = "cmaprfresh"
'================================================================
'Font viewer 1.0
'preview installed/uninstalled .ttf fonts
'This projcet is done in visual basic 6.0.
'email: alexkcherian@rediffmail.com
'july 2004
'================================================================
Option Explicit
Public Sub cmaprfrsh()
    On Error GoTo ercmaprfrsh
    DoEvents
    mainfrm.StatusBar1.Panels(2).Text = "starting"
    'charactermap refresh on charmap form
    Dim rownum, colnum, charnum As Integer
    rownum = 0 'charmap.MSFlexGrid1.Rows
    colnum = 0 'charmap.MSFlexGrid1.Cols
    charnum = 0
    
    charmap.Label1.Caption = ""
    charmap.Label2.Caption = ""
    charmap.Label3.Caption = ""
    '++++++++++++to be reactivated
    'charmap.MSFlexGrid1.Font = charmap.Combo1.Text
    '+++++++++++++
    'charmap.MSFlexGrid1.CellAlignment = flexAlignCenterCenter '4
    
    Do While rownum <= (charmap.MSFlexGrid1.Rows) - 1
        DoEvents
        'charmap.MSFlexGrid1.RowSel = rownum
        'charmap.MSFlexGrid1.ColSel = colnum
        Do While colnum <= (charmap.MSFlexGrid1.Cols) - 1
            If charnum < 256 Then '= 256 Then
                'charmap.MSFlexGrid1.Col = 0
                'charmap.MSFlexGrid1.Row = 0
                'Exit Sub '256 Then Exit Sub
            'Else
                'charmap.MSFlexGrid1.TextMatrix(rownum, colnum) = ""
                charmap.MSFlexGrid1.TextMatrix(rownum, colnum) = Chr(charnum)
                
                mainfrm.StatusBar1.Panels(2).Text = "Refreshing Cell - " & (charnum) '"starting"
                
                charmap.MSFlexGrid1.Col = colnum
                charmap.MSFlexGrid1.Row = rownum
                charmap.MSFlexGrid1.CellFontName = charmap.Text1.Font 'charmap.Combo1.Text
                charmap.MSFlexGrid1.CellAlignment = 4
                charmap.MSFlexGrid1.CellFontSize = 20
                charnum = charnum + 1
                colnum = colnum + 1
            Else
                mainfrm.StatusBar1.Panels(2).Text = "finished"
                mainfrm.StatusBar1.Panels(2).Text = ""
                Exit Sub
            End If
        Loop
        colnum = 0
        rownum = rownum + 1
   
    Loop
    'mainfrm.StatusBar1.Panels(1).Text = "finished"
ercmaprfrsh:
    'Debug.Print charnum
    'Debug.Print colnum
    'Debug.Print rownum
    Exit Sub
End Sub
Public Sub fontmaprfrsh()
'for refreshing fontmap display according to the combo selection on charmap form
Dim curfontname As String
Dim flexrownum, flexcolnum As Integer
    If selfnttype = "inst" Then
        charmap.Text1.Font = charmap.Combo1.Text
        fontmap.Text1.Font = charmap.Combo1.Text
    
        flexcolnum = 1
        curfontname = charmap.Combo1.Text
        For flexrownum = 0 To fontmap.MSFlexGrid1.Rows - 1
            If fontmap.MSFlexGrid1.TextMatrix(flexrownum, flexcolnum) = curfontname Then
                fontmap.MSFlexGrid1.TopRow = flexrownum
                fontmap.MSFlexGrid2.TopRow = flexrownum
                Exit Sub
            End If
        Next
    End If
End Sub
