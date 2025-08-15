VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form branchpalship 
   Caption         =   "Branch Pallet Order Totals"
   ClientHeight    =   8265
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7575
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   13361
      _Version        =   327680
      ForeColor       =   128
      BackColorFixed  =   16777152
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   6000
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Plant:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "branchpalship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()
    Dim i As Integer
    Dim ds As ADODB.Recordset, sqlx As String
    Dim ss As ADODB.Recordset, s As String, oqty As Long
    Dim rs As ADODB.Recordset, ts As ADODB.Recordset, gs As ADODB.Recordset         'jv081516
    Dim c As Long
    Dim pbr As String
    Dim t2 As Integer, t3 As Integer, t4 As Integer, t5 As Integer
    t2 = 0: t3 = 0: t4 = 0: t5 = 0
    'On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 6: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    sqlx = "select plantwhs,branchwhs,sum(onorder / roqty),sum(thiswknewpals),sum(nextwknewpals)"
    sqlx = sqlx & " from bimp"
    If List1 = "All" Then
        sqlx = sqlx & " where plantwhs not in ('VENDOR', 'DRY')"
    Else
        sqlx = sqlx & " where plantwhs = '" & List1 & "'"
    End If
    sqlx = sqlx & " group by plantwhs, branchwhs"
    sqlx = sqlx & " order by branchwhs, plantwhs"
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            oqty = ds(2)
            s = "select branch, sum(netqty) from brorders where branch = " & Val(ds!branchwhs)
            If ds!plantwhs = "T10" Then s = s & " and plant = 50"
            If ds!plantwhs = "K10" Then s = s & " and plant = 51"
            If ds!plantwhs = "A10" Then s = s & " and plant = 52"
            s = s & " group by branch having sum(netqty) <> 0"
            Set ss = wdb.Execute(s)
            If ss.BOF = False Then
                oqty = oqty + ss(1)
                'MsgBox s & " = " & oqty
            End If
            ss.Close
            
            
            'Find pallet qtys in groupitems that have not been posted to trailers.      'jv081516
            s = "select id, loaded, trldate from runs where destination = '" & Val(ds!branchwhs) & "'"  'jv081916
            If ds!plantwhs = "T10" Then s = s & " and loaded = '50'"                    'jv081916
            If ds!plantwhs = "K10" Then s = s & " and loaded = '51'"                    'jv081916
            If ds!plantwhs = "A10" Then s = s & " and loaded = '52'"                    'jv081916
            Set rs = wdb.Execute(s)
            If rs.BOF = False Then
                rs.MoveFirst
                Do Until rs.EOF
                    s = "select * from trgroups where run1 = " & rs!id
                    s = s & " or run2 = " & rs!id
                    s = s & " or run3 = " & rs!id
                    s = s & " or run4 = " & rs!id
                    Set ts = wdb.Execute(s)
                    If ts.BOF = False Then
                        ts.MoveFirst
                        Do Until ts.EOF
                            s = "select * from groupitems where groupcode = '" & ts!groupcode & "'"
                            s = s & " and groupcode not in (select groupcode from trailers)"
                            Set gs = wdb.Execute(s)
                            If gs.BOF = False Then
                                gs.MoveFirst
                                Do Until gs.EOF
                                    If ts!run1 = rs!id And gs!qty1 > 0 Then oqty = oqty + gs!qty1   'jv081916
                                    If ts!run2 = rs!id And gs!qty2 > 0 Then oqty = oqty + gs!qty2   'jv081916
                                    If ts!run3 = rs!id And gs!qty3 > 0 Then oqty = oqty + gs!qty3   'jv081916
                                    If ts!run4 = rs!id And gs!qty4 > 0 Then oqty = oqty + gs!qty4   'jv081916
                                    gs.MoveNext
                                Loop
                            End If
                            gs.Close
                            ts.MoveNext
                        Loop
                    End If
                    ts.Close
                    rs.MoveNext
                Loop
            End If
            rs.Close
            '-------------------------------------------------------------------------------
            
            
            
            
            i = Val(ds(1))
            sqlx = ds(0) & Chr(9)                                                   'Plant
            sqlx = sqlx & ds(1) & "-" & branchrec(i).branchname & Chr(9)
            'sqlx = sqlx & Format(ds(2), "#") & Chr(9)
            sqlx = sqlx & Format(oqty, "#") & Chr(9)
            sqlx = sqlx & Format(ds(3), "#") & Chr(9)
            sqlx = sqlx & Format(ds(4), "#") & Chr(9)
            'sqlx = sqlx & Format(ds(3) + ds(4) + ds(2), "#")
            sqlx = sqlx & Format(ds(3) + ds(4) + oqty, "#")
            'If (ds(2) + ds(3) + ds(4)) <> 0 Then Grid1.AddItem sqlx
            'If (oqty + ds(3) + ds(4)) <> 0 Then Grid1.AddItem sqlx
            Grid1.AddItem sqlx                                                      'jv091416
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^Plant|<Branch|^Active|^This Week|^Next Week|^Total"
    
    Grid1.FillStyle = flexFillRepeat
    c = Grid1.BackColor
    pbr = " "
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 2))                                'jv091416
            Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + Val(Grid1.TextMatrix(i, 3))  'jv091416
            Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + Val(Grid1.TextMatrix(i, 4))  'jv091416
            t2 = t2 + Val(Grid1.TextMatrix(i, 2))
            t3 = t3 + Val(Grid1.TextMatrix(i, 3))
            t4 = t4 + Val(Grid1.TextMatrix(i, 4))
            t5 = t5 + Val(Grid1.TextMatrix(i, 5))
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            If pbr <> Grid1.TextMatrix(i, 1) Then
                If c = Grid1.BackColorFixed Then
                    c = Grid1.BackColor
                Else
                    c = Grid1.BackColorFixed
                End If
                pbr = Grid1.TextMatrix(i, 1)
            End If
            Grid1.CellBackColor = c
        Next i
        Grid1.Row = 1
    End If
    Grid1.AddItem List1 & Chr(9) & "Totals" & Chr(9) & t2 & Chr(9) & t3 & Chr(9) & t4 & Chr(9) & t5
    
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1400 '6 '00
    Grid1.ColWidth(3) = 1400
    Grid1.ColWidth(4) = 1400 '6 '00
    Grid1.ColWidth(5) = 1400
    Grid1.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.Description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "plantot", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " plantot - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Form_Load()
    Combo1.Clear: List1.Clear
    Combo1.AddItem "T10 - Brenham": List1.AddItem "T10"
    Combo1.AddItem "K10 - Broken Arrow": List1.AddItem "K10"
    Combo1.AddItem "A10 - Sylacauga": List1.AddItem "A10"
    Combo1.AddItem "All Plants": List1.AddItem "All"
    Me.Height = whssales.Height
    Me.Top = whssales.Top
    Me.Left = whssales.Width - Me.Width
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 4)
End Sub

Private Sub Grid1_DblClick()
    Dim i As Integer
    i = Val(Left(Grid1.TextMatrix(Grid1.Row, 1), 3))
    If Val(i) > 0 Then
        branchtrailers.bkey = i
        branchtrailers.Show
    End If
End Sub

Private Sub List1_Click()
    refresh_grid1
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    rt = Combo1 & " Branch Orders"
    rh = "Branch Orders"
    If Option1 = True Then
        rh = rh & " - Units"
    Else
        rh = rh & " - Pallets"
    End If
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    'htdc(0) = "lightcyan": gndc(0) = Me.bcolor.BackColor
    'htdc(1) = "yellow": gndc(1) = Me.ycolor.BackColor
    'htdc(2) = "white": gndc(2) = Me.wcolor.BackColor
    Grid1.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "lightyellow", "white")
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Grid1.Redraw = True: Grid1.Row = 1
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "lightyellow", "white")
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Grid1.Redraw = True: Grid1.Row = 1
        Exit Sub
    End If
End Sub

