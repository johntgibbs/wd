VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bimphubrpt 
   Caption         =   "Hub Inventories"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   13770
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Show Branch Totals"
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
      Left            =   10320
      TabIndex        =   8
      Top             =   240
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6975
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   12303
      _Version        =   327680
      Rows            =   1
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   11640
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label tcolor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7680
      TabIndex        =   9
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label scolor 
      BackColor       =   &H00808000&
      Caption         =   "Label2"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7680
      TabIndex        =   6
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label hubdesc 
      Caption         =   "Label2"
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
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Hub Location:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "bimphubrpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function branch_plant(pbr As String, pplant As String) As Boolean
    Dim s As String, ds As ADODB.Recordset, pflag As Boolean
    s = "select listdisplay from valuelists where listname = 'branchplants'"
    s = s & " and listreturn = '" & pbr & "'"
    s = s & " and listdisplay = '" & pplant & "'"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        pflag = True
    Else
        pflag = False
    End If
    ds.Close
    branch_plant = pflag
End Function

Private Sub sku_totals()
    Dim s As String, ds As ADODB.Recordset, i As Integer, k As Integer
    Dim tot3 As Long, tot4 As Long, tot5 As Long, udiff As Long, psize As Integer
    s = "select listdisplay from valuelists where listname = 'hubskus'"
    s = s & " and listreturn = '" & Combo1 & "' order by listdisplay"               'jv120617
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            tot3 = 0: tot4 = 0: tot5 = 0: psize = 0: udiff = 0
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = Trim(ds!listdisplay) Then
                    tot3 = tot3 + Val(Grid1.TextMatrix(i, 3))
                    tot4 = tot4 + Val(Grid1.TextMatrix(i, 4))
                    tot5 = tot5 + Val(Grid1.TextMatrix(i, 5))
                    psize = Val(Grid1.TextMatrix(i, 9))
                    udiff = (tot3 + tot4) - tot5
                End If
            Next i
            If psize <> 0 Then
                s = Trim(ds!listdisplay) & Chr(9)
                s = s & "0000" & Chr(9)
                's = s & "ZZZZ" & Chr(9)
                s = s & Chr(9)
                s = s & Format(tot3, "#") & Chr(9)
                s = s & Format(tot4, "#") & Chr(9)
                s = s & Format(tot5, "#") & Chr(9)
                s = s & Format(udiff, "#") & Chr(9)
                s = s & Format(udiff / psize, "#") & Chr(9)
                If tot5 <> 0 Then
                    s = s & Format((tot3 + tot4) / tot5, ".000") & Chr(9) & psize & Chr(9)
                    s = s & Format(psize / tot5, ".000") & Chr(9)
                    s = s & Chr(9) & hubdesc & Chr(9)
                    s = s & Format(((tot3 + tot4) / tot5) * 30, "#")
                Else
                    s = s & ".000" & Chr(9) & psize & Chr(9) & ".000" & Chr(9)
                End If
                s = s & Chr(9)
                s = s & hubdesc & Chr(9)
                Grid1.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FillStyle = flexFillRepeat
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 1
    Grid1.Sort = 5
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 1) <> "0000" Then
        'If Grid1.TextMatrix(i, 1) <> "ZZZZ" Then
            pdesc = Grid1.TextMatrix(i, 1)
            'Grid1.TextMatrix(i, 1) = " "
            Grid1.TextMatrix(i, 1) = Grid1.TextMatrix(i, 12)
            Grid1.TextMatrix(i, 12) = " "
        Else
            k = Val(Grid1.TextMatrix(i, 0))
            pdesc = skurec(k).unit & " " & skurec(k).desc
            Grid1.TextMatrix(i, 1) = pdesc
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
            If Check1.Value = 1 Then
                Grid1.CellBackColor = scolor.BackColor
                Grid1.CellForeColor = scolor.ForeColor
            Else
                Grid1.CellBackColor = tcolor.BackColor
                Grid1.CellForeColor = tcolor.ForeColor
            End If
        End If
    Next i
    If Check1.Value = 0 Then
        If Grid1.Rows > 1 Then
            For i = Grid1.Rows - 1 To 1 Step -1
                'If Grid1.TextMatrix(i, 1) = " " Then
                If Grid1.TextMatrix(i, 12) = " " Then
                    If Grid1.Rows > 2 Then
                        Grid1.RemoveItem i
                    Else
                        Grid1.Rows = 1
                    End If
                End If
            Next i
        End If
        Grid1.TextMatrix(0, 12) = "Hub Name"
    End If
    Grid1.Row = 1
End Sub

Private Sub refresh_grid()
    Dim query As String, i As Integer, ss As ADODB.Recordset
    Dim db As ADODB.Connection, ds As ADODB.Recordset, sqlx As String
    Dim tc As Integer, stot As Long
    Dim pstat As String                                                             'jv053018
    'On Error GoTo vberror
    'Label3.Caption = bimp_status_time                                                        'jv022316
    'If Label3.Caption > " " Then Label3.Caption = "Last R12 import @ " & Label3.Caption      'jv022316
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 14: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    sqlx = "select sku,lastrecpt,"
    sqlx = sqlx & "onhand,onorder,sales,"
    sqlx = sqlx & "undiff,paldiff,ohpct,roqty,"
    sqlx = sqlx & "pctgain,needqty,branchwhs"
    sqlx = sqlx & ",thiswknewpals,nextwknewpals,plantwhs"                             'jv072216
    sqlx = sqlx & " from bimp"
    sqlx = sqlx & " where branchwhs in (select listdisplay from valuelists where listname = 'hubbranches'"
    sqlx = sqlx & " and listreturn = '" & Combo1 & "')"
    sqlx = sqlx & " and sku in (select listdisplay from valuelists where listname = 'hubskus'"
    sqlx = sqlx & " and listreturn = '" & Combo1 & "')"
    sqlx = sqlx & " order by sku, branchwhs"
    'MsgBox sqlx
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!sku & Chr(9)
            i = Val(ds!sku)
            If skurec(i).sku <> ds(0) Then
                sqlx = sqlx & "..." & Chr(9)
            Else
                sqlx = sqlx & skurec(i).unit & " " & skurec(i).desc & Chr(9)
            End If
            sqlx = sqlx & DateDiff("d", ds(1), Now) & Chr(9)
            sqlx = sqlx & Format(ds(2), "#") & Chr(9)
            sqlx = sqlx & Format(ds(3), "#") & Chr(9)
            'sqlx = sqlx & Format(ds(3) + (ds(12) * ds(8)) + (ds(13) * ds(8)), "#") & Chr(9)         'jv072216
            sqlx = sqlx & Format(ds(4), "#") & Chr(9)
            sqlx = sqlx & Format(ds(5), "#") & Chr(9)
            sqlx = sqlx & Format(ds(6), "#") & Chr(9)
            sqlx = sqlx & Format(ds(7), ".000") & Chr(9)
            sqlx = sqlx & Format(ds(8), "#") & Chr(9)
            sqlx = sqlx & Format(ds(9), ".000") & Chr(9)
            sqlx = sqlx & Format(ds(10), "#") & Chr(9)
            sqlx = sqlx & ds(11)
            sqlx = sqlx & " " & branchrec(Val(ds(11))).branchname
            'If Val(age) > 0 Then
            '    If (DateDiff("d", ds(1), Now) >= Val(age) Or IsDate(ds(1)) = False) And ds(2) > 0 Then Grid1.AddItem sqlx
            'Else
            '    If ds!onhand <> 0 Or ds!onorder <> 0 Or ds!sales <> 0 Then Grid1.AddItem sqlx   'jv101316
            If branch_plant(ds!branchwhs, ds!plantwhs) = True Then
                Grid1.AddItem sqlx
            End If
            'End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
    'Grid1.FormatString = "^SKU|<Product|^Days|^OnHand|^OnOrder|^Sales|^UnitDiff|^PalletDiff|^OH%|^PalSize|^%Gain|^Need|^Branch|^Days Supply"
    Grid1.FormatString = "^SKU|<Product|^Days|^OnHand|^OnOrder|^Sales|^UnitDiff|^PalletDiff|^OH%|^PalSize|^%Gain|^Need||^Days Supply"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 700
    Grid1.ColWidth(3) = 1100
    Grid1.ColWidth(4) = 1100
    Grid1.ColWidth(5) = 1100
    Grid1.ColWidth(6) = 1100
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1100
    Grid1.ColWidth(9) = 1100
    Grid1.ColWidth(10) = 1100
    Grid1.ColWidth(11) = 1100
    Grid1.ColWidth(12) = 0 '1800
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 8)) > 0 Then                                         'jv041116
            Grid1.TextMatrix(i, 13) = Format(Val(Grid1.TextMatrix(i, 8)) * 30, "0")     'jv041116
        End If                                                                          'jv041116
        'If Val(Grid1.TextMatrix(Grid1.Row, 8)) > 0 Then                                                 'jv122115
        '    Grid1.TextMatrix(Grid1.Row, 13) = Format(Val(Grid1.TextMatrix(Grid1.Row, 8)) * 30, "0")     'jv122115
        'End If                                                                                          'jv122115
        pstat = calc_bimp_status(30, Val(Grid1.TextMatrix(i, 8)), Val(Grid1.TextMatrix(i, 7)))  'jv053018
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        If pstat = "W" Then                                                     'jv053018
        'If Val(Grid1.TextMatrix(i, 11)) > 0 Then                                'need
            Grid1.CellBackColor = whssales.wcolor.BackColor
        Else
            If pstat = "B" Then                                                 'jv053018
            'If Val(Grid1.TextMatrix(i, 7)) = 0 Then                             'pallet diff
                Grid1.CellBackColor = whssales.bcolor.BackColor
            Else
                If pstat = "G" Then                                             'jv053018
                'If Val(Grid1.TextMatrix(i, 7)) > 0 Then                         'pallet diff
                    Grid1.CellBackColor = whssales.gcolor.BackColor
                Else
                    Grid1.CellBackColor = whssales.ycolor.BackColor
                End If
            End If
        End If
        Grid1.FillStyle = flexFillRepeat
    Next i
    If Grid1.Rows > 1 Then
        pstat = calc_bimp_status(30, Val(Grid1.TextMatrix(1, 8)), Val(Grid1.TextMatrix(1, 7)))  'jv053018
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        If pstat = "W" Then                                                     'jv053018
        'If Val(Grid1.TextMatrix(1, 11)) > 0 Then
            Grid1.CellBackColor = whssales.wcolor.BackColor
        Else
            If pstat = "B" Then                                                 'jv053018
            'If Val(Grid1.TextMatrix(1, 7)) = 0 Then
                Grid1.CellBackColor = whssales.bcolor.BackColor
            Else
                If pstat = "G" Then                                             'jv053018
                'If Val(Grid1.TextMatrix(1, 7)) > 0 Then
                    Grid1.CellBackColor = whssales.gcolor.BackColor
                Else
                    Grid1.CellBackColor = whssales.ycolor.BackColor
                End If
            End If
        End If
        Grid1.FillStyle = flexFillRepeat
        Grid1.Row = 1: Grid1.Col = 3
    End If
    sku_totals
    Grid1.Redraw = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.Description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "download_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " download_click - Error Number: " & eno
        End
    End If
End Sub


Private Sub refresh_lists()
    Dim s As String, ds As ADODB.Recordset
    Combo1.Clear: List1.Clear
    s = "select * from valuelists where listname = 'hubnames' order by listreturn"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds!listreturn
            List1.AddItem ds!listdisplay
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End Sub

Private Sub Check1_Click()
    refresh_grid
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
    Dim rt As String, rf As String, rh As String
    rt = Me.Caption
    rh = "Warehouse:  " & hubdesc
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "seagreen": gndc(0) = Me.Grid1.BackColorFixed
    'htdc(1) = "cyan": gndc(1) = Me.rcolor.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid2.BackColor
    Grid1.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub Command2_Click()
    refresh_grid
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    refresh_lists
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 200
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 3.5)
End Sub

Private Sub List1_Click()
    hubdesc.Caption = List1
    refresh_grid
End Sub

