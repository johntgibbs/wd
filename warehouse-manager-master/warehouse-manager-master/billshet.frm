VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "Pallet Spaces"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13860
   LinkTopic       =   "Form9"
   ScaleHeight     =   6045
   ScaleWidth      =   13860
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox antecap 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8880
      TabIndex        =   12
      Text            =   "177"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox spcap 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      TabIndex        =   10
      Text            =   "289"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command2 
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
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
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8493
      _Version        =   327680
      BackColorFixed  =   12648447
      BackColorSel    =   192
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      GridLines       =   2
   End
   Begin VB.Label Label2 
      Caption         =   "Ante Cap:"
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
      Left            =   7920
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "SP Cap:"
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
      Left            =   6240
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label gcolor 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label4"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label bcolor 
      BackColor       =   &H00FFFF80&
      Caption         =   "Label3"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label ycolor 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label wcolor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function calc_date(lotcode As String) As String
    Dim seed As String
    If Left(lotcode, 2) = "00" Then
        seed = "12-31-1999"
    Else
        If Val(lotcode) > 90000 Then
            seed = "12-31-19" & Val(Left(lotcode, 2)) - 1
        Else
            seed = "12-31-20" & Format(Val(Left(lotcode, 2)) - 1, "00")
        End If
    End If
    calc_date = Format(DateAdd("d", Val(Right(lotcode, 3)), seed), "m-d-yyyy")
End Function

Private Sub refresh_skuzone(wh As Integer)
    Dim ds As ADODB.Recordset, ts As ADODB.Recordset, sqlx As String, zqty As Long, oqty As Long
    Dim f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String
    Dim f9 As String, f10 As String, f11 As String, f12 As String
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    Grid1.FormatString = "^SKU|<Product|^Zone|^InQty|^OutQty|^Total|^Surplus"
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 3200
    Grid1.ColWidth(2) = 700
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    sqlx = "select sku_config.sku,uom_type, description,zone_num"
    sqlx = sqlx & " from sku_config,zone_config"
    sqlx = sqlx & " where sku_config.sku = zone_config.sku"
    sqlx = sqlx & " order by zone_num,sku_config.sku"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            zqty = 0: oqty = 0
            sqlx = "select zone_num,sum(qty) from lane"
            sqlx = sqlx & " where sku = '" & ds(0) & "'"
            sqlx = sqlx & " and whse_num = " & wh
            sqlx = sqlx & " group by zone_num"
            Set ts = Wdb.Execute(sqlx)
            If ts.BOF = False Then
                ts.MoveFirst
                Do Until ts.EOF
                    If ts!zone_num = ds!zone_num Then
                        zqty = zqty + ts(1)
                    Else
                        oqty = oqty + ts(1)
                    End If
                    ts.MoveNext
                Loop
            End If
            ts.Close
            If zqty > 0 Or oqty > 0 Then
                sqlx = ds(0) & Chr(9)
                sqlx = sqlx & ds!uom_type & " " & ds!description & Chr(9)
                sqlx = sqlx & ds!zone_num & Chr(9)
                sqlx = sqlx & Format(zqty, "#") & Chr(9)
                sqlx = sqlx & Format(oqty, "#") & Chr(9)
                sqlx = sqlx & Format(zqty + oqty, "0")
                Grid1.AddItem sqlx
                DoEvents
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_all()
    Dim ds As ADODB.Recordset, sqlx As String
    Dim mzone As Integer, i As Integer
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.Font = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 13
    Grid1.FormatString = "^SR|^Cap|^Empty|^1Pal|^2Pal|^3Pal|^4Pal|^_|^Total|^Resv|^Orders|^Inc|^Net"
    For i = 0 To 12
        Grid1.ColWidth(i) = 1000
    Next i
    Grid1.ColWidth(0) = 1200
    mzone = 800
    i = 1
    sqlx = "select * from lane where zone_num > 0"
    sqlx = sqlx & " order by whse_num"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!whse_num <> mzone Then
                Grid1.AddItem ds!whse_num
                mzone = ds!whse_num
                i = Grid1.Rows - 1
                DoEvents
            End If
            Grid1.TextMatrix(i, 1) = Val(Grid1.TextMatrix(i, 1)) + ds!capacity
            If ds!qty = 0 Then Grid1.TextMatrix(i, 2) = Val(Grid1.TextMatrix(i, 2)) + 1
            If ds!qty = 1 Then Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + 1
            If ds!qty = 2 Then Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + 1
            If ds!qty = 3 Then Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + 1
            If ds!qty = 4 Then Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + 1
            Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 8)) + ds!qty
            If ds!resv_sku > "..." Then Grid1.TextMatrix(i, 9) = Val(Grid1.TextMatrix(i, 9)) + ds!capacity - ds!qty
            ds.MoveNext
        Loop
    End If
    ds.Close
    For i = 1 To Grid1.Rows - 1
        sqlx = "select to_whse_num,sum(order_qty - ship_plt_qty) from ship_infc"
        sqlx = sqlx & " where to_whse_num = " & Grid1.TextMatrix(i, 0)
        sqlx = sqlx & " and ship_status <> 'DONE'"
        sqlx = sqlx & " and ship_status <> 'CANC'"
        sqlx = sqlx & " group by to_whse_num"
        Set ds = Wdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            If ds(1) > 0 Then Grid1.TextMatrix(i, 10) = ds(1)
        End If
        ds.Close
    Next i
    sqlx = "select source, count(*) from paltasks where area = 'DOCK' and status = 'PEND' and userid < '0'" 'jv041316
    sqlx = sqlx & " and source = 'SR5'"                                     'jv041316
    sqlx = sqlx & " group by source"                                        'jv041316
    Set ds = Wdb.Execute(sqlx)                                         'jv041316
    If ds.BOF = False Then                                                  'jv041316
        ds.MoveFirst                                                        'jv041316
        Do Until ds.EOF                                                     'jv041316
            If ds!source = "SR1" Then Grid1.TextMatrix(1, 10) = ds(1)       'jv041316
            If ds!source = "SR2" Then Grid1.TextMatrix(2, 10) = ds(1)       'jv041316
            If ds!source = "SR3" Then Grid1.TextMatrix(3, 10) = ds(1)       'jv041316
            If ds!source = "SR5" Then Grid1.TextMatrix(4, 10) = ds(1)       'jv041316
            ds.MoveNext                                                     'jv041316
        Loop                                                                'jv041316
    End If                                                                  'jv041316
    ds.Close                                                                'jv041316
    For i = 1 To Grid1.Rows - 1
        sqlx = "select proddate,sum(sr" & Grid1.TextMatrix(i, 0) & ") from prodrcv"
        sqlx = sqlx & " where recdate1 >= '" & Format(Now, "m/d/yyyy") & "'"    'jv041416
        sqlx = sqlx & " or recdate2 >= '" & Format(Now, "m/d/yyyy") & "'"       'jv041416
        sqlx = sqlx & " or recdate3 >= '" & Format(Now, "m/d/yyyy") & "'"       'jv041416
        sqlx = sqlx & " group by proddate"
        Set ds = Wdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ds(1) > 0 Then
                    Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) + ds(1)
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
        sqlx = "select whse_num,lot_num,sum(pallets) from curr_rcpt"
        sqlx = sqlx & " where rcpt_date = '" & Format(Now, "m/d/yyyy") & "'"
        sqlx = sqlx & " and whse_num = " & Grid1.TextMatrix(i, 0)
        sqlx = sqlx & " group by whse_num,lot_num"
        Set ds = Wdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ds(2) > 0 And calc_date(ds!lot_num) <> Format(Now, "m-d-yyyy") Then
                    Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) - ds(2)
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
    Next i
    For i = 1 To Grid1.Rows - 1
        k = Val(Grid1.TextMatrix(i, 1))
        k = k - Val(Grid1.TextMatrix(i, 8))
        k = k - Val(Grid1.TextMatrix(i, 9))
        k = k + Val(Grid1.TextMatrix(i, 10))
        k = k - Val(Grid1.TextMatrix(i, 11))
        Grid1.TextMatrix(i, 12) = k
    Next i
    Grid1.AddItem ""
    Grid1.AddItem "SR Sum"
    k = Grid1.Rows - 1
    For i = 1 To Grid1.Rows - 2
        For j = 1 To 12
            Grid1.TextMatrix(k, j) = Val(Grid1.TextMatrix(k, j)) + Val(Grid1.TextMatrix(i, j))
            DoEvents
        Next j
    Next i
    Grid1.TextMatrix(k, 7) = "_"

    Grid1.AddItem "."
    sqlx = "Racks" & Chr(9) & "Cap" & Chr(9) & "Open" & Chr(9)
    sqlx = sqlx & "BB" & Chr(9) & "4Way" & Chr(9) & "Ingredient" & Chr(9) & "ReWork" & Chr(9) & "Jobbing" & Chr(9) & "Total" & Chr(9)
    sqlx = sqlx & "Resv" & Chr(9) & "Orders" & Chr(9)
    sqlx = sqlx & "Hold" & Chr(9) & "Net"
    Grid1.AddItem sqlx
    Grid1.Row = Grid1.Rows - 1
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
    Grid1.FillStyle = flexFillRepeat
    Grid1.CellBackColor = Grid1.BackColorFixed
    'Aisle 1-H
    Grid1.AddItem "1-H"
    i = Grid1.Rows - 1
    Grid1.Row = i: Grid1.Col = Grid1.Cols - 1
    sqlx = "select p.sku, r.hold, p.bbc, count(*)"          'jv033115
    sqlx = sqlx & " from racks r, rackpos p"                'jv033115
    sqlx = sqlx & " where r.id = p.rackno"                  'jv033115
    sqlx = sqlx & " and r.aisle < 'M'"                      'jv033115
    sqlx = sqlx & " group by p.sku, r.hold, p.bbc"          'jv033115
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Grid1.TextMatrix(i, 1) = Val(Grid1.TextMatrix(i, 1)) + ds(3)    'Count all including blanks for cap
            If ds!sku < "0" Then Grid1.TextMatrix(i, 2) = Val(Grid1.TextMatrix(i, 2)) + ds(3)   'Empty spaces
            If ds!sku >= "100" And ds!sku <= "9999" Then                        'jv082415
                If ds!bbc = "Y" Then
                    Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + ds(3)
                Else
                    Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + ds(3)
                End If
            Else
                If UCase(ds!sku) = "ING" Then Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(3)
                If UCase(ds!sku) = "REW" Then Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds(3)
                If UCase(ds!sku) = "JOB" Then Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
            End If
            If ds!sku > "0" Then Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 8)) + ds(3)
            If ds!hold = "Y" Or ds!hold = "1" Then
                Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) + ds(3)
            End If
            ds.MoveNext
        Loop
    End If
    DoEvents
    ds.Close
    sqlx = "select count(*) from paltasks"
    sqlx = sqlx & " where area = 'FORKLIFT'"
    sqlx = sqlx & " and description > ' '"
    sqlx = sqlx & " and target in ('STAGING', 'ORDER PICK')"
    sqlx = sqlx & " and status = 'PEND'"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If ds(0) > 0 Then Grid1.TextMatrix(i, 10) = ds(0)
    End If
    ds.Close
    
    k = Val(Grid1.TextMatrix(i, 1))
    k = k - Val(Grid1.TextMatrix(i, 8))
    k = k - Val(Grid1.TextMatrix(i, 9))
    k = k + Val(Grid1.TextMatrix(i, 10))
    Grid1.TextMatrix(i, 12) = k
    
    
    'Ante Room - Removed 5-24-16
    'Grid1.AddItem "Ante"
    'i = Grid1.Rows - 1
    'Grid1.Row = i: Grid1.Col = Grid1.Cols - 1
    'sqlx = "select sku, hold, bbc, count(*) from rackpos"
    'sqlx = sqlx & " where rackno in (select id from racks where rack = 'ANTE')"
    'sqlx = sqlx & " group by sku, hold, bbc"
    'sqlx = sqlx & ""
    'Set ds = db.OpenRecordset(sqlx)
    'If ds.BOF = False Then
    '    ds.MoveFirst
    '    Do Until ds.EOF
    '        Grid1.TextMatrix(i, 1) = antecap.Text
    '        If ds!sku >= "100" And ds!sku <= "9999" Then                        'jv082415
    '            If ds!bbc = "Y" Then
    '                Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + ds(3)
    '            Else
    '                Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + ds(3)
    '            End If
    '        Else
    '            If ds!sku > "9999" Then Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(3) 'jv082415
    '        End If
    '        If ds!sku > "0" Then Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 8)) + ds(3)
    '        If ds!hold = "Y" Then
    '            Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) + ds(3)
    '        End If
    '        'DoEvents
    '        ds.MoveNext
    '    Loop
    'End If
    'DoEvents
    'ds.Close
    'k = Val(Grid1.TextMatrix(i, 1))
    'k = k - Val(Grid1.TextMatrix(i, 8))
    'k = k - Val(Grid1.TextMatrix(i, 9))
    'k = k + Val(Grid1.TextMatrix(i, 10))
    'Grid1.TextMatrix(i, 12) = k
    'If k > 0 Then Grid1.TextMatrix(i, 2) = k
    
    'Snack Plant - Removed 5-24-16
    'Grid1.AddItem "Snack"
    'i = Grid1.Rows - 1
    'Grid1.Row = i: Grid1.Col = Grid1.Cols - 1
    'sqlx = "select sku, hold, bbc, count(*) from rackpos"
    'sqlx = sqlx & " where rackno in (select id from racks where rack = 'SP')"
    'sqlx = sqlx & " group by sku, hold, bbc"
    'sqlx = sqlx & ""
    'Set ds = db.OpenRecordset(sqlx)
    'If ds.BOF = False Then
    '    ds.MoveFirst
    '    Do Until ds.EOF
    '        Grid1.TextMatrix(i, 1) = spcap.Text
    '        If ds!sku >= "100" And ds!sku <= "9999" Then                            'jv082415
    '            If ds!bbc = "Y" Then
    '                Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + ds(3)
    '            Else
    '                Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + ds(3)
    '            End If
    '        Else
    '            If ds!sku > "9999" Then Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(3) 'Ing
    '        End If
    '        If ds!sku > "0" Then Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 8)) + ds(3)
    '        If ds!hold = "Y" Then
    '            Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) + ds(3)
    '        End If
    '        'DoEvents
    '        ds.MoveNext
    '    Loop
    'End If
    'DoEvents
    'ds.Close
    
    k = Val(Grid1.TextMatrix(i, 1))
    k = k - Val(Grid1.TextMatrix(i, 8))
    k = k - Val(Grid1.TextMatrix(i, 9))
    k = k + Val(Grid1.TextMatrix(i, 10))
    Grid1.TextMatrix(i, 12) = k
    If k > 0 Then Grid1.TextMatrix(i, 2) = k
    
    Grid1.AddItem ""
    Grid1.AddItem "Rack Sum"
    k = Grid1.Rows - 1
    For i = 1 To Grid1.Rows - 2
        s = Grid1.TextMatrix(i, 0)
        If s = "1-H" Or s = "Ante" Or s = "Snack" Then
            For j = 1 To 11
                Grid1.TextMatrix(k, j) = Format(Val(Grid1.TextMatrix(k, j)) + Val(Grid1.TextMatrix(i, j)), "#")
                DoEvents
            Next j
        End If
    Next i
    
    Grid1.AddItem ""
    Grid1.AddItem "Totals"
    k = Grid1.Rows - 1
    For i = 1 To Grid1.Rows - 2
        s = Grid1.TextMatrix(i, 0)
        If s = "SR Sum" Or s = "Rack Sum" Then
            Grid1.TextMatrix(k, 1) = Val(Grid1.TextMatrix(k, 1)) + Val(Grid1.TextMatrix(i, 1))
            'Grid1.TextMatrix(k, 2) = Val(Grid1.TextMatrix(k, 2)) + Val(Grid1.TextMatrix(i, 2))
            Grid1.TextMatrix(k, 8) = Val(Grid1.TextMatrix(k, 8)) + Val(Grid1.TextMatrix(i, 8))
            Grid1.TextMatrix(k, 9) = Val(Grid1.TextMatrix(k, 9)) + Val(Grid1.TextMatrix(i, 9))
            Grid1.TextMatrix(k, 10) = Val(Grid1.TextMatrix(k, 10)) + Val(Grid1.TextMatrix(i, 10))
            Grid1.TextMatrix(k, 12) = Val(Grid1.TextMatrix(k, 12)) + Val(Grid1.TextMatrix(i, 12))
        End If
    Next i
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub
Private Sub refresh_crane()
    Dim ds As ADODB.Recordset, sqlx As String
    Dim mzone As Integer, i As Integer
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 13
    Grid1.FormatString = "^Whs|^Zone|^Cap|^Empty|^1Pal|^2Pal|^3Pal|^4Pal|^Total|^Resv|^Orders|^Inc|^Net"
    For i = 0 To 12
        Grid1.ColWidth(i) = 1000
    Next i
    mzone = 800
    DoEvents
    i = 1
    If Text1 = 5 Then Grid1.AddItem "5" & Chr(9) & "0"                      'jv041316
    sqlx = "select * from lane where whse_num = " & Text1
    sqlx = sqlx & " and zone_num > 0"
    sqlx = sqlx & " order by zone_num"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!zone_num <> mzone Then
                Grid1.AddItem Text1 & Chr(9) & ds!zone_num
                mzone = ds!zone_num
                i = Grid1.Rows - 1
            End If
            Grid1.TextMatrix(i, 2) = Val(Grid1.TextMatrix(i, 2)) + ds!capacity
            If ds!qty = 0 Then Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + 1
            If ds!qty = 1 Then Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + 1
            If ds!qty = 2 Then Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + 1
            If ds!qty = 3 Then Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + 1
            If ds!qty = 4 Then Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + 1
            Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 8)) + ds!qty
            If ds!resv_sku > "..." Then Grid1.TextMatrix(i, 9) = Val(Grid1.TextMatrix(i, 9)) + ds!capacity - ds!qty
            ds.MoveNext
        Loop
    End If
    ds.Close
    For i = 1 To Grid1.Rows - 1
        sqlx = "select to_whse_num,sum(order_qty - ship_plt_qty) from ship_infc"
        sqlx = sqlx & " where to_whse_num = " & Text1
        sqlx = sqlx & " and ship_status <> 'DONE'"
        sqlx = sqlx & " and ship_status <> 'CANC'"
        sqlx = sqlx & " and sku in ("
        sqlx = sqlx & "select sku from zone_config where zone_num = " & Grid1.TextMatrix(i, 1) & ")"
        sqlx = sqlx & " group by to_whse_num"
        Set ds = Wdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            If ds(1) > 0 Then Grid1.TextMatrix(i, 10) = ds(1)
        End If
        ds.Close
    Next i
    If Text1 = 5 Then                                                           'jv041316
        sqlx = "select source, count(*) from paltasks where area = 'DOCK' and status = 'PEND' and userid < '0'" 'jv041316
        sqlx = sqlx & " and source = 'SR5'"                                     'jv041316
        sqlx = sqlx & " group by source"                                        'jv041316
        Set ds = Wdb.Execute(sqlx)                                         'jv041316
        If ds.BOF = False Then                                                  'jv041316
            ds.MoveFirst                                                        'jv041316
            Do Until ds.EOF                                                     'jv041316
                If ds!source = "SR5" Then Grid1.TextMatrix(1, 10) = ds(1)       'jv041316
                ds.MoveNext                                                     'jv041316
            Loop                                                                'jv041316
        End If                                                                  'jv041316
        ds.Close                                                                'jv041316
    End If                                                                      'jv041316
    
    If Text1 = 5 Then                                                           'jv041316
        sqlx = "select proddate, sum(sr5) from prodrcv"                         'jv041316
        sqlx = sqlx & " where recdate1 >= '" & Format(Now, "m/d/yyyy") & "'"    'jv041416
        sqlx = sqlx & " or recdate2 >= '" & Format(Now, "m/d/yyyy") & "'"       'jv041416
        sqlx = sqlx & " or recdate3 >= '" & Format(Now, "m/d/yyyy") & "'"       'jv041416
        sqlx = sqlx & " group by proddate"                                      'jv041316
        Set ds = Wdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ds(1) > 0 Then
                    Grid1.TextMatrix(1, 11) = Val(Grid1.TextMatrix(1, 11)) + ds(1)  'jv041316
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
    Else                                                                        'jv041316
        For i = 1 To Grid1.Rows - 1
            sqlx = "select sr" & Text1 & " from prodrcv"
            sqlx = sqlx & " where (recdate1 >= '" & Format(Now, "m/d/yyyy") & "'"      'jv041416
            sqlx = sqlx & " or recdate2 >= '" & Format(Now, "m/d/yyyy") & "'"       'jv041416
            sqlx = sqlx & " or recdate3 >= '" & Format(Now, "m/d/yyyy") & "')"       'jv041416
            sqlx = sqlx & " and sku in (select sku from zone_config"
            sqlx = sqlx & " where zone_num = " & Grid1.TextMatrix(i, 1) & ")"
            Set ds = Wdb.Execute(sqlx)
            If ds.BOF = False Then
                ds.MoveFirst
                Do Until ds.EOF
                    If ds(0) > 0 Then
                        Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) + ds(0)
                    End If
                    ds.MoveNext
                Loop
            End If
            ds.Close
        Next i
    End If
    For i = 1 To Grid1.Rows - 1
        k = Val(Grid1.TextMatrix(i, 2))
        k = k - Val(Grid1.TextMatrix(i, 8))
        k = k - Val(Grid1.TextMatrix(i, 9))
        k = k + Val(Grid1.TextMatrix(i, 10))
        k = k - Val(Grid1.TextMatrix(i, 11))
        Grid1.TextMatrix(i, 12) = k
    Next i
    If Grid1.Rows = 1 Then Exit Sub
    Grid1.AddItem Text1 & Chr(9) & "Total"
    k = Grid1.Rows - 1
    For i = 1 To Grid1.Rows - 2
        For j = 2 To 12
            Grid1.TextMatrix(k, j) = Val(Grid1.TextMatrix(k, j)) + Val(Grid1.TextMatrix(i, j))
        Next j
    Next i
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub
Private Sub Refresh_racks()
    Dim ds As ADODB.Recordset, sqlx As String
    Dim mzone As String, i As Integer
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 13
    Grid1.FormatString = "^Aisle|^Cap|^Open|^BB|^4Way|^Ingredient|^ReWork|^Jobbing|^Total|^Resv|^Orders|^Hold|^Net"
    For i = 0 To 12
        Grid1.ColWidth(i) = 1000
    Next i
    Grid1.ColWidth(0) = 1100
    mzone = "800"
    i = 1
    sqlx = "select r.aisle, r.rack, p.sku, r.resv_sku, r.hold, p.bbc, count(*)"
    sqlx = sqlx & " from racks r, rackpos p"
    sqlx = sqlx & " where r.id = p.rackno"
    sqlx = sqlx & " and r.aisle < 'S' and r.rack <> 'OP'"
    sqlx = sqlx & " group by r.aisle, r.rack, p.sku, r.resv_sku, r.hold, p.bbc"
    sqlx = sqlx & " order by r.aisle, r.rack, p.sku, r.resv_sku, r.hold, p.bbc"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!aisle <> mzone Then
                Grid1.AddItem ds!aisle
                mzone = ds!aisle
                i = Grid1.Rows - 1
            End If
            If ds!aisle = "M" Then
                Grid1.TextMatrix(i, 1) = Val(spcap) + Val(antecap)
            Else
                Grid1.TextMatrix(i, 1) = Val(Grid1.TextMatrix(i, 1)) + ds(6)
                If ds!sku < "0" Then Grid1.TextMatrix(i, 2) = Val(Grid1.TextMatrix(i, 2)) + ds(6)
            End If
            If ds!sku >= "100" And ds!sku <= "9999" Then                            'jv082415
                If ds!bbc = "Y" Then
                    Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + ds(6)
                Else
                    Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + ds(6)
                End If
            Else
                If UCase(ds!sku) = "ING" Then Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(6)
                If UCase(ds!sku) = "REW" Then Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds(6)
                If UCase(ds!sku) = "JOB" Then Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(6)
            End If
            If ds!sku > "0" Then Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 8)) + ds(6)
            If ds!resv_sku > "." Then
                If ds!sku < "0" Then Grid1.TextMatrix(i, 9) = Val(Grid1.TextMatrix(i, 9)) + ds(6)
            End If
            If ds!hold <> 0 Then
                Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) + ds(6)
            End If
            DoEvents
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.AddItem "Orders"
    k = Grid1.Rows - 1
    sqlx = "select count(*) from paltasks"
    sqlx = sqlx & " where area = 'FORKLIFT'"
    sqlx = sqlx & " and description > ' '"
    sqlx = sqlx & " and target in ('STAGING', 'ORDER PICK')"
    sqlx = sqlx & " and status = 'PEND'"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If ds(0) > 0 Then Grid1.TextMatrix(k, 10) = ds(0)
    End If
    ds.Close
    For i = 1 To Grid1.Rows - 1
        k = Val(Grid1.TextMatrix(i, 1))
        k = k - Val(Grid1.TextMatrix(i, 8))
        k = k - Val(Grid1.TextMatrix(i, 9))
        k = k + Val(Grid1.TextMatrix(i, 10))
        Grid1.TextMatrix(i, 12) = k
        If Grid1.TextMatrix(i, 0) = "M" And Val(Grid1.TextMatrix(i, 12)) > 0 Then
            Grid1.TextMatrix(i, 2) = Grid1.TextMatrix(i, 12)
        End If
    Next i
    If Grid1.Rows = 1 Then Exit Sub
    Grid1.AddItem "Total"
    k = Grid1.Rows - 1
    For i = 1 To Grid1.Rows - 2
        For j = 1 To Grid1.Cols - 2 '9              'jv033115
            Grid1.TextMatrix(k, j) = Val(Grid1.TextMatrix(k, j)) + Val(Grid1.TextMatrix(i, j))
            DoEvents
        Next j
    Next i
    Grid1.Redraw = True
End Sub

Private Sub Combo1_Click()
    Text1 = Combo1.ListIndex + 1
    Command2.Visible = False
End Sub

Private Sub Command1_Click()
    If Val(Text1) > 0 And Val(Text1) < 4 Then refresh_crane
    If Val(Text1) = 4 Then Refresh_racks
    If Val(Text1) = 5 Then refresh_crane
    If Val(Text1) = 6 Then refresh_all
    If Val(Text1) = 7 Then refresh_skuzone (1)
    If Val(Text1) = 8 Then refresh_skuzone (2)
    If Val(Text1) = 9 Then refresh_skuzone (3)
    Command2.Visible = True
End Sub

Private Sub Command2_Click()
    Dim rt As String, rh As String, rf As String, i
    rt = "Pallet Spaces"
    rh = Combo1
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Combo1.AddItem "SR-1"
    Combo1.AddItem "SR-2"
    Combo1.AddItem "SR-3"
    Combo1.AddItem "Racks"
    Combo1.AddItem "SR-5"
    Combo1.AddItem "All Whs"
    Combo1.AddItem "SKU Zones - SR1"
    Combo1.AddItem "SKU Zones - SR2"
    Combo1.AddItem "SKU Zones - SR3"
    Combo1.ListIndex = 5
    Command1_Click
    DoEvents
    Grid1.Row = 5: Grid1.Col = 8
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Form9.Width - 80
    If Form9.Height > 2000 Then Grid1.Height = Form9.Height - 1005
End Sub
