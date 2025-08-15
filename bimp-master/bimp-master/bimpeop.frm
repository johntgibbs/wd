VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bimpeop 
   Caption         =   "E-O-P Units"
   ClientHeight    =   10920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Include Pallet Orders"
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
      Left            =   3960
      TabIndex        =   9
      Top             =   120
      Width           =   2655
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   4080
      TabIndex        =   7
      Top             =   7560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   1680
      TabIndex        =   6
      Top             =   7560
      Visible         =   0   'False
      Width           =   2055
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
      Height          =   495
      Left            =   10320
      TabIndex        =   5
      Top             =   0
      Width           =   1815
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
      Height          =   495
      Left            =   7800
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.TextBox sdate 
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
      Left            =   12360
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox edate 
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
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   16113
      _Version        =   327680
      BackColorFixed  =   16777152
      FocusRect       =   0
   End
   Begin VB.Label wcolor 
      Caption         =   "wcolor"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9120
      TabIndex        =   8
      Top             =   10200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "E-O-P Date:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "bimpeop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function branch_plant(pwhs As String) As String
    Dim i As Integer, s As String
    s = " "
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = pwhs Then
            s = List2.List(i)
            Exit For
        End If
    Next i
    branch_plant = s
    'MsgBox pwhs & " = " & s
End Function

Private Sub process_orders()                                    'jv101116
    Dim ds As ADODB.Recordset, s As String, punits As Long
    Dim i As Integer, psku As String, oqty As Long
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.Cols = 13
    
    
    s = "select plantwhs,branchwhs,sku,onorder,roqty from bimp"                         'jv062817
    s = s & " where plantwhs not in ('VENDOR', 'DRY')"
    s = s & " and onorder > 0"
    s = s & " order by plantwhs,branchwhs,sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            oqty = ds(3)
            'If ds!sku = "319" Then MsgBox "bimp onorder " & ds!branchwhs & " = " & oqty
            s = "select branch, sum(netqty) from brorders where branch = " & Val(ds!branchwhs)
            s = s & " and sku = '" & ds!sku & "'"
            If ds!plantwhs = "T10" Then s = s & " and plant = 50"
            If ds!plantwhs = "K10" Then s = s & " and plant = 51"
            If ds!plantwhs = "A10" Then s = s & " and plant = 52"
            s = s & " group by branch having sum(netqty) <> 0"
            Set ss = wdb.Execute(s)
            If ss.BOF = False Then
                oqty = oqty + (ss(1) * ds!roqty)
                'If ds!sku = "319" Then MsgBox s & " = " & oqty
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
                            s = s & " and sku = '" & ds!sku & "'"
                            Set gs = wdb.Execute(s)
                            If gs.BOF = False Then
                                gs.MoveFirst
                                Do Until gs.EOF
                                    If ts!run1 = rs!id And gs!qty1 > 0 Then oqty = oqty + (gs!qty1 * ds!roqty)   'jv081916
                                    If ts!run2 = rs!id And gs!qty2 > 0 Then oqty = oqty + (gs!qty2 * ds!roqty)   'jv081916
                                    If ts!run3 = rs!id And gs!qty3 > 0 Then oqty = oqty + (gs!qty3 * ds!roqty)   'jv081916
                                    If ts!run4 = rs!id And gs!qty4 > 0 Then oqty = oqty + (gs!qty4 * ds!roqty)   'jv081916
                                    'If ds!sku = "319" Then MsgBox s & " = " & oqty
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
            
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = ds!plantwhs And Grid1.TextMatrix(i, 2) = ds!sku Then
                    Grid1.TextMatrix(i, 9) = Val(Grid1.TextMatrix(i, 9)) + oqty
                    Exit For
                End If
            Next i
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = "ALL" And Grid1.TextMatrix(i, 2) = ds!sku Then
                    Grid1.TextMatrix(i, 9) = Val(Grid1.TextMatrix(i, 9)) + oqty
                    Exit For
                End If
            Next i
            
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    
    
    
    
    s = "select plantwhs, sku, thiswknewpals, nextwknewpals from bimp"
    s = s & " where thiswknewpals > 0 or nextwknewpals > 0"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!thiswknewpals > 0 Then
                punits = ds!thiswknewpals * skurec(Val(ds!sku)).pallet
                For i = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 0) = ds!plantwhs And Grid1.TextMatrix(i, 2) = ds!sku Then
                        Grid1.TextMatrix(i, 10) = Val(Grid1.TextMatrix(i, 10)) + punits
                        Exit For
                    End If
                Next i
                For i = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 0) = "ALL" And Grid1.TextMatrix(i, 2) = ds!sku Then
                        Grid1.TextMatrix(i, 10) = Val(Grid1.TextMatrix(i, 10)) + punits
                        Exit For
                    End If
                Next i
            End If
            If ds!nextwknewpals > 0 Then
                punits = ds!nextwknewpals * skurec(Val(ds!sku)).pallet
                For i = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 0) = ds!plantwhs And Grid1.TextMatrix(i, 2) = ds!sku Then
                        Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) + punits
                        Exit For
                    End If
                Next i
                For i = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 0) = "ALL" And Grid1.TextMatrix(i, 2) = ds!sku Then
                        Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) + punits
                        Exit For
                    End If
                Next i
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FillStyle = flexFillRepeat
    psku = Grid1.TextMatrix(1, 2)
    For i = 1 To Grid1.Rows - 1
        punits = Val(Grid1.TextMatrix(i, 7))
        punits = punits - Val(Grid1.TextMatrix(i, 9))
        punits = punits - Val(Grid1.TextMatrix(i, 10))
        punits = punits - Val(Grid1.TextMatrix(i, 11))
        Grid1.TextMatrix(i, 12) = Format(punits, "#")
        If punits < 0 Then
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellForeColor = wcolor.ForeColor
        End If
        If psku <> Grid1.TextMatrix(i, 2) Then
            hrow = Not hrow
            psku = Grid1.TextMatrix(i, 2)
        End If
        If hrow = True Then
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellBackColor = Grid1.BackColorFixed
        End If
    Next i
    Grid1.Row = 1
    s = "^Whs|<Location|^SKU|^Unit|<Flavor|^Plant|^Branches|^Total|<|^Active|^ThisWeek|^NextWeek|^Net"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 900
    Grid1.ColWidth(1) = 1600
    Grid1.ColWidth(2) = 900
    Grid1.ColWidth(3) = 900
    Grid1.ColWidth(4) = 3000
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1500
    Grid1.ColWidth(7) = 1500
    Grid1.ColWidth(8) = 0 '1500
    Grid1.ColWidth(9) = 1500
    Grid1.ColWidth(10) = 1500
    Grid1.ColWidth(11) = 1500
    Grid1.ColWidth(12) = 1500
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_bimp_history()
    Dim tdate As String, ds As ADODB.Recordset, t1 As String, t2 As String
    Dim t3 As Long, t4 As Long, t5 As Long, t6 As Long, t7 As Long, t8 As Long
    Dim brp As String, rc As Long, cldate As String
    rc = 0
    tdate = Format(DateAdd("d", 1, edate), "mm-dd-yyyy")
    'date1 = "07-04-2015"
    date1 = Mid(tdate, 1, 3) & "01" & Mid(tdate, 6, 5)
    cldate = Format(DateAdd("d", -1, date1), "mm-dd-yyyy")
    'MsgBox tdate & " " & date1 & " " & cldate
    Screen.MousePointer = 11
    
    q = "select g.inventory_item_id, m.segment1," & _
        " g.subinventory_code," & _
        " g.primary_quantity" & _
        " from gmf_period_balances g, mtl_system_items_b m" & _
        " where g.subinventory_code >= '001' and g.subinventory_code <= '090'" & _
        " and g.subinventory_code not in ('012', '015', '016')" & _
        " and m.organization_id = g.organization_id" & _
        " and m.inventory_item_id = g.inventory_item_id" & _
        " and m.segment1 >= '100' and m.segment1 <= '9999'" & _
        " and g.acct_period_id in" & _
        " (select acct_period_id from org_acct_periods" & _
        " Where organization_id = g.organization_id" & _
        " and schedule_close_date = TO_DATE('" & Format(cldate, "dd-mmm-yyyy") & "'))" & _
        " order by m.segment1"
    'MsgBox q
    
    t1 = Format(Now, "hh:mm:ss")
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            brp = branch_plant(ds(2))
            If brp > "A" Then
            '    MsgBox ds(2) & " " & ds(1)
            'Else
                For i = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = brp Then
                        Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds(3)
                        Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                        'If i = 1 Then DoEvents
                    End If
                    If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = "ALL" Then
                        Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds(3)
                        Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                        'If i = 16 Then DoEvents
                        'if ds(1) = "777" Then DoEvents
                        Exit For
                    End If
                Next i
            End If
            
            rc = rc + 1
            If Right(Format(rc, "0000000"), 4) = "0200" Then DoEvents
            ds.MoveNext
        Loop
    End If
    ds.Close
    DoEvents
    'MsgBox "check"
    
    'Exit Sub
    q = "select g.inventory_item_id, m.segment1," & _
        " g.subinventory_code," & _
        " g.primary_quantity" & _
        " from gmf_period_balances g, mtl_system_items_b m" & _
        " where g.subinventory_code in ('A10', 'K10', 'T10')" & _
        " and m.organization_id = g.organization_id" & _
        " and m.inventory_item_id = g.inventory_item_id" & _
        " and m.segment1 >= '100' and m.segment1 <= '9999'" & _
        " and g.acct_period_id in" & _
        " (select acct_period_id from org_acct_periods" & _
        " Where organization_id = g.organization_id" & _
        " and schedule_close_date = TO_DATE('" & Format(cldate, "dd-mmm-yyyy") & "'))" & _
        " order by m.segment1"
    'MsgBox q
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = ds(2) Then
                    Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(3)
                    Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                End If
                If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = "ALL" Then
                    Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(3)
                    Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                    Exit For
                    If ds(1) = "777" Then DoEvents
                End If
            Next i
            rc = rc + 1
            ds.MoveNext
        Loop
    End If
    ds.Close
    DoEvents
    'MsgBox "check plant"
        
    q = "select o.inventory_item_id, m.segment1," & _
        " o.subinventory_code," & _
        " o.transaction_quantity" & _
        " from mtl_material_transactions o, mtl_system_items_b m" & _
        " where o.subinventory_code >= '001' and o.subinventory_code <= '090'" & _
        " and o.subinventory_code not in ('012', '015', '016')" & _
        " and m.organization_id = o.organization_id" & _
        " and m.inventory_item_id = o.inventory_item_id" & _
        " and m.segment1 >= '100' and m.segment1 <= '9999'" & _
        " and o.transaction_date < TO_DATE('" & Format(tdate, "dd-mmm-yyyy") & "')" & _
        " and o.transaction_date >= TO_DATE('" & Format(date1, "dd-mmm-yyyy") & "')"
        
    'MsgBox q
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            brp = branch_plant(ds(2))
            If brp > "A" Then
            '    MsgBox ds(2) & " " & ds(1)
            'Else
                For i = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = brp Then
                        Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds(3)
                        Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                        'If i = 1 Then DoEvents
                    End If
                    If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = "ALL" Then
                        Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds(3)
                        Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                        Exit For
                    End If
                Next i
            End If
            rc = rc + 1
            If Right(Format(rc, "0000000"), 4) = "0200" Then DoEvents
            ds.MoveNext
        Loop
    End If
    ds.Close
    DoEvents
    
    q = "select o.inventory_item_id, m.segment1," & _
        " o.subinventory_code," & _
        " o.transaction_quantity" & _
        " from mtl_material_transactions o, mtl_system_items_b m" & _
        " where o.subinventory_code in ('A10', 'K10', 'T10')" & _
        " and m.organization_id = o.organization_id" & _
        " and m.inventory_item_id = o.inventory_item_id" & _
        " and m.segment1 >= '100' and m.segment1 <= '9999'" & _
        " and o.transaction_date < TO_DATE('" & Format(tdate, "dd-mmm-yyyy") & "')" & _
        " and o.transaction_date >= TO_DATE('" & Format(date1, "dd-mmm-yyyy") & "')"
        
    'MsgBox q
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = ds(2) Then
                    Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(3)
                    Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                End If
                If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = "ALL" Then
                    Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(3)
                    Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                    Exit For
                    If ds(1) = "777" Then DoEvents
                End If
            Next i
            rc = rc + 1
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t8 = 0
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = "ALL" Then
            t5 = t5 + Val(Grid1.TextMatrix(i, 5))
            t6 = t6 + Val(Grid1.TextMatrix(i, 6))
            t7 = t7 + Val(Grid1.TextMatrix(i, 7))
        End If
    Next i
    Grid1.AddItem "..."
    s = "..." & Chr(9) & "..." & Chr(9) & "..." & Chr(9) & "..." & Chr(9) & "Totals" & Chr(9)
    s = s & t5 & Chr(9)
    s = s & t6 & Chr(9)
    s = s & t7 & Chr(9)
    Grid1.AddItem s
    Screen.MousePointer = 0
    MsgBox t1 & " - " & Format(Now, "hh:mm:ss") & " Records: " & rc, vbOKOnly + vbInformation, "Time...."
End Sub

Private Sub refresh_branch_r12()
    Dim tdate As String, ds As ADODB.Recordset, t1 As String, t2 As String
    Dim t3 As Long, t4 As Long, t5 As Long, t6 As Long, t7 As Long, t8 As Long
    Dim brp As String, rc As Long
    rc = 0
    tdate = Format(DateAdd("d", 1, edate), "mm-dd-yyyy")
    date1 = "07-04-2015"

    Screen.MousePointer = 11
        
    q = "select o.inventory_item_id, m.segment1," & _
        " o.subinventory_code," & _
        " o.transaction_quantity" & _
        " from mtl_material_transactions o, mtl_system_items_b m" & _
        " where o.subinventory_code >= '001' and o.subinventory_code <= '075'" & _
        " and o.subinventory_code not in ('012', '015', '016')" & _
        " and m.organization_id = o.organization_id" & _
        " and m.inventory_item_id = o.inventory_item_id" & _
        " and m.segment1 >= '100' and m.segment1 <= '9999'" & _
        " and o.transaction_date < TO_DATE('" & Format(tdate, "dd-mmm-yyyy") & "')" & _
        " and o.transaction_date >= TO_DATE('" & Format(date1, "dd-mmm-yyyy") & "')"
        
    'MsgBox q
    t1 = Format(Now, "hh:mm:ss")
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            brp = branch_plant(ds(2))
            If brp > "A" Then
            '    MsgBox ds(2) & " " & ds(1)
            'Else
                For i = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = brp Then
                        Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds(3)
                        Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                        'If i = 1 Then DoEvents
                    End If
                    If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = "ALL" Then
                        Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds(3)
                        Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                        'If i = 16 Then DoEvents
                        'if ds(1) = "777" Then DoEvents
                        Exit For
                    End If
                Next i
            End If
            
            rc = rc + 1
            If Right(Format(rc, "0000000"), 5) = "00200" Then DoEvents
            'For i = 1 To Grid1.Rows - 1
            '    If Grid1.TextMatrix(i, 0) = ds(1) Then
            '        Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
            '        If i = 1 Then DoEvents
            '        Exit For
            '    End If
            'Next i
            ''DoEvents
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    q = "select o.inventory_item_id, m.segment1," & _
        " o.subinventory_code," & _
        " o.transaction_quantity" & _
        " from mtl_material_transactions o, mtl_system_items_b m" & _
        " where o.subinventory_code in ('A10', 'K10', 'T10')" & _
        " and m.organization_id = o.organization_id" & _
        " and m.inventory_item_id = o.inventory_item_id" & _
        " and m.segment1 >= '100' and m.segment1 <= '9999'" & _
        " and o.transaction_date < TO_DATE('" & Format(tdate, "dd-mmm-yyyy") & "')" & _
        " and o.transaction_date >= TO_DATE('" & Format(date1, "dd-mmm-yyyy") & "')"
        
    'MsgBox q
    't1 = Format(Now, "hh:mm:ss")
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = ds(2) Then
                    Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(3)
                    Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                    'If i = 1 Then DoEvents
                End If
                If Grid1.TextMatrix(i, 2) = ds(1) And Grid1.TextMatrix(i, 0) = "ALL" Then
                    Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(3)
                    Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds(3)
                    Exit For
                    'If i = 16 Then DoEvents
                    If ds(1) = "777" Then DoEvents
                End If
                
                
                'If Grid1.TextMatrix(i, 0) = ds(1) Then
                '    If ds(2) = "T10" Then Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + ds(3)
                '    If ds(2) = "K10" Then Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + ds(3)
                '    If ds(2) = "A10" Then Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(3)
                '    Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds(3)
                '    If i = 1 Then DoEvents
                '    Exit For
                'End If
            Next i
            'DoEvents
            rc = rc + 1
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t8 = 0
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = "ALL" Then
        '    Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 6)) + Val(Grid1.TextMatrix(i, 7))
        '    t3 = t3 + Val(Grid1.TextMatrix(i, 3))
        '    t4 = t4 + Val(Grid1.TextMatrix(i, 4))
            t5 = t5 + Val(Grid1.TextMatrix(i, 5))
            t6 = t6 + Val(Grid1.TextMatrix(i, 6))
            t7 = t7 + Val(Grid1.TextMatrix(i, 7))
        '    t8 = t8 + Val(Grid1.TextMatrix(i, 8))
        End If
    Next i
    Grid1.AddItem " "
    s = " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & "Totals" & Chr(9)
    's = s & t3 & Chr(9)
    's = s & t4 & Chr(9)
    s = s & t5 & Chr(9)
    s = s & t6 & Chr(9)
    s = s & t7 & Chr(9)
    's = s & t8
    Grid1.AddItem s
    Screen.MousePointer = 0
    MsgBox t1 & " - " & Format(Now, "hh:mm:ss") & " Records: " & rc, vbOKOnly + vbInformation, "Time...."
End Sub

Private Sub refresh_skus()
    Dim ds As ADODB.Recordset, s As String, sdesc As String, hrow As Boolean, psku As String
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 9
    's = "select sku, count(*) from bimp where plantwhs <> 'DRY' group by sku"
    s = "select sku, count(*) from bimp where plantwhs in ('A10', 'K10', 'T10') group by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & Chr(9)
            s = s & skurec(Val(ds!sku)).desc & Chr(9)
            's = s & "0" & Chr(9)
            's = s & "0" & Chr(9)
            's = s & "0" & Chr(9)
            s = s & "0" & Chr(9)
            s = s & "0" & Chr(9)
            s = s & "0"
            sdesc = skurec(Val(ds!sku)).unit & skurec(Val(ds!sku)).desc
            Grid1.AddItem "T10" & Chr(9) & "Brenham" & Chr(9) & s & Chr(9) & sdesc & "500"
            Grid1.AddItem "K10" & Chr(9) & "Broken Arrow" & Chr(9) & s & Chr(9) & sdesc & "501"
            Grid1.AddItem "A10" & Chr(9) & "Sylacauga" & Chr(9) & s & Chr(9) & sdesc & "502"
            Grid1.AddItem "ALL" & Chr(9) & "All" & Chr(9) & s & Chr(9) & sdesc & "999"
            ds.MoveNext
        Loop
    End If
    ds.Close
    List1.Clear: List2.Clear
    s = "select * from valuelists where listname = 'branchplants' order by listreturn"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            List1.AddItem ds!listreturn
            List2.AddItem ds!listdisplay
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 8: Grid1.ColSel = 8
    Grid1.Sort = 5
    Grid1.FillStyle = flexFillRepeat
    psku = Grid1.TextMatrix(1, 2)
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) <> "T10" Then
            Grid1.TextMatrix(i, 3) = "..."
            Grid1.TextMatrix(i, 4) = "..."
        End If
        If psku <> Grid1.TextMatrix(i, 2) Then
            hrow = Not hrow
            psku = Grid1.TextMatrix(i, 2)
        End If
        'If Grid1.TextMatrix(i, 0) = "ALL" Then
        If hrow = True Then
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellBackColor = Grid1.BackColorFixed
        End If
    Next i
    Grid1.Row = 1: Grid1.Col = 1
    s = "^Whs|<Location|^SKU|^Unit|<Flavor|^Plant|^Branches|^Total"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 900
    Grid1.ColWidth(1) = 1600
    Grid1.ColWidth(2) = 900
    Grid1.ColWidth(3) = 900
    Grid1.ColWidth(4) = 3000
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1500
    Grid1.ColWidth(7) = 1500
    Grid1.ColWidth(8) = 0 '1500
    Grid1.Redraw = True
End Sub


Private Sub Check1_Click()
    If Check1.Value = 1 Then
        process_orders                              'jv101116
    Else
        Grid1.Cols = 9
    End If
End Sub

Private Sub Command1_Click()
    'refresh_bimp_history
    'Exit Sub
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    refresh_skus
    'refresh_branch_r12
    refresh_bimp_history
    If Check1.Value = 1 Then process_orders                              'jv101116
End Sub

Private Sub Command2_Click()
    Dim rt As String, rh As String, rf As String
    rt = Combo1 & " Branch E-O-P Units"
    rh = "Date: " & Format(edate, "m-d-yyyy")
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    'htdc(0) = "lightcyan": gndc(0) = Me.bcolor.BackColor
    htdc(0) = "lightcyan": gndc(0) = Me.Grid1.BackColorFixed
    'htdc(1) = "yellow": gndc(1) = Me.ycolor.BackColor
    'htdc(2) = "lightgrey": gndc(2) = Me.wcolor.BackColor
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

Private Sub Form_Load()
    edate = Format(Now, "mm-dd-yyyy")
    sdate = Left(edate, 2) & "-01-" & Right(edate, 4)
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    refresh_skus
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 200
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Command1.Height * 2)
End Sub

Private Sub Grid1_RowColChange()
    Dim pals As Currency, psku As String
    Grid1.ToolTipText = ""
    If Grid1.Col >= 5 And Val(Grid1.TextMatrix(Grid1.Row, 2)) > 0 Then
        If Val(Grid1.TextMatrix(Grid1.Row, Grid1.Col)) <> 0 Then
            psku = Grid1.TextMatrix(Grid1.Row, 2)
            pals = Format(Val(Grid1.TextMatrix(Grid1.Row, Grid1.Col)) / skurec(Val(psku)).pallet, "0.00")
            Grid1.ToolTipText = Grid1.TextMatrix(0, Grid1.Col) & ": " & pals & " Pallets"
        End If
    End If
End Sub
