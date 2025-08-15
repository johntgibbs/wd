VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bimpoutstock 
   Caption         =   "Out of Stock"
   ClientHeight    =   11385
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   ScaleHeight     =   11385
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   9135
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   16113
      _Version        =   327680
      BackColorFixed  =   12648384
      BackColorSel    =   12583104
      FocusRect       =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Post to W/D Browser"
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
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Width           =   2655
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
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6165
      _Version        =   327680
   End
   Begin VB.Label rcolor 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label1"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu eddate 
         Caption         =   "Change Available Date"
      End
      Begin VB.Menu delsku 
         Caption         =   "Drop Product From List"
      End
      Begin VB.Menu inssku 
         Caption         =   "Add Product to List"
      End
   End
End
Attribute VB_Name = "bimpoutstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub plant_trailers()
    Dim s As String, ds As ADODB.Recordset, i As Integer
    For i = 0 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 3) = "Not Scheduled" Then
            s = "select shipdate from trailers where sku = '" & Grid2.TextMatrix(i, 1) & "' and branch = "
            If Grid2.TextMatrix(i, 0) = "A10" Then s = s & "52"
            If Grid2.TextMatrix(i, 0) = "K10" Then s = s & "47"
            If Grid2.TextMatrix(i, 0) = "T10" Then s = s & "1"
            s = s & " order by shipdate"
            'MsgBox s
            Set ds = wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                Grid2.TextMatrix(i, 3) = Format(DateAdd("d", 1, ds!shipdate), "M-d-yyyy")
                'MsgBox s
            End If
            ds.Close
        End If
    Next i
End Sub

Private Sub prod_schedule(porg As String, pdays As Integer, hdays As Integer)
    Dim q As String, ds As ADODB.Recordset, i As Integer
    'R12
    q = "select itm.segment1,mtl.organization_code,MIN(plan_start_date)"
    q = q & " from gme_batch_header hdr,mtl_parameters mtl,"
    q = q & "mtl_system_items_b itm ,gme_material_details dtl"
    q = q & " Where trunc(hdr.plan_start_date) >= trunc(SYSDATE - " & pdays & ")"
    q = q & " and hdr.batch_status in (1,2)"
    q = q & " and hdr.batch_id = dtl.batch_id"
    q = q & " and mtl.organization_code = '" & porg & "'"
    q = q & " and mtl.organization_id = hdr.organization_id"
    q = q & " and dtl.line_type = 1"
    q = q & " and dtl.inventory_item_id = itm.inventory_item_id"
    q = q & " and itm.organization_id = hdr.organization_id"
    q = q & " and itm.segment1 < '999'"
    q = q & " group by itm.segment1,mtl.organization_code"
    q = q & " order by 1,3"
    'MsgBox q
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid2.Rows - 1
                If Grid2.TextMatrix(i, 0) = "T10" And ds(1) = "500" And Grid2.TextMatrix(i, 1) = ds(0) Then
                'If Grid2.TextMatrix(i, 0) = "T10" And ds(1) = porg And Grid2.TextMatrix(i, 1) = ds(0) Then
                    'Grid2.TextMatrix(i, 3) = Format(ds(2), "M-d-yyyy")
                    Grid2.TextMatrix(i, 3) = Format(DateAdd("d", hdays, ds(2)), "M-d-yyyy")         'jv083117
                    'If ds(0) = "765" Then MsgBox ":" & ds(1) & " T10:" & ds(2)
                End If
                If Grid2.TextMatrix(i, 0) = "K10" And ds(1) = "501" And Grid2.TextMatrix(i, 1) = ds(0) Then
                'If Grid2.TextMatrix(i, 0) = "K10" And ds(1) = porg And Grid2.TextMatrix(i, 1) = ds(0) Then
                    'Grid2.TextMatrix(i, 3) = Format(ds(2), "M-d-yyyy")
                    Grid2.TextMatrix(i, 3) = Format(DateAdd("d", hdays, ds(2)), "M-d-yyyy")         'jv083117
                    'If ds(0) = "765" Then MsgBox ":" & porg & ":" & ds(1) & ": K10:" & ds(2)
                End If
                If Grid2.TextMatrix(i, 0) = "A10" And ds(1) = "502" And Grid2.TextMatrix(i, 1) = ds(0) Then
                'If Grid2.TextMatrix(i, 0) = "A10" And ds(1) = porg And Grid2.TextMatrix(i, 1) = ds(0) Then
                    'Grid2.TextMatrix(i, 3) = Format(ds(2), "M-d-yyyy")
                    Grid2.TextMatrix(i, 3) = Format(DateAdd("d", hdays, ds(2)), "M-d-yyyy")         'jv083117
                    'If ds(0) = "765" Then MsgBox porg & " A10:" & ds(2)
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub


Private Sub refresh_grid(plantcode As String)
    Dim i As Integer, ss As ADODB.Recordset
    Dim ds As ADODB.Recordset, sqlx As String, oqty As Long
    Dim tc As Integer, mpal As Integer, np As Long, c As Long, k As Long
    Dim t2 As Long, t3 As Long, t4 As Long, t5 As Long
    Dim psource As Integer
    'On Error GoTo vberror
    sqlx = "select sku,plantpool,sum(onorder),sum(thiswknewpals),sum(nextwknewpals)"
    sqlx = sqlx & " from bimp"
    sqlx = sqlx & " where plantwhs = '" & plantcode & "'"
    sqlx = sqlx & " and sku not in (select sku from discont where plantwhs = '" & plantcode & "')"
    sqlx = sqlx & " group by sku, plantpool"
    sqlx = sqlx & " order by sku"
    'MsgBox sqlx
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = Val(ds(0))
            sqlx = plantcode & Chr(9)
            sqlx = sqlx & ds(0) & Chr(9)                                            'SKU
            If skurec(i).sku <> ds(0) Then
                sqlx = sqlx & "...." & Chr(9)
                mpal = 0: psource = 0
            Else
                sqlx = sqlx & skurec(i).unit & " " & skurec(i).desc & Chr(9)        'Product
                mpal = skurec(i).pallet: psource = skurec(i).psrc
            End If
            np = 0
            s = "select poolsched from bimp where sku = '" & ds(0) & "'"
            s = s & " and plantwhs = '" & plantcode & "'"
            s = s & " and poolsched > 0"
            Set ss = wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                np = ss(0)
            End If
            ss.Close
            
            np = np + plant_transfers(plantcode, ds(0))                              'jv090216
            
            
            oqty = ds(2)
            If IsNull(ds(3)) = False Then oqty = oqty + (ds(3) * mpal)
            If IsNull(ds(4)) = False Then oqty = oqty + (ds(4) * mpal)
            s = "select sku, sum(netqty) from brorders where sku = '" & ds(0) & "'"
            If plantcode = "T10" Then s = s & " and plant = 50"
            If plantcode = "K10" Then s = s & " and plant = 51"
            If plantcode = "A10" Then s = s & " and plant = 52"
            s = s & " group by sku having sum(netqty) <> 0"
            Set ss = wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                'MsgBox s
                If IsNull(ss(1)) = False Then
                    oqty = oqty + (ss(1) * mpal)
                    'If ss(1) > 0 Then MsgBox s & " = " & ss(1), vbOKOnly + vbInformation, "Branch Orders.."
                End If
            End If
            ss.Close
            oqty = oqty + (groupitems_qty(ds(0), plantcode, "ALL") * mpal)          'jv081516
            
            sqlx = sqlx & Format(ds(1) + ds(2), "#") & Chr(9)                                       'Plant Units
            sqlx = sqlx & Format((np * mpal), "#") & Chr(9)                              'New Pool Units
            'sqlx = sqlx & Format(ds(2) + (ds(3) * mpal) + (ds(4) * mpal), "#") & Chr(9)     'Branch Orders
            sqlx = sqlx & Format(oqty, "#") & Chr(9)     'Branch Orders
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    
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

Private Sub build_grid2()
    Dim s As String, i As Integer, ds As ADODB.Recordset
    Dim t10daze As Integer, k10daze As Integer, a10daze As Integer
    Dim t10date As String, k10date As String, a10date As String
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Cols = 6: Grid2.Rows = 1
    Grid2.FixedCols = 2
    Grid2.Clear
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 8)) >= Val(Grid1.TextMatrix(i, 6)) Then
                s = Grid1.TextMatrix(i, 0) & Chr(9)
                s = s & Grid1.TextMatrix(i, 1) & Chr(9)
                s = s & Grid1.TextMatrix(i, 2) & Chr(9)
                s = s & "Not Scheduled" & Chr(9)
                s = s & Grid1.TextMatrix(i, 6) & Chr(9)
                s = s & Grid1.TextMatrix(i, 8)
                Grid2.AddItem s
            End If
        Next i
    End If
    s = "select listreturn, listdisplay from valuelists where listname = 'bimpproddates'"   'jv091616
    's = s & " and listreturn = '" & List1 & "'"                                 'jv091616
    Set ds = wdb.Execute(s)                                                     'jv091616
    If ds.BOF = False Then                                                      'jv091616
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!listreturn & Chr(9) & "_" & Chr(9)
            If ds!listreturn = "T10" Then
                t10date = ds!listdisplay                                                        'jv083117
                t10date = InputBox("T10 Production thru:", "T10 production...", t10date)        'jv083117
                If Len(t10date) = 0 Then t10date = ds!listdisplay                               'jv083117
                If IsDate(t10date) = False Then t10date = ds!listdisplay                        'jv083117
                s = s & "Brenham: Production Thru: " & t10date                                  'jv083117
                t10daze = DateDiff("d", t10date, Now) - 1                                       'jv083117
                's = s & "Brenham: Production Thru: " & ds!listdisplay
                't10daze = DateDiff("d", ds!listdisplay, Now) '- 1
            End If
            If ds!listreturn = "K10" Then
                k10date = ds!listdisplay                                                        'jv083117
                k10date = InputBox("K10 Production thru:", "K10 production...", k10date)        'jv083117
                If Len(k10date) = 0 Then k10date = ds!listdisplay                               'jv083117
                If IsDate(k10date) = False Then k10date = ds!listdisplay                        'jv083117
                s = s & "Broken Arrow: Production Thru: " & k10date                             'jv083117
                k10daze = DateDiff("d", k10date, Now) - 1                                       'jv083117
                's = s & "Broken Arrow: Production Thru: " & ds!listdisplay
                'k10daze = DateDiff("d", ds!listdisplay, Now) '- 1
            End If
            If ds!listreturn = "A10" Then
                a10date = ds!listdisplay                                                        'jv083117
                a10date = InputBox("A10 Production thru:", "A10 production...", a10date)        'jv083117
                If Len(a10date) = 0 Then a10date = ds!listdisplay                               'jv083117
                If IsDate(a10date) = False Then a10date = ds!listdisplay                        'jv083117
                s = s & "Sylacauga: Production Thru: " & a10date                                'jv083117
                a10daze = DateDiff("d", a10date, Now) - 1                                       'jv083117
                's = s & "Sylacauga: Production Thru: " & ds!listdisplay
                'a10daze = DateDiff("d", ds!listdisplay, Now) '- 1
            End If
            s = s & Chr(9) & "_"
            'daze = DateDiff("d", ds!listdisplay, Now) '+ 1                                     'jv083117
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close                                                                    'jv091616
    
    s = InputBox("T10 Product Hold Days", "T10 Product Hold days...", "7")                      'jv083117
    If Val(s) < 0 Then s = "0"                                                                  'jv083117
    Call prod_schedule("500", t10daze, Val(s))                                                  'jv083117
    'Call prod_schedule("500", t10daze)
    s = InputBox("K10 Product Hold Days", "K10 Product Hold days...", "7")                      'jv083117
    If Val(s) < 0 Then s = "0"                                                                  'jv083117
    Call prod_schedule("501", k10daze, Val(s))                                                  'jv083117
    'Call prod_schedule("501", k10daze)
    s = InputBox("A10 Product Hold Days", "A10 Product Hold days...", "7")                      'jv083117
    If Val(s) < 0 Then s = "0"                                                                  'jv083117
    Call prod_schedule("502", a10daze, Val(s))                                                  'jv083117
    'Call prod_schedule("502", a10daze)
    
    Call plant_trailers                                 'jv100617
    
    Grid2.Row = 1: Grid2.RowSel = 1
    Grid2.Col = 0: Grid2.ColSel = 1
    Grid2.Sort = 5
    Grid2.FillStyle = flexFillRepeat
    For i = 1 To Grid2.Rows - 1
        If Val(Grid2.TextMatrix(i, 1)) = 0 Then
            Grid2.Row = i: Grid2.RowSel = i
            Grid2.Col = 0: Grid2.ColSel = Grid2.Cols - 1
            Grid2.CellBackColor = rcolor.BackColor
        End If
    Next i
    Grid2.Row = 1
    'Grid2.FormatString = "^Plant|^SKU|<Product|^Schedule Date|^Net|^OutQty"
    Grid2.FormatString = "^Plant|^SKU|<Product|^Next Available Date|^Net|^OutQty"    'jv083117
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 4000
    Grid2.ColWidth(3) = 2000
    Grid2.ColWidth(4) = 1000
    Grid2.ColWidth(5) = 1000
    Grid2.Redraw = True
End Sub

Private Sub Command1_Click()
    Dim i As Long, np As Long, mpal As Integer, k As Long
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub


    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 9: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    
    

    'Grid1.AddItem "T10" & Chr(9) & " " & Chr(9) & "Brenham"
    Call refresh_grid("T10")
    'Grid1.AddItem "K10" & Chr(9) & " " & Chr(9) & "Broken Arrow"
    Call refresh_grid("K10")
    Call refresh_grid("A10")

    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            np = Val(Grid1.TextMatrix(i, 3)) + Val(Grid1.TextMatrix(i, 4))
            np = np - Val(Grid1.TextMatrix(i, 5))
            Grid1.TextMatrix(i, 6) = Format(np, "#")
            'If Option2 = True Then
                mpal = 0
                If skurec(Val(Grid1.TextMatrix(i, 1))).sku = Grid1.TextMatrix(i, 1) Then
                    mpal = skurec(Val(Grid1.TextMatrix(i, 1))).pallet
                End If
                If mpal <> 0 Then
                    Grid1.TextMatrix(i, 3) = CInt(Val(Grid1.TextMatrix(i, 3)) / mpal)
                    Grid1.TextMatrix(i, 4) = CInt(Val(Grid1.TextMatrix(i, 4)) / mpal)
                    Grid1.TextMatrix(i, 5) = CInt(Val(Grid1.TextMatrix(i, 5)) / mpal)
                    np = Val(Grid1.TextMatrix(i, 3)) + Val(Grid1.TextMatrix(i, 4))
                    np = np - Val(Grid1.TextMatrix(i, 5))
                    Grid1.TextMatrix(i, 6) = Format(np, "#")
                End If
        Next i
    End If
    
    For i = Grid1.Rows - 1 To 1 Step -1                                                             'jv101016
        k = Val(Grid1.TextMatrix(i, 3)) + Val(Grid1.TextMatrix(i, 4)) + Val(Grid1.TextMatrix(i, 5)) 'jv101016
        If k = 0 Then                                                                               'jv101016
            'If Grid1.Rows > 2 Then                                                                  'jv101016
            '    Grid1.RemoveItem i                                                                  'jv101016
            'Else                                                                                    'jv101016
            '    Grid1.Rows = 1                                                                      'jv101016
            'End If                                                                                  'jv101016
        End If                                                                                      'jv101016
    Next i                                                                                          'jv101016
    
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            Grid1.TextMatrix(i, 7) = plant_lowstock(Grid1.TextMatrix(i, 0), Grid1.TextMatrix(i, 1))
            Grid1.TextMatrix(i, 8) = plant_outstock(Grid1.TextMatrix(i, 0), Grid1.TextMatrix(i, 1))
        Next i
    End If
    
    
    Grid1.FormatString = "^Plant|^SKU|<Product|^Plant Pallets|^New Pool|^Branch Orders|^Net|^LowQty|^OutQty"
    
    Grid1.FillStyle = flexFillRepeat
    c = Grid1.BackColor
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
        Grid1.CellBackColor = c
        If c = Grid1.BackColorFixed Then
            c = Grid1.BackColor
        Else
            c = Grid1.BackColorFixed
        End If
        'If Val(Grid1.TextMatrix(i, 5)) < 0 Then Grid1.CellForeColor = rcolor.ForeColor
    Next i
    Grid1.Row = 1
    
    
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 600
    Grid1.ColWidth(2) = 4000
    Grid1.ColWidth(3) = 1400 '6 '00
    Grid1.ColWidth(4) = 1400
    Grid1.ColWidth(5) = 1400 '6 '00
    Grid1.ColWidth(6) = 1400
    Grid1.ColWidth(7) = 1400
    Grid1.ColWidth(8) = 1400
    Grid1.Redraw = True
    
    build_grid2
    Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
    Dim rt As String, rf As String, rh As String
    rt = "Out of Stock Items"
    rh = Me.Caption
    rf = "Posted: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "seagreen": gndc(0) = Me.Grid1.BackColorFixed
    htdc(1) = "cyan": gndc(1) = Me.rcolor.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid2.BackColor
    Grid2.Redraw = False
    Grid2.ColWidth(4) = 0
    Grid2.ColWidth(5) = 0
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        'Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid2, rt, rh, rf, "linen", "khaki", "white")
        Call htmlcolorgrid(Me, "\\BBC-03-FILESVR\SharedGroups\wd\html\stock.htm", Grid2, rt, rh, rf, "linen", "khaki", "white")
        Grid2.ColWidth(4) = 1000
        Grid2.ColWidth(5) = 1000
        Grid2.Redraw = True
        'i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        i = Shell("C:\program files\internet explorer\iexplore.exe \\BBC-03-FILESVR\SharedGroups\wd\html\stock.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        'Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid2, rt, rh, rf, "linen", "khaki", "white")
        Call htmlcolorgrid(Me, "\\BBC-03-FILESVR\SharedGroups\wd\html\stock.htm", Grid2, rt, rh, rf, "linen", "khaki", "white")
        Grid2.ColWidth(4) = 1000
        Grid2.ColWidth(5) = 1000
        Grid2.Redraw = True
        'i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe \\BBC-03-FILESVR\SharedGroups\wd\html\stock.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub delsku_Click()                                              'jv083117
    Dim s As String
    If Val(Grid2.TextMatrix(Grid2.Row, 1)) = 0 Then Exit Sub
    s = "Remove " & Grid2.TextMatrix(Grid2.Row, 2) & " from " & Grid2.TextMatrix(Grid2.Row, 0)
    s = s & " Plant?"
    If MsgBox(s, vbYesNo + vbQuestion, "remove product...") = vbNo Then Exit Sub
    If Grid2.Rows > 2 Then
        Grid2.RemoveItem Grid2.Row
    Else
        Grid2.Rows = 1
    End If
End Sub

Private Sub eddate_Click()                                              'jv083117
    Dim pdate As String
    If Val(Grid2.TextMatrix(Grid2.Row, 1)) = 0 Then Exit Sub
    pdate = Grid2.TextMatrix(Grid2.Row, 3)
    pdate = InputBox("Available Date:", "Available data....", pdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then Exit Sub
    Grid2.TextMatrix(Grid2.Row, 3) = Format(pdate, "M-d-yyyy")
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    'Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    Grid2.Width = Me.Width - 180
    If Me.Height > 2000 Then
        Grid2.Height = Me.Height - (Command1.Height * 2.5)
    End If
End Sub

Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub inssku_Click()                                              'jv083117
    Dim s As String, psku As String
    If Val(Grid2.TextMatrix(Grid2.Row, 1)) = 0 Then Exit Sub
    psku = InputBox("SKU:", "Add SKU.....", " ")
    If Len(psku) = 0 Then Exit Sub
    If skurec(Val(psku)).sku <> psku Then Exit Sub
    s = Grid2.TextMatrix(Grid2.Row, 0) & Chr(9)
    s = s & psku & Chr(9)
    s = s & skurec(Val(psku)).unit & " " & skurec(Val(psku)).desc & Chr(9)
    s = s & Grid2.TextMatrix(Grid2.Row, 3) & Chr(9)
    s = s & Grid2.TextMatrix(Grid2.Row, 4) & Chr(9)
    s = s & Grid2.TextMatrix(Grid2.Row, 5)
    Grid2.AddItem s, Grid2.Row
End Sub
