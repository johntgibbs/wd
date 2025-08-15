Attribute VB_Name = "Module1"
Public WDUserId     As String
Public WDbbsr       As String
Public Wdb          As ADODB.Connection
Public Sdb          As ADODB.Connection         'jv060216
Public logdir       As String '= "\\bbc-01-wdmgmt\wd\data\testlog"
Public wdlogdir     As String
Global eno          As Long
Global edesc        As String
Global vberror_log  As String
Public localAppDataPath As String

Type ptask
    id As Long
    area As String
    description As String
    source As String
    target As String
    product As String
    palletid As String
    qty As String
    uom As String
    lotnum As String
    units As String
    lotnum2 As String
    units2 As String
    status As String
    userid As String
    trandate As String
    reqid As String
End Type

Type daimessagerec
    dhostmodifytime As String
    imessagesequence As String
    smessageidentifier As String
    smessage As String
    bbcidentity As String
    bbcstatus As String
End Type

Type skuinfo
    sku As String
    uom_type As String
    desc As String
    prodname As String
    uom_per_pallet As Integer
    qty_per_pallet As Integer
End Type

Global skurec(0 To 9999) As skuinfo

Function bc000(bc As String) As String
    Dim s As String
    s = bc
    s = Mid(bc, 1, 4)
    s = s & Mid(bc, 5, 6) & "  "
    s = s & Mid(bc, 11, 3) & "  "
    s = s & Mid(bc, 14, 3)
    bc000 = s
End Function

Sub build_sku_config()
    Dim ds As ADODB.Recordset, sqlx As String, i As Integer
    'On Error GoTo vberror
    sqlx = "select * from sku_config order by sku"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = Val(ds!sku)
            If Len(ds!sku) > 0 Then skurec(i).sku = ds!sku                      'jv082415
            If Len(ds!uom_type) > 0 Then
                skurec(i).uom_type = Trim(ds!uom_type & " ")
                skurec(i).prodname = Trim(ds!uom_type & " ")
            End If
            If Len(ds!description) > 0 Then
                skurec(i).desc = Trim(ds!description & " ")
                skurec(i).prodname = skurec(i).prodname & " " & Trim(ds!description & " ")
            End If
            If Len(ds!uom_per_pallet) > 0 Then skurec(i).uom_per_pallet = ds!uom_per_pallet
            If Len(ds!qty_per_pallet) > 0 Then skurec(i).qty_per_pallet = ds!qty_per_pallet
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Function check_hold(p As ptask) As Boolean                                  'jv040615
    Dim psku As String, pcode As String, hflag As Boolean, s As String
    Dim ds As ADODB.Recordset, palno As String
    psku = Trim(Mid(p.palletid, 1, 4))
    'pcode = Mid(p.palletid, 12, 1)
    pcode = Trim(Mid(p.palletid, 11, 3))                                    'jv052515
    palno = Mid(p.palletid, 14, 3)
    hflag = False
    s = "select listreturn from valuelists where listname = 'wmsexpdate'"       'jv042115
    Set ds = Wdb.Execute(s)                                                     'jv042115
    If ds.BOF = False Then                                                      'jv042115
        ds.MoveFirst                                                            'jv042115
        If p.lotnum <= ds(0) Then hflag = True                                  'jv042115
        If p.lotnum2 > "0" Then                                                 'jv042115
            If hflag = False And p.lotnum2 <= ds(0) Then hflag = True           'jv042115
        End If                                                                  'jv042115
    End If                                                                      'jv042115
    ds.Close                                                                    'jv042115
    If hflag = False Then                                                       'jv042115
        s = "select id, spallet, epallet from holdlist where sku = '" & psku & "'"
        s = s & " and lot_num = '" & p.lotnum & "'"
        s = s & " and opcode = '" & pcode & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If palno >= ds!spallet And palno <= ds!epallet And hflag = False Then
                    's = s & " and palno >= " & ds!spallet
                    's = s & " and palno <= " & ds!epallet
                    'MsgBox "lot1: " & s, vbOKOnly, p.palletid
                    hflag = True
                End If
                If hflag = True Then Exit Do
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If                                                                      'jv042115
    If p.lotnum2 > "0" And hflag = False Then
        s = "select id, spallet, epallet from holdlist where sku = '" & psku & "'"
        s = s & " and lot_num = '" & Mid(p.lotnum2, 1, 5) & "'"
        'If Len(p.lotnum2) = 7 Then pcode = Mid(p.lotnum2, 7, 1)
        If Len(p.lotnum2) > 5 Then pcode = Trim(Mid(p.lotnum2, 6, 5))           'jv052515
        s = s & " and opcode = '" & pcode & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If palno >= ds!spallet And palno <= ds!epallet Then
                    's = s & " and palno >= " & ds!spallet
                    's = s & " and palno <= " & ds!epallet
                    'MsgBox "lot2: " & s, vbOKOnly, p.palletid
                    hflag = True
                End If
                If hflag = True Then Exit Do
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    check_hold = hflag
End Function

Public Function fixamps(s As String) As String
    Dim i As Integer, k As Integer, rs As String
    rs = ""
    i = 1: k = 1
    Do Until i = 0
        i = InStr(k, s, "&")
        If i = 0 Then
            rs = rs & Mid(s, k, Len(s))
            Exit Do
        Else
            rs = rs & Mid(s, k, i - k) & "&&"
            k = i + 1
        End If
    Loop
    fixamps = rs
End Function

Public Function fixquotes(s As String) As String
    Dim i As Integer, k As Integer, rs As String
    rs = ""
    i = 1: k = 1
    Do Until i = 0
        i = InStr(k, s, "'")
        If i = 0 Then
            rs = rs & Mid(s, k, Len(s))
            Exit Do
        Else
            rs = rs & Mid(s, k, i - k) & "''"
            k = i + 1
        End If
    Loop
    fixquotes = rs
End Function

Public Function new_pallet_task_record(parea As String) As Long
    'On Error Resume Next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    Dim sSql As String, sCols As String, sRows As String
    Dim zid As Long
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    zid = 0
    sSql = "Select id, status From paltasks Where area = '" & parea & "'"
    sSql = sSql & " and status = 'COMP'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        zid = ds!id
    End If
    ds.Close
    If zid = 0 Then
        sSql = "Select id, status From paltasks Where status = 'COMP'"
        Set ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst
            zid = ds!id
        End If
        ds.Close
    End If
    If zid > 0 Then
        sSql = "Update paltasks Set status = 'PEND' Where id = " & zid
        Wdb.Execute (sSql)
    Else
        zid = wd_seq("PalTasks")
        sSql = "Insert Into paltasks (ID) Values (" & zid & ")"
        Wdb.Execute (sSql)
    End If
    new_pallet_task_record = zid
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "new_pallet_task_record", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "new_pallet_task_record", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: new_pallet_task_record: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Function insert_rack_pallet(p As ptask) As Boolean
    'On Error Resume Next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    Dim hs As ADODB.Recordset                                                   'jv111314
    Dim sSql As String
    Dim k As Integer, j As Integer
    Dim pbbc As Boolean
    Dim pqty As Integer
    Dim pqty4 As Integer
    Dim pcap As Integer
    Dim zid As Long
    Dim i As Integer
    Dim rlot As String
    Dim psku As String, olot As String, opflag As Boolean
    Dim pa As String, ps As String, pt As String                                'jv111314
    On Error GoTo vberror
    'tracelist = tracelist & "<!-- insert_rack_pallet(" & p.target & "," & p.palletid & "," & p.lotnum & ") -->" & vbCrLf
    psku = Trim(Mid(p.palletid, 1, 4))
    'If psku <> sku_info(psku, "sku") Then
    '    Call debug_log(p.palletid & " invalid SKU in pallet barcode, during insert_rack_pallet.", p, p.userid)
    '    insert_rack_pallet = False
    '    Exit Function
    'End If

    pqty = Val(p.units) + Val(p.units2)
    'If pqty > Val(sku_info(psku, "units")) Then
    '    pbbc = False
    'Else
        pbbc = True
    'End If
    pqty = 0

    If p.target = "ORDER PICK" Or p.target = "SNACK PLANT" Or p.target = "ANTE ROOM" Then
        sSql = "Select id, capacity From racks Where aisle = 'M'"
        If p.target = "ORDER PICK" Then sSql = sSql & " And rack = 'OP'"
        If p.target = "SNACK PLANT" Then sSql = sSql & " And rack = 'SP'"
        If p.target = "ANTE ROOM" Then sSql = sSql & " And rack = 'ANTE'"
        opflag = True
    Else
        If p.target = "M-OP" Or p.target = "M-SP" Or p.target = "M-ANTE" Then
            sSql = "Select id, capacity From racks Where aisle = 'M'"
            sSql = sSql & " And rack = '" & Mid(p.target, 3, Len(p.target) - 2) & "'"
            opflag = True
        Else
            sSql = "Select rackno From rackpos Where count_qty = 0 And rackno in (" & _
                "Select id From racks Where aisle = '" & Mid(p.target, 1, 1) & "'" & _
                " And rack = '" & Mid(p.target, 3, Len(p.target) - 2) & "')"
            opflag = False
        End If
    End If

    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr

    k = 0
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        k = ds(0)
        sSql = "select hold from racks where id = " & k                 'jv111314
        Set hs = Wdb.Execute(sSql)                                      'jv111314
        If hs.BOF = False Then                                          'jv111314
            hs.MoveFirst                                                'jv111314
            If hs!hold = "1" Then                                       'jv111314
                pa = p.area                                             'jv111314
                ps = p.source                                           'jv111314
                pt = p.target                                           'jv111314
                p.area = "HOLD"                                         'jv111314
                If p.source = "ORDER PICK" Or p.source = "M-OP" Then    'jv111314
                    p.source = p.target                                 'jv111314
                End If                                                  'jv111314
                p.target = "HOLD"                                       'jv111314
                Call post_move_trans(p)                                 'jv111314
                p.area = pa                                             'jv111314
                p.source = ps                                           'jv111314
                p.target = pt                                           'jv111314
            End If                                                      'jv111314
        End If                                                          'jv111314
        hs.Close                                                        'jv111314
    End If
    ds.Close

    'If k = 0 Then
    '    'db.Close
    '    Call debug_log("failed to update rack in insert_pallet_rack: " & p.target, p, p.userid)
    '    insert_rack_pallet = False
    '    Exit Function
    'End If

    j = 0
    sSql = "Select id From rackpos Where rackno = " & k & _
           " And count_qty = 0 Order by posn_num Desc"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        j = ds!id
    End If
    ds.Close

    If j > 0 Then
        sSql = "Update rackpos Set sku = '" & psku & "'"
        sSql = sSql & ",lot_num = '" & p.lotnum & "'"
        sSql = sSql & ",pallet_num = '" & Mid(p.palletid, 14, 3) & "'"
        sSql = sSql & ",count_qty = " & Val(p.units)
        sSql = sSql & ",recv_date = '" & Format(Now, "MM-dd-yyyy") & "'"
        sSql = sSql & ",barcode = '" & p.palletid & "'"
        sSql = sSql & ",lot2 = '" & p.lotnum2 & "'"
        sSql = sSql & ",qty2 = " & Val(p.units2)
        sSql = sSql & ",wrapped = 'Y',hold = 'N'"
        If opflag = True Then
            sSql = sSql & ",posn_num = " & order_pick_position(psku)
        End If
        If pbbc = True Then
            sSql = sSql & ",bbc = 'Y'"
        Else
            sSql = sSql & ",bbc = 'N'"
        End If
        sSql = sSql & " Where id = " & j
        Wdb.Execute (sSql)
    Else
        If opflag = True Then
            zid = wd_seq("RackPos")
            sSql = "Insert Into rackpos (ID,Rackno,Posn_num,SKU,Lot_num,Pallet_num,"
            sSql = sSql & "count_qty,Recv_date,BBC,Barcode,Lot2,Qty2,Wrapped,Hold)"
            sSql = sSql & " Values (" & zid & "," & k & "," & order_pick_position(psku) & ","
            sSql = sSql & "'" & psku & "','" & p.lotnum & "',"
            sSql = sSql & "'" & Mid(p.palletid, 14, 3) & "'," & Val(p.units) & ","
            sSql = sSql & "'" & Format(Now, "MM-dd-yyyy") & "'"
            If pbbc = True Then
                sSql = sSql & ",'Y',"
            Else
                sSql = sSql & ",'N',"
            End If
            sSql = sSql & "'" & p.palletid & "',"
            sSql = sSql & "'" & p.lotnum2 & "',"
            sSql = sSql & Val(p.units2) & ","
            sSql = sSql & "'Y',"
            sSql = sSql & "'N')"
            Wdb.Execute (sSql)
        Else
            'db.Close
            'Call debug_log(sSql & "failed to update in insert_rack_pallet", p, "0")
            insert_rack_pallet = False
            Exit Function
        End If
    End If

    olot = "99999"
    pqty = 0
    pqty4 = 0

    sSql = "Select sku,lot_num,lot2,bbc From rackpos Where rackno = " & k & _
            " And count_qty > 0"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!bbc = "Y" Then
                pqty = pqty + 1
            Else
                pqty4 = pqty4 + 1
            End If
            rlot = ds!lot_num
            If Val(rlot) > 0 And rlot < olot Then olot = rlot
            rlot = ds!lot2
            If rlot > "0" Then
                If Val(rlot) > 0 And rlot < olot Then olot = rlot
            End If
            ds.MoveNext
        Loop
    Else
        ds.Close ': db.Close
        'Call debug_log(sSql & " failed in insert_rack_pallet..", p, "0")
        insert_rack_pallet = False
        Exit Function
    End If
    ds.Close

    If pqty + pqty4 = 0 Then
        sSql = "Update racks set sku = ' ',lot_num = ' ',qty = 0,qty4 = 0,resv_sku = ' ',resv_lot = ' '"
    Else
        sSql = "Update racks set sku = '" & psku & "'," & _
               "lot_num = '" & olot & "'," & _
               "qty = " & pqty & "," & _
               "qty4 = " & pqty4
    End If
    sSql = sSql & " Where id = " & k
    Wdb.Execute (sSql)
    'db.Close
    insert_rack_pallet = True
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "insert_rack_pallet", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "insert_rack_pallet", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: insert_rack_pallet: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Function insert_trans(pt As ptask) As Long
    'On Error Resume Next
    'Dim db As ADODB.Connection
    Dim sSql As String
    Dim zid As Long
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    zid = new_pallet_task_record(pt.area)
    sSql = "Update paltasks Set area='" & pt.area & "'"
    sSql = sSql & ",description='" & pt.description & "'"
    sSql = sSql & ",source='" & pt.source & "'"
    sSql = sSql & ",target='" & pt.target & "'"
    sSql = sSql & ",product='" & pt.product & "'"
    sSql = sSql & ",palletid='" & pt.palletid & "'"
    sSql = sSql & ",qty=" & Val(pt.qty)
    sSql = sSql & ",uom='" & pt.uom & "'"
    sSql = sSql & ",lotnum='" & pt.lotnum & "'"
    sSql = sSql & ",units=" & Val(pt.units)
    sSql = sSql & ",lotnum2='" & pt.lotnum2 & "'"
    sSql = sSql & ",units2=" & Val(pt.units2)
    sSql = sSql & ",status='" & pt.status & "'"
    sSql = sSql & ",userid='" & pt.userid & "'"
    sSql = sSql & ",trandate='" & pt.trandate & "'"
    sSql = sSql & ",reqid='" & pt.reqid & "'"
    sSql = sSql & " Where id = " & zid
    Wdb.Execute (sSql)
    insert_trans = zid
    'db.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "insert_trans", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "insert_trans", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: insert_trans: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Function order_pick_position(psku As String) As Integer
    'On Error Resume Next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    Dim sSql As String
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    sSql = "Select opseq From oplist Where sku = '" & psku & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        order_pick_position = ds!opseq
    Else
        order_pick_position = 0
    End If
    ds.Close ': db.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "order_pick_position", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "order_pick_position", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: order_pick_position: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Sub post_hold_log(zid As Long, ttype As String)                                      'jv010616
    Dim i As Integer, k As Integer, cfile As String
    If Form1.plantno <> "50" Then Exit Sub
    'cfile = wdlogdir & "holdlogs\hold" & holdlist.Grid5.TextMatrix(zid, 1) & ".txt"
    'Open cfile For Append As #1
    For i = 0 To holdlist.Grid5.Rows - 1
        If Val(holdlist.Grid5.TextMatrix(i, 0)) = zid Then
            cfile = wdlogdir & "holdlogs\hold" & holdlist.Grid5.TextMatrix(i, 1) & ".txt"
            Open cfile For Append As #1
            For k = 0 To holdlist.Grid5.Cols - 2
                Write #1, holdlist.Grid5.TextMatrix(i, k);
            Next k
            Write #1, Format(Now, "yyMMdd hh:mm:ss");
            Write #1, ttype
            Close #1
            Exit For
        End If
    Next i
    'Close #1
End Sub

Public Sub post_move_trans(m As ptask)
    'On Error Resume Next
    Dim cfile As String
    If UCase(m.description) = "2 STEP REQUEST" Then Exit Sub
    On Error GoTo vberror
    'tracelist = tracelist & "<!-- post_move_trans(" & m.id & ") -->" & vbCrLf
    If UCase(m.target) = "M-OP" Or UCase(m.target) = "M OP" Then m.target = "ORDER PICK"
    If UCase(m.target) = "M-ANTE" Or UCase(m.target) = "M ANTE" Then m.target = "ANTE ROOM"
    cfile = logdir & "move" & Format(Now, "MMddyyyy") & ".txt"
    Open cfile For Append As #1
    Write #1, m.id, m.area, m.description, m.source, m.target, m.product;
    Write #1, m.palletid, m.qty, m.uom, m.lotnum, m.units, m.lotnum2, m.units2;
    'Write #1, m.status, m.userid, m.trandate, m.reqid
    Write #1, m.status, WDUserId, m.trandate, m.reqid                   'jv121614
    Close #1
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "post_move_trans", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "post_move_trans", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: post_move_trans: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Sub post_pick_trans(m As ptask, fdate As String)
    'On Error Resume Next
    Dim cfile As String
    If UCase(m.description) = "2 STEP REQUEST" Then Exit Sub
    On Error GoTo vberror
    'tracelist = tracelist & "<!-- post_pick_trans(" & m.id & ") -->" & vbCrLf
    If UCase(m.target) = "M-OP" Or UCase(m.target) = "M OP" Then m.target = "ORDER PICK"
    If UCase(m.target) = "M-ANTE" Or UCase(m.target) = "M ANTE" Then m.target = "ANTE ROOM"
    'cfile = wdlogdir & "pick" & Format(Now, "MMddyyyy") & ".txt"
    cfile = wdlogdir & "pick" & fdate & ".txt"
    Open cfile For Append As #1
    'MsgBox cfile
    Write #1, m.id, m.area, m.description, m.source, m.target, m.product;
    Write #1, m.palletid, m.qty, m.uom, m.lotnum, m.units, m.lotnum2, m.units2;
    'Write #1, m.status, m.userid, m.trandate, m.reqid
    Write #1, m.status, WDUserId, m.trandate, m.reqid                   'jv121614
    Close #1
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "post_pick_trans", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "post_pick_trans", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: post_pick_trans: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub


Sub post_queue_to_sr(Whs As Integer, p As ptask)
    Dim ds As ADODB.Recordset, s As String, zid As Long
    Dim qid As Long, hf As Boolean, nque As Integer
    hf = False
    zid = new_pallet_queue()
    'Process Queue
    s = "select * from queue_infc where id = " & zid
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "update queue_infc set whse_num = " & Whs
        s = s & ",sku = '" & Trim(Left(p.product, 4)) & "'"
        s = s & ",lot_num = '" & p.lotnum & "'"
        If hf Then
            s = s & ",drop_flag = 'H'"
        Else
            s = s & ",drop_flag = ' '"
        End If
        s = s & ",rack_num = " & Val(p.qty)
        s = s & ",units = " & Val(p.units)
        s = s & ",lot_num2 = '" & p.lotnum2 & "'"
        s = s & ",units2 = " & Val(p.units2)
        s = s & ",palletid = '" & p.palletid & "'"
        s = s & ",source = 'FG" & Whs & "'" ''TML'"
        s = s & " where id = " & ds!id
        Wdb.Execute s
    Else
        nque = 50
        qid = wd_seq("Queue_Infc")
        s = "INSERT INTO Queue_Infc (ID, Whse_Num, SKU, Lot_Num, Drop_Flag, Queue_Num,"
        s = s & " Rack_Num, Units, Lot_Num2, Units2, PalletID, Source)"
        s = s & " VALUES (" & qid & ","
        s = s & Whs & ","
        s = s & "'" & Trim(Left(p.product, 4)) & "',"
        s = s & "'" & p.lotnum & "',"
        If hf Then
            s = s & "'H',"
        Else
            s = s & "' ',"
        End If
        s = s & nque & ","
        s = s & Val(p.qty) & ","
        s = s & Val(p.units) & ","
        s = s & "'" & p.lotnum2 & "',"
        s = s & Val(p.units2) & ","
        s = s & "'" & p.palletid & "',"
        s = s & "'FG" & Whs & "')"     '"'TML')"
        Wdb.Execute s
    End If
    ds.Close
End Sub
Public Function r12lot(plot As String) As String
    Dim s As String, t As String
    If Val(plot) > 0 Then
        t = "1-1-20" & Left(plot, 2)
        'MsgBox t
        s = Format(DateAdd("d", Val(Right(plot, 3)) - 1, t), "MM-dd-yyyy")
        'MsgBox s
        s = Format(DateAdd("yyyy", 2, s), "MM-dd-yyyy")
        'MsgBox s
        s = Format(s, "MMddyy")
        If Left(plot, 2) Mod 4 = 0 And Val(Right(plot, 3)) = 60 Then
            s = "0229"
            s = s & Val(Left(plot, 2)) + 2
        End If
    End If
    r12lot = s
End Function

Public Sub update_trans(pt As ptask)
    'On Error Resume Next
    Dim sSql As String
    'Dim db As ADODB.Connection
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr

    sSql = "Update paltasks Set area = '" & pt.area & "'," & _
            "description = '" & pt.description & "'," & _
            "source = '" & pt.source & "'," & _
            "target = '" & pt.target & "'," & _
            "product = '" & pt.product & "'," & _
            "palletid = '" & pt.palletid & "'," & _
            "qty = " & Val(pt.qty) & "," & _
            "uom = '" & pt.uom & "'," & _
            "lotnum = '" & pt.lotnum & "'," & _
            "units = " & Val(pt.units) & "," & _
            "lotnum2 = '" & pt.lotnum2 & "'," & _
            "units2 = " & Val(pt.units2) & "," & _
            "status = '" & pt.status & "'," & _
            "userid = '" & pt.userid & "'," & _
            "trandate = '" & pt.trandate & "'" & _
            " Where id = " & pt.id

    Wdb.Execute (sSql)
    'db.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "update_trans", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "update_trans", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: update_trans: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Sub vb_elog(eno As Long, edesc As String, pform As String, psub As String, puser As String)
    Dim i As Integer, s As String, cfile As String
    On Error GoTo vberror
    'cfile = "\\bbc-01-prodtrk\wd\temp\sqlerrors.txt"
    cfile = vberror_log
    'i = FreeFile(1)
    i = 88
    Open cfile For Append As #i
    Write #i, eno, edesc, pform, psub, Format(Now, "M-d-yyyy h:mm am/pm"), puser
    Close #i
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    If eno = 52 Then
        If MsgBox("Local network connection has been lost.", vbRetryCancel + vbInformation, "try another location...") = vbCancel Then
            End
        Else
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: vb_elog: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End If
End Sub
Public Function wd_seq(tbname As String) As Long
    'On Error Resume Next
    Dim sSql As String
    Dim i As Long
    'Dim db As ADODB.Connection
    Dim ds As ADODB.Recordset
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    sSql = "Select sequence_id From sequences where seq = '" & tbname & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        i = ds!sequence_id + 1
        sSql = "Update sequences Set sequence_id = " & i & " Where seq = '" & tbname & "'"
        Wdb.Execute (sSql)
    Else
        i = 100
        sSql = "Insert Into sequences (sequence_id, seq) Value (" & i & ",'" & tbname & "')"
    End If
    'tracelist = tracelist & "<!-- wd_seq(" & tbname & ") = " & i & " -->" & vbCrLf
    ds.Close ': db.Close
    wd_seq = i
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "wd_seq", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "wd_seq", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: wd_seq: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Function new_pallet_queue() As Long
    'On Error Resume Next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    Dim sSql As String, sCols As String, sRows As String
    Dim k As Integer, ncnt As Integer
    Dim zid As Long
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
  
    sSql = "Select max(queue_num) from queue_infc"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        k = ds(0) + 1
    Else
        k = 100
    End If
    ds.Close

    sSql = "Select id, queue_num from queue_infc where queue_num = 0"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        zid = ds!id
        sSql = "update queue_infc set queue_num = " & k & _
               "Where id = " & zid
        Wdb.Execute (sSql)
    Else
        zid = wd_seq("Queue_infc")
        sSql = "Insert into queue_infc (ID, Queue_num) VALUES (" & _
               zid & "," & k & ")"
        Wdb.Execute (sSql)
    End If
    ds.Close ': db.Close
    new_pallet_queue = zid
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "new_pallet_queue", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "new_pallet_queue", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: new_pallet_queue: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Function barcode_to_lotnum(mbar As String) As String
    Dim s1 As String, s2 As String, s As String, j As Long
    If Len(mbar) <> 16 Then
        barcode_to_lotnum = "01001"
        'MsgBox "len<>16"
    Else
        j = Val(Mid(mbar, 5, 2))
        If j < 1 Or j > 12 Then s = "01001"
        'MsgBox "month=" & j
        j = Val(Mid(mbar, 7, 2))
        If j < 1 Or j > 31 Then s = "01001"
        'MsgBox "day=" & j
        j = Val(Mid(mbar, 9, 2))
        If j < 11 Or j > 44 Then s = "01001"
        'MsgBox "Year=" & j
        If s <> "01001" Then
            s1 = "01-01-20" & Format(CInt(Mid(mbar, 9, 2)) - 2, "00")
            s2 = Mid(mbar, 5, 2) & "-" & Mid(mbar, 7, 2) & "-20" & Format(CInt(Mid(mbar, 9, 2)) - 2, "00")
            j = DateDiff("d", s1, s2) + 1
            s = Format(CInt(Mid(mbar, 9, 2)) - 2, "00")
            s = s & Format(j, "000")
        End If
        barcode_to_lotnum = s
    End If
End Function

Public Function wdempname(empid As String) As String
    Dim s As String, ds As ADODB.Recordset
    s = "Select listdisplay from valuelists where listname = 'wdempid'"
    s = s & " and listreturn = '" & empid & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        wdempname = ds(0)
    Else
        If empid = "999999" Then
            wdempname = "Maintenance"
        Else
            wdempname = empid
        End If
    End If
    ds.Close
End Function

