Attribute VB_Name = "Module1"

Public bbsr As String
Public daioradb As String
Public daisqldb As String
Public tbbsr As String
Public w1cap As Integer
Public w2cap As Integer
Public w3cap As Integer
Public w4cap As Integer
Public w5cap As Integer
Public pallogs As String
Public daimesstext As String
Public daiplate As String
Public daibay As String
Public dailogs As String
Public daiorderid As String
Public daiitem As String
Public dailotnum As String
Public bbcbranches(99) As String
Public shipdb As String
Public lastque As Long
Public sr5_lane_data As String
Public wms_sr5_data As String
Public Wdb As adodb.Connection
Public DaiDb As adodb.Connection
Public vberror_log As String
Global eno As Long
Global edesc As String
Public WDUserId   As String
Global labpix(9999) As labpic               'jv082415
Global labfmtfile As String

Type daiexprct
    action As String
    sOrderID As String
    dExpectedDate As String
    sitem As String
    slot As String
    fExpectedQuantity As String
    sStoreDestination As String
    sRouteID As String
    sHoldReason As String
End Type

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

Public Type labpic
    sku As String
    package As String
    name1 As String
    name2 As String
    name3 As String
End Type


Function barcode_to_lotnum(mbar As String) As String
    Dim s1 As String, s2 As String, s As String, j As Long
    If Len(mbar) <> 16 Then
        barcode_to_lotnum = "01001"
    Else
        j = Val(Mid(mbar, 5, 2))
        If j < 1 Or j > 12 Then s = "01001"
        j = Val(Mid(mbar, 7, 2))
        If j < 1 Or j > 31 Then s = "01001"
        j = Val(Mid(mbar, 9, 2))
        If j < 11 Or j > 44 Then s = "01001"
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

Function bbpallet_units(psku As String) As Long
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    s = "select uom_per_pallet from sku_config where sku = '" & psku & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        bbpallet_units = ds(0)
    Else
        bbpallet_units = 1
    End If
    ds.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "bbpallet_units", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: bbpallet_units: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Sub build_branch_tab()
    Dim Sdb As adodb.Connection
    Dim ds As adodb.Recordset, s As String
    Set Sdb = CreateObject("ADODB.Connection")
    Sdb.Open shipdb
    s = "select branch,branchname from branches where branch < 100"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            bbcbranches(ds!branch) = StrConv(ds!branchname, vbProperCase)
            ds.MoveNext
        Loop
    End If
    ds.Close
    Sdb.Close
End Sub
Function check_dai_plate(bc As String) As String
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    s = "select plateno from pallets where barcode = '" & bc & "'"
    s = s & " and status not in ('Shipped', 'Order Pick')"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        check_dai_plate = ds!plateno
    Else
        check_dai_plate = "None"
    End If
    ds.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "check_dai_plate", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: check_dai_plate: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Function check_hold(p As ptask) As Boolean                                  'jv040615
    Dim psku As String, pcode As String, hflag As Boolean, s As String
    Dim ds As adodb.Recordset, palno As String
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
                If palno >= ds!spallet And palno <= ds!epallet Then
                    hflag = True
                End If
                If hflag = True Then Exit Do
                ds.MoveNext
            Loop
        End If
        ds.Close                                                        'jv081716
    End If                                                                      'jv042115
    'ds.Close
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

Sub clear_sr_lane(pno As String, pitem As String, plot As String)
    Dim ds As adodb.Recordset, s As String, cfile As String
    Dim lkey As Long, pkey As Long, plane As String, pbar As String
    Dim plot1 As String, plot2 As String, pqty1 As Integer, pqty2 As Integer, psku As String    'jv060117
    On Error GoTo vberror
    'cfile = dailogs & "SR5" & Format(Now, "MMdd") & ".csv"
    cfile = pallogs & "SR" & Format(Now, "MMddyyyy") & ".txt"                                   'jv060117
    's = "select id,laneno,barcode from position where barcode in ("
    s = "select id,laneno,barcode,sku,lot_num,count_qty,lot2,qty2 from position where barcode in (" 'jv060117
    If pitem <= "0" Or plot <= "0" Then
        s = s & "select barcode from pallets where plateno = '" & pno & "')"
    Else
        s = s & "select barcode from pallets where plateno = '" & pno
        s = s & "' and sku = '" & pitem & "' and lot1 = '" & Left(plot, 5) & "')"
    End If
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        lkey = ds!laneno
        pkey = ds!id
        pbar = ds!barcode
        psku = ds!sku                               'jv060117
        plot1 = ds!lot_num                          'jv060117
        pqty1 = ds!count_qty                        'jv060117
        plot2 = ds!lot2                             'jv060117
        pqty2 = ds!qty2                             'jv060117
    End If
    ds.Close
    Open cfile For Append As 3
    If lkey > 0 And pkey > 0 Then
        s = "select * from lane where id = " & lkey
        Set ds = Wdb.Execute(s)
        If ds.BOF Then
            ds.MoveFirst
            plane = ds!zone_num & " "
            plane = plane & ds!vert_loc & " "
            plane = plane & ds!horz_loc & " "
            plane = plane & ds!rack_side
        End If
        ds.Close
        s = "Update lane Set sku = ' ',lot_num = ' ', qty = 0, gmasize = 0, lane_status = ' '" 'jv082313
        s = s & " Where id = " & lkey
        Wdb.Execute s
        s = "Update position Set sku = ' ',lot_num = ' ',pallet_num = ' '"
        s = s & ",count_qty = 0,barcode = ' ',lot2 = ' ',qty2 = 0"
        s = s & " Where id = " & pkey
        Wdb.Execute s
        'Write #3, "SR-5";
        'Write #3, "...";
        'Write #3, Trim(Left(pbar, 4));
        'Write #3, barcode_to_lotnum(pbar);
        'Write #3, Right(pbar, 3);
        'Write #3, pbar;
        'Write #3, pno;
        'Write #3, plane;
        'Write #3, "DOCK";
        'Write #3, Format(Now, "h:mm am/pm")
        Write #3, lkey;                             'p.id                   jv060117
        Write #3, "SR-5";                           'p.area                 jv060117
        Write #3, daiorderid;                       'p.description          jv060117
        Write #3, "SR-5";                           'p.source               jv060117
        Write #3, "DOCK";                           'p.target               jv060117
        s = psku
        If labpix(Val(psku)).package > " " Then s = s & " " & labpix(Val(psku)).package
        If labpix(Val(psku)).name1 > " " Then s = s & " " & labpix(Val(psku)).name1
        If labpix(Val(psku)).name2 > " " Then s = s & " " & labpix(Val(psku)).name2
        If labpix(Val(psku)).name3 > " " Then s = s & " " & labpix(Val(psku)).name3
        Write #3, s;                                'p.product              jv060117
        Write #3, pbar;                             'p.palletid             jv060117
        Write #3, "1";                              'p.qty                  jv060117
        Write #3, "Pallet";                         'p.uom                  jv060117
        Write #3, plot1;                            'p.lot_num              jv060117
        Write #3, pqty1;                            'p.units                jv060117
        Write #3, plot2;                            'p.lot2                 jv060117
        Write #3, pqty2;                            'p.units2               jv060117
        Write #3, "COMP";                           'p.status               jv060117
        Write #3, "WMS";                            'p.userid               jv060117
        Write #3, Format(Now, "yyMMdd hh:mm:ss");   'p.trandate             jv060117
        Write #3, pno                               'p.reqid                jv060117
        Call post_dock_barcode(pbar, daiorderid)
    Else
        s = "select * from pallets where plateno = '" & pno & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            pbar = ds!barcode
            psku = Trim(Left(pbar, 4))                  'jv060117
            Write #3, ds!id                             'p.id               jv060117
            Write #3, "SR-5";                           'p.area             jv060117
            Write #3, daiorderid;                       'p.description      jv060117
            Write #3, "SR-5";                           'p.source           jv060117
            Write #3, "DOCK";                           'p.target           jv060117
            s = psku
            If labpix(Val(psku)).package > " " Then s = s & " " & labpix(Val(psku)).package
            If labpix(Val(psku)).name1 > " " Then s = s & " " & labpix(Val(psku)).name1
            If labpix(Val(psku)).name2 > " " Then s = s & " " & labpix(Val(psku)).name2
            If labpix(Val(psku)).name3 > " " Then s = s & " " & labpix(Val(psku)).name3
            Write #3, s;                                'p.product          jv060117
            Write #3, ds!barcode;                       'p.palletid         jv060117
            Write #3, "1";                              'p.qty              jv060117
            Write #3, "Pallet";                         'p.uom              jv060117
            Write #3, ds!lot1;                          'p.lot_num          jv060117
            Write #3, ds!qty1;                          'p.units            jv060117
            Write #3, ds!lot2;                          'p.lot2             jv060117
            Write #3, ds!qty2;                          'p.units2           jv060117
            Write #3, "COMP";                           'p.status           jv060117
            Write #3, "WMS";                            'p.userid               jv060117
            Write #3, Format(Now, "yyMMdd hh:mm:ss");   'p.trandate             jv060117
            Write #3, pno                               'p.reqid                jv060117
            'Write #3, "SR-5";
            'Write #3, "...";
            'Write #3, Trim(Left(pbar, 4));
            'Write #3, barcode_to_lotnum(pbar);
            'Write #3, Right(pbar, 3);
            'Write #3, pbar;
            'Write #3, pno;
            'Write #3, "AS-1";
            'Write #3, "NoLane";
            'Write #3, Format(Now, "h:mm am/pm")
            Call post_dock_barcode(pbar, daiorderid)
        Else
            pbar = dai_lot_barcode(pitem, plot)         'jv060117
            psku = Trim(Left(pbar, 4))                  'jv060117
            Write #3, "0";                              'p.id               jv060117
            Write #3, "SR-5";                           'p.area             jv060117
            Write #3, "NoPlate in WMS";                 'p.description      jv060117
            Write #3, "SR-5";                           'p.source           jv060117
            Write #3, "DOCK";                           'p.target           jv060117
            s = psku
            If labpix(Val(psku)).package > " " Then s = s & " " & labpix(Val(psku)).package
            If labpix(Val(psku)).name1 > " " Then s = s & " " & labpix(Val(psku)).name1
            If labpix(Val(psku)).name2 > " " Then s = s & " " & labpix(Val(psku)).name2
            If labpix(Val(psku)).name3 > " " Then s = s & " " & labpix(Val(psku)).name3
            Write #3, s;                                'p.product          jv060117
            Write #3, pbar;                             'p.palletid         jv060117
            Write #3, "1";                              'p.qty              jv060117
            Write #3, "Pallet";                         'p.uom              jv060117
            Write #3, plot;                             'p.lot_num          jv060117
            Write #3, "0";                              'p.units            jv060117
            Write #3, " ";                              'p.lot2             jv060117
            Write #3, "0";                              'p.units2           jv060117
            Write #3, "COMP";                           'p.status           jv060117
            Write #3, "WMS";                            'p.userid               jv060117
            Write #3, Format(Now, "yyMMdd hh:mm:ss");   'p.trandate             jv060117
            Write #3, pno                               'p.reqid                jv060117
            'Write #3, "SR-5";
            'Write #3, "...";
            'Write #3, "...";
            'Write #3, ".....";
            'Write #3, "...";
            'Write #3, "Plate not in WMS.";
            'Write #3, pno;
            'Write #3, "AS-1";
            'Write #3, "NoLane";
            'Write #3, Format(Now, "h:mm am/pm")
        End If
        ds.Close
    End If
    Close #3
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "clear_sr_lane", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: clear_sr_lane: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Function Dai_expected_receipt(d As daiexprct) As String
    Dim s As String, cfile As String
    s = "<?xml version=" & Chr(34) & "1.0" & Chr(34)
    s = s & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>" & vbCrLf
    s = s & "<!DOCTYPE ExpectedReceiptMessage SYSTEM " & Chr(34) & "wrxj.dtd" & Chr(34) & ">" & vbCrLf
    s = s & "<ExpectedReceiptMessage>" & vbCrLf
    s = s & "<ExpectedReceipt action=" & Chr(34) & d.action & Chr(34)
    s = s & " sOrderID=" & Chr(34) & Format(Val(d.sOrderID), "000000") & Chr(34) & ">" & vbCrLf
    s = s & "<ExpectedReceiptHeader>" & vbCrLf
    s = s & "<dExpectedDate>" & d.dExpectedDate & "</dExpectedDate>" & vbCrLf
    s = s & "</ExpectedReceiptHeader>" & vbCrLf
    s = s & "<ExpectedReceiptLine sItem=" & Chr(34) & d.sitem & Chr(34) & " sLot=" & Chr(34) & d.slot & Chr(34) & ">" & vbCrLf
    s = s & "<fExpectedQuantity>" & d.fExpectedQuantity & "</fExpectedQuantity>" & vbCrLf
    s = s & "<sStoreDestination>" & d.sStoreDestination & "</sStoreDestination>" & vbCrLf
    s = s & "<sRouteID/>" & vbCrLf
    If d.sHoldReason > " " Then                                                 'jv092314
        s = s & "<sHoldReason>" & d.sHoldReason & "</sHoldReason>" & vbCrLf     'jv092314
    Else                                                                        'jv092314
        s = s & "<sHoldReason/>" & vbCrLf
    End If                                                                      'jv092314
    s = s & "</ExpectedReceiptLine>" & vbCrLf
    s = s & "</ExpectedReceipt>" & vbCrLf
    s = s & "</ExpectedReceiptMessage>"
    cfile = dailogs & "daimessages" & Format(Now, "MMddyy") & ".txt"
    Open cfile For Append As #8
    Print #8, "-------------"
    Print #8, s
    Close #8
    DoEvents
    Dai_expected_receipt = s
End Function

Function dai_lot_barcode(ssku As String, slot As String) As String                          'jv060117
    Dim s As String, syr As String, sdate As String, scode As String, spal As String
    syr = Mid(slot, 1, 2)
    sdate = Mid(slot, 3, 3)
    scode = Mid(slot, 6, 3)
    spal = Mid(slot, 9, 3)
    If Val(syr) > 0 And Val(sdate) > 0 And Val(scode) > 0 And Val(ssku) > 0 Then
        If Len(ssku) = 4 Then
            s = ssku
        Else
            s = ssku & " "
        End If
        sdate = DateAdd("d", Val(sdate), "12-31-20" & Format(Val(syr) - 1, "00"))
        sdate = DateAdd("yyyy", 2, sdate)
        sdate = Format(sdate, "MMddyy")
        dai_lot_barcode = s & sdate & scode & spal
    Else
        dai_lot_barcode = "999 010101999000"
    End If
End Function

Function full_pallet(psku As String, qty As Integer) As Boolean
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    s = "select uom_per_pallet from sku_config where sku = '" & psku & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If qty >= ds(0) Then
            full_pallet = True
        Else
            full_pallet = False
        End If
    Else
        full_pallet = False
    End If
    ds.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "full_pallet", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: full_pallet: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Sub import_sr5_lanes(cflag As String)                               'jv070116
    Dim db As dao.Database
    Dim ds As dao.Recordset, s As String
    Dim ls As dao.Recordset, pid As Long, cfile As String
    Dim sitem As String, slot As String, sposn As String, sload As String
    Dim swarehouse As String, saddress As String, sqty As String, salloc As String
    Dim sorder As String, sordlot As String, sdate As String, sagedate As String
    Dim sholdtype As String, sreason As String, spriority As String, sline As String
    Dim prow As String, pzone As Integer, pside As String, phorz As Integer, pvert As Integer
    Dim palunits(0 To 9999) As Integer, nr As Integer
    Dim bc As String, t As String, plot As String

    nr = 0
    For i = 0 To 9999                                               'jv090115
        palunits(i) = 1
    Next i
    
    If Len(Dir(sr5_lane_data)) = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    Open sr5_lane_data For Input As #1
    Open wms_sr5_data For Output As #2
    Do Until EOF(1)
        Line Input #1, s
        'If Val(Left(s, 4)) > 0 Then
        If Val(Left(s, 6)) > 0 Then                                 'jv062017
            Print #2, s
        End If
    Loop
    Close #1
    Close #2
    
    On Error Resume Next
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, tbbsr)
    If cflag = "Y" Then                 'Clear Flag                 'jv070116
        s = "update position set sku = ' ',lot_num = ' ',pallet_num = ' '"
        s = s & ",count_qty = 0,barcode = ' ',lot2 = ' ',qty2 = 0"
        s = s & " where laneno in (select id from lane where whse_num = 5)"
        db.Execute s
        s = "update lane set qty = 0,sku = ' ',lot_num = ' ', gmasize = 0, lane_status = ' ' where whse_num = 5"
        db.Execute s
    End If                                                          'jv070116
    
    Set ds = db.OpenRecordset("select * from sku_config where sku > '0' and sku <= '9999'") 'jv090115
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            palunits(Val(ds!sku)) = ds!uom_per_pallet
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    Open wms_sr5_data For Input As #1
    Do Until EOF(1)
        Input #1, sload, sitem, slot, saddress, sqty, sholdtype
        If Len(sitem) < 5 Then                                      'jv070116
            'If Len(slot) <> 9 Then slot = "13001_000"
            'If Len(slot) <> 10 Then slot = "1500150000"
            If Len(slot) <> 11 Then slot = "15001500000"                            'jv120415
            If Len(slot) >= 5 Then
                s1 = "12-31-20" & Format(Val(Left(slot, 2)) - 1, "00")
                s2 = Format(DateAdd("d", Val(Mid(slot, 3, 3)), s1), "MM-dd-yyyy")
                sagedate = s2
            End If
            If Len(saddress) = 9 Then
                nr = nr + 1
                prow = Left(saddress, 3)
                If prow = "025" Then
                    pzone = 5
                    pside = "L"
                End If
                If prow = "026" Then
                    pzone = 5
                    pside = "R"
                End If
                If prow = "027" Then
                    pzone = 6
                    pside = "L"
                End If
                If prow = "028" Then
                    pzone = 6
                    pside = "R"
                End If
                If prow = "029" Then
                    pzone = 7
                    pside = "L"
                End If
                If prow = "030" Then
                    pzone = 7
                    pside = "R"
                End If
                If prow = "031" Then
                    pzone = 8
                    pside = "L"
                End If
                If prow = "032" Then
                    pzone = 8
                    pside = "R"
                End If
                If prow = "033" Then
                    pzone = 9
                    pside = "L"
                End If
                If prow = "034" Then
                    pzone = 9
                    pside = "R"
                End If
                phorz = Val(Mid(saddress, 4, 3))
                pvert = Val(Mid(saddress, 7, 3))
                pid = 0
                
                
                bc = sitem                                                              'jv043018
                If Len(bc) = 3 Then bc = bc & " "                                       'jv043018
                plot = Mid(slot, 1, 5)                                                  'jv043018
                If Val(plot) > 0 Then                                                   'jv043018
                    t = "1-1-20" & Left(plot, 2)                                        'jv043018
                    s = Format(DateAdd("d", Val(Right(plot, 3)) - 1, t), "MM-dd-yyyy")  'jv043018
                    s = Format(DateAdd("yyyy", 2, s), "MM-dd-yyyy")                     'jv043018
                    s = Format(s, "MMddyy")                                             'jv043018
                End If                                                                  'jv043018
                bc = bc & s                                                             'jv043018
                bc = bc & Mid(slot, 6, 6)                                               'jv043018
                
                
                's = "select * from pallets where plateno = '" & Trim(sload) & "'"
                's = s & " and sku = '" & Trim(sitem) & "'"
                's = s & " and lot1 = '" & Left(slot, 5) & "'"
                s = "select * from pallets where barcode = '" & bc & "'"                'jv043018
                Set ds = db.OpenRecordset(s)
                If ds.BOF = False Then
                    ds.MoveFirst
                    s = "select * from lane where whse_num = 5"
                    s = s & " and zone_num = " & pzone
                    s = s & " and vert_loc = " & pvert
                    s = s & " and horz_loc = " & phorz
                    s = s & " and rack_side = '" & pside & "'"
                    Set ls = db.OpenRecordset(s)
                    If ls.BOF = False Then
                        ls.MoveFirst
                        pid = ls!id
                        ls.Edit
                        ls!sku = ds!sku
                        ls!lot_num = ds!lot1
                        ls!qty = 1
                        ls!lot_date = sagedate
                        If Val(sholdtype) <> 168 Then ls!lane_status = "H"
                        If palunits(Val(ds!sku)) < (ds!qty1 + ds!qty2) Then ls!gmasize = ds!qty1 + ds!qty2
                        ls.Update
                    End If
                    ls.Close
                    s = "select * from position where laneno = " & pid
                    Set ls = db.OpenRecordset(s)
                    If ls.BOF = False Then
                        ls.MoveFirst
                        ls.Edit
                        ls!sku = ds!sku
                        ls!lot_num = ds!lot1
                        ls!pallet_num = Right(ds!barcode, 3)
                        ls!count_qty = ds!qty1
                        ls!recv_date = sagedate
                        ls!barcode = ds!barcode
                        ls!lot2 = ds!lot2
                        ls!qty2 = ds!qty2
                        ls.Update
                    End If
                    ls.Close
                Else
                    s = "select * from lane where whse_num = 5"
                    s = s & " and zone_num = " & pzone
                    s = s & " and vert_loc = " & pvert
                    s = s & " and horz_loc = " & phorz
                    s = s & " and rack_side = '" & pside & "'"
                    Set ls = db.OpenRecordset(s)
                    If ls.BOF = False Then
                        ls.MoveFirst
                        pid = ls!id
                        ls.Edit
                        ls!sku = sitem
                        ls!lot_num = Left(slot, 5)
                        ls!qty = 1
                        ls!lot_date = sagedate
                        If Val(sholdtype) <> 168 Then ls!lane_status = "H"
                        If palunits(Val(sitem)) < Val(sqty) Then ls!gmasize = Val(sqty)
                        ls.Update
                    End If
                    ls.Close
                    s = "select * from position where laneno = " & pid
                    Set ls = db.OpenRecordset(s)
                    If ls.BOF = False Then
                        ls.MoveFirst
                        ls.Edit
                        ls!sku = sitem
                        ls!lot_num = Left(slot, 5)
                        'ls!pallet_num = Right(slot, 3)
                        ls!pallet_num = Mid(slot, 9, 3)                     'jv090115
                        ls!count_qty = sqty
                        ls!recv_date = sagedate
                        If Len(slot) > 5 Then
                        
                            s2 = Format(DateAdd("yyyy", 2, s2), "MMddyy")

                            If Len(sitem) = 3 Then                          'jv090115
                                'ls!barcode = sitem & " " & s2 & " " & Mid(slot, 6, 1) & " " & Right(slot, 3)
                                ls!barcode = sitem & " " & s2 & Mid(slot, 6, 3) & Mid(slot, 9, 3)   'jv090115
                            Else
                                ls!barcode = sitem & s2 & Mid(slot, 6, 3) & Mid(slot, 9, 3)   'jv090115
                            End If
                        Else
                            'ls!barcode = sitem & " " & Left(slot, 5) & " " & Mid(slot, 6, 1) & " " & Right(slot, 3)
                            If Len(sitem) = 3 Then                          'jv090115
                                ls!barcode = sitem & " " & Left(slot, 5) & " " & Mid(slot, 6, 3) & " " & Mid(slot, 9, 3)    'jv090115
                            Else
                                ls!barcode = sitem & Left(slot, 5) & " " & Mid(slot, 6, 3) & " " & Mid(slot, 9, 3)    'jv090115
                            End If
                        End If
                        ls!lot2 = " " 'ds!lot2
                        ls!qty2 = 0   'ds!qty2
                        ls.Update
                    End If
                    ls.Close
                End If
                ds.Close
            End If
        End If
    Loop
    Close #1

    cfile = "\\BBC-01-PRODTRK\wd\sr5\bin\sr5implog.txt"
    Open cfile For Append As #5
    s = "SR5 Lanes updated: " & Format(Now, "mm-dd-yyyy hh:mm")
    Print #5, s
    
    s = nr & " Daifuku pallets."
    Print #5, s
    
    s = "select count(*) from lane where whse_num = 5 and qty > 0"
    Set ds = db.OpenRecordset(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = ds(0) & " WMS lanes updated."
    Else
        s = "WMS lane count failed."
    End If
    ds.Close
    Print #5, s
    
    s = "select count(*) from position where whse_num = 5 and count_qty > 0"
    Set ds = db.OpenRecordset(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = ds(0) & " WMS positions updated."
    Else
        s = "WMS position count failed."
    End If
    ds.Close
    Print #5, s
    
    Close #5
    db.Close
    Screen.MousePointer = 0
End Sub

Sub import_sr5_lanes_ado(cflag As String)
    Dim ds As adodb.Recordset, s As String
    Dim ls As adodb.Recordset, pid As Long, cfile As String
    Dim sitem As String, slot As String, sposn As String, sload As String
    Dim swarehouse As String, saddress As String, sqty As String, salloc As String
    Dim sorder As String, sordlot As String, sdate As String, sagedate As String
    Dim sholdtype As String, sreason As String, spriority As String, sline As String
    Dim prow As String, pzone As Integer, pside As String, phorz As Integer, pvert As Integer
    Dim palunits(0 To 9999) As Integer, nr As Integer

    nr = 0
    For i = 0 To 9999                                       'jv082415
        palunits(i) = 1
    Next i
    
    If Len(Dir(sr5_lane_data)) = 0 Then Exit Sub
    
    Screen.MousePointer = 11
    Open sr5_lane_data For Input As #1
    Open wms_sr5_data For Output As #2
    Do Until EOF(1)
        Line Input #1, s
        If Val(Left(s, 4)) > 0 Then
            Print #2, s
        End If
    Loop
    Close #1
    Close #2
    
    On Error Resume Next
    If cflag = "Y" Then                 'Clear Flag         'jv070116
        s = "update position set sku = ' ',lot_num = ' ',pallet_num = ' '"
        s = s & ",count_qty = 0,barcode = ' ',lot2 = ' ',qty2 = 0"
        s = s & " where laneno in (select id from lane where whse_num = 5)"
        Wdb.Execute s
        s = "update lane set qty = 0,sku = ' ',lot_num = ' ', gmasize = 0, lane_status = ' ' where whse_num = 5"
        Wdb.Execute s
    End If                                                  'jv070116
    
    Set ds = Wdb.Execute("select sku, uom_per_pallet from sku_config where sku > '0' and sku <= '9999'") 'jv082415
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            palunits(Val(ds!sku)) = ds!uom_per_pallet
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    Open wms_sr5_data For Input As #1
    Do Until EOF(1)
        Input #1, sload, sitem, slot, saddress, sqty, sholdtype
        If Len(sitem) < 5 Then                              'jv070116
            'If Len(slot) <> 9 Then slot = "13001_000"
            If Len(slot) <> 11 Then slot = "13001___000"                            'jv120415
            If Len(slot) >= 5 Then
                s1 = "12-31-20" & Format(Val(Left(slot, 2)) - 1, "00")
                s2 = Format(DateAdd("d", Val(Mid(slot, 3, 3)), s1), "MM-dd-yyyy")
                sagedate = s2
            End If
            If Len(saddress) = 9 Then
                nr = nr + 1
                prow = Left(saddress, 3)
                If prow = "025" Then
                    pzone = 5
                    pside = "L"
                End If
                If prow = "026" Then
                    pzone = 5
                    pside = "R"
                End If
                If prow = "027" Then
                    pzone = 6
                    pside = "L"
                End If
                If prow = "028" Then
                    pzone = 6
                    pside = "R"
                End If
                If prow = "029" Then
                    pzone = 7
                    pside = "L"
                End If
                If prow = "030" Then
                    pzone = 7
                    pside = "R"
                End If
                If prow = "031" Then
                    pzone = 8
                    pside = "L"
                End If
                If prow = "032" Then
                    pzone = 8
                    pside = "R"
                End If
                If prow = "033" Then
                    pzone = 9
                    pside = "L"
                End If
                If prow = "034" Then
                    pzone = 9
                    pside = "R"
                End If
                phorz = Val(Mid(saddress, 4, 3))
                pvert = Val(Mid(saddress, 7, 3))
                pid = 0
                s = "select * from pallets where plateno = '" & Trim(sload) & "'"
                s = s & " and sku = '" & Trim(sitem) & "'"
                s = s & " and lot1 = '" & Left(slot, 5) & "'"
                Set ds = Wdb.Execute(s)
                If ds.BOF = False Then
                    ds.MoveFirst
                    s = "select * from lane where whse_num = 5"
                    s = s & " and zone_num = " & pzone
                    s = s & " and vert_loc = " & pvert
                    s = s & " and horz_loc = " & phorz
                    s = s & " and rack_side = '" & pside & "'"
                    Set ls = Wdb.Execute(s)
                    If ls.BOF = False Then
                        ls.MoveFirst
                        pid = ls!id
                        s = "Update lane set sku = '" & ds!sku & "'"
                        s = s & ", lot_num = '" & ds!lot1 & "'"
                        s = s & ", qty = 1"
                        s = s & ", lot_date = '" & sagedate & "'"
                        If Val(sholdtype) <> 168 Then s = s & ", lane_status = 'H'"
                        If palunits(Val(ds!sku)) < (ds!qty1 + ds!qty2) Then s = s & ", gmasize = " & ds!qty1 + ds!qty2
                        s = s & ") Where id = " & ls!id
                        Wdb.Execute s
                    End If
                    ls.Close
                    s = "Update position set sku = '" & ds!sku & "'"
                    s = s & ", lot_num = '" & ds!lot1 & "'"
                    s = s & ", pallet_num = '" & Right(ds!barcode, 3) & "'"
                    s = s & ", count_qty = " & ds!qty1
                    s = s & ", recv_date = '" & sagedate & "'"
                    s = s & ", barcode = '" & ds!barcode & "'"
                    s = s & ", lot2 = '" & ds!lot2 & "'"
                    s = s & ", qty2 = " & ds!qty2
                    s = s & ") Where laneno = " & pid
                    Wdb.Execute s
                Else
                    s = "select * from lane where whse_num = 5"
                    s = s & " and zone_num = " & pzone
                    s = s & " and vert_loc = " & pvert
                    s = s & " and horz_loc = " & phorz
                    s = s & " and rack_side = '" & pside & "'"
                    Set ls = Wdb.Execute(s)
                    If ls.BOF = False Then
                        ls.MoveFirst
                        pid = ls!id
                        s = "Update lane set sku = '" & sitem & "'"
                        s = s & ", lot_num = '" & Left(slot, 5) & "'"
                        s = s & ", qty = 1"
                        s = s & ", lot_date = '" & sagedate & "'"
                        If Val(sholdtype) <> 168 Then s = s & ", lane_status = 'H'"
                        If palunits(Val(sitem)) < (ds!qty1 + ds!qty2) Then s = s & ", gmasize = " & Val(sqty)
                        s = s & ") Where id = " & ls!id
                        Wdb.Execute s
                    End If
                    ls.Close
                    s = "Update position set sku = '" & sitem & "'"
                    s = s & ", lot_num = '" & Left(slot, 5) & "'"
                    s = s & ", pallet_num = '" & Right(slot, 3) & "'"
                    s = s & ", count_qty = " & Val(sqty)
                    s = s & ", recv_date = '" & sagedate & "'"
                    If Len(slot) > 5 Then
                        s2 = Format(DateAdd("yyyy", 2, s2), "MMddyy")
                        s = s & ",barcode = '" & sitem & " " & s2 & " " & Mid(slot, 6, 1) & " " & Right(slot, 3) & "'"
                    Else
                        s = s & ",barcode = '" & sitem & " " & Left(slot, 5) & " " & Mid(slot, 6, 1) & " " & Right(slot, 3) & "'"
                    End If
                    s = s & ", lot2 = ' '"
                    s = s & ", qty2 = 0"
                    s = s & ") Where laneno = " & pid
                    Wdb.Execute s
                End If
                ds.Close
            End If
        End If
    Loop
    Close #1

    cfile = "\\BBC-01-PRODTRK\wd\sr5\bin\sr5implog.txt"
    Open cfile For Append As #5
    s = "SR5 Lanes updated: " & Format(Now, "mm-dd-yyyy hh:mm")
    Print #5, s
    
    s = nr & " Daifuku pallets."
    Print #5, s
    
    s = "select count(*) from lane where whse_num = 5 and qty > 0"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = ds(0) & " WMS lanes updated."
    Else
        s = "WMS lane count failed."
    End If
    ds.Close
    Print #5, s
    
    s = "select count(*) from position where whse_num = 5 and count_qty > 0"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = ds(0) & " WMS positions updated."
    Else
        s = "WMS position count failed."
    End If
    ds.Close
    Print #5, s
    
    Close #5
    Screen.MousePointer = 0
End Sub

Sub insert_trans(pt As ptask)
    Dim ds As adodb.Recordset, s As String, rid As Long
    On Error GoTo vberror
    s = "select * from paltasks where id = " & new_pallet_task_record(pt.area)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set area = '" & pt.area & "'"
        s = s & ", description = '" & pt.description & "'"
        s = s & ", source = '" & pt.source & "'"
        s = s & ", target = '" & pt.target & "'"
        s = s & ", product = '" & pt.product & "'"
        s = s & ", palletid = '" & pt.palletid & "'"
        s = s & ", qty = " & CLng(Val(pt.qty))
        s = s & ", uom = '" & pt.uom & "'"
        s = s & ", lotnum = '" & pt.lotnum & "'"
        s = s & ", units = " & CLng(Val(pt.units))
        s = s & ", lotnum2 = '" & pt.lotnum2 & "'"
        s = s & ", units2 = " & CLng(Val(pt.units2))
        s = s & ", status = '" & pt.status & "'"
        s = s & ", userid = '" & pt.userid & "'"
        s = s & ", trandate = '" & pt.trandate & "'"
        s = s & ", reqid = '" & pt.reqid & "'"
        s = s & " Where id = " & ds!id
        Wdb.Execute s
    Else
        rid = wd_seq("PalTasks")
        s = "INSERT INTO PalTasks (ID, Area, Description, Source, Target, Product,"
        s = s & " PalletID, Qty, Uom, LotNum, Units, LotNum2, Units2, Status, UserID,"
        s = s & " TranDate, ReqID) VALUES (" & rid & ","
        s = s & "'" & pt.area & "',"
        s = s & "'" & pt.description & "',"
        s = s & "'" & pt.source & "',"
        s = s & "'" & pt.target & "',"
        s = s & "'" & pt.product & "',"
        s = s & "'" & pt.palletid & "',"
        s = s & CLng(Val(pt.qty)) & ","
        s = s & "'" & pt.uom & "',"
        s = s & "'" & pt.lotnum & "',"
        s = s & CLng(Val(pt.units)) & ","
        s = s & "'" & pt.lotnum2 & "',"
        s = s & CLng(Val(pt.units2)) & ","
        s = s & "'" & pt.status & "',"
        s = s & "'" & pt.userid & "',"
        s = s & "'" & pt.trandate & "',"
        s = s & "'" & pt.reqid & "')"
        Wdb.Execute s
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "insert_trans", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: insert_trans: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Sub LoadDocument(docfile As String, mtype As String)
    Dim xdoc As MSXML2.DOMDocument60
    Set xdoc = New MSXML2.DOMDocument60
    xdoc.validateOnParse = False
    If xdoc.Load(docfile) Then
    ' The document loaded successfully.
    ' Now do something intersting.
        daimesstext = xdoc.documentElement.nodeName & vbCrLf
        DisplayNode xdoc.childNodes, 0, mtype
    Else
        'MsgBox " The document failed to load."
        ' See the previous listing for error information.
    End If
    If mtype = "LocationArrivalMessage" Then Call update_sr_lane(daiplate, daibay)
    If mtype = "PickCompleteMessage" Then
        Call clear_sr_lane(daiplate, daiitem, dailotnum)
        DoEvents
        daiitem = " "
        dailotnum = " "
    End If
End Sub

Public Sub DisplayNode(ByRef Nodes As IXMLDOMNodeList, ByVal Indent As Integer, mtype As String)

    Dim xNode As IXMLDOMNode
    Dim xattr As IXMLDOMAttribute
    Indent = Indent + 2
    For Each xNode In Nodes
        If xNode.nodeType = NODE_TEXT Then
            daimesstext = daimesstext & Space$(Indent) & xNode.parentNode.nodeName & _
            ":" & xNode.nodeValue & vbCrLf '" type: " & xNode.nodeType & vbCrLf
            If mtype = "LocationArrivalMessage" Then
                If xNode.parentNode.nodeName = "sLoadID" Then daiplate = xNode.nodeValue
                If xNode.parentNode.nodeName = "sLocation" Then daibay = xNode.nodeValue
            End If
            If mtype = "PickCompleteMessage" Then
                If xNode.parentNode.nodeName = "sLoadID" Then daiplate = xNode.nodeValue
                If xNode.parentNode.nodeName = "sOrderID" Then daiorderid = xNode.nodeValue
                If xNode.parentNode.nodeName = "sItem" Then daiitem = xNode.nodeValue
                If xNode.parentNode.nodeName = "sLot" Then dailotnum = xNode.nodeValue
            End If
        Else
            'daimesstext = daimesstext & xNode.parentNode.nodeName & " type: " & xNode.nodeType & vbCrLf
            'daimesstext = daimesstext & xNode.nodeName & " type: " & xNode.nodeType & vbCrLf
        End If
      
        If xNode.nodeType = 1 Then
            If xNode.Attributes.length > 0 Then
                For i = 0 To xNode.Attributes.length - 1
                    'daimesstext.Text = daimesstext.Text & "attr(" & i & "): " & xNode.Attributes(i).nodeName
                    daimesstext = daimesstext & xNode.nodeName & " " & xNode.Attributes(i).nodeName & _
                    "=" & xNode.Attributes(i).nodeValue & vbCrLf
                Next i
            'Else
            '    daimesstext = daimesstext & xNode.nodeName & vbCrLf
            End If
        End If
        
      If xNode.hasChildNodes Then
         DisplayNode xNode.childNodes, Indent, mtype
      End If
   Next xNode
   
End Sub

Public Sub load_labpics()
    Dim s As String, ts As Integer, te As Integer, psku As String
    Open labfmtfile For Input As #1
    Do Until EOF(1)
        Line Input #1, s
        te = InStr(1, s, Chr(9))
        psku = Mid(s, 1, te - 1)
        labpix(Val(psku)).sku = psku
        ts = te + 1
        te = InStr(ts, s, Chr(9))
        labpix(Val(psku)).package = Mid(s, ts, te - ts)
        ts = te + 1
        te = InStr(ts, s, Chr(9))
        labpix(Val(psku)).name1 = Mid(s, ts, te - ts)
        ts = te + 1
        te = InStr(ts, s, Chr(9))
        labpix(Val(psku)).name2 = Mid(s, ts, te - ts)
        ts = te + 1
        If ts < Len(s) Then
            labpix(Val(psku)).name3 = Mid(s, ts, Len(s) - (ts - 1))
        Else
            labpix(Val(psku)).name3 = " "
        End If
    Loop
    Close #1
End Sub

Function mastrec(taskid As Long) As ptask
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    s = "select * from paltasks where id = " & taskid
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        mastrec.id = ds!id
        mastrec.area = ds!area
        mastrec.description = ds!description
        mastrec.source = ds!source
        mastrec.target = ds!target
        mastrec.product = ds!product
        mastrec.palletid = ds!palletid
        mastrec.qty = ds!qty
        mastrec.uom = ds!uom
        mastrec.lotnum = ds!lotnum
        mastrec.units = ds!units
        mastrec.lotnum2 = ds!lotnum2
        mastrec.units2 = ds!units2
        mastrec.status = ds!status
        mastrec.userid = ds!userid
        mastrec.trandate = ds!trandate
        mastrec.reqid = ds!reqid
    Else
        mastrec.id = 0
        mastrec.area = " "
        mastrec.description = " "
        mastrec.source = " "
        mastrec.target = " "
        mastrec.product = " "
        mastrec.palletid = " "
        mastrec.qty = " "
        mastrec.uom = " "
        mastrec.lotnum = " "
        mastrec.units = " "
        mastrec.lotnum2 = " "
        mastrec.units2 = " "
        mastrec.status = " "
        mastrec.userid = " "
        mastrec.trandate = " "
        mastrec.reqid = " "
    End If
    ds.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "mastrec", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: mastrec: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Function new_pallet_queue(flag1 As Boolean) As Long
    Dim ds As adodb.Recordset, s As String, k As Long, zid As Long
    On Error GoTo vberror
    If flag1 = True Then            '1st Queue in the list
        s = "select queue_num from queue_infc where queue_num > 0"
        s = s & " order by queue_num"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            k = ds!queue_num - 1
        Else
            k = 100
        End If
        ds.Close
    Else
        s = "select max(queue_num) from queue_infc"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            k = ds(0) + 1
        Else
            k = 100
        End If
        ds.Close
    End If
    lastque = k
    s = "select id, queue_num from queue_infc where queue_num = 0"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "update queue_infc set queue_num = " & k
        s = s & " where id = " & ds!id
        Wdb.Execute s
        new_pallet_queue = ds!id
    Else
        zid = wd_seq("Queue_Infc")
        s = "INSERT INTO Queue_Infc (ID, Queue_num) VALUES ("
        s = s & zid & ", " & k & ")"
        Wdb.Execute s
        new_pallet_queue = zid
    End If
    ds.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "new_pallet_queue", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: new_pallet_queue: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Function new_pallet_task_record(parea As String) As Long
    Dim ds As adodb.Recordset, s As String, zid As Long
    On Error GoTo vberror
    s = "select id, status from paltasks where area = '" & parea & "'"
    s = s & " and status = 'COMP'"
    If parea <> "TRAFFIC MASTER" Then s = s & " and id >= 400"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set status = 'PEND' Where id = " & ds!id
        Wdb.Execute s
        new_pallet_task_record = ds!id
    Else
        ds.Close
        s = "select id, status from paltasks where status = 'COMP'"
        If parea <> "TRAFFIC MASTER" Then s = s & " and id >= 400"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            s = "Update paltasks set status = 'PEND' Where id = " & ds!id
            Wdb.Execute s
            new_pallet_task_record = ds!id
        Else
            zid = wd_seq("PalTasks")
            s = "INSERT INTO PalTasks (ID) VALUES (" & zid & ")"
            Wdb.Execute s
            new_pallet_task_record = zid
        End If
    End If
    ds.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "new_pallet_task_record", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: new_pallet_task_record: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Function part_pallet_whs(psku As String) As String
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    's = "select sku from opbays where sku = '" & psku & "'"
    'Set ds = Wdb.Execute(s)
    'If ds.BOF = False Then
    '    part_pallet_whs = "1"
    'Else
        part_pallet_whs = "4"                                   'jv022216
    'End If
    'ds.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "part_pallet_whs", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: part_pallet_whs: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Sub poll_logs()
    Dim sdate As String, t As String, ft As String, s As String
    Do While True
        If form1.scanlogs.Value = 0 Then Exit Do
        sdate = Format(Now, "MMddyyyy")
        'form1.logfile1 = "v:\pallogs\recv" & sdate & ".txt"
        form1.logfile1 = pallogs & "recv" & sdate & ".txt"                      'jv092316
        If Len(Dir(form1.logfile1)) > 0 Then
            form1.logsize1 = FileLen(form1.logfile1)
            DoEvents
        End If
        'form1.logfile2 = "v:\pallogs\move" & sdate & ".txt"
        form1.logfile2 = pallogs & "move" & sdate & ".txt"                      'jv092316
        If Len(Dir(form1.logfile2)) > 0 Then
            form1.logsize2 = FileLen(form1.logfile2)
            DoEvents
        End If
        t = Format(Now, "hh:mm:ss")                                     'jv100714
        If Right(t, 1) = "0" Or Right(t, 1) = "5" Then                  'jv100714
            form1.timelog = Format(Now, "h:mm:ss am/pm")                'jv100714
        End If                                                          'jv100714
        't = Format(Now, "hh:mm")
        'form1.timelog = Format(Now, "h:mm am/pm")
        
        'Download Lane Data @ 1:00 am
        If t >= "01:00" And t < "02:30" Then
            If Len(Dir(sr5_lane_data)) > 0 Then
                ft = Format(FileDateTime(sr5_lane_data), "M-d-yyyy h:mm am/pm")
                If ft <> form1.slanedate Then
                    Call import_sr5_lanes("Y")                          'jv070116
                    DoEvents
                    form1.slanedate = ft
                End If
            End If
        End If
        
        'If Len(Dir(form1.shipordfile)) > 0 Then
        '    If FileLen(form1.shipordfile) > 100 Then                    'jv012914
        '        form1.shipordtime = Format(FileDateTime(form1.shipordfile), "MM-dd-yyyy hh:mm:ss am/pm")
        '        DoEvents
        '    End If
        'End If
        If Len(Dir(form1.cobrcptfile)) > 0 Then                         'jv100714
            If FileLen(form1.cobrcptfile) > 100 Then                    'jv100714
                form1.cobrcpttime = Format(FileDateTime(form1.cobrcptfile), "MM-dd-yyyy hh:mm:ss am/pm")
                DoEvents                                                'jv100714
            End If                                                      'jv100714
        End If                                                          'jv100714
        
        'If Len(Dir(form1.additemholdfile.Caption)) > 0 Then                     'jv042015
        '    If FileLen(form1.additemholdfile.Caption) > 100 Then                'jv042015
        '        form1.additemholdtime.Caption = Format(FileDateTime(form1.additemholdfile.Caption), "MM-dd-yyyy hh:mm:ss am/pm")
        '        DoEvents                                                'jv042015
        '    End If                                                      'jv042015
        'End If                                                          'jv042015
        
        'If Len(Dir(form1.remitemholdfile.Caption)) > 0 Then                     'jv042015
        '    If FileLen(form1.remitemholdfile.Caption) > 100 Then                'jv042015
        '        form1.remitemholdtime.Caption = Format(FileDateTime(form1.remitemholdfile.Caption), "MM-dd-yyyy hh:mm:ss am/pm")
        '        DoEvents                                                'jv042015
        '    End If                                                      'jv042015
        'End If                                                          'jv042015
        
        Call poll_sql_requests                      'jv042915
        
    Loop
End Sub

Sub poll_queue_tasks()
    Dim ds As adodb.Recordset, s As String, p As ptask
    Dim pno As Long, splate As String
    'On Error GoTo vberror
    form1.pqflag.Value = 1: DoEvents                                    'jv070214
    s = "select * from queue_infc where queue_num > 0 and source in ('FG3', 'FG5', 'FG6')"  'jvsp27
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            splate = check_dai_plate(ds!palletid)
            If splate = "None" Then
                pno = wd_seq("BHBarcode")
                splate = Format(Val(pno), "000000")         'jvsp27
                p.id = 0
                p.area = "DOCK"
                p.description = " "
                p.source = "BACKHAUL"
                p.target = "SR" & ds!whse_num
                p.product = ds!sku
                p.palletid = ds!palletid
                p.qty = "1"
                p.uom = "Pallet"
                p.lotnum = ds!lot_num
                p.units = ds!units
                p.lotnum2 = ds!lot_num2
                p.units2 = ds!units2
                p.status = "COMP"
                p.userid = "WMSDAI"
                p.trandate = Format(Now, "yyMMdd hh:mm:ss")
                p.reqid = splate 'pno
                form1.wbh = splate 'pno
                'Call record_pallet(Str(pno), p, Str(ds!whse_num), "Backhaul")
                'Call record_pallet(Str(pno), p, Str(ds!whse_num), "BHtest")
                Call record_pallet(splate, p, Str(ds!whse_num), "BHtest")
            End If
            p.id = 0
            p.area = "DOCK"
            p.description = " "
            p.source = "BACKHAUL"
            p.target = "SR" & ds!whse_num
            p.product = ds!sku
            p.palletid = ds!palletid
            p.qty = "1"
            p.uom = "Pallet"
            p.lotnum = ds!lot_num
            p.units = ds!units
            p.lotnum2 = ds!lot_num2
            p.units2 = ds!units2
            p.status = "COMP"
            p.userid = "WMSDAI"
            p.trandate = Format(Now, "yyMMdd hh:mm:ss")
            p.reqid = Format(Val(splate), "000000")             'jvsp27
            Call process_fg_pallet(p)
            ds.MoveNext
        Loop
    End If
    ds.Close
    form1.pqflag.Value = 0: DoEvents                                    'jv070214
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "poll_queue_tasks", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: poll_queue_tasks: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Sub poll_sql_requests()                                                     'jv042915
    Dim s As String, ds As adodb.Recordset, i As Integer, rkey As Long
    If form1.Grid2.Rows > 1 Then
        'form1.Grid2.Redraw = False
        'form1.Grid2.Clear: form1.Grid2.Rows = 1: form1.Grid2.Cols = 6
        form1.Grid2.Rows = 1
    End If
    s = "select * from BBC_HostToWrx where bbcstatus = 'PEND'"
    s = s & " Order by imessagesequence, dhostmodifytime"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = Format(ds(0), "M-d-yy h:mm:ss am/pm") & Chr(9) & ds(1) & Chr(9) & ds(2) & Chr(9)
            s = s & ds(3) & Chr(9) & ds(4) & Chr(9) & ds(5)
            form1.Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    'form1.Grid2.FormatString = "<Time|^ID|<MessageType|<Message|<BBC ID|^Status"
    'form1.Grid2.ColWidth(0) = 1800
    'form1.Grid2.ColWidth(1) = 800
    'form1.Grid2.ColWidth(2) = 1800
    'form1.Grid2.ColWidth(3) = 3000
    'form1.Grid2.ColWidth(4) = 1600
    'form1.Grid2.ColWidth(5) = 800
    'form1.Grid2.Redraw = True
    If form1.Grid2.Rows > 1 Then
        'form1.Grid2.Redraw = True
        For i = 1 To form1.Grid2.Rows - 1
            s = form1.Grid2.TextMatrix(i, 2) & ">"
            If Right(form1.Grid2.TextMatrix(i, 3), Len(s)) = s Then
                rkey = wd_seq("DAIRequests")
                's = "Insert Into HostToWrx (iMessageSequence, sMessageIdentifier, sMessage)"
                's = s & " Values (" & rkey  'form1.Grid2.TextMatrix(i, 1)
                's = s & ", '" & form1.Grid2.TextMatrix(i, 2) & "'"
                's = s & ", '" & form1.Grid2.TextMatrix(i, 3) & "')"
                ''MsgBox s
                'DaiDb.Execute s
                
                s = "Insert Into HostData.HostToWrx (iMessageSequence, sMessageIdentifier, sMessage)"
                s = s & " Values (" & rkey
                s = s & ", '" & form1.Grid2.TextMatrix(i, 2) & "'"
                s = s & ", '" & form1.Grid2.TextMatrix(i, 3) & "')"
                DaiDb.Execute s
                
                s = "Update BBC_HostToWrx set bbcstatus = 'COMP' where imessagesequence = "
                s = s & form1.Grid2.TextMatrix(i, 1)
                'MsgBox s
                Wdb.Execute s
            Else
                s = "Update BBC_HostToWrx set bbcstatus = 'JUNK' where imessagesequence = "
                s = s & form1.Grid2.TextMatrix(i, 1)
                'MsgBox s
                Wdb.Execute s
            End If
        Next i
    End If
End Sub

Sub poll_wrapper_tasks()
    Dim ds As adodb.Recordset, s As String, p As ptask
    Dim psku As String, pqty As Integer
    Dim plot As String, pcode As String, plot2 As String, pcode2 As String  'jv062614
    On Error GoTo vberror
    form1.pwflag.Value = 1: DoEvents                                    'jv070214
    s = "select id,palletid from paltasks where units > 0"
    s = s & " and palletid > '100 000000 0 000'"
    s = s & " and area = 'TRAFFIC MASTER'"
    s = s & " and status = 'PEND'"                                    'jv020514
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If check_dai_plate(ds!palletid) = "None" Then
                p = mastrec(ds!id)
                form1.wrapbc = p.palletid
                psku = Trim(Left(p.palletid, 4))
                plot = p.lotnum                                         'jv062614
                'pcode = Mid(p.palletid, 12, 1)                          'jv062614
                pcode = Trim(Mid(p.palletid, 11, 3))                'jv052515
                If Len(p.lotnum2) = 7 Then                              'jv062614
                    plot2 = Mid(p.lotnum2, 1, 5)                        'jv062614
                    pcode2 = Mid(p.lotnum2, 7, 1)                       'jv062614
                Else                                                    'jv062614
                    If Len(p.lotnum2) = 5 Then                          'jv062614
                        plot2 = p.lotnum2                               'jv062614
                        pcode2 = pcode                                  'jv062614
                    Else                                                'jv062614
                        plot2 = " "                                     'jv062614
                        pcode2 = " "                                    'jv062614
                    End If                                              'jv062614
                End If                                                  'jv062614
                pqty = Val(p.units) + Val(p.units2)
                If p.target = "SR4" Or p.target = "SR5" Then                                    'jv092314
                    'Call form1.refresh_grid1(psku, "HOLD", pcode, plot2, pcode2, pqty)         'jv092314
                    form1.Grid1.Rows = 1                                                        'jv092314
                    form1.Grid1.AddItem Mid(p.target, 3, 1)                                     'jv092314
                Else                                                                            'jv092314
                    Call form1.refresh_grid1(psku, plot, pcode, plot2, pcode2, pqty)    'jv062614
                End If                                                                          'jv092314
                DoEvents
                If form1.Grid1.Rows > 1 Then Call process_wrapper_pallet(p)
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    form1.pwflag.Value = 0: DoEvents                                    'jv070214
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "poll_wrapper_tasks", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: poll_wrapper_tasks: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Sub post_dai_exp_rcpt(p As ptask, pwhs As String, pno As String)
    Dim d As daiexprct, rkey As Long
    Dim q As Integer, psku As String
    Dim s As String                 'jv090913
    If p.source = "TRI-LEVEL 1" Then                'jv091313
        If pwhs = "1" Then Exit Sub                 'jv091313
        If pwhs = "2" Then Exit Sub                 'jv091313
        If pwhs = "4" Then Exit Sub                 'jv091313
    End If                                          'jv091313
    If p.source = "TRI-LEVEL 2" Then                'jv091313
        If pwhs = "1" Then Exit Sub                 'jv091313
        If pwhs = "2" Then Exit Sub                 'jv091313
        If pwhs = "4" Then Exit Sub                 'jv091313
    End If                                          'jv091313
    
    d.sHoldReason = " "                                                             'jv092314
    If check_hold(p) = True Then d.sHoldReason = "PC"           'jv040615
    q = Val(p.units) + Val(p.units2)                            'jv090513
    psku = Trim(Left(p.palletid, 4))                            'jv090513
    If full_pallet(psku, q) = False Then                        'jv090513
        If p.source = "TRI-LEVEL 3" Or p.source = "TRI-LEVEL 4" Then                'jv092314
            If check_hold(p) = True Then                        'jv041015
                pwhs = 4                                        'jv041015
            Else
                pwhs = part_pallet_whs(psku)                                        'jv092314
            End If
            d.sHoldReason = "PC"                                                    'jv092314
        Else                                                                        'jv092314
            pwhs = "2207"                                       'jv090513
            p.source = "SR5"                                    'jv090513
            p.target = "STAGING"                                'jv090513
            p.status = "PEND"                                   'jv090513
            p.userid = ""                                       'jv090513
            p.reqid = pno                                       'jv090513
            Call insert_trans(p)                                'jv090513
            s = "Update queue_infc Set queue_num = 0"           'jv090913
            s = s & " Where whse_num in (5, 6)"                 'jv090913
            s = s & " and palletid = '" & p.palletid & "'"      'jv090913
            Wdb.Execute s                                       'jv090913
        End If                                                                      'jv092314
    End If                                                      'jv090513
    d.action = "ADD"
    d.sOrderID = pno
    d.dExpectedDate = Format(Now, "MM/dd/yyyy hh:mm:ss")
    d.sitem = Trim(Left(p.palletid, 4))
    'd.slot = Trim(p.lotnum & Mid(p.palletid, 12, 1) & Mid(p.palletid, 14, 3))
    d.slot = Trim(p.lotnum & Trim(Mid(p.palletid, 11, 3)) & Mid(p.palletid, 14, 3))     'jv052515
    d.fExpectedQuantity = Val(p.units) + Val(p.units2)
    d.sStoreDestination = pwhs
    If pwhs = "5" And p.description = "HOLD" Then                                   'jv092314
        d.sHoldReason = "PC"                                                        'jv092314
    End If                                                                          'jv092314
    'If d.sStoreDestination = "2" Or d.sStoreDestination = "3" Or d.sStoreDestination = "5" Or d.sStoreDestination = "6" Or d.sStoreDestination = "2207" Then
        Open dailogs & "daiExpectedReceiptMessage.xml" For Output As #1
        Print #1, Dai_expected_receipt(d)
        Close #1
        DoEvents
        rkey = wd_seq("DAIRequests")
        
        Call write_oracle_request("ExpectedReceiptMessage", rkey)
        form1.WebBrowser1.Navigate2 dailogs & "daiExpectedReceiptMessage.xml"
    'End If             'jv091313
End Sub

Sub post_dock_barcode(pbar As String, gcode As String)
    Dim ds As adodb.Recordset, s As String
    Dim psku As String, plot As String, pbr As Integer, gc As String
    On Error GoTo vberror
    psku = Trim(Left(pbar, 4))
    plot = barcode_to_lotnum(pbar)
    pbr = Val(Right(gcode, 2))
    gc = Left(gcode, Len(gcode) - 3)
    s = "select id from paltasks where area = 'DOCK'"
    s = s & " and description = '" & gc & "'"
    s = s & " and source = 'SR5'"
    s = s & " and product like '" & psku & "%'"
    s = s & " and lotnum < '0'"
    If pbr <> 15 And pbr <> 16 Then
        If Right(gc, 2) = "ZO" Then
            s = s & " and target like '%" & gc & "%'"
        Else
            s = s & " and target like '" & bbcbranches(pbr) & "%'"
        End If
    End If
    s = s & " and status = 'PEND' and userid < '0'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set palletid = '" & pbar & "',"
        s = s & "lotnum = '" & plot & "' where id = " & ds(0)
        Wdb.Execute s
    End If
    ds.Close
    s = "Update queue_infc Set queue_num = 0 Where whse_num in (2, 3, 5, 6)"  'jv011714
    s = s & " and palletid = '" & pbar & "'"
    Wdb.Execute s
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "post_dock_barcode", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: post_dock_barcode: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Sub post_queue_to_sr(whs As Integer, p As ptask)
    Dim ds As adodb.Recordset, s As String
    Dim qid As Long, hf As Boolean, nque As Integer
    On Error GoTo vberror
    hf = False
    hf = check_hold(p)                              'jv040615
    'Process Queue
    s = "select id from queue_infc where palletid = '" & p.palletid & "'"
    s = s & " and whse_num = " & whs
    s = s & " and queue_num > 0"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then      'palletid already found in queues
        ds.MoveFirst
    Else
        ds.Close
        s = "select * from queue_infc where id = " & new_pallet_queue(False)
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            s = "update queue_infc set whse_num = " & whs
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
            s = s & ",source = 'TML'"
            s = s & " where id = " & ds!id
            Wdb.Execute s
        Else
            nque = 50
            qid = wd_seq("Queue_Infc")
            s = "INSERT INTO Queue_Infc (ID, Whse_Num, SKU, Lot_Num, Drop_Flag, Queue_Num,"
            s = s & " Rack_Num, Units, Lot_Num2, Units2, PalletID, Source)"
            s = s & " VALUES (" & qid & ","
            s = s & whs & ","
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
            s = s & "'TML')"
            Wdb.Execute s
        End If
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "post_queue_to_sr", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: post_queue_to_sr: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Sub post_route_to_kep(wrapper As Integer, whs As Integer, palid As Long)
    Dim ds As adodb.Recordset, s As String
    Dim croute As Integer
    On Error GoTo vberror
    croute = whs
    s = "update wrapper_config set pallet_id = " & palid
    s = s & ", sr_destination = " & whs
    s = s & ", conv_route = " & croute
    s = s & " where wrapper_id = " & wrapper
    Wdb.Execute s
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "post_route_to_kep", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: post_route_to_kep: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Sub post_sr4_efl(p As ptask)
    p.area = "FORKLIFT"
    p.source = "TRI LEVEL"
    p.target = "RACKS"
    p.status = "PEND"
    p.userid = ""
    Call insert_trans(p)
End Sub

Sub post_tm_log(pwhs As Integer, p As ptask, pno As String)
    Dim cfile As String
    'cfile = "v:\pallogs\tml" & Format(Now, "MMddyyyy") & ".txt"
    cfile = pallogs & "tml" & Format(Now, "MMddyyyy") & ".txt"              'jv092316
    Open cfile For Append As #1
    Write #1, p.id;
    Write #1, "TRAFFIC MASTER";
    Write #1, " ";
    Write #1, p.source;
    Write #1, "SR" & pwhs;
    Write #1, p.product;
    Write #1, p.palletid;
    Write #1, p.qty;
    Write #1, p.uom;
    Write #1, p.lotnum;
    Write #1, p.units;
    Write #1, p.lotnum2;
    Write #1, p.units2;
    Write #1, p.status;
    Write #1, p.userid;
    Write #1, Format(Now, "yyMMdd hh:mm:ss");
    Write #1, pno
    Close #1
End Sub

Sub process_fg_pallet(p As ptask)
    Dim ds As adodb.Recordset, s As String
    'On Error GoTo vberror
    s = "select * from pallets where barcode = '" & p.palletid & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If ds!status <> "Queue" Then
            If p.target = "CRANE 3" Or p.target = "SR3" Then Call post_dai_exp_rcpt(p, "3", p.reqid)
            If p.target = "CRANE 5" Or p.target = "SR5" Then Call post_dai_exp_rcpt(p, "5", p.reqid)
            If p.target = "CRANE 6" Or p.target = "SR6" Then Call post_dai_exp_rcpt(p, "6", p.reqid)
            s = "Update pallets set status = 'Queue'"
            s = s & ",target = '" & p.target & "'"
            s = s & " where barcode = '" & p.palletid & "'"
            Wdb.Execute (s)
            form1.queuebc = p.palletid
        End If
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "process_fg_pallet", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: process_fg_pallet: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Sub process_tmaster_tasks(p As ptask)
    Dim ds As adodb.Recordset, s As String, sr5cnt As Integer, sr4cnt As Integer
    On Error GoTo vberror
    s = "select id from paltasks where area = 'TRAFFIC MASTER' and id = " & p.id
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set target = '" & p.target & "' Where id = " & ds!id
        Wdb.Execute s
    End If
    ds.Close
    
    s = "select id from paltasks where area = 'TRAFFIC MASTER' and status = 'PEND'"
    s = s & " and target > 'SR'"
    s = s & " and palletid in (select palletid from queue_infc where queue_num = 0)"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update paltasks set status = 'COMP', userid = 'TMaster'"
            s = s & ", trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
            s = s & " Where id = " & ds!id
            Wdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    s = "select id from paltasks where area = 'TRAFFIC MASTER' and status = 'PEND'"
    s = s & " and target > 'SR' "
    s = s & " and palletid not in (select palletid from queue_infc)"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update paltasks set status = 'COMP', userid = 'TMaster'"
            s = s & ", trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
            s = s & " Where id = " & ds!id
            Wdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    'Daifuku Pallet Drops   - jv041514
    s = "select id from paltasks where area = 'TRAFFIC MASTER' and palletid > '0' and status = 'PEND'"
    s = s & " and palletid in (select palletid from paltasks where area = 'DOCK' and userid > '0')"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update paltasks set status = 'COMP', userid = 'TMaster'"
            s = s & ", trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
            s = s & " Where id = " & ds!id
            Wdb.Execute s
            ds.MoveNext
        Loop
    End If
    
    s = "select id,target from paltasks where area = 'TRAFFIC MASTER' and status = 'PEND'"
    s = s & " and target in ('SR4', 'SR5')"
    s = s & " order by trandate desc"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        sr5cnt = 0
        Do Until ds.EOF
            If ds!target = "SR4" Then sr4cnt = sr4cnt + 1
            If ds!target = "SR5" Then sr5cnt = sr5cnt + 1
            If ds!target = "SR4" And sr4cnt > 3 Then
                s = "Update paltasks set status = 'COMP', userid = 'TMaster'"
                s = s & ", trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
                s = s & " Where id = " & ds!id
                Wdb.Execute s
            End If
            If ds!target = "SR5" And sr5cnt > 24 Then                'jv041514
                s = "Update paltasks set status = 'COMP', userid = 'TMaster'"
                s = s & ", trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
                s = s & " Where id = " & ds!id
                Wdb.Execute s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "process_tmaster_tasks", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: process_tmaster_tasks: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Sub process_wrapper_pallet(p As ptask)
    Dim pno As Long, pwhs As String
    If p.area = "TRAFFIC MASTER" Then
        pno = Val(p.reqid)                                               'jv091313
        pwhs = form1.Grid1.TextMatrix(1, 0)
        Call post_dai_exp_rcpt(p, pwhs, Format(pno, "000000"))          'jv091313
        If p.source = "TRI-LEVEL 1" Then Call post_route_to_kep(1, Val(pwhs), Format(pno, "000000"))  'jv012914
        If p.source = "TRI-LEVEL 2" Then Call post_route_to_kep(2, Val(pwhs), Format(pno, "000000"))  'jv012914
        If p.source = "TRI-LEVEL 3" Then Call post_route_to_kep(3, Val(pwhs), Format(pno, "000000"))  'jv012914
        If p.source = "TRI-LEVEL 4" Then Call post_route_to_kep(4, Val(pwhs), Format(pno, "000000"))  'jv012914
        Call post_tm_log(Val(pwhs), p, Format(pno, "000000"))           'jv091313
        
        Call record_pallet(Format(pno, "000000"), p, "SR" & pwhs, "Warehouse")  'jv091313
        If Val(pwhs) > 0 And Val(pwhs) < 4 Then
            Call post_queue_to_sr(Val(pwhs), p)
        End If
        If Val(pwhs) = 5 Then
            Call post_queue_to_sr(Val(pwhs), p)
        End If
        
        If Val(pwhs) = 4 Then Call post_sr4_efl(p)              'jv091313
        If Val(pwhs) >= 1 And Val(pwhs) <= 5 Then Call update_prodrcv(pwhs, p)
        p.target = "SR" & pwhs                                  'jv091313
        p.userid = "TMaster"                                    'jv091313
        Call process_tmaster_tasks(p)                           'jv091313
        If p.source = "TRI-LEVEL 1" Then form1.ws1 = pno        'jv091313
        If p.source = "TRI-LEVEL 2" Then form1.ws2 = pno        'jv091313
        If p.source = "TRI-LEVEL 3" Then form1.ws3 = pno        'jv091313
        If p.source = "TRI-LEVEL 4" Then form1.ws4 = pno        'jv091313
    End If
End Sub

Function queue_count(pwhs As String) As Integer
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    s = "select whse_num, count(*) from queue_infc"
    s = s & " where whse_num = " & pwhs
    s = s & " and queue_num > 0"
    s = s & " and source = 'TML'"
    s = s & " group by whse_num"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        queue_count = ds(1)
    Else
        queue_count = 0
    End If
    ds.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "queue_count", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: queue_count: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Sub read_barcode_sequences()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    sqlx = "select * from sequences"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!seq = "TLW1Barcode" Then form1.ws1 = ds!sequence_id
            If ds!seq = "TLW2Barcode" Then form1.ws2 = ds!sequence_id
            If ds!seq = "TLW3Barcode" Then form1.ws3 = ds!sequence_id
            If ds!seq = "TLW4Barcode" Then form1.ws4 = ds!sequence_id
            If ds!seq = "SPBarcode" Then form1.wsp = ds!sequence_id
            If ds!seq = "R0Barcode" Then form1.ws0 = ds!sequence_id
            If ds!seq = "RBBarcode" Then form1.wrb = ds!sequence_id
            If ds!seq = "BHBarcode" Then form1.wbh = ds!sequence_id
            ds.MoveNext
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "read_barcode_sequences", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: read_barcode_sequences: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Sub read_dai_message(xname As String, seqid As Long)
    Dim ds As adodb.Recordset, sqlx As String
    Dim cfile As String, f0 As String, clength As Long, i As Long
    On Error GoTo vberror
    clength = 0
    
    If daisqldb = bbsr Then                     'Testing detected
        sqlx = "select sMessage From WrxToHost"
        sqlx = sqlx & " Where iMessageSequence = " & seqid
        Set ds = DaiDb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            clength = Len(ds(0))
        End If
        If clength > 0 Then
            cfile = dailogs & "dai" & xname & ".xml"
            Open cfile For Output As #1
            Print #1, ds(0)
            Close #1
        Else
            cfile = dailogs & "dai" & xname & ".xml"
            Open cfile For Output As #1
            Print #1, "<" & xname & ">"
            Print #1, "<Sequence>" & seqid & "</Sequence>"
            Print #1, "<! Zero Length Message -->"
            Print #1, "</" & xname & ">"
            Close #1
        End If
        ds.Close
        Exit Sub
    End If
            
            
    
    sqlx = "select LEN(sMessage) FROM WrxToHost"
    sqlx = sqlx & " WHERE iMessageSequence = " & seqid
    Set ds = DaiDb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If IsNull(ds(0)) = False Then clength = ds(0)
    End If
    ds.Close
    If clength > 0 Then
        cfile = dailogs & "dai" & xname & ".xml"
        Open cfile For Output As #1
        For i = 1 To clength Step 256
            sqlx = "select SUBSTRING(sMessage, " & i & ", 256)"
            sqlx = sqlx & " FROM WrxToHost"
            sqlx = sqlx & " WHERE iMessageSequence = " & seqid
            Set ds = DaiDb.Execute(sqlx)
            If ds.BOF = False Then
                ds.MoveFirst
                Print #1, ds(0);
            End If
            ds.Close
        Next i
        Close #1
    Else
        cfile = dailogs & "dai" & xname & ".xml"
        Open cfile For Output As #1
        Print #1, "<" & xname & ">"
        Print #1, "<Sequence>" & seqid & "</Sequence>"
        Print #1, "<! Zero Length Message -->"
        Print #1, "</" & xname & ">"
        Close #1
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "read_dai_message", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: read_dai_message: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Private Sub record_pallet(pno As String, p As ptask, pwhs As String, pstat As String)
    Dim ds As adodb.Recordset, s As String
    Dim pid As Long, psku As String, recid As Long, q As Integer
    On Error GoTo vberror
    psku = Trim(Left(p.palletid, 4))
    q = Val(p.units) + Val(p.units2)
    recid = 0
    
    s = "select * from pallets where barcode = '" & p.palletid & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        recid = ds!id
    Else
        ds.Close
        s = "select * from pallets where status in ('Shipped','Order Pick')"
        s = s & " order by trandate"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            recid = ds!id
        End If
    End If
    ds.Close
    If recid > 0 Then
        s = "Update pallets set plateno = '" & Trim(pno) & "'"
        s = s & ",barcode = '" & p.palletid & "'"
        s = s & ",qty1 = " & Val(p.units)
        s = s & ",lot1 = '" & p.lotnum & "'"
        s = s & ",qty2 = " & Val(p.units2)
        s = s & ",lot2 = '" & p.lotnum2 & "'"
        s = s & ",source = '" & p.source & "'"
        s = s & ",target = '" & pwhs & "'"
        s = s & ",bbc = 'Y'"
        s = s & ",status = '" & pstat & "'"
        s = s & ",trandate = '" & p.trandate & "'"
        s = s & ",sku = '" & psku & "'"
        s = s & " Where id = " & recid
        Wdb.Execute s
    Else
        pid = wd_seq("Pallets")
        s = "Insert Into pallets Values (" & pid
        s = s & ",'" & Trim(pno) & "'"
        s = s & ",'" & p.palletid & "'"
        s = s & "," & Val(p.units)
        s = s & ",'" & p.lotnum & "'"
        s = s & "," & Val(p.units2)
        s = s & ",'" & p.lotnum2 & "'"
        s = s & ",'" & p.source & "'"
        s = s & ",'" & pwhs & "'"
        s = s & ",'Y'"
        If p.target = "ORDER PICK" Then
            s = s & ",'Order Pick'"
        Else
            s = s & ",'" & pstat & "'"
        End If
        s = s & ",'" & p.trandate & "'"
        s = s & ",'" & psku & "')"
        Wdb.Execute s
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "record_pallet", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: record_pallet: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Function sku_alloc(psku As String, plot As String, pcode As String, plot2 As String, pcode2 As String, pwhs As String) As Integer
    Dim ds As adodb.Recordset, s As String, pq As Integer                   'jv062614
    On Error GoTo vberror
    
    pq = 0                                                                  'jv062614
    s = "select sr" & pwhs & " from prodrcv"                                'jv062614
    s = s & " Where sku = '" & psku & "'"                                   'jv062614
    s = s & " and lot_num = '" & plot & "'"                                 'jv062614
    s = s & " and sp_flag = '" & pcode & "'"                                'jv062614
    Set ds = Wdb.Execute(s)                                                 'jv062614
    If ds.BOF = False Then                                                  'jv062614
        ds.MoveFirst                                                        'jv062614
        pq = ds(0)                                                          'jv062614
    End If                                                                  'jv062614
    ds.Close                                                                'jv062614
    If pq = 0 And plot2 > " " Then                                          'jv062614
        s = "select sr" & pwhs & " from prodrcv"                            'jv062614
        s = s & " Where sku = '" & psku & "'"                               'jv062614
        s = s & " and lot_num = '" & plot2 & "'"                            'jv062614
        s = s & " and sp_flag = '" & pcode2 & "'"                           'jv062614
        Set ds = Wdb.Execute(s)                                             'jv062614
        If ds.BOF = False Then                                              'jv062614
            ds.MoveFirst                                                    'jv062614
            pq = ds(0)                                                      'jv062614
        End If                                                              'jv062614
        ds.Close                                                            'jv062614
    End If                                                                  'jv062614
    If pq = 0 Then                                                          'jv062614
        s = "select sr" & pwhs & " from prodrcv"                            'jv062614
        s = s & " Where sku = '" & psku & "'"                               'jv062614
        s = s & " and lot_num in ('" & plot & "', '" & plot2 & "')"         'jv062614
        s = s & " and sp_flag in ('0', '1')"                                'jv062614
        s = s & " and sr" & pwhs & " > 0"                                   'jv062614
        Set ds = Wdb.Execute(s)                                             'jv062614
        If ds.BOF = False Then                                              'jv062614
            ds.MoveFirst                                                    'jv062614
            pq = ds(0)                                                      'jv062614
        End If                                                              'jv062614
        ds.Close                                                            'jv062614
    End If                                                                  'jv062614
    sku_alloc = pq                                                          'jv062614
    
    
    's = "select sr" & pwhs & " from prodrcv"
    's = s & " where sku = '" & psku & "'"
    's = s & " and lot_num = '" & plot & "'"
    'Set ds = Wdb.Execute(s)
    'If ds.BOF = False Then
    '    ds.MoveFirst
    '    sku_alloc = ds(0)
    'Else
    '    sku_alloc = 0
    'End If
    'ds.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "sku_alloc", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: sku_alloc: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Function sr_single_sku(pwhs As String, psku As String) As String
    Dim i As Integer, c As Integer, q As Long
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    q = 200
    If lastque > 200 Then q = lastque - 12
    c = 0
    s = "select whse_num, count(*) from queue_infc"
    s = s & " where whse_num = " & pwhs
    s = s & " and sku = '" & psku & "'"
    s = s & " and queue_num >= " & q
    s = s & " group by whse_num"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        c = ds(1)
    Else
        c = 0
    End If
    ds.Close
    If c = 1 Or c = 3 Or c = 5 Or c = 7 Then
        If Val(pwhs) < 4 Then
            sr_single_sku = "2"
        Else
            sr_single_sku = "1"
        End If
    Else
        sr_single_sku = "0"
    End If
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "sr_single_sku", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: sr_single_sku: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Sub update_prodrcv(pwhs As String, p As ptask)
    Dim s As String, psku As String
    On Error GoTo vberror
    psku = Trim(Left(p.palletid, 4))
    s = "Update prodrcv set sr" & pwhs & " = sr" & pwhs & " -1"
    s = s & " where sku = '" & psku & "'"
    s = s & " and lot_num = '" & p.lotnum & "'"
    's = s & " and sp_flag in ('" & Mid(p.palletid, 12, 1) & "', '0', '1')"          'jv062614
    s = s & " and sp_flag in ('" & Trim(Mid(p.palletid, 11, 3)) & "', '0', '1')"          'jv052515
    s = s & " and sr" & pwhs & " > 0"                                               'jv062614
    Wdb.Execute s
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "update_prodrcv", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: update_prodrcv: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Sub update_sr_lane(pno As String, paddr As String)
    Dim ds As adodb.Recordset, s As String, pid As Long, ls As adodb.Recordset
    Dim prow As String, pzone As Integer, pside As String, phorz As Integer, pvert As Integer
    Dim swarehouse As String, saddress As String, cfile As String, pq As Long
    Dim p As ptask                  'jv042015
    On Error GoTo vberror
    If Len(paddr) < 13 Then
        s = "Lane update failed with address: " & paddr
        Exit Sub
    End If
    If pno < "0" Then
        s = "Invalid plateno for lane update."
        Exit Sub
    End If
    saddress = Mid(paddr, 5, 9)
    prow = Left(saddress, 3)
    If prow = "025" Then
        pzone = 5
        pside = "L"
    End If
    If prow = "026" Then
        pzone = 5
        pside = "R"
    End If
    If prow = "027" Then
        pzone = 6
        pside = "L"
    End If
    If prow = "028" Then
        pzone = 6
        pside = "R"
    End If
    If prow = "029" Then
        pzone = 7
        pside = "L"
    End If
    If prow = "030" Then
        pzone = 7
        pside = "R"
    End If
    If prow = "031" Then
        pzone = 8
        pside = "L"
    End If
    If prow = "032" Then
        pzone = 8
        pside = "R"
    End If
    If prow = "033" Then
        pzone = 9
        pside = "L"
    End If
    If prow = "034" Then
        pzone = 9
        pside = "R"
    End If
    phorz = Val(Mid(saddress, 4, 3))
    pvert = Val(Mid(saddress, 7, 3))
    'cfile = dailogs & "SR5" & Format(Now, "MMdd") & ".csv"
    cfile = pallogs & "SR" & Format(Now, "MMddyyyy") & ".txt"                                   'jv060117
    Open cfile For Append As 6
    s = "select * from pallets where plateno = '" & pno & "'"
    s = s & " and target in ('SR5', 'SR6')"                 'jv070714
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        pq = ds!qty1 + ds!qty2                              'jv082313
        p.id = 0                                'jv042015
        p.area = "SR5"                          'jv042015
        p.description = " "                     'jv042015
        p.lotnum = ds!lot1                      'jv042015
        p.lotnum2 = ds!lot2                     'jv042015
        p.palletid = ds!barcode                 'jv042015
        p.product = ds!sku                      'jv042015
        p.qty = "1"                             'jv042015
        p.reqid = ds!plateno                    'jv042015
        p.source = ds!source                    'jv042015
        p.status = "PEND"                       'jv042015
        p.target = ds!target                    'jv042015
        p.trandate = ds!trandate                'jv042015
        p.units = ds!qty1                       'jv042015
        p.units2 = ds!qty2                      'jv042015
        p.uom = "Pallet"                        'jv042015
        p.userid = "WMS"                        'jv042015
        s = "select id from lane where whse_num = 5"
        s = s & " and zone_num = " & pzone
        s = s & " and vert_loc = " & pvert
        s = s & " and horz_loc = " & phorz
        s = s & " and rack_side = '" & pside & "'"
        Set ls = Wdb.Execute(s)
        If ls.BOF = False Then
            ls.MoveFirst
            pid = ls!id
            s = "Update lane set sku = '" & ds!sku & "'"
            s = s & ",lot_num = '" & ds!lot1 & "'"
            s = s & ",qty = 1"
            s = s & ",lot_date = GETDATE()"
            If pq > bbpallet_units(ds!sku) Then             'jv082313
                s = s & ",gmasize = " & pq                  'jv082313
            End If                                          'jv082313
            If check_hold(p) = True Then           'jv042015
                s = s & ",lane_status = 'H'"       'jv042015
            End If                                 'jv042015
            s = s & " Where id = " & pid
            Wdb.Execute s
            s = "Update position set sku = '" & ds!sku & "'"
            s = s & ",lot_num = '" & ds!lot1 & "'"
            s = s & ",pallet_num = '" & Right(ds!barcode, 3) & "'"
            s = s & ",count_qty = " & ds!qty1
            s = s & ",recv_date = GETDATE()"
            s = s & ",barcode = '" & ds!barcode & "'"
            s = s & ",lot2 = '" & ds!lot2 & "'"
            s = s & ",qty2 = " & ds!qty2
            s = s & " Where laneno = " & pid
            Wdb.Execute s
            's = pzone & " " & pvert & " " & phorz & " " & pside
            Write #6, p.id;                                         'jv060117
            Write #6, p.area;                                       'jv060117
            Write #6, p.description;                                'jv060117
            'Write #6, p.source;                                     'jv060117
            Write #6, "1300";                                       'jv060117
            'Write #6, p.target;                                     'jv060117
            s = pzone & " " & pvert & " " & phorz & " " & pside
            Write #6, s;                                            'jv060117
            
            s = ds!sku
            If labpix(Val(ds!sku)).package > " " Then s = s & " " & labpix(Val(ds!sku)).package
            If labpix(Val(ds!sku)).name1 > " " Then s = s & " " & labpix(Val(ds!sku)).name1
            If labpix(Val(ds!sku)).name2 > " " Then s = s & " " & labpix(Val(ds!sku)).name2
            If labpix(Val(ds!sku)).name3 > " " Then s = s & " " & labpix(Val(ds!sku)).name3
            Write #6, s;                                'p.product          jv060117
            
            'Write #6, p.product;                                    'jv060117
            Write #6, p.palletid;                                   'jv060117
            Write #6, p.qty;                                        'jv060117
            Write #6, p.uom;                                        'jv060117
            Write #6, p.lotnum;                                     'jv060117
            Write #6, p.units;                                      'jv060117
            Write #6, p.lotnum2;                                    'jv060117
            Write #6, p.units2                                      'jv060117
            Write #6, "COMP";                                       'jv060117
            Write #6, p.userid;                                     'jv060117
            Write #6, Format(Now, "yyMMdd hh:mm:ss");               'jv060117
            Write #6, p.reqid                                       'jv060117
            'Write #6, "SR-5";
            'Write #6, "...";
            'Write #6, ds!sku;
            'Write #6, barcode_to_lotnum(ds!barcode);
            'Write #6, Right(ds!barcode, 3);
            'Write #6, ds!barcode;
            'Write #6, pno;
            'Write #6, "1300";
            'Write #6, s;
            'Write #6, Format(Now, "h:mm am/pm")
        End If
        ls.Close
        s = "Update queue_infc Set queue_num = 0 Where whse_num in (2, 3, 5, 6)"  'jv011714
        s = s & " and palletid = '" & ds!barcode & "'"
        Wdb.Execute s
    End If
    ds.Close
    Close #6
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "update_sr_lane", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: update_sr_lane: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Sub vb_elog(eno As Long, edesc As String, pform As String, psub As String, puser As String)
    Dim i As Integer, s As String, cfile As String
    On Error GoTo vberror
    cfile = vberror_log
    'i = FreeFile(1)
    i = 88
    Open cfile For Append As #i
    Write #i, eno, edesc, pform, psub, Format(Now, "M-d-yyyy h:mm am/pm"), puser
    Close #i
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
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

Function wd_seq(tbname As String) As Long
    Dim ds As adodb.Recordset, s As String, i As Long
    On Error GoTo vberror
    i = 1
    s = "select sequence_id from sequences where seq = '" & tbname & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        i = ds!sequence_id + 1
        s = "update sequences set sequence_id = " & i
        s = s & " where seq = '" & tbname & "'"
        Wdb.Execute s
    End If
    ds.Close
    wd_seq = i
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "wd_seq", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: wd_seq: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Sub write_oracle_request(xname As String, mssgseq As Long)
    Dim sqlx As String
    Dim cfile As String, f0 As String, sxml As String
    On Error GoTo vberror
    sxml = ""
    cfile = dailogs & "dai" & xname & ".xml"
    Open cfile For Input As #1
    Do Until EOF(1)
        Line Input #1, f0
        sxml = sxml & f0
    Loop
    Close #1
    
    'Set db = CreateObject("ADODB.Connection")
    'db.Open daioradb
    ''sqlx = "Insert Into HostToWrx (iMessageSequence, sMessageIdentifier)"
    ''sqlx = sqlx & " Values (" & mssgseq & ", '" & xname & "')"
    '''MsgBox sqlx
    ''db.Execute sqlx
    
    
    ''sqlx = "Update HostToWrx"
    ''sqlx = sqlx & " Set sMessage = sMessage || '" & sxml & "'"
    ''sqlx = sqlx & " Where iMessageSequence = " & mssgseq
    ''db.Execute sqlx
    '''MsgBox sqlx
    If xname = "COBPalletReceipt" Then xname = "ExpectedReceiptMessage"                 'jv101414
    If xname = "AddItemHold" Then xname = "InventoryHoldMessage"                        'jv042015
    If xname = "RemItemHold" Then xname = "InventoryHoldMessage"                        'jv042015
    'If Mid(xname, 1, 11) = "AddItemHold" Then xname = "InventoryHoldMessage"         'jv042315
    'If Mid(xname, 1, 11) = "RemItemHold" Then xname = "InventoryHoldMessage"         'jv042315
    
    If Len(sxml) > 100 Then                         'jv042415
        sqlx = "Insert Into HostData.HostToWrx (iMessageSequence, sMessageIdentifier, sMessage)"
        sqlx = sqlx & " Values (" & mssgseq & ", '" & xname & "', '" & sxml & "')"
        DaiDb.Execute sqlx                          'jv042415
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "bbcdai.bas", "write_oracle_request", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: write_oracle_request: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

