Attribute VB_Name = "Module1"
Option Explicit
'Public skutab(1000, 4) As String
Public skutab(9999, 4) As String                                        'jv082415
Public logdir     As String '= "\\bbc-01-wdmgmt\wd\data\testlog"
Public tracelist  As String '= " "
Public debflag    As Boolean '= False
Public SPTarget   As String '= "SNACK PLANT"
Public BHDest     As String '= "STAGING"
Public srflag     As Boolean '= False
Public ARFlag     As Boolean '= False
Public TCarFlag   As Boolean '= False
Public WDOrg      As String
Public WDUserId   As String
Public WDbbsr     As String
Public daioradb   As String
Public daidock    As String
Public ship_units As String
Public ship_lotnum As String
Public ship_units2 As String
Public ship_lotnum2 As String
Public ship_plate As String
Public histbc     As String                                             'jv052515
Public Wdb As ADODB.Connection
Global eno As Long
Global edesc As String
Global vberror_log As String
'Global labpix(1000) As labpic
Global labpix(9999) As labpic                                           'jv082415
Global labfmtfile As String

Public Type ptask
    id          As Long
    area        As String
    description As String
    source      As String
    target      As String
    product     As String
    palletid    As String
    qty         As String
    uom         As String
    lotnum      As String
    units       As String
    lotnum2     As String
    units2      As String
    status      As String
    userid      As String
    trandate    As String
    reqid       As String
End Type

Type daimessagerec
    dhostmodifytime As String
    imessagesequence As String
    smessageidentifier As String
    smessage As String
    bbcidentity As String
    bbcstatus As String
End Type

Public Type labpic
    sku As String
    package As String
    name1 As String
    name2 As String
    name3 As String
End Type

Public Sub add_alternate_dock_pallet(recid As Long, psource As String, paltarg As String)
    Dim p As ptask
    'db As ADODB.Connection,
    Dim rs As ADODB.Recordset, s As String, zid As Long
    Dim sSql As String, i As Long, psku As String, pgroup As String, sCols As String, sRows As String
    On Error GoTo vberror
    p = masterec(recid)
    psku = Trim(Left(p.product, 4))
    pgroup = Trim(p.description)
    p.source = psource
    p.target = paltarg
    If psource = "SR5" Or psource = "SR6" Then
        If Len(psku) = 3 Then                                   'jv082415
            p.palletid = psku & " ...... . ..."
        Else
            p.palletid = psku & "...... . ..."                  'jv082415
        End If
    End If
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")
    zid = insert_trans(p)
    If psource = "STAGING" Then
        p.area = "FORKLIFT"
        p.description = p.description & Space(8 - Len(p.description)) & p.target
        p.source = "RACKS"
        p.target = "STAGING"
        p.trandate = Format(Now, "yyMMdd hh:mm:ss")
        zid = insert_trans(p)
    End If
    If psource = "SR1" Or psource = "SR2" Or psource = "SR3" Then
        'Set db = CreateObject("ADODB.Connection")
        'db.Open WDbbsr
        sSql = "Select id, ship_status From ship_infc Where order_num = '" & pgroup & "'"
        sSql = sSql & " And sku = '" & psku & "'"
        sSql = sSql & " And to_whse_num = " & Mid(psource, 3, 1)
        Set rs = Wdb.Execute(sSql)
        If rs.BOF = False Then
            rs.MoveFirst
            i = rs!id
            s = rs!ship_status
            If s = "DONE" Or s = "CANC" Then
                sSql = "Update ship_infc Set order_qty = 1, ship_uom_qty = 0"
                sSql = sSql & ", ship_plt_qty = 0, ship_status = 'NEW'"
                sSql = sSql & " Where id = " & i
                Wdb.Execute (sSql)
            Else
                sSql = "Update ship_infc Set order_qty = order_qty + 1"
                sSql = sSql & " Where id = " & i
                Wdb.Execute (sSql)
            End If
        Else
            rs.Close
            sSql = "Select id, ship_status From ship_infc"
            sSql = sSql & " Where ship_status in ('CANC','DONE')"
            sSql = sSql & " Order By id"
            Set rs = Wdb.Execute(sSql)
            If rs.BOF = False Then
                rs.MoveFirst
                i = rs!id
                sSql = "Update ship_infc Set order_num = '" & pgroup & "'"
                sSql = sSql & ",sku = '" & psku & "'"
                sSql = sSql & ",ship_date = '" & Format(Now, "M/d/yyyy") & "'"
                sSql = sSql & ",order_qty = 1"
                sSql = sSql & ",ship_uom_qty = 0"
                sSql = sSql & ",ship_plt_qty = 0"
                sSql = sSql & ",ship_status = 'NEW'"
                If psource = "SR1" Then
                    sSql = sSql & ",to_whse_num = 1"
                    sSql = sSql & ",to_vert_loc = 2"
                    sSql = sSql & ",to_horz_loc = 18"
                    sSql = sSql & ",to_rack_side = 'L'"
                End If
                If psource = "SR2" Then
                    sSql = sSql & ",to_whse_num = 2"
                    sSql = sSql & ",to_vert_loc = 2"
                    sSql = sSql & ",to_horz_loc = 39"
                    sSql = sSql & ",to_rack_side = 'L'"
                End If
                If psource = "SR3" Then
                    sSql = sSql & ",to_whse_num = 3"
                    sSql = sSql & ",to_vert_loc = 2"
                    sSql = sSql & ",to_horz_loc = 43"
                    sSql = sSql & ",to_rack_side = 'R'"
                End If
                sSql = sSql & " Where id = " & i
                'MsgBox sSql
                Wdb.Execute (sSql)
            End If
            rs.Close
        End If
        'db.Close
    End If
    'If psource = "SR5" Or psource = "SR6" Then
    '    MsgBox "daifuku task"
    'End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "add_alternate_dock_pallet", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "add_alternate_dock_pallet", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: add_alternate_dock_pallet: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Sub add_alternate_daifuku(ptarget As String, pdock As String, psku As String, pdesc As String)
    Dim xname As String, cfile As String, s As String
    On Error GoTo vberror
    xname = "OrderItemMessage"
    'cfile = "c:\jvwork\dai" & xname & ".xml"
    cfile = "\\bbc-01-prodtrk\wd\sr5\bin\dai" & xname & ".xml"
    Open cfile For Output As #1
    s = "<?xml version=" & Chr(34) & "1.0" & Chr(34)
    s = s & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>" & vbCrLf
    s = s & "<!DOCTYPE OrderItemMessage SYSTEM " & Chr(34) & "wrxj.dtd" & Chr(34) & ">" & vbCrLf
    s = s & "<OrderItemMessage>" & vbCrLf
    s = s & "<Order action=" & Chr(34) & "ADD" & Chr(34)
    s = s & " sOrderID=" & Chr(34) & "ALT" & Right(DateDiff("s", "1-1-13 01:00:00 am", Now), 5) & Chr(34)
    s = s & " iPriority=" & Chr(34) & "3" & Chr(34)
    s = s & " iOrderStatus=" & Chr(34) & "READY" & Chr(34) & ">" & vbCrLf
    Print #1, s
    
    s = "<OrderHeader>" & vbCrLf
    s = s & "<sDestinationStation>" & pdock & "</sDestinationStation>" & vbCrLf
    s = s & "<sDescription>" & ptarget & "</sDescription>" & vbCrLf
    s = s & "<sOrderMessage/>" & vbCrLf
    s = s & "</OrderHeader>" & vbCrLf
    Print #1, s
    
    s = "<OrderLine sItem=" & Chr(34) & psku & Chr(34) & ">" & vbCrLf
    s = s & "<sRouteID/>" & vbCrLf
    s = s & "<fOrderQuantity>1</fOrderQuantity>" & vbCrLf
    s = s & "<sDescription>" & pdesc & "</sDescription>" & vbCrLf
    s = s & "</OrderLine>" & vbCrLf
    Print #1, s
    
    s = "</Order>" & vbCrLf
    s = s & "</OrderItemMessage>"
    Print #1, s
    Close #1
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "add_alternate_daifuku", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "add_alternate_daifuku", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: add_alternate_daifuku: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Sub add_back_prodrcv(pwhs As String, psku As String, plot As String, pcode As String)
    Dim s As String
    On Error GoTo vberror
    s = "Update prodrcv set " & pwhs & " = " & pwhs & " + 1"
    s = s & " where sku = '" & psku & "'"
    s = s & " and lot_num = '" & plot & "'"
    s = s & " and sp_flag = '" & pcode & "'"
    Wdb.Execute s
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "add_alternate_daifuku", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "add_back_prodrcv", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: add_back_prodrcv: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Function barcode_profile(ubar As String) As Boolean
    'On Error Resume Next
    Dim s As String
    barcode_profile = True

    'Check SKU
    s = Trim(Mid(ubar, 1, 4))
    If s <> sku_info(s, "sku") Then barcode_profile = False

    'Check Month
    s = Mid(ubar, 5, 2)
    If Val(s) < 1 Or Val(s) > 12 Then barcode_profile = False

    'Check Day
    s = Mid(ubar, 7, 2)
    If Val(s) < 1 Or Val(s) > 31 Then barcode_profile = False

    'Check Year
    s = Mid(ubar, 9, 2)
    If Val(s) < 11 Or Val(s) > 44 Then barcode_profile = False
    
    'Check 3-digit OpCode           jv072915
    s = Mid(ubar, 11, 3)
    If Val(s) < 100 Or Val(s) > 699 Then barcode_profile = False    'jv081818

    'Check spaces
    'If Mid(ubar, 11, 1) <> " " Then barcode_profile = False
    'If Mid(ubar, 13, 1) <> " " Then barcode_profile = False

    'Check OpCode
    's = UCase(Mid(ubar, 12, 1))
    'If s < "A" Or s > "Z" Then barcode_profile = False

    'Check pallet sequence number for spaces
    If Mid(ubar, 14, 1) = " " Then barcode_profile = False
    If Mid(ubar, 15, 1) = " " Then barcode_profile = False
    If Mid(ubar, 16, 1) = " " Then barcode_profile = False
End Function

Public Function barcode_to_lotnum(mbar As String)
    'On Error Resume Next
    Dim s1 As String
    Dim s2 As String
    Dim s As String
    Dim j As Long
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
            s1 = "01-01-20" & Format(Val(Mid(mbar, 9, 2)) - 2, "00")
            s2 = Mid(mbar, 5, 2) & "-" & Mid(mbar, 7, 2) & "-20" & Format(Val(Mid(mbar, 9, 2)) - 2, "00")
            j = DateDiff("d", s1, s2) + 1
            s = Format(Val(Mid(mbar, 9, 2)) - 2, "00")
            s = s & Format(j, "000")
        End If
        barcode_to_lotnum = s
    End If
End Function

Public Function check_anteroom(msrc As String, mbar As String) As String
    'On Error Resume Next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    Dim sSql As String
    'tracelist = tracelist & "<!-- check anteroom(" & msrc & "," & mbar & ")"
    sSql = "Select sku from rackpos where barcode = '" & mbar & "'" & _
           " and rackno in (select id from racks where aisle = 'M' and rack = 'ANTE')"

    Set ds = Wdb.Execute(sSql)
    
    If ds.BOF = False Then
      check_anteroom = "ANTE ROOM"
      'tracelist = tracelist & " = ANTE ROOM"
    Else
      check_anteroom = msrc
    End If
    'tracelist = tracelist & " -->" & vbCrLf
    ds.Close ': db.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "check_anteroom", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "check_anteroom", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: check_anteroom: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Function check_hold(p As ptask) As Boolean                                      'jv040615
    Dim psku As String, pcode As String, hflag As Boolean, s As String
    Dim ds As ADODB.Recordset, palno As String
    psku = Trim(Mid(p.palletid, 1, 4))
    'pcode = Mid(p.palletid, 12, 1)
    pcode = Trim(Mid(p.palletid, 11, 3))                                        'jv052515
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
                If palno >= ds!spallet And palno < ds!epallet Then
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

Public Sub crane_finished_goods_lane(m As ptask)
    'On Error Resume Next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    Dim sSql As String, sRows As String, sCols As String
    Dim K As Integer, zid As Long, ncnt As Long, dflag As String
    On Error GoTo vberror
    zid = new_pallet_queue()
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    'tracelist = tracelist & "<!-- crane_finished_goods_lane(" & m.id & ") -->" & vbCrLf
    If check_hold(m) = True Then        'jv040615
        dflag = "H"                     'jv040615
    Else                                'jv040615
        dflag = " "                     'jv040615
    End If                              'jv040615
    sSql = "select * from queue_infc where id = " & zid
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        sSql = "update queue_infc set whse_num = " & Right(m.target, 1) & "," & _
                "sku = '" & Trim(Mid(m.palletid, 1, 4)) & "'," & _
                "lot_num = '" & m.lotnum & "'," & _
                "drop_flag = '" & dflag & "'," & _
                "rack_num = " & Val(m.qty) & "," & _
                "units = " & Val(m.units) & "," & _
                "lot_num2 = '" & m.lotnum2 & "'," & _
                "units2 = " & Val(m.units2) & "," & _
                "palletid = '" & m.palletid & "'," & _
                "source = 'FG" & Right(m.target, 1) & "'" & _
                " where id = " & zid
        Wdb.Execute (sSql)
    Else
        zid = wd_seq("Queue_Infc")
        sSql = "INSERT INTO queue_infc (ID,Whse_num,SKU,Lot_num,Drop_Flag,Queue_Num," & _
               "Rack_Num,Units,Lot_Num2,Units2,PalletId,Source) VALUES (" & _
               zid & "," & _
               Right(m.target, 1) & ",'" & _
               Trim(Mid(m.palletid, 1, 4)) & "','" & _
               m.lotnum & "'," & _
                "'" & dflag & "'," & _
                "0," & _
                Val(m.qty) & "," & _
                Val(m.units) & ",'" & _
                m.lotnum2 & "'," & _
                Val(m.units2) & ",'" & _
                m.palletid & "'," & _
                "'FG" & Right(m.target, 1) & "'"
           Wdb.Execute (sSql)
    End If
    ds.Close

    sSql = "Update prodrcv set sr" & Right(m.target, 1) & "= sr" & Right(m.target, 1) & " - 1" & _
           " Where sku = '" & Trim(Mid(m.palletid, 1, 4)) & "'" & _
           " And lot_num = '" & m.lotnum & "'"
    Wdb.Execute (sSql)
    'db.Close
    m.status = "COMP"
    m.trandate = Format(Now, "yyMMdd HH:mm:ss")
    Call post_move_trans(m)
    Call update_trans(m)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "crane_finished_goods_lane", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "crane_finished_goods_lane", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: crane_finished_goods_lane: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Sub dai_ExpectedReceipt_Cancel(bc As String, plateno As String, swhs As String)      'jv101414
    Dim psku As String, plot As String, pqty As String
    Dim s As String, cfile As String, dailogs As String
    'dailogs = "v:\testlogs\"
    dailogs = "\\bbc-01-prodtrk\wd\sr5\bin\"
    psku = Trim(Left(bc, 4))
    plot = barcode_to_lotnum(bc)
    plot = plot & Mid(bc, 12, 1) & Right(bc, 3)
    pqty = sku_info(psku, "units")
    
    s = "<?xml version=" & Chr(34) & "1.0" & Chr(34)
    s = s & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>" & vbCrLf
    s = s & "<!DOCTYPE ExpectedReceiptMessage SYSTEM " & Chr(34) & "wrxj.dtd" & Chr(34) & ">" & vbCrLf
    s = s & "<ExpectedReceiptMessage>" & vbCrLf
    s = s & "<ExpectedReceipt action=" & Chr(34) & "DELETE" & Chr(34)
    s = s & " sOrderID=" & Chr(34) & Format(Val(plateno), "000000") & Chr(34) & ">" & vbCrLf
    s = s & "<ExpectedReceiptHeader>" & vbCrLf
    s = s & "<dExpectedDate>" & Format(Now, "MM/dd/yyyy hh:mm:ss") & "</dExpectedDate>" & vbCrLf
    s = s & "</ExpectedReceiptHeader>" & vbCrLf
    s = s & "<ExpectedReceiptLine sItem=" & Chr(34) & psku & Chr(34) & " sLot=" & Chr(34) & plot & Chr(34) & ">" & vbCrLf
    s = s & "<fExpectedQuantity>" & pqty & "</fExpectedQuantity>" & vbCrLf
    s = s & "<sStoreDestination>" & swhs & "</sStoreDestination>" & vbCrLf
    s = s & "<sRouteID/>" & vbCrLf
    s = s & "<sHoldReason/>" & vbCrLf
    s = s & "</ExpectedReceiptLine>" & vbCrLf
    s = s & "</ExpectedReceipt>" & vbCrLf
    s = s & "</ExpectedReceiptMessage>"
        
    cfile = dailogs & "daiCOBPalletReceipt.xml"
    'cfile = dailogs & "daiOrderItemMessage.xml"
    Open cfile For Output As #1
    Print #1, s
    Close #1
    
    cfile = dailogs & "daimessages" & Format(Now, "MMddyy") & ".txt"
    Open cfile For Append As #8
    Print #8, "-------------"
    Print #8, s
    Close #8
    
    DoEvents
    
End Sub

Public Sub debug_log(s As String, m As ptask, muser As String)
    MsgBox s, vbOKOnly, "Debug..."
End Sub

Public Sub dockfl_to_rstg(m As ptask)
    'On Error Resume Next
    'tracelist = tracelist & "<!-- dockfl_to_rstg(" & m.id & ") -->" & vbCrLf
    'MsgBox "dockfl_to_rstg"
    If m.id = 0 Then Exit Sub
    If srflag = True Or TCarFlag = True Then
        ship_units = m.units
        ship_lotnum = m.lotnum
        ship_units2 = m.units2
        ship_lotnum2 = m.lotnum2
        ship_plate = m.reqid
        Call pallet_lots(m.palletid)
        m.units = ship_units
        m.lotnum = ship_lotnum
        m.units2 = ship_units2
        m.lotnum2 = ship_lotnum2
        m.reqid = ship_plate
        If m.lotnum < "0" Then m.lotnum = barcode_to_lotnum(m.palletid)
    End If
    Call post_move_trans(m)
    m.area = "FORKLIFT"
    m.description = " "
    m.source = "STAGING"
    m.target = "..."
    m.status = "PEND"
    m.userid = " "
    m.trandate = Format(Now, "yyMMdd HH:mm:ss")
    'Call post_ship_trans(m)
    Call update_trans(m)
End Sub

Public Sub dockfl_to_trailer(m As ptask)
    'On Error Resume Next
    'tracelist = tracelist & "<!-- dockfl_to_trailer(" & m.id & ") -->" & vbCrLf
    If m.id = 0 Then Exit Sub
    If srflag = True Or TCarFlag = True Then
        ship_units = m.units
        ship_lotnum = m.lotnum
        ship_units2 = m.units2
        ship_lotnum2 = m.lotnum2
        ship_plate = m.reqid
        Call pallet_lots(m.palletid)
        m.units = ship_units
        m.lotnum = ship_lotnum
        m.units2 = ship_units2
        m.lotnum2 = ship_lotnum2
        m.reqid = ship_plate
        If m.lotnum < "0" Then m.lotnum = barcode_to_lotnum(m.palletid)
    End If
    m.status = "PEND"
    m.trandate = Format(Now, "yyMMdd HH:mm:ss")
    Call post_ship_trans(m)
    Call update_trans(m)
End Sub

Public Sub efl_rack_moves(m As ptask)
    'On Error Resume Next
    Dim psku As String
    Dim paisle As String
    Dim prack As String
    Dim s As String
    'tracelist = tracelist & "<!-- efl_rack_moves(" & m.id & ") -->" & vbCrLf
    If m.id = 0 Then Exit Sub
    If m.target = "ORDER PICK" Or m.target = "SNACK PLANT" Or m.target = "ANTE ROOM" Then
    Else
        If m.target = "M-OP" Or m.target = "M-SP" Or m.target = "M-ANTE" Then
        Else
            psku = Trim(Mid(m.palletid, 1, 4))
            paisle = Mid(m.target, 1, 1)
            If Mid(m.target, 2, 1) = " " Or Mid(m.target, 2, 1) = "-" Then
                prack = Mid(m.target, 3, Len(m.target) - 2)
            Else
                prack = Mid(m.target, 2, Len(m.target) - 1)
                m.target = paisle & "-" & Trim(prack)
            End If
            If space_in_rack(paisle, prack, psku, m) = False Then
                'Complete and log task even though inventory is not correct.
                m.status = "COMP"
                m.trandate = Format(Now, "yyMMdd HH:mm:ss")
                Call post_move_trans(m)
                Call update_trans(m)
                Call debug_log("Database shows no space in target rack: " & m.target, m, m.userid)
                Exit Sub
            End If
        End If
    End If
    If m.source = "ROBOT ZERO" Or m.source = "WRAPPER" Or m.source = "TRI LEVEL" Or m.source = "TRI-LEVEL" Or m.source = "PALLET" Or m.source = "CRANE" Or m.source = "STAGING" Or m.source = "BACKHAUL" Or m.source = "QUEUE" Then
        If insert_rack_pallet(m) = True Then
            m.status = "COMP"
            m.trandate = Format(Now, "yyMMdd HH:mm:ss")
            Call post_move_trans(m)
            Call update_trans(m)
        Else
            s = "Error in efl_rack_moves, updating target: " & m.target
            s = s & " source: " & m.source & " bc: " & m.palletid
            Call debug_log(s, m, m.userid)
        End If
    Else
        If remove_rack_pallet(m) = True Then
            If insert_rack_pallet(m) = True Then
                m.status = "COMP"
                m.trandate = Format(Now, "yyMMdd HH:mm:ss")
                Call post_move_trans(m)
                Call update_trans(m)
            Else
                s = "Error in efl_rack_moves, updating target: " & m.target
                s = s & " source: " & m.source & " bc: " & m.palletid
                Call debug_log(s, m, m.userid)
            End If
        Else
            s = "Error in efl_rack_moves, removing from source: " & m.source
            s = s & " bc: " & m.palletid
            Call debug_log(s, m, m.userid)
        End If
    End If
End Sub

Public Sub efl_to_dstg(m As ptask)
    'On Error Resume Next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    Dim sSql As String, sCols As String, sRows As String
    Dim K As Long
    On Error GoTo vberror
    'tracelist = tracelist & "<!-- efl_to_dstg(" & m.id & ") -->" & vbCrLf
    If m.id = 0 Then Exit Sub
    If remove_rack_pallet(m) = True Then
        m = masterec(m.id)
    Else
        m = masterec(m.id)
    End If
  
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
  
    sSql = "Select id From paltasks Where area = 'DOCK'"
    sSql = sSql & " And description = '" & Trim(Mid(m.description, 1, 6)) & "'"
    sSql = sSql & " And target = '" & Right(m.description, Len(m.description) - 8) & "'"
    sSql = sSql & " And source in ('STAGING','ANTE ROOM','M-ANTE')"
    sSql = sSql & " And product = '" & m.product & "'"
    sSql = sSql & " And palletid < '0'"
    sSql = sSql & " And status = 'PEND'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        K = ds!id
        sSql = "Update paltasks Set source = 'STAGING',"
        sSql = sSql & "palletid = '" & m.palletid & "',"
        sSql = sSql & "qty = " & Val(m.qty) & ","
        sSql = sSql & "lotnum = '" & m.lotnum & "',"
        sSql = sSql & "units = " & Val(m.units) & ","
        sSql = sSql & "lotnum2 = '" & m.lotnum2 & "',"
        sSql = sSql & "units2 = " & Val(m.units2)
        sSql = sSql & " Where id = " & K
        Wdb.Execute (sSql)
        m.status = "COMP"
        m.trandate = Format(Now, "yyMMdd HH:mm:ss")
        Call update_trans(m)
    Else
        Call debug_log("Failed efl_to_dstg: " & sSql, m, m.userid)
    End If
    ds.Close ': db.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "efl_to_dstg", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "efl_to_dstg", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: efl_to_dstg: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Function insert_rack_pallet(p As ptask) As Boolean
    'On Error Resume Next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    Dim hs As ADODB.Recordset                                                   'jv111314
    Dim sSql As String
    Dim K As Integer, j As Integer
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
    If psku <> sku_info(psku, "sku") Then
        Call debug_log(p.palletid & " invalid SKU in pallet barcode, during insert_rack_pallet.", p, p.userid)
        insert_rack_pallet = False
        Exit Function
    End If

    pqty = Val(p.units) + Val(p.units2)
    If pqty > Val(sku_info(psku, "units")) Then
        pbbc = False
    Else
        pbbc = True
    End If
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

    K = 0
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        K = ds(0)
        sSql = "select hold from racks where id = " & K                 'jv111314
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

    If K = 0 Then
        'db.Close
        Call debug_log("failed to update rack in insert_pallet_rack: " & p.target, p, p.userid)
        insert_rack_pallet = False
        Exit Function
    End If

    j = 0
    sSql = "Select id From rackpos Where rackno = " & K & _
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
            sSql = sSql & " Values (" & zid & "," & K & "," & order_pick_position(psku) & ","
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
            Call debug_log(sSql & "failed to update in insert_rack_pallet", p, "0")
            insert_rack_pallet = False
            Exit Function
        End If
    End If

    olot = "99999"
    pqty = 0
    pqty4 = 0

    sSql = "Select sku,lot_num,lot2,bbc From rackpos Where rackno = " & K & _
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
        Call debug_log(sSql & " failed in insert_rack_pallet..", p, "0")
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
    sSql = sSql & " Where id = " & K
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

Public Function masterec(taskid As Long) As ptask
    'On Error Resume Next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    Dim sSql As String
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
  
    sSql = "Select * From paltasks Where id = " & taskid
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        masterec.id = ds!id
        masterec.area = ds!area
        masterec.description = ds!description
        masterec.source = ds!source
        masterec.target = ds!target
        masterec.product = ds!product
        masterec.palletid = ds!palletid
        masterec.qty = ds!qty
        masterec.uom = ds!uom
        masterec.lotnum = ds!lotnum
        masterec.units = ds!units
        masterec.lotnum2 = ds!lotnum2
        masterec.units2 = ds!units2
        masterec.status = ds!status
        masterec.userid = ds!userid
        masterec.trandate = ds!trandate
        masterec.reqid = ds!reqid
    Else
        masterec.id = 0
        masterec.area = " "
        masterec.description = " "
        masterec.source = " "
        masterec.target = " "
        masterec.product = " "
        masterec.palletid = " "
        masterec.qty = " "
        masterec.uom = " "
        masterec.lotnum = " "
        masterec.units = " "
        masterec.lotnum2 = " "
        masterec.units2 = " "
        masterec.status = " "
        masterec.userid = " "
        masterec.trandate = " "
        masterec.reqid = " "
    End If
    ds.Close ': db.Close
    'tracelist = tracelist & "<!-- masterec(" & taskid & ")=" & masterec.palletid
    'tracelist = tracelist & "," & masterec.lotnum & "," & masterec.units & ","
    'tracelist = tracelist & masterec.lotnum2 & "," & masterec.units2 & ","
    'tracelist = tracelist & masterec.uom & "," & masterec.qty & ","
    'tracelist = tracelist & masterec.target & " -->" & vbCrLf
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "masterec", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "masterec", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: masterec: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Function move_task_crane(srkey As Long, bc As String, ptarget As String, puser As String) As ptask
    'on error resume next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset, ss As ADODB.Recordset
    Dim sSql As String, sCols As String, sRows As String
    Dim paisle As String, prack As String
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    sSql = "select * from position where id = " & srkey & " and barcode = '" & bc & "'"
    'MsgBox sSql
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        move_task_crane.area = "EFLMove"
        move_task_crane.description = "Rack Move"
        move_task_crane.source = "CRANE"
        move_task_crane.target = ptarget
        move_task_crane.product = ds!sku & " " & sku_info(ds!sku, "desc")
        move_task_crane.palletid = ds!barcode
        move_task_crane.qty = "1"
        move_task_crane.uom = "Pallet"
        move_task_crane.lotnum = ds!lot_num
        move_task_crane.units = ds!count_qty
        move_task_crane.lotnum2 = ds!lot2
        move_task_crane.units2 = ds!qty2
        move_task_crane.status = "PEND"
        move_task_crane.userid = puser
        move_task_crane.trandate = Format(Now, "yyMMdd HH:mm:ss")
        move_task_crane.reqid = " "
    Else
        move_task_crane.area = "FAILED"
        move_task_crane.description = " "
        move_task_crane.source = " "
        move_task_crane.target = " "
        move_task_crane.product = " "
        move_task_crane.palletid = " "
        move_task_crane.qty = " "
        move_task_crane.uom = " "
        move_task_crane.lotnum = " "
        move_task_crane.units = " "
        move_task_crane.lotnum2 = " "
        move_task_crane.units2 = " "
        move_task_crane.status = " "
        move_task_crane.userid = " "
        move_task_crane.trandate = " "
        move_task_crane.reqid = " "
    End If
    ds.Close ': db.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "move_task_crane", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "move_task_crane", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: move_task_crane: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Function move_task_pallet(palkey As Long, bc As String, ptarget As String, puser As String) As ptask
    'on error resume next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset, ss As ADODB.Recordset
    Dim sSql As String, sCols As String, sRows As String
    Dim paisle As String, prack As String
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    'sSql = "select * from rackpos where id = " & rackkey
    sSql = "select * from pallets where id = " & palkey & " and barcode = '" & bc & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        move_task_pallet.area = "EFLMove"
        move_task_pallet.description = "Rack Move"
        move_task_pallet.source = "PALLET"
        move_task_pallet.target = ptarget
        move_task_pallet.product = ds!sku & " " & sku_info(ds!sku, "desc")
        move_task_pallet.palletid = ds!barcode
        move_task_pallet.qty = "1"
        move_task_pallet.uom = "Pallet"
        move_task_pallet.lotnum = ds!lot1
        move_task_pallet.units = ds!qty1
        move_task_pallet.lotnum2 = ds!lot2
        move_task_pallet.units2 = ds!qty2
        move_task_pallet.status = "PEND"
        move_task_pallet.userid = puser
        move_task_pallet.trandate = Format(Now, "yyMMdd HH:mm:ss")
        move_task_pallet.reqid = ds!plateno
    Else
        move_task_pallet.area = "FAILED"
        move_task_pallet.description = " "
        move_task_pallet.source = " "
        move_task_pallet.target = " "
        move_task_pallet.product = " "
        move_task_pallet.palletid = " "
        move_task_pallet.qty = " "
        move_task_pallet.uom = " "
        move_task_pallet.lotnum = " "
        move_task_pallet.units = " "
        move_task_pallet.lotnum2 = " "
        move_task_pallet.units2 = " "
        move_task_pallet.status = " "
        move_task_pallet.userid = " "
        move_task_pallet.trandate = " "
        move_task_pallet.reqid = " "
    End If
    ds.Close ': db.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "move_task_pallet", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "move_task_pallet", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: move_task_pallet: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Function move_task_queue(quekey As Long, bc As String, ptarget As String, puser As String) As ptask
    'on error resume next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset, ss As ADODB.Recordset
    Dim sSql As String, sCols As String, sRows As String
    Dim paisle As String, prack As String
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    sSql = "select * from queue_infc where id = " & quekey & " and palletid = '" & bc & "'"
    'MsgBox sSql
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        move_task_queue.area = "EFLMove"
        move_task_queue.description = "Rack Move"
        move_task_queue.source = "QUEUE"
        move_task_queue.target = ptarget
        move_task_queue.product = ds!sku & " " & sku_info(ds!sku, "desc")
        move_task_queue.palletid = ds!palletid
        move_task_queue.qty = "1"
        move_task_queue.uom = "Pallet"
        move_task_queue.lotnum = ds!lot_num
        move_task_queue.units = ds!units
        move_task_queue.lotnum2 = ds!lot_num2
        move_task_queue.units2 = ds!units2
        move_task_queue.status = "PEND"
        move_task_queue.userid = puser
        move_task_queue.trandate = Format(Now, "yyMMdd HH:mm:ss")
        move_task_queue.reqid = " "
    Else
        move_task_queue.area = "FAILED"
        move_task_queue.description = " "
        move_task_queue.source = " "
        move_task_queue.target = " "
        move_task_queue.product = " "
        move_task_queue.palletid = " "
        move_task_queue.qty = " "
        move_task_queue.uom = " "
        move_task_queue.lotnum = " "
        move_task_queue.units = " "
        move_task_queue.lotnum2 = " "
        move_task_queue.units2 = " "
        move_task_queue.status = " "
        move_task_queue.userid = " "
        move_task_queue.trandate = " "
        move_task_queue.reqid = " "
    End If
    If move_task_queue.area <> "FAILED" Then                                        'jv111314
        sSql = "Update queue_infc set queue_num = 0 where palletid = '" & bc & "'"  'jv111314
        Wdb.Execute sSql                                                            'jv111314
    End If                                                                          'jv111314
    ds.Close ': db.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "move_task_queue", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "move_task_queue", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: move_task_queue: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Function move_task_rack(rackkey As Long, bc As String, ptarget As String, puser As String) As ptask
    'on error resume next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset, ss As ADODB.Recordset
    Dim sSql As String, sCols As String, sRows As String
    Dim paisle As String, prack As String
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    'sSql = "select * from rackpos where id = " & rackkey
    sSql = "select * from rackpos where rackno = " & rackkey & " and barcode = '" & bc & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        move_task_rack.area = "EFLMove"
        move_task_rack.description = "Rack Move"
        sSql = "Select aisle, rack from racks where id = " & ds!rackno
        Set ss = Wdb.Execute(sSql)
        If ss.BOF = False Then
            ss.MoveFirst
            move_task_rack.source = ss!aisle & "-" & ss!rack
        End If
        ss.Close
        move_task_rack.target = ptarget
        move_task_rack.product = ds!sku & " " & sku_info(ds!sku, "desc")
        move_task_rack.palletid = ds!barcode
        move_task_rack.qty = "1"
        move_task_rack.uom = "Pallet"
        move_task_rack.lotnum = ds!lot_num
        move_task_rack.units = ds!count_qty
        move_task_rack.lotnum2 = ds!lot2
        move_task_rack.units2 = ds!qty2
        move_task_rack.status = "PEND"
        move_task_rack.userid = puser
        move_task_rack.trandate = Format(Now, "yyMMdd HH:mm:ss")
        move_task_rack.reqid = " "
    Else
        move_task_rack.area = "FAILED"
        move_task_rack.description = " "
        move_task_rack.source = " "
        move_task_rack.target = " "
        move_task_rack.product = " "
        move_task_rack.palletid = " "
        move_task_rack.qty = " "
        move_task_rack.uom = " "
        move_task_rack.lotnum = " "
        move_task_rack.units = " "
        move_task_rack.lotnum2 = " "
        move_task_rack.units2 = " "
        move_task_rack.status = " "
        move_task_rack.userid = " "
        move_task_rack.trandate = " "
        move_task_rack.reqid = " "
    End If
    ds.Close ': db.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "move_task_rack", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "move_task_rack", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: move_task_rack: " & eno) = vbRetry Then
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
    Dim K As Integer, ncnt As Integer
    Dim zid As Long
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
  
    sSql = "Select max(queue_num) from queue_infc"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        K = ds(0) + 1
    Else
        K = 100
    End If
    ds.Close

    sSql = "Select id, queue_num from queue_infc where queue_num = 0"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        zid = ds!id
        sSql = "update queue_infc set queue_num = " & K & _
               "Where id = " & zid
        Wdb.Execute (sSql)
    Else
        zid = wd_seq("Queue_infc")
        sSql = "Insert into queue_infc (ID, Queue_num) VALUES (" & _
               zid & "," & K & ")"
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

Public Sub post_recv_trans(m As ptask)
    'On Error Resume Next
    Dim cfile As String
    If UCase(m.description) = "2 STEP REQUEST" Then Exit Sub
    On Error GoTo vberror
    'tracelist = tracelist & "<!-- post_move_trans(" & m.id & ") -->" & vbCrLf
    If UCase(m.target) = "M-OP" Or UCase(m.target) = "M OP" Then m.target = "ORDER PICK"
    If UCase(m.target) = "M-ANTE" Or UCase(m.target) = "M ANTE" Then m.target = "ANTE ROOM"
    cfile = logdir & "recv" & Format(Now, "MMddyyyy") & ".txt"
    Open cfile For Append As #1
    Write #1, m.id, m.area, m.description, m.source, m.target, m.product;
    Write #1, m.palletid, m.qty, m.uom, m.lotnum, m.units, m.lotnum2, m.units2;
    'Write #1, m.status, m.userid, m.trandate, m.reqid
    Write #1, m.status, WDUserId, m.trandate, m.reqid                   'jv121614
    Close #1
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "post_recv_trans", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "post_recv_trans", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: post_recv_trans: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Sub post_ship_trans(m As ptask)
    'On Error Resume Next
    Dim cfile As String
    'Dim db As ADODB.Connection
    Dim sSql As String
    If UCase(m.description) = "2 STEP REQUEST" Then Exit Sub
    On Error GoTo vberror
    'tracelist = tracelist & "<!-- post_ship_trans(" & m.id & ") -->" & vbCrLf
    If UCase(m.target) = "M-OP" Or UCase(m.target) = "M OP" Then m.target = "ORDER PICK"
    If UCase(m.target) = "M-ANTE" Or UCase(m.target) = "M ANTE" Then m.target = "ANTE ROOM"
    cfile = logdir & "ship" & Format(Now, "MMddyyyy") & ".txt"
    Open cfile For Append As #1
    Write #1, m.id, m.area, m.description, m.source, UCase(m.target), m.product;        'jv122314
    Write #1, m.palletid, m.qty, m.uom, m.lotnum, m.units, m.lotnum2, m.units2;
    'Write #1, m.status, m.userid, m.trandate, m.reqid
    Write #1, m.status, WDUserId, m.trandate, m.reqid                   'jv121614
    Close #1
    If srflag = True Or TCarFlag = True Then
        'Set db = CreateObject("ADODB.Connection")
        'db.Open WDbbsr
        sSql = "Update pallets Set source = '" & m.source & "', target = '" & UCase(m.target) & "'"     'jv122314
        sSql = sSql & ", status = 'Shipped', trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
        sSql = sSql & " Where barcode = '" & m.palletid & "'"
        'MsgBox sSql
        Wdb.Execute (sSql)
        'db.Close
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "post_ship_trans", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "post_ship_trans", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: post_ship_trans: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Function pallet_history_text(bc As String) As String
    Dim ds As ADODB.Recordset, ss As ADODB.Recordset
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String, wsku As String, wlot As String
    Dim wzone As String, wstat As String, wgma As Integer, wside As String
    Dim waisle As String, wrack As String
    Dim cfile As String, s As String, srflag As Boolean
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String, f16 As String
    Dim pht As String, t As String, wqty As Integer
    Dim syear As Integer, eyear As Integer, i As Integer                        'jv061215
    Dim logpath As String
    logpath = logdir
    pht = "": bc = UCase(bc)
    sdate = Format(Val(Mid(bc, 9, 2)) - 2, "00")
    sdate = "20" & sdate & Mid(bc, 5, 4)
    edate = Format(Now, "yyyymmdd")
    wsku = Trim(Left(bc, 4))
    wlot = barcode_to_lotnum(bc)
    If wlot = "01001" Then
        pallet_history_text = "Error!!  Invalid BarCode:  " & bc
        Exit Function
    End If
    Screen.MousePointer = 11
    wqty = Val(sku_info(wsku, "units")) / Val(sku_info(wsku, "wraps"))

    
    'Current location
    'If Form1.plantno = "50" Then                'Search Cranes
        s = "select * from position where barcode = '" & bc & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                wzone = "0": wstat = " ": wgma = 0: wside = " "
                s = "select zone_num, rack_side, lane_status, gmasize from lane where id = " & ds!laneno
                Set ss = Wdb.Execute(s)
                If ss.BOF = False Then
                    ss.MoveFirst
                    wzone = ss!zone_num
                    wstat = ss!lane_status
                    wgma = ss!gmasize
                    wside = ss!rack_side
                End If
                
                pht = pht & "Crane Location: SR-" & ds!whse_num & " "  '& ds!vert_loc & "-" & ds!horz_loc & "-" & ds!rack_side
                
                If ds!whse_num < 4 Then                         'Target
                    pht = pht & ds!vert_loc & "-" & ds!horz_loc & "-" & ds!rack_side & " " & ds!posn_num
                Else
                    pht = pht & wzone & "> " & ds!vert_loc & "-" & ds!horz_loc & "-" & wside
                End If
                If wstat = "H" Then pht = pht & " On Hold"
                If wstat = "B" Then pht = pht & " Blocked"
                ss.Close
                pht = pht & " " & Format(ds!count_qty + ds!qty2, "0") & " units, "
                pht = pht & Format((ds!count_qty + ds!qty2) / wqty, "0") & " wraps"
                pht = pht & vbCrLf & vbCrLf
                ds.MoveNext
            Loop
        End If
        ds.Close
    'End If
    
    s = "select * from rackpos where barcode = '" & bc & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            waisle = " ": wstat = " ": wrack = " "
            s = "select aisle, rack, hold from racks where id = " & ds!rackno
            Set ss = Wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                waisle = Trim(ss!aisle)
                wrack = Trim(ss!rack)
                If ss!hold = 0 Then
                    wstat = " "
                Else
                    wstat = "On Hold"
                End If
            End If
            pht = pht & "Rack Location:  " & waisle & "-" & wrack & " " & wstat
            pht = pht & " " & Format(ds!count_qty + ds!qty2, "0") & " units, "
            pht = pht & Format((ds!count_qty + ds!qty2) / wqty, "0") & " wraps"
            pht = pht & vbCrLf & vbCrLf
            ss.Close
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    syear = Val(Left(sdate, 4))                                             'jv061215
    eyear = Val(Left(edate, 4))                                             'jv061215
    s = ""
    spath = logpath & "recv*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Right(sdir, 12)                                                         'jv061215
        fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = bc And Val(f10) > 0 Then
                    s = ""
                    s = s & "Product:        " & f5 & vbCrLf & vbCrLf
                    s = s & "Label:          " & bc & vbCrLf & vbCrLf
                    If Len(f16) = 6 Then s = s & "Plate:          " & f16 & vbCrLf & vbCrLf
                    If Val(f12) > 0 Then
                        s = s & "Lot1 (" & f9 & " " & Mid(bc, 12, 1) & "):  " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                        s = s & "Lot2 (" & f11 & "):  " & f12 & " units, " & Format(Val(f12) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                    Else
                        s = s & "Quantity:       " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                    End If
                    pht = pht & s
                    s = ""
                    t = Mid(f15, 3, 2) & "-" & Mid(f15, 5, 2) & "-" & Mid(f15, 1, 2) & " "
                    t = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                    s = s & "Wrapped:        " & f3 & " "
                    If Len(s) < 50 Then s = s & Space(50 - Len(s))
                    's = s & t & " " & f14 & vbCrLf & vbCrLf
                    s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    pht = pht & s
    
    For i = syear To eyear
        s = ""
        spath = logpath & Format(i, "0000") & "\recv*.txt"                                  'jv061215
        sdir = Dir$(spath)
        Do While sdir <> ""
            fdate = Right(sdir, 12)                                                         'jv061215
            fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1        'jv061215
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = bc And Val(f10) > 0 Then
                        s = ""
                        s = s & "Product:        " & f5 & vbCrLf & vbCrLf
                        s = s & "Label:          " & bc & vbCrLf & vbCrLf
                        If Len(f16) = 6 Then s = s & "Plate:          " & f16 & vbCrLf & vbCrLf
                        If Val(f12) > 0 Then
                            s = s & "Lot1 (" & f9 & " " & Mid(bc, 12, 1) & "):  " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                            s = s & "Lot2 (" & f11 & "):  " & f12 & " units, " & Format(Val(f12) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                        Else
                            s = s & "Quantity:       " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                        End If
                        pht = pht & s
                        s = ""
                        t = Mid(f15, 3, 2) & "-" & Mid(f15, 5, 2) & "-" & Mid(f15, 1, 2) & " "
                        t = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                        s = s & "Wrapped:        " & f3 & " "
                        If Len(s) < 50 Then s = s & Space(50 - Len(s))
                        's = s & t & " " & f14 & vbCrLf & vbCrLf
                        s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
        pht = pht & s
    Next i
    
    s = ""
    spath = logpath & "tml*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Right(sdir, 12)                                                         'jv061215
        fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = bc Then
                    s = ""
                    t = Mid(f15, 3, 2) & "-" & Mid(f15, 5, 2) & "-" & Mid(f15, 1, 2) & " "
                    t = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                    s = "Traffic Master: " & f4 & " "
                    If Len(s) < 50 Then s = s & Space(50 - Len(s))
                    's = s & t & " " & f14 & vbCrLf & vbCrLf
                    s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    pht = pht & s
    For i = syear To eyear
        s = ""
        spath = logpath & Format(i, "0000") & "\tml*.txt"                                   'jv061215
        sdir = Dir$(spath)
        Do While sdir <> ""
            fdate = Right(sdir, 12)                                                         'jv061215
            fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1        'jv061215
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = bc Then
                        s = ""
                        t = Mid(f15, 3, 2) & "-" & Mid(f15, 5, 2) & "-" & Mid(f15, 1, 2) & " "
                        t = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                        s = "Traffic Master: " & f4 & " "
                        If Len(s) < 50 Then s = s & Space(50 - Len(s))
                        's = s & t & " " & f14 & vbCrLf & vbCrLf
                        s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
        pht = pht & s
    Next i
    
    spath = logpath & "move*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Right(sdir, 12)                                                         'jv061215
        fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = bc Then
                    If Len(pht) = 0 Then
                        s = s & "Product:        " & f5 & vbCrLf & vbCrLf
                        s = s & "Label:          " & bc & vbCrLf & vbCrLf
                        If Len(f16) = 6 Then s = s & "Plate:          " & f16 & vbCrLf & vbCrLf
                        
                        If Val(f12) > 0 Then
                            s = s & "Lot1 (" & f9 & " " & Mid(bc, 12, 1) & "):  " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                            s = s & "Lot2 (" & f11 & "):  " & f12 & " units, " & Format(Val(f12) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                        Else
                            s = s & "Quantity:       " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                        End If
                        pht = pht & s
                    End If
                    s = ""
                    t = Mid(f15, 3, 2) & "-" & Mid(f15, 5, 2) & "-" & Mid(f15, 1, 2) & " "
                    t = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                    s = "Moved:          " & Trim(f3) & " to " & Trim(f4) & " " '& t & " " & f14 & vbCrLf & vbCrLf
                    If Len(s) < 50 Then s = s & Space(50 - Len(s))
                    's = s & t & " " & f14 & vbCrLf & vbCrLf
                    s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                    pht = pht & s
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\move*.txt"                                  'jv061215
        sdir = Dir$(spath)
        Do While sdir <> ""
            fdate = Right(sdir, 12)                                                         'jv061215
            fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1        'jv061215
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = bc Then
                        If Len(pht) = 0 Then
                            s = s & "Product:        " & f5 & vbCrLf & vbCrLf
                            s = s & "Label:          " & bc & vbCrLf & vbCrLf
                            If Len(f16) = 6 Then s = s & "Plate:          " & f16 & vbCrLf & vbCrLf
                            If Val(f12) > 0 Then
                                s = s & "Lot1 (" & f9 & " " & Mid(bc, 12, 1) & "):  " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                                s = s & "Lot2 (" & f11 & "):  " & f12 & " units, " & Format(Val(f12) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                            Else
                                s = s & "Quantity:       " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                            End If
                            pht = pht & s
                        End If
                        s = ""
                        t = Mid(f15, 3, 2) & "-" & Mid(f15, 5, 2) & "-" & Mid(f15, 1, 2) & " "
                        t = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                        s = "Moved:          " & Trim(f3) & " to " & Trim(f4) & " " '& t & " " & f14 & vbCrLf & vbCrLf
                        If Len(s) < 50 Then s = s & Space(50 - Len(s))
                        's = s & t & " " & f14 & vbCrLf & vbCrLf
                        s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                        pht = pht & s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    spath = logpath & "ship*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Right(sdir, 12)                                                         'jv061215
        fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = bc Then
                    If Len(pht) = 0 Then
                        s = s & "Product:        " & f5 & vbCrLf & vbCrLf
                        s = s & "Label:          " & bc & vbCrLf & vbCrLf
                        If Len(f16) = 6 Then s = s & "Plate:          " & f16 & vbCrLf & vbCrLf
                        If Val(f12) > 0 Then
                            s = s & "Lot1 (" & f9 & " " & Mid(bc, 12, 1) & "):  " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                            s = s & "Lot2 (" & f11 & "):  " & f12 & " units, " & Format(Val(f12) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                        Else
                            s = s & "Quantity:       " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                        End If
                        pht = pht & s
                    End If
                    s = ""
                    t = Mid(f15, 3, 2) & "-" & Mid(f15, 5, 2) & "-" & Mid(f15, 1, 2) & " "
                    t = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                    s = "Shipped:        " & f3 & " " & f2 & " " & f4 & " "
                    If Len(s) < 50 Then s = s & Space(50 - Len(s))
                    's = s & t & " " & f14 & vbCrLf & vbCrLf
                    s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                    pht = pht & s
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\ship*.txt"                                  'jv061215
        sdir = Dir$(spath)
        Do While sdir <> ""
            fdate = Right(sdir, 12)                                                         'jv061215
            fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1        'jv061215
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = bc Then
                        If Len(pht) = 0 Then
                            s = s & "Product:        " & f5 & vbCrLf & vbCrLf
                            s = s & "Label:          " & bc & vbCrLf & vbCrLf
                            If Len(f16) = 6 Then s = s & "Plate:          " & f16 & vbCrLf & vbCrLf
                            If Val(f12) > 0 Then
                                s = s & "Lot1 (" & f9 & " " & Mid(bc, 12, 1) & "):  " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                                s = s & "Lot2 (" & f11 & "):  " & f12 & " units, " & Format(Val(f12) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                            Else
                                s = s & "Quantity:       " & f10 & " units, " & Format(Val(f10) / wqty, "0") & " wraps" & vbCrLf & vbCrLf
                            End If
                            pht = pht & s
                        End If
                        s = ""
                        t = Mid(f15, 3, 2) & "-" & Mid(f15, 5, 2) & "-" & Mid(f15, 1, 2) & " "
                        t = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                        s = "Shipped:        " & f3 & " " & f2 & " " & f4 & " "
                        If Len(s) < 50 Then s = s & Space(50 - Len(s))
                        's = s & t & " " & f14 & vbCrLf & vbCrLf
                        s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                        pht = pht & s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    pallet_history_text = pht
    Screen.MousePointer = 0
End Function

Public Sub pallet_lots(bc As String)
    Dim ds As ADODB.Recordset
    Dim sSql As String
    'tracelist = tracelist & "<!-- pallet_lots(" & bc & ") -->" & vbCrLf
    On Error GoTo vberror
    sSql = "select qty1,lot1,qty2,lot2,plateno from pallets where barcode = '" & bc & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        ship_units = ds!qty1
        ship_lotnum = ds!lot1
        ship_units2 = ds!qty2
        ship_lotnum2 = ds!lot2
        ship_plate = ds!plateno
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "post_sr4_remove", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "pallet_lots", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: pallet_lots: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub
Public Sub post_sr4_remove(m As ptask, rknote As String)
    'On Error Resume Next
    Dim cfile As String
    If UCase(m.description) = "2 STEP REQUEST" Then Exit Sub
    On Error GoTo vberror
    'tracelist = tracelist & "<!-- post_sr4_remove(" & m.id & "," & rknote & ") -->" & vbCrLf
    If UCase(m.target) = "M-OP" Or UCase(m.target) = "M OP" Then m.target = "ORDER PICK"
    If UCase(m.target) = "M-ANTE" Or UCase(m.target) = "M ANTE" Then m.target = "ANTE ROOM"
    cfile = logdir & "move" & Format(Now, "MMddyyyy") & ".txt"
    Open cfile For Append As #1
    Write #1, m.id, m.area, rknote, m.source, m.target, m.product;                  'jv111314
    Write #1, m.palletid, m.qty, m.uom, m.lotnum, m.units, m.lotnum2, m.units2;
    'Write #1, m.status, m.userid, m.trandate, m.reqid
    Write #1, m.status, WDUserId, m.trandate, m.reqid                   'jv121614
    Close #1
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "post_sr4_remove", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "post_sr4_remove", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: post_sr4_remove: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Function remove_rack_pallet(m As ptask) As Boolean
    'On Error Resume Next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    Dim sSql As String, sCols As String, sRows As String
    Dim olot As String, rlot As String
    Dim K As Integer
    Dim pqty As Integer
    Dim pqty4 As Integer
    Dim rknote As String
    Dim s As String
    Dim zid As Long
    Dim ncnt As Integer, i As Integer
    Dim ma As String, ms As String, mt As String                                            'jv111314
    'tracelist = tracelist & "<!-- remove_rack_pallet(" & m.palletid & ") -->" & vbCrLf
    On Error GoTo vberror
    If m.source = "BACKHAUL" Or m.source = "SR1" Or m.source = "SR2" Or m.source = "SR3" Or m.source = "STAGING" Then
        remove_rack_pallet = True
        Exit Function
    End If
    If m.source = "ALT" Or m.source = "ROBOT ZERO" Or m.source = "WRAPPER" Or m.source = "ROLLER BED" Or m.source = "TRI LEVEL" Or m.source = "TRI-LEVEL" Then
        remove_rack_pallet = True
        Exit Function
    End If
    
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    zid = 0
    s = Mid(m.palletid, 1, 11) & "_" & Mid(m.palletid, 13, 4)
    sSql = "Select * From rackpos Where barcode in ('" & m.palletid & "','" & s & "')"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        zid = ds!id
        K = ds!rackno
        rknote = ds!barcode
        m.lotnum = ds!lot_num
        m.units = ds!count_qty
        s = ds!lot2
        If Len(s) > 0 Then
            m.lotnum2 = s
        Else
            m.lotnum2 = " "
        End If
        s = ds!qty2
        If Len(s) > 0 Then
            m.units2 = s
        Else
        m.units2 = "0"
        End If
    End If
    ds.Close
    'Call debug_log(sSql, m, m.userid)
    If zid > 0 Then
        sSql = "Update rackpos Set sku=' '"
        sSql = sSql & ",lot_num=' '"
        sSql = sSql & ",pallet_num=' '"
        sSql = sSql & ",count_qty=0"
        sSql = sSql & ",recv_date='" & Format(Now, "MM-dd-yyyy") & "'"
        sSql = sSql & ",bbc='Y'"
        sSql = sSql & ",barcode=' '"
        sSql = sSql & ",lot2=' '"
        sSql = sSql & ",qty2=0"
        sSql = sSql & ",wrapped='Y'"
        sSql = sSql & ",hold='N'"
        sSql = sSql & " Where id = " & zid
        Wdb.Execute (sSql)
        'Call debug_log(sSql, m, m.userid)
    Else
        m.lotnum = barcode_to_lotnum(m.palletid)
        Call update_trans(m)
    End If
    m.trandate = Format(Now, "yyMMdd HH:mm:ss")

    If K > 0 Then
        sSql = "Select aisle,rack,hold From racks Where id = " & K              'jv111314
        Set ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst
            rknote = rknote & " @ "
            rknote = rknote & ds!aisle
            rknote = rknote & "-"
            rknote = rknote & ds!rack
            If ds!hold = 1 Then                                                 'jv111314
                ma = m.area                                                     'jv111314
                ms = m.source                                                   'jv111314
                mt = m.target                                                   'jv111314
                m.area = "HOLD"                                                 'jv111314
                If m.target = "ORDER PICK" Or m.target = "M-OP" Then            'jv111314
                    m.target = m.source                                         'jv111314
                End If                                                          'jv111314
                m.source = "HOLD"                                               'jv111314
                Call post_move_trans(m)                                         'jv111314
                m.area = ma                                                     'jv111314
                m.source = ms                                                   'jv111314
                m.target = mt                                                   'jv111314
            End If                                                              'jv111314
        Else
            rknote = rknote & " " & m.source & " not removed."
        End If
        ds.Close
    End If
    'Call post_sr4_remove(m, rknote)                                            'jv111314
    'Call debug_log(sSql, m, m.userid)

    If K = 0 Then
        'db.Close
        If m.source = "BACKHAUL" Or m.source = "SR1" Or m.source = "SR2" Or m.source = "SR3" Or m.source = "STAGING" Then
            remove_rack_pallet = True
            Exit Function
        End If
        If m.source = "ALT" Or m.source = "ROBOT ZERO" Or m.source = "WRAPPER" Or m.source = "ROLLERBED" Or m.source = "TRI LEVEL" Or m.source = "TRI-LEVEL" Then
            remove_rack_pallet = True
            Exit Function
        End If
        s = m.palletid & " barcode was not found in rack inventory, remove rack pallet."
        Call debug_log(s, m, m.userid)
        remove_rack_pallet = False
        Exit Function
    End If
    Call update_trans(m)

    olot = "99999"
    pqty = 0
    pqty4 = 0

    sSql = "Select sku,lot_num,lot2,bbc From rackpos Where rackno = " & K & _
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
    End If
    ds.Close
    'Call debug_log(sSql, m, m.userid)

    If pqty + pqty4 = 0 Then
        sSql = "Update racks set sku = ' ',lot_num = ' ',qty = 0,qty4 = 0"
    Else
        sSql = "Update racks set lot_num = '" & olot & "',"
        sSql = sSql & "qty = " & pqty & ","
        sSql = sSql & "qty4 = " & pqty4
    End If
    sSql = sSql & " Where id = " & K
    Wdb.Execute (sSql)
    'db.Close
    'Call debug_log(sSql, m, m.userid)
    remove_rack_pallet = True
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "remove_rack_pallet", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "remove_rack_pallet", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: remove_rack_pallet: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Sub remove_sp_order(m As ptask)
    'On Error Resume Next
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    Dim sSql As String, sCols As String, sRows As String
    Dim zid As Long
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
  
    sSql = "Select id From paltasks Where area = 'SNACK PLANT WRAPPER'"
    sSql = sSql & " And source = 'M-SP'"
    sSql = sSql & " And status = 'PEND'"
    sSql = sSql & " And product > '" & Mid(m.palletid, 1, 4) & "'"
    sSql = sSql & " And product < '" & Mid(m.palletid, 1, 4) & "ZZZZ'"
    
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        zid = ds!id
        sSql = "Update paltasks Set palletid = '" & m.palletid & "'"
        sSql = sSql & ",status = 'COMP"
        sSql = sSql & ",userid = '" & m.userid & "'"
        sSql = sSql & " Where id = " & zid
        Wdb.Execute (sSql)
    End If
    ds.Close ': db.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "remove_sp_order", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "remove_sp_order", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: remove_sp_order: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Function req_door_barcode(pdesc As String, ptarget As String) As String
    Dim ds As ADODB.Recordset, s As String
    s = "select target from paltasks where area = 'GROUP'"
    s = s & " and product = '" & pdesc & "   " & ptarget & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        req_door_barcode = ds!target
    Else
        req_door_barcode = "NODOOR"
    End If
    ds.Close
End Function

Public Function return_to_wrapper(ubar As String, uname As String, uarea As String, ureq As String) As String
    'On Error Resume Next
    'Dim db As ADODB.Connection
    Dim ds As ADODB.Recordset
    Dim sSql As String, sCols As String, sRows As String
    Dim p As ptask, s As String, i As Long, n As Integer, K As Integer
    If Len(ubar) < 16 Then
        return_to_wrapper = "Illegal Barcode: " & ubar * "!!"
        Exit Function
    End If
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    'tracelist = tracelist & "<!-- return_to_wrapper(" & ubar & "," & uname & "," & uarea & "," & ureq & ") -->" & vbCrLf
    p.id = 0
    p.area = uarea
    p.description = " "
    p.source = " "
    p.target = "WRAPPER"
    s = Trim(Mid(ubar, 1, 4))
    p.product = s & sku_info(s, "desc")
    p.palletid = ubar
    p.qty = "-1"
    p.uom = "Pallet"
    p.lotnum = ".."
    p.units = "0"
    p.lotnum2 = ".."
    p.units2 = "0"
    p.status = "COMP"
    p.userid = uname
    p.trandate = Format(Now, "yyMMdd HH:mm:ss")
    p.reqid = "0"
    sSql = "Select id From paltasks Where palletid = '" & ubar & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        i = ds!id
        p = masterec(i)
        p.target = "WRAPPER"
        p.qty = Format(Val(p.qty) * -1, "0")
        p.units = Format(Val(p.units) * -1, "0")
        p.units2 = Format(Val(p.units2) * -1, "0")
        p.status = "COMP"
        p.userid = uname
        p.trandate = Format(Now, "yyMMdd HH:mm:ss")
        Do Until ds.EOF
            i = ds!id
            sSql = "Update paltasks Set status = 'COMP', palletid = '...'"
            sSql = sSql & " Where id = " & i
            Wdb.Execute (sSql)
            ds.MoveNext
        Loop
    End If
    ds.Close

    sSql = "Select * From queue_infc Where palletid = '" & ubar & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        i = ds!id
        p.id = ds!id
        p.area = uarea
        p.description = " "
        p.source = "SR-" & ds!whse_num
        p.target = "WRAPPER"
        s = ds!sku
        p.product = s & " " & sku_info(s, "desc")
        p.palletid = ubar
        p.qty = "-1"
        p.uom = "Pallet"
        p.lotnum = ds!lot_num
        p.units = Format(Val(ds!units) * -1, "0")
        p.lotnum2 = ds!lot_num2
        p.units2 = Format(Val(ds!units2) * -1, "0")
        p.status = "COMP"
        p.userid = uname
        p.trandate = Format(Now, "yyMMdd HH:mm:ss")
        p.reqid = "0"
        Do Until ds.EOF
            i = ds!id
            sSql = "Update queue_infc Set queue_num = 0"
            sSql = sSql & " Where id = " & i
            Wdb.Execute (sSql)
            ds.MoveNext
        Loop
    End If
    ds.Close

    sSql = "Select * From rackpos Where barcode = '" & ubar & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        If p.id = 0 Then
            p.id = ds!id
            p.area = uarea
            p.description = " "
            p.source = "RACKS"
            p.target = "WRAPPER"
            s = ds!sku
            p.product = s & " " & sku_info(s, "desc")
            p.palletid = ubar
            p.qty = "-1"
            p.uom = "Pallet"
            p.lotnum = ds!lot_num
            p.units = Format(Val(ds!count_qty) * -1, "0")
            p.lotnum2 = ds!lot2
            p.units2 = Format(Val(ds!qty2) * -1, "0")
            p.status = "COMP"
            p.userid = uname
            p.trandate = Format(Now, "yyMMdd HH:mm:ss")
            p.reqid = "0"
        End If
        p.source = "RACKS"
        Call remove_rack_pallet(p)
        If Val(p.units) > 0 Then p.units = Format(Val(p.units) * -1, "0")
        If Val(p.units2) > 0 Then p.units2 = Format(Val(p.units2) * -1, "0")
    End If
    ds.Close
    
    sSql = "select id from pallets where barcode = '" & ubar & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sSql = "Update pallets set plateno = '0', barcode = '..', status = 'Shipped'"
            sSql = sSql & ", sku = ' '"
            sSql = sSql & " Where id = " & ds!id
            Wdb.Execute (sSql)
            ds.MoveNext
        Loop
    End If
    ds.Close ': db.Close
    
    If p.id > 0 Or uarea = "EFLMove" Then
        Call post_recv_trans(p)
        return_to_wrapper = ubar & " returned."
    Else
        return_to_wrapper = ubar & " was not found."
    End If
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "return_to_wrapper", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "return_to_wrapper", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: return_to_wrapper: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Sub robot0_pickup(m As ptask)
    Dim sSql As String
    Dim p As ptask, i As Long
    Dim s As String, t As String
    Dim p4way As Boolean
    On Error GoTo vberror
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    t = "ORDER PICK"
    p4way = False
    'tracelist = tracelist & "<!-- robot0_pickup() -->" & vbCrLf
    sSql = "Select palletid,source From paltasks Where source in ('ROBOT ZERO','WRAPPER')"
    sSql = sSql & " And area = 'FORKLIFT'"
    sSql = sSql & " And palletid = '" & m.palletid & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        s = ds!palletid
        s = s & " already at "
        s = s & ds!source
        s = s & " Pick Up."
        ds.Close ': db.Close
        Call debug_log(s, m, m.userid)
        Exit Sub
    End If
    ds.Close
    m.status = "COMP"
    m.trandate = Format(Now, "yyMMdd HH:mm:ss")
    Call post_recv_trans(m)

    s = Trim(Mid(m.product, 1, 4))
    If Val(m.qty) > Val(sku_info(s, "wraps")) Then p4way = True
    sSql = "Select aisle,rack From racks Where resv_sku = '" & s & "'"
    sSql = sSql & " And (qty + qty4) < capacity"
    If p4way = True Then
        sSql = sSql & " And qty = 0"
        sSql = sSql & " Order By resv_sku,qty4 Desc"
    Else
        sSql = sSql & " And qty4 = 0"
        sSql = sSql & " Order By resv_sku,qty Desc"
    End If
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        t = Trim(ds!aisle) & "-" & Trim(ds!rack)
    End If
    ds.Close ': db.Close
    p.area = "FORKLIFT"
    p.description = " "
    p.source = m.target
    p.target = t
    s = wrapdesc(Trim(Mid(m.product, 1, 4)), Val(m.units) + Val(m.units2))
    If Len(s) > 1 Then
        p.product = Trim(Mid(m.product, 1, 4)) & s & " " & StrConv(Mid(m.product, 5, Len(m.product) - 4), vbProperCase)
    Else
        p.product = m.product
    End If
    p.palletid = m.palletid
    p.qty = m.qty
    p.uom = m.uom
    p.lotnum = m.lotnum
    p.units = m.units
    p.lotnum2 = m.lotnum2
    p.units2 = m.units2
    p.status = "PEND"
    p.userid = " "
    p.trandate = Format(Now, "yyMMdd HH:mm:ss")
    'p.reqid = " "
    p.reqid = m.reqid
    i = insert_trans(p)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "robot0_pickup", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "robot0_pickup", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: robot0_pickup: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Sub roller_bed_pickup(m As ptask)
    Dim p As ptask
    Dim s As String
    tracelist = tracelist & "<!-- roller_bed_pickup(" & m.id & ") -->" & vbCrLf
    p.area = "FORKLIFT"
    p.description = " "
    p.source = m.source
    p.target = m.target
    s = wrapdesc(Trim(Mid(m.product, 1, 4)), Val(m.units) + Val(m.units2))
    If Len(s) > 1 Then
        p.product = Trim(Mid(m.product, 1, 4)) & s & " " & StrConv(Mid(m.product, 5, Len(m.product) - 4), vbProperCase)
    Else
        p.product = m.product
    End If
    p.palletid = m.palletid
    p.qty = m.qty
    p.uom = m.uom
    p.lotnum = m.lotnum
    p.units = m.units
    p.lotnum2 = m.lotnum2
    p.units2 = m.units2
    p.status = m.status
    p.userid = m.userid
    p.trandate = Format(Now, "yyMMdd HH:mm:ss")
    p.reqid = m.reqid
    If insert_rack_pallet(p) = False Then Exit Sub
    Call post_recv_trans(p)
End Sub

Public Sub Set_WD_Org(orgcode As String)
    If orgcode = "500" Then
        logdir = "\\bbc-01-wdmgmt\wd\data\testlog"
        tracelist = " "
        debflag = False
        SPTarget = "SNACK PLANT"
        BHDest = "STAGING"
        srflag = True
        ARFlag = True
        TCarFlag = False
        'WDOrg = gsOrgCode
    End If

    If orgcode = "501" Then
        logdir = "\\bbba-02-dc\f\user\waredist\data\pallogs"
        tracelist = " "
        debflag = False
        SPTarget = "SNACK PLANT"
        BHDest = "STAGING"
        srflag = False
        ARFlag = False
        TCarFlag = False
        'WDOrg = gsOrgCode
    End If

    If orgcode = "502" Then
        logdir = "\\bbsy-02-dc\f\user\waredist\data\pallogs"""
        tracelist = " "
        debflag = False
        SPTarget = "SNACK PLANT"
        BHDest = "STAGING"
        srflag = True
        ARFlag = True
        TCarFlag = False
        'WDOrg = gsOrgCode
    End If

End Sub

Public Function sku_info(psku As String, pfld As String) As String
    'On Error Resume Next
    Dim sSql As String
    Dim i As Integer
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset
    On Error GoTo vberror
    i = Val(psku)
    'If i > 1000 Then i = 1
    If i > 9999 Then i = 1                                      'jv082415

    If skutab(i, 0) = psku Then
    Else
        'Set db = CreateObject("ADODB.Connection")
        'db.Open WDbbsr
 
        sSql = "select sku, uom_type, description, uom_per_pallet, qty_per_pallet" & _
               " from sku_config where sku = '" & psku & "'"

        Set ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst
            skutab(i, 0) = ds!sku
            skutab(i, 1) = ds!uom_type & " "
            skutab(i, 1) = skutab(i, 1) & ds!description
            skutab(i, 2) = ds!uom_per_pallet
            If Val(skutab(i, 2)) < 1 Then skutab(i, 2) = "1"
            skutab(i, 3) = ds!qty_per_pallet
            If Val(skutab(i, 3)) < 1 Then skutab(i, 3) = "1"
        Else
            skutab(i, 0) = "0"
            skutab(i, 1) = "unrecognized SKU"
            skutab(i, 2) = "1"
            skutab(i, 3) = "1"
        End If
        ds.Close ': db.Close
    End If

    sku_info = skutab(i, 0)
    If LCase(pfld) = "wraps" Then sku_info = skutab(i, 3)
    If LCase(pfld) = "units" Then sku_info = skutab(i, 2)
    If LCase(pfld) = "desc" Then sku_info = skutab(i, 1)
    'tracelist = tracelist & "<!-- sku_info(" & psku & "," & pfld & ") = " & sku_info & " -->" & vbCrLf
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "sku_info", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "sku_info", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: sku_info: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Function space_in_rack(maisle As String, mrack As String, msku As String, m As ptask) As Boolean
    'On Error Resume Next
    Dim sSql As String
    'Dim db As ADODB.Connection
    Dim ds As ADODB.Recordset
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    'tracelist = tracelist & "<!-- space_in_rack(" & maisle & "," & mrack & "," & msku & ") -->" & vbCrLf
    space_in_rack = False
    sSql = "Select rackno from rackpos where sku <= '000'" & _
            " And rackno in (select id from racks where aisle = '" & maisle & "'" & _
            " And rack = '" & mrack & "')"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then space_in_rack = True
    ds.Close ': db.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "space_in_rack", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "space_in_rack", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: space_in_rack: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Sub spt_to_dock(m As ptask)
    'On Error Resume Next
    Dim sSql As String
    Dim p As ptask, i As Long
    Dim s As String, t As String
    'Dim db As ADODB.Connection
    Dim ds As ADODB.Recordset
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    'tracelist = tracelist & "<!-- spt_to_dock() -->" & vbCrLf
    t = "ORDER PICK"
    SPTarget = m.target
    sSql = "Select palletid From paltasks Where area = 'DOCK' And status = 'PEND'"
    sSql = sSql & " And palletid = '" & m.palletid & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        s = m.palletid & " aleady at DOCK."
        Call debug_log(s, m, m.userid)
        ds.Close ': db.Close
        Exit Sub
    End If
    ds.Close
    If m.source = "SNACK PLANT WRAPPER" Then
        m.status = "COMP"
        m.trandate = Format(Now, "yyMMdd HH:mm:ss")
        Call post_recv_trans(m)
    End If
    If m.target = "SNACK PLANT" Then
        If insert_rack_pallet(m) = True Then
            Call debug_log("staying at SP.", m, "0")
        End If
    Else
        p.area = "DOCK"
        p.description = " "
        If m.target = "1405" Or m.target = "1406" Or m.target = "1731" Then
            p.source = m.target
        Else
            p.source = "SNACK PLANT"
        End If
        If m.source = "SNACK PLANT WRAPPER" Then
            p.target = "CRANE 5"
        Else
            'p.target = "ANTE ROOM"
            p.target = "CRANE 5"
            sSql = "Update paltasks Set status = 'COMP' Where id = " & m.id
            Wdb.Execute (sSql)
        End If
        s = wrapdesc(Trim(Mid(m.product, 1, 4)), Val(m.units) + Val(m.units2))
        If Len(s) > 1 Then
            p.product = Trim(Mid(m.product, 1, 4)) & s & " " & StrConv(Mid(m.product, 5, Len(m.product) - 4), vbProperCase)
        Else
            p.product = m.product
        End If
        p.palletid = m.palletid
        If Val(m.units2) > 0 Then       'jv091113
            p.qty = "1"
            p.uom = "Pallet"
        Else
            p.qty = m.qty
            p.uom = m.uom
        End If
        p.lotnum = m.lotnum
        p.units = m.units
        p.lotnum2 = m.lotnum2
        p.units2 = m.units2
        p.status = "PEND"
        p.userid = " "
        p.trandate = Format(Now, "yyMMdd HH:mm:ss")
        'p.reqid = " "
        p.reqid = m.reqid
        i = insert_trans(p)
        Call remove_sp_order(p)
    End If
    'db.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "spt_to_dock", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "spt_to_dock", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: spt_to_dock: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Sub spt_to_group(m As ptask)
    'On Error Resume Next
    Dim sSql As String, sCols As String, sRows As String
    Dim zid As Long
    'Dim db As ADODB.Connection
    Dim ds As ADODB.Recordset
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    'tracelist = tracelist & "<!-- spt_to_group(" & m.id & "," & m.target & "," & m.product & ") -->" & vbCrLf
    sSql = "Select id,description From paltasks Where area = 'DOCK'"
    sSql = sSql & " And description > '0 '"
    sSql = sSql & " And target = '" & m.target & "'"
    sSql = sSql & " And product >= '" & Trim(Mid(m.product, 1, 4)) & "'"
    sSql = sSql & " And product < '" & Trim(Mid(m.product, 1, 4)) & "ZZZZ'"
    sSql = sSql & " And source <> 'ALT'"
    sSql = sSql & " And status = 'PEND'"
    sSql = sSql & " And userid < '0'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        zid = ds!id
        m.description = ds!description
        sSql = "Update paltasks Set source = "
        If m.target = "1405" Or m.target = "1406" Or m.target = "1731" Then
            sSql = sSql & "'" & m.target & "'"
        Else
            If m.source = "BACKHAUL" Then
                sSql = sSql & "'BACKHAUL'"
            Else
                sSql = sSql & "'SNACK PLANT'"
            End If
        End If
        sSql = sSql & ",palletid='" & m.palletid & "'"
        sSql = sSql & ",qty=" & Val(m.qty)
        sSql = sSql & ",uom='" & m.uom & "'"
        sSql = sSql & ",lotnum='" & m.lotnum & "'"
        sSql = sSql & ",units=" & Val(m.units)
        sSql = sSql & ",lotnum2='" & m.lotnum2 & "'"
        sSql = sSql & ",units2=" & Val(m.units2)
        sSql = sSql & ",status='COMP'"
        sSql = sSql & ",userid='" & m.userid & "'"
        sSql = sSql & ",trandate='" & Format(Now, "yyMMdd HH:mm:ss") & "'"
        sSql = sSql & " Where id = " & zid
        Wdb.Execute (sSql)
    End If
    ds.Close ': db.Close
    m.status = "COMP"
    m.trandate = Format(Now, "yyMMdd HH:mm:ss")
    Call post_ship_trans(m)
    Call update_trans(m)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "spt_to_group", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "spt_to_group", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: spt_to_group: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Function sr_receiving(sr As String, mbarcode As String) As Boolean
    'On Error Resume Next
    Dim sSql As String, sCols As String, sRows As String
    'Dim db As ADODB.Connection
    Dim ds As ADODB.Recordset, pcode As String
    On Error GoTo vberror
    'pcode = Mid(mbarcode, 12, 1)                                            'jv062614
    pcode = Trim(Mid(mbarcode, 11, 3))                                      'jv052515
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    sSql = "Select id From prodrcv Where sku = '" & Trim(Mid(mbarcode, 1, 4)) & "'"
    sSql = sSql & " And " & sr & " > 0"
    sSql = sSql & " And lot_num = '" & barcode_to_lotnum(mbarcode) & "'"
    sSql = sSql & " And sp_flag in ('" & pcode & "', '0', '1')"             'jv062614
    'tracelist = tracelist & "<!-- sr_receiving(" & sr & "," & mbarcode & ")" & vbCrLf
    'tracelist = tracelist & sSql & " -->" & vbCrLf
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        sr_receiving = True
    Else
        sr_receiving = False
    End If
    ds.Close ': db.Close
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "sr_receiving", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "sr_receiving", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: sr_receving: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Function tag_alternate(psource As String, puser As String) As String
    'On Error Resume Next
    Dim sSql As String, sCols As String, sRows As String
    Dim p As ptask, d As ptask
    Dim nsrc As String, psku As String
    Dim i As Long, s As String
    'Dim db As ADODB.Connection
    Dim ds As ADODB.Recordset, rs As ADODB.Recordset
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    'tracelist = tracelist "(!-- tag_alternate() -->" & vbCrLf
    If psource = "ALTERNATES" Then
        tag_alternate = "<!-- No alternates are defined.. -->"
        Exit Function
    End If
    sSql = "Select * From paltasks Where id = " & Mid(psource, Len(psource) - 4, 5)
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        d.id = ds!id
        d.area = ds!area
        d.description = ds!description
        d.source = ds!source
        d.target = ds!target
        d.product = ds!product
        d.palletid = ds!palletid
        d.qty = ds!qty
        d.uom = ds!uom
        d.lotnum = ds!lotnum
        d.units = ds!units
        d.lotnum2 = ds!lotnum2
        d.units2 = ds!units2
        d.status = ds!status
        d.userid = ds!userid
        d.trandate = ds!trandate
        d.reqid = ds!reqid

        psku = Trim(Mid(d.product, 1, 4))
        sSql = "Select aisle,rack,fo,lot_num From racks"
        sSql = sSql & " Where sku = '" & psku & "'"
        sSql = sSql & " And aisle <> 'M' And resv_sku <> 'ALL'"
        sSql = sSql & " And hold <> 1"
        sSql = sSql & " And id in (Select rackno From rackpos"
        sSql = sSql & " Where sku = '" & psku & "'"
        sSql = sSql & " And count_qty > 0)"
        sSql = sSql & " Order By fo Desc, lot_num"
        Set rs = Wdb.Execute(sSql)
        If rs.BOF = False Then
            rs.MoveFirst
            nsrc = "STAGING"
            p.area = "FORKLIFT"
            p.description = d.description & Space(8 - Len(d.description)) & d.target
            p.source = rs!aisle & "-"
            p.source = psource & Trim(rs!rack)
            p.target = "STAGING"
            p.product = d.product
            p.palletid = d.palletid
            p.qty = d.qty
            p.uom = d.uom
            p.lotnum = d.lotnum
            p.units = d.units
            p.lotnum2 = d.lotnum2
            p.units2 = d.units2
            p.status = d.status
            p.userid = d.userid
            p.trandate = Format(Now, "yyMMdd HH:mm:ss")
            p.reqid = d.reqid
            i = insert_trans(p)
        End If
        rs.Close

        If nsrc = "..." And srflag = True Then
            sSql = "Select whse_num, lot_num From lane"
            sSql = sSql & " Where sku = '" & psku & "'"
            sSql = sSql & " And lane_status <= ' '"
            sSql = sSql & " Order By lot_num, whse_num"
            Set rs = Wdb.Execute(sSql)
            If rs.BOF = False Then
                rs.MoveFirst
                nsrc = "SR" & rs!whse_num
            End If
            rs.Close
'            If nsrc <> "..." Then
                sSql = "Select id, ship_status From ship_infc Where order_num = '" & d.description & "'"
                sSql = sSql & " And sku = '" & psku & "'"
                sSql = sSql & " And to_whse_num = " & Mid(nsrc, 3, 1)
                Set rs = Wdb.Execute(sSql, sCols, sRows)
                If rs.BOF = False Then
                    rs.MoveFirst
                    i = rs!id
                    s = rs!ship_status
                    If s = "DONE" Or s = "CANC" Then
                        sSql = "Update ship_infc Set order_qty = 1, ship_uom_qty = 0"
                        sSql = sSql & ", ship_plt_qty = 0, ship_status = 'NEW'"
                        sSql = sSql & " Where id = " & i
                        Wdb.Execute (sSql)
                    Else
                        sSql = "Update ship_infc Set order_qty = order_qty + 1"
                        sSql = sSql & " Where id = " & i
                        Wdb.Execute (sSql)
                    End If
                Else
                    rs.Close
                    sSql = "Select id, ship_status From ship_infc"
                    sSql = sSql & " Where ship_status in ('CANC','DONE')"
                    sSql = sSql & " Order By id"
                    Set rs = Wdb.Execute(sSql)
                    If rs.BOF = False Then
                        rs.MoveFirst
                        i = rs!id
                        sSql = "Update ship_infc Set order_num = '" & d.description & "'"
                        sSql = sSql & ",sku = '" & psku & "'"
                        sSql = sSql & ",ship_date = '" & Format(Now, "M/d/yyyy") & "'"
                        sSql = sSql & ",order_qty = 1"
                        sSql = sSql & ",ship_uom_qty = 0"
                        sSql = sSql & ",ship_plt_qty = 0"
                        sSql = sSql & ",ship_status = 'NEW"
                        If nsrc = "SR1" Then
                            sSql = sSql & ",to_whse_num = 1"
                            sSql = sSql & ",to_vert_loc = 2"
                            sSql = sSql & ",to_horz_loc = 18"
                            sSql = sSql & ",to_rack_side = 'L'"
                        End If
                        If nsrc = "SR2" Then
                            sSql = sSql & ",to_whse_num = 2"
                            sSql = sSql & ",to_vert_loc = 2"
                            sSql = sSql & ",to_horz_loc = 39"
                            sSql = sSql & ",to_rack_side = 'L'"
                        End If
                        If nsrc = "SR3" Then
                            sSql = sSql & ",to_whse_num = 3"
                            sSql = sSql & ",to_vert_loc = 2"
                            sSql = sSql & ",to_horz_loc = 43"
                            sSql = sSql & ",to_rack_side = 'R'"
                        End If
                        sSql = sSql & " Where id = " & i
                        Wdb.Execute (sSql)
                    End If
                    rs.Close
                End If
'            End If
        End If
        p.area = d.area
        p.description = d.description
        p.source = nsrc
        p.target = d.target
        p.product = d.product
        p.palletid = d.palletid
        p.qty = d.qty
        p.uom = d.uom
        p.lotnum = d.lotnum
        p.units = d.units
        p.lotnum2 = d.lotnum2
        p.units2 = d.units2
        p.status = "PEND"
        p.userid = " "
        p.trandate = Format(Now, "yyMMdd HH:mm:ss")
        'p.reqid = " "
        p.reqid = d.reqid
        i = insert_trans(p)
        s = p.target & " Alternate , " & p.product & ", assigned to "
        If nsrc = "STAGING" Then
            s = s & "FORKLIFT."
        Else
            s = s & nsrc & "."
        End If
        Call debug_log(s, p, puser)
    End If
    ds.Close ': db.Close
    tag_alternate = "<!-- alternate assigned -->"
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "tag_alternate", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "tag_alternate", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: tag_alternate: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Function

Public Sub tlw_to_tml(m As ptask)
    'On Error Resume Next
    Dim sSql As String, sCols As String, sRows As String
    Dim p As ptask, i As Long, s As String
    'Dim db As ADODB.Connection
    Dim ds As ADODB.Recordset
    On Error GoTo vberror
    'Set db = CreateObject("ADODB.Connection")
    'db.Open WDbbsr
    'tracelist = tracelist & "<!-- tlw_to_tml() -->" & vbCrLf
    sSql = "Select palletid From paltasks Where source in ('TRAFFIC MASTER','RC119')"
    sSql = sSql & " and palletid = '" & m.palletid & "'"
    Set ds = Wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        s = ds!palletid
        s = s & " is already at " & m.target & "."
        Call debug_log(s, m, m.userid)
        ds.Close ': db.Close
        Exit Sub
    End If
    ds.Close ': db.Close
    m.status = "COMP"
    m.trandate = Format(Now, "yyMMdd HH:mm:ss")
    If UCase(m.uom) = "PALLET" Or m.target = "RC119" Then
        Call post_move_trans(m)
    Else
        Call post_recv_trans(m)
    End If
    p.area = "TRAFFIC MASTER"
    If m.description = "HOLD" Then                      'jv092316
        p.description = "HOLD"                          'jv092316
    Else                                                'jv092316
        p.description = " "                             'jv092316
    End If                                              'jv092316
    'p.source = m.target
    p.source = m.source
    
    If m.target = "SR4" Or m.target = "SR5" Then        'jv092316
        p.target = m.target                             'jv092316
    Else                                                'jv092316
        p.target = "..."                                'jv092316
    End If                                              'jv092316
    s = wrapdesc(Trim(Mid(m.product, 1, 4)), Val(m.units) + Val(m.units2))
    If Len(s) > 1 Then
        p.product = Trim(Mid(m.product, 1, 4)) & s & " " & StrConv(Mid(m.product, 5, Len(m.product) - 4), vbProperCase)
    Else
        p.product = m.product
    End If
    p.palletid = m.palletid
    p.qty = m.qty
    p.uom = m.uom
    p.lotnum = m.lotnum
    p.units = m.units
    p.lotnum2 = m.lotnum2
    p.units2 = m.units2
    'If m.source = "TRI-LEVEL 3" Or m.source = "TRI-LEVEL 4" Then        'jv092614
    '    p.status = "GATE"                                               'jv092614
    'Else                                                                'jv092614
        p.status = "PEND"
    'End If                                                              'jv092614
    p.userid = " "
    p.trandate = Format(Now, "yyMMdd HH:mm:ss")
    'p.reqid = " "
    p.reqid = m.reqid
    'MsgBox m.target & " tlw_to_tml " & p.target
    i = insert_trans(p)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "tlw_to_tml", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "tlw_to_tml", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: tlw_to_tml: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Public Sub update_ante_room(m As ptask)
    'On Error Resume Next
    Dim s As String
    s = m.target
    'tracelist = tracelist & "<!-- update_ante_room() -->" & vbCrLf
    If check_anteroom("CHECK", m.palletid) <> "ANTE ROOM" Then
        Call insert_rack_pallet(m)
    End If
    m.status = "COMP"
    m.trandate = Format(Now, "yyMMdd HH:mm:ss")
    Call post_move_trans(m)
    m.target = s
    If m.description > "." And UCase(Mid(m.description, 1, 4)) <> "DROP" Then
        m.source = "ANTE ROOM"
        m.status = "PEND"
        m.userid = " "
    Else
        m.status = "COMP"
    End If
    Call update_trans(m)
End Sub

Public Sub update_op_rack(m As ptask)
    'On Error Resume Next
    'tracelist = tracelist & "<!-- update_op_rack(" & m.palletid & "," & m.lotnum & ") -->" & vbCrLf
    Call insert_rack_pallet(m)
    m.status = "COMP"
    m.trandate = Format(Now, "yyMMdd HH:mm:ss")
    Call post_move_trans(m)
    Call update_trans(m)
End Sub

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

Public Function wdempname(empid As String) As String
    Dim s As String, ds As ADODB.Recordset
    s = "Select listdisplay from valuelists where listname = 'wdempid'"
    s = s & " and listreturn = '" & empid & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        wdempname = ds(0)
    Else
        wdempname = empid
    End If
    ds.Close
End Function

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

Public Function wrapdesc(psku As String, punits As Integer)
    Dim wc As Integer
    wc = 1
    If Val(sku_info(psku, "units")) <> punits Then
        wc = Val(sku_info(psku, "units")) / Val(sku_info(psku, "wraps"))
        wc = punits / wc
        wrapdesc = "  " & wc & " Wraps!"
    Else
        wrapdesc = ""
    End If
End Function


