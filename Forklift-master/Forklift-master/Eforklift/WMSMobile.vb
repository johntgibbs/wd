Module WMSMobile
    Public skutab(9999, 4) As String
    Public logdir As String '= "\\bbc-01-wdmgmt\wd\data\testlog"
    Public tracelist As String '= " "
    Public debflag As Boolean '= False
    Public SPTarget As String '= "SNACK PLANT"
    Public BHDest As String '= "STAGING"
    Public SRFlag As Boolean '= False
    Public ARFlag As Boolean '= False
    Public TCarFlag As Boolean '= False
    Public WDOrg As String
    Public WDUserId As String
    Public WDbbsr As String
    Public daioradb As String
    Public daidock As String
    Public ship_units As String
    Public ship_lotnum As String
    Public ship_units2 As String
    Public ship_lotnum2 As String
    Public ship_plate As String
    Public Wdb As ADODB.Connection
    Public eno As Long
    Public edesc As String
    Public vberror_log As String
    Public labpix(9999) As labpic
    Public labfmtfile As String
    Public histbc As String

    Public Structure ptask
        Public id As Long
        Public area As String
        Public description As String
        Public source As String
        Public target As String
        Public product As String
        Public palletid As String
        Public qty As String
        Public uom As String
        Public lotnum As String
        Public units As String
        Public lotnum2 As String
        Public units2 As String
        Public status As String
        Public userid As String
        Public trandate As String
        Public reqid As String
    End Structure

    Public Structure labpic
        Public sku As String
        Public package As String
        Public name1 As String
        Public name2 As String
        Public name3 As String
    End Structure

    Public Sub add_alternate_dock_pallet(ByVal recid As Long, ByVal psource As String, ByVal paltarg As String)
        Dim p As ptask
        Dim rs As ADODB.Recordset, s As String, zid As Long
        Dim sSql As String, i As Long, psku As String, pgroup As String, sCols As String, sRows As String
        On Error GoTo vberror
        p = masterec(recid)
        psku = Trim(Left(p.product, 4))
        pgroup = Trim(p.description)
        p.source = psource
        p.target = paltarg
        If psource = "SR5" Or psource = "SR6" Then p.palletid = psku & " ...... . ..."
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
            sSql = "Select id, ship_status From ship_infc Where order_num = '" & pgroup & "'"
            sSql = sSql & " And sku = '" & psku & "'"
            sSql = sSql & " And to_whse_num = " & Mid(psource, 3, 1)
            rs = Wdb.Execute(sSql)
            If rs.BOF = False Then
                rs.MoveFirst()
                i = rs.Fields(0).Value
                s = rs.Fields(1).Value
                If s = "DONE" Or s = "CANC" Then
                    sSql = "Update ship_infc Set order_qty = 1, ship_uom_qty = 0"
                    sSql = sSql & ", ship_plt_qty = 0, ship_status = 'NEW'"
                    sSql = sSql & " Where id = " & i
                    Wdb.Execute(sSql)
                Else
                    sSql = "Update ship_infc Set order_qty = order_qty + 1"
                    sSql = sSql & " Where id = " & i
                    Wdb.Execute(sSql)
                End If
            Else
                rs.Close()
                sSql = "Select id, ship_status From ship_infc"
                sSql = sSql & " Where ship_status in ('CANC','DONE')"
                sSql = sSql & " Order By id"
                rs = Wdb.Execute(sSql)
                If rs.BOF = False Then
                    rs.MoveFirst()
                    i = rs.Fields(0).Value
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
                    MsgBox(sSql)
                    Wdb.Execute(sSql)
                End If
                rs.Close()
            End If
        End If
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub add_alternate_daifuku(ByVal ptarget As String, ByVal pdock As String, ByVal psku As String, ByVal pdesc As String)
        Dim xname As String, cfile As String, s As String
        On Error GoTo vberror
        xname = "OrderItemMessage"
        cfile = "c:\jvwork\dai" & xname & ".xml"
        'cfile = "\\bbc-01-prodtrk\wd\sr5\bin\dai" & xname & ".xml"
        FileOpen(1, cfile, OpenMode.Append, OpenAccess.Default, OpenShare.Shared)

        s = "<?xml version=" & Chr(34) & "1.0" & Chr(34)
        s = s & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>" & vbCrLf
        s = s & "<!DOCTYPE OrderItemMessage SYSTEM " & Chr(34) & "wrxj.dtd" & Chr(34) & ">" & vbCrLf
        s = s & "<OrderItemMessage>" & vbCrLf
        s = s & "<Order action=" & Chr(34) & "ADD" & Chr(34)
        s = s & " sOrderID=" & Chr(34) & "ALT" & Right(DateDiff("s", "1-1-13 01:00:00 am", Now), 5) & Chr(34)
        s = s & " iPriority=" & Chr(34) & "3" & Chr(34)
        s = s & " iOrderStatus=" & Chr(34) & "READY" & Chr(34) & ">" & vbCrLf
        Print(1, s)

        s = "<OrderHeader>" & vbCrLf
        s = s & "<sDestinationStation>" & pdock & "</sDestinationStation>" & vbCrLf
        s = s & "<sDescription>" & ptarget & "</sDescription>" & vbCrLf
        s = s & "<sOrderMessage/>" & vbCrLf
        s = s & "</OrderHeader>" & vbCrLf
        Print(1, s)

        s = "<OrderLine sItem=" & Chr(34) & psku & Chr(34) & ">" & vbCrLf
        s = s & "<sRouteID/>" & vbCrLf
        s = s & "<fOrderQuantity>1</fOrderQuantity>" & vbCrLf
        s = s & "<sDescription>" & pdesc & "</sDescription>" & vbCrLf
        s = s & "</OrderLine>" & vbCrLf
        Print(1, s)

        s = "</Order>" & vbCrLf
        s = s & "</OrderItemMessage>"
        Print(1, s)
        FileClose(1)
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function barcode_profile(ByVal ubar As String) As Boolean
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

        'Check spaces
        If Mid(ubar, 11, 1) <> " " Then barcode_profile = False
        If Mid(ubar, 13, 1) <> " " Then barcode_profile = False

        'Check OpCode
        s = UCase(Mid(ubar, 12, 1))
        If s < "A" Or s > "Z" Then barcode_profile = False

        'Check pallet sequence number for spaces
        If Mid(ubar, 14, 1) = " " Then barcode_profile = False
        If Mid(ubar, 15, 1) = " " Then barcode_profile = False
        If Mid(ubar, 16, 1) = " " Then barcode_profile = False
    End Function

    Public Function barcode_to_lotnum(ByVal mbar As String)
        Dim s1 As String
        Dim s2 As String
        Dim s As String = " "
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

    Public Function check_anteroom(ByVal msrc As String, ByVal mbar As String) As String
        Dim ds As ADODB.Recordset
        On Error GoTo vberror
        Dim sSql As String
        'tracelist = tracelist & "<!-- check anteroom(" & msrc & "," & mbar & ")"
        sSql = "Select sku from rackpos where barcode = '" & mbar & "'" & _
               " and rackno in (select id from racks where aisle = 'M' and rack = 'ANTE')"

        ds = Wdb.Execute(sSql)

        If ds.BOF = False Then
            check_anteroom = "ANTE ROOM"
            'tracelist = tracelist & " = ANTE ROOM"
        Else
            check_anteroom = msrc
        End If
        'tracelist = tracelist & " -->" & vbCrLf
        ds.Close() ': db.Close
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Function check_hold(ByVal p As ptask) As Boolean                                      'jv040615
        Dim psku As String, pcode As String, hflag As Boolean, s As String
        Dim ds As ADODB.Recordset, palno As String
        psku = Trim(Mid(p.palletid, 1, 4))
        pcode = Trim(Mid(p.palletid, 11, 3))                                        'jv052515
        palno = Mid(p.palletid, 14, 3)
        hflag = False
        s = "select listreturn from valuelists where listname = 'wmsexpdate'"       'jv042115
        ds = Wdb.Execute(s)                                                     'jv042115
        If ds.BOF = False Then                                                      'jv042115
            ds.MoveFirst()                                                            'jv042115
            If p.lotnum <= ds.Fields(0).Value Then hflag = True 'jv042115
            If p.lotnum2 > "0" Then                                                 'jv042115
                If hflag = False And p.lotnum2 <= ds.Fields(0).Value Then hflag = True 'jv042115
            End If                                                                  'jv042115
        End If                                                                      'jv042115
        ds.Close()                                                                    'jv042115
        If hflag = False Then                                                       'jv042115
            s = "select id, spallet, epallet from holdlist where sku = '" & psku & "'"
            s = s & " and lot_num = '" & p.lotnum & "'"
            s = s & " and opcode = '" & pcode & "'"
            ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst()
                Do Until ds.EOF
                    If palno >= ds.Fields(1).Value And palno < ds.Fields(2).Value Then
                        hflag = True
                    End If
                    If hflag = True Then Exit Do
                    ds.MoveNext()
                Loop
            End If
            ds.Close()
        End If                                                                      'jv042115
        If p.lotnum2 > "0" And hflag = False Then
            s = "select id, spallet, epallet from holdlist where sku = '" & psku & "'"
            s = s & " and lot_num = '" & Mid(p.lotnum2, 1, 5) & "'"
            If Len(p.lotnum2) > 5 Then pcode = Trim(Mid(p.lotnum2, 6, 5)) 'jv052515
            s = s & " and opcode = '" & pcode & "'"
            ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst()
                Do Until ds.EOF
                    If palno >= ds.Fields(1).Value And palno <= ds.Fields(2).Value Then
                        hflag = True
                    End If
                    If hflag = True Then Exit Do
                    ds.MoveNext()
                Loop
            End If
            ds.Close()
        End If
        check_hold = hflag
    End Function



    Public Sub crane_finished_goods_lane(ByVal m As ptask)
        Dim ds As ADODB.Recordset
        Dim sSql As String, sRows As String, sCols As String
        Dim K As Integer, zid As Long, ncnt As Long
        On Error GoTo vberror
        zid = new_pallet_queue()
        'tracelist = tracelist & "<!-- crane_finished_goods_lane(" & m.id & ") -->" & vbCrLf
        sSql = "select * from queue_infc where id = " & zid
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            sSql = "update queue_infc set whse_num = " & Right(m.target, 1) & "," & _
                    "sku = '" & Trim(Mid(m.palletid, 1, 4)) & "'," & _
                    "lot_num = '" & m.lotnum & "'," & _
                    "drop_flag = ' '," & _
                    "rack_num = " & Val(m.qty) & "," & _
                    "units = " & Val(m.units) & "," & _
                    "lot_num2 = '" & m.lotnum2 & "'," & _
                    "units2 = " & Val(m.units2) & "," & _
                    "palletid = '" & m.palletid & "'," & _
                    "source = 'FG" & Right(m.target, 1) & "'" & _
                    " where id = " & zid
            Wdb.Execute(sSql)
        Else
            zid = wd_seq("Queue_Infc")
            sSql = "INSERT INTO queue_infc (ID,Whse_num,SKU,Lot_num,Drop_Flag,Queue_Num," & _
                   "Rack_Num,Units,Lot_Num2,Units2,PalletId,Source) VALUES (" & _
                   zid & "," & _
                   Right(m.target, 1) & ",'" & _
                   Trim(Mid(m.palletid, 1, 4)) & "','" & _
                   m.lotnum & "'," & _
                    "' '," & _
                    "0," & _
                    Val(m.qty) & "," & _
                    Val(m.units) & ",'" & _
                    m.lotnum2 & "'," & _
                    Val(m.units2) & ",'" & _
                    m.palletid & "'," & _
                    "'FG" & Right(m.target, 1) & "'"
            Wdb.Execute(sSql)
        End If
        ds.Close()

        sSql = "Update prodrcv set sr" & Right(m.target, 1) & "= sr" & Right(m.target, 1) & " + 1" & _
               " Where sku = '" & Trim(Mid(m.palletid, 1, 4)) & "'" & _
               " And lot_num = '" & m.lotnum & "'"
        Wdb.Execute(sSql)
        m.status = "COMP"
        m.trandate = Format(Now, "yyMMdd HH:mm:ss")
        Call post_move_trans(m)
        Call update_trans(m)
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub debug_log(ByVal s As String, ByVal m As ptask, ByVal muser As String)
        MsgBox(s, vbOKOnly, "Debug...")
    End Sub

    Public Sub dockfl_to_rstg(ByVal m As ptask)
        'tracelist = tracelist & "<!-- dockfl_to_rstg(" & m.id & ") -->" & vbCrLf
        If m.id = 0 Then Exit Sub
        Call post_move_trans(m)
        m.area = "FORKLIFT"
        m.description = " "
        m.source = "STAGING"
        m.target = "..."
        m.status = "PEND"
        m.userid = " "
        m.trandate = Format(Now, "yyMMdd HH:mm:ss")
        Call post_ship_trans(m)
        Call update_trans(m)
    End Sub

    Public Sub dockfl_to_trailer(ByVal m As ptask)
        'tracelist = tracelist & "<!-- dockfl_to_trailer(" & m.id & ") -->" & vbCrLf
        If m.id = 0 Then Exit Sub
        If SRFlag = True Or TCarFlag = True Then
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

    Public Sub efl_rack_moves(ByVal m As ptask)
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

    Public Sub efl_to_dstg(ByVal m As ptask)
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

        sSql = "Select id From paltasks Where area = 'DOCK'"
        sSql = sSql & " And description = '" & Trim(Mid(m.description, 1, 6)) & "'"
        sSql = sSql & " And target = '" & Right(m.description, Len(m.description) - 8) & "'"
        sSql = sSql & " And source in ('STAGING','ANTE ROOM','M-ANTE')"
        sSql = sSql & " And product = '" & m.product & "'"
        sSql = sSql & " And palletid < '0'"
        sSql = sSql & " And status = 'PEND'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            K = ds.Fields(0).Value
            sSql = "Update paltasks Set source = 'STAGING',"
            sSql = sSql & "palletid = '" & m.palletid & "',"
            sSql = sSql & "qty = " & Val(m.qty) & ","
            sSql = sSql & "lotnum = '" & m.lotnum & "',"
            sSql = sSql & "units = " & Val(m.units) & ","
            sSql = sSql & "lotnum2 = '" & m.lotnum2 & "',"
            sSql = sSql & "units2 = " & Val(m.units2)
            sSql = sSql & " Where id = " & K
            Wdb.Execute(sSql)
            m.status = "COMP"
            m.trandate = Format(Now, "yyMMdd HH:mm:ss")
            Call update_trans(m)
        Else
            Call debug_log("Failed efl_to_dstg: " & sSql, m, m.userid)
        End If
        ds.Close()
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function insert_rack_pallet(ByVal p As ptask) As Boolean
        Dim ds As ADODB.Recordset
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

        K = 0
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            K = ds.Fields(0).Value
        End If
        ds.Close()

        If K = 0 Then
            Call debug_log("failed to update rack in insert_pallet_rack: " & p.target, p, p.userid)
            insert_rack_pallet = False
            Exit Function
        End If

        j = 0
        sSql = "Select id From rackpos Where rackno = " & K & _
               " And count_qty = 0 Order by posn_num Desc"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            j = ds.Fields(0).Value
        End If
        ds.Close()

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
            Wdb.Execute(sSql)
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
                Wdb.Execute(sSql)
            Else
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
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                If ds.Fields(3).Value = "Y" Then
                    pqty = pqty + 1
                Else
                    pqty4 = pqty4 + 1
                End If
                rlot = ds.Fields(1).Value
                If Val(rlot) > 0 And rlot < olot Then olot = rlot
                rlot = ds.Fields(2).Value
                If rlot > "0" Then
                    If Val(rlot) > 0 And rlot < olot Then olot = rlot
                End If
                ds.MoveNext()
            Loop
        Else
            ds.Close()
            Call debug_log(sSql & " failed in insert_rack_pallet..", p, "0")
            insert_rack_pallet = False
            Exit Function
        End If
        ds.Close()

        If pqty + pqty4 = 0 Then
            sSql = "Update racks set sku = ' ',lot_num = ' ',qty = 0,qty4 = 0,resv_sku = ' ',resv_lot = ' '"
        Else
            sSql = "Update racks set sku = '" & psku & "'," & _
                   "lot_num = '" & olot & "'," & _
                   "qty = " & pqty & "," & _
                   "qty4 = " & pqty4
        End If
        sSql = sSql & " Where id = " & K
        Wdb.Execute(sSql)
        insert_rack_pallet = True
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function insert_trans(ByVal pt As ptask) As Long
        Dim sSql As String
        Dim zid As Long
        On Error GoTo vberror
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
        Wdb.Execute(sSql)
        insert_trans = zid
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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
        Dim s As String = ""
        Dim ts As Integer = 1
        Dim te As Integer = 1
        Dim psku As String = " "
        On Error GoTo vberror
        'tracelist = tracelist & "<!-- post_move_trans(" & m.id & ") -->" & vbCrLf
        FileOpen(1, labfmtfile, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
        Do Until EOF(1)
            s = LineInput(1)
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
        FileClose(1)
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.Description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, "wmsmobile.bas", "load_labpics", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: load_labpics: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Sub

    Public Function masterec(ByVal taskid As Long) As ptask
        Dim ds As ADODB.Recordset
        Dim sSql As String
        On Error GoTo vberror

        sSql = "Select * From paltasks Where id = " & taskid
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            masterec.id = ds.Fields(0).Value
            masterec.area = ds.Fields(1).Value
            masterec.description = ds.Fields(2).Value
            masterec.source = ds.Fields(3).Value
            masterec.target = ds.Fields(4).Value
            masterec.product = ds.Fields(5).Value
            masterec.palletid = ds.Fields(6).Value
            masterec.qty = ds.Fields(7).Value
            masterec.uom = ds.Fields(8).Value
            masterec.lotnum = ds.Fields(9).Value
            masterec.units = ds.Fields(10).Value
            masterec.lotnum2 = ds.Fields(11).Value
            masterec.units2 = ds.Fields(12).Value
            masterec.status = ds.Fields(13).Value
            masterec.userid = ds.Fields(14).Value
            masterec.trandate = ds.Fields(15).Value
            masterec.reqid = ds.Fields(16).Value
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
        ds.Close() ': db.Close
        'tracelist = tracelist & "<!-- masterec(" & taskid & ")=" & masterec.palletid
        'tracelist = tracelist & "," & masterec.lotnum & "," & masterec.units & ","
        'tracelist = tracelist & masterec.lotnum2 & "," & masterec.units2 & ","
        'tracelist = tracelist & masterec.uom & "," & masterec.qty & ","
        'tracelist = tracelist & masterec.target & " -->" & vbCrLf
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.Description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function move_task_crane(ByVal srkey As Long, ByVal bc As String, ByVal ptarget As String, ByVal puser As String) As ptask
        Dim ds As ADODB.Recordset, ss As ADODB.Recordset
        Dim sSql As String, sCols As String, sRows As String
        Dim paisle As String, prack As String
        On Error GoTo vberror
        sSql = "select sku,barcode,lot_num,count_qty,lot2,qty2 from position where id = " & srkey & " and barcode = '" & bc & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            move_task_crane.area = "EFLMove"
            move_task_crane.description = "Rack Move"
            move_task_crane.source = "CRANE"
            move_task_crane.target = ptarget
            move_task_crane.product = ds.Fields(0).Value & " " & sku_info(ds.Fields(0).Value, "desc")
            move_task_crane.palletid = ds.Fields(1).Value
            move_task_crane.qty = "1"
            move_task_crane.uom = "Pallet"
            move_task_crane.lotnum = ds.Fields(2).Value
            move_task_crane.units = ds.Fields(3).Value
            move_task_crane.lotnum2 = ds.Fields(4).Value
            move_task_crane.units2 = ds.Fields(5).Value
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
        ds.Close()
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function move_task_pallet(ByVal palkey As Long, ByVal bc As String, ByVal ptarget As String, ByVal puser As String) As ptask
        Dim ds As ADODB.Recordset, ss As ADODB.Recordset
        Dim sSql As String, sCols As String, sRows As String
        Dim paisle As String, prack As String
        On Error GoTo vberror
        sSql = "select sku,barcode,lot1,qty1,lot2,qty2,plateno from pallets where id = " & palkey & " and barcode = '" & bc & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            move_task_pallet.area = "EFLMove"
            move_task_pallet.description = "Rack Move"
            move_task_pallet.source = "PALLET"
            move_task_pallet.target = ptarget
            move_task_pallet.product = ds.Fields(0).Value & " " & sku_info(ds.Fields(0).Value, "desc")
            move_task_pallet.palletid = ds.Fields(1).Value
            move_task_pallet.qty = "1"
            move_task_pallet.uom = "Pallet"
            move_task_pallet.lotnum = ds.Fields(2).Value
            move_task_pallet.units = ds.Fields(3).Value
            move_task_pallet.lotnum2 = ds.Fields(4).Value
            move_task_pallet.units2 = ds.Fields(5).Value
            move_task_pallet.status = "PEND"
            move_task_pallet.userid = puser
            move_task_pallet.trandate = Format(Now, "yyMMdd HH:mm:ss")
            move_task_pallet.reqid = ds.Fields(6).Value
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
        ds.Close()
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function move_task_queue(ByVal quekey As Long, ByVal bc As String, ByVal ptarget As String, ByVal puser As String) As ptask
        Dim ds As ADODB.Recordset, ss As ADODB.Recordset
        Dim sSql As String, sCols As String, sRows As String
        Dim paisle As String, prack As String
        On Error GoTo vberror
        sSql = "select sku,palletid,lot_num,units,lot_num2,units2 from queue_infc where id = " & quekey & " and palletid = '" & bc & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            move_task_queue.area = "EFLMove"
            move_task_queue.description = "Rack Move"
            move_task_queue.source = "QUEUE"
            move_task_queue.target = ptarget
            move_task_queue.product = ds.Fields(0).Value & " " & sku_info(ds.Fields(0).Value, "desc")
            move_task_queue.palletid = ds.Fields(1).Value
            move_task_queue.qty = "1"
            move_task_queue.uom = "Pallet"
            move_task_queue.lotnum = ds.Fields(2).Value
            move_task_queue.units = ds.Fields(3).Value
            move_task_queue.lotnum2 = ds.Fields(4).Value
            move_task_queue.units2 = ds.Fields(5).Value
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
        ds.Close()
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function move_task_rack(ByVal rackkey As Long, ByVal bc As String, ByVal ptarget As String, ByVal puser As String) As ptask
        Dim ds As ADODB.Recordset, ss As ADODB.Recordset
        Dim sSql As String, sCols As String, sRows As String
        Dim paisle As String, prack As String
        On Error GoTo vberror
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
        sSql = "select rackno,sku,barcode,lot_num,count_qty,lot2, qty2 from rackpos where rackno = " & rackkey & " and barcode = '" & bc & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            move_task_rack.area = "EFLMove"
            move_task_rack.description = "Rack Move"
            sSql = "Select aisle, rack from racks where id = " & ds.Fields(0).Value
            ss = Wdb.Execute(sSql)
            If ss.BOF = False Then
                ss.MoveFirst()
                move_task_rack.source = ss.Fields(0).Value & "-" & ss.Fields(1).Value
            End If
            ss.Close()
            move_task_rack.target = ptarget
            move_task_rack.product = ds.Fields(1).Value & " " & sku_info(ds.Fields(1).Value, "desc")
            move_task_rack.palletid = ds.Fields(2).Value
            move_task_rack.qty = "1"
            move_task_rack.uom = "Pallet"
            move_task_rack.lotnum = ds.Fields(3).Value
            move_task_rack.units = ds.Fields(4).Value
            move_task_rack.lotnum2 = ds.Fields(5).Value
            move_task_rack.units2 = ds.Fields(6).Value
            move_task_rack.status = "PEND"
            move_task_rack.userid = puser
            move_task_rack.trandate = Format(Now, "yyMMdd HH:mm:ss")
            move_task_rack.reqid = " "
        End If
        ds.Close()
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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
        Dim ds As ADODB.Recordset
        Dim sSql As String, sCols As String, sRows As String
        Dim K As Integer, ncnt As Integer
        Dim zid As Long
        On Error GoTo vberror

        sSql = "Select max(queue_num) from queue_infc"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            K = ds.Fields(0).Value + 1
        Else
            K = 100
        End If
        ds.Close()

        sSql = "Select id, queue_num from queue_infc where queue_num = 0"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            zid = ds.Fields(0).Value
            sSql = "update queue_infc set queue_num = " & K & _
                   "Where id = " & zid
            Wdb.Execute(sSql)
        Else
            zid = wd_seq("Queue_infc")
            sSql = "Insert into queue_infc (ID, Queue_num) VALUES (" & _
                   zid & "," & K & ")"
            Wdb.Execute(sSql)
        End If
        ds.Close()
        new_pallet_queue = zid
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function new_pallet_task_record(ByVal parea As String) As Long
        Dim ds As ADODB.Recordset
        Dim sSql As String, sCols As String, sRows As String
        Dim zid As Long
        On Error GoTo vberror
        zid = 0
        sSql = "Select id, status From paltasks Where area = '" & parea & "'"
        sSql = sSql & " and status = 'COMP'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            zid = ds.Fields(0).Value
        End If
        ds.Close()
        If zid = 0 Then
            sSql = "Select id, status From paltasks Where status = 'COMP'"
            ds = Wdb.Execute(sSql)
            If ds.BOF = False Then
                ds.MoveFirst()
                zid = ds.Fields(0).Value
            End If
            ds.Close()
        End If
        If zid > 0 Then
            sSql = "Update paltasks Set status = 'PEND' Where id = " & zid
            Wdb.Execute(sSql)
        Else
            zid = wd_seq("PalTasks")
            sSql = "Insert Into paltasks (ID) Values (" & zid & ")"
            Wdb.Execute(sSql)
        End If
        new_pallet_task_record = zid
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.Description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function order_pick_position(ByVal psku As String) As Integer
        Dim ds As ADODB.Recordset
        Dim sSql As String
        On Error GoTo vberror
        sSql = "Select opseq From oplist Where sku = '" & psku & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            order_pick_position = ds.Fields(0).Value
        Else
            order_pick_position = 0
        End If
        ds.Close()
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub post_move_trans(ByVal m As ptask)
        Dim cfile As String
        If UCase(m.description) = "2 STEP REQUEST" Then Exit Sub
        On Error GoTo vberror
        'tracelist = tracelist & "<!-- post_move_trans(" & m.id & ") -->" & vbCrLf
        cfile = logdir & "move" & Format(Now, "MMddyyyy") & ".txt"
        FileOpen(1, cfile, OpenMode.Append, OpenAccess.Default, OpenShare.Shared)
        WriteLine(1, m.id, m.area, m.description, m.source, m.target, m.product, m.palletid, m.qty, m.uom, m.lotnum, m.units, m.lotnum2, m.units2, m.status, m.userid, m.trandate, m.reqid)
        FileClose(1)
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub post_recv_trans(ByVal m As ptask)
        Dim cfile As String
        If UCase(m.description) = "2 STEP REQUEST" Then Exit Sub
        On Error GoTo vberror
        'tracelist = tracelist & "<!-- post_move_trans(" & m.id & ") -->" & vbCrLf
        cfile = logdir & "recv" & Format(Now, "MMddyyyy") & ".txt"
        FileOpen(1, cfile, OpenMode.Append, OpenAccess.Default, OpenShare.Shared)
        WriteLine(1, m.id, m.area, m.description, m.source, m.target, m.product, m.palletid, m.qty, m.uom, m.lotnum, m.units, m.lotnum2, m.units2, m.status, m.userid, m.trandate, m.reqid)
        FileClose(1)
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub post_ship_trans(ByVal m As ptask)
        Dim cfile As String
        Dim sSql As String
        If UCase(m.description) = "2 STEP REQUEST" Then Exit Sub
        On Error GoTo vberror
        'tracelist = tracelist & "<!-- post_ship_trans(" & m.id & ") -->" & vbCrLf
        cfile = logdir & "ship" & Format(Now, "MMddyyyy") & ".txt"
        FileOpen(1, cfile, OpenMode.Append, OpenAccess.Default, OpenShare.Shared)
        WriteLine(1, m.id, m.area, m.description, m.source, m.target, m.product, m.palletid, m.qty, m.uom, m.lotnum, m.units, m.lotnum2, m.units2, m.status, m.userid, m.trandate, m.reqid)
        FileClose(1)
        If SRFlag = True Or TCarFlag = True Then
            sSql = "Update pallets Set source = '" & m.source & "', target = '" & m.target & "'"
            sSql = sSql & ", status = 'Shipped', trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
            sSql = sSql & " Where barcode = '" & m.palletid & "'"
            Wdb.Execute(sSql)
        End If
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function pallet_history_text(ByVal bc As String) As String
        Dim ds As ADODB.Recordset, ss As ADODB.Recordset
        Dim spath As String, sdir As String, fdate As String
        Dim sdate As String, edate As String, wsku As String, wlot As String
        Dim wzone As String, wstat As String, wgma As Integer, wside As String
        Dim waisle As String, wrack As String
        Dim s As String
        Dim f0 As String, f1 As String, f2 As String, f3 As String
        Dim f4 As String, f5 As String, f6 As String, f7 As String
        Dim f8 As String, f9 As String, f10 As String, f11 As String
        Dim f12 As String, f13 As String, f14 As String, f15 As String, f16 As String
        Dim pht As String, t As String, wqty As Integer
        Dim syear As Integer, eyear As Integer, i As Integer                        'jv061215
        Dim logpath As String
        f0 = " " : f1 = " " : f2 = " " : f3 = " " : f4 = " " : f5 = " " : f6 = " " : f7 = " " : f8 = " " : f9 = " "
        f10 = " " : f11 = " " : f12 = " " : f13 = " " : f14 = " " : f15 = " " : f16 = " "
        logpath = logdir
        pht = "" : bc = UCase(bc)
        sdate = Format(Val(Mid(bc, 9, 2)) - 2, "00")
        sdate = "20" & sdate & Mid(bc, 5, 4)
        edate = Format(Now, "yyyymmdd")
        wsku = Trim(Left(bc, 4))
        wlot = barcode_to_lotnum(bc)
        If wlot = "01001" Then
            pallet_history_text = "Error!!  Invalid BarCode:  " & bc
            Exit Function
        End If
        'Screen.MousePointer = 11
        Eforklift2.Cursor = Cursors.WaitCursor
        wqty = Val(sku_info(wsku, "units")) / Val(sku_info(wsku, "wraps"))


        'Current location
        'If Form1.plantno = "50" Then                'Search Cranes
        s = "select * from position where barcode = '" & bc & "'"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                wzone = "0" : wstat = " " : wgma = 0 : wside = " "
                s = "select zone_num, rack_side, lane_status, gmasize from lane where id = " & ds.Fields(1).Value
                ss = Wdb.Execute(s)
                If ss.BOF = False Then
                    ss.MoveFirst()
                    wzone = ss.Fields(0).Value
                    wstat = ss.Fields(2).Value '!lane_status
                    wgma = ss.Fields(3).Value '!gmasize
                    wside = ss.Fields(1).Value '!rack_side
                End If

                pht = pht & "Crane Location: SR-" & ds.Fields(2).Value & "  " '!whse_num & " "  '& ds!vert_loc & "-" & ds!horz_loc & "-" & ds!rack_side

                If ds.Fields(2).Value < 4 Then                         'Target
                    pht = pht & ds.Fields(3).Value & "-" & ds.Fields(4).Value & "-" & ds.Fields(5).Value & " " & ds.Fields(6).Value
                Else
                    pht = pht & wzone & "> " & ds.Fields(3).Value & "-" & ds.Fields(4).Value & "-" & wside
                End If
                If wstat = "H" Then pht = pht & " On Hold"
                If wstat = "B" Then pht = pht & " Blocked"
                ss.Close()
                pht = pht & " " & Format(ds.Fields(13).Value + ds.Fields(19).Value, "0") & " units, "
                pht = pht & Format((ds.Fields(13).Value + ds.Fields(19).Value) / wqty, "0") & " wraps"
                pht = pht & vbCrLf & vbCrLf
                ds.MoveNext()
            Loop
        End If
        ds.Close()
        'End If

        s = "select * from rackpos where barcode = '" & bc & "'"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                waisle = " " : wstat = " " : wrack = " "
                s = "select aisle, rack, hold from racks where id = " & ds.Fields(1).Value
                ss = Wdb.Execute(s)
                If ss.BOF = False Then
                    ss.MoveFirst()
                    waisle = Trim(ss.Fields(0).Value)
                    wrack = Trim(ss.Fields(1).Value)
                    If ss.Fields(2).Value = 0 Then
                        wstat = " "
                    Else
                        wstat = "On Hold"
                    End If
                End If
                pht = pht & "Rack Location:  " & waisle & "-" & wrack & " " & wstat
                pht = pht & " " & Format(ds.Fields(6).Value + ds.Fields(11).Value, "0") & " units, "
                pht = pht & Format((ds.Fields(6).Value + ds.Fields(11).Value) / wqty, "0") & " wraps"
                pht = pht & vbCrLf & vbCrLf
                ss.Close()
                ds.MoveNext()
            Loop
        End If
        ds.Close()

        syear = Val(Left(sdate, 4))                                             'jv061215
        eyear = Val(Left(edate, 4))                                             'jv061215
        s = ""
        spath = logpath & "recv*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            fdate = Right(sdir, 12)                                                         'jv061215
            fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
            If fdate >= sdate And fdate <= edate Then
                FileOpen(1, logpath & sdir, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
                Do Until EOF(1)
                    Input(1, f0)
                    Input(1, f1)
                    Input(1, f2)
                    Input(1, f3)
                    Input(1, f4)
                    Input(1, f5)
                    Input(1, f6)
                    Input(1, f7)
                    Input(1, f8)
                    Input(1, f9)
                    Input(1, f10)
                    Input(1, f11)
                    Input(1, f12)
                    Input(1, f13)
                    Input(1, f14)
                    Input(1, f15)
                    Input(1, f16)
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
                        't = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                        t = t & Mid(f15, 8, 8)
                        s = s & "Wrapped:        " & f3 & " "
                        If Len(s) < 50 Then s = s & Space(50 - Len(s))
                        's = s & t & " " & f14 & vbCrLf & vbCrLf
                        s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                    End If
                Loop
                FileClose(1)
            End If
            sdir = Dir$
            'DoEvents()
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
                    'Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1        'jv061215
                    FileOpen(1, logpath & Format(i, "0000") & "\" & sdir, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
                    Do Until EOF(1)
                        Input(1, f0)
                        Input(1, f1)
                        Input(1, f2)
                        Input(1, f3)
                        Input(1, f4)
                        Input(1, f5)
                        Input(1, f6)
                        Input(1, f7)
                        Input(1, f8)
                        Input(1, f9)
                        Input(1, f10)
                        Input(1, f11)
                        Input(1, f12)
                        Input(1, f13)
                        Input(1, f14)
                        Input(1, f15)
                        Input(1, f16)
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
                            't = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                            t = t & Mid(f15, 8, 8)
                            s = s & "Wrapped:        " & f3 & " "
                            If Len(s) < 50 Then s = s & Space(50 - Len(s))
                            's = s & t & " " & f14 & vbCrLf & vbCrLf
                            s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                        End If
                    Loop
                    FileClose(1)
                End If
                sdir = Dir$
                'DoEvents()
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
                FileOpen(1, logpath & sdir, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
                Do Until EOF(1)
                    Input(1, f0)
                    Input(1, f1)
                    Input(1, f2)
                    Input(1, f3)
                    Input(1, f4)
                    Input(1, f5)
                    Input(1, f6)
                    Input(1, f7)
                    Input(1, f8)
                    Input(1, f9)
                    Input(1, f10)
                    Input(1, f11)
                    Input(1, f12)
                    Input(1, f13)
                    Input(1, f14)
                    Input(1, f15)
                    Input(1, f16)
                    If f6 = bc Then
                        s = ""
                        t = Mid(f15, 3, 2) & "-" & Mid(f15, 5, 2) & "-" & Mid(f15, 1, 2) & " "
                        't = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                        t = t & Mid(f15, 8, 8)
                        s = "Traffic Master: " & f4 & " "
                        If Len(s) < 50 Then s = s & Space(50 - Len(s))
                        's = s & t & " " & f14 & vbCrLf & vbCrLf
                        s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                    End If
                Loop
                FileClose(1)
            End If
            sdir = Dir$
            'DoEvents()
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
                    'Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1        'jv061215
                    FileOpen(1, logpath & Format(i, "0000") & "\" & sdir, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
                    Do Until EOF(1)
                        Input(1, f0)
                        Input(1, f1)
                        Input(1, f2)
                        Input(1, f3)
                        Input(1, f4)
                        Input(1, f5)
                        Input(1, f6)
                        Input(1, f7)
                        Input(1, f8)
                        Input(1, f9)
                        Input(1, f10)
                        Input(1, f11)
                        Input(1, f12)
                        Input(1, f13)
                        Input(1, f14)
                        Input(1, f15)
                        Input(1, f16)
                        If f6 = bc Then
                            s = ""
                            t = Mid(f15, 3, 2) & "-" & Mid(f15, 5, 2) & "-" & Mid(f15, 1, 2) & " "
                            't = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                            t = t & Mid(f15, 8, 8)
                            s = "Traffic Master: " & f4 & " "
                            If Len(s) < 50 Then s = s & Space(50 - Len(s))
                            's = s & t & " " & f14 & vbCrLf & vbCrLf
                            s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                        End If
                    Loop
                    FileClose(1)
                End If
                sdir = Dir$
                'DoEvents()
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
                'Open logpath & sdir For Input Shared As #1
                FileOpen(1, logpath & sdir, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
                Do Until EOF(1)
                    Input(1, f0)
                    Input(1, f1)
                    Input(1, f2)
                    Input(1, f3)
                    Input(1, f4)
                    Input(1, f5)
                    Input(1, f6)
                    Input(1, f7)
                    Input(1, f8)
                    Input(1, f9)
                    Input(1, f10)
                    Input(1, f11)
                    Input(1, f12)
                    Input(1, f13)
                    Input(1, f14)
                    Input(1, f15)
                    Input(1, f16)
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
                        't = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                        t = t & Mid(f15, 8, 8)
                        s = "Moved:          " & Trim(f3) & " to " & Trim(f4) & " " '& t & " " & f14 & vbCrLf & vbCrLf
                        If Len(s) < 50 Then s = s & Space(50 - Len(s))
                        's = s & t & " " & f14 & vbCrLf & vbCrLf
                        s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                        pht = pht & s
                    End If
                Loop
                FileClose(1)
            End If
            sdir = Dir$
            'DoEvents()
        Loop
        For i = syear To eyear
            spath = logpath & Format(i, "0000") & "\move*.txt"                                  'jv061215
            sdir = Dir$(spath)
            Do While sdir <> ""
                fdate = Right(sdir, 12)                                                         'jv061215
                fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
                'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
                If fdate >= sdate And fdate <= edate Then
                    'Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1        'jv061215
                    FileOpen(1, logpath & Format(i, "0000") & "\" & sdir, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
                    Do Until EOF(1)
                        Input(1, f0)
                        Input(1, f1)
                        Input(1, f2)
                        Input(1, f3)
                        Input(1, f4)
                        Input(1, f5)
                        Input(1, f6)
                        Input(1, f7)
                        Input(1, f8)
                        Input(1, f9)
                        Input(1, f10)
                        Input(1, f11)
                        Input(1, f12)
                        Input(1, f13)
                        Input(1, f14)
                        Input(1, f15)
                        Input(1, f16)
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
                            't = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                            t = t & Mid(f15, 8, 8)
                            s = "Moved:          " & Trim(f3) & " to " & Trim(f4) & " " '& t & " " & f14 & vbCrLf & vbCrLf
                            If Len(s) < 50 Then s = s & Space(50 - Len(s))
                            's = s & t & " " & f14 & vbCrLf & vbCrLf
                            s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                            pht = pht & s
                        End If
                    Loop
                    FileClose(1)
                End If
                sdir = Dir$
                'DoEvents()
            Loop
        Next i

        spath = logpath & "ship*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            fdate = Right(sdir, 12)                                                         'jv061215
            fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                'Open logpath & sdir For Input Shared As #1
                FileOpen(1, logpath & sdir, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
                Do Until EOF(1)
                    Input(1, f0)
                    Input(1, f1)
                    Input(1, f2)
                    Input(1, f3)
                    Input(1, f4)
                    Input(1, f5)
                    Input(1, f6)
                    Input(1, f7)
                    Input(1, f8)
                    Input(1, f9)
                    Input(1, f10)
                    Input(1, f11)
                    Input(1, f12)
                    Input(1, f13)
                    Input(1, f14)
                    Input(1, f15)
                    Input(1, f16)
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
                        't = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                        t = t & Mid(f15, 8, 8)
                        s = "Shipped:        " & f3 & " " & f2 & " " & f4 & " "
                        If Len(s) < 50 Then s = s & Space(50 - Len(s))
                        's = s & t & " " & f14 & vbCrLf & vbCrLf
                        s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                        pht = pht & s
                    End If
                Loop
                FileClose(1)
            End If
            sdir = Dir$
            'DoEvents()
        Loop
        For i = syear To eyear
            spath = logpath & Format(i, "0000") & "\ship*.txt"                                  'jv061215
            sdir = Dir$(spath)
            Do While sdir <> ""
                fdate = Right(sdir, 12)                                                         'jv061215
                fdate = Mid(fdate, 5, 4) & Mid(fdate, 1, 4)                                     'jv061215
                'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
                If fdate >= sdate And fdate <= edate Then
                    'Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1        'jv061215
                    FileOpen(1, logpath & Format(i, "0000") & "\" & sdir, OpenMode.Input, OpenAccess.Default, OpenShare.Shared)
                    Do Until EOF(1)
                        Input(1, f0)
                        Input(1, f1)
                        Input(1, f2)
                        Input(1, f3)
                        Input(1, f4)
                        Input(1, f5)
                        Input(1, f6)
                        Input(1, f7)
                        Input(1, f8)
                        Input(1, f9)
                        Input(1, f10)
                        Input(1, f11)
                        Input(1, f12)
                        Input(1, f13)
                        Input(1, f14)
                        Input(1, f15)
                        Input(1, f16)
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
                            't = t & Format(Mid(f15, 8, 8), "hh:mm am/pm")
                            t = t & Mid(f15, 8, 8)
                            s = "Shipped:        " & f3 & " " & f2 & " " & f4 & " "
                            If Len(s) < 50 Then s = s & Space(50 - Len(s))
                            's = s & t & " " & f14 & vbCrLf & vbCrLf
                            s = s & t & " " & wdempname(f14) & vbCrLf & vbCrLf
                            pht = pht & s
                        End If
                    Loop
                    FileClose(1)
                End If
                sdir = Dir$
                'DoEvents()
            Loop
        Next i
        pallet_history_text = pht
        'Screen.MousePointer = 0
        Eforklift2.Cursor = Cursors.Default
    End Function


    Public Sub pallet_lots(ByVal bc As String)
        Dim ds As ADODB.Recordset
        Dim sSql As String
        'tracelist = tracelist & "<!-- pallet_lots(" & bc & ") -->" & vbCrLf
        On Error GoTo vberror
        sSql = "select qty1,lot1,qty2,lot2,plateno from pallets where barcode = '" & bc & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            ship_units = ds.Fields(0).Value
            ship_lotnum = ds.Fields(1).Value
            ship_units2 = ds.Fields(2).Value
            ship_lotnum2 = ds.Fields(3).Value
            ship_plate = ds.Fields(4).Value
        End If
        ds.Close()
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.Description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub post_sr4_remove(ByVal m As ptask, ByVal rknote As String)
        Dim cfile As String
        If UCase(m.description) = "2 STEP REQUEST" Then Exit Sub
        On Error GoTo vberror
        'tracelist = tracelist & "<!-- post_sr4_remove(" & m.id & "," & rknote & ") -->" & vbCrLf
        cfile = logdir & "move" & Format(Now, "MMddyyyy") & ".txt"
        FileOpen(1, cfile, OpenMode.Append, OpenAccess.Default, OpenShare.Shared)
        WriteLine(1, m.id, m.area, m.description, m.source, m.target, m.product, m.palletid, m.qty, m.uom, m.lotnum, m.units, m.lotnum2, m.units2, m.status, m.userid, m.trandate, m.reqid)
        FileClose(1)
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.Description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function remove_rack_pallet(ByVal m As ptask) As Boolean
        Dim ds As ADODB.Recordset
        Dim sSql As String, sCols As String, sRows As String
        Dim olot As String, rlot As String
        Dim K As Integer
        Dim pqty As Integer
        Dim pqty4 As Integer
        Dim rknote As String = ""
        Dim s As String
        Dim zid As Long
        Dim ncnt As Integer, i As Integer
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

        zid = 0
        s = Mid(m.palletid, 1, 11) & "_" & Mid(m.palletid, 13, 4)
        sSql = "Select id,rackno,barcode,lot_num,count_qty,lot2,qty2 From rackpos Where barcode in ('" & m.palletid & "','" & s & "')"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            zid = ds.Fields(0).Value
            K = ds.Fields(1).Value
            rknote = ds.Fields(2).Value
            m.lotnum = ds.Fields(3).Value
            m.units = ds.Fields(4).Value
            s = ds.Fields(5).Value
            If Len(s) > 0 Then
                m.lotnum2 = s
            Else
                m.lotnum2 = " "
            End If
            s = ds.Fields(6).Value
            If Len(s) > 0 Then
                m.units2 = s
            Else
                m.units2 = "0"
            End If
        End If
        ds.Close()
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
            Wdb.Execute(sSql)
            'Call debug_log(sSql, m, m.userid)
        Else
            m.lotnum = barcode_to_lotnum(m.palletid)
            Call update_trans(m)
        End If
        m.trandate = Format(Now, "yyMMdd HH:mm:ss")

        If K > 0 Then
            sSql = "Select aisle,rack From racks Where id = " & K
            ds = Wdb.Execute(sSql)
            If ds.BOF = False Then
                ds.MoveFirst()
                rknote = rknote & " @ "
                rknote = rknote & ds.Fields(0).Value
                rknote = rknote & "-"
                rknote = rknote & ds.Fields(1).Value
            Else
                rknote = rknote & " " & m.source & " not removed."
            End If
            ds.Close()
        End If
        Call post_sr4_remove(m, rknote)
        'Call debug_log(sSql, m, m.userid)

        If K = 0 Then
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
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                If ds.Fields(3).Value = "Y" Then
                    pqty = pqty + 1
                Else
                    pqty4 = pqty4 + 1
                End If
                rlot = ds.Fields(1).Value
                If Val(rlot) > 0 And rlot < olot Then olot = rlot
                rlot = ds.Fields(2).Value
                If rlot > "0" Then
                    If Val(rlot) > 0 And rlot < olot Then olot = rlot
                End If
                ds.MoveNext()
            Loop
        End If
        ds.Close()
        'Call debug_log(sSql, m, m.userid)

        If pqty + pqty4 = 0 Then
            sSql = "Update racks set sku = ' ',lot_num = ' ',qty = 0,qty4 = 0"
        Else
            sSql = "Update racks set lot_num = '" & olot & "',"
            sSql = sSql & "qty = " & pqty & ","
            sSql = sSql & "qty4 = " & pqty4
        End If
        sSql = sSql & " Where id = " & K
        Wdb.Execute(sSql)
        'Call debug_log(sSql, m, m.userid)
        remove_rack_pallet = True
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub remove_sp_order(ByVal m As ptask)
        Dim ds As ADODB.Recordset
        Dim sSql As String, sCols As String, sRows As String
        Dim zid As Long
        On Error GoTo vberror

        sSql = "Select id From paltasks Where area = 'SNACK PLANT WRAPPER'"
        sSql = sSql & " And source = 'M-SP'"
        sSql = sSql & " And status = 'PEND'"
        sSql = sSql & " And product > '" & Mid(m.palletid, 1, 4) & "'"
        sSql = sSql & " And product < '" & Mid(m.palletid, 1, 4) & "ZZZZ'"

        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            zid = ds.Fields(0).Value
            sSql = "Update paltasks Set palletid = '" & m.palletid & "'"
            sSql = sSql & ",status = 'COMP"
            sSql = sSql & ",userid = '" & m.userid & "'"
            sSql = sSql & " Where id = " & zid
            Wdb.Execute(sSql)
        End If
        ds.Close()
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function return_to_wrapper(ByVal ubar As String, ByVal uname As String, ByVal uarea As String, ByVal ureq As String) As String
        Dim ds As ADODB.Recordset
        Dim sSql As String, sCols As String, sRows As String
        Dim p As ptask, s As String, i As Long, n As Integer, K As Integer
        If Len(ubar) < 16 Then
            return_to_wrapper = "Illegal Barcode: " & ubar * "!!"
            Exit Function
        End If
        On Error GoTo vberror
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
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            i = ds.Fields(0).Value
            p = masterec(i)
            p.target = "WRAPPER"
            p.qty = Format(Val(p.qty) * -1, "0")
            p.units = Format(Val(p.units) * -1, "0")
            p.units2 = Format(Val(p.units2) * -1, "0")
            p.status = "COMP"
            p.userid = uname
            p.trandate = Format(Now, "yyMMdd HH:mm:ss")
            Do Until ds.EOF
                i = ds.Fields(0).Value
                sSql = "Update paltasks Set status = 'COMP', palletid = '...'"
                sSql = sSql & " Where id = " & i
                Wdb.Execute(sSql)
                ds.MoveNext()
            Loop
        End If
        ds.Close()

        sSql = "Select id,whse_num,sku,lot_num,units,lot_num2,units2 From queue_infc Where palletid = '" & ubar & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            i = ds.Fields(0).Value
            p.id = ds.Fields(0).Value
            p.area = uarea
            p.description = " "
            p.source = "SR-" & ds.Fields(1).Value
            p.target = "WRAPPER"
            s = ds.Fields(2).Value
            p.product = s & " " & sku_info(s, "desc")
            p.palletid = ubar
            p.qty = "-1"
            p.uom = "Pallet"
            p.lotnum = ds.Fields(3).Value
            p.units = Format(Val(ds.Fields(4).Value) * -1, "0")
            p.lotnum2 = ds.Fields(5).Value
            p.units2 = Format(Val(ds.Fields(6).Value) * -1, "0")
            p.status = "COMP"
            p.userid = uname
            p.trandate = Format(Now, "yyMMdd HH:mm:ss")
            p.reqid = "0"
            Do Until ds.EOF
                i = ds.Fields(0).Value
                sSql = "Update queue_infc Set queue_num = 0"
                sSql = sSql & " Where id = " & i
                Wdb.Execute(sSql)
                ds.MoveNext()
            Loop
        End If
        ds.Close()

        sSql = "Select id,sku,lot_num,count_qty,lot2,qty2 From rackpos Where barcode = '" & ubar & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            If p.id = 0 Then
                p.id = ds.Fields(0).Value
                p.area = uarea
                p.description = " "
                p.source = "RACKS"
                p.target = "WRAPPER"
                s = ds.Fields(1).Value
                p.product = s & " " & sku_info(s, "desc")
                p.palletid = ubar
                p.qty = "-1"
                p.uom = "Pallet"
                p.lotnum = ds.Fields(2).Value
                p.units = Format(Val(ds.Fields(3).Value) * -1, "0")
                p.lotnum2 = ds.Fields(4).Value
                p.units2 = Format(Val(ds.Fields(5).Value) * -1, "0")
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
        ds.Close()

        sSql = "select id from pallets where barcode = '" & ubar & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                sSql = "Update pallets set plateno = '0', barcode = '..', status = 'Shipped'"
                sSql = sSql & ", sku = ' '"
                sSql = sSql & " Where id = " & ds.Fields(0).Value
                Wdb.Execute(sSql)
                ds.MoveNext()
            Loop
        End If
        ds.Close()

        If p.id > 0 Or uarea = "EFLMove" Then
            Call post_recv_trans(p)
            return_to_wrapper = ubar & " returned."
        Else
            return_to_wrapper = ubar & " was not found."
        End If
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub robot0_pickup(ByVal m As ptask)
        Dim sSql As String
        Dim p As ptask, i As Long
        Dim s As String, t As String
        Dim p4way As Boolean
        On Error GoTo vberror
        Dim ds As ADODB.Recordset
        t = "ORDER PICK"
        p4way = False
        'tracelist = tracelist & "<!-- robot0_pickup() -->" & vbCrLf
        sSql = "Select palletid,source From paltasks Where source in ('ROBOT ZERO','WRAPPER')"
        sSql = sSql & " And area = 'FORKLIFT'"
        sSql = sSql & " And palletid = '" & m.palletid & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            s = ds.Fields(0).Value
            s = s & " already at "
            s = s & ds.Fields(1).Value
            s = s & " Pick Up."
            ds.Close()
            Call debug_log(s, m, m.userid)
            Exit Sub
        End If
        ds.Close()
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
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            t = Trim(ds.Fields(0).Value) & "-" & Trim(ds.Fields(1).Value)
        End If
        ds.Close()
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
        p.reqid = m.reqid
        i = insert_trans(p)
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub roller_bed_pickup(ByVal m As ptask)
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

    Public Sub Set_WD_Org(ByVal orgcode As String)
        If orgcode = "500" Then
            logdir = "\\bbc-01-wdmgmt\wd\data\testlog"
            tracelist = " "
            debflag = False
            SPTarget = "SNACK PLANT"
            BHDest = "STAGING"
            SRFlag = True
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
            SRFlag = False
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
            SRFlag = True
            ARFlag = True
            TCarFlag = False
            'WDOrg = gsOrgCode
        End If

    End Sub

    Public Function sku_info(ByVal psku As String, ByVal pfld As String) As String
        Dim sSql As String
        Dim i As Integer
        Dim ds As ADODB.Recordset
        On Error GoTo vberror
        i = Val(psku)
        If i > 1000 Then i = 1

        If skutab(i, 0) = psku Then
        Else
            sSql = "select sku, uom_type, description, uom_per_pallet, qty_per_pallet" & _
                   " from sku_config where sku = '" & psku & "'"

            ds = Wdb.Execute(sSql)
            If ds.BOF = False Then
                ds.MoveFirst()
                skutab(i, 0) = ds.Fields(0).Value
                skutab(i, 1) = ds.Fields(1).Value & " "
                skutab(i, 1) = skutab(i, 1) & ds.Fields(2).Value
                skutab(i, 2) = ds.Fields(3).Value
                If Val(skutab(i, 2)) < 1 Then skutab(i, 2) = "1"
                skutab(i, 3) = ds.Fields(4).Value
                If Val(skutab(i, 3)) < 1 Then skutab(i, 3) = "1"
            Else
                skutab(i, 0) = "0"
                skutab(i, 1) = "unrecognized SKU"
                skutab(i, 2) = "1"
                skutab(i, 3) = "1"
            End If
            ds.Close()
        End If

        sku_info = skutab(i, 0)
        If LCase(pfld) = "wraps" Then sku_info = skutab(i, 3)
        If LCase(pfld) = "units" Then sku_info = skutab(i, 2)
        If LCase(pfld) = "desc" Then sku_info = skutab(i, 1)
        'tracelist = tracelist & "<!-- sku_info(" & psku & "," & pfld & ") = " & sku_info & " -->" & vbCrLf
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function space_in_rack(ByVal maisle As String, ByVal mrack As String, ByVal msku As String, ByVal m As ptask) As Boolean
        Dim sSql As String
        Dim ds As ADODB.Recordset
        On Error GoTo vberror
        'tracelist = tracelist & "<!-- space_in_rack(" & maisle & "," & mrack & "," & msku & ") -->" & vbCrLf
        space_in_rack = False
        sSql = "Select rackno from rackpos where sku <= '000'" & _
                " And rackno in (select id from racks where aisle = '" & maisle & "'" & _
                " And rack = '" & mrack & "')"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then space_in_rack = True
        ds.Close()
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub spt_to_dock(ByVal m As ptask)
        Dim sSql As String
        Dim p As ptask, i As Long
        Dim s As String, t As String
        Dim ds As ADODB.Recordset
        On Error GoTo vberror
        'tracelist = tracelist & "<!-- spt_to_dock() -->" & vbCrLf
        t = "ORDER PICK"
        SPTarget = m.target
        sSql = "Select palletid From paltasks Where area = 'DOCK' And status = 'PEND'"
        sSql = sSql & " And palletid = '" & m.palletid & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            s = m.palletid & " aleady at DOCK."
            Call debug_log(s, m, m.userid)
            ds.Close()
            Exit Sub
        End If
        ds.Close()
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
                Wdb.Execute(sSql)
            End If
            s = wrapdesc(Trim(Mid(m.product, 1, 4)), Val(m.units) + Val(m.units2))
            If Len(s) > 1 Then
                p.product = Trim(Mid(m.product, 1, 4)) & s & " " & StrConv(Mid(m.product, 5, Len(m.product) - 4), vbProperCase)
            Else
                p.product = m.product
            End If
            p.palletid = m.palletid
            If Val(m.units2) > 0 Then
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
        eno = Err.Number : edesc = Err.description
        'Call vb_elog(eno, edesc, "wmsmobile.bas", "spt_to_dock", wduserid)
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub spt_to_group(ByVal m As ptask)
        Dim sSql As String, sCols As String, sRows As String
        Dim zid As Long
        Dim ds As ADODB.Recordset
        On Error GoTo vberror
        'tracelist = tracelist & "<!-- spt_to_group(" & m.id & "," & m.target & "," & m.product & ") -->" & vbCrLf
        sSql = "Select id,description From paltasks Where area = 'DOCK'"
        sSql = sSql & " And description > '0 '"
        sSql = sSql & " And target = '" & m.target & "'"
        sSql = sSql & " And product >= '" & Trim(Mid(m.product, 1, 4)) & "'"
        sSql = sSql & " And product < '" & Trim(Mid(m.product, 1, 4)) & "ZZZZ'"
        sSql = sSql & " And source <> 'ALT'"
        sSql = sSql & " And status = 'PEND'"
        sSql = sSql & " And userid < '0'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            zid = ds.Fields(0).Value
            m.description = ds.Fields(1).Value
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
            Wdb.Execute(sSql)
        End If
        ds.Close()
        m.status = "COMP"
        m.trandate = Format(Now, "yyMMdd HH:mm:ss")
        Call post_ship_trans(m)
        Call update_trans(m)
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function sr_receiving(ByVal sr As String, ByVal mbarcode As String) As Boolean
        Dim sSql As String, sCols As String, sRows As String
        Dim ds As ADODB.Recordset
        On Error GoTo vberror
        sSql = "Select id From prodrcv Where sku = '" & Trim(Mid(mbarcode, 1, 4)) & "'"
        sSql = sSql & " And " & sr & " > 0"
        sSql = sSql & " And lot_num = '" & barcode_to_lotnum(mbarcode) & "'"
        'tracelist = tracelist & "<!-- sr_receiving(" & sr & "," & mbarcode & ")" & vbCrLf
        'tracelist = tracelist & sSql & " -->" & vbCrLf
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            sr_receiving = True
        Else
            sr_receiving = False
        End If
        ds.Close()
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function tag_alternate(ByVal psource As String, ByVal puser As String) As String
        Dim sSql As String, sCols As String, sRows As String
        Dim p As ptask, d As ptask
        Dim nsrc As String, psku As String
        Dim i As Long, s As String
        Dim ds As ADODB.Recordset, rs As ADODB.Recordset
        On Error GoTo vberror
        'tracelist = tracelist "(!-- tag_alternate() -->" & vbCrLf
        If psource = "ALTERNATES" Then
            tag_alternate = "<!-- No alternates are defined.. -->"
            Exit Function
        End If
        nsrc = "..."
        sSql = "Select * From paltasks Where id = " & Mid(psource, Len(psource) - 4, 5)
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            d.id = ds.Fields(0).Value
            d.area = ds.Fields(1).Value
            d.description = ds.Fields(2).Value
            d.source = ds.Fields(3).Value
            d.target = ds.Fields(4).Value
            d.product = ds.Fields(5).Value
            d.palletid = ds.Fields(6).Value
            d.qty = ds.Fields(7).Value
            d.uom = ds.Fields(8).Value
            d.lotnum = ds.Fields(9).Value
            d.units = ds.Fields(10).Value
            d.lotnum2 = ds.Fields(11).Value
            d.units2 = ds.Fields(12).Value
            d.status = ds.Fields(13).Value
            d.userid = ds.Fields(14).Value
            d.trandate = ds.Fields(15).Value
            d.reqid = ds.Fields(16).Value

            psku = Trim(Mid(d.product, 1, 4))
            sSql = "Select aisle,rack,fo,lot_num From racks"
            sSql = sSql & " Where sku = '" & psku & "'"
            sSql = sSql & " And aisle <> 'M' And resv_sku <> 'ALL'"
            sSql = sSql & " And hold <> 1"
            sSql = sSql & " And id in (Select rackno From rackpos"
            sSql = sSql & " Where sku = '" & psku & "'"
            sSql = sSql & " And count_qty > 0)"
            sSql = sSql & " Order By fo Desc, lot_num"
            rs = Wdb.Execute(sSql)
            If rs.BOF = False Then
                rs.MoveFirst()
                nsrc = "STAGING"
                p.area = "FORKLIFT"
                p.description = d.description & Space(8 - Len(d.description)) & d.target
                p.source = rs.Fields(0).Value & "-"
                p.source = psource & Trim(rs.Fields(1).Value)
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
            rs.Close()

            If nsrc = "..." And SRFlag = True Then
                sSql = "Select whse_num, lot_num From lane"
                sSql = sSql & " Where sku = '" & psku & "'"
                sSql = sSql & " And lane_status <= ' '"
                sSql = sSql & " Order By lot_num, whse_num"
                rs = Wdb.Execute(sSql)
                If rs.BOF = False Then
                    rs.MoveFirst()
                    nsrc = "SR" & rs.Fields(0).Value
                End If
                rs.Close()
                sSql = "Select id, ship_status From ship_infc Where order_num = '" & d.description & "'"
                sSql = sSql & " And sku = '" & psku & "'"
                sSql = sSql & " And to_whse_num = " & Mid(nsrc, 3, 1)
                rs = Wdb.Execute(sSql)
                If rs.BOF = False Then
                    rs.MoveFirst()
                    i = rs.Fields(0).Value
                    s = rs.Fields(1).Value
                    If s = "DONE" Or s = "CANC" Then
                        sSql = "Update ship_infc Set order_qty = 1, ship_uom_qty = 0"
                        sSql = sSql & ", ship_plt_qty = 0, ship_status = 'NEW'"
                        sSql = sSql & " Where id = " & i
                        Wdb.Execute(sSql)
                    Else
                        sSql = "Update ship_infc Set order_qty = order_qty + 1"
                        sSql = sSql & " Where id = " & i
                        Wdb.Execute(sSql)
                    End If
                Else
                    rs.Close()
                    sSql = "Select id, ship_status From ship_infc"
                    sSql = sSql & " Where ship_status in ('CANC','DONE')"
                    sSql = sSql & " Order By id"
                    rs = Wdb.Execute(sSql)
                    If rs.BOF = False Then
                        rs.MoveFirst()
                        i = rs.Fields(0).Value
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
                        Wdb.Execute(sSql)
                    End If
                    rs.Close()
                End If
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
        ds.Close()
        tag_alternate = "<!-- alternate assigned -->"
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub tlw_to_tml(ByVal m As ptask)
        Dim sSql As String, sCols As String, sRows As String
        Dim p As ptask, i As Long, s As String
        Dim ds As ADODB.Recordset
        On Error GoTo vberror
        'tracelist = tracelist & "<!-- tlw_to_tml() -->" & vbCrLf
        sSql = "Select palletid From paltasks Where source in ('TRAFFIC MASTER','RC119')"
        sSql = sSql & " and palletid = '" & m.palletid & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            s = ds.Fields(0).Value
            s = s & " is already at " & m.target & "."
            Call debug_log(s, m, m.userid)
            ds.Close()
            Exit Sub
        End If
        ds.Close()
        m.status = "COMP"
        m.trandate = Format(Now, "yyMMdd HH:mm:ss")
        If UCase(m.uom) = "PALLET" Or m.target = "RC119" Then
            Call post_move_trans(m)
        Else
            Call post_recv_trans(m)
        End If
        p.area = "TRAFFIC MASTER"
        p.description = " "
        p.source = m.source
        p.target = "..."
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
        p.reqid = m.reqid
        i = insert_trans(p)
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub update_ante_room(ByVal m As ptask)
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

    Public Sub update_op_rack(ByVal m As ptask)
        'tracelist = tracelist & "<!-- update_op_rack(" & m.palletid & "," & m.lotnum & ") -->" & vbCrLf
        Call insert_rack_pallet(m)
        m.status = "COMP"
        m.trandate = Format(Now, "yyMMdd HH:mm:ss")
        Call post_move_trans(m)
        Call update_trans(m)
    End Sub

    Public Sub update_trans(ByVal pt As ptask)
        Dim sSql As String
        On Error GoTo vberror

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

        Wdb.Execute(sSql)
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Sub vb_elog(ByVal eno As Long, ByVal edesc As String, ByVal pform As String, ByVal psub As String, ByVal puser As String)
        Dim i As Integer, s As String, cfile As String
        On Error GoTo vberror
        'cfile = "S:\wd\test\truckerrors.txt"
        cfile = vberror_log
        'i = FreeFile(1)
        i = 88
        FileOpen(i, cfile, OpenMode.Append, OpenAccess.Default, OpenShare.Shared)
        WriteLine(i, eno, edesc, pform, psub, Format(Now, "M-d-yyyy h:mm am/pm"), puser)
        FileClose(i)
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.Description
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

    Public Function wdempname(ByVal empid As String) As String
        Dim s As String, ds As ADODB.Recordset
        s = "Select listdisplay from valuelists where listname = 'wdempid'"
        s = s & " and listreturn = '" & empid & "'"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            wdempname = ds.Fields(0).Value
        Else
            wdempname = empid
        End If
        ds.Close()
    End Function

    Public Function wd_seq(ByVal tbname As String) As Long
        Dim sSql As String
        Dim i As Long
        Dim ds As ADODB.Recordset
        On Error GoTo vberror
        sSql = "Select sequence_id From sequences where seq = '" & tbname & "'"
        ds = Wdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst()
            i = ds.Fields(0).Value + 1
            sSql = "Update sequences Set sequence_id = " & i & " Where seq = '" & tbname & "'"
            Wdb.Execute(sSql)
        Else
            i = 100
            sSql = "Insert Into sequences (sequence_id, seq) Value (" & i & ",'" & tbname & "')"
        End If
        'tracelist = tracelist & "<!-- wd_seq(" & tbname & ") = " & i & " -->" & vbCrLf
        ds.Close() ': db.Close
        wd_seq = i
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.description
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
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

    Public Function wrapdesc(ByVal psku As String, ByVal punits As Integer)
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

End Module
