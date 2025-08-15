Public Class EForklift4
    Private Function barcode_to_lotnum(ByVal bc As String) As String
        Dim s1 As String, s2 As String, s As String, j As Long
        If Len(bc) <> 16 Then
            barcode_to_lotnum = "01001"
        Else
            s = " "
            j = Val(Mid(bc, 5, 2))
            If j < 1 Or j > 12 Then s = "01001"
            j = Val(Mid(bc, 7, 2))
            If j < 1 Or j > 31 Then s = "01001"
            j = Val(Mid(bc, 9, 2))
            If j < 11 Or j > 44 Then s = "01001"
            If s <> "01001" Then
                s1 = "01-01-20" & Format(CInt(Mid(bc, 9, 2)) - 2, "00")
                s2 = Mid(bc, 5, 2) & "-" & Mid(bc, 7, 2) & "-20" & Format(CInt(Mid(bc, 9, 2)) - 2, "00")
                j = DateDiff("d", s1, s2) + 1
                s = Format(CInt(Mid(bc, 9, 2)) - 2, "00")
                s = s & Format(j, "000")
            End If
            barcode_to_lotnum = s
        End If
    End Function

    Private Sub complete_efl_movetask()
        Dim m As ptask, k As Long, efltarget As String
        If Mid(Combo1.SelectedItem, 1, 8) = "Override" Then
            efltarget = Combo3.SelectedItem
        Else
            efltarget = Combo1.SelectedItem
        End If

        If Val(taskkey.Text) > 0 Then    'Pick Up Tasks
            m = masterec(Val(taskkey.Text))

            If efltarget = "Return to Wrapper" Then
                Call return_to_wrapper(m.palletid, WDUserId, m.area, m.reqid)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
                Exit Sub
            End If

            'Crane Finished Good Lanes
            If UCase(efltarget) = "CRANE 1" Or UCase(efltarget) = "CRANE 2" Or UCase(efltarget) = "CRANE 3" Or UCase(efltarget) = "CRANE 5" Or UCase(efltarget) = "CRANE 6" Or efltarget = "TRI LEVEL" Or efltarget = "RC119" Then
                m.target = efltarget
                m.palletid = UCase(Text1.Text)
                m.qty = "1"
                m.uom = "Pallet"
                m.status = "COMP"
                If efltarget = "TRI LEVEL" Or efltarget = "RC119" Then
                    Call tlw_to_tml(m)
                Else
                    Call crane_finished_goods_lane(m)
                End If
                If m.area = "EFLMove" Then Call remove_rack_pallet(m)
                If efltarget = "RC119" And (m.source = "STAGING" Or m.source = "WRAPPER" Or m.source = "BACKHAUL") Then
                    m.status = "COMP"
                    m.userid = " "
                    Call update_trans(m)
                End If
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = ""
                'refresh_task_set_pallets
                Text1.Select()
                Exit Sub
            End If


            If efltarget = "ANTE ROOM" Or efltarget = "M-ANTE" Then
                'If m.description > "." And UCase(Mid(m.description, 1, 4)) <> "DROP" Then
                'Else
                m.target = "ANTE ROOM"
                'End If
                m.palletid = Text1.Text
                m.userid = WDUserId
                If remove_rack_pallet(m) = True Then
                    m = masterec(Val(taskkey.Text))
                    m.target = "ANTE ROOM"
                    m.userid = WDUserId
                End If
                Call update_ante_room(m)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
                Exit Sub
            End If

            If efltarget = "ORDER PICK" Or efltarget = "M-OP" Or efltarget = "M OP" Then
                m.target = "ORDER PICK"
                m.palletid = Text1.Text
                m.userid = WDUserId
                If remove_rack_pallet(m) = True Then
                    m = masterec(Val(taskkey))
                    m.target = "ORDER PICK"
                    m.userid = WDUserId
                End If
                Call update_op_rack(m)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
                Exit Sub
            End If

            If (m.area = "FORKLIFT" And m.description < "0") Or m.area = "EFLMove" Then
                m.target = efltarget
                m.palletid = Text1.Text
                m.userid = WDUserId
                Call efl_rack_moves(m)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
                Exit Sub
            End If

        End If

        If Val(rackkey.Text) > 0 Then
            m = move_task_rack(rackkey.Text, Text1.Text, efltarget, WDUserId)

            If efltarget = "Return to Wrapper" Then
                Call return_to_wrapper(m.palletid, WDUserId, m.area, m.reqid)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
                Exit Sub
            End If

            'Crane Finished Good Lanes - moved here:  jv111314
            If UCase(efltarget) = "CRANE 1" Or UCase(efltarget) = "CRANE 2" Or UCase(efltarget) = "CRANE 3" Or UCase(efltarget) = "CRANE 5" Or UCase(efltarget) = "CRANE 6" Or efltarget = "TRI LEVEL" Or efltarget = "RC119" Then
                'MsgBox efltarget
                m.target = efltarget
                m.palletid = UCase(Text1.Text)
                m.qty = "1"
                m.uom = "Pallet"
                m.status = "COMP"
                If efltarget = "TRI LEVEL" Or efltarget = "RC119" Then
                    Call tlw_to_tml(m)
                Else
                    Call crane_finished_goods_lane(m)
                End If
                If m.area = "EFLMove" Then Call remove_rack_pallet(m)
                If efltarget = "RC119" And (m.source = "STAGING" Or m.source = "WRAPPER" Or m.source = "BACKHAUL") Then
                    m.status = "COMP"
                    m.userid = " "
                    Call update_trans(m)
                End If
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = ""
                'refresh_task_set_pallets
                Text1.Select()
                Exit Sub
            End If


            If m.area = "EFLMove" Then
                k = insert_trans(m)
                m.id = k
                Call efl_rack_moves(m)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
            Else
                emess.Text = "Move task has failed to update racks.."
            End If
            Exit Sub
        End If

        If Val(quekey.Text) > 0 Then
            m = move_task_queue(Val(quekey.Text), Text1.Text, efltarget, WDUserId)

            If efltarget = "Return to Wrapper" Then
                Call return_to_wrapper(m.palletid, WDUserId, m.area, m.reqid)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
                Exit Sub
            End If

            'Crane Finished Good Lanes - moved here:  jv111314
            If UCase(efltarget) = "CRANE 1" Or UCase(efltarget) = "CRANE 2" Or UCase(efltarget) = "CRANE 3" Or UCase(efltarget) = "CRANE 5" Or UCase(efltarget) = "CRANE 6" Or efltarget = "TRI LEVEL" Or efltarget = "RC119" Then
                'MsgBox efltarget
                m.target = efltarget
                m.palletid = UCase(Text1.Text)
                m.qty = "1"
                m.uom = "Pallet"
                m.status = "COMP"
                If efltarget = "TRI LEVEL" Or efltarget = "RC119" Then
                    Call tlw_to_tml(m)
                Else
                    Call crane_finished_goods_lane(m)
                End If
                If m.area = "EFLMove" Then Call remove_rack_pallet(m)
                If efltarget = "RC119" And (m.source = "STAGING" Or m.source = "WRAPPER" Or m.source = "BACKHAUL") Then
                    m.status = "COMP"
                    m.userid = " "
                    Call update_trans(m)
                End If
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = ""
                'refresh_task_set_pallets
                Text1.Select()
                Exit Sub
            End If


            If m.area = "EFLMove" Then
                k = insert_trans(m)
                m.id = k
                Call efl_rack_moves(m)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
            Else
                emess.Text = "Move task has failed to update racks.."
            End If
            Exit Sub
        End If


        If Val(palkey.Text) > 0 Then
            m = move_task_pallet(Val(palkey.Text), Text1.Text, efltarget, WDUserId)

            If efltarget = "Return to Wrapper" Then
                Call return_to_wrapper(m.palletid, WDUserId, m.area, m.reqid)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
                Exit Sub
            End If

            'Crane Finished Good Lanes - moved here:  jv111314
            If UCase(efltarget) = "CRANE 1" Or UCase(efltarget) = "CRANE 2" Or UCase(efltarget) = "CRANE 3" Or UCase(efltarget) = "CRANE 5" Or UCase(efltarget) = "CRANE 6" Or efltarget = "TRI LEVEL" Or efltarget = "RC119" Then
                'MsgBox efltarget
                m.target = efltarget
                m.palletid = UCase(Text1.Text)
                m.qty = "1"
                m.uom = "Pallet"
                m.status = "COMP"
                If efltarget = "TRI LEVEL" Or efltarget = "RC119" Then
                    Call tlw_to_tml(m)
                Else
                    Call crane_finished_goods_lane(m)
                End If
                If m.area = "EFLMove" Then Call remove_rack_pallet(m)
                If efltarget = "RC119" And (m.source = "STAGING" Or m.source = "WRAPPER" Or m.source = "BACKHAUL") Then
                    m.status = "COMP"
                    m.userid = " "
                    Call update_trans(m)
                End If
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = ""
                'refresh_task_set_pallets
                Text1.Select()
                Exit Sub
            End If


            If m.area = "EFLMove" Then
                k = insert_trans(m)
                m.id = k
                Call efl_rack_moves(m)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
            Else
                emess.Text = "Move task has failed to update pallet location.."
            End If
            Exit Sub
        End If

        If Val(srkey.Text) > 0 Then
            m = move_task_crane(Val(srkey.Text), Text1.Text, efltarget, WDUserId)

            If efltarget = "Return to Wrapper" Then
                Call return_to_wrapper(m.palletid, WDUserId, m.area, m.reqid)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
                Exit Sub
            End If

            'Crane Finished Good Lanes - moved here:  jv111314
            If UCase(efltarget) = "CRANE 1" Or UCase(efltarget) = "CRANE 2" Or UCase(efltarget) = "CRANE 3" Or UCase(efltarget) = "CRANE 5" Or UCase(efltarget) = "CRANE 6" Or efltarget = "TRI LEVEL" Or efltarget = "RC119" Then
                'MsgBox efltarget
                m.target = efltarget
                m.palletid = UCase(Text1.Text)
                m.qty = "1"
                m.uom = "Pallet"
                m.status = "COMP"
                If efltarget = "TRI LEVEL" Or efltarget = "RC119" Then
                    Call tlw_to_tml(m)
                Else
                    Call crane_finished_goods_lane(m)
                End If
                If m.area = "EFLMove" Then Call remove_rack_pallet(m)
                If efltarget = "RC119" And (m.source = "STAGING" Or m.source = "WRAPPER" Or m.source = "BACKHAUL") Then
                    m.status = "COMP"
                    m.userid = " "
                    Call update_trans(m)
                End If
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = ""
                'refresh_task_set_pallets
                Text1.Select()
                Exit Sub
            End If


            If m.area = "EFLMove" Then
                k = insert_trans(m)
                m.id = k
                Call efl_rack_moves(m)
                Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
            Else
                emess.Text = "Move task has failed to update pallet location.."
            End If
            Exit Sub
        End If

        Text1.Text = "" : Combo1.Items.Clear() : emess.Text = "" : Text1.Select()
    End Sub

    Private Sub draw_label(ByVal bc As String)
        Dim i As Integer
        skupic.Text = Trim(Mid(bc, 1, 4))
        i = Val(skupic.Text)
        lotpic.Text = Mid(bc, 5, 6)                      'jv052515
        oppic.Text = Mid(bc, 11, 3)                      'jv082415
        palnopic.Text = Mid(bc, 14, 3)
        If Val(palnopic.Text) > 0 Then palnopic.Text = Format(Val(palnopic.Text), "0")
        pkgpic.Text = labpix(i).package
        name1pic.Text = labpix(i).name1
        name2pic.Text = labpix(i).name2
        name3pic.Text = labpix(i).name3
        Frame1.Visible = True
        histbc = bc                                 'jv052515
    End Sub

    Private Function sr_receiving(ByVal sr As String, ByVal bc As String) As Boolean
        Dim ds As ADODB.Recordset, s As String
        On Error GoTo vberror
        s = "select id from prodrcv where sku = '" & Trim(Mid(bc, 1, 4)) & "'"
        s = s & " and " & sr & " > 0"
        s = s & " and lot_num = '" & barcode_to_lotnum(bc) & "'"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            sr_receiving = True
        Else
            sr_receiving = False
        End If
        ds.Close()
        Exit Function
vberror:
        eno = Err.Number : edesc = Err.Description : Err.Clear()
        'Call vb_elog(eno, edesc, Me.Name, "sr_receiving", Form1.userid)
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, Me.Name, "sr_receiving", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: sr_receiving: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Function

    Private Sub barcode_scanned(ByVal bc As String)
        Dim ds As ADODB.Recordset, s As String
        Dim tmon As String, tday As String, tyr As String, topc As String, tpal As String, tdate As String
        On Error GoTo vberror
        emess.Text = "" : Combo1.Items.Clear() : taskkey.Text = "0" : rackkey.Text = "0" : palkey.Text = "0"
        srkey.Text = "0" : quekey.Text = "0"
        bc = UCase(bc)
        s = " "
        If Len(bc) <> 16 Then
            s = "Invalid barcode length: " & bc & "."
        Else
            tmon = Mid(bc, 5, 2)
            tday = Mid(bc, 7, 2)
            tyr = Mid(bc, 9, 2)
            'topc = Mid(bc, 12, 1)
            topc = Mid(bc, 11, 3)                                       'jv052515
            tpal = Mid(bc, 14, 3)
            tdate = tmon & "-" & tday & "-20" & tyr
            'If Mid(bc, 4, 1) > " " Then
            '    s = "Invalid character found in barcode: " & Left(bc, 3) & " _" & Mid(bc, 4, 1) & "_ " & Right(bc, 12) & "."
            'End If
            'If Mid(bc, 11, 1) > " " Then
            '    s = "Invalid character found in barcode: " & Left(bc, 10) & " _" & Mid(bc, 11, 1) & "_ " & Right(bc, 5) & "."
            'End If
            'If Mid(bc, 13, 1) > " " Then
            '    s = "Invalid character found in barcode: " & Left(bc, 12) & " _" & Mid(bc, 13, 1) & "_ " & Right(bc, 3) & "."
            'End If
            If Val(tmon) < 1 Or Val(tmon) > 12 Then
                s = "Invalid Month found in code date: " & Mid(bc, 1, 4) & " [" & tmon & "] " & Mid(bc, 7, 10) & "."
            End If
            If Val(tday) < 1 Or Val(tday) > 31 Then
                s = "Invalid Day found in code date: " & Mid(bc, 1, 6) & " [" & tday & "] " & Mid(bc, 9, 8) & "."
                If Val(tyr) < 12 Or Val(tyr) > 20 Then
                    s = "Invalid Year found in code date: " & Mid(bc, 1, 8) & " [" & tyr & "] " & Mid(bc, 11, 6) & "."
                End If
                If IsDate(tdate) = False Then
                    s = "Invalid code date (" & tdate & "): " & Mid(bc, 1, 4) & " [" & tmon & tday & tyr & "] " & Mid(bc, 11, 6) & "."
                End If
                'If topc < "100" Or topc > "599" Then                                                'jv052515
                '    s = "Invalid Operation Code found in barcode: " & Left(bc, 10) & "_" & topc & "_" & Right(bc, 3) & "."  'jv052515
                'End If                                                                              'jv052515
                'If topc < "A" Or topc > "Z" Then
                '    s = "Invalid Operation Code found in barcode: " & Left(bc, 10) & " _" & topc & "_ " & Right(bc, 3) & "."
                'End If
                If Val(tpal) > 0 Or tpal = "EOR" Then
                Else
                    s = "Invalid Pallet # found in barcode: " & Mid(bc, 1, 12) & " _" & tpal & "_."
                End If
            End If
        End If

        If s > " " Then
            emess.Text = s
            Button1.Enabled = False
            Exit Sub
        End If


        Label2.Text = "BarCode not found."
        s = "select area, source, target, description, id from paltasks where palletid = '" & bc & "'"
        's = s & " and userid < '0'"
        s = s & " and status = 'PEND'"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            If Trim(ds.Fields(0).Value) <> "FORKLIFT" Then
                emess.Text = "PENDING " & ds.Fields(0).Value & " task for this barcode." & vbCrLf & ds.Fields(1).Value & " >> " & ds.Fields(2).Value & " " & ds.Fields(3).Value
                Label2.Text = "Task "
            Else
                Label2.Text = ds.Fields(1).Value & " >> " & ds.Fields(2).Value & " " & ds.Fields(3).Value
                taskkey.Text = ds.Fields(4).Value
            End If
        End If
        ds.Close()

        If Label2.Text = "BarCode not found." Then
            s = "select aisle, rack, id from racks where id in"
            s = s & " (select rackno from rackpos where barcode = '" & bc & "')"
            ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst()
                Label2.Text = "Rack: " & ds.Fields(0).Value & "-" & ds.Fields(1).Value
                rackkey.Text = ds.Fields(2).Value
            End If
            ds.Close()
        End If

        If Label2.Text = "BarCode not found." Then
            s = "select whse_num, queue_num, id from queue_infc where palletid = '" & bc & "' order by queue_num desc"
            ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst()
                Label2.Text = "Crane " & ds.Fields(0).Value & " Queue # " & ds.Fields(1).Value
                quekey = ds.Fields(2).Value
            End If
            ds.Close()
        End If

        If Label2.Text = "BarCode not found." Then
            s = "select id from pallets where barcode = '" & bc & "'"
            ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst()
                Label2.Text = "Pallet"
                palkey.Text = ds.Fields(0).Value
            End If
            ds.Close()
        End If

        If Label2.Text = "BarCode not found." Then
            s = "select whse_num, vert_loc, horz_loc, rack_side, posn_num, id from position"
            s = s & " where barcode = '" & bc & "'"
            ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst()
                Label2.Text = "SR-" & ds.Fields(0).Value & " " & ds.Fields(1).Value & "-" & ds.Fields(2).Value & "-" & ds.Fields(3).Value & "-" & ds.Fields(4).Value
                srkey.Text = ds.Fields(5).Value
            End If
            ds.Close()
        End If

        If Label2.Text = "BarCode not found." Then
            emess.Text = "Barcode: " & Text1.Text & " is not found."
            Combo1.Items.Clear()
            Text1.Text = ""
            Button1.Enabled = False
        Else
            If emess.Text > " " Then
                Button1.Enabled = False
            Else
                refresh_target_racks()
                Button1.Enabled = True
                Call draw_label(Text1.Text)
            End If
        End If
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.Description : Err.Clear()
        'Call vb_elog(eno, edesc, Me.Name, "barcode_scanned", Form1.userid)
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, Me.Name, "barcode_scanned", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: barcode_scanned: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Sub

    Private Sub refresh_products()
        Dim ds As ADODB.Recordset, s As String
        Combo2.Items.Clear()
        On Error GoTo vberror
        s = "select sku, uom_type, description from sku_config"
        s = s & " where sku in (select sku from rackpos where rackno not in (select id from racks where rack in('OP','SP') or hold <> 0))"
        s = s & " order by sku"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                Combo2.Items.Add(ds.Fields(0).Value & " " & ds.Fields(1).Value & " " & ds.Fields(2).Value)
                ds.MoveNext()
            Loop
            Combo2.SelectedIndex = 0
        End If
        ds.Close()
        s = "select distinct aisle from racks where aisle <> 'M' order by aisle"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                Combo2.Items.Add("Aisle-" & ds.Fields(0).Value & " open positions")
                ds.MoveNext()
            Loop
        End If
        ds.Close()
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.Description : Err.Clear()
        'Call vb_elog(eno, edesc, Me.Name, "refresh_products", Form1.userid)
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, Me.Name, "refresh_products", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: refresh_products: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Sub

    Private Sub refresh_rack_locations()
        Dim ds As ADODB.Recordset, s As String, msku As String
        Combo3.Items.Clear()
        On Error GoTo vberror
        If Mid(Combo2.SelectedItem, 1, 5) = "Aisle" Then
            s = "select aisle, rack from racks where aisle = '" & Mid(Combo2.SelectedItem, 7, 1) & "'"
            s = s & " and (qty + qty4) < capacity and resv_sku < '0' and hold = 0"
            s = s & " order by slot"
        Else
            msku = Trim(Mid(Combo2.SelectedItem, 1, 4))
            s = "select aisle, rack from racks where hold = 0" 'False"
            s = s & " and rack not in ('OP','SP')"
            s = s & " and id in (select rackno from rackpos where sku = '" & msku & "')"
            s = s & " order by fo desc, aisle, rack"
        End If
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                Combo3.Items.Add(ds.Fields(0).Value & "-" & Trim(ds.Fields(1).Value))
                ds.MoveNext()
            Loop
            Combo3.SelectedIndex = 0
        End If
        ds.Close()
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.Description : Err.Clear()
        'Call vb_elog(eno, edesc, Me.Name, "refresh_rack_locations", Form1.userid)
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, Me.Name, "refresh_rack_locations", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: refresh_rack_locations: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Sub

    Private Sub refresh_target_racks()
        Dim ds As ADODB.Recordset, s As String, msku As String
        Dim p As ptask                                                  'jv052515
        p.area = " " : p.description = " " : p.id = "0" : p.lotnum = " " : p.lotnum2 = " " : p.palletid = " "
        p.product = " " : p.qty = "0" : p.reqid = "0" : p.source = " " : p.target = " " : p.trandate = " "
        p.units = "0" : p.units2 = "0" : p.uom = " " : p.userid = " " : p.status = " "
        On Error GoTo vberror
        Combo1.Items.Clear() : List1.Items.Clear()
        msku = Trim(Mid(Text1.Text, 1, 4))
        If sr_receiving("sr1", Text1.Text) Then
            Combo1.Items.Add("CRANE 1")
            List1.Items.Add("SR1")
        End If
        If sr_receiving("sr2", Text1.Text) Then
            Combo1.Items.Add("CRANE 2")
            List1.Items.Add("SR2")
        End If
        If sr_receiving("sr3", Text1.Text) Then
            Combo1.Items.Add("TRI LEVEL")
            List1.Items.Add("TMASTER")
        End If
        s = "select aisle, rack, id from racks where resv_sku in ('" & msku & "', 'ALL')"
        s = s & " and (qty + qty4) < capacity"
        s = s & " order by resv_sku, qty desc, aisle, rack"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                Combo1.Items.Add(ds.Fields(0).Value & "-" & Trim(ds.Fields(1).Value))
                List1.Items.Add(ds.Fields(2).Value)
                ds.MoveNext()
            Loop
        End If
        ds.Close()

        s = "select aisle, rack, id from racks where resv_sku <= ' ' and sku = '" & msku & "'"
        s = s & " and hold = 0"
        s = s & " and (qty + qty4) < capacity"
        s = s & " order by resv_sku, qty desc, aisle, rack"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                Combo1.Items.Add(ds.Fields(0).Value & "-" & Trim(ds.Fields(1).Value))
                List1.Items.Add(ds.Fields(2).Value)
                ds.MoveNext()
            Loop
        End If
        ds.Close()
        p.palletid = Text1.Text                                                      'jv052515
        s = "select lot1, lot2 from pallets where barcode = '" & Text1.Text & "'"    'jv052515
        ds = Wdb.Execute(s)                                                 'jv052515
        If ds.BOF = False Then                                                  'jv052515
            ds.MoveFirst()                                                        'jv052515
            p.lotnum = ds.Fields(0).Value                                                  'jv052515
            p.lotnum2 = ds.Fields(1).Value                                                 'jv052515
        Else                                                                    'jv052515
            p.lotnum = barcode_to_lotnum(p.palletid)                            'jv052515
            p.lotnum2 = ""                                                      'jv052515
        End If                                                                  'jv052515
        ds.Close()                                                                'jv052515
        If check_hold(p) = False Then                                           'jv052515
            Combo1.Items.Add("M-OP") : List1.Items.Add("OP")                           'jv052515
        Else                                                                    'jv052515
            Label2.Text = "On Hold! " & Label2.Text                       'jv052515
        End If                                                                  'jv052515
        Combo1.Items.Add("M-ANTE") : List1.Items.Add("ANTE")
        Combo1.Items.Add("Return to Wrapper") : List1.Items.Add("Return")
        If Combo3.Visible = True Then
            Combo1.Items.Add("Override - Goto " & Combo3.SelectedItem)
            List1.Items.Add(Combo3.SelectedItem)
        End If

        If Val(taskkey.Text) = 0 Then
            Combo1.Items.Add("CRANE 1") : List1.Items.Add("SR1")               'jv111314
            Combo1.Items.Add("CRANE 2") : List1.Items.Add("SR2")               'jv111314
            Combo1.Items.Add("TRI LEVEL") : List1.Items.Add("TMASTER")         'jv111314
        End If



        Combo1.SelectedIndex = 0
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.Description : Err.Clear()
        'Call vb_elog(eno, edesc, Me.Name, "refresh_target_racks", Form1.userid)
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, Me.Name, "refresh_target_racks", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: refresh_target_racks: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Sub

    Private Sub Combo1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo1.SelectedIndexChanged
        List1.SelectedIndex = Combo1.SelectedIndex
    End Sub

    Private Sub Combo2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Combo2.SelectedIndexChanged
        refresh_rack_locations()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        complete_efl_movetask()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Combo2.Visible = True Then
            Combo2.Visible = False
            Combo3.Visible = False
            Label4.Visible = False
            Label5.Visible = False
        Else
            Combo2.Visible = True
            Combo3.Visible = True
            Label4.Visible = True
            Label5.Visible = True
            refresh_products()
        End If
    End Sub

    Private Sub EForklift4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Text1.Text = ""
        Combo1.Items.Clear()
        emess.Text = ""
        refresh_products()
    End Sub

    Private Sub EForklift4_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        apphdr.Left = (Me.Width - apphdr.Width) * 0.5
        Button3.Left = (Me.Width - Button3.Width) * 0.5
        apphdr.Left = (Me.Width - apphdr.Width) * 0.5
        Label1.Left = apphdr.Left : Text1.Left = Label1.Left + Label1.Width + 100
        Label2.Left = Text1.Left 'apphdr.Left
        Label3.Left = apphdr.Left : Combo1.Left = Label3.Left + Label3.Width + 100
        Button1.Left = (Me.Width - Button1.Width) * 0.5
        emess.Left = (Me.Width - emess.Width) * 0.5
        Frame1.Left = (Me.Width - Frame1.Width) * 0.5
        Button2.Left = (Me.Width - Button2.Width) * 0.3
        'Command2.Left = apphdr.Left
        Label4.Left = Button2.Left : Combo2.Left = Label4.Left + Label4.Width + 100
        Label5.Left = Button2.Left : Combo3.Left = Label5.Left + Label5.Width + 100
        xit.Left = Me.Width - xit.Width
    End Sub

    Private Sub Frame1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Frame1.Click
        If Button1.Enabled = True Then complete_efl_movetask()
    End Sub

    Private Sub lotpic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lotpic.Click
        If Button1.Enabled = True Then complete_efl_movetask()
    End Sub

    Private Sub name1pic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles name1pic.Click
        If Button1.Enabled = True Then complete_efl_movetask()
    End Sub

    Private Sub name2pic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles name2pic.Click
        If Button1.Enabled = True Then complete_efl_movetask()
    End Sub

    Private Sub name3pic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles name3pic.Click
        If Button1.Enabled = True Then complete_efl_movetask()
    End Sub

    Private Sub oppic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles oppic.Click
        If Button1.Enabled = True Then complete_efl_movetask()
    End Sub

    Private Sub palnopic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles palnopic.Click
        If Button1.Enabled = True Then complete_efl_movetask()
    End Sub

    Private Sub pkgpic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pkgpic.Click
        If Button1.Enabled = True Then complete_efl_movetask()
    End Sub

    Private Sub skupic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles skupic.Click
        If Button1.Enabled = True Then complete_efl_movetask()
    End Sub

    Private Sub Text1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.DoubleClick
        tpad.Text = "BarCode"
        tpad.cname.Text = "Text1"
        tpad.fname.Text = Me.Name
        tpad.trig.Text = Text1.Text
        tpad.Show()
    End Sub

    Private Sub Text1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.GotFocus
        Text1.SelectionStart = 0
        Text1.SelectionLength = Len(Text1.Text)
        Text1.BackColor = Frame1.BackColor
    End Sub

    Private Sub Text1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.LostFocus
        Text1.BackColor = Me.BackColor
    End Sub

    Private Sub Text1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.TextChanged
        Frame1.Visible = False
        If Len(Text1.Text) > 15 Then
            Call barcode_scanned(Text1.Text)
        Else
            Button1.Enabled = False
        End If
    End Sub

    Private Sub xit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles xit.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Text1.Select()
    End Sub
End Class