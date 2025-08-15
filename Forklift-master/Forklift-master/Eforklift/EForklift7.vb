Public Class EForklift7
    Private Sub draw_label(ByVal bc As String)
        Dim i As Integer
        bc = UCase(bc)
        skupic.Text = Trim(Mid(bc, 1, 4))
        i = Val(skupic.Text)
        'lotpic = Mid(bc, 5, 8)
        lotpic.Text = Mid(bc, 5, 6)                          'jv052515
        oppic.Text = Mid(bc, 11, 3)                          'jv082415
        palnopic.Text = Mid(bc, 14, 3)
        'MsgBox(bc & " " & palnopic.Text)
        If Val(palnopic.Text) > 0 Then palnopic.Text = Format(Val(palnopic.Text), "0")
        pkgpic.Text = labpix(i).package
        name1pic.Text = labpix(i).name1
        name2pic.Text = labpix(i).name2
        name3pic.Text = labpix(i).name3
        Frame1.Visible = True
        histbc = bc                                     'jv052515
    End Sub

    Private Sub record_pallet(ByVal pno As String, ByVal p As ptask, ByVal pwhs As String, ByVal pstat As String)
        Dim ds As ADODB.Recordset, s As String
        Dim pid As Long, psku As String, recid As Long
        'Screen.MousePointer = 11
        psku = Trim(Mid(p.palletid, 1, 4))
        recid = 0
        On Error GoTo vberror

        s = "select id from pallets where barcode = '" & p.palletid & "'"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            recid = ds.Fields(0).Value
        Else
            ds.Close()
            s = "select id from pallets where status in ('Shipped','Order Pick')"
            s = s & " order by trandate"
            ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst()
                recid = ds.Fields(0).Value
            End If
        End If
        ds.Close()
        If recid > 0 Then
            s = "Update pallets set plateno = '" & Format(Val(pno), "000000") & "'"
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
            Wdb.Execute(s)
        Else
            pid = wd_seq("Pallets")
            s = "Insert Into pallets Values (" & pid
            s = s & ",'" & Format(Val(pno), "000000") & "'"
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
            Wdb.Execute(s)
        End If
        'Screen.MousePointer = 0
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description : Err.Clear()
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, Me.Name, "record_pallet", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: record_pallet: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Sub


    Private Sub barcode_scanned(ByVal bc As String)
        Dim ds As ADODB.Recordset, s As String
        Dim i As Integer, cd As String, ssku As String, cc As String, td As String
        If Len(bc) < 15 Then Exit Sub
        On Error GoTo vberror
        'Test for previous scan
        If Combo1.Items.Count > 0 Then
            For i = 0 To Combo1.Items.Count - 1
                'If Left(Combo1.List(i), 16) = bc Then
                If Mid(Combo1.Items.Item(i), 1, 16) = bc Then
                    emess.Text = bc & " has already been scanned."
                    Combo1.SelectedIndex = i
                    Text1.Text = ""
                    Text1.Select()
                    Exit Sub
                End If
                If Len(Text4) = 6 Then
                    'If Mid(Combo1.List(i), 18, 6) = Format(Val(Text4), "000000") Then
                    If Mid(Combo1.Items.Item(i), 18, 6) = Format(Val(Text4.Text), "000000") Then
                        emess.Text = "Plate " & Text4.Text & " has already been scanned."
                        Combo1.SelectedIndex = i
                        Text4.Text = ""
                        Text4.Select()
                        Exit Sub
                    End If
                End If
            Next i
        End If

        'Build 2nd code date list
        cd = Mid(bc, 5, 2) & "-" & Mid(bc, 7, 2) & "-20" & Mid(bc, 9, 2)
        td = Format(DateAdd("yyyy", 2, Now), "MM-dd-yyyy")
        Combo2.Items.Clear()
        Combo2.Items.Add(" ")
        For i = 1 To DateDiff("d", cd, td)
            Combo2.Items.Add(Format(DateAdd("d", i, cd), "MMddyy"))
        Next i

        'Check sku and get wrap qty
        ssku = Trim(Mid(bc, 1, 4))
        wrapspal.Text = "0"
        unitspal.Text = "0"
        unitswrap.Text = "0"
        s = "select uom_per_pallet, qty_per_pallet, uom_type, description from sku_config where sku = '" & ssku & "'"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            wrapspal.Text = ds.Fields(1).Value      '!qty_per_pallet
            unitspal.Text = ds.Fields(0).Value           '!uom_per_pallet
            unitswrap.Text = ds.Fields(0).Value / ds.Fields(1).Value  '!qty_per_pallet
            Text2.Text = ds.Fields(1).Value              '!qty_per_pallet
            emess.Text = ds.Fields(2).Value & " " & ds.Fields(3).Value
            Text3.Text = ""
            Text4.Select()
        Else
            emess.Text = "Invalid SKU: " & ssku & " found in the barcode."
            ds.Close()
            Text1.Text = ""
            Text2.Text = ""
            Text3.Text = ""
            Text1.Select()
            Exit Sub
        End If
        ds.Close()
        refresh_target_racks()
        Button1.Enabled = True
        cc = Mid(bc, 11, 3)
        Call draw_label(Text1.Text)
        For i = 0 To Combo4.Items.Count - 1
            If Combo4.Items.Item(i) = cc Then
                Combo4.SelectedIndex = i
                Exit For
            End If
        Next i
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description : Err.Clear()
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

    Private Sub refresh_rollerbed_pallets()
        Dim ds As ADODB.Recordset, s As String
        Combo1.Items.Clear() : List1.Items.Clear()
        On Error GoTo vberror
        s = "select palletid, qty, uom, id, reqid, target from paltasks where source = 'ROLLER BED'"
        s = s & " order by trandate desc"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                Combo1.Items.Add(ds.Fields(0).Value & " " & Format(Val(ds.Fields(4).Value), "000000") & " " & ds.Fields(1).Value & " " & ds.Fields(2).Value & " " & StrConv(ds.Fields(5).Value, vbProperCase))
                List1.Items.Add(ds.Fields(3).Value)
                ds.MoveNext()
            Loop
            Combo1.SelectedIndex = 0
        End If
        ds.Close()
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description : Err.Clear()
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, Me.Name, "refresh_rollerbed_pallets", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: refresh_rollerbed_pallets: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Sub

    Private Sub refresh_target_racks()
        Dim ds As ADODB.Recordset, s As String, msku As String
        On Error GoTo vberror
        Combo3.Items.Clear() : List3.Items.Clear()
        msku = Trim(Mid(Text1.Text, 1, 4))

        'If sr_receiving("sr1", Text1) Then
        '    Combo3.AddItem "Crane 1"
        '    List3.AddItem "SR1"
        'End If
        'If sr_receiving("sr2", Text1) Then
        '    Combo3.AddItem "Crane 2"
        '    List3.AddItem "SR2"
        'End If
        'If sr_receiving("sr3", Text1) Then
        '    Combo3.AddItem "Tri Level"
        '    List3.AddItem "TMASTER"
        'End If

        s = "select aisle, rack, id from racks where resv_sku in ('" & msku & "', 'ALL')"
        s = s & " and (qty + qty4) < capacity"
        s = s & " order by resv_sku, qty desc, aisle, rack"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                Combo3.Items.Add(ds.Fields(0).Value & "-" & Trim(ds.Fields(1).Value))
                List3.Items.Add(ds.Fields(2).Value)
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
                Combo3.Items.Add(ds.Fields(0).Value & "-" & Trim(ds.Fields(1).Value))
                List3.Items.Add(ds.Fields(2).Value)
                ds.MoveNext()
            Loop
        End If
        ds.Close()


        Combo3.Items.Add("M-OP") : List3.Items.Add("OP")
        'Combo3.AddItem "M-ANTE": List3.AddItem "ANTE"
        Combo3.SelectedIndex = 0
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description : Err.Clear()
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

    Private Sub refresh_plate(ByVal pno As Long)
        Dim ds As ADODB.Recordset, s As String
        On Error GoTo vberror
        If pno = 0 Then
            s = "select sequence_id from sequences where seq = 'RBBarcode'"
            ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst()
                Label8.Text = Format(ds.Fields(0).Value + 1, "000000")
            End If
            ds.Close()
        Else
            s = "update sequences set sequence_id = " & pno & " where seq = 'RBBarcode'"
            Wdb.Execute(s)
            Label8.Text = Format(pno + 1, "000000")
        End If
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.description : Err.Clear()
        'Call vb_elog(eno, edesc, Me.Name, "refresh_plate", Form1.userid)
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, Me.Name, "refresh_plate", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: refresh_plate: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Sub

    Private Sub Combo1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        List1.SelectedIndex = Combo1.SelectedIndex
    End Sub

    Private Sub Combo2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo2.SelectedIndexChanged
        If Combo2.SelectedItem > " " Then Text3.Text = Val(wrapspal.Text) - Val(Text2.Text)
    End Sub

    Private Sub Combo3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo3.SelectedIndexChanged
        List3.SelectedIndex = Combo3.SelectedIndex
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Frame1.Click
        Dim p As ptask, s As String, psku As String, wcnt As Integer
        If Button1.Enabled = False Then Exit Sub
        If Len(Text4.Text) <> 6 Then
            MsgBox("Invalid or Missing Plate...", vbOKOnly + vbExclamation, "Sorry, try again...")
            Text4.Select()
            Exit Sub
        End If
        'Test for previous scan
        If Combo1.Items.Count > 0 Then
            For i = 0 To Combo1.Items.Count - 1
                If Mid(Combo1.Items.Item(i), 1, 16) = Text1.Text Then
                    emess.Text = Text1.Text & " has already been scanned."
                    Combo1.SelectedIndex = i
                    Text1.Text = ""
                    Text1.Select()
                    Exit Sub
                End If
                If Mid(Combo1.Items.Item(i), 18, 6) = Format(Val(Text4.Text), "000000") Then
                    emess.Text = "Plate " & Text4.Text & " has already been scanned."
                    Combo1.SelectedIndex = i
                    Text4.Text = ""
                    Text4.Select()
                    Exit Sub
                End If
            Next i
        End If
        s = Text1.Text & " " & Text4.Text & " " & Text2.Text & " Wraps NEW"
        Combo1.Items.Insert(0, s)
        'MsgBox s
        psku = Trim(Mid(Text1.Text, 1, 4))
        p.area = "ROLLER BED"
        p.description = " "
        p.source = "ROLLER BED"
        p.target = Combo3.SelectedItem
        p.product = psku & " " & sku_info(psku, "desc")
        p.palletid = Text1.Text
        p.qty = Val(Text2.Text) + Val(Text3.Text)
        p.uom = "Wraps"
        p.lotnum = barcode_to_lotnum(Text1.Text)
        wcnt = sku_info(psku, "units")
        wcnt = wcnt / sku_info(psku, "wraps")
        p.units = Val(Text2) * wcnt
        If Combo2.SelectedItem > " " Then
            s = Mid(Text1.Text, 1, 4) & Combo2.SelectedItem & Mid(Text1.Text, 11, 6)
            'p.lotnum2 = barcode_to_lotnum(s) & " " & Combo4
            p.lotnum2 = barcode_to_lotnum(s) & Combo4.SelectedItem               'jv052515
            p.units2 = Val(Text3.Text) * wcnt
        Else
            p.lotnum2 = " "
            p.units2 = "0"
        End If
        p.status = "PEND"
        p.userid = WDUserId  '"131052"
        p.trandate = Format(Now, "yyMMdd hh:mm:ss")
        p.reqid = Text4.Text
        p.id = insert_trans(p)
        Call roller_bed_pickup(p)
        p.status = "COMP"
        p.userid = " "
        p.reqid = Text4.Text
        Call update_trans(p)

        Call record_pallet(Text4.Text, p, Combo3.SelectedItem, "Wrapper")
        refresh_rollerbed_pallets()
        emess.Text = ""
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Combo2.SelectedIndex = 0
        Call refresh_plate(Val(Text4.Text))
        Text4.Text = ""
        Text1.Select()
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim s As String, i As Integer, p As ptask
        s = " "
        If Combo1.SelectedItem > "0" Then s = Mid(Combo1.SelectedItem, 1, 16)
        If Val(List1) > 0 Then
            p = masterec(Val(List1))
            Call return_to_wrapper(p.palletid, WDUserId, p.area, p.reqid)
        End If
        If Combo1.Items.Count > 1 Then
            i = Combo1.SelectedIndex
            Combo1.Items.RemoveAt(i)
            'List1.RemoveItem(i)
            List1.Items.RemoveAt(i)
            'DoEvents()
            Combo1.SelectedIndex = 0
        Else
            Combo1.Items.Clear()
            List1.Items.Clear()
        End If
        Text1.Text = s
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Text4.Text = Label8.Text
    End Sub

    Private Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        refresh_rollerbed_pallets()
    End Sub

    Private Sub EForklift7_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        emess.Text = ""
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Combo3.Items.Clear()
        Combo4.Items.Clear()
        For i = 500 To 599                  'jv052515
            Combo4.Items.Add(i)                'jv052515
        Next i                              'jv052515
        Combo4.SelectedIndex = 0
        'apphdr = "ROLLER BED Pallet"
        refresh_rollerbed_pallets()
        Call refresh_plate(0)
    End Sub

    Private Sub EForklift7_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Combo1.Left = (Me.Width - Combo1.Width) * 0.5
        apphdr.Left = Combo1.Left
        Label1.Left = Combo1.Left
        Text1.Left = Label1.Left + Label1.Width

        Label6.Left = Combo1.Left
        Text4.Left = Label6.Left + Label6.Width
        Button3.Left = Text4.Left + Text4.Width + Button1.Width ' + 600
        'Label8.Left = Text4.Left + Text4.Width
        Label8.Left = Button3.Left + Button3.Width

        Label7.Left = Combo1.Left
        Combo3.Left = Label7.Left + Label7.Width

        Label2.Left = Combo1.Left
        Text2.Left = Label2.Left + Label2.Width

        Label3.Left = Combo1.Left
        Combo2.Left = Label3.Left + Label3.Width
        Label5.Left = Combo2.Left + Combo2.Width
        Combo4.Left = Label5.Left + Label5.Width


        Label4.Left = Combo1.Left
        Text3.Left = Label4.Left + Label4.Width
        'Label5.Left = Combo1.Left
        Button4.Left = Combo1.Left
        Button1.Left = (Me.Width - Button1.Width) * 0.5
        Button2.Left = (Me.Width - Button2.Width) * 0.5
        emess.Left = (Me.Width - emess.Width) * 0.5
        xit.Left = Me.Width - xit.Width
        Frame1.Left = (Me.Width - Frame1.Width) * 0.5
    End Sub

    Private Sub Text1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.DoubleClick
        tpad.Text = "BarCode"
        tpad.cname.Text = "Text1"
        tpad.fname.Text = Me.Name
        tpad.trig.Text = Text1.Text 'Val(tpad.trig) + 1
        tpad.Show()
    End Sub

    Private Sub Text1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.GotFocus
        Text1.SelectionStart = 0
        Text1.SelectionLength = Len(Text1.Text)
    End Sub

    Private Sub Text1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text1.TextChanged
        Frame1.Visible = False
        emess.Visible = False
        If Len(Text1.Text) > 15 Then
            Text1.Text = UCase(Text1.Text)
            Call barcode_scanned(Text1.Text)
            Button3.Select()
        Else
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Text2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text2.Click
        tpad.Text = "Qty"
        tpad.cname.Text = "Text2"
        tpad.fname.Text = Me.Name
        tpad.trig.Text = Text3.Text 'Val(tpad.trig) + 1
        tpad.Show()
    End Sub

    Private Sub Text2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text2.TextChanged
        If Combo2.SelectedItem > " " Then Text3.Text = Val(wrapspal.Text) - Val(Text2.Text)
    End Sub

    Private Sub Text3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text3.Click
        tpad.Text = "Qty"
        tpad.cname.Text = "Text3"
        tpad.fname.Text = Me.Name
        tpad.trig.Text = Text3.Text 'Val(tpad.trig) + 1
        tpad.Show()
    End Sub

    Private Sub Text3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text3.TextChanged
        If Combo2.SelectedItem > " " Then
            Text2.Text = Val(wrapspal.Text) - Val(Text3.Text)
        Else
            If Text3.Text > "0" Then emess.Text = "2nd code date is not specified."
        End If
    End Sub

    Private Sub Text4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text4.Click
        tpad.Text = "Plate"
        tpad.cname.Text = "Text4"
        tpad.fname.Text = Me.Name
        tpad.trig.Text = Format(Val(Label8.Text), "000000") 'Val(tpad.trig) + 1
        tpad.Show()
    End Sub

    Private Sub Text4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text4.TextChanged
        If Len(Text4.Text) > 5 And Button1.Enabled = True Then Button1.Select()
    End Sub

    Private Sub xit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles xit.Click
        Me.Close()
    End Sub
End Class