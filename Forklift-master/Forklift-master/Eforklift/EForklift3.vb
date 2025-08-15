Public Class EForklift3

    Private Function check_staging(ByVal bc As String) As String                  'jv121714
        Dim ds As ADODB.Recordset, s As String, m As String
        s = "select target from paltasks where palletid = '" & bc & "' and area = 'DOCK'"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            m = bc & " had already been scanned to the DOCK for trailer: " & ds.Fields(0).Value
        Else
            m = " "
        End If
        ds.Close()
        check_staging = m
    End Function

    Private Sub complete_efl_shiptask()
        Dim p As ptask, k As Long
        k = Val(Mid(List3.SelectedItem, 1, 6))
        p = masterec(k)
        p.source = Combo3.SelectedItem
        If p.target = "ORDER PICK" Then
            p.target = "ORDER PICK"
            p.palletid = UCase(Text1.Text)
            p.qty = "1"
            p.uom = "Pallet"
            p.userid = WDUserId
            If remove_rack_pallet(p) = True Then
                p = masterec(p.id)
                p.target = "ORDER PICK"
            End If
            Call update_op_rack(p)
        End If

        If p.target = "STAGING" Then
            p.target = "STAGING"
            p.palletid = UCase(Text1.Text)
            p.qty = "1"
            p.uom = "Pallet"
            p.userid = WDUserId
            Call efl_to_dstg(p)
        End If

        Text1.Text = "" : emess.Text = ""
        If Combo2.Items.Count > 1 Then
            refresh_products()
        Else
            refresh_orders()
        End If
        Text1.Select()
    End Sub

    Private Sub barcode_scanned(ByVal bc As String)
        Dim s As String
        Dim p As ptask, k As Long
        Dim tmon As String, tday As String, tyr As String, topc As String, tpal As String, tdate As String

        s = " "
        If Combo2.Items.Count < 1 Then Exit Sub
        bc = UCase(bc)
        If Len(bc) <> 16 Then
            s = "Invalid barcode length: " & bc & "."
        Else
            tmon = Mid(bc, 5, 2)
            tday = Mid(bc, 7, 2)
            tyr = Mid(bc, 9, 2)
            topc = Mid(bc, 11, 3)                               'jv052515
            tpal = Mid(bc, 14, 3)
            tdate = tmon & "-" & tday & "-20" & tyr
            If Val(tmon) < 1 Or Val(tmon) > 12 Then
                's = "Invalid Month found in code date: " & Mid(bc, 1, 4) & " [" & tmon & "] " & Right(bc, 10) & "."
                s = "Invalid Month found in code date: " & Mid(bc, 1, 4) & " [" & tmon & "] " & Mid(bc, 7, 10) & "."
            End If
            If Val(tday) < 1 Or Val(tday) > 31 Then
                's = "Invalid Day found in code date: " & Mid(bc, 1, 6) & " [" & tday & "] " & Right(bc, 8) & "."
                s = "Invalid Day found in code date: " & Mid(bc, 1, 6) & " [" & tday & "] " & Mid(bc, 9, 8) & "."
            End If
            If Val(tyr) < 12 Then
                's = "Invalid Year found in code date: " & Mid(bc, 1, 8) & " [" & tyr & "] " & Right(bc, 6) & "."
                s = "Invalid Year found in code date: " & Mid(bc, 1, 8) & " [" & tyr & "] " & Mid(bc, 11, 6) & "."
            End If
            If IsDate(tdate) = False Then
                's = "Invalid code date (" & tdate & "): " & Mid(bc, 1, 4) & " [" & tmon & tday & tyr & "] " & Right(bc, 6) & "."
                s = "Invalid code date (" & tdate & "): " & Mid(bc, 1, 4) & " [" & tmon & tday & tyr & "] " & Mid(bc, 11, 6) & "."
            End If
            'If topc < "100" Or topc > "599" Then                    'jv052515
            '    s = "Invalid Operation Code found in barcode: " & Left(bc, 10) & "_" & topc & "_" & Right(bc, 3) & "."  'jv052515
            'End If                                                  'jv052515
            'If topc < "A" Or topc > "Z" Then
            '    s = "Invalid Operation Code found in barcode: " & Left(bc, 10) & " _" & topc & "_ " & Right(bc, 3) & "."
            'End If
            If Val(tpal) > 0 Or tpal = "EOR" Then
            Else
                s = "Invalid Pallet # found in barcode: " & Mid(bc, 1, 12) & " _" & tpal & "_."
            End If
            If Mid(Text1.Text, 1, 4) <> Mid(Combo2.SelectedItem, 1, 4) Then
                s = "Scanned Barcode: " & UCase(Text1.Text) & " does not match with: " & vbCrLf & Combo2.SelectedItem
            End If
        End If

        k = Val(Mid(List3.SelectedItem, 1, 6))                                                 'jv121714
        p = masterec(k)                                                         'jv121714
        If p.status <> "PEND" Then                                              'jv121714
            s = "Selected task is not in PEND status."                          'jv121714
            s = s & "  Try to re-scan the pallet."                              'jv121714
        End If                                                                  'jv121714
        If p.userid > " " Then                                                  'jv121714
            's = "Selected task is already assigned to user id: " & p.userid     'jv121714
            s = "Selected task is already assigned to user id: " & wdempname(p.userid)     'jv052515
            s = s & ".  Try to re-scan the pallet."                             'jv121714
        End If                                                                  'jv121714

        If s = " " Then s = check_staging(bc) 'jv121714

        If s = " " Then                                                         'jv052515
            p.palletid = bc                                                     'jv052515
            p.lotnum = barcode_to_lotnum(bc)
            If check_hold(p) = True Then                                        'jv052515
                s = bc & " is On-Hold.  Do not ship!"                           'jv052515
            End If                                                              'jv052515
        End If                                                                  'jv052515

        If s > " " Then
            emess.Text = s
            Button1.Enabled = False
            Text1.Text = ""                                                          'jv121714
            refresh_products()                                                    'jv121714
        Else
            'emess = List4 & vbCrLf & Combo2
            emess.Text = "__"
            Call draw_label(UCase(bc))
            Button1.Enabled = True
            'Button1.SetFocus()
            Button1.Select()
        End If
    End Sub

    Private Sub draw_label(ByVal bc As String)
        Dim i As Integer
        'i = Val(Left(bc, 3))
        'DoEvents()
        skupic.Text = Trim(Mid(bc, 1, 4))
        i = Val(skupic.Text)
        'lotpic = Mid(bc, 5, 8)
        lotpic.Text = Mid(bc, 5, 6)                  'jv052515
        oppic.Text = Mid(bc, 11, 3)                  'jv082415
        palnopic.Text = Mid(bc, 13, 3)
        If Val(palnopic.Text) > 0 Then palnopic.Text = Format(Val(palnopic.Text), "0")
        pkgpic.Text = labpix(i).package
        name1pic.Text = labpix(i).name1
        name2pic.Text = labpix(i).name2
        name3pic.Text = labpix(i).name3
        Frame1.Visible = True
        histbc = bc                             'jv052515
    End Sub

    Private Sub refresh_orders()
        Dim ds As ADODB.Recordset, s As String
        Combo1.Items.Clear() : List1.Items.Clear() : List2.Items.Clear()
        Combo2.Items.Clear() : List3.Items.Clear() : List4.Items.Clear()
        Combo3.Items.Clear()
        On Error GoTo vberror
        s = "select source, target, product from paltasks where area = 'GROUP'"
        s = s & " and status = 'ACTV'"
        s = s & " and product in (select description from paltasks where area = 'FORKLIFT' and status = 'PEND')"
        s = s & " order by product"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                Combo1.Items.Add(ds.Fields(2).Value)
                List1.Items.Add(ds.Fields(1).Value)
                List2.Items.Add(ds.Fields(2).Value)
                ds.MoveNext()
            Loop
            Combo1.SelectedIndex = 0
        End If
        ds.Close()
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.Description : Err.Clear()
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, "eforklift3", "refresh_orders", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: refresh_orders: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Sub

    Private Sub refresh_products()
        Dim ds As ADODB.Recordset, s As String
        Combo2.Items.Clear() : List3.Items.Clear() : List4.Items.Clear()
        On Error GoTo vberror
        s = "select id, product, units, source, target from paltasks where area = 'FORKLIFT'"
        s = s & " and description = '" & List2.SelectedItem
        s = s & " and userid <= ' '"
        s = s & " and status = 'PEND'"
        s = s & " order by product"
        ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst()
            Do Until ds.EOF
                Combo2.Items.Add(ds.Fields(1).Value)
                List3.Items.Add(Format(ds.Fields(0).Value, "000000") & " " & ds.Fields(1).Value & " " & Format(ds.Fields(2).Value, "0000"))
                List4.Items.Add(Trim(ds.Fields(3).Value) & " >> " & ds.Fields(4).Value)
                ds.MoveNext()
            Loop
            Combo2.SelectedIndex = 0
        End If
        ds.Close()
        Exit Sub
vberror:
        eno = Err.Number : edesc = Err.Description : Err.Clear()
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, "eforklift3", "refresh_products", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: refresh_products: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Sub

    Private Sub refresh_rack_locations(ByVal pqty As Integer)
        Dim ds As ADODB.Recordset, s As String, msku As String
        Combo3.Items.Clear()
        msku = Trim(Mid(Combo2.SelectedItem, 1, 4))
        If msku < "0" Then Exit Sub
        On Error GoTo vberror
        s = "select aisle, rack, fo from racks where hold = '0'"
        s = s & " and rack not in ('OP', 'SP')"
        s = s & " and id in (select rackno from rackpos where sku = '" & msku & "' and count_qty + qty2 = " & pqty
        s = s & " order by fo desc, aisle, rack"
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
        If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
            MsgBox("Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location...")
            Resume
        Else
            Call vb_elog(eno, edesc, "eforklift3", "refresh_rack_locations", WDUserId)
            If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: refresh_rack_locations: " & eno) = vbRetry Then
                Resume
            Else
                End
            End If
        End If
    End Sub

    Private Sub xit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles xit.Click
        Me.Close()
    End Sub

    Private Sub EForklift3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Text1.Text = ""
        emess.Text = ""
        refresh_orders()
    End Sub

    Private Sub Combo1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo1.SelectedIndexChanged
        List1.SelectedIndex = Combo1.SelectedIndex
        List2.SelectedIndex = Combo1.SelectedIndex
        Text1.Text = ""
    End Sub

    Private Sub Combo2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Combo2.SelectedIndexChanged
        List3.SelectedIndex = Combo2.SelectedIndex
        List4.SelectedIndex = Combo2.SelectedIndex
        Text1.Text = ""
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        complete_efl_shiptask()        
    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim s As String
        If Frame1.Visible = True Then
            s = "Active Pallet BarCode!"
            MsgBox(s, vbOKOnly + vbExclamation, "active task has not been completed...")
        Else
            refresh_orders()
        End If
    End Sub

    Private Sub Frame1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Frame1.Click
        If Button1.Enabled Then complete_efl_shiptask()
    End Sub

    Private Sub List1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles List1.SelectedIndexChanged
        Label2.Text = List1.SelectedItem
    End Sub

    Private Sub List2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles List2.SelectedIndexChanged
        refresh_products()
    End Sub

    Private Sub List3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles List3.SelectedIndexChanged
        Dim k As Integer
        k = Val(Mid(List3.SelectedItem, Len(List3.SelectedItem) - 4, 4))
        Call refresh_rack_locations(k)
    End Sub

    Private Sub List4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles List4.SelectedIndexChanged
        Label6.Text = List4.SelectedItem
    End Sub

    Private Sub lotpic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lotpic.Click
        If Button1.Enabled Then complete_efl_shiptask()
    End Sub

    Private Sub name1pic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles name1pic.Click
        If Button1.Enabled Then complete_efl_shiptask()
    End Sub

    Private Sub name2pic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles name2pic.Click
        If Button1.Enabled Then complete_efl_shiptask()
    End Sub

    Private Sub name3pic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles name3pic.Click
        If Button1.Enabled Then complete_efl_shiptask()
    End Sub

    Private Sub oppic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles oppic.Click
        If Button1.Enabled Then complete_efl_shiptask()
    End Sub

    Private Sub palnopic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles palnopic.Click
        If Button1.Enabled Then complete_efl_shiptask()
    End Sub

    Private Sub pkgpic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pkgpic.Click
        If Button1.Enabled Then complete_efl_shiptask()
    End Sub

    Private Sub skupic_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles skupic.Click
        If Button1.Enabled Then complete_efl_shiptask()
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
        Text1.BackColor = Combo1.BackColor
    End Sub

    Private Sub Text1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Text1.LostFocus
        Text1.BackColor = Me.BackColor
    End Sub

    Private Sub Text1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Text1.TextChanged
        Frame1.Visible = False
        If Len(Text1.Text) > 15 Then
            Call barcode_scanned(Text1.Text)
        Else
            Button1.Enabled = False
        End If
    End Sub

    Private Sub EForklift3_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Label2.Left = (Me.Width - Label2.Width) * 0.5
        Frame1.Left = (Me.Width - Frame1.Width) * 0.5
        Button1.Left = (Me.Width - Button1.Width) * 0.5
        apphdr.Left = Label2.Left
        Label1.Left = Label2.Left : Combo1.Left = Label1.Left + Label1.Width

        Label3.Left = Label2.Left : Combo2.Left = Label3.Left + Label3.Width
        Label4.Left = Label2.Left : Combo3.Left = Label4.Left + Label4.Width
        Label5.Left = Label2.Left : Text1.Left = Label5.Left + Label5.Width
        Label6.Left = Label2.Left
        emess.Left = Label2.Left

        'Combo2.Top = Me.Height * 0.3
        'Label3.Top = Combo2.Top
        'Label2.Top = Combo2.Top - (Combo2.Height * 2)
        'Label1.Top = Label2.Top - (Label2.Height * 2)
        'Combo1.Top = Label1.Top
        'apphdr.Top = Label1.Top - (Label1.Height * 2)
        'Label6.Top = Combo2.Top + (Combo2.Height * 2)
        'Label4.Top = Label6.Top + (Label6.Height * 2)
        'Combo3.Top = Label4.Top
        'Label5.Top = Label4.Top + (Label4.Height * 2)
        'Text1.Top = Label5.Top
        'Button1.Top = Text1.Top + (Text1.Height * 2)
        'emess.Top = Button1.Top + (Button1.Height * 2)
        Frame1.Top = emess.Top
        xit.Left = Me.Width - xit.Width
    End Sub
End Class