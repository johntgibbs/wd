Attribute VB_Name = "quotautil"
Public wdb As ADODB.Connection
Public r12db As ADODB.Connection
Public tsb As ADODB.Connection
Public a10shipdb As ADODB.Connection
Public k10shipdb As ADODB.Connection
Public t10bbsr As String
Public k10bbsr As String
Public a10bbsr As String
Public k10ship As String
Public a10ship As String
Public cs5db As String
Public r12access As Boolean
Public r12connection As String
Public bimpuserid As String
Public bimpstat As String                           'jv022316
Public salesdays As Integer                         'jv050117
Public brzloss As String

Type skuinfo
    sku As String
    unit As String
    desc As String
    psrc As Integer
    pallet As Integer
    whs As Integer
    wrapunits As Integer
End Type

Type branchinfo
    branchno As String
    branchname As String
    oraloc As String
    capacity As Long
    usable As Long
    supplier As String
    region As String
End Type

Type plantinfo
    plantno As String
    plantname As String
    orawhs As String
End Type

Type bimprec
    id As Long
    plantwhs As String
    branchwhs As String
    sku As String
    onhand As Long
    onorder As Long
    sales As Long
    undiff As Long
    paldiff As Integer
    ohpct As Single
    roqty As Integer
    pctgain As Single
    needqty As Integer
    bimpstatus As String
    quotapct As Single
    plantpool As Long
    poolsched As Long
    thiswknewpals As Integer
    nextwknewpals As Integer
    lowqty As Integer
    outqty As Integer
    lowflag As String
    outflag As String
    lastrcpt As String
End Type

Type loadloss
    id As Long
    sku As String
    palsize As Integer
    plantwhs As String
    branchwhs As String
    lastrcpt As String
    lastissue As String
    sales As Long
End Type

Type stkhist
    id As Long
    branchwhs As String
    sku As String
    startdate As String
    enddate As String
    postdate As String
    totaldays As Integer
    daysin As Integer
    daysout As Integer
    loads As Long
End Type

Global skurec(0 To 9999) As skuinfo
Global branchrec(1 To 99) As branchinfo
Global plantrec(50 To 52) As plantinfo
Global eno As Long
Global edesc As String

Function avg_sales(bwhs As String, bsku As String, sdaze As Integer) As Long        'jv051618
    Dim ds As ADODB.Recordset, s As String, i As Long
    i = 0
    s = "select daysin, loads from stockhistory"
    s = s & " Where branchwhs = '" & bwhs & "'"
    s = s & " and sku = '" & bsku & "'"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If ds!daysin > 0 And ds!loads > 0 Then
            i = (ds!loads / ds!daysin) * sdaze
        End If
    End If
    avg_sales = i
End Function

Sub build_plants()
    plantrec(50).plantno = "50"
    plantrec(51).plantno = "51"
    plantrec(52).plantno = "52"
    plantrec(50).plantname = "Brenham"
    plantrec(51).plantname = "Broken Arrow"
    plantrec(52).plantname = "Sylacauga"
    plantrec(50).orawhs = "T10"
    plantrec(51).orawhs = "K10"
    plantrec(52).orawhs = "A10"
End Sub

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

Sub bimp_missing_skus()
    Dim ds As ADODB.Recordset, bs As ADODB.Recordset, s As String
    Dim rt As String, rf As String, rh As String
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    Screen.MousePointer = 11
    bimpbanner.pgrid.Clear: bimpbanner.pgrid.Rows = 1: bimpbanner.pgrid.Cols = 5
    'R12 Onhand
    s = "select o.subinventory_code, m.segment1, sum(o.transaction_quantity)"
    s = s & " from mtl_onhand_quantities o, mtl_system_items_b m"
    s = s & ", mtl_item_locations l"
    's = s & " where o.subinventory_code > '001'"
    s = s & " where o.subinventory_code > '000'"                        'jv120715
    s = s & " and o.subinventory_code not in ('A01', 'A10', 'K01', 'K10', 'T01', 'T10')"
    s = s & " and m.organization_id = o.organization_id"
    s = s & " and m.inventory_item_id = o.inventory_item_id"
    s = s & " and m.segment1 >= '100' and m.segment1 <= '9999'"
    's = s & " and m.segment1 = '777'"
    s = s & " and l.inventory_location_id = o.locator_id"
    s = s & " and l.segment1 > 'FLOOR   '"
    s = s & " and l.segment1 < 'FLOORZZZ'"
    s = s & " group by o.subinventory_code, m.segment1"
    s = s & " order by m.segment1"
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds(0) <> "A01" And ds(0) <> "K01" And ds(0) <> "T01" Then
                If ds(0) = "T10" Or ds(0) = "A10" Or ds(0) = "K10" Then
                    s = "Select sku from bimp"
                    s = s & " Where plantwhs = '" & ds(0) & "'"
                    s = s & " And sku = '" & ds(1) & "'"
                Else
                    s = "Select sku from bimp"
                    s = s & " Where branchwhs = '" & ds(0) & "'"
                    s = s & " And sku = '" & ds(1) & "'"
                End If
                Set bs = wdb.Execute(s)
                If bs.BOF = True Then
                    s = ds(0) & Chr(9)
                    If ds(0) = "T10" Then
                        s = s & "Brenham Plant" & Chr(9)
                    Else
                        If ds(0) = "K10" Then
                            s = s & "Broken Arrow" & Chr(9)
                        Else
                            If ds(0) = "A10" Then
                                s = s & "Sylacauga Plant" & Chr(9)
                            Else
                                If Val(ds(0)) > 0 And Val(ds(0)) < 1000 Then
                                    s = s & branchrec(Val(ds(0))).branchname & Chr(9)
                                Else
                                    s = s & "Warehouse Unknown" & Chr(9)
                                End If
                            End If
                        End If
                    End If
                    s = s & ds(1) & Chr(9)
                    s = s & skurec(Val(ds(1))).unit & " " & skurec(Val(ds(1))).desc & Chr(9)
                    s = s & ds(2)
                    bimpbanner.pgrid.AddItem s
                End If
                bs.Close
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    bimpbanner.pgrid.FormatString = "^Whs|<Branch Name|^SKU|<Product|^Units"
    bimpbanner.pgrid.ColWidth(0) = 1000
    bimpbanner.pgrid.ColWidth(1) = 3000
    bimpbanner.pgrid.ColWidth(2) = 1000
    bimpbanner.pgrid.ColWidth(3) = 3000
    bimpbanner.pgrid.ColWidth(4) = 1000
    Screen.MousePointer = 0
    rt = "Missing BIMP SKUs"
    rh = " "
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    'htdc(0) = "cyan": gndc(0) = Me.Grid1.BackColorFixed
    'htdc(1) = "yellow": gndc(1) = Me.Grid1.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid1.BackColor
    bimpbanner.pgrid.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(bimpbanner, "c:\htmlgrid.htm", bimpbanner.pgrid, rt, rh, rf, "linen", "khaki", "white")
        bimpbanner.pgrid.Redraw = True
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(bimpbanner, "c:\htmlgrid.htm", bimpbanner.pgrid, rt, rh, rf, "linen", "khaki", "white")
        bimpbanner.pgrid.Redraw = True
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Function bimp_sales_days()                              'jv050117
    Dim ds As ADODB.Recordset, s As String, i As Integer
    i = 30
    s = "select listreturn from valuelists where listname = 'bimploaddays'"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        i = ds!listreturn
    End If
    ds.Close
    bimp_sales_days = i
End Function

Function bimp_status_time() As String                   'jv022316
    Dim s As String
    s = " "
    If Len(Dir(bimpstat)) > 0 Then
        Open bimpstat For Input As #1
        Line Input #1, s
        Close #1
    End If
    bimp_status_time = s
End Function

Sub build_skumast()
    Dim ds As ADODB.Recordset, sqlx As String, i As Integer
    'On Error GoTo vberror
    sqlx = "select * from skumast order by sku"
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = Val(ds!sku)
            If Len(ds!sku) > 0 Then skurec(i).sku = ds!sku                      'jv082415
            If Len(ds!fgunit) > 0 Then skurec(i).unit = Trim(ds!fgunit & " ")   'jv082415
            If Len(ds!fgdesc) > 0 Then skurec(i).desc = Trim(ds!fgdesc & " ")   'jv082415
            If Len(ds!psource) > 0 Then skurec(i).psrc = ds!psource             'jv082415
            If Len(ds!pallet) > 0 Then skurec(i).pallet = ds!pallet             'jv082415
            If Len(ds!whs_num) > 0 Then skurec(i).whs = ds!whs_num              'jv082415
            If Len(ds!numwrap) > 0 Then skurec(i).wrapunits = ds!numwrap
            ds.MoveNext
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.Description: Err.Clear
    Call vb_elog(eno, edesc, "Sub", "build_skumast", "quota")
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, "Sub: build_skumast - Error Number: " & eno
        End
    End If
End Sub

Sub build_branches()
    Dim ds As ADODB.Recordset, sqlx As String, i As Integer
    'On Error GoTo vberror
    sqlx = "select * from branches where gemmsid > ' ' order by branch"
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = ds!branch
            branchrec(i).branchno = ds!branch
            branchrec(i).branchname = ds!branchname
            branchrec(i).oraloc = ds!gemmsid
            branchrec(i).capacity = Val(ds!modem)
            branchrec(i).usable = Val(ds!fax)
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "select * from valuelists where listname = 'branchplants'"
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = Val(ds!listreturn)
            branchrec(i).supplier = ds!listdisplay
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "select * from valuelists where listname = 'brzdivmap'"
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = Val(Left(ds!listreturn, 2))
            branchrec(i).region = ds!listdisplay
            ds.MoveNext
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.Description: Err.Clear
    Call vb_elog(eno, edesc, "Sub", "build_branches", "quota")
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, "Sub: build_branches - Error Number: " & eno
        End
    End If
End Sub

Sub calc_bimp_all()
    Dim ds As ADODB.Recordset, s As String
    Dim ponhand As Long, ponorder As Long, psales As Long, pundiff As Long, ppaldiff As Long
    Dim pohpct As Currency, ppctgain As Currency, pneed As Long, pstat As String
    Dim plowflag As String, poutflag As String, ptestqty As Long
    salesdays = bimp_sales_days             'jv053018
    'MsgBox "calc_bimp_all"
    Screen.MousePointer = 11
    's = "select id, sku, onhand, onorder, sales, lowqty, outqty, plantpool, branchwhs from bimp"
    s = "select id, sku, onhand, onorder, sales, lowqty, outqty, plantpool, branchwhs,"         'jv072216
    s = s & " thiswknewpals, nextwknewpals, roqty from bimp"                                    'jv072216
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            ponhand = ds!onhand
            ponorder = ds!onorder
            'ponorder = ponorder + (ds!thiswknewpals * ds!roqty)                 'jv072216
            'ponorder = ponorder + (ds!nextwknewpals * ds!roqty)                 'jv072216
            psales = ds!sales
            pundiff = (ponhand + ponorder) - psales
            If skurec(Val(ds!sku)).pallet > 0 Then
                ppaldiff = Format(pundiff / skurec(Val(ds!sku)).pallet, "0")
            Else
                ppaldiff = 0
            End If
            If psales <> 0 Then
                pohpct = Format((ponhand + ponorder) / psales, "0.000")
                If pohpct > 9999 Then pohpct = 9999                            'jv113015
                If skurec(Val(ds!sku)).pallet > 0 Then
                    ppctgain = Format(skurec(Val(ds!sku)).pallet / psales, "0.000")
                Else
                    ppctgain = 0
                End If
                pneed = 0
                If pohpct < 0.5 And ppctgain < 1 And ppctgain <> 0 Then
                    'MsgBox ds!sku & " " & ds!branchwhs
                    pneed = Val(Format((0.5 - pohpct) / ppctgain, "0"))
                End If
            Else
                pohpct = 0
                ppctgain = 0
                pneed = 0
            End If
            pstat = calc_bimp_status(salesdays, pohpct, ppaldiff)           'jv053018
            'If pneed > 0 Then
            '    pstat = "W"
            '
            'Else
            '    If ppaldiff = 0 Then
            '        pstat = "B"
            '    Else
            '        If ppaldiff > 0 Then
            '            pstat = "G"
            '        Else
            '            pstat = "Y"
            '        End If
            '    End If
            'End If
            ptestqty = ds!lowqty * skurec(Val(ds!sku)).pallet
            If ds!plantpool > ptestqty Then
                plowflag = "N"
            Else
                plowflag = "Y"
            End If
            ptestqty = ds!outqty * skurec(Val(ds!sku)).pallet
            If ds!plantpool > ptestqty Then
                poutflag = "N"
            Else
                poutflag = "Y"
            End If
            s = "Update bimp set sales = " & psales
            's = s & ", onorder = " & ponorder                       'jv072216
            s = s & ", undiff = " & pundiff
            s = s & ", paldiff = " & ppaldiff
            s = s & ", ohpct = " & pohpct
            s = s & ", roqty = " & skurec(Val(ds!sku)).pallet
            s = s & ", pctgain = " & ppctgain
            s = s & ", needqty = " & pneed
            s = s & ", bimpstatus = '" & pstat & "'"
            s = s & ", lowflag = '" & plowflag & "'"
            s = s & ", outflag = '" & poutflag & "'"
            s = s & " Where id = " & ds!id
            'MsgBox s, vbOKOnly, "calc_bimp..."
            wdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
End Sub

Function calc_bimp_line(bline As bimprec) As bimprec
    Dim ds As ADODB.Recordset, s As String
    Dim ponhand As Long, ponorder As Long, psales As Long, pundiff As Long, ppaldiff As Long
    Dim pohpct As Currency, ppctgain As Currency, pneed As Long, pstat As String
    Dim plowflag As String, poutflag As String, ptestqty As Long
    salesdays = bimp_sales_days             'jv053018
    'MsgBox "calc_bimp_line"
    ponhand = bline.onhand
    ponorder = bline.onorder
    ponorder = ponorder + (bline.thiswknewpals * bline.roqty)
    ponorder = ponorder + (bline.nextwknewpals * bline.roqty)
    psales = bline.sales
    pundiff = (ponhand + ponorder) - psales
    
    If skurec(Val(bline.sku)).pallet > 0 Then
        ppaldiff = Format(pundiff / skurec(Val(bline.sku)).pallet, "0")
    Else
        ppaldiff = 0
    End If
    
    If psales <> 0 Then
        pohpct = Format((ponhand + ponorder) / psales, "0.000")
        If pohpct > 9999 Then pohpct = 9999                            'jv113015
        If skurec(Val(bline.sku)).pallet > 0 Then
            ppctgain = Format(skurec(Val(bline.sku)).pallet / psales, "0.000")
        Else
            ppctgain = 0
        End If
        pneed = 0
        If pohpct < 0.5 And ppctgain < 1 And ppctgain <> 0 Then
            'MsgBox ds!sku & " " & ds!branchwhs
            pneed = Val(Format((0.5 - pohpct) / ppctgain, "0"))
        End If
    Else
        pohpct = 0
        ppctgain = 0
        pneed = 0
    End If
    
    pstat = calc_bimp_status(salesdays, pohpct, ppaldiff)           'jv053018
    'If pneed > 0 Then
    '    pstat = "W"
    'Else
    '    If ppaldiff = 0 Then
    '        pstat = "B"
    '    Else
    '        If ppaldiff > 0 Then
    '            pstat = "G"
    '        Else
    '            pstat = "Y"
    '        End If
    '    End If
    'End If
    
    ptestqty = bline.lowqty * skurec(Val(bline.sku)).pallet
    If bline.plantpool > ptestqty Then
        plowflag = "N"
    Else
        plowflag = "Y"
    End If
    ptestqty = bline.outqty * skurec(Val(bline.sku)).pallet
    If bline.plantpool > ptestqty Then
        poutflag = "N"
    Else
        poutflag = "Y"
    End If
    
    bline.sales = psales
    bline.onorder = ponorder
    bline.undiff = pundiff
    bline.paldiff = ppaldiff
    bline.ohpct = pohpct
    bline.roqty = skurec(Val(bline.sku)).pallet
    bline.pctgain = ppctgain
    bline.needqty = pneed
    bline.bimpstatus = pstat
    bline.lowflag = plowflag
    bline.outflag = poutflag
    calc_bimp_line = bline
End Function

Function calc_bimp_status(sdays As Integer, ohpct As Currency, paldiff As Long) As String      'jv053018
    Dim ohdays As Long, pstat As String
    ohdays = sdays * ohpct
    If ohdays > 0 Then
        If ohdays < 14 Then
            pstat = "W"
        Else
            If ohdays < 30 Then
                pstat = "Y"
            Else
                If paldiff > 0 Then
                    pstat = "G"
                Else
                    pstat = "B"
                End If
            End If
        End If
    Else
        pstat = "B"
    End If
    'MsgBox ohdays & " " & pstat
    calc_bimp_status = pstat
End Function

Public Sub connect_r12()
    Dim s As String
    s = "This event requires an ODBC connection to the Oracle R12 Database."
    s = s & "  Do you wish to try to connect?"
    If MsgBox(s, vbYesNo + vbQuestion, "Connect R12....") = vbNo Then Exit Sub
    On Error GoTo r12err
    Set r12db = CreateObject("ADODB.Connection")
    r12db.Open r12connection
    r12access = True
    Exit Sub
r12err:
    MsgBox "R12 Connection failed.", vbOKOnly + vbInformation, "Sorry, no connection...."
End Sub

Public Sub import_r12_sales()
    Dim ds As ADODB.Recordset, os As ADODB.Recordset, s As String
    Dim ponhand As Long, ponorder As Long, psales As Long, pundiff As Long, ppaldiff As Long
    'Dim phpct As Single, ppctgain As Single, pneed As Long
    Dim pohpct As Currency, ppctgain As Currency, pneed As Long
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    'MsgBox "import_r12_sales"
    Screen.MousePointer = 11
    
    salesdays = bimp_sales_days                                                 'jv050117
    'Clear Sales Qtys
    s = "Update bimp set sales = 0, undiff = 0, paldiff = 0, ohpct = 0, pctgain = 0, needqty = 0"
    wdb.Execute s
    
    'R12 Onhand
    s = "select product_no,branch_no,sum(tran_qty) from bolinf.inv_adj_input_detail"
    s = s & " where tran_type = '1'"
    's = s & " and trunc(tran_date) > trunc(SYSDATE - 30)"
    's = s & " and trunc(tran_date) > trunc(SYSDATE - 31)"                           'jv031816
    s = s & " and trunc(tran_date) > trunc(SYSDATE - " & Format(salesdays + 1, "0") & ")"    'jv050117
    s = s & " group by product_no,branch_no"
    s = s & " order by product_no,branch_no"
    'MsgBox s
    Set os = r12db.Execute(s)
    If os.BOF = False Then
        os.MoveFirst
        Do Until os.EOF
            s = "select * from bimp where branchwhs = '" & os!branch_no & "'"
            s = s & " and sku = '" & os!product_no & "'"
            Set ds = wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                Do Until ds.EOF
                    ponhand = ds!onhand
                    ponorder = ds!onorder
                    'psales = os(2)
                    psales = (os(2) * 30) / salesdays                           'jv050117
                    pundiff = (ponhand + ponorder) - psales
                    If skurec(Val(ds!sku)).pallet <> 0 Then
                        ppaldiff = Format(pundiff / skurec(Val(ds!sku)).pallet, "0")
                    Else
                        ppaldiff = 0
                    End If
                    If psales <> 0 Then
                        pohpct = Format((ponhand + ponorder) / psales, "0.000")
                        'If pohpct > 16000 Then pohpct = 1000                                    'jv113015
                        If pohpct > 9999 Then pohpct = 9999                                    'jv113015
                        If skurec(Val(ds!sku)).pallet <> 0 Then
                            ppctgain = Format(skurec(Val(ds!sku)).pallet / psales, "0.000")
                        Else
                            ppctgain = 1
                        End If
                    Else
                        pohpct = 1
                        ppctgain = 1
                    End If
                    pneed = 0
                    If pohpct < 0.5 And ppctgain < 1 And ppctgain <> 0 Then
                        pneed = Val(Format((0.5 - pohpct) / ppctgain, "0"))
                    End If
                    s = "Update bimp set sales = " & psales
                    s = s & ", undiff = " & pundiff
                    s = s & ", paldiff = " & ppaldiff
                    s = s & ", ohpct = " & pohpct
                    s = s & ", roqty = " & skurec(Val(ds!sku)).pallet
                    s = s & ", pctgain = " & ppctgain
                    s = s & ", needqty = " & pneed
                    s = s & ", bimpstatus = '" & calc_bimp_status(salesdays, pohpct, ppaldiff) & "'"    'jv053018
                    'If pneed > 0 Then
                    '    s = s & ", bimpstatus = 'W'"
                    'Else
                    '    If ppaldiff = 0 Then
                    '        s = s & ", bimpstatus = 'B'"
                    '    Else
                    '        If ppaldiff > 0 Then
                    '            s = s & ", bimpstatus = 'G'"
                    '        Else
                    '            s = s & ", bimpstatus = 'Y'"
                    '        End If
                    '    End If
                    'End If
                    s = s & " Where id = " & ds!id
                    'MsgBox s
                    wdb.Execute s
                    ds.MoveNext
                Loop
            End If
            ds.Close
            os.MoveNext
        Loop
    End If
    os.Close
    Screen.MousePointer = 0
End Sub

Public Sub import_r12_branch_sales(pwhs As String)
    Dim ds As ADODB.Recordset, os As ADODB.Recordset, s As String
    Dim ponhand As Long, ponorder As Long, psales As Long, pundiff As Long, ppaldiff As Long
    Dim pohpct As Currency, ppctgain As Currency, pneed As Long
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    'MsgBox "import_r12_branch_sales"
    Screen.MousePointer = 11
    
    salesdays = bimp_sales_days                                                     'jv050117
    'Clear Sales Qtys
    s = "Update bimp set sales = 0, undiff = 0, paldiff = 0, ohpct = 0, pctgain = 0, needqty = 0"
    s = s & " where branchwhs = '" & pwhs & "'"
    wdb.Execute s
    
    'R12 Onhand
    s = "select product_no,branch_no,sum(tran_qty) from bolinf.inv_adj_input_detail"
    s = s & " where tran_type = '1'"
    s = s & " and branch_no = '" & pwhs & "'"
    's = s & " and trunc(tran_date) > trunc(SYSDATE - 30)"
    's = s & " and trunc(tran_date) > trunc(SYSDATE - 31)"                           'jv031816
    s = s & " and trunc(tran_date) > trunc(SYSDATE - " & Format(salesdays + 1, "0") & ")"    'jv050117
    s = s & " group by product_no,branch_no"
    s = s & " order by product_no,branch_no"
    'MsgBox s
    Set os = r12db.Execute(s)
    If os.BOF = False Then
        os.MoveFirst
        Do Until os.EOF
            s = "select * from bimp where branchwhs = '" & os!branch_no & "'"
            s = s & " and sku = '" & os!product_no & "'"
            Set ds = wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                Do Until ds.EOF
                    ponhand = ds!onhand
                    ponorder = ds!onorder
                    psales = os(2)
                    psales = (os(2) * 30) / salesdays                                       'jv05117
                    pundiff = (ponhand + ponorder) - psales
                    If skurec(Val(ds!sku)).pallet <> 0 Then
                        ppaldiff = Format(pundiff / skurec(Val(ds!sku)).pallet, "0")
                    Else
                        ppaldiff = 0
                    End If
                    If psales <> 0 Then
                        pohpct = Format((ponhand + ponorder) / psales, "0.000")
                        'If pohpct > 16000 Then pohpct = 1000                                    'jv113015
                        If pohpct > 9999 Then pohpct = 9999                                    'jv113015
                        If skurec(Val(ds!sku)).pallet <> 0 Then
                            ppctgain = Format(skurec(Val(ds!sku)).pallet / psales, "0.000")
                        Else
                            ppctgain = 1
                        End If
                    Else
                        pohpct = 1
                        ppctgain = 1
                    End If
                    pneed = 0
                    If pohpct < 0.5 And ppctgain < 1 And ppctgain <> 0 Then
                        pneed = Val(Format((0.5 - pohpct) / ppctgain, "0"))
                    End If
                    s = "Update bimp set sales = " & psales
                    s = s & ", undiff = " & pundiff
                    s = s & ", paldiff = " & ppaldiff
                    s = s & ", ohpct = " & pohpct
                    s = s & ", roqty = " & skurec(Val(ds!sku)).pallet
                    s = s & ", pctgain = " & ppctgain
                    s = s & ", needqty = " & pneed
                    s = s & ", bimpstatus = '" & calc_bimp_status(salesdays, pohpct, ppaldiff) & "'"    'jv053018
                    'If pneed > 0 Then
                    '    s = s & ", bimpstatus = 'W'"
                    'Else
                    '    If ppaldiff = 0 Then
                    '        s = s & ", bimpstatus = 'B'"
                    '    Else
                    '        If ppaldiff > 0 Then
                    '            s = s & ", bimpstatus = 'G'"
                    '        Else
                    '            s = s & ", bimpstatus = 'Y'"
                    '        End If
                    '    End If
                    'End If
                    s = s & " Where id = " & ds!id
                    'MsgBox s
                    wdb.Execute s
                    ds.MoveNext
                Loop
            End If
            ds.Close
            os.MoveNext
        Loop
    End If
    os.Close
    Screen.MousePointer = 0
End Sub

Public Sub import_r12_route_loads()
    Dim ds As ADODB.Recordset
    Dim q As String, lbr As String, ldate As String
    Dim cfile As String
    
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    
    lbr = InputBox("Warehouse:", "Oracle Warehouse..", "001")
    If Len(lbr) = 0 Then Exit Sub
    ldate = InputBox("Load Date:", "Load Date..", Format(Now, "m-d-yyyy"))
    If Len(ldate) = 0 Then Exit Sub
    If IsDate(ldate) = False Then
        MsgBox "Invalid Date: " & ldate, vbOKOnly + vbInformation, "try again..."
        Exit Sub
    End If
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\routes." & lbr
    cfile = InputBox("Route Countsheet File:", "Export File..", cfile)
    If Len(cfile) = 0 Then Exit Sub
    Screen.MousePointer = 11
    Open cfile For Append As #1
    q = "select product_no,route_no,sum(tran_qty)"
    q = q & " from bolinf.inv_adj_input_detail"
    q = q & " where tran_type = '1'"
    q = q & " and tran_date = TO_DATE('" & Format(ldate, "dd-mmm-yy") & "')"
    q = q & " and branch_no = '" & lbr & "'"
    q = q & " group by product_no, route_no"
    q = q & " order by product_no, route_no"
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds(2) <> 0 Then
                Write #1, "RT" & ds(1);
                Write #1, ds(0);
                i = Val(ds(0))
                Write #1, skurec(i).unit & " " & skurec(i).desc;
                Write #1, ""; "";
                Write #1, ds(2) * -1;
                Write #1, ds(2) * -1;
                Write #1, Format(ldate, "m-d-yyyy")
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Close #1
    Screen.MousePointer = 0
End Sub

Public Sub import_r12_qtys()
    Dim ds As ADODB.Recordset, s As String
    Dim pb As ADODB.Connection, servup As Boolean
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    Screen.MousePointer = 11
    
    'Clear Quota Qtys
    s = "Update bimp set onhand = 0, onorder = 0, plantpool = 0"
    wdb.Execute s
    
    'R12 Onhand
    s = "select o.subinventory_code, m.segment1, sum(o.transaction_quantity)"
    s = s & " from mtl_onhand_quantities o, mtl_system_items_b m"
    s = s & ", mtl_item_locations l"
    's = s & " where o.subinventory_code > '001'"
    s = s & " where o.subinventory_code > '000'"                        'jv120715
    s = s & " and o.subinventory_code not in ('A01', 'K01', 'T01')"     'jv022717
    s = s & " and m.organization_id = o.organization_id"
    s = s & " and m.inventory_item_id = o.inventory_item_id"
    's = s & " and m.segment1 >= '100' and m.segment1 <= '967'"
    's = s & " and m.segment1 = '777'"
    s = s & " and l.inventory_location_id = o.locator_id"
    s = s & " and l.segment1 > 'FLOOR   '"
    s = s & " and l.segment1 < 'FLOORZZZ'"
    s = s & " group by o.subinventory_code, m.segment1"
    s = s & " order by m.segment1"
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            'If ds(0) = "T10" Or ds(0) = "001" Then
            '    s = "Update bimp set onhand = onhand + " & ds(2)
            '    s = s & " Where branchwhs = '001'"
            '    s = s & " And sku = '" & ds(1) & "'"
            'Else
                If ds(0) = "T10" Or ds(0) = "A10" Or ds(0) = "K10" Then
                'If ds(0) = "A10" Or ds(0) = "K10" Then
                    s = "Update bimp set plantpool = " & ds(2)
                    s = s & " Where plantwhs = '" & ds(0) & "'"
                    s = s & " And sku = '" & ds(1) & "'"
                Else
                    s = "Update bimp set onhand = " & ds(2)
                    s = s & " Where branchwhs = '" & ds(0) & "'"
                    s = s & " And sku = '" & ds(1) & "'"
                End If
            'End If
            'MsgBox s
            wdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    'WMS Trailers - Brenham
    's = "select runid, plant, branch, sku, sum(units) from trailers where pb_flag = 'N'"
    's = s & " and shipdate > '" & Format(Now, "M-d-yyyy") & "'"
    s = "select runid, plant, branch, sku, sum(units) from trailers"        'jv102615
    s = s & " where shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv102615
    s = s & " and pallets > 0"                  'jv040318
    s = s & " and plant = 50"
    s = s & " Group by runid, plant, branch, sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ticket_post(ds!runid) = False Then
                'If ds!plant = "50" Then
                    s = "Update bimp set plantpool = plantpool - " & ds(4)
                    s = s & " Where plantwhs = 'T10' and sku = '" & ds!sku & "'"
                'End If
                'If ds!plant = "51" Then
                '    s = "Update bimp set plantpool = plantpool - " & ds(4)
                '    s = s & " Where plantwhs = 'K10' and sku = '" & ds!sku & "'"
                'End If
                'If ds!plant = "52" Then
                '    s = "Update bimp set plantpool = plantpool - " & ds(4)
                '    s = s & " Where plantwhs = 'A10' and sku = '" & ds!sku & "'"
                'End If
                'MsgBox s
                wdb.Execute s
            End If
            
            If ticket_receipt(ds!runid) = False Then
                s = "Update bimp set onorder = onorder + " & ds(4)
                s = s & " where branchwhs = '" & Format(ds!branch, "000") & "'"
                If ds!plant = "50" Then s = s & " and plantwhs = 'T10'"
                'If ds!plant = "51" Then s = s & " and plantwhs = 'K10'"
                'If ds!plant = "52" Then s = s & " and plantwhs = 'A10'"
                s = s & " and sku = '" & ds!sku & "'"
                'MsgBox s
                wdb.Execute s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    'WMS Trailers - Broken Arrow
    If plant_server_status("K10") = True Then                       'jv010417
        Set pb = CreateObject("ADODB.Connection")
        pb.Open k10ship
        's = "select runid, plant, branch, sku, sum(units) from trailers where pb_flag = 'N'"
        's = s & " and shipdate > '" & Format(Now, "M-d-yyyy") & "'"
        s = "select runid, plant, branch, sku, sum(units) from trailers"        'jv102615
        s = s & " where shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv102615
        s = s & " and pallets > 0"                  'jv040318
        s = s & " and plant = 51"
        s = s & " Group by runid, plant, branch, sku"
        Set ds = pb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ticket_post(ds!runid) = False Then
                    s = "Update bimp set plantpool = plantpool - " & ds(4)
                    s = s & " Where plantwhs = 'K10' and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
            
                If ticket_receipt(ds!runid) = False Then
                    s = "Update bimp set onorder = onorder + " & ds(4)
                    s = s & " where branchwhs = '" & Format(ds!branch, "000") & "'"
                    s = s & " and plantwhs = 'K10'"
                    s = s & " and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close: pb.Close
    Else                                                                'jv010417
        MsgBox "Broken Arrow is offline.", vbOKOnly + vbInformation, "BA Trailers not processed..."
        s = "select runid, plant, branch, sku, sum(units) from trailers"        'jv102615
        s = s & " where shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv102615
        s = s & " and plant = 51"
        s = s & " Group by runid, plant, branch, sku"
        Set ds = wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ticket_post(ds!runid) = False Then
                    s = "Update bimp set plantpool = plantpool - " & ds(4)
                    s = s & " Where plantwhs = 'K10' and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
            
                If ticket_receipt(ds!runid) = False Then
                    s = "Update bimp set onorder = onorder + " & ds(4)
                    s = s & " where branchwhs = '" & Format(ds!branch, "000") & "'"
                    s = s & " and plantwhs = 'K10'"
                    s = s & " and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If                                                              'jv010417
    'WMS Trailers - Sylacauga
    If plant_server_status("A10") = True Then                           'jv010417
        Set pb = CreateObject("ADODB.Connection")
        pb.Open a10ship
        's = "select runid, plant, branch, sku, sum(units) from trailers where pb_flag = 'N'"
        's = s & " and shipdate > '" & Format(Now, "M-d-yyyy") & "'"
        s = "select runid, plant, branch, sku, sum(units) from trailers"        'jv102615
        s = s & " where shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv102615
        s = s & " and pallets > 0"                  'jv040318
        s = s & " and plant = 52"
        s = s & " Group by runid, plant, branch, sku"
        Set ds = pb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ticket_post(ds!runid) = False Then
                    s = "Update bimp set plantpool = plantpool - " & ds(4)
                    s = s & " Where plantwhs = 'A10' and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
    
                If ticket_receipt(ds!runid) = False Then
                    s = "Update bimp set onorder = onorder + " & ds(4)
                    s = s & " where branchwhs = '" & Format(ds!branch, "000") & "'"
                    s = s & " and plantwhs = 'A10'"
                    s = s & " and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close: pb.Close
    Else                                                                'jv010417
        MsgBox "Sylacauga is offline.", vbOKOnly + vbInformation, "Sylacauga trailers not processed..."
        s = "select runid, plant, branch, sku, sum(units) from trailers"        'jv102615
        s = s & " where shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv102615
        s = s & " and plant = 52"
        s = s & " Group by runid, plant, branch, sku"
        Set ds = wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ticket_post(ds!runid) = False Then
                    s = "Update bimp set plantpool = plantpool - " & ds(4)
                    s = s & " Where plantwhs = 'A10' and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
    
                If ticket_receipt(ds!runid) = False Then
                    s = "Update bimp set onorder = onorder + " & ds(4)
                    s = s & " where branchwhs = '" & Format(ds!branch, "000") & "'"
                    s = s & " and plantwhs = 'A10'"
                    s = s & " and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If                                                              'jv010417
    
    Call process_r12_nonreceipts("ALL")                                 'JV090116
    
    calc_bimp_all
    
    Open bimpstat For Output As #1                          'jv022316
    Print #1, Format(Now, "M-dd-yyyy hh:mm am/pm")          'jv022316
    Close #1                                                'jv022316
    Screen.MousePointer = 0
End Sub

Sub delete_bimp_log(uname As String, fname As String, sqlx As String)           'jv030817
    Dim cfile As String
    cfile = "\\BBC-01-PRODTRK\wd\data\bimpdels.csv"
    Open cfile For Append As #1
    Write #1, Format(Now, "M-dd-yyyy hh:mm am/pm"); uname; fname; sqlx
    Close #1
End Sub

Sub export_branchbarcodes_bills()
    Dim sdate As String, edate As String, fdate As String
    Dim s As String, cfile As String, spath As String, sdir As String
    Dim logpath As String
    Screen.MousePointer = 11
    sdate = Format(DateAdd("d", -2, Now), "yyyymmdd")
    'edate = Format(Now, "yyyymmdd")
    edate = Format(DateAdd("d", 2, Now), "yyyymmdd")
    
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\brbarcodes.txt"
    Open cfile For Output As #4
    
    If plant_server_status("T10") = True Then
        logpath = "\\bbc-01-prodtrk\wd\pallogs\"
        spath = logpath & "bill*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                s = Mid(fdate, 5, 2) & "-" & Mid(fdate, 7, 2) & "-" & Mid(fdate, 1, 4)
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    Write #4, f16;              'Ticket
                    Write #4, s;                'File Date
                    Write #4, f4;               'Target
                    Write #4, f5;               'Product
                    Write #4, f6                'BarCode
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    
    If plant_server_status("K10") = True Then
        logpath = "\\bbba-03-dc\f\user\waredist\data\pallogs\"
        spath = logpath & "bill*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                s = Mid(fdate, 5, 2) & "-" & Mid(fdate, 7, 2) & "-" & Mid(fdate, 1, 4)
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    Write #4, f16;              'Ticket
                    Write #4, s;                'File Date
                    Write #4, f4;               'Target
                    Write #4, f5;               'Product
                    Write #4, f6                'BarCode
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    
    If plant_server_status("A10") = True Then
        logpath = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
        spath = logpath & "bill*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                s = Mid(fdate, 5, 2) & "-" & Mid(fdate, 7, 2) & "-" & Mid(fdate, 1, 4)
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    Write #4, f16;              'Ticket
                    Write #4, s;                'File Date
                    Write #4, f4;               'Target
                    Write #4, f5;               'Product
                    Write #4, f6                'BarCode
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    
    
    Close #4
    Screen.MousePointer = 0
End Sub


Sub export_branchbarcodes_ships()
    Dim sdate As String, edate As String, fdate As String
    Dim s As String, cfile As String, spath As String, sdir As String
    Dim logpath As String
    Screen.MousePointer = 11
    sdate = Format(DateAdd("d", -2, Now), "yyyymmdd")
    'edate = Format(Now, "yyyymmdd")
    edate = Format(DateAdd("d", 2, Now), "yyyymmdd")
    
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\brbarcodes.txt"
    Open cfile For Output As #4
    
    If plant_server_status("T10") = True Then
        logpath = "\\bbc-01-prodtrk\wd\pallogs\"
        'spath = logpath & "ship*.txt"
        spath = logpath & "bill*.txt"                               'jv030518
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                s = Mid(fdate, 5, 2) & "-" & Mid(fdate, 7, 2) & "-" & Mid(fdate, 1, 4)
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    Write #4, f2;              'Group
                    Write #4, s;                'File Date
                    Write #4, f4;               'Target
                    Write #4, f5;               'Product
                    Write #4, f6                'BarCode
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    
    If plant_server_status("K10") = True Then
        logpath = "\\bbba-03-dc\f\user\waredist\data\pallogs\"
        'spath = logpath & "ship*.txt"
        spath = logpath & "bill*.txt"                               'jv030518
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                s = Mid(fdate, 5, 2) & "-" & Mid(fdate, 7, 2) & "-" & Mid(fdate, 1, 4)
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    Write #4, f2;               'Group
                    Write #4, s;                'File Date
                    Write #4, f4;               'Target
                    Write #4, f5;               'Product
                    Write #4, f6                'BarCode
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    
    If plant_server_status("A10") = True Then
        logpath = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
        'spath = logpath & "ship*.txt"
        spath = logpath & "bill*.txt"                                   'jv030518
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                s = Mid(fdate, 5, 2) & "-" & Mid(fdate, 7, 2) & "-" & Mid(fdate, 1, 4)
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    Write #4, f2;               'Group
                    Write #4, s;                'File Date
                    Write #4, f4;               'Target
                    Write #4, f5;               'Product
                    Write #4, f6                'BarCode
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    
    
    Close #4
    Screen.MousePointer = 0
End Sub

Function groupitems_qty(psku As String, pplant As String, pbranch As String) As Long
    Dim ds As ADODB.Recordset, s As String, gqty As Integer
    Dim ts As ADODB.Recordset, rs As ADODB.Recordset                    'jv081516
    Dim rplant As Integer, rbranch As Integer
    If pplant = "T10" Then rplant = 50
    If pplant = "K10" Then rplant = 51
    If pplant = "A10" Then rplant = 52
    If pbranch = "ALL" Then
        rbranch = 0
    Else
        rbranch = Val(pbranch)
    End If
    gqty = 0
    s = "Select * from groupitems Where sku = '" & psku & "'"
    s = s & " and groupcode not in (select groupcode from trailers)"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "select * from trgroups where groupcode = '" & ds!groupcode & "'"
            Set ts = wdb.Execute(s)
            If ts.BOF = False Then
                ts.MoveFirst
                If ts!run1 > 0 Then
                    s = "select * from runs where id = " & ts!run1
                    Set rs = wdb.Execute(s)
                    If rs.BOF = False Then
                        If rs!loaded = rplant Then
                            If pbranch = "ALL" Then
                                If ds!qty1 > 0 Then gqty = gqty + ds!qty1           'jv081916
                            Else
                                If rs!destination = rbranch Then
                                    If ds!qty1 > 0 Then gqty = gqty + ds!qty1       'jv081916
                                End If
                            End If
                        End If
                    End If
                    rs.Close
                End If
                If ts!run2 > 0 Then
                    s = "select * from runs where id = " & ts!run2
                    Set rs = wdb.Execute(s)
                    If rs.BOF = False Then
                        If rs!loaded = rplant Then
                            If pbranch = "ALL" Then
                                If ds!qty2 > 0 Then gqty = gqty + ds!qty2           'jv081916
                            Else
                                If rs!destination = rbranch Then
                                    If ds!qty2 > 0 Then gqty = gqty + ds!qty2       'jv081916
                                End If
                            End If
                        End If
                    End If
                    rs.Close
                End If
                If ts!run3 > 0 Then
                    s = "select * from runs where id = " & ts!run3
                    Set rs = wdb.Execute(s)
                    If rs.BOF = False Then
                        If rs!loaded = rplant Then
                            If pbranch = "ALL" Then
                                If ds!qty3 > 0 Then gqty = gqty + ds!qty3           'jv081916
                            Else
                                If rs!destination = rbranch Then
                                    If ds!qty3 > 0 Then gqty = gqty + ds!qty3       'jv081916
                                End If
                            End If
                        End If
                    End If
                    rs.Close
                End If
                If ts!run4 > 0 Then
                    s = "select * from runs where id = " & ts!run4
                    Set rs = wdb.Execute(s)
                    If rs.BOF = False Then
                        If rs!loaded = rplant Then
                            If pbranch = "ALL" Then
                                If ds!qty4 > 0 Then gqty = gqty + ds!qty4           'jv081916
                            Else
                                If rs!destination = rbranch Then
                                    If ds!qty4 > 0 Then gqty = gqty + ds!qty4       'jv081916
                                End If
                            End If
                        End If
                    End If
                    rs.Close
                End If
            End If
            ts.Close
            ds.MoveNext
        Loop
    End If
    ds.Close
    'If gqty > 0 Then MsgBox s, vbOKOnly, "gq=" & gqty
    groupitems_qty = gqty
End Function

Public Sub import_r12_branch_qty(pwhs As String)
    Dim ds As ADODB.Recordset, s As String
    Dim pb As ADODB.Connection
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    Screen.MousePointer = 11
    
    'Clear Quota Qtys
    s = "Update bimp set onhand = 0, onorder = 0 where branchwhs = '" & pwhs & "'"
    wdb.Execute s
    
    'R12 Onhand
    s = "select o.subinventory_code, m.segment1, sum(o.transaction_quantity)"
    s = s & " from mtl_onhand_quantities o, mtl_system_items_b m"
    s = s & ", mtl_item_locations l"
    s = s & " where o.subinventory_code = '" & pwhs & "'"
    s = s & " and m.organization_id = o.organization_id"
    s = s & " and m.inventory_item_id = o.inventory_item_id"
    s = s & " and l.inventory_location_id = o.locator_id"
    s = s & " and l.segment1 > 'FLOOR   '"
    s = s & " and l.segment1 < 'FLOORZZZ'"
    s = s & " group by o.subinventory_code, m.segment1"
    s = s & " order by m.segment1"
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update bimp set onhand = " & ds(2)
            s = s & " Where branchwhs = '" & ds(0) & "'"
            s = s & " And sku = '" & ds(1) & "'"
            'MsgBox s
            wdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    'WMS Trailers - Brenham
    s = "select runid, plant, branch, sku, sum(units) from trailers"        'jv102615
    s = s & " where shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv102615
    s = s & " and plant = 50"
    s = s & " and branch = " & Format(Val(pwhs), "0")
    s = s & " Group by runid, plant, branch, sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ticket_receipt(ds!runid) = False Then
                s = "Update bimp set onorder = onorder + " & ds(4)
                s = s & " where branchwhs = '" & Format(ds!branch, "000") & "'"
                s = s & " and plantwhs = 'T10'"
                s = s & " and sku = '" & ds!sku & "'"
                'MsgBox s
                wdb.Execute s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    'WMS Trailers - Broken Arrow
    If plant_server_status("K10") = True Then                               'jv010417
        Set pb = CreateObject("ADODB.Connection")
        pb.Open k10ship
        s = "select runid, plant, branch, sku, sum(units) from trailers"        'jv102615
        s = s & " where shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv102615
        s = s & " and plant = 51"
        s = s & " and branch = " & Format(Val(pwhs), "0")
        s = s & " Group by runid, plant, branch, sku"
        Set ds = pb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ticket_receipt(ds!runid) = False Then
                    s = "Update bimp set onorder = onorder + " & ds(4)
                    s = s & " where branchwhs = '" & Format(ds!branch, "000") & "'"
                    s = s & " and plantwhs = 'K10'"
                    s = s & " and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close: pb.Close
    Else                                                                    'jv010417
        MsgBox "Broken Arrow is offline.", vbOKOnly + vbInformation, "Trailers not processed.."
        s = "select runid, plant, branch, sku, sum(units) from trailers"        'jv102615
        s = s & " where shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv102615
        s = s & " and plant = 51"
        s = s & " and branch = " & Format(Val(pwhs), "0")
        s = s & " Group by runid, plant, branch, sku"
        Set ds = wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ticket_receipt(ds!runid) = False Then
                    s = "Update bimp set onorder = onorder + " & ds(4)
                    s = s & " where branchwhs = '" & Format(ds!branch, "000") & "'"
                    s = s & " and plantwhs = 'K10'"
                    s = s & " and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If                                                                  'jv010417
    'WMS Trailers - Sylacauga
    If plant_server_status("A10") = True Then                               'jv010417
        Set pb = CreateObject("ADODB.Connection")
        pb.Open a10ship
        s = "select runid, plant, branch, sku, sum(units) from trailers"        'jv102615
        s = s & " where shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv102615
        s = s & " and plant = 52"
        s = s & " and branch = " & Format(Val(pwhs), "0")
        s = s & " Group by runid, plant, branch, sku"
        Set ds = pb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ticket_receipt(ds!runid) = False Then
                    s = "Update bimp set onorder = onorder + " & ds(4)
                    s = s & " where branchwhs = '" & Format(ds!branch, "000") & "'"
                    s = s & " and plantwhs = 'A10'"
                    s = s & " and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close: pb.Close
    Else                                                                    'jv010417
        MsgBox "Sylacauga server is offline.", vbOKOnly + vbInformation, "Trailers not processed.."
        s = "select runid, plant, branch, sku, sum(units) from trailers"        'jv102615
        s = s & " where shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv102615
        s = s & " and plant = 52"
        s = s & " and branch = " & Format(Val(pwhs), "0")
        s = s & " Group by runid, plant, branch, sku"
        Set ds = wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ticket_receipt(ds!runid) = False Then
                    s = "Update bimp set onorder = onorder + " & ds(4)
                    s = s & " where branchwhs = '" & Format(ds!branch, "000") & "'"
                    s = s & " and plantwhs = 'A10'"
                    s = s & " and sku = '" & ds!sku & "'"
                    wdb.Execute s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If                                                                  'jv010417
    
    Call process_r12_nonreceipts(pwhs)                          'jv090116
    'Open bimpstat For Output As #1                          'jv022316
    'Print #1, Format(Now, "M-dd-yyyy hh:mm am/pm")          'jv022316
    'Close #1                                                'jv022316
    Screen.MousePointer = 0
End Sub

Function last_branch_issue(psku As String, pwhs As String) As String
    Dim ds As ADODB.Recordset, s As String
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Function
    Screen.MousePointer = 11
    s = "select product_no, max(tran_date) from bolinf.inv_adj_input_detail"
    s = s & " where branch_no = '" & pwhs & "'"
    s = s & " and product_no = '" & psku & "'"
    s = s & " and tran_type = '1'"
    s = s & " and trunc(tran_date) > trunc(SYSDATE - 91)"
    If pwhs = "036" Then s = s & " and route_no not in ('01', '02', '03', '04')"
    s = s & " group by product_no"
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = Format(ds(1), "M-d-yyyy")
    Else
        s = Format(DateAdd("d", -91, Now), "M-d-yyyy")
    End If
    ds.Close
    last_branch_issue = s
    Screen.MousePointer = 0
End Function

Function last_branch_loads(psku As String, pwhs As String, sdate As String, edate As String) As Long
    Dim ds As ADODB.Recordset, s As String
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Function
    Screen.MousePointer = 11
    s = "select product_no, sum(tran_qty) from bolinf.inv_adj_input_detail"
    s = s & " where branch_no = '" & pwhs & "'"
    s = s & " and product_no = '" & psku & "'"
    s = s & " and tran_type = '1'"
    s = s & " and tran_date >= TO_DATE('" & Format(sdate, "dd-mmm-yy") & "')"
    s = s & " and tran_date <= TO_DATE('" & Format(edate, "dd-mmm-yy") & "')"
    If pwhs = "036" Then s = s & " and route_no not in ('01', '02', '03', '04')"
    s = s & " group by product_no"
    'MsgBox s
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = Format(ds(1), "0")
    Else
        s = "0"
    End If
    ds.Close
    last_branch_loads = Val(s)
    Screen.MousePointer = 0
End Function

Function last_branch_receipt(psku As String, pwhs As String, pqty As Integer) As String
    Dim ds As ADODB.Recordset, s As String
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Function
    Screen.MousePointer = 11
    s = "Select m.segment1, max(o.transaction_date)"
    s = s & " from mtl_material_transactions o, mtl_system_items_b m"
    s = s & " where o.subinventory_code = '" & pwhs & "'"
    s = s & " and m.organization_id = o.organization_id"
    s = s & " and m.inventory_item_id = o.inventory_item_id"
    s = s & " and m.segment1 = '" & psku & "'"
    s = s & " and o.transaction_quantity >= " & pqty
    s = s & " and o.source_code = 'RCV'"
    s = s & " and trunc(o.transaction_date) > trunc(SYSDATE - 181)"
    s = s & " group by m.segment1"
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = Format(ds(1), "M-d-yyyy")
    Else
        s = Format(DateAdd("d", -181, Now), "M-d-yyyy")
        'MsgBox psku & " " & pwhs, vbOKOnly + vbInformation, "no receipt"
    End If
    ds.Close
    last_branch_receipt = s
    Screen.MousePointer = 0
End Function

Sub Main()
    Dim ret As Long, s As String, i As Integer
    Dim lpbuff As String * 25
    check_hax
    ret = GetUserName(lpbuff, 25)
    bimpuserid = Left(lpbuff, InStr(lpbuff, Chr(0)) - 1)


    'Me.shipdb = "ODBC;DATABASE=WDShip;DSN=wdship"
    ''shipdb = "Driver={SQL Server};Server=bbc-01-wdsql;DATABASE=WDShip;UID=bbcship500;PWD=brenham500"
    'schdb = "Driver={SQL Server};Server=10.100.1.181;DATABASE=WDTruck;uid=bbctruck500;pwd=brenham500"
    a10bbsr = "Driver={SQL Server};Server=bbsy-01-wdsql;DATABASE=SYRacks;UID=bbcwd502;PWD=alabama502"
    k10bbsr = "Driver={SQL Server};Server=bbba-01-wdsql;DATABASE=BARacks;UID=bbcwd501;PWD=barrow501"
    t10bbsr = "Driver={SQL Server};Server=bbc-01-wdsql;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500"
    a10ship = "Driver={SQL Server};Server=bbsy-01-wdsql;DATABASE=SYShip;UID=bbcship502;PWD=Alabama502"
    k10ship = "Driver={SQL Server};Server=bbba-01-wdsql;DATABASE=BAShip;UID=bbcship501;PWD=Barrow501"
    'a10ship = "ODBC;DATABASE=WDShip;DSN=wdship"
    'k10ship = "ODBC;DATABASE=WDShip;DSN=wdship"
    
    
    'cs5db = "Driver={SQL Server};Server=bbsy-01-sqlsvr;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907"
    'Westfalia Upgrade
    cs5db = "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=sywms;PWD=!Sylacauga_WMS1907"
    'db5.Open "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
    
    'oradb = "odbc;database=pbelle;uid=apps;pwd=pb3113tx;dsn=pbelle"
    'r12connection = oradb
    'Set wdb = CreateObject("ADODB.Connection")
    'wdb.Open Me.shipdb
    'Set tsb = CreateObject("ADODB.Connection")
    'tsb.Open Me.schdb
    'Call build_skumast
    'Call build_branches
    'Call build_plants
    'refresh_branches
    'refresh_skus
    'Combo2.Clear
    'Combo2.AddItem "A10"
    'Combo2.AddItem "K10"
    'Combo2.AddItem "T10"
    'Combo2.AddItem "ALL"
    'Combo2.ListIndex = 0
    'r12access = False
    brzloss = "\\BBC-03-FILESVR\SharedGroups\wd\html\boutstock.csv"                            'jv032318
    bimpstat = "\\BBC-03-FILESVR\SharedGroups\wd\html\bimpstat.txt"
    r12connection = "odbc;database=pbelle;uid=apps;pwd=pb3113tx;dsn=pbelle"
    'r12connection = "Driver={Microsoft ODBC for Oracle};Server=pbelle;uid=apps;pwd=pb3113tx"
    Set wdb = CreateObject("ADODB.Connection")
    'wdb.Open "ODBC;DATABASE=WDShip;DSN=wdship"
    wdb.Open "Driver={SQL Server};Server=BBC-08-SQLSVR;DATABASE=WDShip;UID=bbcship500;PWD=brenham500"
    Set tsb = CreateObject("ADODB.Connection")
    tsb.Open "Driver={SQL Server};Server=10.100.1.181;DATABASE=WDTruck;uid=bbctruck500;pwd=brenham500"
    Call build_skumast
    Call build_branches
    Call build_plants
    r12access = False
    'plandist.Show
    bimpbanner.Show
End Sub

Function net_order_qty(pwhs As String, bwhs As String, psku As String) As Integer
    Dim ds As ADODB.Recordset, s As String, oplant As Integer, obranch As Integer, oqty As Integer
    oqty = 0
    oplant = 50
    If pwhs = "A10" Then oplant = 52
    If pwhs = "K10" Then oplant = 51
    obranch = Val(bwhs)
    If bwhs = "T10" Then obranch = 1
    If bwhs = "K10" Then obranch = 47
    If bwhs = "A10" Then obranch = 52
    s = "select netqty from brorders where plant = " & oplant
    s = s & " and branch = " & obranch
    s = s & " and sku = '" & psku & "'"
    s = s & " and netqty > 0"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            oqty = oqty + ds!netqty
            ds.MoveNext
        Loop
    End If
    ds.Close
    net_order_qty = oqty
End Function

Function plant_lowstock(pwhs As String, psku As String) As Integer
    Dim ss As ADODB.Recordset, s As String, i As Integer
    i = 1
    s = "select lowqty from bimp where plantwhs = '" & pwhs & "'"
    s = s & " and sku = '" & psku & "' and lowqty > 0 order by lowqty desc"
    Set ss = wdb.Execute(s)
    If ss.BOF = False Then
        ss.MoveFirst
        i = ss!lowqty
    End If
    ss.Close
    plant_lowstock = i
End Function

Function plant_outstock(pwhs As String, psku As String) As Integer
    Dim ss As ADODB.Recordset, s As String, i As Integer
    i = 1
    s = "select outqty from bimp where plantwhs = '" & pwhs & "'"
    s = s & " and sku = '" & psku & "' and outqty > 0 order by outqty desc"
    Set ss = wdb.Execute(s)
    If ss.BOF = False Then
        ss.MoveFirst
        i = ss!outqty
    End If
    ss.Close
    plant_outstock = i
End Function

Function plant_server_status(pwhs As String) As Boolean                 'jv010417
    Dim ss As ADODB.Recordset, s As String, sstat As Boolean
    sstat = False
    If pwhs = "001" Then pwhs = "T10"
    If pwhs = "047" Then pwhs = "K10"
    If pwhs = "052" Then pwhs = "A10"
    s = "select listdisplay from valuelists where listname = 'wdserverstatus'"
    s = s & " and listreturn = '" & pwhs & " OnLine'"
    Set ss = wdb.Execute(s)
    If ss.BOF = False Then
        ss.MoveFirst
        If Trim(LCase(ss!listdisplay)) = "true" Then sstat = True
    End If
    ss.Close
    plant_server_status = sstat
End Function

Function plant_transfers(plantcode As String, psku As String) As Long
    Dim ss As ADODB.Recordset, s As String, np As Long
    
    np = 0
    's = "select poolsched from bimp where sku = '" & psku & "'"
    's = s & " and plantwhs = '" & plantcode & "'"
    's = s & " and poolsched > 0"
    'Set ss = wdb.Execute(s)
    'If ss.BOF = False Then
    '    ss.MoveFirst
    '    np = ss(0)
    'End If
    'ss.Close
            
    s = "select sku, sum(thiswknewpals), sum(nextwknewpals), sum(onorder / roqty) from bimp"
    s = s & " where sku = '" & psku & "'"
    If plantcode = "T10" Then
        s = s & " and plantwhs in ('A10', 'K10') and branchwhs = '001'"
    End If
    If plantcode = "K10" Then
        s = s & " and plantwhs in ('A10', 'T10') and branchwhs = '047'"
    End If
    If plantcode = "A10" Then
        s = s & " and plantwhs in ('K10', 'T10') and branchwhs = '052'"
    End If
    s = s & " group by sku"
    Set ss = wdb.Execute(s)
    If ss.BOF = False Then
        ss.MoveFirst
        If IsNull(ss(1)) = False Then
            np = np + ss(1)
        End If
        If IsNull(ss(2)) = False Then
            np = np + ss(2)
        End If
        If IsNull(ss(3)) = False Then
            np = np + ss(3)
        End If
    End If
    ss.Close
            
    s = "select sku, sum(netqty) from brorders where sku = '" & psku & "'"
    If plantcode = "T10" Then
        s = s & " and plant in (52, 51) and branch = 1"
    End If
    If plantcode = "K10" Then
        s = s & " and plant in (52, 50) and branch = 47"
    End If
    If plantcode = "A10" Then
        s = s & " and plant in (51, 50) and branch = 52"
    End If
    s = s & " group by sku having sum(netqty) <> 0"
    Set ss = wdb.Execute(s)
    If ss.BOF = False Then
        ss.MoveFirst
        If IsNull(ss(1)) = False Then
            np = np + ss(1)
        End If
    End If
    ss.Close
    'If np > 0 Then MsgBox plantcode & " " & psku & " = " & np
    plant_transfers = np
End Function

Public Sub process_bimp_discontinued()
    Dim ds As ADODB.Recordset, s As String, t As String
    Screen.MousePointer = 11
    'clear discflags
    s = "Update bimp set skunotes = ' ' where discflag = 'Y'"   'jv082818
    wdb.Execute s                                               'jv082818
    s = "Update bimp set discflag = 'N' where discflag <> 'B'"  'jv082818
    wdb.Execute s
    's = "Update bimp set skunotes = ' ' where promoflag = 'N'"
    'wdb.Execute s
    s = "select * from discont" ' where sku = '729'"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            t = "Discontinued: " & Format(ds!discdate, "MM-dd-yyyy")
            t = t & " - " & ds!discomm
            s = "Update bimp set discflag = 'Y', skunotes = '" & t & "'"
            s = s & " Where sku = '" & ds!sku & "'"
            'MsgBox s
            wdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
End Sub

Function pallet_space(psku As String, pqty As Long) As Integer          'jv022516
    Dim pc As Integer, i As Integer
    pc = skurec(Val(psku)).pallet
    If Int(pqty / pc) = pqty / pc Then
        i = Int(pqty / pc)
    Else
        i = Int(pqty / pc) + 1
    End If
    pallet_space = i
End Function

Public Sub process_bimp_lastissue()
    Dim ds As ADODB.Recordset, s As String, os As ADODB.Recordset, i As Long
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    Screen.MousePointer = 11
    salesdays = bimp_sales_days                                                 'jv050117
    'R12 Last Receipt Data
    s = "select product_no,branch_no,max(tran_date) from bolinf.inv_adj_input_detail"
    s = s & " where tran_type = '1'"
    's = s & " and branch_no = '003'"
    's = s & " and trunc(tran_date) > trunc(SYSDATE - 30)"
    s = s & " and trunc(tran_date) > trunc(SYSDATE - 91)"                           'jv031816
    's = s & " and trunc(tran_date) > trunc(SYSDATE - " & Format(salesdays + 1, "0") & ")"    'jv050117
    s = s & " group by product_no,branch_no"
    s = s & " order by product_no,branch_no"
    'MsgBox s
    i = 0
    Set os = r12db.Execute(s)
    If os.BOF = False Then
        os.MoveFirst
        Do Until os.EOF
            s = "Select id, discflag from bimp where branchwhs = '" & os(1) & "'"
            s = s & " and sku = '" & os(0) & "'"
            s = s & " and onhand < 1"
            s = s & " and discflag = 'N'"
            Set ds = wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                Do Until ds.EOF
                    'If ds!discflag = "N" Then
                        s = "Update bimp set skunotes = '" & Format(os(2), "M-d-yyyy") & "'"
                        s = s & " Where id = " & ds(0)
                        'MsgBox s, vbOKOnly, os(0) & " " & os(1)
                        i = i + 1
                        wdb.Execute s
                    'End If
                    ds.MoveNext
                Loop
            End If
            ds.Close
            os.MoveNext
        Loop
    End If
    os.Close
    Screen.MousePointer = 0
    'MsgBox i
End Sub

Public Sub process_bimp_lastreceipt()
    Dim ds As ADODB.Recordset, s As String
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    Screen.MousePointer = 11
    'R12 Last Receipt Data
    s = "select m.segment1, o.subinventory_code, max(date_received)" & _
        " from mtl_system_items_b m, mtl_onhand_quantities o, mtl_item_locations l" & _
        " where m.segment1 >= '100' and m.segment1 <= '9999'" & _
        " and o.organization_id = m.organization_id" & _
        " and o.inventory_item_id = m.inventory_item_id" & _
        " and trunc(o.date_received) > trunc(sysdate - 30)" & _
        " and l.inventory_location_id = o.locator_id" & _
        " and l.segment1 > 'FLOOR   '" & _
        " and l.segment1 < 'FLOORZZZ'" & _
        " group by m.segment1, o.subinventory_code" & _
        " order by m.segment1, o.subinventory_code"
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update bimp set lastrecpt = '" & Format(ds(2), "M-d-yyyy") & "'"
            s = s & " Where branchwhs = '" & ds(1) & "'"
            s = s & " and sku = '" & ds(0) & "'"
            'MsgBox s, vbOKOnly, ds(2)
            wdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
End Sub

Sub process_r12_nonreceipts(pwhs As String)
    Dim ds As ADODB.Recordset, ts As ADODB.Recordset, s As String
    Dim r12 As Boolean, blit As String, cfile As String, rid As String, elit As String
    Dim pplant As String, pb As ADODB.Connection
    Dim b1 As Integer, b2 As Integer
    If pwhs = "ALL" Then
        b1 = 1: b2 = 99
    Else
        b1 = Val(pwhs): b2 = b1
    End If
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\data\norcptr12.txt"
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    Screen.MousePointer = 11
    Open cfile For Output As #1
    s = "Select t.shipment_number, sum(t.transaction_quantity)"
    s = s & " From mtl_material_transactions t"
    s = s & " Where t.transaction_date > sysdate - 5"
    s = s & " and t.shipment_number > ' '"
    s = s & " and t.source_code in ('RCV', 'TRAILER TRANSFER')"
    s = s & " group by t.shipment_number"
    s = s & " Having Sum(t.transaction_quantity) < 0"
    s = s & " order by t.shipment_number"
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            r12 = False
            rid = Left(ds(0), Len(ds(0)) - 1)
            
            
            pplant = "T10"
            s = "select t.subinventory_code from mtl_material_transactions t"
            s = s & " where t.shipment_number = '" & ds(0) & "'"
            s = s & " and t.source_code = 'TRAILER TRANSFER'"
            Set ts = r12db.Execute(s)
            If ts.BOF = False Then
                ts.MoveFirst
                pplant = ts(0)
            End If
            ts.Close
                
            If plant_server_status(pplant) = True Then                          'jv010417
                If pplant <> "T10" And pplant <> "001" Then
                    Set pb = CreateObject("ADODB.Connection")
                    If pplant = "K10" Or pplant = "047" Then pb.Open k10ship
                    If pplant = "A10" Or pplant = "052" Then pb.Open a10ship
                End If
            
                s = "select runid from trailers where runid = " & rid
                s = s & " and shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv091416
            
                If pplant = "T10" Or pplant = "001" Then
                    Set ts = wdb.Execute(s)
                Else
                    Set ts = pb.Execute(s)
                End If
            
                If ts.BOF = False Then
                    ts.MoveFirst
                Else
                    r12 = True
                End If
                ts.Close
                If pplant <> "T10" And pplant <> "001" Then pb.Close
            Else                                                                'jv010417
                s = "WD Server for " & pplant & " is not on line.  Non-receipts cannot be processed."
                MsgBox s, vbOKOnly + vbExclamation, "Ticket #" & rid
                s = "select runid from trailers where runid = " & rid
                s = s & " and shipdate >= '" & Format(Now, "M-d-yyyy") & "'"          'jv091416
                Set ts = wdb.Execute(s)
                If ts.BOF = False Then
                    ts.MoveFirst
                Else
                    r12 = True
                End If
                ts.Close
            End If                                                              'jv010417
            
            If r12 = True Then
                s = "select t.subinventory_code, i.segment1, i.description, t.transaction_quantity," & _
                " t.transaction_uom, t.source_code, t.transaction_reference, t.transaction_date" & _
                " from mtl_material_transactions t, mtl_system_items_b i" & _
                " where t.shipment_number in ('" & rid & "P', '" & rid & "W')" & _
                " and i.inventory_item_id = t.inventory_item_id" & _
                " and i.organization_id = t.organization_id" & _
                " order by t.source_code, i.segment1, t.subinventory_code"
                Set ts = r12db.Execute(s)
                'MsgBox s
                elit = " "                                                                      'jv020717
                If ts.BOF = False Then
                    ts.MoveFirst
                    Do Until ts.EOF
                        If ts(6) > "0" Then                                                     'jv020717
                            blit = Left(ts(6), Len(ts(6)) - 7)  'Trans Reference Branch Name
                        Else
                            elit = "Warning!  Check out trailer receipts for ticket number: " & rid & "."   'jv020717
                        End If
                        For k = b1 To b2
                            If UCase(branchrec(k).branchname) = UCase(blit) Then
                                'MsgBox blit & "=" & i
                                s = "Update bimp set onorder = onorder + " & (Val(ts(3)) * -1)
                                s = s & " where plantwhs = '" & ts(0) & "'"
                                s = s & " and branchwhs = '" & Format(k, "000") & "'"
                                s = s & " and sku = '" & ts(1) & "'"
                                'MsgBox s
                                wdb.Execute s
                                Print #1, rid & ": " & blit & " " & s
                                Exit For
                            End If
                        Next k
                        ts.MoveNext
                    Loop
                End If
                ts.Close
                If elit > " " Then                                                      'jv020717
                    MsgBox elit, vbOKOnly + vbInformation, "trailer receipt.."          'jv020717
                End If                                                                  'jv020717
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    Close #1
    Screen.MousePointer = 0
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
    End If
    r12lot = s
End Function

Function ticket_post(rid As String) As Boolean
    Dim ds As ADODB.Recordset, q As String, tflag As Boolean
    q = "select shipment_number, source_code from mtl_material_transactions"
    q = q & " where shipment_number = '" & rid & "P' and source_code = 'TRAILER TRANSFER'"
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        tflag = True
    Else
        tflag = False
    End If
    ds.Close
    If tflag = True Then
        q = "update runs set runstat = 'In Transit' where id = " & rid
        'MsgBox q
        wdb.Execute q
    End If
    ticket_post = tflag
End Function

Function ticket_receipt(rid As String) As Boolean
    Dim ds As ADODB.Recordset, q As String, tflag As Boolean
    q = "select shipment_number, source_code from mtl_material_transactions"
    q = q & " where shipment_number = '" & rid & "P' and source_code = 'RCV'"
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        tflag = True
    Else
        tflag = False
    End If
    ds.Close
    If tflag = True Then
        q = "update runs set runstat = 'Received' where id = " & rid
        'MsgBox q
        wdb.Execute q
    End If
    ticket_receipt = tflag
End Function

Sub update_testdb()
    Dim cb As ADODB.Connection, cs As ADODB.Recordset
    Dim tb As ADODB.Connection, ts As ADODB.Recordset, z As Long, s As String
    Screen.MousePointer = 11
    Set cb = CreateObject("ADODB.Connection")
    Set tb = CreateObject("ADODB.Connection")
    tb.Open "ODBC;DATABASE=WDShip;DSN=wdship"
    cb.Open "Driver={SQL Server};Server=bbc-01-wdsql;DATABASE=WDShip;UID=bbcship500;PWD=brenham500"
    
    'Stock History
    tb.Execute "delete from stockhistory"
    s = "select * from stockhistory order by id"
    Set cs = cb.Execute(s)
    If cs.BOF = False Then
        cs.MoveFirst
        Do Until cs.EOF
            s = "Insert into stockhistory (id, branchwhs, sku, startdate, enddate, postdate,"
            s = s & " totaldays, daysin, daysout, loads) Values (" & cs!id
            s = s & ", '" & cs!branchwhs & "'"
            s = s & ", '" & cs!sku & "'"
            s = s & ", '" & cs!startdate & "'"
            s = s & ", '" & cs!enddate & "'"
            s = s & ", '" & cs!postdate & "'"
            s = s & ", " & cs!totaldays
            s = s & ", " & cs!daysin
            s = s & ", " & cs!daysout
            s = s & ", " & cs!loads & ")"
            tb.Execute s
            z = cs!id + 1
            cs.MoveNext
        Loop
    End If
    cs.Close
    s = "Update sequences set sequence_id = " & z & " where seq = 'stockhistory'"
    tb.Execute s
    
    'BIMP
    tb.Execute "delete from bimp"
    s = "select * from bimp order by id"
    Set cs = cb.Execute(s)
    If cs.BOF = False Then
        cs.MoveFirst
        Do Until cs.EOF
            s = "Insert into bimp (id, plantwhs, branchwhs, sku, onhand, onorder, sales,"
            s = s & " undiff, paldiff, ohpct, roqty, pctgain, needqty, bimpstatus, promoqty, lowqty,"
            s = s & " outqty, quotapct, plantpool, poolsched, discflag, promoflag, lowflag, outflag,"
            s = s & " lastrecpt, skunotes, thiswknewpals, nextwknewpals) Values (" & cs!id
            s = s & ", '" & cs!plantwhs & "'"
            s = s & ", '" & cs!branchwhs & "'"
            s = s & ", '" & cs!sku & "'"
            s = s & ", " & cs!onhand
            s = s & ", " & cs!onorder
            s = s & ", " & cs!sales
            s = s & ", " & cs!undiff
            s = s & ", " & cs!paldiff
            s = s & ", " & cs!ohpct
            s = s & ", " & cs!roqty
            s = s & ", " & cs!pctgain
            s = s & ", " & cs!needqty
            s = s & ", '" & cs!bimpstatus & "'"
            s = s & ", " & cs!promoqty
            s = s & ", " & cs!lowqty
            s = s & ", " & cs!outqty
            s = s & ", " & cs!quotapct
            s = s & ", " & cs!plantpool
            s = s & ", " & cs!poolsched
            s = s & ", '" & cs!discflag & "'"
            s = s & ", '" & cs!promoflag & "'"
            s = s & ", '" & cs!lowflag & "'"
            s = s & ", '" & cs!outflag & "'"
            s = s & ", '" & cs!lastrecpt & "'"
            s = s & ", '" & cs!skunotes & " '"
            s = s & ", 0" '& cs!thiswknewpals
            s = s & ", 0)" '& cs!nextwknewpals & ")"
            tb.Execute s
            z = cs!id + 1
            cs.MoveNext
        Loop
    End If
    cs.Close
    s = "Update sequences set sequence_id = " & z & " where seq = 'bimp'"
    tb.Execute s
    
    'SKU Master
    s = "select * from skumast order by sku"
    Set cs = cb.Execute(s)
    If cs.BOF = False Then
        cs.MoveFirst
        Do Until cs.EOF
            s = "select * from skumast where sku = '" & cs!sku & "'"
            Set ts = tb.Execute(s)
            If ts.BOF = False Then
                s = "Update skumast set pallet = " & cs!pallet & ", numwrap = " & cs!numwrap
                s = s & " Where sku = '" & cs!sku & "'"
            Else
                s = "Insert into skumast (sku, fgdesc, fgunit, psource, whs_num, proddesc, prodtype,"
                s = s & " prodclass, sales_class, invoice_no, pallet, upc,"
                s = s & " numwrap) Values ('" & cs!sku & "'"
                s = s & ", '" & cs!fgdesc & "'"
                s = s & ", '" & cs!fgunit & "'"
                s = s & ", " & cs!psource
                s = s & ", " & cs!whs_num
                s = s & ", '" & cs!proddesc & " '"
                s = s & ", '" & cs!prodtype & " '"
                s = s & ", '" & cs!prodclass & " '"
                s = s & ", '" & cs!sales_class & " '"
                s = s & ", '" & cs!invoice_no & " '"
                's = s & ", " & cs!gal_divisor
                's = s & ", " & cs!gl_number
                s = s & ", " & cs!pallet
                s = s & ", '" & cs!upc & " '"
                's = s & ", " & cs!bulku
                's = s & ", " & cs!bundle
                's = s & ", " & cs!onhand
                's = s & ", " & cs!unlbs
                s = s & ", " & cs!numwrap & ")"
            End If
            ts.Close
            tb.Execute s
            'MsgBox s
            cs.MoveNext
        Loop
    End If
    cs.Close
    
    cb.Close
    tb.Close
    Screen.MousePointer = 0
End Sub

Public Sub vb_elog(eno As Long, edesc As String, pform As String, psub As String, puser As String)
    Dim i As Integer, s As String, cfile As String
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\images\vberrors.txt"
    i = FreeFile(1)
    Open cfile For Append As #i
    Write #i, eno, edesc, pform, psub, Format(Now, "M-d-yyyy h:mm am/pm"), puser
    Close #i
End Sub

Function wd_seq(tbname As String) As Long
    Dim sSql As String
    Dim i As Long
    Dim ds As ADODB.Recordset
    sSql = "Select sequence_id From sequences where seq = '" & tbname & "'"
    Set ds = wdb.Execute(sSql)
    If ds.BOF = False Then
        ds.MoveFirst
        i = ds!sequence_id + 1
        sSql = "Update sequences Set sequence_id = " & i & " Where seq = '" & tbname & "'"
        wdb.Execute (sSql)
    Else
        i = 100
        sSql = "Insert Into sequences (sequence_id, seq) Value (" & i & ",'" & tbname & "')"
        wdb.Execute (sSql)
    End If
    ds.Close
    wd_seq = i
End Function

