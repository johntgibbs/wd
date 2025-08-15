Attribute VB_Name = "shipsaeutil"
Public Sdb As adodb.Connection
Public Wdb As adodb.Connection
Public localAppDataPath As String
Public htmlTempFile As String
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

Type skuinfo
    sku As String
    unit As String
    desc As String
    psrc As Integer
    pallet As Integer
    whs As Integer
End Type

Public Type labpic
    sku As String
    package As String
    name1 As String
    name2 As String
    name3 As String
End Type

Global skurec(0 To 9999) As skuinfo
Global eno As Long
Global edesc As String
Global labpix(9999) As labpic
Global labfmtfile As String

Sub insert_trans(pt As ptask)
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String, zid As Long
    Set db = CreateObject("ADODB.Connection")
    db.Open (Form1.bbsr)
    s = "select * from paltasks where status = 'COMP'"
    'Set ds = db.Execute(s)
    Set ds = New adodb.Recordset
    ds.Open s, db, adOpenKeyset, adLockOptimistic, adCmdText
    If ds.BOF = False Then
        ds.MoveFirst
        ds!area = pt.area
        ds!description = pt.description
        ds!source = pt.source
        ds!target = pt.target
        ds!product = pt.product
        ds!palletid = pt.palletid
        ds!qty = CLng(Val(pt.qty))
        ds!uom = pt.uom
        ds!lotnum = pt.lotnum
        ds!units = CLng(Val(pt.units))
        ds!lotnum2 = pt.lotnum2
        ds!units2 = CLng(Val(pt.units2))
        ds!status = pt.status
        ds!userid = pt.userid
        ds!trandate = pt.trandate
        ds.Update
    Else
        zid = wd_seq("PalTasks", Form1.bbsr)
        s = "INSERT INTO PalTasks (ID, Area, Description, Source, Target, Product,"
        s = s & " PalletID, Qty, Uom, LotNum, Units, LotNum2, Units2, Status, Userid,"
        s = s & " TranDate, ReqID) VALUES (" & zid & ","
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
        s = s & "' ')"
        db.Execute s
        'ds.AddNew
    End If
    ds.Close
    db.Close
End Sub

Function r12_lot(plot As String, ocode As String) As String
    Dim s As String, myear As Integer, mdays As Integer
    If Len(plot) >= 5 Then
        myear = Val(Left(plot, 2))
        mdays = Val(mid(plot, 3, 3)) - 1
        s = "1-1-20" & Left(plot, 2)
        s = Format(DateAdd("d", mdays, s), "MMddyy")
        s = Left(s, 4) & Format(myear + 2, "00")
        If Len(plot) > 5 Then
            s = s & " " & Right(plot, 1)
        Else
            s = s & " " & ocode
        End If
    Else
        s = "LOT1"
    End If
    r12_lot = s
End Function

Sub update_trans(pt As ptask)
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String
    Set db = CreateObject("ADODB.Connection")
    db.Open (Form1.bbsr)
    s = "select * from paltasks where id = " & pt.id
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        ds!area = pt.area
        ds!description = pt.description
        ds!source = pt.source
        ds!target = pt.target
        ds!product = pt.product
        ds!palletid = pt.palletid
        ds!qty = CLng(pt.qty)
        ds!uom = pt.uom
        ds!lotnum = pt.lotnum
        ds!units = CLng(pt.units)
        ds!lotnum2 = pt.lotnum2
        ds!units2 = CLng(pt.units2)
        ds!status = pt.status
        ds!userid = pt.userid
        ds!trandate = pt.trandate
        ds!reqid = pt.reqid
        ds.Update
    Else
        MsgBox "update trans failed to find record..."
    End If
    ds.Close
    db.Close
End Sub

Function wd_seq(tname As String, conString As String) As Long
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String, i As Long
    Set db = CreateObject("ADODB.Connection")
    db.Open (conString)
    s = "select sequence_id from sequences where seq = '" & tname & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        i = ds!sequence_id + 1
        'ds!sequence_id = i
        'ds.Update
    End If
    ds.Close
    Set ds = db.Execute("UPDATE sequences SET sequence_id = " & i & " WHERE seq = '" & tname & "'")
    db.Close
    wd_seq = i
End Function

Sub insert_ship_infc(gc As String, psku As String, pwhs As Integer, pqty As Integer, psize As Integer)
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String, i As Long
    Set db = CreateObject("ADODB.Connection")
    db.Open (Form1.bbsr)
    s = "select * from ship_infc where order_num = '" & gc & "'"
    s = s & " and sku = '" & psku & "'"
    s = s & " and to_whse_num = " & pwhs
    Set ds = New adodb.Recordset
    ds.Open s, db, adOpenKeyset, adLockOptimistic, adCmdText
    If ds.BOF = False Then
        'MsgBox s
        ds.MoveFirst
        ds!order_qty = ds!order_qty + pqty
        If ds!order_qty <= 0 Then
            ds!ship_status = "CANC"
        Else
            If ds!order_qty <= ship_plt_qty Then
                ds!ship_status = "DONE"
            Else
                ds!ship_status = "NEW"
            End If
        End If
        ds!gmasize = psize
        ds.Update
    Else
        ds.Close
        s = "select * from ship_infc where ship_status in ('CANC','DONE')"
        Set ds = New adodb.Recordset
        ds.Open s, db, adOpenKeyset, adLockOptimistic, adCmdText
        If ds.BOF = False Then
            'MsgBox s
            ds.MoveFirst
            ds!order_num = gc
            ds!sku = psku
            ds!lot_num = " "
            ds!ship_date = Format(Now, "MM-dd-yyyy")
            ds!order_qty = pqty
            ds!ship_uom_qty = 0
            ds!ship_plt_qty = 0
            ds!ship_status = "NEW"
            ds!to_whse_num = pwhs
            ds!to_vert_loc = 2
            If pwhs = "1" Then
                ds!to_horz_loc = 18
                ds!to_rack_side = "L"
            Else
                If pwhs = "2" Then
                    ds!to_horz_loc = 22
                    ds!to_rack_side = "R"
                Else
                    If pwhs = "3" Then
                        ds!to_horz_loc = 43
                        ds!to_rack_side = "R"
                    Else
                        ds!to_horz_loc = 0
                        ds!to_rack_side = "R"
                    End If
                End If
            End If
            ds!resv_strategy = "A"
            ds!gmasize = psize
            ds.Update
        Else
            s = "INSERT INTO ship_infc (ID, order_num, sku, lot_num, ship_date, order_qty,"
            s = s & " ship_uom_qty, ship_plt_qty, ship_status, to_whse_num, to_vert_loc,"
            s = s & " to_horz_loc, to_rack_side, resv_strategy, gmasize) VALUES ("
            s = s & wd_seq("Ship_infc", Form1.bbsr) & ","
            s = s & "'" & gc & "',"
            s = s & "'" & psku & "',"
            s = s & "' ',"
            s = s & "'" & Format(Now, "MM-dd-yyyy") & "',"
            s = s & pqty & ","
            s = s & "0,"
            s = s & "0,"
            s = s & "'NEW',"
            s = s & pwhs & ","
            s = s & "2,"
            If pwhs = "1" Then
                s = s & "18,'L',"
            Else
                If pwhs = "2" Then
                    s = s & "22,'R',"
                Else
                    If pwhs = "3" Then
                        s = s & "43,'R',"
                    Else
                        s = s & "0,'R',"
                    End If
                End If
            End If
            s = s & "'A',"
            s = s & psize & ")"
            db.Execute s
            'MsgBox s
        End If
    End If
    ds.Close: db.Close
End Sub

Public Sub vb_elog(eno As Long, edesc As String, pform As String, psub As String, puser As String)
    Dim i As Integer, s As String, cfile As String
    On Error GoTo vberror
    cfile = localAppDataPath & "\error.log"
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

Sub build_skumast()
    Dim ds As adodb.Recordset, sqlx As String, i As Integer
    'On Error GoTo vberror
    sqlx = "select * from skumast order by sku"
    Set ds = Sdb.Execute(sqlx)
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
            ds.MoveNext
        Loop
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, "Sub", "build_skumast", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, "Sub: build_skumast - Error Number: " & eno
        End
    End If
End Sub

Public Function fixquotes(s As String) As String
    Dim i As Integer, k As Integer, rs As String
    rs = ""
    i = 1: k = 1
    Do Until i = 0
        i = InStr(k, s, "'")
        If i = 0 Then
            rs = rs & mid(s, k, Len(s))
            Exit Do
        Else
            rs = rs & mid(s, k, i - k) & "''"
            k = i + 1
        End If
    Loop
    fixquotes = rs
End Function

Public Sub load_labpics()
    Dim s As String, ts As Integer, te As Integer, psku As String
    Open labfmtfile For Input As #1
    Do Until EOF(1)
        Line Input #1, s
        te = InStr(1, s, Chr(9))
        psku = mid(s, 1, te - 1)
        labpix(Val(psku)).sku = psku
        ts = te + 1
        te = InStr(ts, s, Chr(9))
        labpix(Val(psku)).package = mid(s, ts, te - ts)
        ts = te + 1
        te = InStr(ts, s, Chr(9))
        labpix(Val(psku)).name1 = mid(s, ts, te - ts)
        ts = te + 1
        te = InStr(ts, s, Chr(9))
        labpix(Val(psku)).name2 = mid(s, ts, te - ts)
        ts = te + 1
        If ts < Len(s) Then
            labpix(Val(psku)).name3 = mid(s, ts, Len(s) - (ts - 1))
        Else
            labpix(Val(psku)).name3 = " "
        End If
    Loop
    Close #1
    'MsgBox labpix(777).name3
End Sub

Public Function fixamps(s As String) As String
    Dim i As Integer, k As Integer, rs As String
    rs = ""
    i = 1: k = 1
    Do Until i = 0
        i = InStr(k, s, "&")
        If i = 0 Then
            rs = rs & mid(s, k, Len(s))
            Exit Do
        Else
            rs = rs & mid(s, k, i - k) & "&&"
            k = i + 1
        End If
    Loop
    fixamps = rs
End Function

Public Function StringReplace(InputString As String, ReplaceChar As String, ReplaceWith As String)
    Dim Pos As Integer
    Dim Tmp As String
    Pos = InStr(1, InputString, ReplaceChar, vbTextCompare)
    While Pos > 0
        Tmp = Left(InputString, Pos - Len(ReplaceChar)) ' Build left half of string, before the replacement character
        Tmp = Tmp & Right(InputString, Len(InputString) - Pos) ' Build right half of string, after replace.
        Pos = InStr(1, Tmp, ReplaceChar, vbTextCompare)
        InputString = Tmp
    Wend
    
    StringReplace = InputString
End Function

Public Function OpenFileInExcel(FileToOpen As String) As Boolean
    ' Open in Excel 2019 if it's installed
    If Len(Dir("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE")) <> 0 Then
        i = Shell("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE " & FileToOpen, vbNormalFocus)
        OpenFileInExcel = True
        Exit Function
    End If
    
    If Len(Dir("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE")) <> 0 Then
        i = Shell("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE " & FileToOpen, vbNormalFocus)
        OpenFileInExcel = True
        Exit Function
    End If
    
    ' Otherwise open in Excel 2010
    If Len(Dir("C:\Program Files\Microsoft Office\Office14\EXCEL.EXE")) <> 0 Then
        i = Shell("C:\Program Files\Microsoft Office\Office14\EXCEL.EXE " & FileToOpen, vbNormalFocus)
        OpenFileInExcel = True
        Exit Function
    End If
    
    If Len(Dir("C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE")) <> 0 Then
        i = Shell("C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE " & FileToOpen, vbNormalFocus)
        OpenFileInExcel = True
        Exit Function
    End If
    OpenFileInExcel = False
End Function
