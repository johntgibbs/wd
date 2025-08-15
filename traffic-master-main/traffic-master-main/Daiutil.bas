Attribute VB_Name = "Daiutil"
Public tbbsr As String
Public daioradb As String
Public daisqldb As String
Public dailogs As String
Public localAppDataPath As String

Type daiexprct
    action As String
    sOrderID As String
    dExpectedDate As String
    sItem As String
    sLot As String
    fExpectedQuantity As String
    sStoreDestination As String
    sRouteID As String
    sHoldReason As String
End Type

Type daimessagerec
    dhostmodifytime As String
    imessagesequence As String
    smessageidentifier As String
    smessage As String
    bbcidentity As String
    bbcstatus As String
End Type

Type saerequesttype
    id As String
    userid As String
    warehouse As String
    area As String
    func As String
    barcode As String
End Type

Type saeresponsetype
    reqid As String
    moveid As String
    warehouse As String
    area As String
    func As String
    fromloc As String
    toloc As String
    uom As String
    qty As String
    product As String
    barcode As String
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

Function Dai_expected_receipt(d As daiexprct) As String
    Dim s As String
    s = "<?xml version=" & Chr(34) & "1.0" & Chr(34)
    s = s & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>" & vbCrLf
    s = s & "<!DOCTYPE ExpectedReceiptMessage SYSTEM " & Chr(34) & "wrxj.dtd" & Chr(34) & ">" & vbCrLf
    s = s & "<ExpectedReceiptMessage>" & vbCrLf
    s = s & "<ExpectedReceipt action=" & Chr(34) & d.action & Chr(34)
    s = s & " sOrderID=" & Chr(34) & d.sOrderID & Chr(34) & ">" & vbCrLf
    s = s & "<ExpectedReceiptHeader>" & vbCrLf
    s = s & "<dExpectedDate>" & d.dExpectedDate & "</dExpectedDate>" & vbCrLf
    s = s & "</ExpectedReceiptHeader>" & vbCrLf
    s = s & "<ExpectedReceiptLine sItem=" & Chr(34) & d.sItem & Chr(34) & " sLot=" & Chr(34) & d.sLot & Chr(34) & ">" & vbCrLf
    s = s & "<fExpectedQuantity>" & d.fExpectedQuantity & "</fExpectedQuantity>" & vbCrLf
    's = s & "<sStoreDestination=" & Chr(34) & d.sStoreDestination & Chr(34) & ">" & vbCrLf
    s = s & "<sStoreDestination>" & d.sStoreDestination & "</sStoreDestination>" & vbCrLf
    s = s & "<sRouteID/>" & vbCrLf
    s = s & "<sHoldReason/>" & vbCrLf
    s = s & "</ExpectedReceiptLine>" & vbCrLf
    s = s & "</ExpectedReceipt>" & vbCrLf
    s = s & "</ExpectedReceiptMessage>"
    'Call post_oracle_dai_expected_receipt(s)
    Dai_expected_receipt = s
End Function

Sub post_oracle_dai_expected_receipt(rs As String)
    Dim db As ADODB.Connection, ds As Recordset, sqlx As String
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.oradb
    sqlx = "UPDATE hz_cust_acct_sites_all SET ADDRESS_TEXT = '" & rs & "'"
    sqlx = sqlx & " WHERE cust_acct_site_id = 1015"
    'MsgBox sqlx
    db.Execute (sqlx)
    db.Close
End Sub

Sub save_oracle_clob_message(xname As String, recid As Long)
    Dim db As ADODB.Connection, ds As Recordset, sqlx As String
    Dim cfile As String, f0 As String
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.oradb
    sqlx = "update hz_cust_acct_sites_all"
    sqlx = sqlx & " set address_text = ''"
    sqlx = sqlx & " where cust_acct_site_id = " & recid
    db.Execute sqlx
    'cfile = "c:\jvwork\dai" & xname & ".xml"
    cfile = dailogs & "dai" & xname & ".xml"
    Open cfile For Input As #1
    Do Until EOF(1)
        Line Input #1, f0
        sqlx = "update hz_cust_acct_sites_all"
        sqlx = sqlx & " set address_text = address_text || '" & f0 & "'"
        sqlx = sqlx & " where cust_acct_site_id = " & recid
        'MsgBox sqlx
        db.Execute sqlx
    Loop
    Close #1
    db.Close
End Sub

Sub write_oracle_request(xname As String, mssgseq As Long)
    Dim db As ADODB.Connection, ds As Recordset, sqlx As String
    Dim cfile As String, f0 As String, sxml As String
    sxml = ""
    'cfile = "c:\jvwork\dai" & xname & ".xml"
    cfile = dailogs & "dai" & xname & ".xml"
    Open cfile For Input As #1
    Do Until EOF(1)
        Line Input #1, f0
        sxml = sxml & f0
    Loop
    Close #1
    
    Set db = CreateObject("ADODB.Connection")
    db.Open daisqldb
    sqlx = "INSERT INTO HostToWrx (iMessageSequence, sMessageIdentifier) VALUES (" & mssgseq & ", '" & xname & "')"
    'MsgBox sqlx
    db.Execute sqlx
    
    'cfile = "c:\jvwork\dai" & xname & ".xml"
    'Open cfile For Input As #1
    'Do Until EOF(1)
    '    Line Input #1, f0
    '    sqlx = "Update HostToWrx"
    '    sqlx = sqlx & " Set sMessage = sMessage || '" & f0 & "'"
    '    sqlx = sqlx & " Where iMessageSequence = " & mssgseq
    '    'MsgBox sqlx
    '    db.Execute sqlx
    'Loop
    'Close #1
    
    sqlx = "UPDATE HostToWrx SET sMessage = sMessage + '" & sxml & "' WHERE iMessageSequence = " & mssgseq
    'MsgBox sqlx
    db.Execute sqlx
    db.Close
End Sub

Sub read_oracle_clob_message(xname As String, recid As Long)
    Dim db As ADODB.Connection, ds As Recordset, sqlx As String
    Dim cfile As String, f0 As String, clength As Long, i As Long
    clength = 0
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.oradb
    sqlx = "select dbms_lob.getlength(address_text) FROM hz_cust_acct_sites_all"
    sqlx = sqlx & " WHERE cust_acct_site_id = " & recid
    Set ds = db.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        clength = ds(0)
    End If
    ds.Close
    'MsgBox "length=" & clength
    If clength > 0 Then
        'cfile = "c:\jvwork\dai" & xname & ".xml"
        cfile = dailogs & "dai" & xname & ".xml"
        Open cfile For Output As #1
        For i = 1 To clength Step 256
            sqlx = "select dbms_lob.substr(ADDRESS_TEXT, 256, " & i & ")"
            sqlx = sqlx & " FROM hz_cust_acct_sites_all"
            sqlx = sqlx & " WHERE cust_acct_site_id = " & recid
            'MsgBox sqlx
            Set ds = db.Execute(sqlx)
            If ds.BOF = False Then
                ds.MoveFirst
                Print #1, ds(0);
                'MsgBox ds(0)
            End If
            ds.Close
        Next i
        Close #1
    End If
    db.Close
End Sub

Sub read_dai_message(xname As String, seqid As Long)
    Dim db As ADODB.Connection, ds As Recordset, sqlx As String
    Dim cfile As String, f0 As String, clength As Long, i As Long
    clength = 0
    Set db = CreateObject("ADODB.Connection")
    db.Open daisqldb
    sqlx = "SELECT LEN(sMessage) FROM WrxToHost WHERE iMessageSequence = " & seqid
    Set ds = db.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If IsNull(ds(0)) = False Then clength = ds(0)
    End If
    ds.Close
    'MsgBox "length=" & clength
    If clength > 0 Then
        'cfile = "c:\jvwork\dai" & xname & ".xml"
        cfile = dailogs & "dai" & xname & ".xml"
        Open cfile For Output As #1
        For i = 1 To clength Step 256
            sqlx = "SELECT SUBSTRING(sMessage, " & i & ", 256) FROM WrxToHost WHERE iMessageSequence = " & seqid
            'MsgBox sqlx
            Set ds = db.Execute(sqlx)
            If ds.BOF = False Then
                ds.MoveFirst
                Print #1, ds(0);
                'MsgBox ds(0)
            End If
            ds.Close
        Next i
        Close #1
    Else
        'cfile = "c:\jvwork\dai" & xname & ".xml"
        cfile = dailogs & xname & ".xml"
        Open cfile For Output As #1
        Print #1, "<" & xname & ">"
        Print #1, "<Sequence>" & seqid & "</Sequence>"
        Print #1, "<! Zero Length Message -->"
        Print #1, "</" & xname & ">"
        Close #1
    End If
    db.Close
End Sub

Function new_pallet_queue(flag1 As Boolean) As Long
    Dim db As ADODB.Connection, ds As Recordset, s As String, k As Long, zid As Long
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.tbbsr
    If flag1 = True Then            '1st Queue in the list
        s = "select queue_num from queue_infc where queue_num > 0"
        s = s & " order by queue_num"
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            k = ds!queue_num - 1
        Else
            k = 100
        End If
        ds.Close
    Else
        s = "select max(queue_num) from queue_infc"
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            k = ds(0) + 1
        Else
            k = 100
        End If
        ds.Close
    End If
    s = "select id, queue_num from queue_infc where queue_num = 0"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "update queue_infc set queue_num = " & k
        s = s & " where id = " & ds!id
        db.Execute s
        new_pallet_queue = ds!id
    Else
        zid = wd_seq("Queue_Infc")
        s = "INSERT INTO Queue_Infc (ID, Queue_num) VALUES ("
        s = s & zid & ", " & k & ")"
        db.Execute s
        new_pallet_queue = zid
    End If
    ds.Close: db.Close
End Function

Function new_pallet_task_record(parea As String) As Long
    Dim db As ADODB.Connection, ds As Recordset, s As String, zid As Long
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.tbbsr
    s = "select id, status from paltasks where area = '" & parea & "'"
    s = s & " and status = 'COMP'"
    If parea <> "TRAFFIC MASTER" Then s = s & " and id >= 400"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "update paltasks set status = 'PEND'"
        s = s & " where id = " & ds!id
        db.Execute s
        new_pallet_task_record = ds!id
    Else
        ds.Close
        s = "select id, status from paltasks where status = 'COMP'"
        If parea <> "TRAFFIC MASTER" Then s = s & " and id >= 400"
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            s = "update paltasks set status = 'PEND'"
            s = s & " where id = " & ds!id
            db.Execute s
            new_pallet_task_record = ds!id
        Else
            zid = wd_seq("PalTasks")
            s = "INSERT INTO PalTasks (ID) VALUES (" & zid & ")"
            db.Execute s
            new_pallet_task_record = zid
        End If
    End If
    ds.Close: db.Close
End Function

Sub insert_trans(pt As ptask)
    Dim db As ADODB.Connection, ds As Recordset, s As String, rid As Long
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.tbbsr
    s = "select * from paltasks where id = " & new_pallet_task_record(pt.area)
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "update paltasks set area = '" & pt.area & "'"
        s = s & ",description = '" & pt.description & "'"
        s = s & ",source = '" & pt.source & "'"
        s = s & ",target = '" & pt.target & "'"
        s = s & ",product = '" & pt.product & "'"
        s = s & ",palletid = '" & pt.palletid & "'"
        s = s & ",qty = " & CLng(Val(pt.qty))
        s = s & ",uom = '" & pt.uom & "'"
        s = s & ",lotnum = '" & pt.lotnum & "'"
        s = s & ",units = " & CLng(Val(pt.units))
        s = s & ",lotnum2 = '" & pt.lotnum2 & "'"
        s = s & ",units2 = " & CLng(Val(pt.units2))
        s = s & ",status = '" & pt.status & "'"
        s = s & ",userid = '" & pt.userid & "'"
        s = s & ",trandate = '" & pt.trandate & "'"
        s = s & ",reqid = '" & pt.reqid & "'"
        s = s & " where id = " & ds!id
        db.Execute s
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
        db.Execute s
    End If
    ds.Close
    db.Close
End Sub

Sub update_trans(pt As ptask)
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.tbbsr
    s = "select * from paltasks where id = " & pt.id
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "update paltasks set area = '" & pt.area & "'"
        s = s & ",description = '" & pt.description & "'"
        s = s & ",source = '" & pt.source & "'"
        s = s & ",target = '" & pt.target & "'"
        s = s & ",product = '" & pt.product & "'"
        s = s & ",palletid = '" & pt.palletid & "'"
        s = s & ",qty = " & CLng(Val(pt.qty))
        s = s & ",uom = '" & pt.uom & "'"
        s = s & ",lotnum = '" & pt.lotnum & "'"
        s = s & ",units = " & CLng(Val(pt.units))
        s = s & ",lotnum2 = '" & pt.lotnum2 & "'"
        s = s & ",units2 = " & CLng(Val(pt.units2))
        s = s & ",status = '" & pt.status & "'"
        s = s & ",userid = '" & pt.userid & "'"
        s = s & ",trandate = '" & pt.trandate & "'"
        s = s & ",reqid = '" & pt.reqid & "'"
        s = s & " where id = " & ds!id
        db.Execute s
    Else
        MsgBox "update trans failed to find record..." & s
    End If
    ds.Close
    db.Close
End Sub

Function wd_seq(tbname As String) As Long
    Dim db As ADODB.Connection, ds As Recordset, s As String, i As Long
    i = 1
    Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.tbbsr
    db.Open tbbsr
    s = "select sequence_id from sequences where seq = '" & tbname & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        i = ds!sequence_id + 1
        s = "update sequences set sequence_id = " & i
        s = s & " where seq = '" & tbname & "'"
        db.Execute s
    End If
    ds.Close: db.Close
    wd_seq = i
End Function

