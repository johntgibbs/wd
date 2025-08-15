Attribute VB_Name = "Trilevel"
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

Function Dai_expected_receipt(d As daiexprct) As String
    Dim s As String, cfile As String
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
    If d.sHoldReason = "PC" Then                                            'jv010616
        's = s & "<sHoldReason>PC</sHoldReason" & vbCrLf                     'jv010616
        s = s & "<sHoldReason>PC</sHoldReason>" & vbCrLf                     'jv072318
    Else                                                                    'jv010616
        s = s & "<sHoldReason/>" & vbCrLf
    End If                                                                  'jv010616
    s = s & "</ExpectedReceiptLine>" & vbCrLf
    s = s & "</ExpectedReceipt>" & vbCrLf
    s = s & "</ExpectedReceiptMessage>"
    'Call post_oracle_dai_expected_receipt(s)
    cfile = Form1.dailogs & "daimessages" & Format(Now, "MMddyy") & ".txt"
    Open cfile For Append As #8
    Print #8, "------"
    Print #8, s
    Print #8, "------"
    Close #8
    Dai_expected_receipt = s
End Function

Sub send_dai_request(pkey As Long, paction As String, pno As String)
    Dim d As daiexprct, rkey As Long, bc As String
    Dim p As ptask                                                          'jv010616
    Dim db As ADODB.Connection, ds As ADODB.Recordset, s As String
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    s = "select * from queue_infc where id = " & pkey
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        d.action = paction
        d.sOrderID = pno
        d.dExpectedDate = Format(Now, "MM/dd/yyyy hh:mm:ss")
        d.sItem = ds!sku
        'd.sLot = Trim(ds!lot_num & Mid(ds!palletid, 12, 1) & Mid(ds!palletid, 14, 3))
        d.sLot = Trim(ds!lot_num & Trim(Mid(ds!palletid, 11, 3)) & Mid(ds!palletid, 14, 3))     'jv052515
        d.fExpectedQuantity = ds!units + ds!units2
        d.sStoreDestination = ds!whse_num
        p.palletid = ds!palletid                                            'jv010616
        p.lotnum = ds!lot_num                                               'jv010616
        p.lotnum2 = ds!lot_num2                                             'jv010616
        If check_hold(p) = True Then                                        'jv010616
            d.sHoldReason = "PC"                                            'jv010616
        Else                                                                'jv010616
            d.sHoldReason = " "                                             'jv010616
        End If                                                              'jv010616
        If d.sStoreDestination = "2" Or d.sStoreDestination = "3" Or d.sStoreDestination = "5" Or d.sStoreDestination = "6" Then
            'Open "c:\jvwork\daiExpectedReceiptMessage.xml" For Output As #1
            Open Form1.dailogs & "daiExpectedReceiptMessage.xml" For Output As #1
            Print #1, Dai_expected_receipt(d)
            Close #1
            DoEvents
            rkey = wd_seq("DAIRequests")
            Call write_oracle_request("ExpectedReceiptMessage", rkey)
        End If
    End If
    ds.Close: db.Close
End Sub

Sub write_oracle_request(xname As String, mssgseq As Long)
    Dim db As ADODB.Connection, sqlx As String
    Dim cfile As String, f0 As String, sxml As String
    sxml = ""
    'cfile = "c:\jvwork\dai" & xname & ".xml"
    cfile = Form1.dailogs & "dai" & xname & ".xml"
    Open cfile For Input As #1
    Do Until EOF(1)
        Line Input #1, f0
        sxml = sxml & f0
    Loop
    Close #1
    Set db = CreateObject("ADODB.Connection")
    db.Open daioradb
    sqlx = "Insert Into HostToWrx (iMessageSequence, sMessageIdentifier, sMessage)"
    sqlx = sqlx & " Values (" & mssgseq & ", '" & xname & "', '" & sxml & "')"
    'MsgBox sqlx
    db.Execute sqlx
    db.Close
End Sub

