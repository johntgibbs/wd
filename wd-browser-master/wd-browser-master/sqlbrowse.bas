Attribute VB_Name = "sqlbrowse"
Public wdb As ADODB.Connection
Public r12db As ADODB.Connection
Public tsb As ADODB.Connection
Public t10bbsr As String
Public k10bbsr As String
Public a10bbsr As String
Public cs5db As String
Public r12access As Boolean
Public r12connection As String
Public wduserid As String

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
Global localAppDataPath As String

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

Function bimp_status_time() As String                   'jv022316
    Dim s As String, cfile As String, msg As String
    s = " "
    cfile = Form1.webdir & "\bimpstat.txt"
    'msg = MsgBox(cfile, vbOKOnly)
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Line Input #1, s
        Close #1
    End If
    bimp_status_time = s
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
            If Len(ds!numwrap) > 0 Then skurec(i).wrapunits = ds!numwrap        'jv082415
            ds.MoveNext
        Loop
    End If
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
    sqlx = "select branch, branchname, gemmsid, modem, fax from branches where gemmsid > ' ' order by branch"
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = ds!branch
            If Len(ds!branch) > 0 Then branchrec(i).branchno = ds!branch
            If Len(ds!branchname) > 0 Then branchrec(i).branchname = ds!branchname
            If Len(ds!gemmsid) > 0 Then branchrec(i).oraloc = ds!gemmsid
            If Len(ds!modem) > 0 Then branchrec(i).capacity = Val(ds!modem)
            If Len(ds!fax) > 0 Then branchrec(i).usable = Val(ds!fax)
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
                                If ds!qty1 > 0 Then gqty = gqty + ds!qty1       'jv081916
                            Else
                                If rs!destination = rbranch Then
                                    If ds!qty1 > 0 Then gqty = gqty + ds!qty1   'jv081916
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
                                If ds!qty2 > 0 Then gqty = gqty + ds!qty2       'jv081916
                            Else
                                If rs!destination = rbranch Then
                                    If ds!qty2 > 0 Then gqty = gqty + ds!qty2   'jv081916
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
                                If ds!qty3 > 0 Then gqty = gqty + ds!qty3       'jv081916
                            Else
                                If rs!destination = rbranch Then
                                    If ds!qty3 > 0 Then gqty = gqty + ds!qty3   'jv081916
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
                                If ds!qty4 > 0 Then gqty = gqty + ds!qty4       'jv081916
                            Else
                                If rs!destination = rbranch Then
                                    If ds!qty4 > 0 Then gqty = gqty + ds!qty4   'jv081916
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

Public Sub vb_elog(eno As Long, edesc As String, pform As String, psub As String, puser As String)
    Dim i As Integer, s As String, cfile As String
    cfile = "s:\wd\html\images\vberrors.txt"
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
    End If
    ds.Close
    wd_seq = i
End Function

