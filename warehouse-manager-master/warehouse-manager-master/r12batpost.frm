VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form r12batpost 
   Caption         =   "R12 Batch Posting"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form18"
   ScaleHeight     =   8730
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Post"
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   3375
      Left            =   0
      TabIndex        =   6
      Top             =   5280
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5953
      _Version        =   327680
      BackColorFixed  =   12648384
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid oragrid 
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   9360
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   2143
      _Version        =   327680
      BackColorFixed  =   12648384
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1095
      Left            =   0
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   1931
      _Version        =   327680
      BackColorFixed  =   16777152
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read Data"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4695
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8281
      _Version        =   327680
      BackColorFixed  =   12648447
      FocusRect       =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Receiving Date:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "r12batpost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub post_batches()
    Dim i As Integer, k As Integer, cfile As String, s As String, t As String
    Dim dsn As String, userid As String, pwd As String
    If pgrid.Rows < 2 Then Exit Sub
    Screen.MousePointer = 11
    dsn = "pbelle"                            'R12Test
    userid = "Apps"                             'R12Test
    pwd = "papps"                             'R12Test
    If AllocateODBChEnv(hEnv) <> SQL_SUCCESS Then Exit Sub
    If ConnectToDataSource(hEnv, hdbc, hstmt, dsn, userid, pwd) <> SQL_SUCCESS Then
        i = FreeODBChEnv(hEnv)
        Exit Sub
    End If
    'cfile = "s:\wd\test\jvtest.txt"
    'Open cfile For Output As #1
    For i = 1 To pgrid.Rows - 1
        For k = 1 To oragrid.Rows - 1
            If pgrid.TextMatrix(i, 0) = oragrid.TextMatrix(k, 1) Then
                oragrid.TextMatrix(k, 4) = oragrid.TextMatrix(k, 4) & Space(10)     'Pad P-System  jv0213
                s = "insert into bolinf.mixer_batch_hdr (seq_id,orgn_code,queue_id,batch_id,p_system,formula_id,time_started,time_finished,type,produce_date)"
                s = s & " values (bolinf.mixer_batch_seq.NEXTVAL,"
                s = s & "'" & pgrid.TextMatrix(i, 1) & "',"
                s = s & pgrid.TextMatrix(i, 2) & ","
                s = s & pgrid.TextMatrix(i, 0) & ",'"
                s = s & Left(oragrid.TextMatrix(k, 4), 10) & "',"
                s = s & oragrid.TextMatrix(k, 7)
                t = Format(oragrid.TextMatrix(k, 2), "DD-MMM-YY") & " 05:00:00 AM"
                s = s & ",TO_DATE('" & t & "','DD-MON-YY HH:MI:SS AM'),"
                t = Format(Text1, "DD-MMM-YY") & " 04:00:00 PM"
                s = s & "TO_DATE('" & t & "','DD-MON-YY HH:MI:SS AM')"
                s = s & ",'FG'"
                t = Format(oragrid.TextMatrix(k, 2), "DD-MMM-YY")
                s = s & ",TO_DATE('" & t & "','DD-MON-YY') )"
                If Execute_Remote_SQL(s) <> SQL_SUCCESS Then  '9-6
                    MsgBox s, vbOKOnly, "Cannot insert row"     '9-6
                    Exit Sub                                       '9-6
                End If                                             '9-6
                'Print #1, s
                
                s = "insert into bolinf.mixer_batch_dtl (seq_id,item_id,item_qty,item_type,whse_code,loct_code,lot,line)"
                s = s & " values (bolinf.mixer_batch_seq.CURRVAL,"
                s = s & oragrid.TextMatrix(k, 5) & ","
                s = s & pgrid.TextMatrix(i, 10) & ",'P',"
                If pgrid.TextMatrix(i, 1) = "500" Then s = s & "'T10','FLOORT10',"
                If pgrid.TextMatrix(i, 1) = "501" Then s = s & "'K10','FLOORK10',"
                If pgrid.TextMatrix(i, 1) = "502" Then s = s & "'A10','FLOORA10',"
                If pgrid.TextMatrix(i, 1) = "503" Then s = s & "'S10','FLOORS10',"
                s = s & "'" & pgrid.TextMatrix(i, 8) & "',"
                s = s & pgrid.TextMatrix(i, 11) & ")"
                If Execute_Remote_SQL(s) <> SQL_SUCCESS Then  '9-6
                    MsgBox s, vbOKOnly, "Cannot insert row"     '9-6
                    Exit Sub                                       '9-6
                End If                                             '9-6
                'Print #1, s
                Exit For
            End If
        Next k
    Next i
    'Close #1
    i = DisconnectFromDataSource(hdbc, hstmt)
    i = FreeODBChEnv(hEnv)
    Screen.MousePointer = 0
End Sub

Private Sub prtnew()
    Dim rt As String, rf As String, rh As String
    Dim i As Integer, k As Integer, s As String
    Dim spflag As Boolean
    spflag = False
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = Grid1.Cols
    If Form1.plantno = "50" Then
        If MsgBox("Do you want to include Snack Plant?", vbYesNo + vbQuestion, "Organization 503") = vbYes Then
            spflag = True
        End If
    End If
    For i = 1 To Grid1.Rows - 1
        'If Val(Grid1.TextMatrix(i, 10)) <> 0 Then
        If Val(Grid1.TextMatrix(i, 10)) > 0 Then
            s = Grid1.TextMatrix(i, 0)
            For k = 1 To Grid1.Cols - 1
                s = s & Chr(9) & Grid1.TextMatrix(i, k)
            Next k
            If Grid1.TextMatrix(i, 1) = "503" Then
                If spflag = True Then pgrid.AddItem s
            Else
                pgrid.AddItem s
            End If
        End If
    Next i
    pgrid.FormatString = Grid1.FormatString
    For i = 0 To Grid1.Cols - 1
        pgrid.ColWidth(i) = Grid1.ColWidth(i)
    Next i
    
    rt = Me.Caption
    rh = "Date: " & Text1
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, pgrid, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
End Sub

Private Sub refresh_batch_actqtys_r12()
    Dim q As String, i As Integer, k As Integer
    Dim dsn As String, userid As String, pwd As String
    Screen.MousePointer = 11
    dsn = "pbelle"                            'R12Test
    userid = "Apps"                             'R12Test
    pwd = "papps"                             'R12Test

    
    If AllocateODBChEnv(hEnv) <> SQL_SUCCESS Then Exit Sub
    If ConnectToDataSource(hEnv, hdbc, hstmt, dsn, userid, pwd) <> SQL_SUCCESS Then
        i = FreeODBChEnv(hEnv)
        Exit Sub
    End If
    
    q = "select h.batch_no,h.batch_id,TO_CHAR(h.plan_start_date,'MM-DD-YYYY'),h.batch_status,"
    'q = q & "h.attribute1,d.item_id,i.item_no,i.item_desc1,d.plan_qty,"
    q = q & "h.attribute1,d.inventory_item_id,i.segment1,h.formula_id,d.actual_qty,"           'R12Test
    q = q & "d.item_um"
    'q = q & " from apps.gme_batch_header h, apps.gme_material_details d, apps.ic_item_mst_b i"
    q = q & " from apps.gme_batch_header h, apps.gme_material_details d, apps.mtl_system_items_b i" 'R12Test
    q = q & " where h.organization_id in (select organization_id from mtl_parameters where organization_code in ('500', '503', '501', '502'))"
    q = q & " and h.plan_start_date >= TO_DATE('" & Format(DateAdd("d", -7, Text1), "DD-MMM-YYYY") & "')"
    q = q & " and h.plan_start_date <= TO_DATE('" & Format(DateAdd("d", 1, Text1), "DD-MMM-YYYY") & "')"
    q = q & " and h.delete_mark = 0"
    'q = q & " and h.batch_status in (1, 2)"                            'R12Test
    q = q & " and h.batch_id = d.batch_id"
    q = q & " and d.line_type = 1"
    'q = q & " and d.item_id = i.item_id"
    'q = q & " and i.item_no > '000'"
    'q = q & " and i.item_no < '999'"
    q = q & " and i.organization_id = h.organization_id"                'R12Test
    q = q & " and i.inventory_item_id = d.inventory_item_id"            'R12Test
    q = q & " and i.segment1 > '000'"                                   'R12Test
    q = q & " and i.segment1 < '999'"                                   'R12Test
    'q = q & " order by 3, i.item_no, d.plan_qty desc, h.attribute1"
    q = q & " order by 3, i.segment1, d.plan_qty desc, h.attribute1"    'R12Test
    'MsgBox q
    i = LoadGrid(oragrid, q, hstmt, 1, "")
    i = DisconnectFromDataSource(hdbc, hstmt)
    i = FreeODBChEnv(hEnv)
    Screen.MousePointer = 0
    oragrid.FormatString = "^Batch No|Batch_Id|^Plan Start|^Status|<Location|^Item|^SKU|<Formula|^Actual Qty|^UOM"
    oragrid.ColWidth(0) = 800
    oragrid.ColWidth(1) = 800
    oragrid.ColWidth(2) = 1000
    oragrid.ColWidth(3) = 700
    oragrid.ColWidth(4) = 1800
    oragrid.ColWidth(5) = 800
    oragrid.ColWidth(6) = 600
    oragrid.ColWidth(7) = 800
    oragrid.ColWidth(8) = 900
    oragrid.ColWidth(9) = 650
    For i = 1 To oragrid.Rows - 1
        For k = 1 To Grid1.Rows - 1
            If oragrid.TextMatrix(i, 1) = Grid1.TextMatrix(k, 0) Then
                'Grid1.TextMatrix(k, 1) = oragrid.TextMatrix(i, 5)       'Plant
                Grid1.TextMatrix(k, 9) = oragrid.TextMatrix(i, 8)       'Actual Qty
                Exit For
            End If
        Next k
    Next i
End Sub


Private Sub refresh_batch_actqtys()
    Dim q As String, i As Integer, k As Integer
    Dim dsn As String, userid As String, pwd As String
    Screen.MousePointer = 11
    dsn = "pbbcri"
    userid = "bbcgmd"
    pwd = "gmd0207"
    If AllocateODBChEnv(hEnv) <> SQL_SUCCESS Then Exit Sub
    If ConnectToDataSource(hEnv, hdbc, hstmt, dsn, userid, pwd) <> SQL_SUCCESS Then
        i = FreeODBChEnv(hEnv)
        Exit Sub
    End If
    'oragrid.Visible = False
    
    q = "select h.batch_no,h.batch_id,TO_CHAR(h.plan_start_date,'MM-DD-YYYY'),h.batch_status,"
    q = q & "h.attribute1,h.plant_code,i.item_no,i.item_desc1,d.actual_qty,"
    q = q & "d.item_um"
    q = q & " from apps.gme_batch_header h, apps.gme_material_details d, apps.ic_item_mst_b i"
    'q = q & " where h.plant_code = '" & Grid1.TextMatrix(1, 1) & "'"
    q = q & " where h.plant_code in ('500', '503', '501', '502')"
    q = q & " and h.plan_start_date >= TO_DATE('" & Format(DateAdd("d", -7, Text1), "DD-MMM-YYYY") & "')"
    q = q & " and h.plan_start_date <= TO_DATE('" & Format(DateAdd("d", 1, Text1), "DD-MMM-YYYY") & "')"
    q = q & " and h.delete_mark = 0"
    q = q & " and h.batch_id = d.batch_id"
    q = q & " and d.line_type = 1"
    q = q & " and d.item_id = i.item_id"
    q = q & " and i.item_no >= '100'"
    q = q & " and i.item_no <= '999'"
    q = q & " order by h.plant_code desc, 3, i.item_no, d.actual_qty desc, h.attribute1"
    'MsgBox q
    i = LoadGrid(oragrid, q, hstmt, 1, "")
    i = DisconnectFromDataSource(hdbc, hstmt)
    i = FreeODBChEnv(hEnv)
    Screen.MousePointer = 0
    oragrid.FormatString = "^Batch No|Batch_Id|^Plan Start|^Status|<Location|^Plant|^SKU|<Description|^Actual Qty|^UOM"
    oragrid.ColWidth(0) = 800
    oragrid.ColWidth(1) = 800
    oragrid.ColWidth(2) = 1000
    oragrid.ColWidth(3) = 700
    oragrid.ColWidth(4) = 1800
    oragrid.ColWidth(5) = 800
    oragrid.ColWidth(6) = 600
    oragrid.ColWidth(7) = 2000
    oragrid.ColWidth(8) = 900
    oragrid.ColWidth(9) = 650
    For i = 1 To oragrid.Rows - 1
        For k = 1 To Grid1.Rows - 1
            If oragrid.TextMatrix(i, 1) = Grid1.TextMatrix(k, 0) Then
                Grid1.TextMatrix(k, 1) = oragrid.TextMatrix(i, 5)       'Plant
                Grid1.TextMatrix(k, 9) = oragrid.TextMatrix(i, 8)       'Actual Qty
                Exit For
            End If
        Next k
    Next i
End Sub

Private Sub refresh_tickets()
    Dim cfile As String, s As String, i As Integer, bc As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    cfile = "s:\wd\test\r12bats.txt"
    If Form1.plantno = "50" Then cfile = "s:\wd\data\r12bats.txt"
    If Form1.plantno = "51" Then cfile = "\\bbba-03-dc\f\user\waredist\data\r12bats.txt"
    If Form1.plantno = "52" Then cfile = "\\bbsy-02-dc\f\user\waredist\data\r12bats.txt"
    'MsgBox cfile
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7
        s = f0 & Chr(9)
        s = s & f1 & Chr(9)
        s = s & f2 & Chr(9)
        s = s & f3 & Chr(9)
        s = s & f4 & Chr(9)
        s = s & f5 & Chr(9)
        s = s & f6 & Chr(9)
        bc = f3 & " " & f4 & " " & f5 & " 001"
        s = s & barcode_to_lotnum(bc) & Chr(9)
        s = s & f4 & " " & f5 & Chr(9) & Chr(9) & Chr(9) & f7
        
        Grid1.AddItem s
    Loop
    Close #1
    s = "<ID|^Org|<Ticket|^SKU|^CodeDate|^OPCode|^PlanQty|^WDLot|^R12Lot|^Actual|^New|^Line"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 600
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 500
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 900
    Grid1.ColWidth(7) = 700
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 900
    Grid1.ColWidth(10) = 800
    Grid1.ColWidth(11) = 400
End Sub

Private Function pick_ticket(msku As String, mlot As String, mqty As Integer) As Integer
    Dim i As Integer, s As String, t As Long, aqty As Long
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 5
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 3) = msku And Grid1.TextMatrix(i, 8) = mlot Then
            s = i & Chr(9)
            s = s & Grid1.TextMatrix(i, 6) & Chr(9)
            s = s & Grid1.TextMatrix(i, 9) & Chr(9)
            s = s & Grid1.TextMatrix(i, 10) & Chr(9)
            s = s & Val(Grid1.TextMatrix(i, 6)) - (Val(Grid1.TextMatrix(i, 9)) + Val(Grid1.TextMatrix(i, 10)))
            Grid2.AddItem s
        End If
    Next i
    Grid2.FormatString = "^Row|^PlanQty|^Actual|^New|^Diff"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 1000
    't = 0
    'For i = 0 To Grid2.Rows - 1
    '    If Val(Grid2.TextMatrix(i, 3)) > t Then
    '        s = Grid2.TextMatrix(i, 0)
    '        t = Val(Grid2.TextMatrix(i, 3))
    '    End If
    'Next i
    s = Grid2.TextMatrix(1, 0)
    t = 999999
    If mqty > 0 Then
    For i = 1 To Grid2.Rows - 1
        aqty = Val(Grid2.TextMatrix(i, 2)) + Val(Grid2.TextMatrix(i, 3))
        If Val(Grid2.TextMatrix(i, 4)) > mqty And aqty < t Then
            s = Grid2.TextMatrix(i, 0)
            t = aqty
        End If
    Next i
    End If
    pick_ticket = Val(s)
End Function
Private Sub Command1_Click()                    'Read production receive logs
    Dim cfile As String, i As Integer, psku As String, pcode As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim f10 As String, f11 As String, f12 As String, f13 As String, f14 As String
    Dim f15 As String, f16 As String, k As Integer
    'cfile = "v:\pallogs\recv" & Format(Text1, "MMddyyyy") & ".txt"
    cfile = Form1.logdir & "recv" & Format(Text1, "MMddyyyy") & ".txt"
    'MsgBox cfile
    Screen.MousePointer = 11
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0        'ID
        Input #1, f1        'area
        Input #1, f2        'description
        Input #1, f3        'source
        Input #1, f4        'target
        Input #1, f5        'product
        Input #1, f6        'pallet
        Input #1, f7        'qty
        Input #1, f8        'uom
        Input #1, f9        'lot1
        Input #1, f10       'units
        Input #1, f11       'lot2
        Input #1, f12       'units2
        Input #1, f13       'status
        Input #1, f14       'user
        Input #1, f15       'datetime
        Input #1, f16       'reqid
        psku = Trim(Left(f6, 4))
        pcode = Mid(f6, 5, 8)
        For i = 0 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 3) = psku And Grid1.TextMatrix(i, 8) = pcode Then
                k = pick_ticket(psku, pcode, Val(f10))
                Grid1.TextMatrix(k, 10) = Val(Grid1.TextMatrix(k, 10)) + Val(f10)
                Exit For
            End If
        Next i
        If f11 > "0" And f12 > "0" Then
                If Len(f11) = 5 Then                            'jv020614
                    For i = 0 To Grid1.Rows - 1
                        If Grid1.TextMatrix(i, 3) = psku And Grid1.TextMatrix(i, 7) = f11 Then
                            Grid1.TextMatrix(i, 10) = Val(Grid1.TextMatrix(i, 10)) + Val(f12)
                            Exit For
                        End If
                    Next i
                Else
                    For i = 0 To Grid1.Rows - 1
                        If Grid1.TextMatrix(i, 3) = psku And Grid1.TextMatrix(i, 7) = Left(f11, 5) And Grid1.TextMatrix(i, 5) = Right(f11, 1) Then
                            Grid1.TextMatrix(i, 10) = Val(Grid1.TextMatrix(i, 10)) + Val(f12)
                            Exit For
                        End If
                    Next i
                End If
        End If
    Loop
    Close #1
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 10)) > 0 Then
            Grid1.TopRow = i
            Exit For
        End If
    Next i
    'Call prtnew
    'Call post_batches
    Screen.MousePointer = 0
    Command1.Enabled = False
End Sub

Private Sub Command2_Click()                'Print listing
    Call prtnew
    If pgrid.Rows > 2 Then
        Command3.Enabled = True
    End If
End Sub

Private Sub Command3_Click()                'Post batches to R12
    Screen.MousePointer = 11
    Call post_batches
    Screen.MousePointer = 0
    Command3.Enabled = False
    Command1.Enabled = False
End Sub

Private Sub Form_Load()
    Text1 = Format(Now, "mm-dd-yyyy")
    Command3.Enabled = False
    'Text1 = "8-22-2012"
    refresh_tickets
    DoEvents
    'refresh_batch_actqtys
    refresh_batch_actqtys_r12
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    oragrid.Width = Me.Width - 80
    pgrid.Width = Me.Width - 80
End Sub

Private Sub Grid1_Click()
    Dim i As Integer
    i = pick_ticket(Grid1.TextMatrix(Grid1.Row, 3), Grid1.TextMatrix(Grid1.Row, 8), 0)
    'MsgBox i
End Sub
