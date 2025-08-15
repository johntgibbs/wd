VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form r12wbatpost 
   Caption         =   "R12 WMS Batch Post"
   ClientHeight    =   11985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14430
   LinkTopic       =   "Form2"
   ScaleHeight     =   11985
   ScaleWidth      =   14430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Post WMS Units"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   0
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   6735
      Left            =   0
      TabIndex        =   3
      Top             =   5280
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   11880
      _Version        =   327680
      ForeColor       =   12582912
      BackColorFixed  =   16777152
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4215
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7435
      _Version        =   327680
      ForeColor       =   8388736
      BackColorFixed  =   12648384
      BackColorSel    =   12583104
      FocusRect       =   0
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label bunits 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "bunits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label bproduct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "bproduct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   5040
      Width           =   3615
   End
   Begin VB.Label bbarcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "bbarcode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label batchno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "batchno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WMS Units:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BarCode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Batch Ticket:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Production Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "r12wbatpost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wdb As ADODB.Connection
Dim r12db As ADODB.Connection
Dim t10bbsr As String
Dim k10bbsr As String
Dim a10bbsr As String
Dim cs5db As String
Dim r12access As Boolean
Dim r12connection As String

Function format_bc(bc As String) As String
    Dim s As String
    s = Trim(Mid(bc, 1, 4)) & " "
    s = s & Mid(bc, 5, 6) & "  "
    s = s & Mid(bc, 11, 3) & "  "
    s = s & Mid(bc, 14, 3)
    format_bc = s
End Function

Function barcode_to_lotnum(mbar As String) As String
    Dim s1 As String, s2 As String, s As String, j As Long
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
            s1 = "01-01-20" & Format(CInt(Mid(mbar, 9, 2)) - 2, "00")
            s2 = Mid(mbar, 5, 2) & "-" & Mid(mbar, 7, 2) & "-20" & Format(CInt(Mid(mbar, 9, 2)) - 2, "00")
            j = DateDiff("d", s1, s2) + 1
            s = Format(CInt(Mid(mbar, 9, 2)) - 2, "00")
            s = s & Format(j, "000")
        End If
        barcode_to_lotnum = s
    End If
End Function


Public Sub connect_r12()
    Dim s As String
    's = "This event requires an ODBC connection to the Oracle R12 Database."
    's = s & "  Do you wish to try to connect?"
    'If MsgBox(s, vbYesNo + vbQuestion, "Connect R12....") = vbNo Then Exit Sub
    On Error GoTo r12err
    Set r12db = CreateObject("ADODB.Connection")
    r12db.Open r12connection
    r12access = True
    Exit Sub
r12err:
    MsgBox "R12 Connection failed.", vbOKOnly + vbInformation, "Sorry, no connection...."
End Sub

Function check_post(pid As Long) As Long
    Dim i As Integer, s As String, ds As ADODB.Recordset
    i = 0
    s = "select h.batch_id, d.item_qty from bolinf.mixer_batch_hdr h, bolinf.mixer_batch_dtl d"
    s = s & " where h.batch_id = " & pid
    s = s & " and d.seq_id = h.seq_id"
    'MsgBox s
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = i + ds(1)
            ds.MoveNext
        Loop
    End If
    ds.Close
    check_post = i
End Function

Private Sub post_batches()
    Dim i As Integer, k As Integer, cfile As String, s As String, t As String
    Dim dsn As String, UserId As String, pwd As String
    Dim porg As String, plot As String
    porg = Form1.Combo1
    i = Grid1.Row
    s = "insert into bolinf.mixer_batch_hdr (seq_id,orgn_code,queue_id,batch_id,p_system,formula_id,time_started,time_finished,type,produce_date)"
    s = s & " values (bolinf.mixer_batch_seq.NEXTVAL, "
    s = s & "'" & porg & "', "
    s = s & Grid1.TextMatrix(i, 0) & ", "
    s = s & Grid1.TextMatrix(i, 1) & ", '"
    s = s & Left(Grid1.TextMatrix(i, 3), 10) & "', "
    s = s & Grid1.TextMatrix(i, 8)
    t = Format(Combo1, "DD-MMM-YY") & " 05:00:00 AM"
    s = s & ", TO_DATE('" & t & "', 'DD-MON-YY HH:MI:SS AM'), "
    t = Format(Combo1, "DD-MMM-YY") & " 04:00:00 PM"
    s = s & " TO_DATE('" & t & "', 'DD-MON-YY HH:MI:SS AM')"
    s = s & ", 'FG'"
    t = Format(Combo1, "DD-MMM-YY")
    s = s & ", TO_DATE('" & t & "','DD-MON-YY') )"
    MsgBox s
                
    plot = Format(DateAdd("yyyy", 2, Combo1), "MMddyy")
    plot = plot & Right(Trim(Grid1.TextMatrix(i, 3)), 3)
    s = "insert into bolinf.mixer_batch_dtl (seq_id, item_id, item_qty, item_type, whse_code, loct_code, lot, line)"
    s = s & " values (bolinf.mixer_batch_seq.CURRVAL, "
    s = s & Grid1.TextMatrix(i, 9) & ", "
    s = s & Grid1.TextMatrix(i, 12) & ", 'P', "
    If porg = "500" Then s = s & "'T10', 'FLOORT10', "
    If porg = "501" Then s = s & "'K10', 'FLOORK10', "
    If porg = "502" Then s = s & "'A10', 'FLOORA10', "
    If porg = "503" Then s = s & "'S10', 'FLOORS10', "
    s = s & "'" & plot & "', "
    's = s & pgrid.TextMatrix(i, 11) & ")"
    s = s & "1)"
    MsgBox s
End Sub

Private Sub refresh_grid1()
    Dim q As String, i As Integer
    'Dim dsn As String, userid As String, pwd As String
    Dim db As ADODB.Connection, ds As ADODB.Recordset, s As String, hs As ADODB.Recordset
    Dim t6 As Long, t7 As Long, t8 As Long                      'jv121415
    Dim t9 As Long, t10 As Long, t11 As Long, t13 As Long, t14 As Long, t15 As Long
    Dim k As Long, nl As Boolean, wdlot As String, pflag As String
    Dim psku As String, plot As String, sp As String, ep As String
    
    wdlot = Right(Combo1, 2)
    wdlot = wdlot & Format(DateDiff("d", "1-1-" & Right(Combo1, 4), Combo1) + 1, "000")
    
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    
    'On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 13
    'q = "select h.batch_no,TO_CHAR(h.plan_start_date,'MM-DD-YYYY'),h.batch_status,"
    q = "select h.batch_no,h.batch_id,h.batch_status,"
    q = q & "h.attribute1,i.segment1,i.description,d.plan_qty,"
    q = q & "d.actual_qty,h.formula_id,d.inventory_item_id"
    q = q & " from apps.gme_batch_header h, apps.gme_material_details d, apps.mtl_system_items_b i"
    If Form1.Combo1 = "500" Then
        q = q & " where h.organization_id in (select organization_id from mtl_parameters where organization_code in ('500','503'))"
    Else
        q = q & " where h.organization_id in (select organization_id from mtl_parameters where organization_code in ('" & Form1.Combo1 & "'))"
    End If
    q = q & " and h.plan_start_date >= TO_DATE('" & Format(Combo1, "DD-MMM-YYYY") & "')"
    q = q & " and h.plan_start_date <= TO_DATE('" & Format(DateAdd("d", 1, Combo1), "DD-MMM-YYYY") & "')"
    q = q & " and h.delete_mark = 0"
    q = q & " and h.batch_id = d.batch_id"
    q = q & " and h.batch_status in (1, 2, 3, 4)"
    q = q & " and d.line_type = 1"
    q = q & " and i.organization_id = d.organization_id"
    q = q & " and i.inventory_item_id = d.inventory_item_id"
    q = q & " and i.segment1 >= '100' and i.segment1 <= '9999'"             'jv082415
    q = q & " order by i.segment1, 2, d.plan_qty desc, h.attribute1"
    'MsgBox q
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(0) & Chr(9)                              'Batch No
            s = s & ds(1) & Chr(9)                          'Batch ID
            If ds(2) = 1 Then s = s & "PEND" & Chr(9)       'Status
            If ds(2) = 2 Then s = s & "WIP" & Chr(9)
            If ds(2) = 3 Then s = s & "CERT" & Chr(9)
            If ds(2) = 4 Then s = s & "Closed" & Chr(9)
            s = s & ds(3) & Chr(9)                          'Location
            s = s & ds(4) & Chr(9)                          'SKU
            s = s & ds(5) & Chr(9)                          'Product Name
            s = s & ds(6) & Chr(9)                          'Planned Qty
            s = s & ds(7) & Chr(9)                          'Actual Qty
            s = s & ds(8) & Chr(9)                          'Formula_ID
            s = s & ds(9) & Chr(9)
            pflag = Trim(ds(4))
            If Len(pflag) = 3 Then pflag = pflag & " "
            pflag = pflag & Format(DateAdd("yyyy", 2, Combo1), "MMddyy")
            pflag = pflag & Right(ds(3), 3)
            s = s & Chr(9) & pflag
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    
    If Grid1.Rows > 1 Then
        Grid1_RowColChange
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 0)) > 0 Then
                Grid1.Row = i
                DoEvents
                Grid1.TextMatrix(i, 12) = bunits
                Grid1.TextMatrix(i, 10) = check_post(Grid1.TextMatrix(i, 1))
            End If
        Next i
        Grid1.Row = 1: Grid1.Col = 3
    End If
    
    Screen.MousePointer = 0
    
    If Form1.Combo1 = "502" Then
        Grid1.FormatString = "^Batch No|^Batch ID|^Status|<Location|^SKU|<Description|^Planned|^Actual|^Formula|^Item|^Posted|^Flag|^WMS Qty"
    Else
        Grid1.FormatString = "^Batch No|^Batch ID|^Status|<Location|^SKU|<Description|^Planned|^Actual|^Formula|^Item|^Posted|^BarCode|^WMS Qty"
    End If
    Grid1.ColWidth(0) = 900
    Grid1.ColWidth(1) = 1100
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 2000
    Grid1.ColWidth(4) = 700
    Grid1.ColWidth(5) = 2200
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 800
    Grid1.ColWidth(10) = 800
    Grid1.ColWidth(11) = 1300
    Grid1.ColWidth(12) = 900
    Grid1.FillStyle = flexFillRepeat
    DoEvents
    Grid1.Redraw = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.Description: Err.Clear
    'Call vb_elog(eno, edesc, Me.Name, "refresh_tickets", Form1.UserId)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_tickets - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_grid2()
    Dim db As ADODB.Connection, ds As ADODB.Recordset, s As String, i As Integer, wdlot As String
    Dim cb As ADODB.Connection, cs As ADODB.Recordset, t As String
    Dim aqty As Long, hqty As Long
    Screen.MousePointer = 11
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 9
    bunits = "0"
    wdlot = barcode_to_lotnum(bbarcode.Caption & "EOR")
    wdlot = wdlot & Right(bbarcode, 3)
    Set db = CreateObject("ADODB.Connection")
    If Form1.Combo1 = "500" Then db.Open t10bbsr
    If Form1.Combo1 = "501" Then db.Open k10bbsr
    If Form1.Combo1 = "502" Then db.Open a10bbsr
    
    If Form1.Combo1 = "500" Then          'T10 Cranes
        s = "select l.whse_num, l.zone_num, l.vert_loc, l.horz_loc, l.rack_side, p.posn_num, p.barcode,"
        s = s & " p.lot_num, p.count_qty, p.lot2, p.qty2"
        s = s & " from lane l, position p"
        s = s & " where ((p.barcode >= '" & bbarcode & "'"
        s = s & " and p.barcode <= '" & bbarcode & "EOR') or p.lot2 = '" & wdlot & "')"
        s = s & " and l.id = p.laneno"
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds(0) & Chr(9)
                If ds(0) = 5 Then
                    s = s & ds(1) & " " & ds(2) & "-" & ds(3) & "-" & ds(4) & Chr(9)
                Else
                    s = s & ds(2) & "-" & ds(3) & "-" & ds(4) & " " & ds(5) & Chr(9)
                End If
                s = s & ds(6) & Chr(9)
                s = s & ds(7) & Chr(9)
                s = s & ds(8) & Chr(9)
                s = s & ds(9) & Chr(9)
                s = s & ds(10)
                Grid2.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    
    If Form1.Combo1 = "502" Then
        Set cb = CreateObject("ADODB.Connection")
        cb.Open cs5db
        t = Format(DateAdd("yyyy", 2, Me.Combo1), "M/d/yyyy")
        ' Old version below for reference
        's = "Select location, [Pal ID] from vContainerLocation_1033 where item = '" & Trim(Left(bbarcode, 4))
        's = s & "-" & Right(bbarcode, 3) & "' and expiration = '" & t & "'"
        ' New version by Reece that uses a stored procedure to recreate the view used above on the old (BBSY-01-SQLSVR) server
        s = "EXEC bb_get_pallet_locations '" & Trim(Left(bbarcode, 4))
        s = s & "-" & Right(bbarcode, 3) & "', '" & t & "'"
        'MsgBox s
        Set cs = cb.Execute(s)
        If cs.BOF = False Then
            cs.MoveFirst
            Do Until cs.EOF
                's = "CS5" & Chr(9) & Left(cs(8), 8) & Chr(9) '& Trim(cs(18))
                't = "select barcode, lot1, qty1, lot2, qty2 from pallets where plateno = '" & Trim(cs(18)) & "'"
                s = "CS5" & Chr(9) & Left(cs(0), 8) & Chr(9) '& Trim(cs(18))
                t = "select barcode, lot1, qty1, lot2, qty2 from pallets where barcode = '" & Trim(cs(1)) & "'"
                Set ds = db.Execute(t)
                If ds.BOF = False Then
                    ds.MoveFirst
                    s = s & ds!barcode & Chr(9)
                    s = s & ds!lot1 & Chr(9)
                    s = s & ds!qty1 & Chr(9)
                    s = s & ds!lot2 & Chr(9)
                    s = s & ds!qty2 & Chr(9)
                End If
                ds.Close
                Grid2.AddItem s
                cs.MoveNext
            Loop
        End If
        cs.Close: cb.Close
    End If
    
    s = "select r.aisle, r.rack, p.barcode, p.lot_num, p.count_qty, p.lot2, p.qty2 from racks r, rackpos p"
    s = s & " Where ((p.barcode >= '" & bbarcode & "'"
    s = s & " and p.barcode <= '" & bbarcode & "EOR') or p.lot2 = '" & wdlot & "')"
    s = s & " and r.id = p.rackno"
    'MsgBox s
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "4" & Chr(9)
            s = s & Trim(ds(0)) & "-" & Trim(ds(1)) & Chr(9)
            s = s & ds(2) & Chr(9)
            s = s & ds(3) & Chr(9)
            s = s & ds(4) & Chr(9)
            s = s & ds(5) & Chr(9)
            s = s & ds(6) & Chr(9)
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid2.Rows > 1 Then
        Grid2.FillStyle = flexFillRepeat
        Grid2.Row = 1: Grid2.RowSel = 1
        Grid2.Col = 2: Grid2.ColSel = 2
        Grid2.Sort = 5
        For i = 1 To Grid2.Rows - 1
            If Grid2.TextMatrix(i, 5) > " " Then            '2nd lot
                If Grid2.TextMatrix(i, 5) = wdlot Then
                    s = "select id from holdlist where sku = '" & Trim(Left(bbarcode, 4)) & "'"
                    s = s & " and lot_num = '" & Grid2.TextMatrix(i, 3) & "'"
                    s = s & " and opcode = '" & Mid(Grid2.TextMatrix(i, 2), 11, 3) & "'"
                    s = s & " and spallet <= '" & Right(Grid2.TextMatrix(i, 2), 3) & "'"
                    s = s & " and epallet >= '" & Right(Grid2.TextMatrix(i, 2), 3) & "'"
                    'MsgBox s
                    Set ds = db.Execute(s)
                    If ds.BOF = False Then
                        ds.MoveFirst
                        Grid2.TextMatrix(i, 7) = "Yes"
                        Grid2.TextMatrix(i, 8) = Grid2.TextMatrix(i, 6)
                    End If
                    ds.Close
                Else
                     s = "select id from holdlist where sku = '" & Trim(Left(bbarcode, 4)) & "'"
                     s = s & " and lot_num = '" & Left(Grid2.TextMatrix(i, 5), 5) & "'"
                     s = s & " and opcode = '" & Right(Grid2.TextMatrix(i, 5), 3) & "'"
                     'MsgBox s
                     Set ds = db.Execute(s)
                     If ds.BOF = False Then
                        ds.MoveFirst
                        Grid2.TextMatrix(i, 7) = "Yes"
                        Grid2.TextMatrix(i, 8) = Grid2.TextMatrix(i, 4)
                    End If
                    ds.Close
                End If
            End If
        Next i
        aqty = 0: hqty = 0
        For i = 1 To Grid2.Rows - 1
            If Grid2.TextMatrix(i, 5) > " " Then
                Grid2.Row = i: Grid2.RowSel = i
                If Grid2.TextMatrix(i, 5) = wdlot Then
                    If Grid2.TextMatrix(i, 7) > " " Then
                        hqty = hqty + Val(Grid2.TextMatrix(i, 8))
                    Else
                        aqty = aqty + Val(Grid2.TextMatrix(i, 6))
                    End If
                    Grid2.Col = 5: Grid2.ColSel = 6
                Else
                    If Grid2.TextMatrix(i, 7) > " " Then
                        hqty = hqty + Val(Grid2.TextMatrix(i, 8))
                    Else
                        aqty = aqty + Val(Grid2.TextMatrix(i, 4))
                    End If
                    Grid2.Col = 3: Grid2.ColSel = 4
                End If
                Grid2.CellBackColor = ycolor.BackColor
            Else
                Grid2.Row = i: Grid2.RowSel = i
                Grid2.Col = 3: Grid2.ColSel = 4
                Grid2.CellBackColor = ycolor.BackColor
                aqty = aqty + Val(Grid2.TextMatrix(i, 4))
            End If
            If Len(Grid2.TextMatrix(i, 2)) = 16 Then Grid2.TextMatrix(i, 2) = format_bc(Grid2.TextMatrix(i, 2))
        Next i
        bunits = Format(aqty + hqty, "0")
        s = "All" & Chr(9)
        s = s & " " & Chr(9)
        s = s & "Totals" & Chr(9) & Chr(9) & aqty & Chr(9)
        s = s & Chr(9) & Chr(9) & Chr(9) & hqty
        Grid2.AddItem s
        Grid2.Row = 1
    End If
    db.Close
    Grid2.FormatString = "^SR|^Rack|^BarCode|^Lot|^Units|^Lot2|^Units|^2nd Lot OnHold|^Hold Qty"
    Grid2.ColWidth(0) = 600
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1900
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 1000
    Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 1000
    Grid2.ColWidth(7) = 1600
    Grid2.ColWidth(8) = 1000
    Grid2.Redraw = True
    Screen.MousePointer = 0

End Sub

Private Sub Combo1_Click()
    refresh_grid1
End Sub

Private Sub Command1_Click()
    Dim rt As String, rh As String, rf As String
    refresh_grid1
    rt = Me.Caption
    rh = "Date: " & Combo1
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Grid1.Redraw = False
        Call htmlcolorgrid(Me, "c:\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        Grid1.Redraw = True
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
End Sub

Private Sub Command2_Click()
    post_batches
    DoEvents
    refresh_grid1
End Sub

Private Sub Form_Load()
    a10bbsr = "Driver={SQL Server};Server=bbsy-01-wdsql;DATABASE=SYRacks;UID=bbcwd502;PWD=alabama502"
    k10bbsr = "Driver={SQL Server};Server=bbba-01-wdsql;DATABASE=BARacks;UID=bbcwd501;PWD=barrow501"
    t10bbsr = "Driver={SQL Server};Server=bbc-01-wdsql;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500"
    'cs5db = "Driver={SQL Server};Server=bbsy-01-sqlsvr;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907"
    cs5db = "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=sywms;PWD=!Sylacauga_WMS1907"
    r12connection = "odbc;database=pbelle;uid=apps;pwd=pb3113tx;dsn=pbelle"
    Set wdb = CreateObject("ADODB.Connection")
    wdb.Open "Driver={SQL Server};Server=bbc-01-wdsql;DATABASE=WDShip;UID=bbcship500;PWD=brenham500"
    connect_r12
    Combo1.Clear
    For i = 0 To 21
        Combo1.AddItem Format(DateAdd("d", i * -1, Now), "MM-dd-yyyy")
    Next i
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 200
    Grid2.Width = Me.Width - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wdb.Close
    If r12access = True Then r12db.Close
End Sub

Private Sub Grid1_RowColChange()
    If batchno = Grid1.TextMatrix(Grid1.Row, 0) Then Exit Sub
    Command2.Visible = False
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) > 0 Then
        bunits = " "
        batchno = Grid1.TextMatrix(Grid1.Row, 0)
        bbarcode = Grid1.TextMatrix(Grid1.Row, 11)
        bproduct = Grid1.TextMatrix(Grid1.Row, 5)
        refresh_grid2
        If Val(Grid1.TextMatrix(Grid1.Row, 7)) = 0 And Val(Grid1.TextMatrix(Grid1.Row, 10)) = 0 Then
            If Val(bunits) > 0 Then
                Command2.Visible = True
            End If
        End If
    End If
End Sub
