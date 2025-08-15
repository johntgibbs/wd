VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Trailcycle 
   Caption         =   "Trailer Cycle"
   ClientHeight    =   9510
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11175
   LinkTopic       =   "Form3"
   ScaleHeight     =   9510
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   7560
      TabIndex        =   8
      Top             =   7200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   3495
      Left            =   0
      TabIndex        =   4
      Top             =   6000
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6165
      _Version        =   327680
      BackColorFixed  =   12648384
      BackColorSel    =   0
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2415
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4260
      _Version        =   327680
      BackColorFixed  =   8454143
      BackColorSel    =   255
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4260
      _Version        =   327680
      BackColorFixed  =   16777152
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dock Tasks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Top             =   5760
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Crane Tasks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Trailers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Group Codes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu edtrls 
         Caption         =   "Trailers"
         Begin VB.Menu cansku 
            Caption         =   "Cancel Product"
         End
         Begin VB.Menu addsku 
            Caption         =   "Add Product"
            Begin VB.Menu ap 
               Caption         =   "Trailer 1"
               Index           =   0
            End
            Begin VB.Menu ap 
               Caption         =   "Trailer 2"
               Index           =   1
            End
            Begin VB.Menu ap 
               Caption         =   "Trailer 3"
               Index           =   2
            End
            Begin VB.Menu ap 
               Caption         =   "Trailer 4"
               Index           =   3
            End
            Begin VB.Menu ap 
               Caption         =   "Trailer 5"
               Index           =   4
            End
            Begin VB.Menu ap 
               Caption         =   "Trailer 6"
               Index           =   5
            End
            Begin VB.Menu ap 
               Caption         =   "Trailer 7"
               Index           =   6
            End
            Begin VB.Menu ap 
               Caption         =   "Trailer 8"
               Index           =   7
            End
         End
         Begin VB.Menu chgwhs 
            Caption         =   "Change Warehouse"
            Begin VB.Menu cw 
               Caption         =   "Crane 1"
               Index           =   1
            End
            Begin VB.Menu cw 
               Caption         =   "Crane 2"
               Index           =   2
            End
            Begin VB.Menu cw 
               Caption         =   "Crane 3"
               Index           =   3
            End
            Begin VB.Menu cw 
               Caption         =   "Regular"
               Index           =   4
            End
            Begin VB.Menu cw 
               Caption         =   "Ante Room"
               Index           =   11
            End
            Begin VB.Menu snackplant 
               Caption         =   "Snack Plant"
               Index           =   6
            End
            Begin VB.Menu c4way 
               Caption         =   "4 Way"
               Index           =   13
            End
         End
      End
      Begin VB.Menu edship 
         Caption         =   "Shipping"
         Begin VB.Menu chsr 
            Caption         =   "Change Warehouse"
         End
         Begin VB.Menu chsrqty 
            Caption         =   "Change Qty"
         End
         Begin VB.Menu chsrstat 
            Caption         =   "Cancel SKU"
         End
      End
      Begin VB.Menu edtask 
         Caption         =   "Pallet Tasks"
         Begin VB.Menu mc 
            Caption         =   "Mark Complete"
         End
      End
   End
End
Attribute VB_Name = "Trailcycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_groups()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    Combo1.Clear
    s = "select groupcode, count(*) from trailers"
    s = s & " where ra_flag = 'N' and plant = 50"
    s = s & " group by groupcode order by groupcode"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds!groupcode
            ds.MoveNext
        Loop
    Else
        Combo1.AddItem "None"
    End If
    ds.Close
    Combo1.ListIndex = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_groups", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_groups - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_trailers()
    Dim ds As adodb.Recordset, s As String, addt As Boolean, i As Integer
    Dim ss As adodb.Recordset, tname As String, pdesc As String, wdesc As String
    On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 16
    List1.Clear
    s = "select * from trailers where groupcode = '" & Combo1 & "'"
    s = s & " order by sku,branch,account"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            tname = "..."
            If ds!branch <> 16 And ds!branch <> 15 Then
                s = "select branchname from branches where branch = " & ds!branch
                Set ss = Sdb.Execute(s)
                If ss.BOF = False Then
                    ss.MoveFirst
                    tname = StrConv(ss!branchname, vbProperCase) & " " & ds!trlno
                End If
            Else
                s = "select acctdesc from jobbing where branch = " & ds!branch
                s = s & " and account = '" & ds!account & "'"
                Set ss = Sdb.Execute(s)
                If ss.BOF = False Then
                    ss.MoveFirst
                    tname = StrConv(ss!acctdesc, vbProperCase) & " " & ds!trlno
                End If
            End If
            ss.Close
            addt = True
            For i = 0 To List1.ListCount - 1
                If List1.List(i) = tname Then
                    addt = False
                    Exit For
                End If
            Next i
            If addt = True Then List1.AddItem tname
            pdesc = "..."
            s = "select fgunit,fgdesc from skumast where sku = '" & ds!sku & "'"
            Set ss = Sdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                pdesc = StrConv(ss!fgunit & " " & ss!fgdesc, vbProperCase)
            End If
            ss.Close
            wdesc = "..."
            s = "select whsname from warehouses where whs_num = " & ds!whs_num
            Set ss = Sdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                wdesc = ss!whsname
            End If
            ss.Close
            s = ds!id & Chr(9)
            's = s & ds!plant & Chr(9)
            's = s & ds!branch & Chr(9)
            's = s & ds!account & Chr(9)
            s = s & tname & Chr(9)
            s = s & Format(ds!shipdate, "mm-dd-yyyy") & Chr(9)
            's = s & ds!trlno & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & pdesc & Chr(9)
            s = s & ds!pallets & Chr(9)
            s = s & ds!wraps & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!whs_num & Chr(9)
            s = s & wdesc & Chr(9)
            s = s & ds!pb_flag & Chr(9)
            s = s & ds!plant & Chr(9)
            s = s & ds!branch & Chr(9)
            s = s & ds!account & Chr(9)
            s = s & ds!trlno & Chr(9)
            s = s & ds!runid
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    's = "^ID|^Plant|^Branch|^Account|^ShipDate|^Trlno|^SKU|^Pallet|^Wraps|^Units|^Whs|^PB"
    s = "^ID|<Trailer|^ShipDate|^SKU|<Product|^Pallets|^Wraps|^Units|^Whs|^Source|^PB"
    's = "^ID|<Trailer|^ShipDate|^SKU|<Product|^Pallets|^Wraps|^Units|^Whs|^Source|^PB|^Plant|^Branch|^Account|^TrlNo|^RunId"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 1800
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 2800
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 800
    Grid1.ColWidth(11) = 1 '800
    Grid1.ColWidth(12) = 1 '800
    Grid1.ColWidth(13) = 1 '800
    Grid1.ColWidth(14) = 1 '800
    Grid1.ColWidth(15) = 1 '800
    For i = 0 To 7
        'ap(i).Visible = False
        ap(i).Caption = "..."
    Next i
    For i = 0 To List1.ListCount - 1
        If i <= 7 Then
            ap(i).Caption = List1.List(i)
            'ap(i).Visible = True
        End If
    Next i
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_trailers", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_trailers - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_ship_infc()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 9
    s = "select * from ship_infc where order_num = '" & Combo1 & "'"
    s = s & " and ship_status in ('NEW','ACTV')"
    s = s & " order by sku, to_whse_num"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!to_whse_num & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & Format(ds!ship_date, "mm-dd-yyyy") & Chr(9)
            s = s & ds!order_qty & Chr(9)
            s = s & ds!ship_plt_qty & Chr(9)
            s = s & Format(ds!order_qty - ds!ship_plt_qty, "0") & Chr(9)
            s = s & ds!ship_status & Chr(9)
            s = s & ds!gmasize
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "^ID|^SR|^SKU|^ShipDate|^OrderQty|^ShipQty|^Net Qty|^Status|^4Way Size"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 800
    Grid2.ColWidth(1) = 800
    Grid2.ColWidth(2) = 800
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 800
    Grid2.ColWidth(5) = 800
    Grid2.ColWidth(6) = 800
    Grid2.ColWidth(7) = 800
    Grid2.ColWidth(8) = 900
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_ship_infc", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_ship_infc - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_dock_tasks()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 11
    s = "select * from paltasks where description >= '" & Combo1 & "'"
    s = s & " and description < '" & Combo1 & "XXXXX'"
    s = s & " and area in ('DOCK','FORKLIFT')"
    s = s & " and status = 'PEND'"
    s = s & " order by product,area"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!area & Chr(9)
            s = s & ds!source & Chr(9)
            s = s & ds!target & Chr(9)
            s = s & StrConv(ds!product, vbProperCase) & Chr(9)
            s = s & ds!palletid & Chr(9)
            s = s & ds!qty & Chr(9)
            s = s & ds!uom & Chr(9)
            s = s & ds!status & Chr(9)
            s = s & ds!userid & Chr(9)
            s = s & ds!description
            Grid3.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "^ID|^Area|^Source|^Target|<Product|^BarCode|^Qty|^Uom|^Status|^UserID|<"
    Grid3.FormatString = s
    Grid3.ColWidth(0) = 800
    Grid3.ColWidth(1) = 1000
    Grid3.ColWidth(2) = 1000
    Grid3.ColWidth(3) = 1800
    Grid3.ColWidth(4) = 2800
    Grid3.ColWidth(5) = 1800
    Grid3.ColWidth(6) = 800
    Grid3.ColWidth(7) = 800
    Grid3.ColWidth(8) = 800
    Grid3.ColWidth(9) = 800
    Grid3.ColWidth(10) = 1800
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_dock_tasks", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_dock_tasks - Error Number: " & eno
        End
    End If
End Sub

Private Sub ap_Click(Index As Integer)
    Dim ds As adodb.Recordset, s As String
    Dim i As Integer, k As Integer, pdesc As String, p As ptask, pkey As Long
    Dim psku As String, ppal As String, pwhs As String, punits As Integer
    On Error GoTo vberror
    psku = InputBox("SKU:", "Add product......", "777")
    If Len(psku) = 0 Then Exit Sub
    ppal = InputBox("Pallets:", "Pallet Quantity....", "1")
    If Len(ppal) = 0 Then Exit Sub
    s = "select fgunit,fgdesc,pallet from skumast where sku = '" & psku & "'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        pdesc = psku & " " & UCase(ds!fgunit & " " & ds!fgdesc)
        punits = Val(ppal) * ds!pallet
    End If
    ds.Close
    If punits = 0 Then
        Exit Sub
    End If
    pwhs = InputBox("Warehouse:", "Warehouse Code....", "1")
    If Len(pwhs) = 0 Then
        Exit Sub
    End If
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 1) = ap(Index).Caption Then
            Grid1.Row = i
            Exit For
        End If
    Next i
    
    pkey = wd_seq("trailers", Form1.shipdb)
    s = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno, sku"
    s = s & ", pallets, wraps, units, whs_num, pb_flag, ra_flag) Values (" & pkey
    s = s & ", " & Val(Grid1.TextMatrix(i, 15))
    s = s & ", '" & Combo1 & "'"
    s = s & ", " & Grid1.TextMatrix(i, 11)
    s = s & ", " & Grid1.TextMatrix(i, 12)
    s = s & ", '" & Grid1.TextMatrix(i, 13) & "'"
    s = s & ", '" & Grid1.TextMatrix(i, 2) & "'"
    s = s & ", '" & Grid1.TextMatrix(i, 14) & "'"
    s = s & ", '" & psku & "'"
    s = s & ", " & ppal
    s = s & ", 0"
    s = s & ", " & punits
    s = s & ", " & pwhs
    s = s & ", 'N', 'N')"
    Sdb.Execute s
    
    If pwhs > "0" And p_whs < "4" Then
    'If pwhs > "0" And p_whs < "3" Then
        Call insert_ship_infc(Combo1, psku, Val(pwhs), Val(ppal), 0)
    End If
    For k = 1 To Val(ppal)
        p.area = "DOCK"
        p.description = Combo1
        If pwhs = "1" Then p.source = "SR1"
        If pwhs = "2" Then p.source = "SR2"
        If pwhs = "3" Then p.source = "SR3"
        If pwhs = "4" Then p.source = "STAGING"
        If pwhs = "6" Then p.source = "SNACK PLANT"
        If pwhs = "11" Then p.source = "ANTE ROOM"
        If pwhs = "13" Then p.source = "4WAY"
        p.target = ap(Index).Caption
        p.product = pdesc
        p.palletid = "..."
        p.qty = 1
        p.uom = "Pallet"
        p.lotnum = " "
        p.units = 0
        p.lotnum2 = " "
        p.units2 = 0
        p.status = "PEND"
        p.userid = " " ' '
        p.trandate = Format(Now, "yyMMdd hh:mm")
        p.reqid = " "
        Call insert_trans(p)
        If pwhs > "3" Then
            p.area = "FORKLIFT"
            'p.description = Combo1 & " " & ap(Index).Caption & "'": MsgBox s
            p.description = Combo1 & Space(8 - Len(Combo1)) & ap(Index).Caption
            If pwhs = "4" Then p.source = "RACKS"
            If pwhs = "6" Then p.source = "SNACK PLANT"
            If pwhs = "11" Then p.source = "ANTE ROOM"
            If pwhs = "13" Then p.source = "4WAY"
            p.target = "STAGING"
            p.product = pdesc
            p.palletid = "..."
            p.qty = 1
            p.uom = "Pallet"
            p.lotnum = " "
            p.units = 0
            p.lotnum2 = " "
            p.units2 = 0
            p.status = "PEND"
            p.userid = " "
            p.trandate = Format(Now, "yyMMdd hh:mm")
            p.reqid = " "
            Call insert_trans(p)
        End If
    Next k
    refresh_trailers
    DoEvents
    refresh_ship_infc
    DoEvents
    refresh_dock_tasks
    DoEvents
    Grid1_Click
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "ap_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " ap_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub cansku_Click()
    Dim s As String, i As Integer, k As Integer, psku As String, xsku As String
    Dim ds As adodb.Recordset
    On Error GoTo vberror
    If Grid1.Row < 1 Then Exit Sub
    s = "Ok to cancel " & Grid1.TextMatrix(Grid1.Row, 9) & ", "
    s = s & Grid1.TextMatrix(Grid1.Row, 3) & " "
    s = s & Grid1.TextMatrix(Grid1.Row, 4) & ", from "
    s = s & Grid1.TextMatrix(Grid1.Row, 1)
    If MsgBox(s, vbYesNo + vbQuestion, "are you sure.....") = vbNo Then Exit Sub
    k = Grid1.Row
    If Grid2.Rows > 1 Then
        For i = 1 To Grid2.Rows - 1
            If Grid2.TextMatrix(i, 2) = Grid1.TextMatrix(k, 3) And Grid2.TextMatrix(i, 1) = Grid1.TextMatrix(k, 8) Then
                If Val(Grid2.TextMatrix(i, 6)) <= Val(Grid1.TextMatrix(k, 5)) Then
                    s = "update ship_infc set ship_status = 'CANC' where id = "
                    s = s & Grid2.TextMatrix(i, 0)
                    Wdb.Execute s
                Else
                    s = "update ship_infc set order_qty = order_qty - "
                    s = s & Grid1.TextMatrix(i, 5)
                    s = s & " where id = " & Grid2.TextMatrix(i, 0)
                    Wdb.Execute s
                End If
                Exit For
            End If
        Next i
    End If
    If Grid3.Rows > 1 Then
        psku = Grid1.TextMatrix(k, 3)
        xsku = psku & "XXXX"
        For i = 1 To Grid3.Rows - 1
            If Grid3.TextMatrix(i, 4) >= psku And Grid3.TextMatrix(i, 4) < xsku And Grid3.TextMatrix(i, 8) = "PEND" Then
                If Grid3.TextMatrix(i, 1) = "DOCK" Then
                    If Grid3.TextMatrix(i, 3) = Grid1.TextMatrix(k, 1) Then
                        s = "update paltasks set status = 'COMP' where id = "
                        s = s & Grid3.TextMatrix(i, 0)
                        Wdb.Execute s
                    End If
                Else
                    If Grid3.TextMatrix(i, 10) = Combo1 & " " & Grid1.TextMatrix(i, 1) Then
                        s = "update paltasks set status = 'COMP' where id = "
                        s = s & Grid3.TextMatrix(i, 0)
                        Wdb.Execute s
                    End If
                End If
            End If
        Next i
    End If
    s = "Delete from trailers where id = " & Grid1.TextMatrix(k, 0)
    Sdb.Execute s
    Combo1_Click
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "cansku_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " cansku_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub chsr_Click()
    Dim ds As adodb.Recordset, s As String, pwhs As String
    On Error GoTo vberror
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    pwhs = Grid2.TextMatrix(Grid2.Row, 1)
    pwhs = InputBox("Warehouse:", "Change Warehouse....", pwhs)
    If Len(pwhs) = 0 Then Exit Sub
    If pwhs = Grid2.TextMatrix(Grid2.Row, 1) Then Exit Sub
    If pwhs < "1" And pwhs > "3" Then Exit Sub
    s = "select * from ship_infc where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update ship_infc set to_whse_num = " & Val(pwhs) & ", to_vert_loc = 2"
            If pwhs = "1" Then
                s = s & ", to_horz_loc = 18, to_rack_side = 'L'"
            End If
            If pwhs = "2" Then
                s = s & ", to_horz_loc = 22, to_rack_side = 'R'"
            End If
            If pwhs = "3" Then
                s = s & ", to_horz_loc = 43, to_rack_side = 'R'"
            End If
            If pwhs = "4" Then
                s = s & ", to_horz_loc = 0, to_rack_side = 'R'"
            End If
            s = s & " Where id = " & ds!id
            Wdb.Execute s
            ds.MoveNext
            Grid2.TextMatrix(Grid2.Row, 1) = pwhs
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "chsr_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " chsr_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub chsrqty_Click()
    Dim ds As adodb.Recordset, s As String, pqty As String
    On Error GoTo vberror
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    pqty = Grid2.TextMatrix(Grid2.Row, 4)
    pqty = InputBox("Order Qty:", "Change Order Qty....", pqty)
    If Len(pqty) = 0 Then Exit Sub
    If Val(pqty) = 0 Then Exit Sub
    If pqty = Grid2.TextMatrix(Grid2.Row, 4) Then Exit Sub
    Set db = CreateObject("ADODB.Connection")
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update ship_infc set order_qty = " & Val(pqty) & ", status = 'NEW'"
            s = s & " Where id = " & ds!id
            Wdb.Execute s
            ds.MoveNext
            Grid2.TextMatrix(Grid2.Row, 4) = pqty
            Grid2.TextMatrix(Grid2.Row, 7) = "NEW"
            Grid2.TextMatrix(Grid2.Row, 6) = Val(pqty) - Val(Grid2.TextMatrix(Grid2.Row, 5))
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "chsrqty_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " chsrqty_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub chsrstat_Click()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    s = "select * from ship_infc where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update ship_infc set status = 'CANC' Where id = " & ds!id
            Wdb.Execute s
            ds.MoveNext
            Grid2.TextMatrix(Grid2.Row, 7) = "CANC"
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "chsrstat_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " chsrstat_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo1_Click()
    refresh_trailers
    refresh_ship_infc
    refresh_dock_tasks
End Sub

Private Sub cw_Click(Index As Integer)
    Dim s As String, i As Integer, pprod As String, k As Integer
    Dim ds As adodb.Recordset
    On Error GoTo vberror
    k = Grid1.Row
    s = "Update trailers set whs_num = " & Index & " Where id = " & Grid1.TextMatrix(k, 0)
    Sdb.Execute s
    Grid1.TextMatrix(k, 8) = Index
    If Index < 4 Then       'Assign SR1-SR3
        If Val(Grid1.TextMatrix(k, 8)) < 4 Then         'Original SR1-SR3
            For i = 1 To Grid2.Rows - 1
                If Grid2.TextMatrix(i, 2) = Grid1.TextMatrix(k, 3) Then
                    s = "update ship_infc set to_whse_num = " & Index
                    s = s & " where id = " & Grid2.TextMatrix(i, 0)
                    Wdb.Execute s
                End If
            Next i
        Else            'Not SR Originally so try to update existing SR record
            s = " "
            For i = 1 To Grid2.Rows - 1
                If Grid2.TextMatrix(i, 2) = Grid1.TextMatrix(k, 3) And Grid2.TextMatrix(i, 1) = Index Then
                    s = "update ship_infc set order_qty = order_qty + " & Grid1.TextMatrix(k, 5)
                    s = s & " where id = " & Grid2.TextMatrix(i, 0)
                    Wdb.Execute s
                End If
            Next i
            If s < "up" Then        'Not found so insert new SR Record
                Call insert_ship_infc(Combo1, Grid1.TextMatrix(i, 3), Index, Val(Grid1.TextMatrix(i, 5)), 0)
            End If
        End If
    Else                'Assign non-SR warehouse
        If Val(Grid1.TextMatrix(k, 8)) < 4 Then         'If it was originally used in SR, clean up
            For i = 1 To Grid2.Rows - 1
                If Grid2.TextMatrix(i, 2) = Grid1.TextMatrix(k, 3) And Grid2.TextMatrix(i, 1) = Grid1.TextMatrix(k, 8) Then
                    If Val(Grid1.TextMatrix(k, 5)) >= Val(Grid2.TextMatrix(i, 4)) Then
                        s = "update ship_infc set status = 'CANC'"
                    Else
                        s = "update ship_infc set order_qty = order_qty - " & Grid1.TextMatrix(k, 5)
                    End If
                    s = s & " where id = " & Grid2.TextMatrix(i, 0)
                    Wdb.Execute s
                End If
            Next i
        End If
    End If
    pprod = Grid1.TextMatrix(k, 3) & " " & Grid1.TextMatrix(k, 4)
    For i = 1 To Grid3.Rows - 1
        If Grid3.TextMatrix(i, 1) = "DOCK" Then
            If Grid3.TextMatrix(i, 2) <> "ALT" Then
                If Grid3.TextMatrix(i, 4) = pprod Then
                    s = "update paltasks set source = '"
                    If Index = 1 Then s = s & "SR1"
                    If Index = 2 Then s = s & "SR2"
                    If Index = 3 Then s = s & "SR3"
                    If Index = 4 Then s = s & "STAGING"
                    If Index = 6 Then s = s & "SNACK PLANT"
                    If Index = 11 Then s = s & "ANTE ROOM"
                    If Index = 13 Then s = s & "4WAY"
                    s = s & "' where id = " & Grid3.TextMatrix(i, 0)
                    Wdb.Execute s
                End If
            End If
        End If
    Next i
    refresh_ship_infc
    DoEvents
    refresh_dock_tasks
    DoEvents
    Grid1_Click
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "cw_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " cw_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Load()
    refresh_groups
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    Grid2.Width = Me.Width - 80
    Grid3.Width = Me.Width - 80
    Label2.Width = Me.Width - 80
    Label3.Width = Me.Width - 80
    Label4.Width = Me.Width - 80
End Sub

Private Sub Grid1_Click()
    Dim psku As String, i As Integer
    If Grid1.Row < 1 Then Exit Sub
    psku = Grid1.TextMatrix(Grid1.Row, 3)
    If Grid2.Rows > 1 Then
        For i = 1 To Grid2.Rows - 1
            If Grid2.TextMatrix(i, 2) = psku Then
                If Grid2.Rows > 8 Then Grid2.TopRow = i
                Grid2.Row = i: Grid2.Col = 1
                Exit For
            End If
        Next i
    End If
    If Grid3.Rows > 1 Then
        For i = 1 To Grid3.Rows - 1
            If Left(Grid3.TextMatrix(i, 4), 3) = psku Then
                If Grid3.Rows > 8 Then Grid3.TopRow = i
                Grid3.Row = i: Grid3.Col = 4
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edtrls
End Sub

Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edship
End Sub

Private Sub Grid3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edtask
End Sub

Private Sub mc_Click()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    If Val(Grid3.TextMatrix(Grid3.Row, 0)) = 0 Then Exit Sub
    s = "select id,status from paltasks where id = " & Grid3.TextMatrix(Grid3.Row, 0)
    s = s & " and palletid = '" & Grid3.TextMatrix(Grid3.Row, 5) & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update paltasks set status = 'COMP' Where id = " & ds!id
            Wdb.Execute s
            ds.MoveNext
            Grid3.TextMatrix(Grid3.Row, 8) = "COMP"
        Loop
    End If
    ds.Close
    'refresh_dock_tasks
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "mc_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " mc_click - Error Number: " & eno
        End
    End If
End Sub
