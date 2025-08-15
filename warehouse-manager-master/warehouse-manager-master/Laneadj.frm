VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Crane Inventory - Left Side - Segment 1"
   ClientHeight    =   9900
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11175
   LinkTopic       =   "Form2"
   ScaleHeight     =   9900
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1935
      Left            =   0
      TabIndex        =   12
      Top             =   7920
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
      _Version        =   327680
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear Reservation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   120
      Width           =   1875
   End
   Begin VB.Frame Frame2 
      Caption         =   "Position Table "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   4800
      Width           =   10215
      Begin VB.CommandButton Command8 
         Caption         =   "Edit Pallet Units"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Insert Pallet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid PGrid 
         Height          =   4455
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   7858
         _Version        =   327680
         BackColor       =   16777152
         FocusRect       =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Warehouse "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton Command7 
         Caption         =   "Hold"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Block"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   10
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear Lane"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reserve Lane"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid LGrid 
         Height          =   4215
         Left            =   0
         TabIndex        =   3
         Top             =   480
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   7435
         _Version        =   327680
         Cols            =   3
         FixedCols       =   2
         BackColor       =   12648447
         BackColorSel    =   16711680
         FocusRect       =   2
      End
      Begin VB.ComboBox Whs 
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Label rcolor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "rcolor"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9840
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Menu qmenu 
      Caption         =   "&Query"
      Begin VB.Menu qs 
         Caption         =   "Left Side - Segment 1"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu qs 
         Caption         =   "Right Side - Segment 2"
         Index           =   1
      End
      Begin VB.Menu qs 
         Caption         =   "Zone"
         Index           =   2
      End
      Begin VB.Menu qs 
         Caption         =   "Bays For SKU"
         Index           =   3
      End
      Begin VB.Menu qs 
         Caption         =   "Bays For SKU - Lot"
         Index           =   4
      End
      Begin VB.Menu qs 
         Caption         =   "Reserved Bays"
         Index           =   5
      End
      Begin VB.Menu qs 
         Caption         =   "LIFO Bays"
         Index           =   6
      End
      Begin VB.Menu qs 
         Caption         =   "Empty Bays"
         Index           =   7
      End
      Begin VB.Menu qs 
         Caption         =   "Unfilled Bays"
         Index           =   8
      End
      Begin VB.Menu qs 
         Caption         =   "Blocked Bays"
         Index           =   9
      End
      Begin VB.Menu qs 
         Caption         =   "Product On Hold"
         Index           =   10
      End
      Begin VB.Menu qs 
         Caption         =   "Activity Date"
         Index           =   11
      End
      Begin VB.Menu qs 
         Caption         =   "GMA Pallets"
         Index           =   12
      End
      Begin VB.Menu qs 
         Caption         =   "BarCode"
         Index           =   13
      End
   End
   Begin VB.Menu prtmenu 
      Caption         =   "&Print"
      Begin VB.Menu crnsheet 
         Caption         =   "Crane Sheet"
      End
      Begin VB.Menu bbpals 
         Caption         =   "BB Pallets"
      End
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu edlane 
         Caption         =   "Lane"
         Begin VB.Menu edlane1 
            Caption         =   "Reserve Lane"
         End
         Begin VB.Menu edlane2 
            Caption         =   "Clear Lane"
         End
         Begin VB.Menu edlane3 
            Caption         =   "Clear Reservation"
         End
         Begin VB.Menu edlane4 
            Caption         =   "Block Lane"
         End
         Begin VB.Menu edlane5 
            Caption         =   "Tag On Hold"
         End
      End
      Begin VB.Menu edpos 
         Caption         =   "Position"
         Begin VB.Menu edpos1 
            Caption         =   "Clear Position"
         End
         Begin VB.Menu edpos2 
            Caption         =   "Insert Pallet"
         End
         Begin VB.Menu edpos3 
            Caption         =   "Edit Pallet Units"
         End
         Begin VB.Menu palhis 
            Caption         =   "View Pallet History"
         End
         Begin VB.Menu batonhand 
            Caption         =   "View Batch Inventory"
         End
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function calc_date(lotcode As String) As String
    Dim seed As String
    If Left(lotcode, 2) = "00" Then
        seed = "12-31-1999"
    Else
        If Val(lotcode) > 90000 Then
            seed = "12-31-19" & Val(Left(lotcode, 2)) - 1
        Else
            seed = "12-31-20" & Format(Val(Left(lotcode, 2)) - 1, "00")
        End If
    End If
    calc_date = Format(DateAdd("d", Val(Right(lotcode, 3)), seed), "m-d-yyyy")
End Function

Private Sub refresh_lanes()
    Dim ds As ADODB.Recordset, sqlx As String, pstr As String
    Screen.MousePointer = 11
    LGrid.Redraw = False
    LGrid.FontName = "Arial"
    LGrid.FontBold = True
    LGrid.FontSize = 8
    
    LGrid.Rows = 2: LGrid.Cols = 14: LGrid.Clear: LGrid.FixedCols = 5
    sqlx = "Select * From Lane where whse_num = " & Whs
    If qs(0).Checked = True Then sqlx = sqlx & " and rack_side in ('1', 'L')"
    If qs(1).Checked = True Then sqlx = sqlx & " and rack_side in ('2', 'R')"
    If qs(2).Checked = True Then sqlx = sqlx & " and zone_num = " & Right$(Form2.Caption, 2)
    If qs(3).Checked = True Then sqlx = sqlx & " and sku = '" & Trim(Right(Form2.Caption, 4)) & "'" 'jv082415
    If qs(4).Checked = True Then
        pstr = Right$(Form2.Caption, 10)                            'jv082415
        'sqlx = sqlx & " and sku = '" & Left$(pstr, 3) & "'"
        'sqlx = sqlx & " and lot_num = '" & Right$(pstr, 5) & "'"
        sqlx = sqlx & " and id in (select laneno from position where sku = '" & Trim(Left(pstr, 4)) & "'"   'jv082415
        sqlx = sqlx & " and lot_num = '" & Right$(pstr, 5) & "')"
        'MsgBox sqlx
    End If
    If qs(5).Checked = True Then
        sqlx = sqlx & " and (resv_sku > ' ' or resv_lot > ' ')"
    End If
    If qs(6).Checked = True Then sqlx = sqlx & " and lock_status = 1"
    If qs(7).Checked = True Then sqlx = sqlx & " and qty = 0"
    If qs(9).Checked = True Then sqlx = sqlx & " and lane_status = 'B'"
    If qs(10).Checked = True Then sqlx = sqlx & " and lane_status = 'H'"
    If qs(11).Checked = True Then
        sqlx = sqlx & " and id in (select laneno from position where recv_date = #"
        sqlx = sqlx & Right(Form2.Caption, 10) & "#)"
    End If
    If qs(12).Checked Then sqlx = sqlx & " and gmasize > 0"
    If qs(13).Checked = True Then
        sqlx = sqlx & " and id in (select laneno from position where barcode = '" & Right(Me.Caption, 16) & "')"
    End If
    If qs(8).Checked = True Then
        sqlx = sqlx & " and qty > 0 and qty < capacity order by sku, horz_loc, vert_loc"
    Else
        If Whs = 5 Then
            sqlx = sqlx & " Order by zone_num,vert_loc,horz_loc,rack_side"
        Else
            sqlx = sqlx & " Order by vert_loc, horz_loc, rack_side"
        End If
    End If
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!id & Chr$(9)
            sqlx = sqlx & ds!zone_num & Chr$(9)
            sqlx = sqlx & ds!vert_loc & Chr$(9)
            sqlx = sqlx & ds!horz_loc & " " & ds!rack_side & Chr$(9)
            sqlx = sqlx & ds!capacity & Chr$(9)
            If ds!qty > 0 Then
                sqlx = sqlx & ds!qty & Chr$(9)
            Else
                sqlx = sqlx & Chr(9)
            End If
            sqlx = sqlx & ds!lane_status & Chr$(9)
            sqlx = sqlx & ds!sku & Chr$(9)
            sqlx = sqlx & ds!lot_num & Chr$(9)
            sqlx = sqlx & ds!resv_sku & Chr$(9)
            sqlx = sqlx & ds!resv_lot & Chr$(9)
            If ds!lock_status > 0 Then sqlx = sqlx & "*"
            sqlx = sqlx & Chr(9)
            sqlx = sqlx & Format(ds!lot_date, "m-d-yyyy") & Chr(9)
            sqlx = sqlx & Format(ds!gmasize, "#")
            'sqlx = sqlx & ds!lock_status
            LGrid.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    LGrid.FillStyle = flexFillRepeat
    If LGrid.Rows > 1 Then
        For i = 1 To LGrid.Rows - 1
            If LGrid.TextMatrix(i, 6) = "H" Then
                LGrid.Row = i: LGrid.RowSel = i
                LGrid.Col = 2: LGrid.ColSel = LGrid.Cols - 1
                LGrid.CellBackColor = rcolor.BackColor
                LGrid.CellForeColor = rcolor.ForeColor
            End If
        Next i
        'LGrid.Row = 1
    End If
    LGrid.FormatString = "ID|^Zone|^Vert|^Horz|^Cap|^Qty|^Status|^SKU|^Lot|^RSku|^RLot|^Lock|^Lot Date|^GMA Size"
    LGrid.Row = 0
    LGrid.ColWidth(0) = 1: LGrid.ColWidth(1) = 700
    LGrid.ColWidth(2) = 650: LGrid.ColWidth(3) = 650
    LGrid.ColWidth(4) = 650: LGrid.ColWidth(5) = 650
    LGrid.ColWidth(6) = 1000: LGrid.ColWidth(7) = 650
    LGrid.ColWidth(8) = 1000: LGrid.ColWidth(9) = 700
    LGrid.ColWidth(10) = 1000: LGrid.ColWidth(11) = 700
    LGrid.ColWidth(12) = 1200
    LGrid.ColWidth(13) = 1200
    If LGrid.Rows > 2 Then
        LGrid.RemoveItem 1
        LGrid.Row = 1: Call LGrid_Click
    End If
    Screen.MousePointer = 0
    LGrid.Redraw = True
End Sub
Private Sub refresh_pos()
    Dim ds As ADODB.Recordset, sqlx As String, i As Integer
    Dim pdesc As String, psku As String, s As String
    Screen.MousePointer = 11
    pgrid.Redraw = False
    pgrid.FontName = "Arial"
    pgrid.FontBold = True
    pgrid.FontSize = 8
    pgrid.Cols = 15: pgrid.FixedCols = 2
    'pgrid.FormatString = "ID|^Pos|^PosStat|^SKU|^Description|^Lot|^Pallet|^LotStat|^PalStat|^Qty|^Date"
    If Val(LGrid.TextMatrix(LGrid.Row, 13)) > 0 Then
        pgrid.FormatString = "ID|^Pos|^PosStat|^SKU|^Description|^Lot|^Pallet||^PalStat|^Qty|^Lot2|^Qty2|^Date|^GMAPos|^BarCode"
    Else
        pgrid.FormatString = "ID|^Pos|^PosStat|^SKU|^Description|^Lot|^Pallet||^PalStat|^Qty|^Lot2|^Qty2|^Date|^BBCPos|^BarCode"
    End If
    pgrid.ColWidth(0) = 1
    pgrid.ColWidth(1) = 600: pgrid.ColWidth(2) = 850
    pgrid.ColWidth(3) = 700: pgrid.ColWidth(4) = 3500
    pgrid.ColWidth(5) = 1000: pgrid.ColWidth(6) = 1000
    pgrid.ColWidth(7) = 1: pgrid.ColWidth(8) = 800
    pgrid.ColWidth(9) = 800: pgrid.ColWidth(10) = 800
    pgrid.ColWidth(11) = 800: pgrid.ColWidth(12) = 1200
    pgrid.ColWidth(13) = 1000: pgrid.ColWidth(14) = 1600
    pgrid.Rows = 1
    psku = LGrid.TextMatrix(LGrid.Row, 7): pdesc = " "
    sqlx = "Select * From Position where laneno = " & LGrid.TextMatrix(LGrid.Row, 0)
    'If Val(LGrid.TextMatrix(LGrid.Row, 13)) > 0 Then
    '    sqlx = sqlx & " and gmapos is not null"
    'End If
    sqlx = sqlx & " Order by posn_num"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!id & Chr$(9)
            sqlx = sqlx & ds!posn_num & Chr$(9)
            sqlx = sqlx & ds!posn_status & Chr$(9)
            sqlx = sqlx & ds!sku & Chr$(9)
            If Len(ds!sku) > 0 Then
            If Val(ds!sku) > 0 Then
                sqlx = sqlx & skurec(Val(ds!sku)).prodname
            End If
            End If
            sqlx = sqlx & Chr$(9)
            sqlx = sqlx & ds!lot_num & Chr$(9)
            sqlx = sqlx & ds!pallet_num & Chr$(9)
            sqlx = sqlx & ds!lot_status & Chr$(9)
            sqlx = sqlx & ds!pallet_status & Chr$(9)
            sqlx = sqlx & ds!count_qty & Chr$(9)
            sqlx = sqlx & ds!lot2 & Chr(9)
            sqlx = sqlx & ds!qty2 & Chr(9)
            sqlx = sqlx & Format$(ds!recv_date, "m-dd-yyyy") & Chr(9)
            If Val(LGrid.TextMatrix(LGrid.Row, 13)) > 0 Then
                sqlx = sqlx & ds!gmapos
            Else
                sqlx = sqlx & ds!bbcpos
            End If
            sqlx = sqlx & Chr(9) & ds!barcode
            pgrid.AddItem sqlx
            ds.MoveNext
        Loop
        pgrid.Row = 1
    End If
    ds.Close
    If LGrid.TextMatrix(LGrid.Row, 6) = "H" Then            'jv112415
        Command8.Visible = True                             'jv112415
        edpos3.Enabled = True                               'jv112415
    Else                                                    'jv112415
        Command8.Visible = False                            'jv112415
        edpos3.Enabled = False
    End If                                                  'jv112415
    If pgrid.Rows > 1 And LGrid.TextMatrix(LGrid.Row, 6) = "H" Then
        pgrid.FillStyle = flexFillRepeat
        For i = 1 To pgrid.Rows - 1
            If Val(pgrid.TextMatrix(i, 3)) > 0 Then
                pgrid.Row = i: pgrid.RowSel = i
                pgrid.Col = 2: pgrid.ColSel = pgrid.Cols - 1
                pgrid.CellBackColor = rcolor.BackColor
                pgrid.CellForeColor = rcolor.ForeColor
            End If
        Next i
        pgrid.Row = 1
    End If
    pgrid.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub batonhand_Click()
    Dim s As String
    s = Left(pgrid.TextMatrix(pgrid.Row, 14), 13)
    tktonhand.bbarcode = s
    tktonhand.bproduct = pgrid.TextMatrix(pgrid.Row, 4)
    tktonhand.Show
End Sub

Private Sub bbpals_Click()
    If MsgBox("Print all cranes?", vbYesNo + vbQuestion, "all cranes....") = vbYes Then
        Screen.MousePointer = 11
        Form12.qstr = "CraneAll"
    Else
        Screen.MousePointer = 11
        Form12.qstr = "Crane" & Whs
    End If
    Form12.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()        'Reserve Lane
    Dim psku As String, plot As String, pkey As Long, psize As String
    Dim ds As ADODB.Recordset, sqlx As String
    If Val(LGrid.TextMatrix(LGrid.Row, 5)) > 0 Then
        MsgBox "Lane Shows Inventory Qty.", vbOKOnly + vbInformation, "Cannot Reserve"
        Exit Sub
    End If
    If LGrid.TextMatrix(LGrid.Row, 6) = "B" Then
        MsgBox "This lane is blocked.", vbOKOnly + vbInformation, "Cannot Reserve..."
        Exit Sub
    End If
    If LGrid.TextMatrix(LGrid.Row, 6) = "H" Then
        MsgBox "This lane has product on hold.", vbOKOnly + vbInformation, "Cannot Reserve..."
        Exit Sub
    End If
    
    sqlx = "select * from queue_infc where whse_num = " & Whs
    sqlx = sqlx & " and queue_num > 0 order by queue_num"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        psku = ds!sku
        plot = ds!lot_num
    Else
        psku = "000": plot = "00000"
    End If
    ds.Close
    psku = InputBox("SKU #:", "Reserve Lane", psku)
    If Len(psku) = 0 Then Exit Sub
    plot = InputBox("Lot #:", "Reserve Lane", plot)
    If Len(plot) = 0 Then Exit Sub
    psize = InputBox("GMA Size:", "GMA Size", 0)
    If Len(psize) = 0 Then Exit Sub
    
    If skurec(Val(psku)).sku <> psku Then
        MsgBox "Invalid SKU", vbOKOnly, "Sorry Cannot Reserve"
        Exit Sub
    End If
    
    pkey = Val(LGrid.TextMatrix(LGrid.Row, 0))
    
    sqlx = "Update lane set resv_sku = '" & psku & "', resv_lot = '" & plot & "'"
    sqlx = sqlx & ", gmasize = " & CInt(psize) & " Where id = " & pkey
    Wdb.Execute sqlx
    
    sqlx = "Update position set posn_status = 'R' Where laneno = " & pkey
    Wdb.Execute sqlx
    
    LGrid.TextMatrix(LGrid.Row, 9) = psku
    LGrid.TextMatrix(LGrid.Row, 10) = plot
    LGrid.TextMatrix(LGrid.Row, 13) = Format(Val(psize), "#")
    Call LGrid_Click
End Sub

Private Sub Command2_Click()    'Clear Lane
    Dim ds As ADODB.Recordset, sqlx As String, i As Integer
    Dim p As ptask, preas As String                                                 'jv060117
    If MsgBox("Ok to clear this lane?", vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    sqlx = "Update lane set qty = 0, sku = ' ', lot_num = ' ', resv_sku = ' ', resv_lot = ' '"
    sqlx = sqlx & ", lane_status = ' ', lot_date = ' ', gmasize = 0, horz_travel = 0"
    sqlx = sqlx & " where id = " & LGrid.TextMatrix(LGrid.Row, 0)
    Wdb.Execute sqlx
    LGrid.TextMatrix(LGrid.Row, 5) = "0"
    LGrid.TextMatrix(LGrid.Row, 6) = " "
    LGrid.TextMatrix(LGrid.Row, 7) = " "
    LGrid.TextMatrix(LGrid.Row, 8) = " "
    LGrid.TextMatrix(LGrid.Row, 9) = " "
    LGrid.TextMatrix(LGrid.Row, 10) = " "
    LGrid.TextMatrix(LGrid.Row, 12) = ""
    LGrid.TextMatrix(LGrid.Row, 13) = ""
    
    preas = InputBox("Reason for delete:", "Reason for delete....")                         'jv060117
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"                         'jv060117
    Open cfile For Append Shared As #1                                                      'jv060117
    For i = 1 To pgrid.Rows - 1                                                             'jv060117
        sqlx = "Update position set posn_status = ' ', sku = ' ', lot_num = ' ', pallet_num = 0"
        sqlx = sqlx & ", lot_status = ' ', pallet_status = ' ', count_qty = 0"
        sqlx = sqlx & ", recv_date = '" & Format(Now, "m-d-yyyy") & "', barcode = ' '"
        sqlx = sqlx & ", lot2 = ' ', qty2 = 0 Where id = " & pgrid.TextMatrix(i, 0)
        Wdb.Execute sqlx                                                                    'jv060117
        If Val(pgrid.TextMatrix(i, 3)) <> 0 Then                                            'jv060117
            p.area = "SR-" & Whs                                                            'jv060117
            If Len(preas) > 0 Then                                                          'jv060117
                p.description = preas                                                       'jv060117
            Else                                                                            'jv060117
                p.description = " "                                                         'jv060117
            End If                                                                          'jv060117
            p.source = "Clear Lane"                                                         'jv060117
            p.target = Frame2.Caption & " " & Trim(pgrid.TextMatrix(i, 1))                  'jv060117
            p.product = pgrid.TextMatrix(i, 3) & " " & UCase(pgrid.TextMatrix(i, 4))        'jv060117
            p.palletid = pgrid.TextMatrix(i, 14)                                            'jv060117
            p.qty = "1"                                                                     'jv060117
            p.uom = "Pallet"                                                                'jv060117
            p.lotnum = pgrid.TextMatrix(i, 5)                                               'jv060117
            p.units = pgrid.TextMatrix(i, 9)                                                'jv060117
            p.lotnum2 = pgrid.TextMatrix(i, 10)                                             'jv060117
            p.units2 = pgrid.TextMatrix(i, 11)                                              'jv060117
            p.status = "COMP"                                                               'jv060117
            p.userid = Form1.userid                                                         'jv060117
            p.trandate = Format(Now, "yyMMdd hh:mm:ss")                                     'jv060117
            p.reqid = ".."                                                                  'jv060117
            'If LCase(Form1.userid) <> "jvierus" Then                                        'jv060117
                Write #1, i;                                                                'jv060117
                Write #1, p.area;                                                           'jv060117
                Write #1, p.description;                                                    'jv060117
                Write #1, p.source;                                                         'jv060117
                Write #1, p.target;                                                         'jv060117
                Write #1, p.product;                                                        'jv060117
                Write #1, p.palletid;                                                       'jv060117
                Write #1, p.qty;                                                            'jv060117
                Write #1, p.uom;                                                            'jv060117
                Write #1, p.lotnum;                                                         'jv060117
                Write #1, p.units;                                                          'jv060117
                Write #1, p.lotnum2;                                                        'jv060117
                Write #1, p.units2;                                                         'jv060117
                Write #1, p.status;                                                         'jv060117
                Write #1, p.userid;                                                         'jv060117
                Write #1, p.trandate;                                                       'jv060117
                Write #1, p.reqid                                                           'jv060117
            'End If                                                                          'jv060117
        End If                                                                              'jv060117
        pgrid.TextMatrix(i, 2) = " ": pgrid.TextMatrix(i, 3) = " "                          'jv060117
        pgrid.TextMatrix(i, 4) = " ": pgrid.TextMatrix(i, 5) = " "                          'jv060117
        pgrid.TextMatrix(i, 6) = "0": pgrid.TextMatrix(i, 7) = " "                          'jv060117
        pgrid.TextMatrix(i, 8) = " ": pgrid.TextMatrix(i, 9) = "0"                          'jv060117
        pgrid.TextMatrix(i, 10) = " ": pgrid.TextMatrix(i, 11) = " "                        'jv060117
        pgrid.TextMatrix(i, 12) = Format$(Now, "m-dd-yyyy")                                 'jv060117
        pgrid.TextMatrix(i, 14) = " "                                                       'jv060117
    Next i                                                                                  'jv060117
    Close #1                                                                                'jv060117
    
    'SR Log
    'If Form1.plantno = "50" And Whs <= "5" Then         'jv070213
    '    'Add to crane movement log - clear
    '    'cfile = "\\bbc-01-wdmgmt\wd\sr" & Whs.Text & "\bin\SR" & Whs.Text & Format(Now, "mmdd") & ".csv"
    '    cfile = Form1.srserv & "\wd\sr" & Whs.Text & "\bin\SR" & Whs.Text & Format(Now, "mmdd") & ".csv"
    '    'MsgBox cfile
    '    'cfile = "c:\sr10430.csv"
    '    Open cfile For Append As #1
    'End If
    'For i = 1 To pgrid.Rows - 1
    '    sqlx = "Update position set posn_status = ' ', sku = ' ', lot_num = ' ', pallet_num = 0"
    '    sqlx = sqlx & ", lot_status = ' ', pallet_status = ' ', count_qty = 0"
    '    sqlx = sqlx & ", recv_date = '" & Format(Now, "M-d-yyyy") & "', barcode = ' '"
    '    sqlx = sqlx & ", lot2 = ' ', qty2 = 0"
    '    sqlx = sqlx & " where id = " & pgrid.TextMatrix(i, 0)
    '    Wdb.Execute sqlx
    '    If Form1.plantno = "50" And Whs <= "5" And Val(pgrid.TextMatrix(i, 3)) <> 0 Then   'jv070213
    '        Write #1, "SR-" & Whs.Text;
    '        Write #1, "...";
    '        Write #1, pgrid.TextMatrix(i, 3);
    '        Write #1, pgrid.TextMatrix(i, 5);
    '        Write #1, pgrid.TextMatrix(i, 6);
    '        Write #1, LTrim(StrConv(pgrid.TextMatrix(i, 4), vbProperCase));
    '        'Write #1, "WMS";
    '        Write #1, Form1.userid;
    '        Write #1, LTrim(Frame2.Caption) & " " & pgrid.TextMatrix(i, 1);
    '        Write #1, "Cleared";
    '        Write #1, Format(Now, "h:mm am/pm")
    '    End If
    '    pgrid.TextMatrix(i, 2) = " "
    '    pgrid.TextMatrix(i, 3) = " "
    '    pgrid.TextMatrix(i, 4) = " "
    '    pgrid.TextMatrix(i, 5) = " "
    '    pgrid.TextMatrix(i, 6) = " "
    '    pgrid.TextMatrix(i, 7) = " "
    '    pgrid.TextMatrix(i, 8) = " "
    '    pgrid.TextMatrix(i, 9) = "0"
    '    pgrid.TextMatrix(i, 10) = " "
    '    pgrid.TextMatrix(i, 11) = "0"
    '    pgrid.TextMatrix(i, 12) = Format$(Now, "m-dd-yyyy")
    '    pgrid.TextMatrix(i, 14) = " "
    'Next i
    'If Form1.plantno = "50" Then Close #1
End Sub

Private Sub Command3_Click()        'Clear Position
    Dim ds As ADODB.Recordset, y As Integer
    Dim sqlx As String, i As Integer, pqty As Integer
    Dim olot As String, poc As Integer
    Dim p As ptask, preas As String                                                 'jv060117
    If pgrid.Row < 1 Then Exit Sub
    y = pgrid.Row
    If MsgBox("Ok to clear position " & pgrid.Row & "?", vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    
    preas = InputBox("Reason for delete:", "Reason for delete....")                 'jv060117
    p.area = "SR-" & Whs                                                            'jv060117
    If Len(preas) > 0 Then                                                          'jv060117
        p.description = preas                                                       'jv060117
    Else                                                                            'jv060117
        p.description = " "                                                         'jv060117
    End If                                                                          'jv060117
    p.source = "Clear Position"                                                     'jv060117
    p.target = Frame2.Caption & " " & Trim(pgrid.TextMatrix(y, 1))                  'jv060117
    p.product = pgrid.TextMatrix(y, 3) & " " & UCase(pgrid.TextMatrix(y, 4))        'jv060117
    p.palletid = pgrid.TextMatrix(y, 14)                                            'jv060117
    p.qty = "1"                                                                     'jv060117
    p.uom = "Pallet"                                                                'jv060117
    p.lotnum = pgrid.TextMatrix(y, 5)                                               'jv060117
    p.units = pgrid.TextMatrix(y, 9)                                                'jv060117
    p.lotnum2 = pgrid.TextMatrix(y, 10)                                             'jv060117
    p.units2 = pgrid.TextMatrix(y, 11)                                              'jv060117
    p.status = "COMP"                                                               'jv060117
    p.userid = Form1.userid                                                         'jv060117
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")                                     'jv060117
    p.reqid = ".."                                                                  'jv060117
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"                 'jv060117
    'If LCase(Form1.userid) <> "jvierus" Then                                        'jv060117
        Open cfile For Append Shared As #1                                          'jv060117
        Write #1, y;                                                                'jv060117
        Write #1, p.area;                                                           'jv060117
        Write #1, p.description;                                                    'jv060117
        Write #1, p.source;                                                         'jv060117
        Write #1, p.target;                                                         'jv060117
        Write #1, p.product;                                                        'jv060117
        Write #1, p.palletid;                                                       'jv060117
        Write #1, p.qty;                                                            'jv060117
        Write #1, p.uom;                                                            'jv060117
        Write #1, p.lotnum;                                                         'jv060117
        Write #1, p.units;                                                          'jv060117
        Write #1, p.lotnum2;                                                        'jv060117
        Write #1, p.units2;                                                         'jv060117
        Write #1, p.status;                                                         'jv060117
        Write #1, p.userid;                                                         'jv060117
        Write #1, p.trandate;                                                       'jv060117
        Write #1, p.reqid                                                           'jv060117
        Close #1                                                                    'jv060117
    'End If                                                                          'jv060117
    
    
    'If Form1.plantno = "50" And Whs <= "5" Then         'jv070213
    '    'Add to crane movement log - clear
    '    'cfile = "\\bbc-01-wdmgmt\wd\sr" & Whs.Text & "\bin\SR" & Whs.Text & Format(Now, "mmdd") & ".csv"
    '    cfile = Form1.srserv & "\wd\sr" & Whs.Text & "\bin\SR" & Whs.Text & Format(Now, "mmdd") & ".csv"
    '    'MsgBox cfile
    '    'cfile = "c:\sr10430.csv"
    '    Open cfile For Append As #1
    '    Write #1, "SR-" & Whs.Text;
    '    Write #1, "...";
    '    Write #1, pgrid.TextMatrix(y, 3);
    '    Write #1, pgrid.TextMatrix(y, 5);
    '    Write #1, pgrid.TextMatrix(y, 6);
    '    Write #1, LTrim(StrConv(pgrid.TextMatrix(y, 4), vbProperCase));
    '    'Write #1, "WMS";
    '    Write #1, Form1.userid;
    '    Write #1, LTrim(Frame2.Caption) & " " & pgrid.TextMatrix(y, 1);
    '    Write #1, "Cleared";
    '    Write #1, Format(Now, "h:mm am/pm")
    '    Close #1
    'End If
    
    sqlx = "Update position set posn_status = ' ', sku = ' ', lot_num = ' '"
    sqlx = sqlx & ", pallet_num = 0, lot_status = ' ', pallet_status = ' '"
    sqlx = sqlx & ", count_qty = 0 , recv_date = '" & Format$(Now, "m-dd-yyyy") & "'"
    sqlx = sqlx & ", barcode = ' ', lot2 = ' ', qty2 = 0"
    sqlx = sqlx & " Where id = " & pgrid.TextMatrix(y, 0)
    Wdb.Execute sqlx
    pgrid.TextMatrix(y, 2) = " ": pgrid.TextMatrix(y, 3) = " "
    pgrid.TextMatrix(y, 4) = " ": pgrid.TextMatrix(y, 5) = " "
    pgrid.TextMatrix(y, 6) = " ": pgrid.TextMatrix(y, 7) = " "
    pgrid.TextMatrix(y, 8) = " ": pgrid.TextMatrix(y, 9) = "0"
    pgrid.TextMatrix(y, 10) = " ": pgrid.TextMatrix(y, 11) = "0"
    pgrid.TextMatrix(y, 12) = Format$(Now, "m-dd-yyyy")
    pgrid.TextMatrix(y, 14) = " "
    pqty = 0
    olot = "99999"
    poc = 0                                                                     'jv011216
    For i = 1 To pgrid.Rows - 1
        If Val(pgrid.TextMatrix(i, 9)) > 0 Then
            pqty = pqty + 1
            If pgrid.TextMatrix(i, 5) < olot Then olot = pgrid.TextMatrix(i, 5)
            If pgrid.TextMatrix(i, 14) > "0" Then poc = Val(Mid(pgrid.TextMatrix(i, 14), 11, 3))  'jv011216
        End If
    Next i
    
    If pqty = 0 Then
        sqlx = "Update lane set qty = 0, sku = ' ', lot_num = ' ', resv_sku = ' ', resv_lot = ' '"
        sqlx = sqlx & ", lane_status = ' ', lot_date = ' ', horz_travel = 0"
        sqlx = sqlx & " Where id = " & LGrid.TextMatrix(LGrid.Row, 0)
        Wdb.Execute sqlx
        LGrid.TextMatrix(LGrid.Row, 5) = "0"
        LGrid.TextMatrix(LGrid.Row, 6) = " "
        LGrid.TextMatrix(LGrid.Row, 7) = " "
        LGrid.TextMatrix(LGrid.Row, 8) = " "
        LGrid.TextMatrix(LGrid.Row, 9) = " "
        LGrid.TextMatrix(LGrid.Row, 10) = " "
        LGrid.TextMatrix(LGrid.Row, 12) = ""
    Else
        sqlx = "Update lane set qty = " & pqty & ", lot_num = '" & olot & "'"
        sqlx = sqlx & ", lot_date = '" & calc_date(olot) & "', horz_travel = " & poc
        sqlx = sqlx & " Where id = " & LGrid.TextMatrix(LGrid.Row, 0)
        Wdb.Execute sqlx
        LGrid.TextMatrix(LGrid.Row, 5) = pqty
        LGrid.TextMatrix(LGrid.Row, 8) = olot
        LGrid.TextMatrix(LGrid.Row, 12) = Format(calc_date(olot), "m-d-yyyy")
    End If
    
    If y < pgrid.Rows - 1 Then y = y + 1
    pgrid.Col = 1: pgrid.Row = y
    Call PGrid_Click
End Sub

Private Sub Command4_Click()            'Insert Pallet
    Dim psku As String, ppal As String, pdesc As String
    Dim plot As String, pdate As String, sqlx As String
    Dim olot As String, psize As String, pbar As String
    Dim i As Integer, pqty As Integer, lqty As Integer
    Dim pqty2 As Integer, plot2 As String, popl As String
    Dim ds As ADODB.Recordset, y As Integer, recid As Long
    Dim pplate As String                                                              'jv070314
    Dim p As ptask                                                                          'jv060117
    psku = " ": ppal = "0": psize = "0": popl = "_"
    If Val(pgrid.TextMatrix(pgrid.Row, 9)) > 0 Then
        MsgBox "Position currently contains a pallet.", vbOKOnly, "Cannot Insert Here..."
        Exit Sub
    End If
    y = pgrid.Row: lqty = 1
    olot = "99999"
    For i = 1 To pgrid.Rows - 1
        If Val(pgrid.TextMatrix(i, 9)) > 0 Then
            lqty = lqty + 1
            psku = pgrid.TextMatrix(i, 3)
            pdesc = Trim(pgrid.TextMatrix(i, 4))
            plot = pgrid.TextMatrix(i, 5)
            If plot < olot Then olot = plot
            If Val(pgrid.TextMatrix(i, 6)) >= Val(ppal) Then
                ppal = Val(pgrid.TextMatrix(i, 6)) + 1
            End If
            pqty = Val(pgrid.TextMatrix(i, 9))
            pdate = pgrid.TextMatrix(i, 10)
            'popl = Mid(pgrid.TextMatrix(i, 14), 12, 1)
            popl = Mid(pgrid.TextMatrix(i, 14), 11, 3)                  'jv052515
        End If
    Next i
    
    'User Prompts
    psku = InputBox("SKU #", "Insert Position " & y, psku)
    If Len(psku) = 0 Then Exit Sub
    plot = InputBox("Lot #", "Insert Position " & y, plot)
    If Len(plot) = 0 Then Exit Sub
    ppal = InputBox("Pallet #", "Insert Position " & y, ppal)
    If Len(ppal) = 0 Then Exit Sub
    popl = InputBox("Operation Code", "Insert Position " & y, popl)
    If Len(popl) = 0 Or Len(popl) > 3 Then Exit Sub                     'jv052515
    'If Len(popl) = 0 Or Len(popl) > 1 Then Exit Sub
    If Len(popl) = 1 Then                                               'jv052515
        popl = " " & popl & " "                                         'jv052515
    Else                                                                'jv052515
        If Len(popl) = 2 Then popl = " " & popl                         'jv052515
    End If                                                              'jv052515
    psize = InputBox("GMA Size:", "Insert Position " & y, 0)
    If Len(psize) = 0 Then Exit Sub
        
    psku = Trim(Left(psku, 4))                                          'jv082415
    plot = Left$(plot, 5)
    ppal = UCase(ppal)
    popl = UCase(popl)
    i = Val(psku)
    If skurec(i).sku <> psku Then
        MsgBox "Invalid SKU...", vbOKOnly, "Cannot Insert...."
        Exit Sub
    End If
    pdesc = skurec(i).prodname
    If Val(psize) = 0 Then pqty = skurec(i).uom_per_pallet
    pdate = Format$(Now, "m-d-yyyy")
    pqty2 = pqty
    pqty2 = InputBox("Unit Qty for Lot " & plot & ":", "Lot " & plot & " units...", pqty2)
    If Len(pqty2) = 0 Then Exit Sub
    If pqty2 <> pqty Then
        i = pqty - pqty2    '308 - 200 = 108
        pqty = pqty2        '200
        pqty2 = i           '108
        plot2 = Format(Val(plot) + 1, "00000") & popl                                   'jv052515
        plot2 = InputBox("Lot 2#", "Lot #2..." & y, plot2)
        If Len(plot2) = 0 Then Exit Sub
    Else
        plot2 = " "
        pqty2 = 0
    End If
    
    pbar = psku                                                         'jv082415
    If Len(psku) = 3 Then                                               'jv082415
        pbar = pbar & " " & Form1.bb_codedate(plot)                     'jv082415
    Else                                                                'jv082415
        pbar = pbar & Form1.bb_codedate(plot)                           'jv082415
    End If                                                              'jv082415
    pbar = pbar & popl & Format(ppal, "000")                            'jv052515
    sqlx = "Update position set posn_status = ' ', sku = '" & psku & "', lot_num = '" & plot & "'"
    sqlx = sqlx & ", pallet_num = '" & ppal & "', lot_status = '" & pallet_status & "'"             'jv072116
    sqlx = sqlx & ", count_qty = " & pqty & ", recv_date = '" & Format(Now, "M-d-yyyy") & "'"
    sqlx = sqlx & ", barcode = '" & pbar & "', lot2 = '" & plot2 & "', qty2 = " & pqty2
    sqlx = sqlx & " Where id = " & pgrid.TextMatrix(y, 0)
    Wdb.Execute sqlx
    pgrid.TextMatrix(y, 3) = psku
    pgrid.TextMatrix(y, 4) = " " & pdesc
    pgrid.TextMatrix(y, 5) = plot
    pgrid.TextMatrix(y, 6) = ppal
    pgrid.TextMatrix(y, 7) = " "
    pgrid.TextMatrix(y, 8) = " "
    pgrid.TextMatrix(y, 9) = pqty
    pgrid.TextMatrix(y, 10) = plot2
    pgrid.TextMatrix(y, 11) = pqty2
    pgrid.TextMatrix(y, 12) = pdate
    pgrid.TextMatrix(y, 14) = pbar
    If plot < olot Then olot = plot
    
    p.area = "SR-" & Whs                                                            'jv060117
    p.description = " "                                                             'jv060117
    p.source = "Insert Pallet"                                                      'jv060117
    p.target = Frame2.Caption & " " & Trim(pgrid.TextMatrix(y, 1))                  'jv060117
    p.product = pgrid.TextMatrix(y, 3) & " " & UCase(pgrid.TextMatrix(y, 4))        'jv060117
    p.palletid = pgrid.TextMatrix(y, 14)                                            'jv060117
    p.qty = "1"                                                                     'jv060117
    p.uom = "Pallet"                                                                'jv060117
    p.lotnum = pgrid.TextMatrix(y, 5)                                               'jv060117
    p.units = pgrid.TextMatrix(y, 9)                                                'jv060117
    p.lotnum2 = pgrid.TextMatrix(y, 10)                                             'jv060117
    p.units2 = pgrid.TextMatrix(y, 11)                                              'jv060117
    p.status = "COMP"                                                               'jv060117
    p.userid = Form1.userid                                                         'jv060117
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")                                     'jv060117
    p.reqid = ".."                                                                  'jv060117
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"                 'jv060117
    'If LCase(Form1.userid) <> "jvierus" Then                                        'jv060117
        Open cfile For Append Shared As #1                                          'jv060117
        Write #1, y;                                                                'jv060117
        Write #1, p.area;                                                           'jv060117
        Write #1, p.description;                                                    'jv060117
        Write #1, p.source;                                                         'jv060117
        Write #1, p.target;                                                         'jv060117
        Write #1, p.product;                                                        'jv060117
        Write #1, p.palletid;                                                       'jv060117
        Write #1, p.qty;                                                            'jv060117
        Write #1, p.uom;                                                            'jv060117
        Write #1, p.lotnum;                                                         'jv060117
        Write #1, p.units;                                                          'jv060117
        Write #1, p.lotnum2;                                                        'jv060117
        Write #1, p.units2;                                                         'jv060117
        Write #1, p.status;                                                         'jv060117
        Write #1, p.userid;                                                         'jv060117
        Write #1, p.trandate;                                                       'jv060117
        Write #1, p.reqid                                                           'jv060117
        Close #1                                                                    'jv060117
    'End If                                                                          'jv060117
    
    
    'If Form1.plantno = "50" And Whs <= "5" Then         'jv070213
    '    'Add to crane movement log
    '    'cfile = "\\bbc-01-wdmgmt\wd\sr" & Whs.Text & "\bin\SR" & Whs.Text & Format(Now, "mmdd") & ".csv"
    '    cfile = Form1.srserv & "\wd\sr" & Whs.Text & "\bin\SR" & Whs.Text & Format(Now, "mmdd") & ".csv"
    '    'MsgBox cfile
    '    'cfile = "c:\sr10430.csv"
    '    Open cfile For Append As #1
    '    Write #1, "SR-" & Whs.Text;
    '    Write #1, "...";
    '    Write #1, pgrid.TextMatrix(y, 3);
    '    Write #1, pgrid.TextMatrix(y, 5);
    '    Write #1, pgrid.TextMatrix(y, 6);
    '    Write #1, LTrim(StrConv(pgrid.TextMatrix(y, 4), vbProperCase));
    '    'Write #1, "WMS";
    '    Write #1, Form1.userid;
    '    Write #1, "Insert";
    '    Write #1, LTrim(Frame2.Caption) & " " & pgrid.TextMatrix(y, 1);
    '    Write #1, Format(Now, "h:mm am/pm")
    '    Close #1
    'End If
    
    sqlx = "Update lane set qty = " & lqty & ", sku = '" & psku & "', lot_num = '" & olot & "'"
    sqlx = sqlx & ", lot_date = '" & calc_date(olot) & "', gmasize = " & CInt(psize)
    sqlx = sqlx & ", horz_travel = " & Val(popl)
    sqlx = sqlx & " where id = " & LGrid.TextMatrix(LGrid.Row, 0)
    Wdb.Execute sqlx
    'Update pallet record
    If Val(psku) >= 100 Then                                                                'jv070314
        recid = 0
        s = "select * from pallets where barcode = '" & pbar & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            pplate = ds!plateno                                                             'jv070314
            recid = ds!id
        Else
            ds.Close
            s = "select * from pallets where status in ('Shipped','Order Pick')"
            s = s & " order by trandate"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                pplate = " "                                                                'jv070314
                recid = ds!id
            End If
        End If
        ds.Close
        If recid > 0 Then
            s = "Update pallets set plateno = '" & pplate & "'"                             'jv070314
            s = s & ",barcode = '" & pbar & "'"
            s = s & ",qty1 = " & Val(pqty)
            s = s & ",lot1 = '" & plot & "'"
            s = s & ",qty2 = " & Val(pqty2)
            s = s & ",lot2 = '" & plot2 & "'"
            s = s & ",source = 'SR-" & Whs & "'"
            If Whs = "5" Then
                s = s & ",target = '" & LGrid.TextMatrix(LGrid.Row, 1) & " " & Frame2.Caption & "'"
            Else
                s = s & ",target = '" & Frame2.Caption & "'"
            End If
            If psize = 0 Then
                s = s & ",bbc = 'Y'"
            Else
                s = s & ",bbc = 'N'"
            End If
            s = s & ",status = 'Warehouse'"
            s = s & ",trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
            s = s & ",sku = '" & psku & "'"
            s = s & " Where id = " & recid
            Wdb.Execute s
        Else
            pid = wd_seq("Pallets")
            s = "Insert Into pallets Values (" & pid
            's = s & ",'" & recid & "'"
            s = s & ",'" & pplate & "'"                                                     'jv070314
            s = s & ",'" & pbar & "'"
            s = s & "," & Val(pqty)
            s = s & ",'" & plot & "'"
            s = s & "," & Val(pqty2)
            s = s & ",'" & plot2 & "'"
            s = s & ",'SR-" & Whs & "'"
            If Whs = "5" Then
                s = s & ",'" & LGrid.TextMatrix(LGrid.Row, 1) & " " & Frame2.Caption & "'"
            Else
                s = s & ",'" & Frame2.Caption & "'"
            End If
            If psize = 0 Then
                s = s & ",'Y'"
            Else
                s = s & ",'N'"
            End If
            s = s & ",'Warehouse'"
            s = s & ",'" & Format(Now, "yyMMdd hh:mm:ss") & "'"
            s = s & ",'" & psku & "')"
            Wdb.Execute s
        End If
    End If                                                                                      'jv070314
    LGrid.TextMatrix(LGrid.Row, 5) = lqty
    LGrid.TextMatrix(LGrid.Row, 7) = psku
    LGrid.TextMatrix(LGrid.Row, 8) = olot
    LGrid.TextMatrix(LGrid.Row, 12) = Format(calc_date(olot), "m-d-yyyy")
    If Val(psize) > 0 Then LGrid.TextMatrix(LGrid.Row, 13) = Format(CInt(psize), "#")
    If y > 1 Then y = y - 1
    pgrid.Col = 1: pgrid.Row = y
    Call PGrid_Click
End Sub

Private Sub Command5_Click()        'Clear Reservation
    Dim psku As String, plot As String, pkey As Long
    Dim ds As ADODB.Recordset, sqlx As String
    pkey = Val(LGrid.TextMatrix(LGrid.Row, 0))
    sqlx = "Update lane set resv_sku = ' ', resv_lot = ' ' Where id = " & pkey
    Wdb.Execute sqlx
    sqlx = "Update position set posn_status = ' ' Where laneno = " & pkey
    Wdb.Execute sqlx
    LGrid.TextMatrix(LGrid.Row, 9) = ""
    LGrid.TextMatrix(LGrid.Row, 10) = ""
    Call LGrid_Click
End Sub

Private Sub Command6_Click()
    Dim sqlx As String
    If Val(LGrid.TextMatrix(LGrid.Row, 0)) = 0 Then Exit Sub
    If LGrid.TextMatrix(LGrid.Row, 6) = "B" Then
        sqlx = "Update lane set lane_status = ' ' Where id = " & LGrid.TextMatrix(LGrid.Row, 0)
        LGrid.TextMatrix(LGrid.Row, 6) = " "
    Else
        sqlx = "Update lane set lane_status = 'B' Where id = " & LGrid.TextMatrix(LGrid.Row, 0)
        LGrid.TextMatrix(LGrid.Row, 6) = "B"
    End If
    Wdb.Execute sqlx
End Sub

Private Sub Command7_Click()
    Dim ds As ADODB.Recordset, sqlx As String
    Dim psku As String, plot As String, pbar As String, pcode As String, zid As Long
    Dim i As Integer
    If Val(LGrid.TextMatrix(LGrid.Row, 0)) = 0 Then Exit Sub
    If LGrid.TextMatrix(LGrid.Row, 6) = "H" Then
        sqlx = "Update lane set lane_status = ' ' Where id = " & LGrid.TextMatrix(LGrid.Row, 0)
        LGrid.TextMatrix(LGrid.Row, 6) = " "
    Else
        sqlx = "Update lane set lane_status = 'H' Where id = " & LGrid.TextMatrix(LGrid.Row, 0)
        LGrid.TextMatrix(LGrid.Row, 6) = "H"
    End If
    Wdb.Execute sqlx
    For i = 1 To pgrid.Rows - 1
        If pgrid.TextMatrix(i, 3) > "0" Then
            zid = wd_seq("HoldList")                                            'jv042015
            psku = pgrid.TextMatrix(i, 3)                                                       'jv040715
            plot = pgrid.TextMatrix(i, 5)                                                       'jv040715
            pcode = Mid(pgrid.TextMatrix(i, 14), 12, 1)                                         'jv040715
            ppal = Mid(pgrid.TextMatrix(i, 14), 14, 3)                                          'jv040715
            s = "select id from pallets where barcode = '" & pgrid.TextMatrix(i, 14) & "'"      'jv040715
            Set ds = Wdb.Execute(s)                                                             'jv040715
            If ds.BOF = False Then                                                              'jv040715
                ds.MoveFirst                                                                    'jv040715
                zid = ds!id                                                                     'jv040715
            End If                                                                              'jv040715
            ds.Close                                                                            'jv040715
            If LGrid.TextMatrix(LGrid.Row, 6) = "H" Then
                s = "Insert into holdlist (id, sku, lot_num, opcode, spallet, epallet, hsource, userid, holddate) values (" & zid  'jv040715
                s = s & ", '" & psku & "', '" & plot & "', '" & pcode & "', '" & ppal & "', '" & ppal & "', 'SR-" & Whs & "'"
                s = s & ", '" & WDUserId & "', '" & Format(Now, "yyMMdd hh:mm:ss") & "')"
                Wdb.Execute s                                                       'jv040715
            Else                                                                    'jv040715
                s = "delete from holdlist where sku = '" & psku & "'"               'jv040715
                s = s & " and lot_num = '" & plot & "'"                             'jv040715
                s = s & " and opcode = '" & pcode & "'"                             'jv040715
                s = s & " and spallet = '" & ppal & "'"                             'jv040715
                s = s & " and epallet = '" & ppal & "'"                             'jv040715
                Wdb.Execute s                                                       'jv040715
            End If
        End If
    Next i
End Sub

Private Sub Command8_Click()
    Dim mlot1 As String, mlot2 As String, mqty1 As String, mqty2 As String, p As ptask
    Dim s As String, preas As String, cfile As String
    If Val(pgrid.TextMatrix(pgrid.Row, 9)) = 0 Then Exit Sub
    If LGrid.TextMatrix(LGrid.Row, 6) <> "H" Then
        MsgBox "Edit is available for on-hold units.", vbOKOnly + vbInformation, "Lane is not on hold.."
        Exit Sub
    End If
    mlot1 = pgrid.TextMatrix(pgrid.Row, 5)
    mqty1 = Val(pgrid.TextMatrix(pgrid.Row, 9))
    mlot2 = pgrid.TextMatrix(pgrid.Row, 10)
    mqty2 = Val(pgrid.TextMatrix(pgrid.Row, 11))
    If Val(mqty1) > 0 Then
        mqty1 = InputBox("Units:", "Lot " & mlot1 & " units.", mqty1)
        If Len(mqty1) = 0 Then Exit Sub
        If Val(mqty1) <= 0 Then
            MsgBox "Quantity: " & mqty1 & " is invalid.", vbOKOnly + vbInformation, "Sorry, try again.."
            Exit Sub
        End If
    End If
    If Val(mqty2) > 0 Then
        mqty2 = InputBox("Units:", "Lot " & mlot2 & " units.", mqty2)
        If Len(mqty2) = 0 Then Exit Sub
        If Val(mqty2) <= 0 Then
            MsgBox "Quantity: " & mqty2 & " is invalid.", vbOKOnly + vbInformation, "Sorry, try again.."
            Exit Sub
        End If
    End If
    
    s = "Update position set count_qty = " & mqty1 & ", qty2 = " & mqty2
    s = s & " Where id = " & pgrid.TextMatrix(pgrid.Row, 0)
    Wdb.Execute s
    s = "Update pallets set qty1 = " & mqty1 & ", qty2 = " & mqty2
    s = s & " Where barcode = '" & pgrid.TextMatrix(pgrid.Row, 14) & "'"
    Wdb.Execute s
    
    preas = InputBox("Reason for edit:", "Reason for edit....")
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"
    'cfile = "v:\testlogs\wms" & Format(Now, "mmddyyyy") & ".txt"
    Open cfile For Append Shared As #1
    p.area = "WMS"
    If Len(preas) > 0 Then
        p.description = preas
    Else
        p.description = " "
    End If
    p.source = "Edit"
    If Whs = "5" Then
        p.target = LGrid.TextMatrix(LGrid.Row, 1) & ":" & Frame2.Caption
    Else
        p.target = Whs & ":" & Frame2.Caption
    End If
    p.product = pgrid.TextMatrix(pgrid.Row, 3) & " " & pgrid.TextMatrix(pgrid.Row, 4)
    p.palletid = pgrid.TextMatrix(pgrid.Row, 14)
    p.qty = "1"
    p.uom = "Pallet"
    p.lotnum = mlot1
    p.units = mqty1
    p.lotnum2 = mlot2
    p.units2 = mqty2
    p.status = "COMP"
    p.userid = Form1.userid
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")
    p.reqid = ".."
    Write #1, "0";
    Write #1, p.area;
    Write #1, p.description;
    Write #1, p.source;
    Write #1, p.target;
    Write #1, p.product;
    Write #1, p.palletid;
    Write #1, p.qty;
    Write #1, p.uom;
    Write #1, p.lotnum;
    Write #1, p.units;
    Write #1, p.lotnum2;
    Write #1, p.units2;
    Write #1, p.status;
    Write #1, p.userid;
    Write #1, p.trandate;
    Write #1, p.reqid
    Close #1
    refresh_pos
End Sub

Private Sub crnsheet_Click()
    Dim ds As ADODB.Recordset, s As String
    Dim rt As String, rh As String, rf As String, i As Double
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
        
    s = "select sku,vert_loc,horz_loc,rack_side,zone_num,qty,lot_num,lane_status"
    s = s & " from lane"
    s = s & " where whse_num = " & Whs
    s = s & " order by horz_loc, vert_loc, rack_side"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!vert_loc & " " & ds!horz_loc & " " & ds!rack_side & Chr(9)
            s = s & ds!zone_num & Chr(9)
            s = s & ds!qty & Chr(9)
            If ds!sku > "0" Then s = s & ds!sku
            s = s & Chr(9)
            If ds!lot_num > "0" Then s = s & ds!lot_num
            s = s & Chr(9)
            If ds!lane_status = "B" Then
                s = s & "Blocked" & Chr(9)
            Else
                If ds!lane_status = "H" Then
                    s = s & "OnHold" & Chr(9)
                Else
                    s = s & Chr(9)
                End If
            End If
            If Val(ds!sku) > 0 Then
                s = s & skurec(Val(ds!sku)).prodname
            End If
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    Grid1.FormatString = "^Bay|^Zone|^Qty|^SKU|^Lot #|^Status|<Product"
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 3000
    Screen.MousePointer = 0
    
    rt = "Crane " & Whs & " Count Sheet"
    rh = "SR-" & Whs & "   " & Format(Now, "mmmm d, yyyy")
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
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

Private Sub edlane1_Click()
    Command1_Click
End Sub

Private Sub edlane2_Click()
    Command2_Click
End Sub

Private Sub edlane3_Click()
    Command5_Click
End Sub

Private Sub edlane4_Click()
    Command6_Click
End Sub

Private Sub edlane5_Click()
    Command7_Click
End Sub

Private Sub edpos1_Click()
    Command3_Click
End Sub

Private Sub edpos2_Click()
    Command4_Click
End Sub

Private Sub edpos3_Click()
    Command8_Click
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Form2.WindowState = 0 Then
        For i = 1 To Form1.Frmgrid.Rows - 1
            If Form1.Frmgrid.TextMatrix(i, 0) = "form2" Then
                Form1.Frmgrid.TextMatrix(i, 1) = Form2.Top
                Form1.Frmgrid.TextMatrix(i, 2) = Form2.Left
                Form1.Frmgrid.TextMatrix(i, 3) = Form2.Height
                Form1.Frmgrid.TextMatrix(i, 4) = Form2.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.Frmgrid.Rows - 1
        If Form1.Frmgrid.TextMatrix(i, 0) = "form2" Then
            Form2.Top = Val(Form1.Frmgrid.TextMatrix(i, 1))
            Form2.Left = Val(Form1.Frmgrid.TextMatrix(i, 2))
            Form2.Height = Val(Form1.Frmgrid.TextMatrix(i, 3))
            Form2.Width = Val(Form1.Frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Whs.AddItem "1"
    Whs.AddItem "2"
    Whs.AddItem "3"
    Whs.AddItem "5"
    Whs.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Frame1.Width = Me.Width - 80
    Frame2.Width = Me.Width - 80
    LGrid.Width = Frame1.Width
    pgrid.Width = Frame2.Width
    Grid1.Width = Me.Width - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub LGrid_Click()
    'rcolor.Caption = "clicked"
    Call refresh_pos
    Frame2.Caption = LGrid.TextMatrix(LGrid.Row, 2) & " "
    Frame2.Caption = Frame2.Caption & LGrid.TextMatrix(LGrid.Row, 3)
    LGrid.Col = 4: LGrid.ColSel = LGrid.Cols - 1
    'rcolor.Caption = "rcolor"
End Sub

Private Sub LGrid_EnterCell()
    'If rcolor.Caption = "rcolor" Then Call LGrid_Click
End Sub

Private Sub LGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edlane
End Sub

Private Sub palhis_Click()
    palhistory.Show
    palhistory.barkey = pgrid.TextMatrix(pgrid.Row, 14)
End Sub

Private Sub PGrid_Click()
    pgrid.ColSel = pgrid.Cols - 1
End Sub

Private Sub PGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call PGrid_Click
End Sub

Private Sub PGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edpos
End Sub

Private Sub qs_Click(Index As Integer)
    Dim i As Integer, pzone As Integer, psku As String, plot As String, pbc As String
    Dim pdate As String
    For i = 0 To 12
        qs(i).Checked = False
    Next i
    qs(Index).Checked = True
    Form2.Caption = "Crane Inventory " & qs(Index).Caption
    If Index = 2 Then
        pzone = InputBox("Zone #", "Enter a zone..", "1")
        Form2.Caption = Form2.Caption & " " & Format$(pzone, "00")
    End If
    If Index = 3 Or Index = 4 Then
        psku = LGrid.TextMatrix(LGrid.Row, 7)
        psku = InputBox("SKU #", "Enter SKU #...", psku)
        If Len(psku) = 0 Then psku = "000"
        Form2.Caption = Form2.Caption & " " & Format$(psku, "000")
    End If
    If Index = 4 Then
        plot = LGrid.TextMatrix(LGrid.Row, 8)
        plot = InputBox("Lot #", "Enter Lot #...", plot)
        If Len(plot) = 0 Then plot = "98000"
        Form2.Caption = Form2.Caption & " " & Format$(plot, "00000")
    End If
    If Index = 11 Then
        pdate = InputBox("Date:", "Enter Date ...", Format(Now, "m-d-yyyy"))
        If Len(pdate) = 0 Then Exit Sub
        If IsDate(pdate) = False Then pdate = Format(Now, "m-d-yyyy")
        Form2.Caption = Form2.Caption & " " & Format(pdate, "mm-dd-yyyy")
    End If
    If Index = 13 Then                                                                  'jv112515
        pbc = InputBox("BarCode:", "Enter Pallet Barcode...", "777 112515570001")
        If Len(pbc) = 0 Then Exit Sub
        If Len(pbc) <> 16 Then
            MsgBox "Invalid barcode: " & pbc, vbOKOnly + vbInformation, "sorry, try again..."
            Exit Sub
        End If
        Form2.Caption = "BarCode: " & pbc
    End If
    Call refresh_lanes
End Sub

Private Sub Whs_Click()
    Call refresh_lanes
End Sub
