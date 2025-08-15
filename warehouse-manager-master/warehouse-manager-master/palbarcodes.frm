VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form palbarcodes 
   Caption         =   "Pallet BarCodes"
   ClientHeight    =   10800
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13740
   LinkTopic       =   "Form14"
   ScaleHeight     =   10800
   ScaleWidth      =   13740
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   8055
      Left            =   10320
      TabIndex        =   9
      Top             =   600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   9720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1931
      _Version        =   327680
      BackColorFixed  =   12648447
      FocusRect       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   15266
      _Version        =   327680
      ForeColor       =   4210688
      BackColorFixed  =   12648384
      FocusRect       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label shipcount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pallet(s) Shipped."
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
      Left            =   10320
      TabIndex        =   8
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label dupcount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Duplicate BarCodes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10320
      TabIndex        =   7
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label reccount 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label pkey 
      Caption         =   "pkey"
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
      TabIndex        =   4
      Top             =   9360
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Warehouse:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu inspallet 
         Caption         =   "Insert Pallet Record"
      End
      Begin VB.Menu delpallet 
         Caption         =   "Clear Pallet Record"
      End
      Begin VB.Menu edpallet 
         Caption         =   "Edit Field"
      End
   End
   Begin VB.Menu lookmenu 
      Caption         =   "Lookup"
      Visible         =   0   'False
      Begin VB.Menu batonhand 
         Caption         =   "Batch Inventory"
      End
      Begin VB.Menu palhis 
         Caption         =   "Pallet History"
      End
   End
   Begin VB.Menu actmenu 
      Caption         =   "Action"
      Begin VB.Menu finddups 
         Caption         =   "Search For Duplicates"
      End
      Begin VB.Menu findships 
         Caption         =   "Search for Shipped Pallets"
      End
   End
End
Attribute VB_Name = "palbarcodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()
    Dim ds As ADODB.Recordset, s As String, i As Integer, rs As ADODB.Recordset
    Screen.MousePointer = 11
    dupcount.Visible = False
    shipcount.Visible = False
    List1.Clear
    List1.Visible = False
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    If Combo1 = "SR1" Or Combo1 = "SR2" Or Combo1 = "SR3" Or Combo1 = "SR5" Then
        s = "Select id,whse_num,vert_loc,horz_loc,rack_side,posn_num,sku,barcode"
        s = s & " from position where whse_num = " & Mid(Combo1, 3, 1)
        s = s & " and count_qty > 0 order by barcode"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds!id & Chr(9)
                s = s & "SR" & ds!whse_num & Chr(9)
                s = s & ds!vert_loc & " "
                s = s & ds!horz_loc & " "
                s = s & ds!rack_side & " "
                s = s & ds!posn_num & Chr(9)
                s = s & Mid(ds!barcode, 1, 10) & " "
                s = s & Mid(ds!barcode, 11, 3) & " "
                s = s & Mid(ds!barcode, 14, 3) & Chr(9)
                i = Val(Trim(Mid(ds!barcode, 1, 4)))
                s = s & skurec(i).prodname & Chr(9)
                s = s & ds!barcode
                Grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    If Combo1 = "SR4" Then
        s = "select rackpos.id,rackpos.sku,rackpos.barcode,racks.aisle,racks.rack"
        s = s & " from rackpos, racks where rackpos.count_qty > 0"
        s = s & " and racks.id = rackpos.rackno"
        s = s & " order by rackpos.barcode"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds!id & Chr(9)
                s = s & "SR4" & Chr(9)
                s = s & ds!aisle & "-"
                s = s & ds!rack & Chr(9)
                s = s & Mid(ds!barcode, 1, 10) & " "
                s = s & Mid(ds!barcode, 11, 3) & " "
                s = s & Mid(ds!barcode, 14, 3) & Chr(9)
                i = Val(Trim(Mid(ds!barcode, 1, 4)))
                s = s & skurec(i).prodname & Chr(9)
                s = s & ds!barcode
                Grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    If Combo1 = "ALL" Then
        s = "Select id,whse_num,vert_loc,horz_loc,rack_side,posn_num,sku,barcode"
        s = s & " from position where whse_num > 0"
        s = s & " and count_qty > 0 order by barcode"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds!id & Chr(9)
                s = s & "SR" & ds!whse_num & Chr(9)
                s = s & ds!vert_loc & " "
                s = s & ds!horz_loc & " "
                s = s & ds!rack_side & " "
                s = s & ds!posn_num & Chr(9)
                s = s & Mid(ds!barcode, 1, 10) & " "
                s = s & Mid(ds!barcode, 11, 3) & " "
                s = s & Mid(ds!barcode, 14, 3) & Chr(9)
                i = Val(Trim(Mid(ds!barcode, 1, 4)))
                s = s & skurec(i).prodname & Chr(9)
                s = s & ds!barcode
                Grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
        s = "select rackpos.id,rackpos.sku,rackpos.barcode,racks.aisle,racks.rack"
        s = s & " from rackpos, racks where rackpos.count_qty > 0"
        s = s & " and racks.id = rackpos.rackno"
        s = s & " order by rackpos.barcode"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds!id & Chr(9)
                s = s & "SR4" & Chr(9)
                s = s & ds!aisle & "-"
                s = s & ds!rack & Chr(9)
                s = s & Mid(ds!barcode, 1, 10) & " "
                s = s & Mid(ds!barcode, 11, 3) & " "
                s = s & Mid(ds!barcode, 14, 3) & Chr(9)
                i = Val(Trim(Mid(ds!barcode, 1, 4)))
                s = s & skurec(i).prodname & Chr(9)
                s = s & ds!barcode
                Grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
        Grid1.Col = 5: Grid1.ColSel = 5
        Grid1.Sort = 5
    End If
    reccount.Caption = Grid1.Rows - 1 & " Records."
    s = "^Id|^Whs|^Rack Pos|^Barcode|<Product"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 1500
    Grid1.ColWidth(3) = 2500
    Grid1.ColWidth(4) = 5000
    Grid1.ColWidth(5) = 0 '1800
    Grid1.Row = 1: Grid1.Col = 3
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_grid2()
    Dim ds As ADODB.Recordset, s As String
    inspallet.Enabled = False
    delpallet.Enabled = False
    edpallet.Enabled = False
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 13
    s = "Select * from pallets where barcode = '" & pkey & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!plateno & Chr(9)
            s = s & ds!barcode & Chr(9)
            s = s & ds!qty1 & Chr(9)
            s = s & ds!lot1 & Chr(9)
            s = s & ds!qty2 & Chr(9)
            s = s & ds!lot2 & Chr(9)
            s = s & ds!source & Chr(9)
            s = s & ds!target & Chr(9)
            s = s & ds!bbc & Chr(9)
            s = s & ds!status & Chr(9)
            s = s & ds!trandate & Chr(9)
            s = s & ds!sku
            Grid2.AddItem s
            ds.MoveNext
        Loop
        delpallet.Enabled = True
        edpallet.Enabled = True
    Else
        If Grid1.TextMatrix(Grid1.Row, 1) <> "SR5" Then inspallet.Enabled = True
    End If
    ds.Close
    s = "^Id|^Plate|<Barcode|^Qty1|^Lot1|^Qty2|^Lot2|<Source|<Target|^BBC|<Status|^TranDate|^SKU"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1800
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 1000
    Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 1000
    Grid2.ColWidth(7) = 1400
    Grid2.ColWidth(8) = 1200
    Grid2.ColWidth(9) = 1000
    Grid2.ColWidth(10) = 1400
    Grid2.ColWidth(11) = 1400
    Grid2.ColWidth(12) = 1000
End Sub

Private Sub batonhand_Click()
    Dim s As String
    s = Left(Grid1.TextMatrix(Grid1.Row, 5), 13)
    tktonhand.bbarcode = s
    tktonhand.bproduct = Grid1.TextMatrix(Grid1.Row, 4)
    tktonhand.Show
End Sub

Private Sub Combo1_Click()
    refresh_grid1
End Sub

Private Sub Command1_Click()
    refresh_grid1
End Sub

Private Sub delpallet_Click()
    Dim s As String
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) > 0 Then
        s = "Update pallets set plateno = '0', status = 'Shipped' where id = " & Grid2.TextMatrix(Grid2.Row, 0)
        MsgBox s
        Wdb.Execute s
        'refresh_grid2
    End If
End Sub

Private Sub edpallet_Click()
    Dim s As String, f As String
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    f = Grid2.TextMatrix(0, Grid2.Col)
    s = InputBox(f, f, Grid2.Text)
    If Len(s) = 0 Then Exit Sub
    s = "Update pallets set " & f & "='" & s & "' Where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    'MsgBox s
    Wdb.Execute s
    refresh_grid2
End Sub

Private Sub finddups_Click()
    Dim i As Integer, c As Integer
    Screen.MousePointer = 11
    Grid1.FillStyle = flexFillRepeat
    dupcount.Visible = False
    shipcount.Visible = False
    List1.Clear
    List1.Visible = False
    c = 0
    For i = 1 To Grid1.Rows - 2
        If Grid1.TextMatrix(i, 5) = Grid1.TextMatrix(i + 1, 5) Then
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 3: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellBackColor = dupcount.BackColor
            Grid1.Row = i + 1: Grid1.RowSel = i + 1
            Grid1.Col = 3: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellBackColor = dupcount.BackColor
            c = c + 1
            Grid1.TopRow = i: Grid1.Col = 2
            List1.AddItem Format(i, "00000") & " " & Grid1.TextMatrix(i, 5)
        End If
    Next i
    If c > 0 Then
        dupcount.Caption = c & " Duplicate BarCodes."
        dupcount.Visible = True
        List1.Visible = True
    End If
    Screen.MousePointer = 0
End Sub

Private Sub findships_Click()
    Dim i As Integer, c As Integer, k As Integer
    Screen.MousePointer = 11
    Grid1.FillStyle = flexFillRepeat
    dupcount.Visible = False
    shipcount.Visible = False
    List1.Clear
    List1.Visible = False
    c = 0
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i
        DoEvents
        If Grid2.Rows > 1 Then
            For k = 1 To Grid2.Rows - 1
                If Grid2.TextMatrix(k, 2) = "Shipped" Then
                    Grid1.Row = i: Grid1.RowSel = i
                    Grid1.Col = 3: Grid1.ColSel = Grid1.Cols - 1
                    Grid1.CellBackColor = shipcount.BackColor
                    c = c + 1
                    Grid1.TopRow = i: Grid1.Col = 2
                    List1.AddItem Format(i, "00000") & " " & Grid1.TextMatrix(i, 5)
                    Exit For
                End If
            Next k
        Else
            If Val(Mid(pkey, 11, 3)) > 100 Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 3: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = shipcount.BackColor
                c = c + 1
                Grid1.TopRow = i: Grid1.Col = 2
                List1.AddItem Format(i, "00000") & " " & Grid1.TextMatrix(i, 5)
            End If
        End If
    Next i
    If c > 0 Then
        shipcount.Caption = c & " Pallets(s) Shipped."
        shipcount.Visible = True
        List1.Visible = True
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "SR1"
    Combo1.AddItem "SR2"
    Combo1.AddItem "SR3"
    Combo1.AddItem "SR4"
    Combo1.AddItem "SR5"
    Combo1.AddItem "ALL"
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 280
    Grid2.Width = Grid1.Width
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu lookmenu
End Sub

Private Sub Grid1_RowColChange()
    pkey.Caption = Grid1.TextMatrix(Grid1.Row, 5)
End Sub

Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub inspallet_Click()
    Dim ds As ADODB.Recordset, s As String, i As Long
    Dim mbc As String, mqty1 As Integer, mlot1 As String
    Dim mqty2 As Integer, mlot2 As String, msku As String, mtype As String
    
    If Grid1.TextMatrix(Grid1.Row, 1) = "SR4" Then
        s = "select * from rackpos where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            mbc = ds!barcode
            mqty1 = ds!count_qty
            mlot1 = ds!lot_num
            mqty2 = ds!qty2
            mlot2 = ds!lot2
            msku = ds!sku
            mtype = ds!bbc
        End If
        ds.Close
    Else
        s = "select * from position where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            mbc = ds!barcode
            mqty1 = ds!count_qty
            mlot1 = ds!lot_num
            mqty2 = ds!qty2
            mlot2 = ds!lot2
            msku = ds!sku
            mtype = "Y"
        End If
        ds.Close
    End If
    
    s = "select id from pallets where status = 'Shipped' order by id"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update pallets set plateno = '0'"
        s = s & ", barcode = '" & mbc & "'"
        s = s & ", qty1 = " & mqty1
        s = s & ", lot1 = '" & mlot1 & "'"
        s = s & ", qty2 = " & mqty2
        s = s & ", lot2 = '" & mlot2 & "'"
        s = s & ", source = 'WMS'"
        s = s & ", target = '" & Grid1.TextMatrix(Grid1.Row, 1) & "'"
        s = s & ", bbc = '" & mtype & "'"
        s = s & ", status = 'Warehouse', trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
        s = s & ", sku = '" & msku & "'"
        s = s & " where id = " & ds!id
        'MsgBox s
        Wdb.Execute s
    End If
    ds.Close
    refresh_grid2
End Sub

Private Sub List1_Click()
    Dim i As Integer
    i = Val(Mid(List1, 1, 5))
    Grid1.Row = i: Grid1.Col = 2
    DoEvents
    Grid1.TopRow = i
End Sub

Private Sub palhis_Click()
    palhistory.Show
    palhistory.barkey = Grid1.TextMatrix(Grid1.Row, 5)
End Sub

Private Sub pkey_Change()
    refresh_grid2
End Sub

