VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form daiwmsplate 
   Caption         =   "Daifuku Plate Maintenance"
   ClientHeight    =   12780
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11565
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form14"
   ScaleHeight     =   12780
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   1215
      Left            =   0
      TabIndex        =   6
      Top             =   11040
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2143
      _Version        =   327680
      BackColorFixed  =   12640511
      BackColorSel    =   32768
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   9360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2143
      _Version        =   327680
      BackColorFixed  =   16761024
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8775
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   15478
      _Version        =   327680
      ForeColor       =   16711680
      BackColorFixed  =   12648447
      FocusRect       =   0
   End
   Begin VB.Label bced 
      Caption         =   "Label3"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   10680
      Width           =   2895
   End
   Begin VB.Label bckey 
      Caption         =   "Label3"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   10680
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "BarCode:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   10680
      Width           =   975
   End
   Begin VB.Label pkey 
      Caption         =   "000001"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   9000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Plate:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   9000
      Width           =   735
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu insplate 
         Caption         =   "Insert Pallet Record"
      End
      Begin VB.Menu clrplate 
         Caption         =   "Clear Pallet Plate"
      End
      Begin VB.Menu edfield 
         Caption         =   "Edit Field"
      End
      Begin VB.Menu scanplates 
         Caption         =   "Scan for Issues"
      End
   End
End
Attribute VB_Name = "daiwmsplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()
    Dim cfile As String, s As String
    Dim sload As String, sitem As String, slot As String, saddress As String
    Dim sqty As String, sholdtype As String
    Screen.MousePointer = 11
    'cfile = "\\bbc-01-daifuku\d\daifuku\data\WRxInventoryData.txt"
    cfile = "v:\data\sr5.csv"
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, sload, sitem, slot, saddress, sqty, sholdtype
        s = Trim(sload) & Chr(9)
        s = s & Trim(sitem) & Chr(9)
        s = s & Trim(slot) & Chr(9)
        s = s & Trim(saddress) & Chr(9)
        s = s & Trim(sqty) & Chr(9)
        s = s & Trim(sholdtype)
        Grid1.AddItem s
    Loop
    Close #1
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 1
    Grid1.Sort = 5
    Grid1.Row = 1: Grid1.Col = 2
    Grid1.FormatString = "^Plate|^SKU|^Lot|^Address|^Qty|^HoldType"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 1600
    Grid1.ColWidth(3) = 1600
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1200
    Screen.MousePointer = 0
End Sub

Private Sub refresh_grid2()
    Dim ds As adodb.Recordset, s As String
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 13
    s = "Select * from pallets where plateno = '" & pkey & "'"
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
    Grid2.ColWidth(11) = 1600
    Grid2.ColWidth(12) = 1000
End Sub

Private Sub refresh_grid3()
    Dim ds As adodb.Recordset, s As String
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 10
    s = "Select * from position where barcode = '" & bckey & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!whse_num & Chr(9)
            s = s & ds!vert_loc & Chr(9)
            s = s & ds!horz_loc & Chr(9)
            s = s & ds!rack_side & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & ds!count_qty & Chr(9)
            s = s & ds!recv_date & Chr(9)
            s = s & ds!lot2 & Chr(9)
            s = s & ds!qty2
            Grid3.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "^Id|^SR|^Vert|^Horz|^Side|^SKU|^Units|^RecDate|^Lot2|^Qty2"
    Grid3.FormatString = s
    Grid3.ColWidth(0) = 1000
    Grid3.ColWidth(1) = 1000
    Grid3.ColWidth(2) = 1000
    Grid3.ColWidth(3) = 1000
    Grid3.ColWidth(4) = 1000
    Grid3.ColWidth(5) = 1000
    Grid3.ColWidth(6) = 1000
    Grid3.ColWidth(7) = 1400
    Grid3.ColWidth(8) = 1000
    Grid3.ColWidth(9) = 1000
End Sub

Private Sub bckey_Change()
    bced = Mid(bckey, 1, 4) & " "
    bced = bced & Mid(bckey, 5, 6) & " "
    bced = bced & Mid(bckey, 11, 3) & " "
    bced = bced & Mid(bckey, 14, 3)
    refresh_grid3
End Sub

Private Sub clrplate_Click()
    Dim s As String
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) > 0 Then
        s = "Update pallets set plateno = '0', status = 'Shipped' where id = " & Grid2.TextMatrix(Grid2.Row, 0)
        'MsgBox s
        Wdb.Execute s
        refresh_grid2
    End If
End Sub

Private Sub edfield_Click()
    Dim s As String, f As String
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    f = Grid2.TextMatrix(0, Grid2.Col)
    s = InputBox(f, f, Grid2.Text)
    If Len(s) = 0 Then Exit Sub
    s = "Update pallets set " & f & "='" & s & "' Where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    MsgBox s
    Wdb.Execute s
    refresh_grid2
End Sub

Private Sub Form_Load()
    refresh_grid1
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 200
    Grid2.Width = Me.Width - 200
    Grid3.Width = Me.Width - 200
End Sub

Private Sub Grid1_RowColChange()
    Dim r12lot As String, s As String, t As String, b As String, plot As String
    pkey.Caption = Grid1.TextMatrix(Grid1.Row, 0)
    b = Trim(Grid1.TextMatrix(Grid1.Row, 1))
    If Len(b) = 3 Then b = b & " "
    'b = b & Mid(Grid1.TextMatrix(Grid1.Row, 2), 1, 5) & " "
    plot = Mid(Grid1.TextMatrix(Grid1.Row, 2), 1, 5)
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
    b = b & r12lot
    b = b & Mid(Grid1.TextMatrix(Grid1.Row, 2), 6, 6)
    bckey.Caption = b
End Sub

Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub insplate_Click()
    Dim ds As adodb.Recordset, s As String, i As Long
    Dim slot As String, s1 As String, s2 As String, sbc As String, sitem As String
    sitem = Grid1.TextMatrix(Grid1.Row, 1)
    slot = Grid1.TextMatrix(Grid1.Row, 2)
    If Len(slot) >= 5 Then
        s1 = "12-31-20" & Format(Val(Left(slot, 2)) - 1, "00")
        s2 = Format(DateAdd("d", Val(Mid(slot, 3, 3)), s1), "MM-dd-yyyy")
    End If

    If Len(slot) > 5 Then
        s2 = Format(DateAdd("yyyy", 2, s2), "MMddyy")
        If Len(sitem) = 3 Then                          'jv090115
            sbc = sitem & " " & s2 & Mid(slot, 6, 3) & Mid(slot, 9, 3)   'jv090115
        Else
            sbc = sitem & s2 & Mid(slot, 6, 3) & Mid(slot, 9, 3)   'jv090115
        End If
    Else
        If Len(sitem) = 3 Then                          'jv090115
            sbc = sitem & " " & Left(slot, 5) & " " & Mid(slot, 6, 3) & " " & Mid(slot, 9, 3)    'jv090115
        Else
            sbc = sitem & Left(slot, 5) & " " & Mid(slot, 6, 3) & " " & Mid(slot, 9, 3)    'jv090115
        End If
    End If
    
    
    s = "select id from pallets where status = 'Shipped' order by id"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update pallets set plateno = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
        s = s & ", barcode = '" & sbc & "'"
        s = s & ", qty1 = " & Val(Grid1.TextMatrix(Grid1.Row, 4))
        s = s & ", lot1 = '" & Left(slot, 5) & "'"
        s = s & ", qty2 = 0, lot2 = ' '"
        s = s & ", source = 'TRI-LEVEL', target = 'SR5', bbc = 'Y'"
        s = s & ", status = 'Warehouse', trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
        s = s & ", sku = '" & sitem & "'"
        s = s & " where id = " & ds!id
        'MsgBox s
        Wdb.Execute s
    End If
    ds.Close
    refresh_grid2
End Sub

Private Sub pkey_Change()
    refresh_grid2
End Sub

Private Sub scanplates_Click()
    Dim i As Integer
    For i = Grid1.Row To Grid1.Rows - 1
        Grid1.Row = i: Grid1.TopRow = i
        DoEvents
        If Grid2.Rows <> 2 Then Exit For
        If Grid3.Rows <> 2 Then Exit For
    Next i
End Sub
