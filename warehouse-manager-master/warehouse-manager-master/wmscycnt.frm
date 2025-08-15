VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form wmscycnt 
   Caption         =   "Cycle Counts"
   ClientHeight    =   9390
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "Form8"
   ScaleHeight     =   9390
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "View Count Listing"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   7858
      _Version        =   327680
      BackColorFixed  =   65280
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8070
      _Version        =   327680
      BackColorFixed  =   65535
      BackColorSel    =   33023
      FocusRect       =   0
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu insrec 
         Caption         =   "New Record"
      End
      Begin VB.Menu edfield 
         Caption         =   "Edit Field"
      End
      Begin VB.Menu mc 
         Caption         =   "Mark Complete"
      End
      Begin VB.Menu mp 
         Caption         =   "Mark Pending"
      End
   End
End
Attribute VB_Name = "wmscycnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()             'Count Items
    Dim ds As ADODB.Recordset, s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 9
    s = "select * from counttasks"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!aisle & Chr(9)
            s = s & ds!rack & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & ds!product & Chr(9)
            s = s & ds!lotnum & Chr(9)
            s = s & ds!status & Chr(9)
            s = s & ds!userid & Chr(9)
            s = s & ds!notes
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "^ID|^Aisle|^Rack|^SKU|<Product|^Lot|^Status|^UserID|<Notes"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 3000
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 2000
End Sub

Private Sub refresh_grid2()             'Count Listing Racks
    Dim ds As ADODB.Recordset, s As String, wc As Integer
    Dim mrack As String, maisle As String, msku As String, mlot As String, mprod As String
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 7
    maisle = Grid1.TextMatrix(Grid1.Row, 1)
    mrack = Grid1.TextMatrix(Grid1.Row, 2)
    msku = Grid1.TextMatrix(Grid1.Row, 3)
    mlot = Grid1.TextMatrix(Grid1.Row, 5)
    s = "select * from rackpos where rackno in (select id from racks where aisle = '" & maisle & "'"
    If mrack <> "ALL" Then s = s & " and rack = '" & mrack & "'"
    s = s & ") and (count_qty + qty2) > 0"
    If msku <> "ALL" Then s = s & " and sku = '" & msku & "'"
    If mlot <> "ALL" Then s = s & " and (lot_num = '" & mlot & "' or lot2 = '" & mlot & "')"
    s = s & " order by rackno, posn_num, sku"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            wc = 1
            If skurec(Val(ds!sku)).sku = ds!sku Then
                mprod = skurec(Val(ds!sku)).prodname
                If skurec(Val(ds!sku)).uom_per_pallet > 0 And skurec(Val(ds!sku)).qty_per_pallet > 0 Then
                    wc = skurec(Val(ds!sku)).uom_per_pallet / skurec(Val(ds!sku)).qty_per_pallet
                End If
            Else
                mprod = "Product: " & ds!sku
            End If
            s = "P" & ds!id & Chr(9)
            s = s & maisle & "-" & mrack & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & mprod & Chr(9)
            s = s & (ds!count_qty + ds!qty2) / wc & Chr(9)
            s = s & "Wraps" & Chr(9)
            s = s & ds!barcode
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "^ID|^Rack|^SKU|<Product|^Qty|^UOM|^Lot"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 3500
    Grid2.ColWidth(4) = 1000
    Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 1800
End Sub

Private Sub Command1_Click()            'View Count Listing
    refresh_grid2
End Sub

Private Sub edfield_Click()
    Dim ds As ADODB.Recordset, s As String, mfld As String, mcol As String, mprod As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    mfld = Grid1.Text
    mcol = UCase(Grid1.TextMatrix(0, Grid1.Col))
    mfld = InputBox(mcol & ":", "Edit " & mcol, mfld)
    If Len(mfld) = 0 Then Exit Sub
    If mfld = Grid1.Text Then Exit Sub
    If mcol <> "NOTES" Then mfld = UCase(mfld)
    
    If mcol = "SKU" Then
        If mfld = "ALL" Then
            mprod = "All Products"
        Else
            If skurec(Val(mfld)).sku = mfld Then
                mprod = skurec(Val(mfld)).prodname
            Else
                mprod = "Invalid SKU"
            End If
        End If
    End If
    
    s = "select * from counttasks where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If mcol = "AISLE" Then
            Grid1.TextMatrix(Grid1.Row, 1) = mfld
            s = "Update counttasks set aisle = '" & mfld & "'"
        End If
        If mcol = "RACK" Then
            Grid1.TextMatrix(Grid1.Row, 2) = mfld
            s = "Update counttasks set rack = '" & mfld & "'"
        End If
        If mcol = "SKU" Then
            Grid1.TextMatrix(Grid1.Row, 3) = mfld
            Grid1.TextMatrix(Grid1.Row, 4) = mprod
            s = "Update counttasks set sku = '" & mfld & "', product = '" & mprod & "'"
        End If
        If mcol = "LOT" Then
            Grid1.TextMatrix(Grid1.Row, 5) = mfld
            s = "Update counttasks set lotnum = '" & mfld & "':"
        End If
        If mcol = "NOTES" Then
            Grid1.TextMatrix(Grid1.Row, 8) = mfld
            s = "Update counttasks set notes = '" & mfld & "'"
        End If
        Grid1.TextMatrix(Grid1.Row, 6) = "NEW"
        s = s & ", status = 'NEW'"
        Grid1.TextMatrix(Grid1.Row, 7) = "."
        s = s & ", userid = '.' Where id = " & ds!id
        Wdb.Execute s
    End If
    ds.Close
End Sub

Private Sub Form_Load()
    refresh_grid1
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    Grid2.Width = Me.Width - 100
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid1_RowColChange()
    edfield.Enabled = False
    If Grid1.Row = 0 Then Exit Sub
    If Grid1.Col = 1 Then edfield.Enabled = True
    If Grid1.Col = 2 Then edfield.Enabled = True
    If Grid1.Col = 3 Then edfield.Enabled = True
    If Grid1.Col = 5 Then edfield.Enabled = True
    If Grid1.Col = 8 Then edfield.Enabled = True
End Sub

Private Sub insrec_Click()              'New Record
    Dim maisle As String, mrack As String, msku As String, mlot As String, mprod As String
    Dim ds As ADODB.Recordset, s As String, mmid As Long
    
    maisle = InputBox("Aisle:", "Specify An Aisle or ALL....", "ALL")
    If Len(maisle) = 0 Then Exit Sub
    maisle = UCase(maisle)
    
    mrack = InputBox("Rack:", "Specify A Rack or ALL....", "ALL")
    If Len(mrack) = 0 Then Exit Sub
    mrack = UCase(mrack)
    
    msku = InputBox("SKU:", "Specify SKU or ALL....", "ALL")
    If Len(msku) = 0 Then Exit Sub
    msku = UCase(msku)
    
    mlot = InputBox("LotCode:", "Specify A Lot or ALL....", "ALL")
    If Len(mlot) = 0 Then Exit Sub
    mlot = UCase(mlot)
    
    If msku <> "ALL" Then
        If skurec(Val(msku)).sku = msku Then
            mprod = skurec(Val(msku)).prodname
        Else
            mprod = "Invalid SKU"
        End If
    Else
        mprod = "All Products"
    End If
    
    mmid = wd_seq("CountTasks")
    s = "INSERT INTO counttasks (ID, Aisle, Rack, SKU, Product, LotNum, Status, UserID)"
    s = s & " VALUES ("
    s = s & mmid & ","
    s = s & "'" & maisle & "',"
    s = s & "'" & mrack & "',"
    s = s & "'" & msku & "',"
    s = s & "'" & mprod & "',"
    s = s & "'" & mlot & "',"
    s = s & "'NEW',"
    s = s & "'.')"
    Wdb.Execute s
    
    s = mmid & Chr(9)
    s = s & maisle & Chr(9)
    s = s & mrack & Chr(9)
    s = s & msku & Chr(9)
    s = s & mprod & Chr(9)
    s = s & mlot & Chr(9)
    s = s & "NEW" & Chr(9)
    s = s & "."
    Grid1.AddItem s
    
End Sub

Private Sub mc_Click()                  'Mark Completed
    Dim s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    Screen.MousePointer = 11
    Grid1.TextMatrix(Grid1.Row, 6) = "COMP"
    Grid1.TextMatrix(Grid1.Row, 7) = "."
    s = "Update counttasks set status = 'COMP', userid = '.' Where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Wdb.Execute s
    Screen.MousePointer = 0
End Sub

Private Sub mp_Click()                  'Mark Pending
    Dim s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    Screen.MousePointer = 11
    Grid1.TextMatrix(Grid1.Row, 6) = "PEND"
    Grid1.TextMatrix(Grid1.Row, 7) = "."
    s = "Update counttasks set status = 'PEND', userid = '.' Where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Wdb.Execute s
    Screen.MousePointer = 0
End Sub
