VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form skurelease 
   Caption         =   "New Product Releases"
   ClientHeight    =   10005
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   6735
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   11880
      _Version        =   327680
      ForeColor       =   128
      BackColorFixed  =   14737632
      ForeColorFixed  =   128
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2175
      Left            =   6720
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3836
      _Version        =   327680
      ForeColor       =   12582912
      BackColorFixed  =   12648384
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3836
      _Version        =   327680
      ForeColor       =   8388736
      BackColorFixed  =   12640511
      FocusRect       =   0
   End
   Begin VB.Label whskey 
      Caption         =   "whskey"
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Label skukey 
      Caption         =   "skukey"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu edsku 
         Caption         =   "SKU List"
         Begin VB.Menu addsku 
            Caption         =   "Add SKU"
         End
         Begin VB.Menu dropsku 
            Caption         =   "Drop SKU"
         End
      End
      Begin VB.Menu edbranch 
         Caption         =   "Branch List"
         Begin VB.Menu addbranch 
            Caption         =   "Add Branch"
         End
         Begin VB.Menu dropbranch 
            Caption         =   "Drop Branch"
         End
         Begin VB.Menu edpqty 
            Caption         =   "Edit New Pallet Qty"
         End
         Begin VB.Menu postnwk 
            Caption         =   "Post Qty to Next Week"
         End
         Begin VB.Menu posttwk 
            Caption         =   "Post Qty to This Week"
         End
      End
   End
End
Attribute VB_Name = "skurelease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()
    Dim ds As ADODB.Recordset, s As String
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 3
    's = "select sku, reldate, count(*) from prodrelease group by sku, reldate order by sku"
    s = "select plantwhs, sku, count(*) from bimp"
    s = s & " where promoflag > 'N' and plantwhs in ('A10', 'K10', 'T10')"
    s = s & " group by plantwhs, sku order by sku, plantwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!plantwhs & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^Plant|^SKU|<Product"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 4000
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_grid2()
    Dim ds As ADODB.Recordset, s As String
    Screen.MousePointer = 11
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 4: Grid3.FixedCols = 1
    s = "select plantwhs, plantpool, sum(nextwknewpals), sum(thiswknewpals) from bimp"
    s = s & " Where sku = '" & skukey.Caption & "'"
    s = s & " group by plantwhs, plantpool order by plantwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!plantwhs & Chr(9)
            s = s & Int(ds!plantpool / skurec(Val(skukey)).pallet) & Chr(9)
            s = s & Format(ds(2) + ds(3), "0") & Chr(9)
            s = s & "0"
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid2.Rows > 1 Then
        For i = 1 To Grid2.Rows - 1
            Grid2.TextMatrix(i, 3) = Val(Grid2.TextMatrix(i, 1)) - Val(Grid2.TextMatrix(i, 2))
        Next i
    End If
    Grid2.FormatString = "^Plant|^OnHand|^Active|^Net"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 1000
    Grid2.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_grid3()
    Dim ds As ADODB.Recordset, s As String, i As Integer, pt As Long
    Dim t3 As Long, t4 As Long, t5 As Long, t6 As Long, t7 As Long, t8 As Currency
    Screen.MousePointer = 11
    Grid3.Redraw = False
    Grid3.FontName = "Arial"
    Grid3.FontBold = True
    Grid3.FontSize = 8
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 9: Grid3.FixedCols = 3
    s = "select id, branchwhs, onorder, thiswknewpals, nextwknewpals, promoqty from bimp"
    s = s & " where sku = '" & skukey.Caption & "' and plantwhs = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
    s = s & " and promoflag = 'Y'"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!branchwhs & Chr(9)
            s = s & branchrec(Val(ds!branchwhs)).branchname & Chr(9)
            If ds!promoqty > 0 Then
                s = s & ds!promoqty & Chr(9)
            Else
                s = s & " " & Chr(9)
            End If
            If ds!onorder > 0 Then
                s = s & CInt(ds!onorder / skurec(Val(skukey)).pallet) & Chr(9)
            Else
                s = s & " " & Chr(9)
            End If
            If ds!thiswknewpals > 0 Then
                s = s & ds!thiswknewpals & Chr(9)
            Else
                s = s & Chr(9)
            End If
            If ds!nextwknewpals > 0 Then
                s = s & ds!nextwknewpals & Chr(9)
            Else
                s = s & Chr(9)
            End If
            If ds!promoqty > 0 Then
                s = s & ds!promoqty & Chr(9)
            Else
                s = s & Chr(9)
            End If
            Grid3.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t8 = 0
    If Grid3.Rows > 1 Then
        For i = 1 To Grid3.Rows - 1
            pt = Val(Grid3.TextMatrix(i, 4))
            pt = pt + Val(Grid3.TextMatrix(i, 5))
            pt = pt + Val(Grid3.TextMatrix(i, 6))
            pt = pt + Val(Grid3.TextMatrix(i, 7))
            Grid3.TextMatrix(i, 3) = Format(pt, "#")
        Next i
                
        For i = 1 To Grid3.Rows - 1
            t3 = t3 + Val(Grid3.TextMatrix(i, 3))
            t4 = t4 + Val(Grid3.TextMatrix(i, 4))
            t5 = t5 + Val(Grid3.TextMatrix(i, 5))
            t6 = t6 + Val(Grid3.TextMatrix(i, 6))
            t7 = t7 + Val(Grid3.TextMatrix(i, 7))
        Next i
        For i = 1 To Grid3.Rows - 1
            If Val(Grid3.TextMatrix(i, 7)) > 0 Then
                Grid3.TextMatrix(i, 8) = Format((Val(Grid3.TextMatrix(i, 7)) / t7) * 100, ".000")
            Else
                Grid3.TextMatrix(i, 8) = " "
            End If
            t8 = t8 + Val(Grid3.TextMatrix(i, 8))
        Next i
        s = " " & Chr(9) & " " & Chr(9) & "Summary" & Chr(9)
        s = s & t3 & Chr(9)
        s = s & t4 & Chr(9)
        s = s & t5 & Chr(9)
        s = s & t6 & Chr(9)
        s = s & t7 & Chr(9)
        s = s & t8 'Format(t8 * 100, ".000")
        Grid3.AddItem s
    End If
    Grid3.FormatString = "^ID|^Branch|<Location|^Total|^Orders|^This Week|^Next Week|^New|^Pct."
    Grid3.ColWidth(0) = 1000
    Grid3.ColWidth(1) = 1000
    Grid3.ColWidth(2) = 2400
    Grid3.ColWidth(3) = 1200
    Grid3.ColWidth(4) = 1200
    Grid3.ColWidth(5) = 1200
    Grid3.ColWidth(6) = 1200
    Grid3.ColWidth(7) = 1200
    Grid3.ColWidth(8) = 1200
    Grid3.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub addbranch_Click()
    Dim s As String, pbranch As String, pt As Long, i As Integer, ds As ADODB.Recordset
    If Val(skukey.Caption) = 0 Then Exit Sub
    pbranch = InputBox("Branch:", "New release branch....", "001")
    If Len(pbranch) = 0 Then Exit Sub
    If Val(branchrec(Val(pbranch)).branchno) <> Val(pbranch) Then
        MsgBox "Invalid r12 warehouse.", vbOKOnly + vbInformation, "sorry, try again..."
        Exit Sub
    End If
    pbranch = Format(Val(pbranch), "000")
    For i = 1 To Grid3.Rows - 1
        If Grid3.TextMatrix(i, 2) = pbranch Then
            s = "Branch: " & pbranch & " " & Grid3.TextMatrix(i, 3) & " already listed.."
            MsgBox s, vbInformation + vbOKOnly, "duplicate record...."
            Exit Sub
        End If
    Next i
    s = "select id, onorder, roqty, thiswknewpals, nextwknewpals from bimp"
    s = s & " Where plantwhs = '" & whskey.Caption & "'"
    s = s & " and branchwhs = '" & pbranch & "'"
    s = s & " and sku = '" & skukey.Caption & "'"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        pt = 0
        s = "Update bimp set promoflag = 'Y', promoqty = 0 where id = " & ds!id
        wdb.Execute s
        i = Grid3.Row
        refresh_grid3
        DoEvents
        Grid3.Row = i
    Else
        s = "Branch " & pbranch & " is not assigned to plant " & whskey.Caption & "."
        MsgBox s, vbOKOnly + vbInformation, "sorry, cannot add branch at this time..."
    End If
    ds.Close
End Sub

Private Sub addsku_Click()
    Dim s As String, psku As String, zid As Long, pwhs As String
    psku = InputBox("SKU:", "New release product....", "735")
    If Len(psku) = 0 Then Exit Sub
    If skurec(Val(psku)).sku <> psku Then
        MsgBox "Invalid sku.", vbOKOnly + vbInformation, "sorry, try again..."
        Exit Sub
    End If
    pwhs = "T10"
    pwhs = InputBox("Plant:", "Plant Warehouse...", pwhs)
    If Len(pwhs) = 0 Then Exit Sub
    If pwhs <> "A10" And pwhs <> "K10" And pwhs <> "T10" Then
        MsgBox "Invalid plant: " & pplant, vbOKOnly + vbInformation, "sorry, try again...."
        Exit Sub
    End If
    s = "Update bimp set promoflag = 'Y', promoqty = 0 where sku = '" & psku & "'"
    s = s & " and plantwhs = '" & pwhs & "'"
    wdb.Execute s
    s = pwhs & Chr(9) & psku & Chr(9) & skurec(Val(psku)).unit & " " & skurec(Val(psku)).desc
    Grid1.AddItem s
    Grid1.Row = Grid1.Rows - 1
End Sub

Private Sub dropbranch_Click()
    Dim s As String
    If Val(Grid3.TextMatrix(Grid3.Row, 0)) = 0 Then Exit Sub
    If MsgBox("Ok to drop " & Grid3.TextMatrix(Grid3.Row, 2) & " from new release list?", vbYesNo + vbQuestion, "are you sure...") = vbNo Then Exit Sub
    s = "Update bimp set promoflag = 'N', promoqty = 0 Where id = " & Grid3.TextMatrix(Grid3.Row, 0)
    wdb.Execute s
    If Grid3.Rows <= 2 Then
        refresh_grid3
    Else
        Grid3.RemoveItem Grid3.Row
    End If
End Sub

Private Sub dropsku_Click()
    Dim s As String, psku As String, pwhs As String
    psku = Grid1.TextMatrix(Grid1.Row, 1)
    pwhs = Grid1.TextMatrix(Grid1.Row, 0)
    If Val(psku) = 0 Then Exit Sub
    If MsgBox("Ok to drop " & psku & " from " & pwhs & " new release list?", vbYesNo + vbQuestion, "are you sure...") = vbNo Then Exit Sub
    s = "Update bimp set promoflag = 'N', promoqty = 0 where sku = '" & psku & "'"
    s = s & " and plantwhs = '" & pwhs & "'"
    wdb.Execute s
    refresh_grid1
    DoEvents
    Grid1_RowColChange
End Sub

Private Sub edpqty_Click()
    Dim s As String, pqty As String, i As Integer
    If Val(Grid3.TextMatrix(Grid3.Row, 0)) = 0 Then Exit Sub
    pqty = Grid3.TextMatrix(Grid3.Row, 7)
    pqty = InputBox("Pallets:", "Pallet qty....", pqty)
    If Val(pqty) = 0 Then Exit Sub
    s = "Update bimp set promoqty = " & Val(pqty)
    s = s & " Where id = " & Grid3.TextMatrix(Grid3.Row, 0)
    wdb.Execute s
    i = Grid3.Row
    refresh_grid3
    Grid3.Row = i
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    'Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    refresh_grid1
    DoEvents
    Grid1_RowColChange
End Sub

Private Sub Form_Resize()
    If Me.Height > 2000 Then
        Grid3.Height = Me.Height - (Grid1.Height * 1.4)
    End If
End Sub

Private Sub Grid1_RowColChange()
    skukey = Grid1.TextMatrix(Grid1.Row, 1)
    whskey = Grid1.TextMatrix(Grid1.Row, 0)
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edsku
End Sub

Private Sub Grid3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edbranch
End Sub

Private Sub postnwk_Click()
    Dim s As String, pqty As String, i As Integer
    If Val(Grid3.TextMatrix(Grid3.Row, 0)) = 0 Then Exit Sub
    pqty = Grid3.TextMatrix(Grid3.Row, 7)
    pqty = InputBox("Pallets:", "Next Week Pallets....", pqty)
    If Val(pqty) = 0 Then Exit Sub
    s = "Update bimp set nextwknewpals = " & pqty
    s = s & ", promoqty = promoqty - " & pqty
    s = s & " where id = " & Grid3.TextMatrix(Grid3.Row, 0)
    wdb.Execute s
    i = Grid3.Row
    refresh_grid3
    Grid3.Row = i
End Sub

Private Sub posttwk_Click()
    Dim s As String, pqty As String, i As Integer
    If Val(Grid3.TextMatrix(Grid3.Row, 0)) = 0 Then Exit Sub
    pqty = Grid3.TextMatrix(Grid3.Row, 7)
    pqty = InputBox("Pallets:", "This Week Pallets....", pqty)
    If Val(pqty) = 0 Then Exit Sub
    s = "Update bimp set thistwknewpals = " & pqty
    s = s & ", promoqty = promoqty - " & pqty
    s = s & " where id = " & Grid3.TextMatrix(Grid3.Row, 0)
    wdb.Execute s
    i = Grid3.Row
    refresh_grid3
    Grid3.Row = i
End Sub

Private Sub skukey_Change()
    refresh_grid2
    refresh_grid3
End Sub

Private Sub whskey_Change()
    'refresh_grid2
    refresh_grid3
End Sub

