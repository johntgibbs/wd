VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form skuconf 
   Caption         =   "SKU Configurations"
   ClientHeight    =   5970
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   ScaleHeight     =   5970
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Rebuild SKU Recs"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete SKU"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add SKU"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   9551
      _Version        =   327680
      BackColorFixed  =   12640511
      BackColorSel    =   192
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu addsku 
         Caption         =   "Add SKU"
      End
      Begin VB.Menu delsku 
         Caption         =   "Delete SKU"
      End
   End
End
Attribute VB_Name = "skuconf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String
Private Sub update_item()
    Dim ds As adodb.Recordset, sqlx As String, zid As Long, s As String
    If edcell = "zone" Then
        sqlx = "select * from zone_config"
    Else
        sqlx = "select * from sku_config"
    End If
    sqlx = sqlx & " where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Grid1.Text = Trim(Grid1.Text)
        If edcell = "uom_type" Then
            If Len(Grid1.Text) = 0 Then Grid1.Text = " "
            If Len(Grid1.Text) > 8 Then Grid1.Text = Left(Grid1.Text, 8)
            sqlx = "Update sku_config set uom_type = '" & Grid1.Text & "'"
            sqlx = sqlx & " Where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
            Wdb.Execute sqlx
            Check1.Value = 1
        End If
        If edcell = "description" Then
            If Len(Grid1.Text) = 0 Then Grid1.Text = " "
            If Len(Grid1.Text) > 40 Then Grid1.Text = Left(Grid1.Text, 40)
            sqlx = "Update sku_config set description = '" & Grid1.Text & "'"
            sqlx = sqlx & " Where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
            Wdb.Execute sqlx
            Check1.Value = 1
        End If
        If edcell = "uom_per_pallet" Then
            Grid1.Text = Val(Grid1.Text)
            sqlx = "Update sku_config set uom_per_pallet = " & Val(Grid1.Text)
            sqlx = sqlx & " Where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
            Wdb.Execute sqlx
            Check1.Value = 1
        End If
        If edcell = "qty_per_pallet" Then
            Grid1.Text = Val(Grid1.Text)
            sqlx = "Update sku_config set qty_per_pallet = " & Val(Grid1.Text)
            sqlx = sqlx & " Where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
            Wdb.Execute sqlx
            Check1.Value = 1
        End If
        If edcell = "zone" Then
            Grid1.Text = Val(Grid1.Text)
            sqlx = "Update zone_config set zone_num = " & Val(Grid1.Text)
            sqlx = sqlx & " Where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
            Wdb.Execute sqlx
        End If
    Else
        Grid1.Text = Trim(Grid1.Text)
        If edcell = "zone" Then
            zid = wd_seq("Zone_Config")
            s = "INSERT INTO Zone_Config (ID, SKU, Whse_Num, Zone_Num, Lot_Size)"
            s = s & " VALUES (" & zid & ","
            s = s & "'" & Grid1.TextMatrix(Grid1.Row, 0) & "',3,"
            Grid1.Text = Val(Grid1.Text)
            s = s & Grid1.Text & ",0)"
            Wdb.Execute s
        End If
    End If
    ds.Close
    edcell = ""
End Sub

Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    If Form1.edlane.Enabled = True Then
        s = "select s.sku,s.uom_type,s.description,s.uom_per_pallet,s.qty_per_pallet,z.zone_num"
        s = s & " from sku_config s, zone_config z"
        s = s & " where z.sku = s.sku"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                sqlx = ds!sku & Chr$(9)
                sqlx = sqlx & " " & ds!uom_type & Chr$(9)
                sqlx = sqlx & " " & ds!description & Chr$(9)
                sqlx = sqlx & ds!uom_per_pallet & Chr$(9)
                sqlx = sqlx & ds!qty_per_pallet & Chr$(9)
                sqlx = sqlx & ds!zone_num
                Grid1.AddItem sqlx
                ds.MoveNext
            Loop
        End If
        ds.Close
        s = "select * from sku_config"
        s = s & " where sku not in (select sku from zone_config)"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                sqlx = ds!sku & Chr$(9)
                sqlx = sqlx & " " & ds!uom_type & Chr$(9)
                sqlx = sqlx & " " & ds!description & Chr$(9)
                sqlx = sqlx & ds!uom_per_pallet & Chr$(9)
                sqlx = sqlx & ds!qty_per_pallet & Chr$(9)
                sqlx = sqlx & "0"
                Grid1.AddItem sqlx
                ds.MoveNext
            Loop
        End If
    Else
        s = "select * from sku_config order by sku"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                sqlx = ds!sku & Chr$(9)
                sqlx = sqlx & " " & ds!uom_type & Chr$(9)
                sqlx = sqlx & " " & ds!description & Chr$(9)
                sqlx = sqlx & ds!uom_per_pallet & Chr$(9)
                sqlx = sqlx & ds!qty_per_pallet & Chr$(9)
                sqlx = sqlx & "0"
                Grid1.AddItem sqlx
                ds.MoveNext
            Loop
        End If
    End If
        
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 0
    Grid1.Sort = 5
    ds.Close
    Grid1.FormatString = "^SKU|^UOM|Description|^U/Pallet|^W/Pallet|^Zone"
    Grid1.ColWidth(0) = 800: Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 3400: Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000: Grid1.ColWidth(5) = 800
    Grid1.Redraw = True
End Sub

Private Sub addsku_Click()
    Command1_Click
End Sub


Private Sub Command1_Click()
    Dim ds As adodb.Recordset, sqlx As String, psku As String
    Dim i As Integer, j As Integer, y As Integer, zid As Long
    psku = InputBox$("SKU #", "New SKU", "000")
    If Len(psku) = 0 Then Exit Sub
    psku = Trim(Left(psku, 4))                                                      'jv082415
    sqlx = "select sku,uom_type,description from sku_config where sku = '" & psku & "'"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        MsgBox "SKU# " & psku & " Already exists for " & ds!uom_type & " " & ds!description, vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    Else
        s = "INSERT INTO SKU_Config (SKU, SKU_Type, UOM_Per_Pallet, Qty_Per_Pallet, Select_Method)"
        s = s & " VALUES ('" & psku & "','F',1,1,'R')"
        Wdb.Execute s
    End If
    ds.Close
    If Form1.edlane.Enabled = True Then
        zid = wd_seq("Zone_Config")
        s = "INSERT INTO Zone_Config (ID, SKU, Whse_Num, Zone_Num, Lot_Size)"
        s = s & " VALUES (" & zid & ","
        s = s & "'" & psku & "',"
        s = s & "3,2,0)"
        Wdb.Execute s
    End If
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) > psku Then Exit For
        y = i
    Next i
    Grid1.AddItem " "
    For i = Grid1.Rows - 2 To y Step -1
        For j = 0 To Grid1.Cols - 1
            Grid1.TextMatrix(i + 1, j) = Grid1.TextMatrix(i, j)
            Grid1.TextMatrix(i, j) = ""
        Next j
    Next i
    Grid1.TextMatrix(y, 0) = psku
    Grid1.TextMatrix(y, 3) = "1"
    Grid1.TextMatrix(y, 4) = "1"
    Grid1.TextMatrix(y, 5) = "2"
    Grid1.Row = y: Grid1.Col = 1
    Check1.Value = 1
End Sub

Private Sub Command2_Click()
    Dim sqlx As String
    If Grid1.Row = 0 Then Exit Sub
    If MsgBox("Ok to delete SKU# " & Grid1.TextMatrix(Grid1.Row, 0), vbYesNo, "Delete SKU") = vbNo Then Exit Sub
    sqlx = "delete from sku_config where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
    Wdb.Execute sqlx
    If Form1.edlane.Enabled = True Then
        sqlx = "delete from zone_config where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
        Wdb.Execute sqlx
    End If
    Grid1.RemoveItem Grid1.Row
    Check1.Value = 1
End Sub

Private Sub delsku_Click()
    Command2_Click
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Len(edcell) > 0 Then
        If MsgBox("Save changes to " & edcell & "?", vbYesNo + vbQuestion, "Save changes..") = vbYes Then
            Call update_item
        End If
    End If
    If skuconf.WindowState = 0 Then
        For i = 1 To Form1.Frmgrid.Rows - 1
            If Form1.Frmgrid.TextMatrix(i, 0) = "skuconf" Then
                Form1.Frmgrid.TextMatrix(i, 1) = skuconf.Top
                Form1.Frmgrid.TextMatrix(i, 2) = skuconf.Left
                Form1.Frmgrid.TextMatrix(i, 3) = skuconf.Height
                Form1.Frmgrid.TextMatrix(i, 4) = skuconf.Width
                Exit For
            End If
        Next i
    End If
    If Check1.Value = 1 Then
        Screen.MousePointer = 11
        Call build_sku_config
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If skuconf.ActiveControl.Name = "Grid1" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command1_Click
        If KeyCode = 46 Or KeyCode = 120 Then Call Command2_Click
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    edcell = ""
    For i = 1 To Form1.Frmgrid.Rows - 1
        If Form1.Frmgrid.TextMatrix(i, 0) = "skuconf" Then
            skuconf.Top = Val(Form1.Frmgrid.TextMatrix(i, 1))
            skuconf.Left = Val(Form1.Frmgrid.TextMatrix(i, 2))
            skuconf.Height = Val(Form1.Frmgrid.TextMatrix(i, 3))
            skuconf.Width = Val(Form1.Frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Call refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Height = Me.Height - 800
    Grid1.Width = Me.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Grid1.Col = Grid1.Cols - 1 Then
            SendKeys "{HOME}{DOWN}"
        Else
            SendKeys "{RIGHT}"
        End If
        Exit Sub
    End If
    If Grid1.Row = 0 Or Grid1.Col = 0 Then Exit Sub
    If Form1.edlane.Enabled = False And Grid1.Col > 4 Then Exit Sub
    If Len(edcell) = 0 Then Grid1.Text = ""
    If Grid1.Col = 1 Then edcell = "uom_type"
    If Grid1.Col = 2 Then edcell = "description"
    If Grid1.Col = 3 Then edcell = "uom_per_pallet"
    If Grid1.Col = 4 Then edcell = "qty_per_pallet"
    If Grid1.Col = 5 Then edcell = "zone"
    If KeyAscii = 8 Then
        If Len(Grid1.Text) > 1 Then
            Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
        Else
            Grid1.Text = ""
        End If
    End If
    If KeyAscii > 31 And KeyAscii < 127 Then
        Grid1.Text = Grid1.Text & Chr(KeyAscii)
    End If

End Sub
Private Sub Grid1_LeaveCell()
    If Len(edcell) > 0 Then Call update_item
End Sub

Private Sub Grid1_LostFocus()
    If Len(edcell) > 0 Then Call update_item
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

