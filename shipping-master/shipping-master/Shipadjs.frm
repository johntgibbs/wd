VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Whstotals 
   Caption         =   "Inventory Adjustments"
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6450
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Drop SKU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add SKU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6015
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   10610
      _Version        =   327680
      Cols            =   6
      FixedCols       =   3
      BackColor       =   16777215
      BackColorFixed  =   12648447
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   1
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
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   2895
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
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu addsku 
         Caption         =   "Add SKU"
      End
      Begin VB.Menu dropsku 
         Caption         =   "Drop SKU"
      End
   End
End
Attribute VB_Name = "Whstotals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String
Private Sub update_inv()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    sqlx = "select * from whstotals where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Grid1.Text = Val(Grid1.Text)
        If Val(Grid1.Text) = 0 Then Grid1.Text = ""
        If edcell = "count_qty" Then
            sqlx = "Update whstotals set count_qty = " & Val(Grid1.TextMatrix(Grid1.Row, 3))
            sqlx = sqlx & ", grp_qty = " & Val(Grid1.TextMatrix(Grid1.Row, 4))
            sqlx = sqlx & ", avail = " & Val(Grid1.TextMatrix(Grid1.Row, 5)) & " Where id = " & ds!id
            Sdb.Execute sqlx
        End If
        If edcell = "grp_qty" Then
            sqlx = "Update whstotals set count_qty = " & Val(Grid1.TextMatrix(Grid1.Row, 3))
            sqlx = sqlx & ", grp_qty = " & Val(Grid1.TextMatrix(Grid1.Row, 4))
            sqlx = sqlx & ", avail = " & Val(Grid1.TextMatrix(Grid1.Row, 5)) & " Where id = " & ds!id
            Sdb.Execute sqlx
        End If
    End If
    ds.Close
    edcell = ""
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "update_inv", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " update_inv - Error Number: " & eno
        End
    End If
End Sub
Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    sqlx = "Select whstotals.id,whstotals.sku,fgunit,fgdesc,count_qty,grp_qty,avail"
    sqlx = sqlx & " from whstotals,skumast"
    sqlx = sqlx & " Where whstotals.whs_num = " & Left$(Combo1, 3)
    sqlx = sqlx & " and whstotals.sku = skumast.sku"
    sqlx = sqlx & " Order by whstotals.sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds(0) & Chr$(9)
            sqlx = sqlx & ds(1) & Chr$(9)
            sqlx = sqlx & " " & ds(2) & " " & ds(3) & Chr$(9)
            sqlx = sqlx & Format(ds(4), "#") & Chr$(9)
            sqlx = sqlx & Format(ds(5), "#") & Chr$(9)
            sqlx = sqlx & Format(ds(6), "#")
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "ID|^SKU|Product|^Count|^Grouped|^Avail"
    Grid1.ColWidth(0) = 1:    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 4000: Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000:  Grid1.ColWidth(5) = 1000
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid - Error Number: " & eno
        End
    End If
End Sub

Private Sub addsku_Click()
    Command2_Click
End Sub

Private Sub Combo1_Click()
    Call refresh_grid
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    Screen.MousePointer = 11
    Printer.FontName = "MS Sans Serif"
    Printer.FontSize = 12
    Printer.Print Combo1 & " Inventory        " & Format(Now, "mmmm d, yyyy  h:m Am/Pm")
    Printer.Print " "
    Printer.FontName = "Courier"
    Printer.FontSize = 10
    Printer.Print Tab(51); "OnHand    Grouped  Available"
    For i = 1 To Grid1.Rows - 1
        Printer.Print Grid1.TextMatrix(i, 1);
        Printer.Print Tab(5); Grid1.TextMatrix(i, 2);
        Printer.Print Tab(50); Space(6 - Len(Grid1.TextMatrix(i, 3))) & Grid1.TextMatrix(i, 3);
        Printer.Print Tab(60); Space(6 - Len(Grid1.TextMatrix(i, 4))) & Grid1.TextMatrix(i, 4);
        Printer.Print Tab(70); Space(6 - Len(Grid1.TextMatrix(i, 5))) & Grid1.TextMatrix(i, 5)
    Next i
    Printer.EndDoc
    Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
    Dim ds As adodb.Recordset, sqlx As String
    Dim msku As String, y As Integer, i As Integer, j As Integer
    Dim pkey As Long, pdesc As String
    On Error GoTo vberror
    msku = InputBox$("SKU #", "Add SKU to " & Right$(Combo1, Len(Combo1) - 4), "777")
    If Len(msku) = 0 Then Exit Sub
    sqlx = "Select * from skumast where sku = '" & msku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid SKU #.. Cannot add..", vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    End If
    pdesc = " " & ds!fgunit & " " & ds!fgdesc
    ds.Close
    sqlx = "select * from whstotals where whs_num = " & Left$(Combo1, 3)
    sqlx = sqlx & " and sku = '" & msku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        MsgBox "SKU " & msku & " quantity detected in " & Right$(Combo1, Len(Combo1) - 4), vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    End If
    pkey = wd_seq("whstotals", Form1.shipdb)
    sqlx = "Insert into whstotals (id, whs_num, sku, count_qty, grp_qty, avail, old_qty, old_lot) Values (" & pkey
    sqlx = sqlx & ", " & Val(Left(Combo1, 3))
    sqlx = sqlx & ", '" & msku & "'"
    sqlx = sqlx & ", 0, 0, 0, 0, ' ')"
    Sdb.Execute sqlx
    ds.Close
    If Grid1.Rows > 2 Then
        For i = 1 To Grid1.Rows - 1
            y = i
            If Grid1.TextMatrix(i, 1) > msku Then Exit For
        Next i
        Grid1.AddItem " "
        If y = Grid1.Rows - 2 Then
            y = y + 1
        Else
            For i = Grid1.Rows - 2 To y Step -1
                For j = 0 To Grid1.Cols - 1
                    Grid1.TextMatrix(i + 1, j) = Grid1.TextMatrix(i, j)
                Next j
            Next i
        End If
        Grid1.TextMatrix(y, 0) = pkey
        Grid1.TextMatrix(y, 1) = msku
        Grid1.TextMatrix(y, 2) = pdesc
        Grid1.TextMatrix(y, 3) = ""
        Grid1.TextMatrix(y, 4) = ""
        Grid1.TextMatrix(y, 5) = ""
        Grid1.Row = y: Grid1.Col = 3
    Else
        Call refresh_grid
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command2_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command2_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command3_Click()
    On Error GoTo vberror
    If Grid1.Row < 1 Then Exit Sub
    If MsgBox("Delete " & Grid1.TextMatrix(Grid1.Row, 2) & " From " & Combo1 & " List?", vbYesNo + vbQuestion, "Are you sure?") = vbNo Then Exit Sub
    Sdb.Execute "delete from whstotals where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_grid
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command3_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command3_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub dropsku_Click()
    Command3_Click
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Len(edcell) > 0 Then
        If MsgBox("Update inventory record?", vbYesNo + vbQuestion, "Save changes...") = vbYes Then
            Call update_inv
        Else
            edcell = ""
        End If
    End If
    If Whstotals.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            If Form1.FrmGrid.TextMatrix(i, 0) = "whstotals" Then
                Form1.FrmGrid.TextMatrix(i, 1) = Whstotals.Top
                Form1.FrmGrid.TextMatrix(i, 2) = Whstotals.Left
                Form1.FrmGrid.TextMatrix(i, 3) = Whstotals.Height
                Form1.FrmGrid.TextMatrix(i, 4) = Whstotals.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Whstotals.ActiveControl.Name = "Grid1" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command2_Click
        If KeyCode = 46 Or KeyCode = 120 Then Call Command3_Click
    End If
End Sub

Private Sub Form_Load()
    Dim ds As adodb.Recordset, sqlx As String
    Dim i As Integer
    On Error GoTo vberror
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    For i = 1 To Form1.FrmGrid.Rows - 1
        If Form1.FrmGrid.TextMatrix(i, 0) = "whstotals" Then
            Whstotals.Top = Val(Form1.FrmGrid.TextMatrix(i, 1))
            Whstotals.Left = Val(Form1.FrmGrid.TextMatrix(i, 2))
            Whstotals.Height = Val(Form1.FrmGrid.TextMatrix(i, 3))
            Whstotals.Width = Val(Form1.FrmGrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    sqlx = "select whs_num,whsname from warehouses order by whsname"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem Format$(ds(0), "000") & " " & ds(1)
            ds.MoveNext
        Loop
        Combo1.ListIndex = 0
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "form_load", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " form_load - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Whstotals.Height > 2000 Then
        Grid1.Height = Whstotals.Height - 1080
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_GotFocus()
    Grid1.FocusRect = flexFocusNone
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
    If Grid1.Row = 0 Then Exit Sub
    If Grid1.Col < 3 Or Grid1.Col > 4 Then Exit Sub
    If Len(edcell) = 0 Then Grid1.Text = ""
    If Grid1.Col = 3 Then edcell = "count_qty"
    If Grid1.Col = 4 Then edcell = "grp_qty"
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
    If edcell = "count_qty" Or edcell = "grp_qty" Then
        Grid1.TextMatrix(Grid1.Row, 5) = Val(Grid1.TextMatrix(Grid1.Row, 3)) - Val(Grid1.TextMatrix(Grid1.Row, 4))
    End If
End Sub

Private Sub Grid1_LeaveCell()
    If Len(edcell) > 0 Then Call update_inv
End Sub

Private Sub Grid1_LostFocus()
    If Len(edcell) > 0 Then Call update_inv
    Grid1.FocusRect = flexFocusLight
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub
