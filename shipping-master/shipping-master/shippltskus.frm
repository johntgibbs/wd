VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Plantskus 
   Caption         =   "Plant SKU Listing"
   ClientHeight    =   5400
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5400
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Drop SKU"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add SKU"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8916
      _Version        =   327680
      Cols            =   7
      FixedCols       =   3
      ForeColor       =   8388608
      BackColorFixed  =   12648447
      BackColorBkg    =   -2147483633
      FocusRect       =   0
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   3495
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu insrec 
         Caption         =   "Insert SKU"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete SKU"
      End
   End
End
Attribute VB_Name = "Plantskus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    sqlx = "select plantskus.id,plantskus.sku,fgunit,fgdesc,lowstk,outstk,lowflag,outflag"
    sqlx = sqlx & " from plantskus,skumast"
    sqlx = sqlx & " where plantskus.plant = " & Left$(Combo1, 3)
    sqlx = sqlx & " and plantskus.sku = skumast.sku"
    sqlx = sqlx & " order by plantskus.sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds(0) & Chr$(9) & ds(1) & Chr$(9)
            sqlx = sqlx & " " & ds(2) & " " & ds(3) & Chr$(9)
            sqlx = sqlx & ds(4) & Chr$(9) & ds(5) & Chr$(9)
            sqlx = sqlx & ds(6) & Chr$(9) & ds(7)
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "ID|^SKU|Product|^Low Qty|^Out Qty|^Low|^Out"
    Grid1.ColWidth(0) = 1
    Grid1.ColWidth(1) = 700: Grid1.ColWidth(2) = 4200
    Grid1.ColWidth(3) = 1000: Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 700: Grid1.ColWidth(6) = 700
    If Grid1.Rows > 1 Then
        Grid1.Row = 1: Grid1.Col = 3
        'Grid1.FixedCols = 3: Grid1.FixedRows = 1
    End If
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

Private Sub Combo1_Click()
    Call refresh_grid
End Sub

Private Sub Command1_Click()            'Add SKU
    Dim msku As String, ds As adodb.Recordset, sqlx As String
    Dim mdesc As String, mid As Long, i As Integer, k As Integer, pkey As Long
    On Error GoTo vberror
    msku = InputBox$("SKU #", "Add SKU to " & Combo1, "777")
    If Len(msku) = 0 Then Exit Sub
    sqlx = "select * from skumast where sku = '" & msku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "SKU " & msku & " not found in SKUMAST.", vbOKOnly, "Sorry can't add"
        ds.Close
        Exit Sub
    End If
    mdesc = ds!fgunit & " " & ds!fgdesc
    ds.Close
    sqlx = "Select * from plantskus where plant = " & Val(Left$(Combo1, 3))
    sqlx = sqlx & " and sku = '" & msku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        MsgBox "SKU " & msku & " already found in " & Combo1 & " list.", vbOKOnly, "Add aborted"
        ds.Close
        Exit Sub
    End If
    pkey = wd_seq("Plantskus", Form1.shipdb)
    sqlx = "Insert into plantskus (id, plant, sku, lowstk, outstk, lowflag, outflag) Values (" & pkey
    sqlx = sqlx & ", " & Val(Left$(Combo1, 3))
    sqlx = sqlx & ", '" & msku & "'"
    sqlx = sqlx & ", 4, 2, 'N', 'N')"
    Sdb.Execute sqlx
    mid = pkey
    ds.Close
    Grid1.AddItem " "
    k = Grid1.Rows - 1
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 1) > msku Then
            k = i: Exit For
        End If
    Next i
    For i = k To Grid1.Rows - 2
        Grid1.TextMatrix(k + 1, 0) = Grid1.TextMatrix(k, 0)
        Grid1.TextMatrix(k + 1, 1) = Grid1.TextMatrix(k, 1)
        Grid1.TextMatrix(k + 1, 2) = Grid1.TextMatrix(k, 2)
        Grid1.TextMatrix(k + 1, 3) = Grid1.TextMatrix(k, 3)
        Grid1.TextMatrix(k + 1, 4) = Grid1.TextMatrix(k, 4)
        Grid1.TextMatrix(k + 1, 5) = Grid1.TextMatrix(k, 5)
        Grid1.TextMatrix(k + 1, 6) = Grid1.TextMatrix(k, 6)
    Next i
    Grid1.TextMatrix(k, 0) = mid
    Grid1.TextMatrix(k, 1) = msku
    Grid1.TextMatrix(k, 2) = " " & mdesc
    Grid1.TextMatrix(k, 3) = "4"
    Grid1.TextMatrix(k, 4) = "2"
    Grid1.TextMatrix(k, 5) = "N"
    Grid1.TextMatrix(k, 6) = "N"
    Grid1.Row = k
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command1_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command1_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command2_Click()            'Drop SKU
    Dim sqlx As String
    On Error GoTo vberror
    If MsgBox("Drop " & Grid1.TextMatrix(Grid1.Row, 2) & " from " & Combo1 & " list?", vbOKCancel, "Are you sure?") = vbCancel Then Exit Sub
    Sdb.Execute "Delete from plantskus where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
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

Private Sub delrec_Click()
    Command2_Click
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Plantskus.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            If Form1.FrmGrid.TextMatrix(i, 0) = "plantskus" Then
                Form1.FrmGrid.TextMatrix(i, 1) = Plantskus.Top
                Form1.FrmGrid.TextMatrix(i, 2) = Plantskus.Left
                Form1.FrmGrid.TextMatrix(i, 3) = Plantskus.Height
                Form1.FrmGrid.TextMatrix(i, 4) = Plantskus.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Plantskus.ActiveControl.Name = "Grid1" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command1_Click
        If KeyCode = 46 Or KeyCode = 120 Then Call Command2_Click
    End If
End Sub

Private Sub Form_Load()
    Dim ds As adodb.Recordset, sqlx As String
    Dim i As Integer
    On Error GoTo vberror
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    For i = 1 To Form1.FrmGrid.Rows - 1
        If Form1.FrmGrid.TextMatrix(i, 0) = "plantskus" Then
            Plantskus.Top = Val(Form1.FrmGrid.TextMatrix(i, 1))
            Plantskus.Left = Val(Form1.FrmGrid.TextMatrix(i, 2))
            Plantskus.Height = Val(Form1.FrmGrid.TextMatrix(i, 3))
            Plantskus.Width = Val(Form1.FrmGrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    sqlx = "select plant,plantname from plants order by plant"
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
    Grid1.Width = Me.Width - 110
    If Plantskus.Height > 2000 Then
        Grid1.Height = Plantskus.Height - 950 '750
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim x As Integer
    Dim sqlx As String
    On Error GoTo vberror
    If KeyAscii = 8 Then
        If Len(Grid1.Text) > 1 Then
            Grid1.Text = Left$(Grid1.Text, Len(Grid1.Text) - 1)
        Else
            Grid1.Text = "0"
        End If
    End If
    If KeyAscii = 45 Then
        If Left(Grid1.Text, 1) = "-" Then
            Grid1.Text = Abs(Val(Grid1.Text))
        Else
            Grid1.Text = "-" & Grid1.Text
            Grid1.Text = Val(Grid1.Text)
        End If
    End If
    If KeyAscii > 47 And KeyAscii < 58 Then
        Grid1.Text = Grid1.Text & Chr$(KeyAscii)
        Grid1.Text = Val(Grid1.Text)
    End If
    If Grid1.Col = 5 Then
        If Grid1.Text = "Y" Then
            Grid1.Text = "N"
            sqlx = "Set lowflag = 'N'"
        Else
            Grid1.Text = "Y"
            sqlx = "Set lowflag = 'Y'"
        End If
    End If
    If Grid1.Col = 6 Then
        If Grid1.Text = "Y" Then
            Grid1.Text = "N"
            sqlx = "Set outflag = 'N'"
        Else
            Grid1.Text = "Y"
            sqlx = "Set outflag = 'Y'"
        End If
    End If
    If Grid1.Col = 3 Then
        Grid1.Text = Val(Grid1.Text)
        sqlx = "Set lowstk = " & Grid1.Text
    End If
    If Grid1.Col = 4 Then
        Grid1.Text = Val(Grid1.Text)
        sqlx = "Set outstk = " & Grid1.Text
    End If
    x = Grid1.Col
    Grid1.Col = 0
    sqlx = "Update plantskus " & sqlx & " Where id = " & Grid1.Text
    Grid1.Col = x
    Sdb.Execute sqlx
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "grid1_keypress", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " grid1_keypress - Error Number: " & eno
        End
    End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub insrec_Click()
    Command1_Click
End Sub
