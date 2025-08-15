VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brprods 
   Caption         =   "Branch Special Products"
   ClientHeight    =   4110
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   4110
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Sort "
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   0
      Width           =   2175
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Branch"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "SKU"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6588
      _Version        =   327680
      BackColor       =   16777215
      BackColorFixed  =   8454016
      BackColorSel    =   255
      BackColorBkg    =   -2147483633
      FocusRect       =   0
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu insrec 
         Caption         =   "Insert Record"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete Record"
      End
   End
End
Attribute VB_Name = "brprods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, sqlx As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    sqlx = "select brprods.branch,branchname,brprods.sku,fgunit,fgdesc"
    sqlx = sqlx & " from brprods,branches,skumast"
    sqlx = sqlx & " Where brprods.branch = branches.branch"
    sqlx = sqlx & " and brprods.sku = skumast.sku"
    If Option1 = True Then
        sqlx = sqlx & " Order by brprods.sku, brprods.branch"
    Else
        sqlx = sqlx & " order by brprods.branch, brprods.sku"
    End If
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Grid1.AddItem ds(0) & Chr$(9) & ds(1) & Chr(9) & ds(2) & Chr(9) & ds(3) & " " & ds(4)
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^Code|<Branch|^SKU|<Product"
    Grid1.ColWidth(0) = 600: Grid1.ColWidth(1) = 2600
    Grid1.ColWidth(2) = 600: Grid1.ColWidth(3) = 4000
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

Private Sub Command1_Click()            'Insert Record
    Dim sqlx As String, msku As String, pkey As Long
    Dim ds As ADODB.Recordset, pdesc As String, mbr As String, bdesc As String
    On Error GoTo vberror
    mbr = InputBox("Branch Code:", "Branch Code....", Grid1.TextMatrix(Grid1.Row, 0))
    If Len(mbr) = 0 Then Exit Sub
    sqlx = "select * from branches where branch = " & mbr
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        bdesc = ds!branchname
    Else
        MsgBox "Invalid Branch Code: " & mbr, vbOKOnly + vbExclamation, "Sorry, keep trying..."
        ds.Close
        Exit Sub
    End If
    ds.Close
    msku = InputBox$("SKU #", "Add product to " & bdesc, Grid1.TextMatrix(Grid1.Row, 2))
    If Len(msku) = 0 Then Exit Sub
    sqlx = "select * from skumast where sku = '" & msku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid SKU #", vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    Else
        pdesc = " " & ds!fgunit & " " & ds!fgdesc
    End If
    ds.Close
    sqlx = "select * from brprods where branch = " & mbr
    sqlx = sqlx & " and sku = '" & msku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        MsgBox "SKU " & msku & " already assigned to " & mbr, vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    Else
        pkey = wd_seq("brprods", Form1.shipdb)
        sqlx = "Insert into brprods (id, branch, sku) Values (" & pkey & ", " & Val(mbr) & ", '" & msku & "')"
        Sdb.Execute sqlx
    End If
    ds.Close
    Grid1.AddItem mbr & Chr(9) & bdesc & Chr(9) & msku & Chr(9) & pdesc
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, Command1.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command1_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command2_Click()            'Delete Record
    Dim sqlx As String
    On Error GoTo vberror
    If Grid1.Row = 0 Then Exit Sub
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    If MsgBox("Delete " & Grid1.TextMatrix(Grid1.Row, 3) & " From " & Grid1.TextMatrix(Grid1.Row, 1), vbYesNo + vbQuestion, "Are you sure?") = vbNo Then Exit Sub
    sqlx = "delete from brprods where branch = " & Grid1.TextMatrix(Grid1.Row, 0)
    sqlx = sqlx & " And sku = '" & Grid1.TextMatrix(Grid1.Row, 2) & "'"
    Sdb.Execute sqlx
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_grid
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, Command2.Caption & "_click", Form1.userid)
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
    If brprods.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            If Form1.FrmGrid.TextMatrix(i, 0) = "brprods" Then
                Form1.FrmGrid.TextMatrix(i, 1) = brprods.Top
                Form1.FrmGrid.TextMatrix(i, 2) = brprods.Left
                Form1.FrmGrid.TextMatrix(i, 3) = brprods.Height
                Form1.FrmGrid.TextMatrix(i, 4) = brprods.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.FrmGrid.Rows - 1
        If Form1.FrmGrid.TextMatrix(i, 0) = "brprods" Then
            brprods.Top = Val(Form1.FrmGrid.TextMatrix(i, 1))
            brprods.Left = Val(Form1.FrmGrid.TextMatrix(i, 2))
            brprods.Height = Val(Form1.FrmGrid.TextMatrix(i, 3))
            brprods.Width = Val(Form1.FrmGrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Call refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = brprods.Width - 80
    If brprods.Height > 2000 Then
        Grid1.Height = brprods.Height - 1020
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If brprods.ActiveControl.Name = "Grid1" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command1_Click
        If KeyCode = 46 Or KeyCode = 120 Then Call Command2_Click
    End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub insrec_Click()
    Command1_Click
End Sub

Private Sub Option1_Click()
    Call refresh_grid
End Sub

Private Sub Option2_Click()
    Call refresh_grid
End Sub
