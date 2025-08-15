VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form oplist 
   Caption         =   "Order Pick List"
   ClientHeight    =   5850
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5850
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   10186
      _Version        =   327680
      BackColorFixed  =   12648447
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu insrec 
         Caption         =   "Insert Record"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "oplist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String
Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 3
    sqlx = "select oplist.sku,fgunit,fgdesc,opseq from oplist,skumast"
    sqlx = sqlx & " where oplist.sku = skumast.sku"
    sqlx = sqlx & " Order by opseq"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Grid1.AddItem ds(0) & Chr$(9) & " " & ds(1) & " " & ds(2) & Chr$(9) & ds(3)
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^SKU|<Product|^Seq #"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 4200
    Grid1.ColWidth(2) = 1000
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

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub insrec_Click()
    Dim ds As adodb.Recordset, sqlx As String
    Dim ncode As String, ndesc As String
    On Error GoTo vberror
    ncode = InputBox("SKU:", "Add SKU...")
    If Len(ncode) = 0 Then Exit Sub
    sqlx = "select * from skumast where sku = '" & ncode & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        ndesc = ds!fgunit & " " & ds!fgdesc
    Else
        sqlx = "SKU: " & ncode & " is not found in skumaster list."
        MsgBox sqlx, vbOKOnly, "Sorry, Try again....."
        ds.Close
        Exit Sub
    End If
    ds.Close
    sqlx = "select * from oplist where sku = '" & ncode & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        sqlx = "SKU already in list at: " & ds!opseq
        MsgBox sqlx, vbOKOnly + vbExclamation, "Cannot insert..."
        ds.Close
        Exit Sub
    Else
        sqlx = "Insert into oplist (sku, opseq) Values ('" & ncode & "', 1)"
        Sdb.Execute sqlx
        Grid1.AddItem ncode & Chr(9) & ndesc
        Grid1.Row = Grid1.Rows - 1
        Grid1.TopRow = Grid1.Rows - 1
        nfile = True
    End If
    ds.Close
    'bbsr update
    sqlx = "INSERT INTO OPList (SKU, OPSeq) VALUES ('" & ncode & "', 1)"
    Wdb.Execute sqlx
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "insrec_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " insrec_click - Error Number: " & eno
        End
    End If
End Sub
Private Sub delrec_Click()
    Dim sqlx As String
    On Error GoTo vberror
    If MsgBox("Ok to delete " & Grid1.TextMatrix(Grid1.Row, 1) & " from list..", vbYesNo + vbQuestion, "Drop user..") = vbNo Then Exit Sub
    sqlx = "delete from oplist where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
    Sdb.Execute sqlx
    'bbsr update
    Wdb.Execute sqlx
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_grid
    End If
    nfile = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "delrec_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " delrec_click - Error Number: " & eno
        End
    End If
End Sub
Private Sub update_item()
    Dim sqlx As String
    On Error GoTo vberror
    Grid1.Text = Trim(Grid1.Text)
    If edcell = "seqno" Then
        Grid1.Text = Val(Grid1.Text)
        sqlx = "Update oplist set opseq = " & Grid1.Text & " where sku = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
        Sdb.Execute sqlx
        'bbsr update
        Wdb.Execute sqlx
    End If
    edcell = ""
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "update_item", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " update_item - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Len(edcell) > 0 Then Call update_item
    If oplist.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
            If Form1.FrmGrid.Text = "oplist" Then
                Form1.FrmGrid.Col = 1: Form1.FrmGrid.Text = oplist.Top
                Form1.FrmGrid.Col = 2: Form1.FrmGrid.Text = oplist.Left
                Form1.FrmGrid.Col = 3: Form1.FrmGrid.Text = oplist.Height
                Form1.FrmGrid.Col = 4: Form1.FrmGrid.Text = oplist.Width
                Exit For
            End If
        Next i
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 121 Then Call insrec_Click
    If KeyCode = 120 Then Call delrec_Click
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.FrmGrid.Rows - 1
        Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
        If Form1.FrmGrid.Text = "oplist" Then
            Form1.FrmGrid.Col = 1: oplist.Top = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 2: oplist.Left = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 3: oplist.Height = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 4: oplist.Width = Val(Form1.FrmGrid.Text)
            Exit For
        End If
    Next i
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Call refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 110
    If oplist.Height > 1000 Then Grid1.Height = oplist.Height - 800
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
    If Grid1.Row = 0 Or Grid1.Col < 2 Then Exit Sub
    If Len(edcell) = 0 Then Grid1.Text = ""
    If Grid1.Col = 2 Then edcell = "seqno"
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

