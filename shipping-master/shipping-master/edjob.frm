VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form edjob 
   Caption         =   "BlueBell Jobbing Accounts"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   4815
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   2990
      _Version        =   327680
      ForeColor       =   4194368
      BackColorFixed  =   12648384
      BackColorSel    =   255
      BackColorBkg    =   -2147483633
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Menu prtmenu 
      Caption         =   "&Print"
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu insrec 
         Caption         =   "Insert Record - F10"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete Record - F9"
      End
      Begin VB.Menu cutfield 
         Caption         =   "Cut Field"
      End
      Begin VB.Menu copyfield 
         Caption         =   "Copy Field"
      End
      Begin VB.Menu pastefield 
         Caption         =   "Paste Field"
      End
   End
End
Attribute VB_Name = "edjob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String

Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 11
    Grid1.FixedCols = 3
    sqlx = "select * from jobbing order by acctdesc"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!id & Chr(9) & ds!branch & Chr(9)
            sqlx = sqlx & ds!account & Chr(9) & ds!acctdesc & Chr(9)
            sqlx = sqlx & ds!terms & Chr(9) & ds!addr1 & Chr(9)
            sqlx = sqlx & ds!addr2 & Chr(9) & ds!addr3 & Chr(9)
            sqlx = sqlx & ds!jzip & Chr(9) & ds!jphone & Chr(9) & ds!schecode
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^ID|^Branch|^Acct|<Customer|<Terms|<Address1|<Address2|<City/State|<ZipCode|<Phone|^SCode"
    Grid1.ColWidth(0) = 1
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 900
    Grid1.ColWidth(3) = 3500
    Grid1.ColWidth(4) = 1500
    Grid1.ColWidth(5) = 3000
    Grid1.ColWidth(6) = 3000
    Grid1.ColWidth(7) = 2500
    Grid1.ColWidth(8) = 1200
    Grid1.ColWidth(9) = 1500
    Grid1.ColWidth(10) = 800
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

Private Sub copyfield_Click()
    If Grid1.Row = 0 Or Grid1.Col < 3 Then Exit Sub
    If Len(Grid1.Text) = 0 Then
        Text1 = " "
    Else
        Text1 = Grid1.Text
    End If
End Sub

Private Sub cutfield_Click()
    If Grid1.Row = 0 Or Grid1.Col < 3 Then Exit Sub
    If Grid1.Col = 3 Then edcell = "acctdesc"
    If Grid1.Col = 4 Then edcell = "terms"
    If Grid1.Col = 5 Then edcell = "addr1"
    If Grid1.Col = 6 Then edcell = "addr2"
    If Grid1.Col = 7 Then edcell = "addr3"
    If Grid1.Col = 8 Then edcell = "jzip"
    If Grid1.Col = 9 Then edcell = "jphone"
    If Grid1.Col = 10 Then edcell = "schecode"
    If Len(edcell) > 0 Then
        Text1 = Grid1.Text
        Grid1.Text = " "
        Call update_item
    End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub insrec_Click()
    Dim sqlx As String
    Dim jbr As String, jacct As String
    Dim i As Integer, nkey As Long
    On Error GoTo vberror
    jbr = InputBox("Branch:", "Branch...", "16")
    If Len(jbr) = 0 Then Exit Sub
    If Val(jbr) = 0 Then Exit Sub
    jacct = InputBox("Account #:", "Account Number...")
    If Len(jacct) = 0 Then Exit Sub
    jacct = Left(jacct, 6)
    For i = 0 To Grid1.Rows - 1
        If jbr = Grid1.TextMatrix(i, 1) And jacct = Grid1.TextMatrix(i, 2) Then
            MsgBox "Account is already on file....", vbOKOnly + vbInformation, "try again..."
            Exit Sub
        End If
    Next i
    nkey = wd_seq("jobbing", Form1.shipdb)
    sqlx = "Insert into jobbing (id, branch, account) Values (" & nkey & ", " & Val(jbr) & ", '" & jacct & "')"
    Sdb.Execute sqlx
    Grid1.AddItem nkey & Chr(9) & jbr & Chr(9) & jacct
    Grid1.Row = Grid1.Rows - 1
    Grid1.TopRow = Grid1.Rows - 1
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
    If MsgBox("Ok to delete " & Grid1.TextMatrix(Grid1.Row, 2) & " from list..", vbYesNo + vbQuestion, "Drop user..") = vbNo Then Exit Sub
    sqlx = "Delete from jobbing where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Sdb.Execute sqlx
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_grid
    End If
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
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    If edcell = "acctdesc" Then Grid1.Text = Left(Grid1.Text, 35)
    If edcell = "terms" Then Grid1.Text = Left(Grid1.Text, 15)
    If edcell = "addr1" Then Grid1.Text = Left(Grid1.Text, 25)
    If edcell = "addr2" Then Grid1.Text = Left(Grid1.Text, 25)
    If edcell = "addr3" Then Grid1.Text = Left(Grid1.Text, 20)
    If edcell = "jzip" Then Grid1.Text = Left(Grid1.Text, 9)
    If edcell = "jphone" Then Grid1.Text = Left(Grid1.Text, 20)
    If edcell = "schecode" Then Grid1.Text = Val(Grid1.Text)
    sqlx = "Update jobbing set " & edcell & " = '" & fixquotes(Grid1.Text) & "'"
    sqlx = sqlx & " Where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Sdb.Execute sqlx
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
    If Me.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
            If Form1.FrmGrid.Text = "jobbing" Then
                Form1.FrmGrid.Col = 1: Form1.FrmGrid.Text = Me.Top
                Form1.FrmGrid.Col = 2: Form1.FrmGrid.Text = Me.Left
                Form1.FrmGrid.Col = 3: Form1.FrmGrid.Text = Me.Height
                Form1.FrmGrid.Col = 4: Form1.FrmGrid.Text = Me.Width
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
    Text1 = " "
    For i = 1 To Form1.FrmGrid.Rows - 1
        Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
        If Form1.FrmGrid.Text = "jobbing" Then
            Form1.FrmGrid.Col = 1: Me.Top = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 2: Me.Left = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 3: Me.Height = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 4: Me.Width = Val(Form1.FrmGrid.Text)
            Exit For
        End If
    Next i
    Grid1.Font = "Arial": Grid1.FontSize = 8: Grid1.FontBold = True
    Call refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    If Me.Height > 1000 Then Grid1.Height = Me.Height - 680
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
    If Grid1.Row = 0 Or Grid1.Col < 3 Then Exit Sub
    If Len(edcell) = 0 Then Grid1.Text = ""
    If Grid1.Col = 3 Then edcell = "acctdesc"
    If Grid1.Col = 4 Then edcell = "terms"
    If Grid1.Col = 5 Then edcell = "addr1"
    If Grid1.Col = 6 Then edcell = "addr2"
    If Grid1.Col = 7 Then edcell = "addr3"
    If Grid1.Col = 8 Then edcell = "jzip"
    If Grid1.Col = 9 Then edcell = "jphone"
    If Grid1.Col = 10 Then edcell = "schecode"
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

Private Sub pastefield_Click()
    If Grid1.Row = 0 Or Grid1.Col < 3 Then Exit Sub
    If Grid1.Col = 3 Then edcell = "acctdesc"
    If Grid1.Col = 4 Then edcell = "terms"
    If Grid1.Col = 5 Then edcell = "addr1"
    If Grid1.Col = 6 Then edcell = "addr2"
    If Grid1.Col = 7 Then edcell = "addr3"
    If Grid1.Col = 8 Then edcell = "jzip"
    If Grid1.Col = 9 Then edcell = "jphone"
    If Grid1.Col = 10 Then edcell = "schecode"
    If Len(edcell) > 0 Then
        If Len(Text1) = 0 Then
            Grid1.Text = " "
        Else
            Grid1.Text = Text1
        End If
        Call update_item
    End If
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    rt = "Jobbing Accounts"
    rh = "W / D"
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Call printflexgrid(Printer, Grid1, rt, rh, rf)
End Sub
