VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Plantbranch 
   Caption         =   "Plant Branches"
   ClientHeight    =   4665
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   4665
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   7646
      _Version        =   327680
      BackColorFixed  =   12648447
      BackColorBkg    =   -2147483633
      FocusRect       =   0
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
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu insrec 
         Caption         =   "Add Branch"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete Branch"
      End
   End
End
Attribute VB_Name = "Plantbranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 2
    sqlx = "select branch,branchname from branches"
    sqlx = sqlx & " where branch in (select branch from plantbranch"
    sqlx = sqlx & " where plant = " & Left$(Combo1, 3) & ")"
    sqlx = sqlx & " order by branch"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Grid1.AddItem ds(0) & Chr$(9) & ds(1)
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^Branch|<Description"
    Grid1.ColWidth(0) = 900: Grid1.ColWidth(1) = 3000
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

Private Sub Command1_Click()            'Add Branch
    Dim ds As adodb.Recordset, sqlx As String, mbr As String, mname As String, pkey As Long
    On Error GoTo vberror
    mbr = InputBox$("Branch #", "Add Branch to " & Combo1, "01")
    If Len(mbr) = 0 Then Exit Sub
    sqlx = "select * from branches where branch = " & mbr
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid Branch # " & mbr, vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    Else
        mname = ds!branchname
    End If
    ds.Close
    sqlx = "select * from plantbranch where plant = " & Left$(Combo1, 3)
    sqlx = sqlx & " and branch = " & mbr
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        MsgBox "Branch " & mbr & " already assigned to " & Combo1, vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    End If
    ds.Close
    pkey = wd_seq("plantbranch", Form1.shipdb)
    sqlx = "Insert into plantbranch(id, plant,branch) values (" & pkey & ", " & Left$(Combo1, 3)
    sqlx = sqlx & "," & mbr & ")"
    Sdb.Execute sqlx
    Grid1.AddItem mbr & Chr(9) & mname
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

Private Sub Command2_Click()            'Delete Branch
    Dim sqlx As String
    On Error GoTo vberror
    If Grid1.Row = 0 Then Exit Sub
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    If MsgBox("Delete " & Grid1.TextMatrix(Grid1.Row, 1) & " From " & Combo1, vbOKCancel, "Are you sure?") = vbCancel Then Exit Sub
    sqlx = "delete from plantbranch where plant = " & Left$(Combo1, 3)
    sqlx = sqlx & " And branch = " & Grid1.TextMatrix(Grid1.Row, 0)
    Sdb.Execute sqlx
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
    If Plantbranch.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            If Form1.FrmGrid.TextMatrix(i, 0) = "plantbranch" Then
                Form1.FrmGrid.TextMatrix(i, 1) = Plantbranch.Top
                Form1.FrmGrid.TextMatrix(i, 2) = Plantbranch.Left
                Form1.FrmGrid.TextMatrix(i, 3) = Plantbranch.Height
                Form1.FrmGrid.TextMatrix(i, 4) = Plantbranch.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Plantbranch.ActiveControl.Name = "Grid1" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command1_Click
        If KeyCode = 46 Or KeyCode = 120 Then Call Command2_Click
    End If
End Sub

Private Sub Form_Load()
    Dim ds As adodb.Recordset, sqlx As String, i As Integer
    On Error GoTo vberror
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    For i = 1 To Form1.FrmGrid.Rows - 1
        If Form1.FrmGrid.TextMatrix(i, 0) = "plantbranch" Then
            Plantbranch.Top = Val(Form1.FrmGrid.TextMatrix(i, 1))
            Plantbranch.Left = Val(Form1.FrmGrid.TextMatrix(i, 2))
            Plantbranch.Height = Val(Form1.FrmGrid.TextMatrix(i, 3))
            Plantbranch.Width = Val(Form1.FrmGrid.TextMatrix(i, 4))
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
    If Plantbranch.Height > 2000 Then
        Grid1.Height = Plantbranch.Height - 1050 '750
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_Click()
    Grid1.Col = 0: Grid1.ColSel = 1
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub insrec_Click()
    Command1_Click
End Sub
