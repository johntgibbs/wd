VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Shipuser 
   Caption         =   "W&D User Accounts"
   ClientHeight    =   9960
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9960
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
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
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7223
      _Version        =   327680
      ForeColor       =   128
      BackColorFixed  =   12648384
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
      FocusRect       =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Special Branch Codes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu insrec 
         Caption         =   "Insert User"
      End
      Begin VB.Menu delrec 
         Caption         =   "Drop User"
      End
   End
End
Attribute VB_Name = "Shipuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String, nfile As Boolean

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub insrec_Click()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    ncode = InputBox("User Computer Code:", "Add User...")
    If Len(ncode) = 0 Then Exit Sub
    sqlx = "select * from wdusers where usercode = '" & ncode & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        sqlx = "Code already exists for: " & ds!wduserName
        sqlx = sqlx & " at branch: " & ds!branch
        MsgBox sqlx, vbOKOnly + vbExclamation, "Cannot insert..."
    Else
        sqlx = "Insert into wdusers (usercode) Values ('" & ncode & "')"
        Sdb.Execute sqlx
        Grid1.AddItem ncode
        Grid1.Row = Grid1.Rows - 1
        Grid1.TopRow = Grid1.Rows - 1
        nfile = True
    End If
    ds.Close
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
    sqlx = "delete from wdusers where usercode = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
    Sdb.Execute sqlx
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
Private Sub export_users()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    sqlx = "select * from wdusers order by usercode"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        Open Form1.webdir & "\userlist" For Output As #1
        ds.MoveFirst
        Do Until ds.EOF
            Write #1, ds!usercode, Format(ds!branch, "00"), ds!wduserName
            ds.MoveNext
        Loop
        Close #1
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "export_users", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " export_users - Error Number: " & eno
        End
    End If
End Sub

Private Sub update_item()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    sqlx = "select * from wdusers where usercode = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
    Set ds = Sdb.Execute(sqlx)
    ds.MoveFirst
    Grid1.Text = Trim(Grid1.Text)
    If edcell = "branch" Then
        Grid1.Text = Val(Grid1.Text)
        sqlx = "Update wdusers set branch = " & Val(Grid1.Text) & " Where usercode = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
        Sdb.Execute sqlx
    End If
    If edcell = "wdusername" Then
        If Len(Grid1.Text) > 50 Then Grid1.Text = Left(Grid1.Text, 50)
        sqlx = "Update wdusers set wdusername = '" & Grid1.Text & "' Where usercode = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
        Sdb.Execute sqlx
    End If
    ds.Close
    edcell = ""
    nfile = True
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
Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 3
    sqlx = "select * from wdusers order by wdusername"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!usercode & Chr(9)
            sqlx = sqlx & ds!branch & Chr(9)
            sqlx = sqlx & ds!wduserName
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "<Code|^Branch|<Name"
    Grid1.ColWidth(0) = 1800
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 4000
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

Private Sub Form_Deactivate()
    If Len(edcell) > 0 Then Call update_item
    If nfile = True Then Call export_users
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 121 Then Call insrec_Click
    If KeyCode = 120 Then Call delrec_Click
End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "500 - Super User Account"
    Combo1.AddItem "401 - Region 1 Manager"
    Combo1.AddItem "402 - Region 2 Manager"
    Combo1.AddItem "403 - Region 3 Manager"
    Combo1.AddItem "404 - Region 4 Manager"
    Combo1.AddItem "405 - Region 5 Manager"
    Combo1.AddItem "406 - Region 6 Manager"
    Combo1.ListIndex = 0
    nfile = False
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Call refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Shipuser.Width - 80
    If Shipuser.Height > 2000 Then
        Grid1.Height = Shipuser.Height - 1050 '740
    End If
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
    If Grid1.Col = 1 Then edcell = "branch"
    If Grid1.Col = 2 Then edcell = "wdusername"
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
