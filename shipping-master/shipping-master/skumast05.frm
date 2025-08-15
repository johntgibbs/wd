VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form wdvalists 
   Caption         =   "Value Lists"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   LinkTopic       =   "Form5"
   ScaleHeight     =   7905
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "New Value List"
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New Value"
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Changes"
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   6840
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   6480
      Width           =   5295
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5535
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9763
      _Version        =   327680
      BackColorSel    =   128
      FocusRect       =   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "List Name"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "wdvalists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub new_value_list()                'jv051617
    Dim ds As adodb.Recordset, s As String, nc As String, i As Integer, zid As Long
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    nc = InputBox("New Value List Name:", "New value list...", "vlist")
    If Len(nc) = 0 Then Exit Sub
    zid = wd_seq("Valuelists", Form1.shipdb)
    s = "Insert into valuelists (id, listname, listreturn, listdisplay) values (" & zid & ", '" & nc & "', ' ', ' ')"
    Sdb.Execute s
    'MsgBox s
    Combo1.AddItem nc
    Combo1.ListIndex = Combo1.ListCount - 1
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "New Value_List", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " New Value List - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_grid()
    Dim ds As adodb.Recordset, s As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 3
    s = "select * from valuelists where listname = '" & Combo1 & "'"
    s = s & " order by listreturn"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(0) & Chr(9) & ds(2) & Chr(9) & ds(3)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^ID|<Return Value|<Display Value"
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 3000
    Grid1.Row = 1
    Grid1_RowColChange
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_list()
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    Combo1.Clear
    s = "select distinct listname from valuelists order by listname"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds(0)
            ds.MoveNext
        Loop
    End If
    ds.Close
    Combo1.ListIndex = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "refresh_list", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_list - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo1_Click()
    refresh_grid
End Sub

Private Sub Command1_Click()
    Dim ds As adodb.Recordset, s As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    If Grid1.Row = 0 Then Exit Sub
    If Len(Grid1.TextMatrix(Grid1.Row, 1)) = 0 Or Len(Grid1.TextMatrix(Grid1.Row, 2)) = 0 Then
        MsgBox "Null values are not allowed within value lists.", vbInformation + vbOKOnly, "try again.."
        Exit Sub
    End If
    Screen.MousePointer = 11
    s = "select * from valuelists where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update valuelists set listreturn = '" & Grid1.TextMatrix(Grid1.Row, 1) & "'"
        s = s & ", listdisplay = '" & Grid1.TextMatrix(Grid1.Row, 2) & "'"
        s = s & " Where id = " & ds!id
        Sdb.Execute s
    End If
    ds.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "Save changes_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " Save changes_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command2_Click()
    Dim ds As adodb.Recordset, s As String, nc As String, i As Integer, zid As Long
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    zid = wd_seq("Valuelists", Form1.shipdb)
    s = "Insert into valuelists (id, listname, listreturn, listdisplay) values (" & zid & ", '" & Combo1 & "', ' ', ' ')"
    Sdb.Execute s
    Grid1.AddItem zid 'nc
    Grid1.Row = Grid1.Rows - 1
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "New Value_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " New Value_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command3_Click()
    new_value_list
End Sub

Private Sub Form_Load()
    refresh_list
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    i = Grid1.Col
    If KeyAscii > 31 And KeyAscii < 128 Then
        If i = 1 Then Text1.Text = Text1.Text & Chr(KeyAscii)
        If i = 2 Then Text2.Text = Text2.Text & Chr(KeyAscii)
    End If
    If KeyAscii = 8 Then
        If i = 1 Then
            If Len(Text1.Text) <= 1 Then
                Text1.Text = ""
            Else
                Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
            End If
        End If
        If i = 2 Then
            If Len(Text2.Text) <= 1 Then
                Text2.Text = ""
            Else
                Text2.Text = Left(Text2.Text, Len(Text2.Text) - 1)
            End If
        End If
    End If
End Sub

Private Sub Grid1_RowColChange()
    If Grid1.Row = 0 Then Exit Sub
    Label1.Caption = Grid1.TextMatrix(0, 1): Text1.Text = Grid1.TextMatrix(Grid1.Row, 1)
    Label2.Caption = Grid1.TextMatrix(0, 2): Text2.Text = Grid1.TextMatrix(Grid1.Row, 2)
End Sub

Private Sub Text1_Change()
    If Grid1.Row = 0 Then Exit Sub
    Grid1.TextMatrix(Grid1.Row, 1) = Text1
End Sub

Private Sub Text2_Change()
    If Grid1.Row = 0 Then Exit Sub
    Grid1.TextMatrix(Grid1.Row, 2) = Text2
End Sub
