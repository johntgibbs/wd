VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "WMS User Menus"
   ClientHeight    =   12885
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13815
   LinkTopic       =   "Form7"
   ScaleHeight     =   12885
   ScaleWidth      =   13815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Apply Menu"
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
      Left            =   8640
      TabIndex        =   14
      Top             =   120
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Refresh Users"
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
      Left            =   11880
      TabIndex        =   13
      Top             =   480
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   11295
      Left            =   7200
      TabIndex        =   11
      Top             =   960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   19923
      _Version        =   327680
      BackColorFixed  =   16777152
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   11295
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   19923
      _Version        =   327680
      BackColorFixed  =   12648447
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
   End
   Begin VB.ListBox List3 
      Height          =   1035
      Left            =   12360
      TabIndex        =   7
      Top             =   12360
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   10800
      TabIndex        =   6
      Text            =   "Combo3"
      Top             =   12360
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   9000
      TabIndex        =   5
      Top             =   12360
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5160
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   7080
      TabIndex        =   3
      Top             =   12360
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label appkey 
      Caption         =   "..."
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
      Left            =   7200
      TabIndex        =   12
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Left            =   3600
      TabIndex        =   10
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
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
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label listlab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Org ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label listlab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu addmenu 
         Caption         =   "Add Menu"
      End
      Begin VB.Menu dropmenu 
         Caption         =   "Drop Menu"
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_lists()
    Dim ds As ADODB.Recordset, s As String
    Combo1.Clear: List1.Clear
    Combo2.Clear: List2.Clear
    Combo3.Clear: List3.Clear
    s = "select listreturn, listdisplay from valuelists where listname = 'wmsuser' order by listdisplay"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds!listreturn
            List1.AddItem ds!listdisplay
            ds.MoveNext
        Loop
    End If
    ds.Close
    's = "select listreturn, listdisplay from valuelists where listname = 'oraorg' order by listdisplay"
    'Set ds = Wdb.Execute(s)
    'If ds.BOF = False Then
    '    ds.MoveFirst
    '    Do Until ds.EOF
    '        Combo2.AddItem ds!listreturn
    '        List2.AddItem ds!listdisplay
    '        ds.MoveNext
    '    Loop
    'End If
    'ds.Close
    Combo2.AddItem "500": List2.AddItem "Brenham"
    Combo2.AddItem "501": List2.AddItem "Broken Arrow"
    Combo2.AddItem "502": List2.AddItem "Sylacauga"
    s = "select listreturn, listdisplay from valuelists where listname = 'wmsmenu' order by listdisplay"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo3.AddItem ds!listreturn
            List3.AddItem ds!listdisplay
            ds.MoveNext
        Loop
    End If
    ds.Close
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
End Sub

Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, s As String, i As Integer
    Grid1.Redraw = False: Check1.Value = 0
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    For i = 0 To Combo3.ListCount - 1
        s = " " & Chr(9) & Combo2 & Chr(9) & Combo3.List(i) & Chr(9) & List3.List(i)
        Grid1.AddItem s
    Next i
    s = "select * from usermenus where userid = '" & Combo1 & "'"
    s = s & " and orgid = '" & Combo2 & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 0 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 2) = ds!menuname Then
                    Grid1.TextMatrix(i, 0) = ds!id
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FillStyle = flexFillRepeat
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 0)) = 0 Then
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellBackColor = Label1.BackColor
        End If
    Next i
    Grid1.Row = 1: Grid1.Col = 1
    Grid1.FormatString = "^ID|^Org|<Menu|<Name"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 3400
    Grid1.Redraw = True
    Grid1_RowColChange
    Check1.Value = 1
    appkey_Change
End Sub

Private Sub refresh_grid2()
    Dim s As String, ds As ADODB.Recordset, i As Integer, nam As String
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 4
    s = "select * from usermenus where menuname = '" & Combo3 & "' and orgid = '" & Combo2 & "' order by userid"
    'MsgBox s
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            nam = " "
            For i = 0 To Combo1.ListCount - 1
                If Combo1.List(i) = ds!userid Then
                    nam = List1.List(i)
                    Exit For
                End If
            Next i
            s = ds!id & Chr(9) & ds!userid & Chr(9) & nam & Chr(9) & ds!orgid
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid2.FormatString = "^ID|<UserID|<Name|^OrgID"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1400
    Grid2.ColWidth(2) = 2400
    Grid2.ColWidth(3) = 1000
End Sub

Private Sub addmenu_Click()
    Dim s As String, i As String, ds As ADODB.Recordset, zid As Long
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 0)) > 0 Then Exit Sub
    s = "select id from usermenus where userid = '...' order by id"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "update usermenus set userid = '" & Combo1 & "', orgid = '" & Combo2 & "'"
        s = s & ", menuname = '" & Grid1.TextMatrix(i, 2) & "' where id = " & ds!id
        'MsgBox s
        Wdb.Execute s
        Grid1.TextMatrix(i, 0) = ds!id
    Else
        zid = wd_seq("Usermenus")
        s = "Insert into usermenus (id, userid, orgid, menuname) Values (" & zid
        s = s & ", '" & Combo1 & "'"
        s = s & ", '" & Combo2 & "'"
        s = s & ", '" & Grid1.TextMatrix(i, 2) & "')"
        'MsgBox s
        Wdb.Execute s
        Grid1.TextMatrix(i, 0) = zid
    End If
    ds.Close
    Grid1.Row = i: Grid1.RowSel = i
    Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
    Grid1.CellBackColor = Grid1.BackColor
    Grid1.Col = 1
    Grid1_RowColChange
End Sub

Private Sub appkey_Change()
    If Check1.Value = 1 Then refresh_grid2
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Combo2_Click()
    List2.ListIndex = Combo2.ListIndex
End Sub

Private Sub Combo3_Click()
    List3.ListIndex = Combo3.ListIndex
End Sub

Private Sub Command1_Click()
    Call Form1.menu_build(Combo1)
End Sub

Private Sub dropmenu_Click()
    Dim s As String, i As String
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 0)) = 0 Then Exit Sub
    s = "update usermenus set userid = '...', orgid = '...', menuname = '...' where id = "
    s = s & Grid1.TextMatrix(i, 0)
    'MsgBox s
    Wdb.Execute s
    Grid1.TextMatrix(i, 0) = " "
    Grid1.Row = i: Grid1.RowSel = i
    Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
    Grid1.CellBackColor = Label1.BackColor
    Grid1.Col = 1
    Grid1_RowColChange
End Sub

Private Sub Form_Load()
    refresh_lists
    'refresh_grid
End Sub

Private Sub Form_Resize()
    'Grid1.Width = Me.Width - 100
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid1_RowColChange()
    Dim s As String, i As Integer
    addmenu.Enabled = False
    dropmenu.Enabled = False
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then
        addmenu.Enabled = True
    Else
        dropmenu.Enabled = True
    End If
    s = Grid1.TextMatrix(Grid1.Row, 2)
    For i = 0 To Combo3.ListCount - 1
        If Combo3.List(i) = s Then
            Combo3.ListIndex = i
            appkey = List3.List(i)
            Exit For
        End If
    Next i
End Sub

Private Sub List1_Click()
    Label1.Caption = List1
    refresh_grid
End Sub

Private Sub List2_Click()
    Label2.Caption = List2
    refresh_grid
End Sub
