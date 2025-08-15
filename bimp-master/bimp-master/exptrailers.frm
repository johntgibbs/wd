VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form exptrailers 
   Caption         =   "Send Trailers to Plant"
   ClientHeight    =   9870
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8055
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   14208
      _Version        =   327680
      BackColorFixed  =   16761024
      FocusRect       =   0
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   8880
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   5160
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   120
      Width           =   2175
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
      Left            =   1440
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label ycolor 
      BackColor       =   &H0080FFFF&
      Caption         =   "ycolor"
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
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
      Left            =   7560
      TabIndex        =   6
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Plant:"
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
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Ship Date:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu postmenu 
      Caption         =   "Post"
      Begin VB.Menu postrun 
         Caption         =   "Post Run"
      End
      Begin VB.Menu postall 
         Caption         =   "Post All"
      End
   End
End
Attribute VB_Name = "exptrailers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function plant_seq(tbname As String) As Long
    Dim sSql As String
    Dim i As Long
    Dim ds As ADODB.Recordset
    sSql = "Select sequence_id From sequences where seq = '" & tbname & "'"
    If Combo2 = "A10" Then
        Set ds = a10shipdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst
            i = ds!sequence_id + 1
            sSql = "Update sequences Set sequence_id = " & i & " Where seq = '" & tbname & "'"
            a10shipdb.Execute (sSql)
        Else
            i = 100
            sSql = "Insert Into sequences (sequence_id, seq) Value (" & i & ",'" & tbname & "')"
            a10shipdb.Execute (sSql)
        End If
        ds.Close
    End If
    If Combo2 = "K10" Then
        Set ds = k10shipdb.Execute(sSql)
        If ds.BOF = False Then
            ds.MoveFirst
            i = ds!sequence_id + 1
            sSql = "Update sequences Set sequence_id = " & i & " Where seq = '" & tbname & "'"
            k10shipdb.Execute (sSql)
        Else
            i = 100
            sSql = "Insert Into sequences (sequence_id, seq) Value (" & i & ",'" & tbname & "')"
            k10shipdb.Execute (sSql)
        End If
        ds.Close
    End If
    plant_seq = i
End Function

Private Sub post_run(rkey As String)
    Dim ds As ADODB.Recordset, s As String, gcode As String, z As Long
    s = "Delete from trailers where runid = " & rkey
    If Combo2 = "A10" Then
        'MsgBox s, vbOKOnly, "Sylacauga"
        a10shipdb.Execute s
    End If
    If Combo2 = "K10" Then
        'MsgBox s, vbOKOnly, "Broken Arrow"
        k10shipdb.Execute s
    End If
    s = "select * from trailers where runid = " & rkey
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            z = plant_seq("Trailers")
            gcode = "T" & Format(ds!shipdate, "MM") & Format(ds!branch, "00") & Right(ds!trlno, 1)
            s = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno, sku"
            s = s & ", pallets, wraps, units, whs_num, pb_flag, ra_flag) Values ("
            s = s & z
            s = s & ", " & ds!runid
            s = s & ", '" & gcode & "'"
            s = s & ", " & ds!plant
            s = s & ", " & ds!branch
            s = s & ", '" & ds!account & "'"
            s = s & ", '" & Format(ds!shipdate, "MM-dd-yyyy") & "'"
            s = s & ", '" & ds!trlno & "'"
            s = s & ", '" & ds!sku & "'"
            s = s & ", " & Format(ds!pallets, "0")
            s = s & ", " & Format(ds!wraps, "0")
            s = s & ", " & Format(ds!units, "0")
            s = s & ", " & ds!whs_num
            s = s & ", 'N', 'N')"
            'MsgBox s
            If Combo2 = "A10" Then a10shipdb.Execute s
            If Combo2 = "K10" Then k10shipdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    refresh_grid
End Sub

Private Sub refresh_dates()
    Dim ds As ADODB.Recordset, s As String
    Combo1.Clear
    s = "select shipdate, count(*) from trailers where ra_flag = 'N'"
    s = s & " and plant in (51, 52)"
    s = s & " and shipdate >= '" & Format(Now, "MM-dd-yyyy") & "'"
    s = s & " group by shipdate order by shipdate"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem Format(ds!shipdate, "MM-dd-yyyy")
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, s As String, i As Integer
    If Combo2 <> "A10" And Combo2 <> "K10" Then Exit Sub            'jv010417
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    s = "select runid, branch, account, trlno, count(*) from trailers where shipdate = '" & Combo1 & "'"
    If Combo2 = "K10" Then s = s & " and plant = 51"
    If Combo2 = "A10" Then s = s & " and plant = 52"
    s = s & " group by runid, branch, account, trlno order by branch, trlno"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!runid & Chr(9)
            s = s & Format(ds!branch, "000") & "-" & branchrec(ds!branch).branchname & Chr(9)
            s = s & ds!account & Chr(9)
            s = s & ds!trlno & Chr(9)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            s = "select id from trailers where runid = " & Grid1.TextMatrix(i, 0)
            If Combo2 = "A10" Then
                Set ds = a10shipdb.Execute(s)
                If ds.BOF = False Then Grid1.TextMatrix(i, 4) = "X"
                ds.Close
            End If
            If Combo2 = "K10" Then
                Set ds = k10shipdb.Execute(s)
                If ds.BOF = False Then Grid1.TextMatrix(i, 4) = "X"
                ds.Close
            End If
            If Grid1.TextMatrix(i, 4) = "X" Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ycolor.BackColor
            End If
        Next i
        Grid1.Row = 1
    End If
    Grid1.FormatString = "^Ticket|<Branch|^Account|^Trailer #|^Sent"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub


Private Sub Combo1_Click()
    refresh_grid
End Sub

Private Sub Combo2_Click()
    List2.ListIndex = Combo2.ListIndex
    Label3 = List2
    refresh_grid
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    'Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    Me.Left = plandist.Width - Me.Width
    'Set a10shipdb = CreateObject("ADODB.Connection")
    'a10shipdb.Open a10ship
    'Set k10shipdb = CreateObject("ADODB.Connection")
    'k10shipdb.Open k10ship
    refresh_dates
    Combo2.Clear: List2.Clear
    If plant_server_status("A10") = True Then                   'jv010417
        Combo2.AddItem "A10": List2.AddItem "Sylacauga"
        Set a10shipdb = CreateObject("ADODB.Connection")
        a10shipdb.Open a10ship
    Else
        MsgBox "Sylacauga Server is not online for posting.", vbOKOnly + vbInformation, "Sylacauga offline.."
    End If
    If plant_server_status("K10") = True Then                   'jv010417
        Combo2.AddItem "K10": List2.AddItem "Broken Arrow"
        Set k10shipdb = CreateObject("ADODB.Connection")
        k10shipdb.Open k10ship
    Else
        MsgBox "Broken Arrow Server in not online for posting.", vbOKOnly + vbInformation, "Broken Arrow offline.."
    End If
    Combo1.ListIndex = 0
    If Combo2.ListCount > 0 Then                                'jv010417
        Combo2.ListIndex = 0
    End If
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 4.5)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If plant_server_status("A10") = True Then                   'jv010417
        a10shipdb.Close
    End If
    If plant_server_status("K10") = True Then                   'jv010417
        k10shipdb.Close
    End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu postmenu
End Sub

Private Sub postall_Click()
    Dim i As Integer
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 0)) > 0 Then Call post_run(Grid1.TextMatrix(i, 0))
    Next i
End Sub

Private Sub postrun_Click()
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) > 0 Then Call post_run(Grid1.TextMatrix(Grid1.Row, 0))
End Sub
