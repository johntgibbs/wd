VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form truckwos 
   Caption         =   "Transport Wos"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form3"
   ScaleHeight     =   7470
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Worksheet"
      Height          =   375
      Left            =   7800
      TabIndex        =   25
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear Date"
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel Run"
      Height          =   375
      Left            =   2040
      TabIndex        =   22
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Record"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6960
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   7080
      TabIndex        =   20
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4080
      TabIndex        =   19
      Text            =   "Combo2"
      Top             =   6480
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1920
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   6480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   6120
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1920
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   5760
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3960
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3375
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5953
      _Version        =   327680
      FocusRect       =   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contents"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   10
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comments"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Startime"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Size"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Trailer #"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Destination"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Origin"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ticket #"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Schedule Date:"
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
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "truckwos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim outfile As Boolean

Private Sub refresh_mlists()
    Dim db As DAO.Database, ds As DAO.Recordset, s As String
    Set db = OpenDatabase("s:\wd\test\trucktest.mdb")
    List2.Clear: Combo2.Clear
    s = "select * from valuelists where listname = 'contents'"
    s = s & " and listreturn > ' ' and listdisplay > ' '"
    s = s & " order by listdisplay"
    Set ds = db.OpenRecordset(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            List2.AddItem ds!listreturn
            Combo2.AddItem ds!listdisplay
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    Combo2.AddItem " ": List2.AddItem " "
End Sub

Private Sub sched_file()
    Dim db As DAO.Database, ds As DAO.Recordset, sqlx As String
    'Set db = OpenDatabase(Form1.shipdb)
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, True, Form1.shipdb)
    sqlx = "select loaded,destination,trldate,sum(trlsize)"
    sqlx = sqlx & " from runs where trldate > '" & Format(Now, "m-d-yyyy") & "'"
    sqlx = sqlx & " group by loaded,destination,trldate"
    sqlx = sqlx & " order by trldate,destination,loaded"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Open Form1.webdir & "\ordsched.txt" For Output As #1
        Do Until ds.EOF
            sqlx = Format(ds!trldate, "m-dd-yyyy") & ","
            sqlx = sqlx & ds!Destination & ","
            sqlx = sqlx & ds!loaded & ","
            sqlx = sqlx & ds(3)
            Print #1, sqlx
            ds.MoveNext
        Loop
        Close #1
    End If
    ds.Close
    db.Close
End Sub

Private Sub refresh_dates()
    Dim db As DAO.Database, ds As DAO.Recordset, s As String, i As Integer
    Dim addate As Boolean, k As Integer
    Combo1.Clear
    'Set db = OpenDatabase(Form1.shipdb)
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, True, Form1.shipdb)
    s = "select distinct trldate from runs"
    Set ds = db.OpenRecordset(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem Format(ds!trldate, "m-dd-yyyy")
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    s = Format(Now, "m-dd-yyyy")
    If Combo1.ListCount = 0 Then Combo1.AddItem s
    addate = True
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) = s Then addate = False
    Next i
    If addate = True Then Combo1.AddItem s
    'MsgBox Format(Now, "ddd")
    For k = 1 To 3
        s = Format(DateAdd("d", k, Now), "m-dd-yyyy")
        If Format(s, "ddd") <> "Sun" Then
            addate = True
            For i = 0 To Combo1.ListCount - 1
                If Combo1.List(i) = s Then addate = False
            Next i
            If addate = True Then Combo1.AddItem s
        End If
    Next k
End Sub

Private Sub refresh_wonum()
    Dim db As DAO.Database, ds As DAO.Recordset, s As String, i As Integer
    Set db = OpenDatabase("s:\wd\test\trucktest.mdb")
    Text1(1).Text = ""
    Text1(2).Text = ""
    Text1(3).Text = ""
    Text1(4).Text = ""
    Text1(5).Text = ""
    Text1(6).Text = ""
    Text1(7).Text = ""
    s = "select * from truckwo where r12ticket = '" & Text1(0).Text & "'"
    Set ds = db.OpenRecordset(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Text1(1).Text = ds!origin
        Text1(2).Text = ds!Destination
        Text1(3).Text = ds!trlno
        Text1(4).Text = ds!trlsize
        Text1(5).Text = ds!startime
        Text1(6).Text = ds!description
        Text1(7).Text = ds!contents
    End If
    ds.Close: db.Close
    For i = 0 To List2.ListCount - 1
        If List2.List(i) = Text1(7).Text Then
            Combo2.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub refresh_wos()
    Dim db As DAO.Database, ds As DAO.Recordset, s As String
    Dim rb As DAO.Database, rs As DAO.Recordset, q As String
    Dim ss As DAO.Recordset, pname As String, rkey As Long
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    'Set db = OpenDatabase("s:\wd\test\trucktest.mdb")
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, True, Form1.schdb)
    'Set rb = OpenDatabase(Form1.shipdb)
    Set rb = OpenDatabase(mysqldev, dbcdrivernoprompt, True, Form1.shipdb)
    s = "select * from truckwo where wodate = '" & Combo1 & "'"
    s = s & " and wtype in ('Start', 'SameDay', 'Bobtail')"
    'If Form1.plantno = 50 Then s = s & " and origin = 'T10'"
    'If Form1.plantno = 51 Then s = s & " and origin = 'K10'"
    'If Form1.plantno = 52 Then s = s & " and origin = 'A10'"
    's = s & " and ucase(contents) <= 'IC'"
    s = s & " and wostatus <> 'CANC'"
    'MsgBox s
    Set ds = db.OpenRecordset(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            'MsgBox UCase(ds!contents) & "ticket:" & ds!r12ticket & ":" & Len(ds!r12ticket)
            If Left(UCase(ds!contents), 2) = "IC" Then
                If Len(ds!r12ticket) <= 3 Or IsNull(Len(ds!r12ticket)) = True Then
                    q = "select * from runs where id = 0"
                    Set rs = rb.OpenRecordset(q)
                    pname = " "
                    s = "select location from locations where lcode = '" & ds!Destination & "'"
                    'MsgBox s
                    Set ss = db.OpenRecordset(s)
                    If ss.BOF = False Then
                        ss.MoveFirst
                        pname = ss!location
                    End If
                    ss.Close
                    rs.AddNew
                    If ds!origin = "A10" Then rs!loaded = 52
                    If ds!origin = "K10" Then rs!loaded = 51
                    If ds!origin = "T10" Then rs!loaded = 50
                    If ds!Destination = "A10" Then
                        rs!Destination = "52"
                    Else
                        If ds!Destination = "T10" Then
                            rs!Destination = "50"
                        Else
                            If ds!Destination = "K10" Then
                                rs!Destination = "51"
                            Else
                                rs!Destination = Format(Val(ds!Destination), "00")
                            End If
                        End If
                    End If
                    rs!locname = pname
                    If ds!wtype = "Bobtail" Then
                        rs!trlno = "B" & ds!trlno
                    Else
                        rs!trlno = "#" & ds!trlno
                    End If
                    rs!trlsize = ds!trlsize
                    rs!trldate = ds!wodate
                    rs!startime = ds!startime
                    rs!pickup = ds!description
                    If ds!drvpool = "Outside Carrier" Then rs!oc = "*"
                    rkey = rs!id
                    rs.Update
                    rs.Close
                    ds.Edit
                    ds!r12ticket = rkey
                    ds.Update
                Else
                    q = "select * from runs where id = " & Val(ds!r12ticket)
                    MsgBox q
                    Set rs = rb.OpenRecordset(q)
                    If rs.BOF = False Then
                        rs.MoveFirst
                        rs.Edit
                        rs!trlsize = ds!trlsize
                        rs!startime = ds!startime
                        rs!pickup = ds!description
                        rs.Update
                    End If
                    rs.Close
                End If
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    q = "select * from runs where trldate = '" & Combo1 & "'"
    q = q & " order by loaded,startime"
    Set rs = rb.OpenRecordset(q)
    If rs.BOF = False Then
        rs.MoveFirst
        Do Until rs.EOF
            s = rs!id & Chr(9)
            If rs!loaded = "50" Then s = s & "Brenham"
            If rs!loaded = "51" Then s = s & "Broken Arrow"
            If rs!loaded = "52" Then s = s & "Sylacauga"
            s = s & Chr(9)
            's = s & rs!loaded & Chr(9)
            's = s & rs!destination & Chr(9)
            s = s & rs!locname & Chr(9)
            s = s & rs!trlno & Chr(9)
            s = s & rs!trlsize & Chr(9)
            's = s & rs!trldate & Chr(9)
            s = s & rs!startime & Chr(9)
            s = s & rs!pickup & Chr(9)
            s = s & rs!oc
            Grid1.AddItem s
            rs.MoveNext
        Loop
    End If
    rs.Close: rb.Close
    Grid1.FormatString = "<ID|<Plant|<Destination|^#|^Size|^Start|<Comments|^OC"
    Grid1.ColWidth(0) = 6000
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 2000
    Grid1.ColWidth(3) = 600
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 4500
    Grid1.ColWidth(7) = 600
    Screen.MousePointer = 0
    If Grid1.Rows > 1 Then Call Grid1_RowColChange
End Sub

Private Sub Combo1_Click()
    refresh_wos
End Sub

Private Sub Combo2_Click()
    List2.ListIndex = Combo2.ListIndex
End Sub

Private Sub Command1_Click()
    Dim db As DAO.Database, ds As DAO.Recordset, s As String
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, True, Form1.schdb)
    'Set db = OpenDatabase("s:\wd\test\trucktest.mdb")
    s = "select * from truckwo where r12ticket = '" & Text1(0).Text & "'"
    Set ds = db.OpenRecordset(s)
    If ds.BOF = False Then
        ds.MoveFirst
        ds.Edit
        ds!trlsize = Val(Text1(4).Text)
        ds!startime = Text1(5).Text
        ds!Comments = Text1(6).Text
        ds!contents = Text1(7).Text
        ds.Update
    End If
    ds.Close: db.Close
    Set db = OpenDatabase("s:\wd\test\trucktest.mdb")
    s = "select * from runs where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = db.OpenRecordset(s)
    If ds.BOF = False Then
        ds.MoveFirst
        ds.Edit
        ds!trlsize = Val(Grid1.TextMatrix(Grid1.Row, 4))
        ds!startime = Grid1.TextMatrix(Grid1.Row, 5)
        ds!pickup = Grid1.TextMatrix(Grid1.Row, 6)
        ds!updatedby = "shipping" 'Form1.UserId
        ds!lastchange = Format(Now, "m-d-yyyy h:mm am/pm")
        ds.Update
    End If
    ds.Close: db.Close
    outfile = True
End Sub

Private Sub Command2_Click()
    Dim db As DAO.Database, sqlx As String
    sqlx = Grid1.TextMatrix(Grid1.Row, 2)
    sqlx = sqlx & " " & Grid1.TextMatrix(Grid1.Row, 3)
    If MsgBox("Clear " & sqlx & " on " & Combo1, vbYesNo + vbQuestion, "Are you sure?") = vbNo Then
        Exit Sub
    End If
    Set db = OpenDatabase(Form1.shipdb)
    sqlx = "delete from runs where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    db.Execute sqlx
    db.Close
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Grid1.Rows = 1
    End If
    outfile = True
End Sub

Private Sub Command3_Click()
    Dim i As Integer, ol As String
    Screen.MousePointer = 11
    Call sched_file
    outfile = False
    Printer.FontSize = 12
    Printer.Print "Transport Schedule  " & Combo1
    Printer.FontName = "Courier New"
    Printer.FontSize = 8
    Printer.Print " "
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 7) > " " Then
            ol = "OC_______  "
        Else
            ol = "_________  "
        End If
        ol = ol & Grid1.TextMatrix(i, 1) & Space$(20 - Len(Grid1.TextMatrix(i, 1)))
        ol = ol & Grid1.TextMatrix(i, 2) & Space$(20 - Len(Grid1.TextMatrix(i, 2)))
        ol = ol & Grid1.TextMatrix(i, 3) & Space$(5 - Len(Grid1.TextMatrix(i, 3)))
        ol = ol & Grid1.TextMatrix(i, 4) & Space$(5 - Len(Grid1.TextMatrix(i, 4)))
        ol = ol & Grid1.TextMatrix(i, 5) & Space$(15 - Len(Grid1.TextMatrix(i, 5)))
        ol = ol & Grid1.TextMatrix(i, 6)
        Printer.Print ol
        Printer.Print " "
    Next i
    Printer.EndDoc
    Screen.MousePointer = 0
End Sub

Private Sub Command4_Click()
    Dim db As DAO.Database, sqlx As String
    If MsgBox("Clear schedule for " & Combo1, vbOKCancel, "Are you sure?") = vbCancel Then
        Exit Sub
    End If
    'Set db = OpenDatabase(Form1.shipdb)
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, True, Form1.shipdb)
    sqlx = "delete from runs where trldate = '" & Combo1 & "'"
    db.Execute sqlx
    Combo1.RemoveItem Combo1.ListIndex
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
    db.Close
    outfile = True
End Sub

Private Sub Command5_Click()
    Dim cfile As String, i As Integer, x
    If Grid1.Rows = 1 Then Exit Sub
    cfile = Form1.tempdir & "\aschedwrk.csv"
    Open cfile For Output As #1
    For i = 1 To Grid1.Rows - 1
        Write #1, Combo1;                           'Date
        Write #1, Grid1.TextMatrix(i, 1);           'Plant
        Write #1, Grid1.TextMatrix(i, 2);           'Branch
        Write #1, Grid1.TextMatrix(i, 3) & "    ";           'Trailer #
        Write #1, Grid1.TextMatrix(i, 4) & "    ";           'Size
        Write #1, Grid1.TextMatrix(i, 5);           'Start
        Write #1, Grid1.TextMatrix(i, 6);           'Notes
        Write #1, "  " & Grid1.TextMatrix(i, 7) & "  "      'Oc
    Next i
    Close #1
    MsgBox "Created file at: " & cfile, vbInformation + vbOKOnly, "Export completed...."
    'x = Shell("notepad.exe " & cfile, vbNormalFocus)
End Sub

Private Sub Form_Load()
    outfile = False
    refresh_dates
    refresh_mlists
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If outfile Then Call sched_file
End Sub

Private Sub Grid1_RowColChange()
    Text1(0).Text = Grid1.TextMatrix(Grid1.Row, 0)
End Sub

Private Sub List2_Click()
    Text1(7).Text = List2
End Sub

Private Sub Text1_Change(Index As Integer)
    If Index = 0 Then refresh_wonum
    If Index = 4 Then Grid1.TextMatrix(Grid1.Row, 4) = Text1(4).Text
    If Index = 5 Then Grid1.TextMatrix(Grid1.Row, 5) = Text1(5).Text
    If Index = 6 Then Grid1.TextMatrix(Grid1.Row, 6) = Text1(6).Text
End Sub
