VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Transched 
   Caption         =   "Transport Schedule"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5655
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2295
      Left            =   0
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   4048
      _Version        =   327680
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Schedule Changes"
      Height          =   375
      Left            =   9120
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Worksheet"
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear Date"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5295
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9340
      _Version        =   327680
      BackColorFixed  =   12648384
      FocusRect       =   0
      HighLight       =   2
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   12240
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   11160
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label pcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Order not received via W/D browser."
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
      Left            =   4080
      TabIndex        =   10
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label ycolor 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label1"
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "Transched"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim outfile As Boolean
Dim edcell As String

Private Sub update_run()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    sqlx = "select * from runs where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If edcell = "locname" Then
            Grid1.Text = Trim(Grid1.Text)
            If Len(Grid1.Text) > 30 Then Grid1.Text = Left(Grid1.Text, 30)
            If Len(Grid1.Text) = 0 Then
                s = "Update runs set locname = ' ' Where id = " & ds!id
                Sdb.Execute s
            Else
                s = "Update runs set locname = '" & Grid1.Text & "' Where id = " & ds!id
                Sdb.Execute s
            End If
        End If
        If edcell = "trlno" Then
            Grid1.Text = Trim(Grid1.Text)
            If Len(Grid1.Text) > 2 Then Grid1.Text = Left(Grid1.Text, 2)
            If Len(Grid1.Text) = 0 Then
                s = "Update runs set trlno = ' ' Where id = " & ds!id
                Sdb.Execute s
            Else
                s = "Update runs set trlno = '" & Grid1.Text & "' Where id = " & ds!id
                Sdb.Execute s
            End If
        End If
        If edcell = "trlsize" Then
            Grid1.Text = Val(Grid1.Text)
            s = "Update runs set trlsize = " & Val(Grid1.Text) & " Where id = " & ds!id
            Sdb.Execute s
        End If
        If edcell = "startime" Then
            If IsDate(Grid1.Text) = False Then
                Grid1.Text = Format(Now, "h:mm am/pm")
                Beep
            Else
                Grid1.Text = Format(Grid1.Text, "h:mm am/pm")
            End If
            s = "Update runs set startime = '" & Format(Grid1.Text, "h:mm am/pm") & "' Where id = " & ds!id
            Sdb.Execute s
        End If
        If edcell = "pickup" Then
            Grid1.Text = Trim(Grid1.Text)
            If Len(Grid1.Text) > 50 Then Grid1.Text = Left(Grid1.Text, 50)
            If Len(Grid1.Text) = 0 Then
                s = "Update runs set pickup = ' ' Where id = " & ds!id
                Sdb.Execute s
            Else
                s = "Update runs set pickup = '" & Grid1.Text & "' Where id = " & ds!id
                Sdb.Execute s
            End If
        End If
        If edcell = "oc" Then
            s = "Update runs set oc = '" & Grid1.Text & "' Where id = " & ds!id
            Sdb.Execute s
        End If
    End If
    ds.Close
    edcell = "": outfile = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "update_run", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " update_run - Error Number: " & eno
        End
    End If
End Sub

Private Sub sched_file()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    sqlx = "select loaded,destination,trldate,sum(trlsize)"
    sqlx = sqlx & " from runs where trldate > '" & Format(Now, "m-d-yyyy") & "'"
    sqlx = sqlx & " group by loaded,destination,trldate"
    sqlx = sqlx & " order by trldate,destination,loaded"
    Set ds = Sdb.Execute(sqlx)
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
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "sched_file", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " sched_file - Error Number: " & eno
        End
    End If
End Sub
Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String, gflag As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Cols = 9: Grid1.Rows = 1
    sqlx = "select runs.id,plantname,locname,trlno,trlsize,startime,pickup,oc, loaded, destination"
    sqlx = sqlx & " from runs,plants"
    sqlx = sqlx & " where trldate = '" & Combo1 & "'"
    'sqlx = sqlx & " and val(runs.loaded) = plants.plant"
    sqlx = sqlx & " and runs.loaded = plants.plant"
    sqlx = sqlx & " order by loaded,startime"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            gflag = " "
            sqlx = ds(0) & Chr$(9) & ds(1) & Chr$(9) & ds(2) & Chr$(9)
            sqlx = sqlx & ds(3) & Chr$(9) & ds(4) & Chr$(9)
            sqlx = sqlx & Format$(ds(5), "h:mm am/pm") & Chr$(9)
            sqlx = sqlx & " " & ds(6) & Chr$(9) & ds(7)
            If ds!Destination = "16" Then gflag = "*"
            If ds!Destination = "15" Then gflag = "*"
            If Val(ds!Destination) = 0 Then gflag = "*"
            If Val(ds!Destination) = 1 Then gflag = "*"
            If ds!loaded = "50" And ds!Destination = "51" Then gflag = "*"
            If ds!loaded = "51" And ds!Destination = "50" Then gflag = "*"
            If ds!loaded = "52" And ds!Destination = "50" Then gflag = "*"
            sqlx = sqlx & Chr(9) & gflag
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
        'Grid1.RemoveItem 1
        Grid1.FixedRows = 1: Grid1.FixedCols = 2
    End If
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 8) > " " Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = pcolor.BackColor
            End If
        Next i
        Grid1.Row = 1
    End If
    Grid1.FormatString = "^BatchID|<Plant|<Destination|^#|^Size|^Start|<Notes and Contents|^OC|^NBO"
    Grid1.ColWidth(0) = 900: Grid1.ColWidth(1) = 1400
    Grid1.ColWidth(2) = 2000: Grid1.ColWidth(3) = 400
    Grid1.ColWidth(4) = 450: Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 4500: Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 800
    ds.Close
    Grid1.Row = 1: Grid1.Col = 2
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
    Form1.cdate = Format(Combo1, "m-d-yyyy")
End Sub

Private Sub Command1_Click()
    Dim ds As adodb.Recordset, sqlx As String, pname As String
    Dim pc As String, dc As String, stime As String, lc As String, pkey As Long
    On Error GoTo vberror
    pc = InputBox$("Plant Code ", "Plant Code", "50")
    If Len(pc) = 0 Then Exit Sub
    dc = InputBox$("Branch Code ", "Branch Code", "28")
    If Len(dc) = 0 Then Exit Sub
    stime = InputBox$("Start Time", "Start Time", Format$(Now, "h:mm am/pm"))
    If Len(stime) = 0 Then Exit Sub
    If IsDate(stime) = False Then
        MsgBox "Invalid Time Format", vbOKOnly, "Sorry"
        Exit Sub
    End If
    sqlx = "select * from plants where plant = " & pc
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid Plant Code " & pc & " used.", vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    Else
        pname = ds!plantname
    End If
    ds.Close
    sqlx = "select * from branches where branch = " & dc
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid Branch Code " & dc & " used.", vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    End If
    lc = ds!branchname
    ds.Close
    pkey = wd_seq("Oratkt", Form1.schdb)
    sqlx = "Insert into runs (id, loaded, destination, locname, trlno, trlsize, trldate, startime, pickup, oc)"
    sqlx = sqlx & " Values (" & pkey
    sqlx = sqlx & ", '" & pc & "'"
    sqlx = sqlx & ", '" & dc & "'"
    sqlx = sqlx & ", '" & lc & "'"
    sqlx = sqlx & ", '#'"
    sqlx = sqlx & ", 32"
    sqlx = sqlx & ", '" & Combo1 & "'"
    sqlx = sqlx & ", '" & stime & "'"
    sqlx = sqlx & ", 'Added'"
    sqlx = sqlx & ", ' ')"
    Sdb.Execute sqlx
    sqlx = pkey & Chr(9) & pname & Chr(9) & lc & Chr(9)
    sqlx = sqlx & "#" & Chr(9) & "32" & Chr(9)
    sqlx = sqlx & stime & Chr(9) & "Added"
    Grid1.AddItem sqlx
    outfile = True
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

Private Sub Command2_Click()
    Dim sqlx As String
    On Error GoTo vberror
    sqlx = Grid1.TextMatrix(Grid1.Row, 2)
    sqlx = sqlx & " " & Grid1.TextMatrix(Grid1.Row, 3)
    If MsgBox("Clear " & sqlx & " on " & Combo1, vbYesNo + vbQuestion, "Are you sure?") = vbNo Then
        Exit Sub
    End If
    sqlx = "delete from runs where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Sdb.Execute sqlx
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_grid
    End If
    outfile = True
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
    Dim sqlx As String
    On Error GoTo vberror
    If MsgBox("Clear schedule for " & Combo1, vbOKCancel, "Are you sure?") = vbCancel Then
        Exit Sub
    End If
    sqlx = "delete from runs where trldate = '" & Combo1 & "'"
    Sdb.Execute sqlx
    Combo1.RemoveItem Combo1.ListIndex
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
    outfile = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command4_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command4_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command5_Click()
    Dim cfile As String, i As Integer, x
    If Grid1.Rows = 1 Then Exit Sub
    cfile = Form1.tempdir & "\aschedwrk.csv"
    Open cfile For Output As #1
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 1) = "Brenham" And Grid1.TextMatrix(i, 8) <= " " Then
            Write #1, Combo1;                               'Date
            Write #1, Grid1.TextMatrix(i, 1);               'Plant
            Write #1, Grid1.TextMatrix(i, 2);               'Branch
            Write #1, Grid1.TextMatrix(i, 3) & "    ";      'Trailer #
            Write #1, Grid1.TextMatrix(i, 4) & "    ";      'Size
            Write #1, Grid1.TextMatrix(i, 5);               'Start
            Write #1, Grid1.TextMatrix(i, 6);               'Notes
            Write #1, "  " & Grid1.TextMatrix(i, 7) & "  "  'Oc
        End If
    Next i
    Close #1
    MsgBox "Created file at: " & cfile, vbInformation + vbOKOnly, "Export completed...."
    'x = Shell("notepad.exe " & cfile, vbNormalFocus)
End Sub

Private Sub Command6_Click()
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String
    Dim rt As String, rf As String, rh As String, hf As String
    On Error GoTo vberror
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 8
    s = "select groupcode,branch,trailers.trlno,shipdate,startime,sum(units) from trailers, runs"
    s = s & " where ra_flag = 'N'"
    s = s & " and plant = 50"
    s = s & " and branch <> 16"
    s = s & " and runs.id = trailers.runid"
    s = s & " group by groupcode,branch,trailers.trlno,shipdate,startime"
    s = s & " order by shipdate,groupcode,branch,trailers.trlno,startime"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!groupcode & Chr(9)
            s = s & Format(ds!branch, "00") & Chr(9)
            s = s & ds!trlno & Chr(9)
            s = s & Format(ds!shipdate, "mm-dd-yyyy") & Chr(9)
            s = s & Format(ds!startime, "hh:mm AM/PM")
            If Left(ds(2), 1) = "#" Then Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.schdb
    For i = 1 To Grid2.Rows - 1
        s = "select * from schedule where trdate = '" & Grid2.TextMatrix(i, 3) & "'"
        s = s & " and trid = (select sched1 from locations where lcode = '" & Grid2.TextMatrix(i, 1) & "')"
        s = s & " and trailer = '" & Grid2.TextMatrix(i, 2) & "'"
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            'Grid2.TextMatrix(i, 5) = ds!trid
            Grid2.TextMatrix(i, 5) = Format(ds!startime, "hh:mm AM/PM")
            Grid2.TextMatrix(i, 6) = ds!crt
            Grid2.TextMatrix(i, 7) = Format(ds!lastchg, "mm-dd-yyyy hh:mm AM/PM")
        End If
        ds.Close
    Next i
    db.Close
    For i = 1 To Grid2.Rows - 1
        s = "select branchname from branches where branch = " & Val(Grid2.TextMatrix(i, 1))
        Set ds = Sdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Grid2.TextMatrix(i, 2) = ds!branchname & " " & Grid2.TextMatrix(i, 2)
        End If
        ds.Close
    Next i
    
    Grid2.FillStyle = flexFillRepeat
    For i = 1 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 4) <> Grid2.TextMatrix(i, 5) Then
            Grid2.Row = i: Grid2.RowSel = i
            Grid2.Col = 0: Grid2.ColSel = Grid2.Cols - 1
            Grid2.CellBackColor = ycolor.BackColor
        Else
            Grid2.TextMatrix(i, 6) = "."
            Grid2.TextMatrix(i, 7) = "."
        End If
    Next i
    Grid2.FormatString = "^Group|^Branch|^Trailer|^Date|^OrigTime|^NewTime|^User|^Changed@"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1 '000
    Grid2.ColWidth(2) = 1800
    Grid2.ColWidth(3) = 1200
    Grid2.ColWidth(4) = 1200
    Grid2.ColWidth(5) = 1200
    Grid2.ColWidth(6) = 1000
    Grid2.ColWidth(7) = 1800
    
    rt = "Grouped Transport Schedule Start Times"
    rh = "Grouped Transport Schedule Start Times"
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    hf = Form1.tempdir & "\schestarts.htm"
    htdc(0) = "Yellow": gndc(0) = ycolor.BackColor
    Call htmlcolorgrid(Me, hf, Grid2, rt, rh, rf, "linen", "lemonchiffon", "white")
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe " & hf, vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & hf, vbNormalFocus)
        Exit Sub
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command6_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command6_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Len(edcell) > 0 Then
        If MsgBox("Update schedule record?", vbYesNo + vbQuestion, "Save changes...") = vbYes Then
            Call update_run
        Else
            edcell = ""
        End If
    End If
    If Transched.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            If Form1.FrmGrid.TextMatrix(i, 0) = "transched" Then
                Form1.FrmGrid.TextMatrix(i, 1) = Transched.Top
                Form1.FrmGrid.TextMatrix(i, 2) = Transched.Left
                Form1.FrmGrid.TextMatrix(i, 3) = Transched.Height
                Form1.FrmGrid.TextMatrix(i, 4) = Transched.Width
                Exit For
            End If
        Next i
    End If
    If outfile Then Call sched_file
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Transched.ActiveControl.Name = "Grid1" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command1_Click
        If KeyCode = 46 Or KeyCode = 120 Then Call Command2_Click
    End If
End Sub

Private Sub Form_Load()
    Dim ds As adodb.Recordset, sqlx As String
    Dim i As Integer
    On Error GoTo vberror
    outfile = False
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    For i = 1 To Form1.FrmGrid.Rows - 1
        If Form1.FrmGrid.TextMatrix(i, 0) = "transched" Then
            Transched.Top = Val(Form1.FrmGrid.TextMatrix(i, 1))
            Transched.Left = Val(Form1.FrmGrid.TextMatrix(i, 2))
            Transched.Height = Val(Form1.FrmGrid.TextMatrix(i, 3))
            Transched.Width = Val(Form1.FrmGrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    sqlx = "select distinct trldate from runs order by trldate"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem Format$(ds(0), "m-d-yyyy")
            ds.MoveNext
        Loop
        For i = 0 To Combo1.ListCount - 1
            If Combo1.List(i) = Form1.cdate Then
                Combo1.ListIndex = i
                Exit For
            End If
        Next i
        If Combo1.ListIndex < 0 Then Combo1.ListIndex = 0
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
    Grid1.Width = Transched.Width - 100
    If Transched.Height > 2000 Then
        Grid1.Height = Transched.Height - 750
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_GotFocus()
    Grid1.FocusRect = flexFocusNone
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
    If Grid1.Row = 0 Then Exit Sub
    If Grid1.Col < 2 Then Exit Sub
    If Len(edcell) = 0 And Grid1.Col <> 7 Then Grid1.Text = ""
    If Grid1.Col = 2 Then edcell = "locname"
    If Grid1.Col = 3 Then edcell = "trlno"
    If Grid1.Col = 4 Then edcell = "trlsize"
    If Grid1.Col = 5 Then edcell = "startime"
    If Grid1.Col = 6 Then edcell = "pickup"
    If Grid1.Col = 7 Then
        edcell = "oc"
        If Grid1.Text = "*" Then
            Grid1.Text = " "
        Else
            Grid1.Text = "*"
        End If
        Exit Sub
    End If
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
    If Len(edcell) > 0 Then Call update_run
End Sub

Private Sub Grid1_LostFocus()
    If Len(edcell) > 0 Then Call update_run
    Grid1.FocusRect = flexFocusLight
End Sub
