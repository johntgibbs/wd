VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Brorders 
   Caption         =   "Branch Orders"
   ClientHeight    =   9135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9135
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Refresh Listing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   16
      Top             =   120
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   9015
      Left            =   9600
      TabIndex        =   15
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   15901
      _Version        =   327680
      ForeColor       =   12582912
      BackColorFixed  =   12632319
      FocusRect       =   0
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   1695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   6720
      Width           =   7935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Drop SKU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert SKU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6240
      Width           =   1455
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   240
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5535
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   9763
      _Version        =   327680
      Cols            =   8
      FixedCols       =   3
      BackColor       =   16777215
      BackColorFixed  =   12648447
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   1
      FormatString    =   "^ID|^sku|Product|^Order|^Grouped|^Net|^Alt|^Wraps"
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   3735
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
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2415
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Account"
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
      Left            =   5640
      TabIndex        =   8
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Branch"
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
      Left            =   3240
      TabIndex        =   7
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Order Date"
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
      Left            =   1920
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Plant"
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
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu inssku 
         Caption         =   "Insert SKU"
      End
      Begin VB.Menu dropsku 
         Caption         =   "Drop SKU"
      End
   End
End
Attribute VB_Name = "Brorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rf As String
Dim edcell As String

Function branch_notes(bcode As String) As String
    Dim cfile As String, s As String, t As String
    t = bcode & " Notes:" & vbCrLf
    cfile = Form1.webdir & "\orders\notes." & Left(bcode, 2)
    If Len(Dir(cfile)) > 0 Then
        t = t & "Updated:  " & Format(FileDateTime(cfile), "dddd M-dd-yyyy h:mm am/pm") & vbCrLf
        Open cfile For Input As #1
        Do Until EOF(1)
            Line Input #1, s
            If s > " " Then t = t & s & vbCrLf
        Loop
        Close #1
    End If
    branch_notes = t
End Function

Private Sub refresh_grid2()
    Dim ds As adodb.Recordset, s As String
    Grid2.Font = "Arial": Grid2.FontSize = 9: Grid2.FontBold = True
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 7
    s = "select plant,branch,account,orddate,sum(ordqty),sum(grpqty),sum(netqty)"
    s = s & " from brorders group by plant,branch,account,orddate"
    s = s & " order by orddate,plant,branch"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = Format(ds!orddate, "M-d-yyyy") & Chr(9)
            s = s & ds!plant
            If ds!plant = 50 Then s = s & " Brenham"
            If ds!plant = 51 Then s = s & " Broken Arrow"
            If ds!plant = 52 Then s = s & " Sylacauga"
            s = s & Chr(9)
            s = s & Format(ds!branch, "00") & Chr(9)
            s = s & ds!account & Chr(9)
            s = s & ds(4) & Chr(9) & ds(5) & Chr(9) & ds(6)
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid2.Rows > 1 Then
        For i = 1 To Grid2.Rows - 1
            If Val(Grid2.TextMatrix(i, 2)) > 0 Then
                s = "select branchname from branches where branch = " & Val(Grid2.TextMatrix(i, 2))
                Set ds = Sdb.Execute(s)
                If ds.BOF = False Then
                    ds.MoveFirst
                    Grid2.TextMatrix(i, 2) = Grid2.TextMatrix(i, 2) & " " & ds!branchname
                End If
                ds.Close
            End If
        Next i
    End If
    Grid2.FormatString = "^Date|<Plant|<Branch|<Account|^Order|^Group|^Net"
    Grid2.ColWidth(0) = 1100
    Grid2.ColWidth(1) = 1700
    Grid2.ColWidth(2) = 2000
    Grid2.ColWidth(3) = 1200
    Grid2.ColWidth(4) = 800
    Grid2.ColWidth(5) = 800
    Grid2.ColWidth(6) = 800
End Sub

Private Sub update_ord()
    Dim db As adodb.Connection, sqlx As String
    On Error GoTo vberror
    If edcell = "ordqty" Then
        Grid1.Text = Val(Grid1.Text)
        If Val(Grid1.Text) = 0 Then Grid1.Text = ""
        sqlx = "Update brorders set"
        sqlx = sqlx & " ordqty = " & Val(Grid1.TextMatrix(Grid1.Row, 3))
        sqlx = sqlx & ",grpqty = " & Val(Grid1.TextMatrix(Grid1.Row, 4))
        sqlx = sqlx & ",netqty = " & Val(Grid1.TextMatrix(Grid1.Row, 5))
        sqlx = sqlx & " Where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Sdb.Execute sqlx
    End If
    If edcell = "grpqty" Then
        Grid1.Text = Val(Grid1.Text)
        If Val(Grid1.Text) = 0 Then Grid1.Text = ""
        sqlx = "Update brorders set"
        sqlx = sqlx & " ordqty = " & Val(Grid1.TextMatrix(Grid1.Row, 3))
        sqlx = sqlx & ",grpqty = " & Val(Grid1.TextMatrix(Grid1.Row, 4))
        sqlx = sqlx & ",netqty = " & Val(Grid1.TextMatrix(Grid1.Row, 5))
        sqlx = sqlx & " Where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Sdb.Execute sqlx
    End If
    If edcell = "altflag" Then
        sqlx = "Update brorders set altflag = '" & Grid1.Text & "'"
        sqlx = sqlx & " Where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Sdb.Execute sqlx
    End If
    If edcell = "partqty" Then
        Grid1.Text = Val(Grid1.Text)
        If Val(Grid1.Text) = 0 Then Grid1.Text = ""
        sqlx = "Update brorders set partqty = " & Val(Grid1.Text)
        sqlx = sqlx & " Where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Sdb.Execute sqlx
    End If
    edcell = ""
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "update_ord", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " update_ord - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo1_Click()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Combo2.Clear
    Combo3.Clear
    If IsDate(Combo1) = True Then
        Screen.MousePointer = 11
        sqlx = "Select branch, branchname from branches where branch in "
        sqlx = sqlx & "(select branch from brorders where orddate = '" & Combo1 & "'"
        sqlx = sqlx & " and plant = " & Left$(Combo4, 2) & ")"
        sqlx = sqlx & " Order by branch"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                Combo2.AddItem Format$(ds(0), "00") & " " & ds(1)
                ds.MoveNext
            Loop
        End If
        ds.Close
        Screen.MousePointer = 0
        Combo2.ListIndex = 0
        Form1.cdate = Format(Combo1, "m-d-yyyy")
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "combo1_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " combo1_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo2_Click()
    Dim ds As adodb.Recordset, sqlx As String
    Dim oq As Integer, pq As Integer
    On Error GoTo vberror
    Label1 = "": oq = 0: pq = 0
    Combo3.Clear
    If Left$(Combo2, 2) = "15" Or Left$(Combo2, 2) = "16" Then
        Screen.MousePointer = 11
        Combo3.Visible = True
        Combo3.Clear
        sqlx = "Select account,acctdesc from jobbing"
        sqlx = sqlx & " Where branch = " & Left$(Combo2, 2)
        sqlx = sqlx & " And Account in (Select account from brorders"
        sqlx = sqlx & " Where branch = " & Left$(Combo2, 2)
        sqlx = sqlx & " And orddate = '" & Combo1 & "')"
        sqlx = sqlx & " order by acctdesc"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                Combo3.AddItem ds(0) & " " & ds(1)
                ds.MoveNext
            Loop
        End If
        ds.Close
        Combo3.ListIndex = 0
        Screen.MousePointer = 0
    Else
        Combo3.Visible = False
        Grid1.Rows = 2
        sqlx = "select id,brorders.sku,fgunit,fgdesc,ordqty,grpqty,netqty,altflag,partqty"
        sqlx = sqlx & " from brorders,skumast where orddate = '" & Combo1 & "'"
        sqlx = sqlx & " and plant = " & Left$(Combo4, 2)
        sqlx = sqlx & " and branch = " & Left$(Combo2, 2)
        sqlx = sqlx & " and brorders.sku = skumast.sku order by brorders.sku"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                sqlx = ds(0) & Chr$(9)
                sqlx = sqlx & ds(1) & Chr$(9)
                sqlx = sqlx & " " & ds(2) & " " & ds(3) & Chr$(9)
                If ds(4) > 0 Then sqlx = sqlx & ds(4)
                sqlx = sqlx & Chr$(9)
                If ds(5) > 0 Then sqlx = sqlx & ds(5)
                sqlx = sqlx & Chr$(9)
                If ds(6) > 0 Then sqlx = sqlx & ds(6)
                sqlx = sqlx & Chr$(9)
                sqlx = sqlx & ds(7) & Chr$(9)
                If ds(8) > 0 Then sqlx = sqlx & ds(8)
                Grid1.AddItem sqlx
                oq = oq + ds(4)
                pq = pq + ds(8)
                ds.MoveNext
            Loop
            Grid1.RemoveItem 1
        End If
        ds.Close
        Label1 = oq & " Pallets " & pq & " Wraps"
    End If
    Text1 = branch_notes(Combo2)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "combo2_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " combo2_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo3_Click()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Grid1.Rows = 2
    Screen.MousePointer = 11
    sqlx = "select id,brorders.sku,fgunit,fgdesc,ordqty,grpqty,netqty,altflag,partqty"
    sqlx = sqlx & " from brorders,skumast where orddate = '" & Combo1 & "'"
    sqlx = sqlx & " and plant = " & Left$(Combo4, 2)
    sqlx = sqlx & " and branch = " & Left$(Combo2, 2)
    sqlx = sqlx & " and account = '" & Left$(Combo3, 6) & "'"
    sqlx = sqlx & " and brorders.sku = skumast.sku order by brorders.sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds(0) & Chr$(9)
            sqlx = sqlx & ds(1) & Chr$(9)
            sqlx = sqlx & " " & ds(2) & " " & ds(3) & Chr$(9)
            If ds(4) > 0 Then sqlx = sqlx & ds(4)
            sqlx = sqlx & Chr$(9)
            If ds(5) > 0 Then sqlx = sqlx & ds(5)
            sqlx = sqlx & Chr$(9)
            If ds(6) > 0 Then sqlx = sqlx & ds(6)
            sqlx = sqlx & Chr$(9)
            sqlx = sqlx & ds(7) & Chr$(9)
            If ds(8) > 0 Then sqlx = sqlx & ds(8)
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
        Grid1.RemoveItem 1
    End If
    ds.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "combo3_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " combo3_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo4_Click()
    Dim ds As adodb.Recordset, sqlx As String
    Dim i As Integer
    Screen.MousePointer = 11
    On Error GoTo vberror
    Combo1.Clear
    Combo2.Clear
    Combo3.Clear
    Grid1.Rows = 2
    Grid1.AddItem " "
    Grid1.RemoveItem 1
    sqlx = "Select distinct orddate from brorders where plant = " & Left$(Combo4, 2)
    sqlx = sqlx & " order by orddate"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem Format(ds!orddate, "m-d-yyyy")
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
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "combo4_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " combo4_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command1_Click()
    Dim mplant As String, mdate As String, mbranch As String, macct As String, pkey As Long
    Dim msku As String, sqlx As String, ds As adodb.Recordset
    On Error GoTo vberror
    mplant = InputBox$("Please Enter Plant Code..", "New Order Plant", Left$(Combo4, 2))
    If Len(mplant) = 0 Then Exit Sub
    mdate = InputBox$("Please Enter Order Date..", "New Order Date", Combo1)
    If Len(mdate) = 0 Then Exit Sub
    If IsDate(mdate) = False Then
        MsgBox "Invalid Date Format..", vbOKOnly, "Sorry"
        Exit Sub
    End If
    Form1.cdate = Format(mdate, "m-d-yyyy")
    mbranch = InputBox$("Please Enter Branch Code..", "New Order Branch", Left$(Combo2, 2))
    If Len(mbranch) = 0 Then Exit Sub
    If mbranch = "15" Or mbranch = "16" Then
        macct = InputBox$("Please Enter Account #..", "New Order Account", "000000")
    Else
        macct = "......"
    End If
    If Len(macct) = 0 Then Exit Sub
    msku = InputBox$("Please Enter a valid SKU #...", "New Order SKU", "777")
    If Len(msku) = 0 Then Exit Sub
    sqlx = "Select plant from plants where plant = " & mplant
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid Plant Code!!!!", vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    End If
    ds.Close
    sqlx = "Select branch from branches where branch = " & mbranch
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid Branch Code!!!", vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    End If
    ds.Close
    If macct <> "......" Then
        sqlx = "Select account from jobbing where branch = " & mbranch
        sqlx = sqlx & " and account = '" & macct & "'"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = True Then
            MsgBox "Invalid Jobbing Account #", vbOKOnly, "Sorry"
            ds.Close
            Exit Sub
        End If
        ds.Close
    End If
    sqlx = "Select sku from skumast where sku = '" & msku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid SKU #!!!", vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    End If
    ds.Close
    pkey = wd_seq("brorders", Form1.shipdb)
    sqlx = "Insert into brorders (id, plant,branch,account,sku,orddate,ordqty,grpqty,netqty,altflag,partqty)"
    sqlx = sqlx & " Values (" & pkey & "," & mplant & ","
    sqlx = sqlx & mbranch & ","
    sqlx = sqlx & "'" & macct & "',"
    sqlx = sqlx & "'" & msku & "',"
    sqlx = sqlx & "'" & mdate & "',"
    sqlx = sqlx & "0,0,0,0,0)"
    Sdb.Execute sqlx
    Call Combo4_Click
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
    Dim msku As String, ds As adodb.Recordset, sqlx As String
    Dim y As Integer, i As Integer, j As Integer, pkey As Long, pdesc As String
    On Error GoTo vberror
    If Len(edcell) > 0 Then Call update_ord
    msku = InputBox$("Please Enter a valid SKU #...", "Insert SKU", "777")
    If Len(msku) = 0 Then Exit Sub
    sqlx = "select sku,fgdesc,fgunit from skumast where sku = '" & msku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "SKU not found...", vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    End If
    pdesc = " " & ds!fgunit & " " & ds!fgdesc
    ds.Close
    
    pkey = wd_seq("brorders", Form1.shipdb)
    sqlx = "Insert into brorders (id, plant, branch, account, sku, orddate, ordqty, grpqty, netqty, altflag, partqty)"
    sqlx = sqlx & " Values (" & pkey
    sqlx = sqlx & ", " & Val(Left(Combo4, 2))
    sqlx = sqlx & ", " & Val(Left(Combo2, 2))
    If Combo3.Visible = True Then
        sqlx = sqlx & ", '" & Left(Combo3, 6) & "'"
    Else
        sqlx = sqlx & ", '......'"
    End If
    sqlx = sqlx & ", '" & msku & "'"
    sqlx = sqlx & ", '" & Combo1 & "'"
    sqlx = sqlx & ", 0, 0, 0, 'N', 0)"
    Sdb.Execute sqlx
    For i = 1 To Grid1.Rows - 1
        y = i
        If Grid1.TextMatrix(i, 1) > msku Then Exit For
    Next i
    Grid1.AddItem " "
    For i = Grid1.Rows - 2 To y Step -1
        For j = 0 To Grid1.Cols - 1
            Grid1.TextMatrix(i + 1, j) = Grid1.TextMatrix(i, j)
        Next j
    Next i
    Grid1.TextMatrix(y, 0) = pkey
    Grid1.TextMatrix(y, 1) = msku
    Grid1.TextMatrix(y, 2) = pdesc
    Grid1.TextMatrix(y, 3) = ""
    Grid1.TextMatrix(y, 4) = ""
    Grid1.TextMatrix(y, 5) = ""
    Grid1.TextMatrix(y, 6) = "N"
    Grid1.TextMatrix(y, 7) = ""
    Grid1.Row = y: Grid1.Col = 3
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
    Dim sqlx As String
    On Error GoTo vberror
    If MsgBox("Are you sure?", vbOKCancel, "Clear Order") = vbCancel Then Exit Sub
    sqlx = "Delete from brorders where plant = " & Left$(Combo4, 2)
    sqlx = sqlx & " And branch = " & Left$(Combo2, 2)
    If Combo3.Visible = True Then
        sqlx = sqlx & " And account = '" & Left$(Combo3, 6) & "'"
    End If
    sqlx = sqlx & " And orddate = '" & Combo1 & "'"
    Sdb.Execute sqlx
    Call Combo4_Click
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command3_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command3_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command4_Click()
    If Grid1.Row < 1 Then Exit Sub
    On Error GoTo vberror
    If MsgBox("Delete " & Grid1.TextMatrix(Grid1.Row, 2) & " From this order?", vbYesNo + vbQuestion, "Are you sure?") = vbNo Then Exit Sub
    Sdb.Execute "delete from brorders where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call Combo2_Click
    End If
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
    refresh_grid2
End Sub

Private Sub dropsku_Click()
    Command4_Click
End Sub

Private Sub Form_Activate()
    rf = "Yes"
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Len(edcell) > 0 Then
        If MsgBox("Update order?", vbYesNo + vbQuestion, "Save changes...") = vbYes Then
            Call update_ord
        Else
            edcell = ""
        End If
    End If
    If Brorders.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            If Form1.FrmGrid.TextMatrix(i, 0) = "brorders" Then
                Form1.FrmGrid.TextMatrix(i, 1) = Brorders.Top
                Form1.FrmGrid.TextMatrix(i, 2) = Brorders.Left
                Form1.FrmGrid.TextMatrix(i, 3) = Brorders.Height
                Form1.FrmGrid.TextMatrix(i, 4) = Brorders.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Brorders.ActiveControl.Name = "Grid1" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command2_Click
        If KeyCode = 46 Or KeyCode = 120 Then Call Command4_Click
    End If
End Sub

Private Sub Form_Load()
    Dim ds As adodb.Recordset, sqlx As String
    Dim i As Integer
    For i = 1 To Form1.FrmGrid.Rows - 1
        If Form1.FrmGrid.TextMatrix(i, 0) = "brorders" Then
            Brorders.Top = Val(Form1.FrmGrid.TextMatrix(i, 1))
            Brorders.Left = Val(Form1.FrmGrid.TextMatrix(i, 2))
            Brorders.Height = Val(Form1.FrmGrid.TextMatrix(i, 3))
            Brorders.Width = Val(Form1.FrmGrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Me.Top = 0 ' Val(Form1.formtop.Caption)          'Commented this out because I have no idea why it exists. Reece - 3/25/2019
    Me.Left = 0
    Me.Width = Screen.Width - 200
    Me.Height = Screen.Height - (Me.Top + Command1.Height)
    refresh_grid2
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    rf = "No"
    On Error GoTo vberror
    sqlx = "Select * from plants order by plant"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo4.AddItem ds(0) & " " & ds(1)
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "ID|^SKU|Product|^Order|^Grouped|^Net|^Alt|^Wraps"
    Grid1.ColWidth(0) = 1: Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 4000: Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000: Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 800: Grid1.ColWidth(7) = 800
    Combo4.ListIndex = 0
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
    If Brorders.Height > 1000 Then
        Command1.Top = Brorders.Height - 1155 ' 855
        Command2.Top = Command1.Top
        Command3.Top = Command1.Top
        Command4.Top = Command1.Top
        Label1.Top = Command1.Top
        Text1.Top = Command1.Top - (Text1.Height + 50)
        'Grid2.Height = Me.Height - (Command1.Height * 2)
        'Grid2.Height = Grid1.Height + Text1.Height + Command5.Height
    End If
    If Brorders.Height > 2000 Then
        Grid1.Height = Text1.Top - (Grid1.Top + 50)
        Grid2.Height = Grid1.Height + Text1.Height + Command5.Height
    End If
    'Grid1.Width = Brorders.Width - 100
    Text1.Width = Grid1.Width
    Grid2.Width = Math.Abs(Me.Width - (Text1.Width + 400))
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
    If Grid1.Col < 3 Then Exit Sub
    If Grid1.Col = 5 Then Exit Sub
    If Len(edcell) = 0 And Grid1.Col <> 6 Then Grid1.Text = ""
    If Grid1.Col = 3 Then edcell = "ordqty"
    If Grid1.Col = 4 Then edcell = "grpqty"
    If Grid1.Col = 6 Then
        edcell = "altflag"
        If Grid1.Text = "Y" Then
            Grid1.Text = "N"
        Else
            Grid1.Text = "Y"
        End If
        Exit Sub
    End If
    If Grid1.Col = 7 Then edcell = "partqty"
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
    If edcell = "ordqty" Or edcell = "grpqty" Then
        Grid1.TextMatrix(Grid1.Row, 5) = Val(Grid1.TextMatrix(Grid1.Row, 3)) - Val(Grid1.TextMatrix(Grid1.Row, 4))
    End If
End Sub

Private Sub Grid1_LeaveCell()
    If Len(edcell) > 0 Then Call update_ord
End Sub

Private Sub Grid1_LostFocus()
    If Len(edcell) > 0 Then Call update_ord
    Grid1.FocusRect = flexFocusLight
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid2_RowColChange()
    Dim i As Integer, j As Integer
    If Grid2.Row < 1 Then Exit Sub
    i = Grid2.Row
    
    For j = 0 To Combo4.ListCount - 1
        If Combo4.List(j) = Grid2.TextMatrix(i, 1) Then
            Combo4.ListIndex = j
            Exit For
        End If
    Next j
    DoEvents
        
    For j = 0 To Combo1.ListCount - 1
        If Combo1.List(j) = Grid2.TextMatrix(i, 0) Then
            Combo1.ListIndex = j
            Exit For
        End If
    Next j
    DoEvents
        
        
    For j = 0 To Combo2.ListCount - 1
        If Combo2.List(j) = Grid2.TextMatrix(i, 2) Then
            Combo2.ListIndex = j
            Exit For
        End If
    Next j
End Sub

Private Sub inssku_Click()
    Command2_Click
End Sub
