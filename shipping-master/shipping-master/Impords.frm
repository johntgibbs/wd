VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Impords 
   Caption         =   "Import Branch Orders"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form2"
   ScaleHeight     =   8670
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Caption         =   "Bobtail Orders"
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
      Left            =   6360
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Branch-Jobbing"
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
      Left            =   6360
      TabIndex        =   9
      Top             =   960
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Branches"
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
      Left            =   6360
      TabIndex        =   8
      Top             =   720
      Value           =   -1  'True
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3855
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6800
      _Version        =   327680
      FixedCols       =   0
      ForeColor       =   32768
      BackColorFixed  =   12648447
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      AllowUserResizing=   1
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
      Height          =   285
      Left            =   7320
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Import All"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
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
      Left            =   6360
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Import Order"
      Enabled         =   0   'False
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
      Left            =   6360
      TabIndex        =   0
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label ordid 
      Caption         =   "0"
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
      Left            =   7320
      TabIndex        =   6
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Order#:"
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
      Left            =   6480
      TabIndex        =   5
      Top             =   3720
      Width           =   735
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
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Impords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imp_brorders()
    Dim ds As adodb.Recordset, sqlx As String
    Dim ofile As String, ofilex As String, pkey As Long
    Dim xbr As String, xsku As String, xodate As String, xoqty As String
    Dim xgqty As String, xaflg As String, xpqty As String, xplant As String
    On Error GoTo vberror
    If IsDate(Text1) = False Then
        MsgBox "Invalid Date Format", vbOKOnly, "Sorry"
        Exit Sub
    End If
    Form1.cdate = Format(Text1, "m-d-yyyy")
    Screen.MousePointer = 11
    ofile = Form1.webdir & "\orders\" & Grid1.TextMatrix(Grid1.Row, 0)
    ofilex = Form1.webdir & "\orders\X" & Grid1.TextMatrix(Grid1.Row, 0)
    sqlx = "Delete From Brorders Where Branch = " & mid$(Grid1.TextMatrix(Grid1.Row, 0), 6, 2)
    sqlx = sqlx & " and Plant = " & mid$(Grid1.TextMatrix(Grid1.Row, 0), 4, 2)
    sqlx = sqlx & " and Orddate = '" & Text1 & "'"
    Sdb.Execute sqlx
    Open ofile For Input As #1
    Do While Not EOF(1)
        Input #1, xbr, xsku, xodate, xoqty, xplant, xaflg, xpqty
        pkey = wd_seq("brorders", Form1.shipdb)
        sqlx = "Insert into brorders (id, plant, branch, account, sku, orddate, ordqty, grpqty, netqty"
        sqlx = sqlx & ", altflag, partqty) Values (" & pkey
        sqlx = sqlx & ", " & Val(xplant)
        sqlx = sqlx & ", " & Val(xbr)
        sqlx = sqlx & ", '......'"
        sqlx = sqlx & ", '" & xsku & "'"
        sqlx = sqlx & ", '" & Format(xodate, "m-d-yyyy") & "'"
        sqlx = sqlx & ", " & Val(xoqty)
        sqlx = sqlx & ", 0"
        sqlx = sqlx & ", " & Val(xoqty)
        If UCase(xaflg) = "Y" Then
            sqlx = sqlx & ", 'Y'"
        Else
            sqlx = sqlx & ", 'N'"
        End If
        sqlx = sqlx & ", " & Val(xpqty) & ")"
        Sdb.Execute sqlx
    Loop
    Close #1
    'On Error Resume Next
    If Len(Dir(ofilex)) > 0 Then Kill ofilex
    'On Error GoTo 0
    Name ofile As ofilex
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_ordlist
    End If
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "imp_brorders", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " imp_brorders - Error Number: " & eno
        End
    End If
End Sub
Private Sub imp_bobtail()
    Dim ds As adodb.Recordset, sqlx As String
    Dim bfile As String, bfilex As String
    Dim bbr As String, btrl As String, bsku As String
    Dim bdate As String, brun As Long, pkey As Long
    Screen.MousePointer = 11
    On Error GoTo vberror
    bfile = Form1.webdir & "\orders\bobtail\" & Grid1.TextMatrix(Grid1.Row, 0)
    bfilex = Form1.webdir & "\orders\bobtail\" & Grid1.TextMatrix(Grid1.Row, 0) & "X"
    btrl = "B" & mid(Grid1.TextMatrix(Grid1.Row, 0), 8, 1)
    Open bfile For Input As #1
    Input #1, bbr, bdate, bsku, bprod, bqty
    pkey = wd_seq("Oratkt", Form1.schdb)
    sqlx = "Insert into runs (id, loaded, destination, locname, trlno, trlsize, trldate, startime, pickup, oc)"
    sqlx = sqlx & " Values (" & pkey
    sqlx = sqlx & ", " & Form1.plantno
    sqlx = sqlx & ", " & bbr
    sqlx = sqlx & ", '" & bbr & "-Bobtail'"
    sqlx = sqlx & ", '" & btrl & "'"
    sqlx = sqlx & ", 0"
    sqlx = sqlx & ", '" & bdate & "'"
    sqlx = sqlx & ", '12:00 PM'"
    sqlx = sqlx & ", 'Branch-Bobtail'"
    sqlx = sqlx & ", '*')"
    Sdb.Execute sqlx
    brun = pkey
    Do Until EOF(1)
        Input #1, bbr, bdate, bsku, bprod, bqty
        pkey = wd_seq("trailers", Form1.shipdb)
        sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno, sku"
        sqlx = sqlx & ", pallets, wraps, units, pb_flag, ra_flag) Values (" & pkey
        sqlx = sqlx & ", " & brun
        sqlx = sqlx & ", '" & btrl & "-Add'"
        sqlx = sqlx & ", " & Form1.plantno
        sqlx = sqlx & ", " & Val(bbr)
        sqlx = sqlx & ", '......'"
        sqlx = sqlx & ", '" & bdate & "'"
        sqlx = sqlx & ", '" & btrl & "'"
        sqlx = sqlx & ", '" & bsku & "'"
        sqlx = sqlx & ", 0, 0, " & Val(bqty) & ", 'N', 'N')"
        Sdb.Execute sqlx
    Loop
    Close #1
    'On Error Resume Next
    If Len(Dir(bfilex)) > 0 Then Kill bfilex
    'On Error GoTo 0
    Name bfile As bfilex
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_ordlist
    End If
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "imp_bobtail", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " imp_bobtail - Error Number: " & eno
        End
    End If
End Sub
Private Sub imp_jobbing()
    Dim ds As adodb.Recordset, sqlx As String
    Dim jfile As String, jfilex As String
    Dim jbr As String, jbr2 As String, jacct As String
    Dim jname As String, jaddr As String, jcity As String
    Dim jdate As String, jrun As Long, jsku As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    jfile = Form1.webdir & "\orders\jobbing\" & Grid1.TextMatrix(Grid1.Row, 0)
    jfilex = Form1.webdir & "\orders\jobbing\" & Grid1.TextMatrix(Grid1.Row, 0) & "X"
    Open jfile For Input As #1
    Input #1, jbr, jbr2, jacct, jname, jaddr, jcity, jdate
    pkey = wd_seq("Oratkt", Form1.schdb)
    sqlx = "Insert into runs (id, loaded, destination, locname, trlno, trlsize, trldate, startime, pickup, oc)"
    sqlx = sqlx & " Values (" & pkey
    sqlx = sqlx & ", " & jbr
    sqlx = sqlx & ", " & jbr2
    sqlx = sqlx & ", '" & jname
    sqlx = sqlx & ", 'J1'"
    sqlx = sqlx & ", 0"
    sqlx = sqlx & ", '" & jdate & "'"
    sqlx = sqlx & ", '12:00 PM'"
    sqlx = sqlx & ", 'Branch-Jobbing'"
    sqlx = sqlx & ", '*')"
    Sdb.Execute sqlx
    jrun = ds!id
    Do Until EOF(1)
        Input #1, jsku, jprod, jqty
        pkey = wd_seq("trailers", Form1.shipdb)
        sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno, sku"
        sqlx = sqlx & ", pallets, wraps, units, pb_flag, ra_flag) Values (" & pkey
        sqlx = sqlx & ", " & jrun
        sqlx = sqlx & ", '" & jbr & "-Add'"
        sqlx = sqlx & ", " & Val(jbr)
        sqlx = sqlx & ", " & Val(jbr2)
        sqlx = sqlx & ", '" & jacct & "'"
        sqlx = sqlx & ", '" & jdate & "'"
        sqlx = sqlx & ", 'J1'"
        sqlx = sqlx & ", '" & jsku & "'"
        sqlx = sqlx & ", 0, 0, " & Val(jqty) & ", 'N', 'N')"
        Sdb.Execute sqlx
    Loop
    Close #1
    'On Error Resume Next
    If Len(Dir(jfilex)) > 0 Then Kill jfilex
    'On Error GoTo 0
    Name jfile As jfilex
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_ordlist
    End If
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "imp_jobbing", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " imp_jobbing - Error Number: " & eno
        End
    End If
End Sub
Private Sub refresh_ordlist()
    Dim spath As String, sdir As String, sqlx As String
    Grid1.Clear: Grid1.Cols = 3: Grid1.Rows = 1
    'Branch Orders
    If Option1.Value = True Then
        spath = Form1.webdir & "\orders\ord????." & ordid
        sdir = Dir$(spath)
        Do While sdir <> ""
            sqlx = sdir & Chr(9)
            sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\orders\" & sdir), "mm-dd-yyyy hh:mm am/pm") & Chr(9)
            sqlx = sqlx & FileLen(Form1.webdir & "\orders\" & sdir)
            Grid1.AddItem sqlx
            sdir = Dir$
        Loop
    End If
    'Branch Jobbing
    If Option2.Value = True Then
        spath = Form1.webdir & "\orders\jobbing\J??????.??"
        sdir = Dir$(spath)
        Do While sdir <> ""
            sqlx = sdir & Chr(9)
            sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\orders\jobbing\" & sdir), "mm-dd-yyyy hh:mm am/pm") & Chr(9)
            sqlx = sqlx & FileLen(Form1.webdir & "\orders\jobbing\" & sdir)
            Grid1.AddItem sqlx
            sdir = Dir$
        Loop
    End If
    'Bobtails
    If Option3.Value = True Then
        spath = Form1.webdir & "\orders\bobtail\bobtail?.??"
        sdir = Dir$(spath)
        Do While sdir <> ""
            sqlx = sdir & Chr(9)
            sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\orders\bobtail\" & sdir), "mm-dd-yyyy hh:mm am/pm") & Chr(9)
            sqlx = sqlx & FileLen(Form1.webdir & "\orders\bobtail\" & sdir)
            Grid1.AddItem sqlx
            sdir = Dir$
        Loop
    End If
    
    Grid1.FormatString = "^File|^Time Created|^Size"
    Grid1.ColWidth(0) = 1800
    Grid1.ColWidth(1) = 2200
    Grid1.ColWidth(2) = 1200
End Sub
Private Sub Command1_Click()
    If Grid1.Row = 0 Then Exit Sub
    If Val(Grid1.TextMatrix(Grid1.Row, 2)) = 0 Then Exit Sub
    If Option1.Value = True Then imp_brorders
    If Option2.Value = True Then imp_jobbing
    If Option3.Value = True Then imp_bobtail
End Sub

Private Sub Command2_Click()
    Unload Impords
End Sub

Private Sub Command3_Click()
    Dim i As Integer
    If Grid1.Rows < 2 Then Exit Sub
    If IsDate(Text1) = False And Option1.Value = True Then
        MsgBox "Invalid Date Format", vbOKOnly, "Sorry"
        Exit Sub
    End If
    For i = Grid1.Rows - 1 To 1 Step -1
        If Val(Grid1.TextMatrix(i, 2)) > 0 Then
            Grid1.Row = i
            Call Command1_Click
        End If
    Next i
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Impords.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
            If Form1.FrmGrid.Text = "impords" Then
                Form1.FrmGrid.Col = 1: Form1.FrmGrid.Text = Impords.Top
                Form1.FrmGrid.Col = 2: Form1.FrmGrid.Text = Impords.Top
                Form1.FrmGrid.Col = 3: Form1.FrmGrid.Text = Impords.Height
                Form1.FrmGrid.Col = 4: Form1.FrmGrid.Text = Impords.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.FrmGrid.Rows - 1
        Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
        If Form1.FrmGrid.Text = "impords" Then
            Form1.FrmGrid.Col = 1: Impords.Top = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 2: Impords.Left = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 3: Impords.Height = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 4: Impords.Width = Val(Form1.FrmGrid.Text)
            Exit For
        End If
    Next i
    Text1 = Format(DateAdd("d", 2, Now), "m-d-yyyy")
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Call refresh_ordlist
End Sub

Private Sub Form_Resize()
    If Impords.Height > 3000 Then Grid1.Height = Impords.Height - 380
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_RowColChange()
    If Val(Grid1.TextMatrix(Grid1.Row, 2)) > 0 Then
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
End Sub

Private Sub Option1_Click()
    refresh_ordlist
End Sub

Private Sub Option2_Click()
    refresh_ordlist
End Sub

Private Sub Option3_Click()
    refresh_ordlist
End Sub

Private Sub ordid_Change()
    Call refresh_ordlist
End Sub

Private Sub Text1_Change()
    If IsDate(Text1) Then
        ordid = DateDiff("d", "01-01-1999", Text1)
    Else
        ordid = "0"
    End If
End Sub
