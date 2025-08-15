VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form EdBills 
   Caption         =   "Bills of Lading"
   ClientHeight    =   7965
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12780
   LinkTopic       =   "Form3"
   ScaleHeight     =   7965
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Save Changes"
      Height          =   375
      Left            =   10800
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox ftplist 
      Height          =   2595
      Left            =   9120
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox runlist 
      Height          =   7470
      Left            =   11520
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid tmpgrid 
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   7200
      Visible         =   0   'False
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   1085
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1508
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   2295
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4048
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1935
      Left            =   3960
      TabIndex        =   10
      Top             =   5280
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3413
      _Version        =   327680
      BackColorFixed  =   16777152
      Appearance      =   0
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View All Fields"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7800
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   480
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   12938
      _Version        =   327680
      Cols            =   19
      BackColorSel    =   4210688
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label rcolor 
      BackColor       =   &H000000FF&
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9240
      TabIndex        =   19
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label gcolor 
      BackColor       =   &H0080FF80&
      Caption         =   "Label3"
      Height          =   255
      Left            =   9240
      TabIndex        =   18
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label rundate 
      Caption         =   "..."
      Height          =   255
      Left            =   10800
      TabIndex        =   17
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label srun 
      Caption         =   "..."
      Height          =   255
      Left            =   10800
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scanned Units Do Not Match Original Order"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label trlkey 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Find:"
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label cntlit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Records"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label hcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "All Records"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Ship Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
      Begin VB.Menu prtbill 
         Caption         =   "Print Bill of Lading"
      End
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Begin VB.Menu canline 
         Caption         =   "Cancel Line"
         Enabled         =   0   'False
      End
      Begin VB.Menu addbc 
         Caption         =   "Add New BarCode"
         Enabled         =   0   'False
      End
      Begin VB.Menu addwraps 
         Caption         =   "Add New Product - Wraps"
         Enabled         =   0   'False
      End
      Begin VB.Menu edunits 
         Caption         =   "Edit Units"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu postmenu 
      Caption         =   "Post"
      Enabled         =   0   'False
      Begin VB.Menu postr12 
         Caption         =   "Post to Oracle Batches"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu renmenu 
      Caption         =   "Rename Trailer"
      Enabled         =   0   'False
      Begin VB.Menu rentrl 
         Caption         =   "Rename Trailer"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "EdBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rename_trailer(runno As String)
    Dim ds As adodb.Recordset, s As String, pkey As Long
    Dim obranch As String, otrlno As String, odate As String
    Dim nbranch As String, ntrlno As String, ndate As String, nbatch As String
    Dim bname As String, cfile As String, buildrun As Boolean, newrun As String
    On Error GoTo vberror
    bname = "none": newrun = runno
    nbranch = InputBox("New Branch Code:", "new branch...")
    If Len(nbranch) = 0 Then Exit Sub
    s = "select branchname from branches where branch = " & nbranch
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        bname = ds!branchname
    End If
    ds.Close
    If bname = "none" Then
        MsgBox "Invalid branch!", vbOKOnly + vbExclamation, "sorry, try again..."
        Exit Sub
    End If
    
    ntrlno = InputBox("Trailer #:", "Trailer #...", "#1")
    If Len(ntrlno) = 0 Then Exit Sub
    If Len(ntrlno) <> 2 Then
        MsgBox "Invalid trailer code entered: " & ntrlno, vbOKOnly + vbExclamation, "sorry, try again.."
        Exit Sub
    End If
    
    ndate = InputBox("Ship Date:", "Shipping date...", rundate)
    If Len(ndate) = 0 Then Exit Sub
    If IsDate(ndate) = False Then
        MsgBox "Invalid date entered: " & ndate, vbOKOnly + vbExclamation, "sorry, try again..."
        Exit Sub
    End If
    
    s = "Ok to rename " & rundate & " " & Grid1.TextMatrix(Grid1.Row, 5)
    s = s & " to " & ndate & " " & bname & " " & ntrlno & "?"
    If MsgBox(s, vbYesNo + vbQuestion, "are you sure...") = vbNo Then Exit Sub
    
    nbatch = DateDiff("d", "1-1-2012", ndate) & Format(Val(nbranch), "00") & Right(ntrlno, 1)
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 1) = nbatch Then
            s = "Trailer " & bname & " " & ntrlno & " already exists for " & ndate & "!"
            MsgBox s, vbOKOnly + vbExclamation, "sorry, try again..."
            Exit Sub
        End If
    Next i
    
    s = "select * from trailers where runid = " & runno
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        buildrun = False
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update trailers set branch = " & Val(nbranch)
            s = s & ", account = '......'"
            s = s & ", shipdate = '" & ndate & "'"
            s = s & ", trlno = '" & ntrlno & "'"
            s = s & ", pb_flag = 'Y'"
            s = s & ", ra_flag = 'N'"
            s = s & " Where id = " & ds!id
            Sdb.Execute s
            ds.MoveNext
        Loop
    Else
        buildrun = True
    End If
    ds.Close
    
    If buildrun = True Then
        s = "select * from runs where loaded = '" & Form1.plantno & "'"
        s = s & " and destination = '" & nbranch & "'"
        s = s & " and trlno = '" & ntrlno & "'"
        s = s & " and startime = '" & ndate & "'"
        Set ds = Sdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            newrun = ds!id
        Else
            pkey = wd_seq("Oratkt", Form1.schdb)
            s = "Insert into runs (loaded, destination, locname, trlno, trlsize, trldate, startime, pickup, oc)"
            s = s & " Values (" & pkey
            s = s & ", '" & Form1.plantno & "'"
            s = s & ", '" & nbranch & "'"
            s = s & ", '" & bname & "'"
            s = s & ", '" & ntrlno & "'"
            s = s & ", 32"
            s = s & ", '" & ndate & "'"
            s = s & ", '12:00 PM'"
            s = s & ", 'Swapped-" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
            s = s & ", ' ')"
            Sdb.Execute s
            newrun = pkey
        End If
        ds.Close
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 17) = runno Then
                pkey = wd_seq("trailers", Form1.shipdb)
                s = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno, sku"
                s = s & ", pallets, wraps, units, whs_num, pb_flag, ra_flag) Values (" & pkey
                s = s & ", " & newrun
                s = s & ", '" & groupcode & "'"
                s = s & ", '" & Form1.plantno & "'"
                s = s & ", '" & nbranch & "'"
                s = s & ", '......'"
                s = s & ", '" & ndate & "'"
                s = s & ", '" & ntrlno & "'"
                s = s & ", '" & Left(Grid1.TextMatrix(i, 6), 3) & "'"
                If Grid1.TextMatrix(i, 7) > "00" Then
                    s = s & ", 1, 0"
                Else
                    s = s & ", 0, 0"
                End If
                s = s & ", " & Val(Grid1.TextMatrix(i, 11)) + Val(Grid1.TextMatrix(i, 13))
                s = s & ", 4, 'Y', 'N')"
                Sdb.Execute s
            End If
        Next i
    End If
    
    If newrun <> runno Then
        cfile = Form1.pallogs & "bill" & Format(ndate, "MMddyyyy") & ".txt"
        Open cfile For Append As #1
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 17) = runno Then
                Write #1, nbatch;
                Write #1, Grid1.TextMatrix(i, 2);
                Write #1, Grid1.TextMatrix(i, 3);
                Write #1, Grid1.TextMatrix(i, 4);
                Write #1, bname & " " & ntrlno;
                Write #1, Grid1.TextMatrix(i, 6);
                Write #1, Grid1.TextMatrix(i, 7);
                Write #1, Grid1.TextMatrix(i, 8);
                Write #1, Grid1.TextMatrix(i, 9);
                Write #1, Grid1.TextMatrix(i, 10);
                Write #1, Grid1.TextMatrix(i, 11);
                Write #1, Grid1.TextMatrix(i, 12);
                Write #1, Grid1.TextMatrix(i, 13);
                Write #1, "PEND";
                Write #1, Grid1.TextMatrix(i, 15);
                Write #1, Grid1.TextMatrix(i, 16);
                Write #1, newrun
            End If
        Next i
        Close #1
    Else
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 17) = runno Then
                s = "B" & Chr(9)
                s = s & nbatch & Chr(9)
                s = s & Grid1.TextMatrix(i, 2) & Chr(9)
                s = s & Grid1.TextMatrix(i, 3) & Chr(9)
                s = s & Grid1.TextMatrix(i, 4) & Chr(9)
                s = s & bname & " " & ntrlno & Chr(9)
                s = s & Grid1.TextMatrix(i, 6) & Chr(9)
                s = s & Grid1.TextMatrix(i, 7) & Chr(9)
                s = s & Grid1.TextMatrix(i, 8) & Chr(9)
                s = s & Grid1.TextMatrix(i, 9) & Chr(9)
                s = s & Grid1.TextMatrix(i, 10) & Chr(9)
                s = s & Grid1.TextMatrix(i, 11) & Chr(9)
                s = s & Grid1.TextMatrix(i, 12) & Chr(9)
                s = s & Grid1.TextMatrix(i, 13) & Chr(9)
                s = s & "PEND" & Chr(9)
                s = s & Grid1.TextMatrix(i, 15) & Chr(9)
                s = s & Grid1.TextMatrix(i, 16) & Chr(9)
                s = s & newrun
                Grid1.AddItem s
            End If
        Next i
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 17) = runno And Grid1.TextMatrix(i, 1) <> nbatch Then
                Grid1.TextMatrix(i, 14) = "CANC"
            End If
        Next i
    End If
    Call save_bills(runno)
    DoEvents
    Call refresh_grid1(Text1)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "rename_trailer", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " rename_trailer - Error Number: " & eno
        End
    End If
End Sub

Private Sub duplex_bill_log(runno As String)
    Dim db As adodb.Connection, ds As adodb.Recordset, sqlx As String, s As String
    Dim js As adodb.Recordset, jobtrail As Boolean, ppflag As Boolean
    Dim j1 As String, j2 As String, j3 As String, j4 As String, j5 As String
    Dim ss As adodb.Recordset, lc As Integer, tc As String
    Dim fcode As String, bcode As String, i As Integer
    Dim scode As String, stot As Currency, gtot As Currency
    Dim pno As Integer, tu As Long, tw As Integer, tp As Integer 'Currency
    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long, p5 As Long
    Dim dbranch As String, daddr1 As String, daddr2 As String, dphone As String, dfax As String
    Dim oplant As String, oaddr1 As String, oaddr2 As String, ophone As String, ofax As String
    Dim ldate As String, ltarget As String, cfile As String
    Dim f1 As String, f2 As String, f3 As String, f4 As String, f5 As String, f6 As String
    Dim f7 As String, f8 As String, f9 As String, f10 As String, f11 As String, f12 As String
    Dim f13 As String, f14 As String, f15 As String, f16 As String, f17 As String
    Dim bno As String, ano As String, pflag As Boolean, wflag As Boolean
    Dim tflag As Boolean                                            'jv072314
    pno = 1: jobtrail = False: pflag = False: wflag = False
    On Error GoTo vberror
    If Val(runno) = 0 Then Exit Sub
    tc = InputBox("Please Enter Trailer Code or 'OC'", "Trailer Code", "OC")
    If Len(tc) = 0 Then Exit Sub
    If UCase(tc) <> "OC" Then
        'Set db = CreateObject("ADODB.Connection")
        'db.Open Form1.shipdb
        'Set ds = db.Execute("select * from trailers where trlcode = " & tc)
        'If ds.BOF = True Then
        '    MsgBox "Invalid Trailer Code Entered", vbOKOnly + vbExclamation, "Sorry Cannot Process.."
        '    ds.Close: db.Close
        '    Exit Sub
        'End If
        'ds.Close: db.Close
        tflag = False                                               'jv072314
        cfile = Form1.webdir & "\bbtcodes.txt"                      'jv072314
        Open cfile For Input As #1                                  'jv072314
        Do Until EOF(1)                                             'jv072314
            Input #1, s                                             'jv072314
            If s = tc Then tflag = True                             'jv072314
        Loop                                                        'jv072314
        Close #1                                                    'jv072314
        If tflag = False Then                                       'jv072314
            MsgBox "Invalid Trailer Code Entered", vbOKOnly + vbExclamation, "Sorry Cannot Process.."
            Exit Sub                                                'jv072314
        End If                                                      'jv072314
    End If                                                          'jv072314
    
    Screen.MousePointer = 11
    
    Printer.Duplex = 3
    Printer.Orientation = 1
    
    oplant = Form1.plantno
    If oplant = "50" Then sqlx = "select * from branches where branch = 1"
    If oplant = "51" Then sqlx = "select * from branches where branch = 47"
    If oplant = "52" Then sqlx = "select * from branches where branch = 52"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        oaddr1 = ds!addr1
        oaddr2 = ds!addr2
        ophone = ds!brphone & " "
        ofax = ds!brfax & " "
    End If
    ds.Close
    
    sqlx = "select * from trailers where runid = " & runno
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        bno = ds!branch
        ano = ds!account
    End If
    ds.Close
    'If Val(bno) = 0 Then
    '    db.Close
    '    Screen.MousePointer = 0
    '    MsgBox "Original order is not available.", vbOKOnly + vbExclamation, "cannot print bill..."
    '    Exit Sub
    'End If
    If Val(bno) = 0 Then
        Screen.MousePointer = 0
        ano = "......"
        bno = InputBox("Branch code:", "Original order is not available...", "")
        If Len(bno) = 0 Or Val(bno) = 0 Then
            Exit Sub
        End If
        If Val(bno) = 15 Or Val(bno) = 16 Then
            ano = InputBox("Jobbing account:", "Original order is not available...", ano)
        End If
        Screen.MousePointer = 11
    End If
        
    sqlx = "select * from branches where branch = " & bno
    Set ds = Sdb.Execute(sqlx)
    ds.MoveFirst
    'Printer.Height = 1440 * 11
    'Printer.Width = 1440 * 8.5
    Printer.FontName = "Arial"
    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.Print Tab(32); " " '"B i l l   O f   L a d i n g"
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.CurrentX = 720: Printer.Print "Origination:";
    Printer.FontBold = False
    Printer.CurrentX = 1440 * 1.5: Printer.Print "Blue Bell Creameries L.P.";
    Printer.FontBold = True
    Printer.CurrentX = 1440 * 4.5: Printer.Print "Destination: ";
    Printer.FontBold = False
    Printer.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        jobtrail = True
        sqlx = "select * from jobbing where branch = " & bno
        sqlx = sqlx & " and account = '" & ano & "'"
        Set js = Sdb.Execute(sqlx)
        If js.BOF = False Then
            js.MoveFirst
            j1 = js!acctdesc & " "
            j2 = js!addr1 & " "
            j3 = js!addr2 & " "
            j4 = js!addr3 & " " & js!jzip
            j5 = js!jphone & " "
        Else
            j1 = " ": j2 = " ": j3 = " ": j4 = " ": j5 = " "
        End If
        js.Close
        If j2 <= " " Then
            j2 = j3: j3 = j4: j4 = j5
        End If
        If j3 <= " " Then
            j3 = j4: j4 = j5
        End If
        If j4 <= " " Then
            j4 = j5
        End If
        Printer.Print "Jobbing Account # "; bno; "-"; ano; " "
    Else
        Printer.Print Format(bno, "00"); " "; ds!branchname; " "; Right(Grid1.TextMatrix(Grid1.Row, 5), 2)
        ltarget = ds!branchname & " " & Right(Grid1.TextMatrix(Grid1.Row, 5), 2)      'jv022811
    End If
    
    Printer.CurrentX = 1440 * 1.5: Printer.Print oaddr1; '"1101 S. Blue Bell Road";
    Printer.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        Printer.Print j1
    Else
        Printer.Print ds!addr1
    End If
    Printer.CurrentX = 1440 * 1.5: Printer.Print oaddr2; '"Brenham, Texas  77834-1807";
    Printer.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        Printer.Print j2
    Else
        Printer.Print ds!addr2
    End If
    Printer.CurrentX = 1440 * 1.5: Printer.Print ophone; '"(979) 836-7977";
    Printer.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        Printer.Print j3
    Else
        Printer.Print ds!brphone
    End If
    Printer.CurrentX = 1440 * 1.5: Printer.Print "Fax: " & ofax; '"Fax: (979) 830-7398";
    Printer.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        Printer.Print j4
    Else
        Printer.Print "Fax: " & ds!brfax
    End If
    Printer.Print String(130, "_")
    ds.Close
    tu = 0: tw = 0: tp = 0
    tmpgrid.Clear: tmpgrid.Rows = 1: tmpgrid.Cols = 5
    ppflag = False
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 17) = runno And Grid1.TextMatrix(i, 14) <> "CANC" Then
            f1 = Grid1.TextMatrix(i, 1)
            f2 = Grid1.TextMatrix(i, 2)
            f3 = Grid1.TextMatrix(i, 3)
            f4 = Grid1.TextMatrix(i, 4)
            f5 = Grid1.TextMatrix(i, 5)
            f6 = Grid1.TextMatrix(i, 6)
            f7 = Grid1.TextMatrix(i, 7)
            f8 = Grid1.TextMatrix(i, 8)
            f9 = Grid1.TextMatrix(i, 9)
            f10 = Grid1.TextMatrix(i, 10)
            f11 = Grid1.TextMatrix(i, 11)
            f12 = Grid1.TextMatrix(i, 12)
            f13 = Grid1.TextMatrix(i, 13)
            f14 = Grid1.TextMatrix(i, 14)
            f15 = Grid1.TextMatrix(i, 15)
            f16 = Grid1.TextMatrix(i, 16)
            f17 = Grid1.TextMatrix(i, 17)
            s = "select sku,fgunit,fgdesc,pallet,numwrap from skumast"
            s = s & " where sku = '" & Trim(Left(f6, 4)) & "'"
            Set ss = Sdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                s = ss!sku & Chr(9)
                s = s & ss!fgunit & " " & ss!fgdesc & Chr(9)
                If f7 > "00" Then
                    s = s & mid(f7, 5, 12) & Chr(9)
                    pflag = True
                Else
                    s = s & "Partial" & Chr(9)
                    ppflag = True
                    wflag = True
                End If
                s = s & Format((Val(f11) + Val(f13)) / ss!numwrap, "0") & Chr(9)
                s = s & Format(Val(f11) + Val(f13), "0")
                tmpgrid.AddItem s
            End If
            ss.Close
            If Grid1.TextMatrix(i, 14) <> "POSTED" Then Grid1.TextMatrix(i, 14) = "PRINTED"
        End If
    Next i
    
    'Partials
    If ppflag = True Then tp = InputBox("# Partial Pallets:", "Partial pallet detected..", tp)
    
    If tmpgrid.Rows > 37 Then '46 Then
        p1 = 1440 * 3.5  '3.25 '2.75
        p2 = 1440 * 4   '3.75 '3.25
        p3 = 1440 * 4.25 '4.75
        p4 = 1440 * 7.5  '7.25
        p5 = 1440 * 8    '7.75
        Printer.FontName = "Arial"
        Printer.FontBold = True
        Printer.CurrentX = 360: Printer.Print "SKU  Description";
        If jobtrail Then
            Printer.CurrentX = p1 - Printer.TextWidth("Wraps")
            Printer.Print "Wraps";
        Else                                                            'jv100609
            Printer.CurrentX = p1 - Printer.TextWidth("Pallets")        'jv100609
            Printer.Print "Pallets";                                    'jv100609
        End If
        Printer.CurrentX = p2 - Printer.TextWidth("Units"): Printer.Print "Units";
        Printer.CurrentX = p3: Printer.Print ("SKU  Description");
        If jobtrail Then
            Printer.CurrentX = p4 - Printer.TextWidth("Wraps")
            Printer.Print "Wraps";
        Else                                                            'jv100609
            Printer.CurrentX = p4 - Printer.TextWidth("Pallets")        'jv100609
            Printer.Print "Pallets";                                    'jv100609
        End If
        Printer.CurrentX = p5 - Printer.TextWidth("Units"): Printer.Print "Units"
        Printer.FontBold = False
        lc = 8
        pgrid.Clear: pgrid.Rows = Int(tmpgrid.Rows / 2) + 1: pgrid.Cols = 8
        For i = 1 To pgrid.Rows - 1
            k = i + pgrid.Rows - 1
            pgrid.TextMatrix(i, 0) = tmpgrid.TextMatrix(i, 0)
            pgrid.TextMatrix(i, 1) = tmpgrid.TextMatrix(i, 1)
            If jobtrail Then                                                    'jv100609
                pgrid.TextMatrix(i, 2) = CInt(Val(tmpgrid.TextMatrix(i, 3)))
            Else                                                                'jv100609
                pgrid.TextMatrix(i, 2) = tmpgrid.TextMatrix(i, 2)     'jv022811
            End If                                                              'jv100609
            pgrid.TextMatrix(i, 3) = tmpgrid.TextMatrix(i, 4)
            tu = tu + Val(tmpgrid.TextMatrix(i, 4))
            tw = tw + Val(tmpgrid.TextMatrix(i, 3))
            'tp = tp + Val(tmpgrid.TextMatrix(i, 2))
            If tmpgrid.TextMatrix(i, 2) <> "Partial" Then
                tp = tp + 1                                 'jv022811
            End If
            If k < tmpgrid.Rows Then
                pgrid.TextMatrix(i, 4) = tmpgrid.TextMatrix(k, 0)
                pgrid.TextMatrix(i, 5) = tmpgrid.TextMatrix(k, 1)
                If jobtrail Then                                                    'jv100609
                    pgrid.TextMatrix(i, 6) = CInt(Val(tmpgrid.TextMatrix(k, 3)))
                Else                                                                'jv100609
                    pgrid.TextMatrix(i, 6) = tmpgrid.TextMatrix(k, 2) 'jv022811
                End If                                                              'jv100609
                pgrid.TextMatrix(i, 7) = tmpgrid.TextMatrix(k, 4)
                tu = tu + Val(tmpgrid.TextMatrix(k, 4))
                tw = tw + Val(tmpgrid.TextMatrix(k, 3))
                'tp = tp + Val(tmpgrid.TextMatrix(k, 2))
                If tmpgrid.TextMatrix(k, 2) <> "Partial" Then
                    tp = tp + 1                             'jv022811
                End If
            End If
        Next i
        For i = 1 To pgrid.Rows - 1
            Printer.FontName = "Arial"
            Printer.CurrentX = 360: Printer.Print pgrid.TextMatrix(i, 0); " ";
            Printer.Print StrConv(pgrid.TextMatrix(i, 1), vbProperCase); " ";
            'If jobtrail Then 'jv100609
                Printer.CurrentX = p1 - Printer.TextWidth(pgrid.TextMatrix(i, 2))
                Printer.Print pgrid.TextMatrix(i, 2);
            'End If jv100609
            Printer.CurrentX = p2 - Printer.TextWidth(pgrid.TextMatrix(i, 3))
            Printer.Print pgrid.TextMatrix(i, 3);
            Printer.CurrentX = p3
            Printer.Print pgrid.TextMatrix(i, 4); " ";
            Printer.Print StrConv(pgrid.TextMatrix(i, 5), vbProperCase); " ";
            'If jobtrail Then jv100609
                Printer.CurrentX = p4 - Printer.TextWidth(pgrid.TextMatrix(i, 6))
                Printer.Print pgrid.TextMatrix(i, 6);
            'End If jv100609
            Printer.CurrentX = p5 - Printer.TextWidth(pgrid.TextMatrix(i, 7))
            Printer.Print pgrid.TextMatrix(i, 7)
        
            If lc > 54 Then
                Printer.NewPage
                pno = pno + 1
                Printer.Print "Page "; pno;
                Printer.CurrentX = 8600: Printer.Print "Policy Number ";
                Printer.FontBold = True
                Printer.FontUnderline = True
                Printer.FontBold = False
                Printer.FontUnderline = False
                Printer.Print " "
                lc = 2: scode = " ": bcode = "N": fcode = "N"
            End If
            lc = lc + 1
        Next i
    Else
        p1 = 1440 * 1.25 '2.25 '2.75
        p2 = 1440 * 5.25
        p3 = 1440 * 6.25 '5.75
        Printer.FontName = "Arial"
        Printer.FontSize = 12
        Printer.FontBold = True
        Printer.CurrentX = p1:  Printer.Print "SKU  Description";
        If jobtrail Then
            Printer.CurrentX = p2 - Printer.TextWidth("Wraps")
            Printer.Print "Wraps";
        Else                                                            'jv100609
            Printer.CurrentX = p2 - Printer.TextWidth("Pallets")        'jv100609
            Printer.Print "Pallets";                                    'jv100609
        End If
        Printer.CurrentX = p3 - Printer.TextWidth("Units"): Printer.Print "Units"
        Printer.FontBold = False
        lc = 8
        pgrid.Clear: pgrid.Rows = Int(tmpgrid.Rows / 2) + 1: pgrid.Cols = 8
        For i = 1 To pgrid.Rows - 1
            k = i + pgrid.Rows - 1
            pgrid.TextMatrix(i, 0) = tmpgrid.TextMatrix(i, 0)
            pgrid.TextMatrix(i, 1) = tmpgrid.TextMatrix(i, 1)
            'Pgrid.TextMatrix(i, 2) = CInt(Val(tmpgrid.TextMatrix(i, 3)))
            pgrid.TextMatrix(i, 2) = tmpgrid.TextMatrix(i, 3)         'jv022811
            pgrid.TextMatrix(i, 3) = tmpgrid.TextMatrix(i, 4)
            tu = tu + Val(tmpgrid.TextMatrix(i, 4))
            tw = tw + Val(tmpgrid.TextMatrix(i, 3))
            'tp = tp + Val(tmpgrid.TextMatrix(i, 2))
            If tmpgrid.TextMatrix(i, 2) <> "Partial" Then
                tp = tp + 1         'jv022811
            End If
            If k < tmpgrid.Rows Then
                pgrid.TextMatrix(i, 4) = tmpgrid.TextMatrix(k, 0)
                pgrid.TextMatrix(i, 5) = tmpgrid.TextMatrix(k, 1)
                'Pgrid.TextMatrix(i, 6) = CInt(Val(tmpgrid.TextMatrix(k, 3)))
                pgrid.TextMatrix(i, 6) = tmpgrid.TextMatrix(k, 3)     'jv022811
                pgrid.TextMatrix(i, 7) = tmpgrid.TextMatrix(k, 4)
                tu = tu + Val(tmpgrid.TextMatrix(k, 4))
                tw = tw + Val(tmpgrid.TextMatrix(k, 3))
                'tp = tp + Val(tmpgrid.TextMatrix(k, 2))
                If tmpgrid.TextMatrix(k, 2) <> "Partial" Then
                    tp = tp + 1     'jv022811
                    pflag = True
                End If
            End If
        Next i
        For i = 1 To tmpgrid.Rows - 1
            Printer.FontName = "Arial"
            Printer.CurrentX = p1
            Printer.Print tmpgrid.TextMatrix(i, 0); " ";
            Printer.Print StrConv(tmpgrid.TextMatrix(i, 1), vbProperCase); " ";
            If jobtrail Then
                Printer.CurrentX = p2 - Printer.TextWidth(tmpgrid.TextMatrix(i, 3))
                Printer.Print tmpgrid.TextMatrix(i, 3);
            Else                                                                    'jv100609
                If Val(tmpgrid.TextMatrix(i, 2)) >= 1 Or tmpgrid.TextMatrix(i, 2) = "Partial" Then                            'jv100609
                    'k = Format(Val(tmpgrid.TextMatrix(i, 2)), "0")                    'jv100609
                    'Printer.CurrentX = p2 - Printer.TextWidth(k)                    'jv100609
                    'Printer.Print k;                                                'jv100609
                    Printer.CurrentX = p2 - Printer.TextWidth(tmpgrid.TextMatrix(i, 2)) 'jv100609
                    Printer.Print tmpgrid.TextMatrix(i, 2);                           'jv100609
                End If                                                              'jv100609
            End If
            Printer.CurrentX = p3 - Printer.TextWidth(tmpgrid.TextMatrix(i, 4))
            Printer.Print tmpgrid.TextMatrix(i, 4)
        
            If lc > 54 Then
                Printer.NewPage
                pno = pno + 1
                Printer.Print "Page "; pno;
                Printer.CurrentX = 8600: Printer.Print "Policy Number ";
                Printer.FontBold = True
                Printer.FontUnderline = True
                Printer.FontBold = False
                Printer.FontUnderline = False
                Printer.Print " "
                lc = 2: scode = " ": bcode = "N": fcode = "N"
            End If
            lc = lc + 1
        Next i
    End If
    Printer.Print " "
    If tmpgrid.Rows > 37 Then '46 Then
        Printer.CurrentX = p3: Printer.Print "Total Units";
        If jobtrail Then
            Printer.CurrentX = p4 - Printer.TextWidth(Format(tw, "#,###,###"))
            Printer.Print Format(tw, "#,###,###");
        End If
        Printer.CurrentX = p5 - Printer.TextWidth(Format(tu, "#,###,###")): Printer.Print Format(tu, "#,###,###")
    Else
        Printer.CurrentX = p1: Printer.Print "Total Units";
        If jobtrail Then
            Printer.CurrentX = p2 - Printer.TextWidth(Format(tw, "#,###,###"))
            Printer.Print Format(tw, "#,###,###");
        End If
        Printer.CurrentX = p3 - Printer.TextWidth(Format(tu, "#,###,###")): Printer.Print Format(tu, "#,###,###")
    End If
    lc = lc + 2
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.CurrentY = 1440 * 9
    For i = lc To 50 '54 '45 '50 '57
        Printer.Print " "
    Next i
    Printer.CurrentX = 720: Printer.Print "Ship Date:";
    Printer.CurrentX = 1440 * 1.5: Printer.Print Text1; 'Edittrl.sd;
    Printer.CurrentX = 1440 * 3: Printer.Print "Trailer #:";
    Printer.CurrentX = 1440 * 4: Printer.Print tc;
    Printer.CurrentX = 1440 * 5: Printer.Print "Total Pallets:";
    Printer.CurrentX = 1440 * 6: Printer.Print tp  'Int(tp + 0.8)
    Printer.Print " "
    Printer.CurrentX = 720: Printer.Print "Inspected By:";                          'jv082415
    Printer.CurrentX = 1440 * 1.5: Printer.Print "_____________________________";
    Printer.CurrentX = 1440 * 4: Printer.Print "Completed By:";
    Printer.CurrentX = 1440 * 5: Printer.Print "_____________________________"
    Printer.Print " "
    Printer.CurrentX = 720: Printer.Print "Seal #:";
    Printer.CurrentX = 1440 * 1.5: Printer.Print "_____________________________";
    Printer.CurrentX = 1440 * 4: Printer.Print "Sealed By:";
    Printer.CurrentX = 1440 * 5: Printer.Print "_____________________________"
    Printer.Print " "
    Printer.CurrentX = 720: Printer.Print "Driver:";
    Printer.CurrentX = 1440 * 1.5: Printer.Print "_____________________________";
    Printer.CurrentX = 1440 * 4: Printer.Print "Freight:";
    Printer.CurrentX = 1440 * 5: Printer.Print "_____________________________"
    Printer.Print " "
    Printer.CurrentX = 720: Printer.Print "Special Instructions:";
    Printer.CurrentX = 1440 * 2: Printer.Print "____________________________________________________________________"
    'db.Close
        
    Printer.NewPage
    Call prtpage2(Printer, pflag, wflag)
    Printer.EndDoc
    Printer.Duplex = 1
        
    'Screen.MousePointer = 0
    'Exit Sub
    'Turn off for testing   jv022811
    sqlx = "Update trailers set pb_flag = 'Y' where runid = " & runno
    Sdb.Execute sqlx
    Open Form1.tempdir & "/billtrl.prn" For Append As #1
    If bno = "16" Or bno = "15" Then
        'Write #1, sccode; Right(Combo1, 2); sd; tc
        'Write #1, sccode; Right(Grid1.TextMatrix(Grid1.Row, 5), 2); sd; tc
        Write #1, sccode; Right(Grid1.TextMatrix(Grid1.Row, 5), 2); Text1; tc
    Else
        'Write #1, Format(bno, "00"); Right(Combo1, 2); sd; tc
        'Write #1, Format(bno, "00"); Right(Grid1.TextMatrix(Grid1.Row, 5), 2); sd; tc
        Write #1, Format(bno, "00"); Right(Grid1.TextMatrix(Grid1.Row, 5), 2); Text1; tc
    End If
    Close #1
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "duplex_bill_log", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " duplex_bill_log - Error Number: " & eno
        End
    End If
End Sub

Private Sub prtpage2(pd As Control, pallets As Boolean, wraps As Boolean)
    Dim dl As String, s As String, i As Long
    Dim xs As Long, xe As Long, st As Long
    xs = 1440 * 0.25
    xe = 1440 * 8
    dl = "_________________________"
    'pd.Height = 1440 * 11
    'pd.Width = 1440 * 8.5
    pd.FontName = "Arial"
    pd.FontSize = 10
    If TypeOf pd Is Printer Then
        pd.DrawWidth = 6
    Else
        pd.DrawWidth = 1
    End If
    pd.Print " ": pd.Print " "
    pd.Print " ": pd.Print " "
    pd.Print " ": pd.Print " "
    s = "DRIVER INFORMATION"
    pd.FontBold = True
    pd.CurrentX = 1440 * 4 - (pd.TextWidth(s) * 0.5)
    pd.Print s
    pd.FontBold = False
    pd.Print " ": pd.Print " "
    st = pd.CurrentY
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 2.5: pd.Print "Driver #1";
    pd.CurrentX = 1440 * 4.5: pd.Print "Driver #2";
    pd.CurrentX = 1440 * 6.5: pd.Print "Driver #3"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Driver Name"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Starting Location"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Date"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Destination"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Seal #"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Depart temp."
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Mid trip temp."             'jv022717
    pd.Print " "                                                    'jv022717
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)                     'jv022717
    pd.Print " "                                                    'jv022717
    pd.CurrentX = 1440 * 0.5: pd.Print "Arrival temp."
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Signature"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 2: pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 4: pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 6: pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 8: pd.Line (xs, st)-(xs, pd.CurrentY)
    pd.Print " "
    pd.Print " ": pd.Print " "
    s = "FINAL DESTINATION INFORMATION"
    pd.FontBold = True
    pd.CurrentX = 1440 * 4 - (pd.TextWidth(s) * 0.5)
    pd.Print s
    pd.FontBold = False

    
    pd.Print " ": pd.Print " "
    pd.CurrentX = 720: pd.Print "Arrival Date:";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "Arrival temperature:";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Seal #:";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "Verified by:";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Time Arrived:";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "Time Departed:";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "# Pallets returned:";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "# Sleeves returned:";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Returns:";
    pd.CurrentX = 1440 * 2: pd.Print dl & dl & dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Comments:";
    pd.CurrentX = 1440 * 2: pd.Print dl & dl & dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Corrections:";
    pd.CurrentX = 1440 * 2: pd.Print dl & dl & dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Received by:";
    pd.CurrentX = 1440 * 2: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Batch Ticket:";
    pd.CurrentX = 1440 * 2
    If pallets = True And wraps = True Then
        pd.Print Grid1.TextMatrix(Grid1.Row, 17) & "P " & Grid1.TextMatrix(Grid1.Row, 17) & "W"
    Else
        If pallets = True Then
            pd.Print Grid1.TextMatrix(Grid1.Row, 17) & "P"
        Else
            If wraps = True Then
                pd.Print Grid1.TextMatrix(Grid1.Row, 17) & "W"
            Else
                pd.Print Grid1.TextMatrix(Grid1.Row, 17)
            End If
        End If
    End If
End Sub

Private Sub postr12_log(mplant As String, sdate As String)
    Dim ofile As String, s As String, rfile As String
    Dim f1 As String, f2 As String, f3 As String, f4 As String, f5 As String
    Dim f6 As String, f7 As String, f8 As String, f9 As String, f10 As String
    Dim f11 As String, f12 As String, f13 As String, f14 As String, f15 As String
    Dim f16 As String, f17 As String
    Dim i As Integer, k As Integer, addfile As Boolean, ftpexe As String
    Dim x
    ftplist.Clear: ftplist.AddItem "..."
    Dim ds As adodb.Recordset
    On Error GoTo vberror
    If mplant = "50" Then
        morg = "500"
        mwhs = "T10"
    End If
    If mplant = "51" Then
        morg = "501"
        mwhs = "K10"
    End If
    If mplant = "52" Then
        morg = "502"
        mwhs = "A10"
    End If
    ofile = Form1.pallogs & "R12" & sdate & ".txt"
    Open ofile For Append As #1
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 14) = "PRINTED" Then
            s = "select * from trailers where runid = " & Grid1.TextMatrix(i, 17)
            Set ds = Sdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                rfile = Form1.pallogs & "RO" & ds!runid & ".txt"
                addfile = True
                For k = 0 To ftplist.ListCount - 1
                    If ftplist.List(k) = ds!runid Then addfile = False
                Next k
                If addfile = True Then ftplist.AddItem ds!runid
                Open rfile For Append As #5
                If Grid1.TextMatrix(i, 7) > "00" Then  'barcode indicates pallet
                    Write #1, ds!runid & "P";: Write #5, ds!runid & "P";
                    Write #1, morg;: Write #5, morg;
                    Write #1, mwhs;: Write #5, mwhs;
                    Write #1, "FLOOR" & mwhs;: Write #5, "FLOOR" & mwhs;
                    If ds!branch = 47 Then
                        Write #1, "501"; "K10"; "FLOORK10";
                        Write #5, "501"; "K10"; "FLOORK10";
                    Else
                        If ds!branch = 52 Then
                            Write #1, "502"; "A10"; "FLOORA10";
                            Write #5, "502"; "A10"; "FLOORA10";
                        Else
                            If ds!branch = 1 Then
                                Write #1, "500"; "T10"; "FLOORT10";
                                Write #5, "500"; "T10"; "FLOORT10";
                            Else
                                Write #1, Format(ds!branch, "000");
                                Write #1, Format(ds!branch, "000");
                                Write #1, "FLOOR" & Format(ds!branch, "000");
                                Write #5, Format(ds!branch, "000");
                                Write #5, Format(ds!branch, "000");
                                Write #5, "FLOOR" & Format(ds!branch, "000");
                            End If
                        End If
                    End If
                    Write #1, ds!account;: Write #5, ds!account;
                    Write #1, Trim(Left(Grid1.TextMatrix(i, 7), 4));
                    Write #5, Trim(Left(Grid1.TextMatrix(i, 7), 4));
                    'Write #1, mid(Grid1.TextMatrix(i, 7), 5, 8);                'lot
                    'Write #5, mid(Grid1.TextMatrix(i, 7), 5, 8);                'lot
                    Write #1, Trim(mid(Grid1.TextMatrix(i, 7), 5, 9));                'jv052515 lot
                    Write #5, Trim(mid(Grid1.TextMatrix(i, 7), 5, 9));                'jv052515 lot
                    Write #1, Format(Val(Grid1.TextMatrix(i, 11)), "0");
                    Write #5, Format(Val(Grid1.TextMatrix(i, 11)), "0");
                    Write #1, "EACH";: Write #5, "EACH";
                    Write #1, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                    Write #1, StrConv(Grid1.TextMatrix(i, 5), vbProperCase) & " " & Right(Grid1.TextMatrix(i, 7), 3);
                    Write #5, StrConv(Grid1.TextMatrix(i, 5), vbProperCase) & " " & Right(Grid1.TextMatrix(i, 7), 3);
                    Write #1, Format(ds!shipdate, "MM-dd-yyyy");
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy");
                    If Left(ds!trlno, 1) = "B" Or ds!branch = 15 Or ds!branch = 16 Then
                        Write #1, "Y": Write #5, "Y"
                    Else
                        Write #1, "N": Write #5, "N"
                    End If
                    'Write #1, "N":Write #5, "N"
                        
                    If Val(Grid1.TextMatrix(i, 12)) > 0 Then   '2nd lot
                        f7 = Grid1.TextMatrix(i, 7)
                        f10 = Grid1.TextMatrix(i, 10)
                        f12 = Grid1.TextMatrix(i, 12)
                        s = mid(f7, 5, 2) & "-" & mid(f7, 7, 2) & "-20" & mid(f7, 9, 2)
                        s = Format(DateAdd("d", Val(f12) - Val(f10), s), "MMddyy")
                        's = s & mid(f7, 11, 2)
                        's = s & RTrim(mid(f7, 11, 3))                                   'jv052515
                        s = s & RTrim(mid(f12, 6, 3))                                       'jv083115
                        Write #1, ds!runid & "P";: Write #5, ds!runid & "P";
                        Write #1, morg;: Write #5, morg;
                        Write #1, mwhs;: Write #5, mwhs;
                        Write #1, "FLOOR" & mwhs;: Write #5, "FLOOR" & mwhs;
                        If ds!branch = 47 Then
                            Write #1, "501"; "K10"; "FLOORK10";
                            Write #5, "501"; "K10"; "FLOORK10";
                        Else
                            If ds!branch = 52 Then
                                Write #1, "502"; "A10"; "FLOORA10";
                                Write #5, "502"; "A10"; "FLOORA10";
                            Else
                                If ds!branch = 1 Then
                                    Write #1, "500"; "T10"; "FLOORT10";
                                    Write #5, "500"; "T10"; "FLOORT10";
                                Else
                                    Write #1, Format(ds!branch, "000");
                                    Write #5, Format(ds!branch, "000");
                                    Write #1, Format(ds!branch, "000");
                                    Write #5, Format(ds!branch, "000");
                                    Write #1, "FLOOR" & Format(ds!branch, "000");
                                    Write #5, "FLOOR" & Format(ds!branch, "000");
                                End If
                            End If
                        End If
                        Write #1, ds!account;: Write #5, ds!account;
                        Write #1, Trim(Left(Grid1.TextMatrix(i, 7), 4));
                        Write #5, Trim(Left(Grid1.TextMatrix(i, 7), 4));
                        's = r12_lot(f12, mid(f7, 12, 1))        'jv020614
                        s = r12_lot(f12, Trim(mid(f7, 11, 3)))        'jv052515
                        Write #1, s;: Write #5, s;
                        Write #1, Format(Val(Grid1.TextMatrix(i, 13)), "0");
                        Write #5, Format(Val(Grid1.TextMatrix(i, 13)), "0");
                        Write #1, "EACH";: Write #5, "EACH";
                        Write #1, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                        Write #5, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                        Write #1, StrConv(Grid1.TextMatrix(i, 5), vbProperCase) & " " & Right(f7, 3);
                        Write #5, StrConv(Grid1.TextMatrix(i, 5), vbProperCase) & " " & Right(f7, 3);
                        Write #1, Format(ds!shipdate, "MM-dd-yyyy");
                        Write #5, Format(ds!shipdate, "MM-dd-yyyy");
                        If Left(ds!trlno, 1) = "B" Or ds!branch = 15 Or ds!branch = 16 Then
                            Write #1, "Y": Write #5, "Y"
                        Else
                            Write #1, "N": Write #5, "N"
                        End If
                        'Write #1, "N": Write #5, "N"
                    End If
                Else
                    Write #1, ds!runid & "W";: Write #5, ds!runid & "W";
                    If mplant = "50" Then Write #1, "001"; "001"; "FLOOR001";
                    If mplant = "50" Then Write #5, "001"; "001"; "FLOOR001";
                    If mplant = "51" Then Write #1, "047"; "047"; "FLOOR047";
                    If mplant = "51" Then Write #5, "047"; "047"; "FLOOR047";
                    If mplant = "52" Then Write #1, "052"; "052"; "FLOOR052";
                    If mplant = "52" Then Write #5, "052"; "052"; "FLOOR052";
                    Write #1, Format(ds!branch, "000");
                    Write #5, Format(ds!branch, "000");
                    Write #1, Format(ds!branch, "000");
                    Write #5, Format(ds!branch, "000");
                    Write #1, "FLOOR" & Format(ds!branch, "000");
                    Write #5, "FLOOR" & Format(ds!branch, "000");
                    Write #1, ds!account;: Write #5, ds!account;
                    Write #1, Trim(Left(Grid1.TextMatrix(i, 6), 4));
                    Write #5, Trim(Left(Grid1.TextMatrix(i, 6), 4));
                    Write #1, "LOT1";: Write #5, "LOT1";
                    Write #1, Format(Val(Grid1.TextMatrix(i, 11)), "0");
                    Write #5, Format(Val(Grid1.TextMatrix(i, 11)), "0");
                    Write #1, "EACH";: Write #5, "EACH";
                    Write #1, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                    Write #1, StrConv(Grid1.TextMatrix(i, 5), vbProperCase);
                    Write #5, StrConv(Grid1.TextMatrix(i, 5), vbProperCase);
                    Write #1, Format(ds!shipdate, "MM-dd-yyyy");
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy");
                    If Left(ds!trlno, 1) = "B" Or ds!branch = 15 Or ds!branch = 16 Then
                        Write #1, "Y": Write #5, "Y"
                    Else
                        Write #1, "N": Write #5, "N"
                    End If
                End If
            End If
            Close #5
            ds.Close
            Grid1.TextMatrix(i, 14) = "POSTED"
        End If
    Next i
        
    Close #1
 
    addfile = False
    ofile = Form1.pallogs & "r12trls.win"
    Open ofile For Output As #1
    Print #1, "open pbelle.bluebell.com"
    Print #1, "infbbcri"
    Print #1, "welcome@2023"
    Print #1, "BINARY"
    'Print #1, "cd /interface/infbbcri/PBELLE/incoming"
    Print #1, "cd PBELLE/incoming"
    Print #1, "lcd "; Left(Form1.pallogs, Len(Form1.pallogs) - 1)
    For i = 0 To ftplist.ListCount - 1
        If ftplist.List(i) > "0" Then
            rfile = Form1.pallogs & "RO" & ftplist.List(i) & ".txt"
            If Len(Dir(rfile)) > 0 Then
                s = "put RO" & ftplist.List(i) & ".txt RO" & ftplist.List(i) & ".txt"
                Print #1, s
                addfile = True
            End If
            s = "Update trailers set pb_flag = 'Y', ra_flag = 'Y' Where runid = " & ftplist.List(i)
            Sdb.Execute s
        End If
    Next i
    Print #1, "close"
    Print #1, "bye"
    Close #1
    If addfile = True Then
        ftpexe = "c:\windows\system32\ftp.exe"
        x = Shell(ftpexe & " -s:" & ofile, vbNormalFocus)
        MsgBox ftpexe & " -s:" & ofile
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "postr12_log", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " postr12_log - Error Number: " & eno
        End
    End If
End Sub

Public Sub refresh_grid1(sd As String)
    Dim cfile As String, s As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim logpath As String
    Text1 = sd
    logpath = Form1.pallogs
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 19: Grid1.Redraw = False: Grid1.Visible = False
   
    cfile = logpath & "bill" & Format(sd, "mmddyyyy") & ".txt"
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input Shared As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 ', f17
            s = "B" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
            s = s & Trim(StrConv(f4, vbProperCase)) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
            s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & Trim(StrConv(f4, vbProperCase)) & f15 'f16
            Grid1.AddItem s
        Loop
        Close #1
    End If

    Grid1.Row = 0: Grid1.RowSel = 0: Grid1.Col = 18: Grid1.ColSel = 18
    Grid1.Sort = 5
    If Check1.Value = 1 Then
        s = "^Type|^Batch|^Area|<Description|^Source|<Target|<Product|^BarCode|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId|<SortKey"
        Grid1.FormatString = s
        Grid1.ColWidth(0) = 600
        Grid1.ColWidth(1) = 700
        Grid1.ColWidth(2) = 1300
        Grid1.ColWidth(3) = 1000
        Grid1.ColWidth(4) = 1300
        Grid1.ColWidth(5) = 1800
        Grid1.ColWidth(6) = 3000
        Grid1.ColWidth(7) = 1800
        Grid1.ColWidth(8) = 600
        Grid1.ColWidth(9) = 800
        Grid1.ColWidth(10) = 800
        Grid1.ColWidth(11) = 800
        Grid1.ColWidth(12) = 800
        Grid1.ColWidth(13) = 800
        Grid1.ColWidth(14) = 800
        Grid1.ColWidth(15) = 1000
        Grid1.ColWidth(16) = 1400
        Grid1.ColWidth(17) = 1000
        Grid1.ColWidth(18) = 1000
    Else
        s = "^Type|^Batch||||<Target|<Product|^BarCode|^Qty|^Uom|||||^Status||||"
        Grid1.FormatString = s
        Grid1.ColWidth(0) = 600
        Grid1.ColWidth(1) = 700
        Grid1.ColWidth(2) = 1 '1300
        Grid1.ColWidth(3) = 1 '1000
        Grid1.ColWidth(4) = 1 '1300
        Grid1.ColWidth(5) = 2500
        Grid1.ColWidth(6) = 3000
        Grid1.ColWidth(7) = 1800
        Grid1.ColWidth(8) = 800
        Grid1.ColWidth(9) = 800
        Grid1.ColWidth(10) = 1 '800
        Grid1.ColWidth(11) = 1 '800
        Grid1.ColWidth(12) = 1 '800
        Grid1.ColWidth(13) = 1 '800
        Grid1.ColWidth(14) = 1000
        Grid1.ColWidth(15) = 1 '1000
        Grid1.ColWidth(16) = 1 '1400
        Grid1.ColWidth(17) = 1 '1000
        Grid1.ColWidth(18) = 1 '1000
    End If
    hcolor.Caption = "All Records"
    cntlit.Caption = Grid1.Rows - 1 & " Records"
    Combo1.Clear: runlist.Clear
    
    If Grid1.Rows > 1 Then
        s = Grid1.TextMatrix(1, 17)
        Combo1.AddItem Grid1.TextMatrix(1, 5)
        runlist.AddItem Grid1.TextMatrix(1, 17)
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 17) <> s Then
                s = Grid1.TextMatrix(i, 17)
                Combo1.AddItem Grid1.TextMatrix(i, 5)
                runlist.AddItem Grid1.TextMatrix(i, 17)
            End If
        Next i
    End If
    
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        s = Grid1.TextMatrix(1, 17)
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 17) <> s Then
                s = Grid1.TextMatrix(i, 17)
                Grid1.AddItem " ", i
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = Grid1.BackColorBkg  'gcolor.BackColor
            End If
        Next i
        Combo1.ListIndex = 0
    End If
    
    
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            If UCase(Grid1.TextMatrix(i, 14)) = "PRINTED" Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 14: Grid1.ColSel = 14
                Grid1.CellBackColor = gcolor.BackColor
                Grid1.CellForeColor = gcolor.ForeColor
            End If
            If UCase(Grid1.TextMatrix(i, 14)) = "POSTED" Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 14: Grid1.ColSel = 14
                Grid1.CellBackColor = rcolor.BackColor
                Grid1.CellForeColor = rcolor.ForeColor
            End If
        Next i
    End If
    Grid1.Redraw = True
    Grid1.Visible = True
    If Grid1.Rows > 1 Then Grid1.Row = 1
    rundate.Caption = sd
End Sub

Private Sub save_bills(runno As String)
    Dim cfile As String, s As String, i As Integer
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim logpath As String
    logpath = Form1.pallogs
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 17
   
    cfile = logpath & "bill" & Format(rundate, "MMddyyyy") & ".txt"
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input Shared As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
            If f16 <> runno Then
                s = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                Grid3.AddItem s
            End If
        Loop
        Close #1
    End If
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 17) = runno And Grid1.TextMatrix(i, 14) <> "CANC" Then
            s = Grid1.TextMatrix(i, 1)
            For k = 2 To 17
                s = s & Chr(9) & Grid1.TextMatrix(i, k)
            Next k
            Grid3.AddItem s
        End If
    Next i
    
    cfile = logpath & "bill" & Format(rundate, "MMddyyyy") & ".txt"
    Open cfile For Output As #1
    For i = 1 To Grid3.Rows - 1
        For k = 0 To Grid3.Cols - 2
            Write #1, Grid3.TextMatrix(i, k);
        Next k
        Write #1, Grid3.TextMatrix(i, Grid3.Cols - 1)
    Next i
    Close #1
End Sub


Private Sub check_totals(runno As String)
    Dim ds As adodb.Recordset, s As String, adflag As Boolean
    On Error GoTo vberror
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 5
    Grid2.Visible = False
    ycolor.Visible = False
    If Val(runno) = 0 Then Exit Sub
    s = "select runid, sku, sum(units) from trailers where runid = " & runno
    s = s & " group by runid, sku"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(0) & Chr(9) & ds(1) & Chr(9) & ds(2)
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.Redraw = False
    Grid1.FillStyle = flexFillRepeat
    Grid2.FillStyle = flexFillRepeat
    If Grid2.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 17) = Grid2.TextMatrix(1, 0) And Grid1.TextMatrix(i, 14) <> "CANC" Then
                adflag = True
                For k = 1 To Grid2.Rows - 1
                    If Grid2.TextMatrix(k, 1) = Left(Grid1.TextMatrix(i, 6), 3) Then
                        Grid2.TextMatrix(k, 3) = Val(Grid2.TextMatrix(k, 3)) + Val(Grid1.TextMatrix(i, 11))
                        Grid2.TextMatrix(k, 3) = Val(Grid2.TextMatrix(k, 3)) + Val(Grid1.TextMatrix(i, 13))
                        adflag = False
                        Exit For
                    End If
                Next k
                If adflag = True Then
                    s = Grid2.TextMatrix(1, 0) & Chr(9)
                    s = s & Left(Grid1.TextMatrix(i, 6), 3) & Chr(9)
                    s = s & "0" & Chr(9)
                    k = Val(Grid1.TextMatrix(i, 11))
                    k = k + Val(Grid1.TextMatrix(i, 13))
                    s = s & k
                    Grid2.AddItem s
                End If
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 6: Grid1.ColSel = 13
                Grid1.CellBackColor = Grid1.BackColor
                Grid1.Col = 1
            End If
        Next i
                
        For i = 1 To Grid2.Rows - 1
            Grid2.TextMatrix(i, 4) = Val(Grid2.TextMatrix(i, 3)) - Val(Grid2.TextMatrix(i, 2))
            If Val(Grid2.TextMatrix(i, 4)) <> 0 Then
                For k = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(k, 17) = Grid2.TextMatrix(i, 0) And Left(Grid1.TextMatrix(k, 6), 3) = Grid2.TextMatrix(i, 1) Then
                        Grid1.Row = k: Grid1.RowSel = k
                        Grid1.Col = 6: Grid1.ColSel = 13
                        Grid1.CellBackColor = ycolor.BackColor
                        Grid1.Col = 1
                    End If
                Next k
                Grid2.Row = i: Grid2.RowSel = i
                Grid2.Col = 1: Grid2.ColSel = Grid2.Cols - 1
                Grid2.CellBackColor = ycolor.BackColor
                Grid2.TopRow = i: Grid2.Col = 1: Grid2.ColSel = 2
                Grid2.Visible = True
                ycolor.Visible = True
            End If
        Next i
                
    End If
                
    Grid1.Redraw = True
    Grid2.FormatString = "^RunId|^SKU|^Ordered|^Scanned|^Diff"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 1000
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "check_totals", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " check_totals - Error Number: " & eno
        End
    End If
End Sub

Private Sub addbc_Click()
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String
    Dim cfile As String, s As String, bc As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    
    Dim t0 As String, t1 As String, t2 As String, t3 As String
    Dim t4 As String, t5 As String, t6 As String, t7 As String
    Dim t8 As String, t9 As String, t10 As String, t11 As String
    Dim t12 As String, t13 As String, t14 As String, t15 As String
    
    Dim dl As Long, wbc As String
    Dim logpath As String
    logpath = Form1.pallogs
    If Val(Grid1.TextMatrix(Grid1.Row, 17)) < 1 Then Exit Sub
    wbc = Grid1.TextMatrix(Grid1.Row, 7)
    wbc = InputBox("Enter a BarCode to search for:", "BarCode Example....", wbc)
    If Len(wbc) = 0 Then Exit Sub
    wbc = UCase(wbc)
    If wbc = Grid1.TextMatrix(Grid1.Row, 7) And Grid1.TextMatrix(Grid1.Row, 14) <> "CANC" Then
        MsgBox "BarCode is already on this bill.", vbOKOnly + vbInformation, "Duplicate barcode..."
        Exit Sub
    End If
    Screen.MousePointer = 11
    t10 = "0"
    'Look for barcode in movement log
    sdate = Format(DateAdd("d", -1, Text1), "MM-dd-yyyy")
    edate = Format(Text1, "MM-dd-yyyy")
    sdate = Format(sdate, "yyyymmdd")
    edate = Format(edate, "yyyymmdd")
    spath = logpath & "move*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = wbc Then
                    t0 = f0: t1 = f1: t2 = f2: t3 = f3: t4 = f4
                    t5 = f5: t6 = f6: t7 = f7: t8 = f8: t9 = f9
                    t10 = f10: t11 = f11: t12 = f12: t13 = f13: t14 = f14
                    t15 = f15: t16 = f16
                    s = f2 & " " & f4 & " " & f5
                    MsgBox s, vbOKOnly + vbInformation, f15 & " received...... " & f6
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    
    If Val(t10) = 0 Then
        'Look for barcodes in shipping tasks
        spath = logpath & "ship*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then
                        t0 = f0: t1 = f1: t2 = f2: t3 = f3: t4 = f4
                        t5 = f5: t6 = f6: t7 = f7: t8 = f8: t9 = f9
                        t10 = f10: t11 = f11: t12 = f12: t13 = f13: t14 = f14
                        t15 = f15: t16 = f16
                        s = f2 & " " & f4 & " " & f5
                        MsgBox s, vbOKOnly + vbInformation, f15 & " shipped...... " & f6
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    
    If Val(t10) = 0 Then
        'Look for barcodes at wrappers
        sdate = mid(wbc, 5, 2) & "-" & mid(wbc, 7, 2) & "-20" & Format(Val(mid(wbc, 9, 2)) - 2, "00")
        edate = Format(DateAdd("d", 5, sdate), "MM-dd-yyyy")
        sdate = Format(sdate, "yyyymmdd")
        edate = Format(edate, "yyyymmdd")
        spath = logpath & "recv*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then
                        t0 = f0: t1 = f1: t2 = f2: t3 = f3: t4 = f4
                        t5 = f5: t6 = f6: t7 = f7: t8 = f8: t9 = f9
                        t10 = f10: t11 = f11: t12 = f12: t13 = f13: t14 = f14
                        t15 = f15: t16 = f16
                        s = f2 & " " & f4 & " " & f5
                        MsgBox s, vbOKOnly + vbInformation, f15 & " received...... " & f6
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    
    Screen.MousePointer = 0
    If Val(t10) <> 0 Then
        i = Grid1.Row
        s = "B" & Chr(9)
        s = s & Grid1.TextMatrix(i, 1) & Chr(9)
        s = s & Grid1.TextMatrix(i, 2) & Chr(9)
        s = s & Grid1.TextMatrix(i, 3) & Chr(9)
        s = s & "ADD" & Chr(9) 'Grid1.TextMatrix(i, 4) & Chr(9)
        s = s & Grid1.TextMatrix(i, 5) & Chr(9)
        s = s & t5 & Chr(9)
        s = s & t6 & Chr(9)
        s = s & t7 & Chr(9)
        s = s & t8 & Chr(9)
        s = s & t9 & Chr(9)
        s = s & t10 & Chr(9)
        s = s & t11 & Chr(9)
        s = s & t12 & Chr(9)
        s = s & "PEND" & Chr(9) 'Grid1.TextMatrix(i, 14) & Chr(9)
        s = s & "wms" & Chr(9) 'Grid1.TextMatrix(i, 15) & Chr(9)
        s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9) 'Grid1.TextMatrix(i, 16) & Chr(9)
        s = s & Grid1.TextMatrix(i, 17) & Chr(9)
        Grid1.AddItem s, i
        srun = Grid1.TextMatrix(i, 17)
        Call check_totals(srun)
        Grid1.Row = i
    End If
End Sub

Private Sub addwraps_Click()
    Dim ds As adodb.Recordset, s As String
    Dim wqty As Integer, uqty As Integer
    Dim sprod As String, wconv As Integer
    On Error GoTo vberror
    If Val(Grid1.TextMatrix(Grid1.Row, 17)) = 0 Then Exit Sub
    wconv = 0
    s = Trim(Left(Grid1.TextMatrix(Grid1.Row, 6), 4))
    s = InputBox("SKU:", "Add partial wraps...", s)
    If Len(s) = 0 Then Exit Sub
    s = "select * from skumast where sku = '" & s & "'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        sprod = ds!sku & " " & ds!fgunit & " " & ds!fgdesc
        wconv = ds!numwrap
    End If
    ds.Close
    If wconv = 0 Then Exit Sub
    wqty = InputBox("# Wraps:", "Add partial wraps...", "1")
    If wqty = 0 Then Exit Sub
    s = Grid1.TextMatrix(Grid1.Row, 0)
    For i = 1 To 5
        s = s & Chr(9) & Grid1.TextMatrix(Grid1.Row, i)
    Next i
    s = s & Chr(9) & sprod & Chr(9) & " " & Chr(9)
    s = s & wqty & Chr(9)
    s = s & "Wraps" & Chr(9)
    s = s & "LOT1" & Chr(9)
    s = s & Format(wconv * wqty, "0") & Chr(9)
    s = s & ".." & Chr(9) & "0" & Chr(9) & "PEND" & Chr(9)
    s = s & "wms" & Chr(9) & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
    s = s & Grid1.TextMatrix(Grid1.Row, 17)
    i = Grid1.Row
    Grid1.AddItem s, Grid1.Row
    srun = Grid1.TextMatrix(Grid1.Row, 17)
    Call check_totals(srun)
    Grid1.Row = i
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "addwraps_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " addwraps_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub canline_Click()
    Dim i As Integer
    If Grid1.TextMatrix(Grid1.Row, 14) = "POSTED" Then      'JV010313
        MsgBox "This line has been POSTED.", vbOKOnly + vbInformation, "Cancel is denied.."
        Exit Sub
    End If
    i = Grid1.Row
    Grid1.TextMatrix(Grid1.Row, 14) = "CANC"
    srun = Grid1.TextMatrix(Grid1.Row, 17)
    Call check_totals(srun)
    Grid1.Row = i
End Sub

Private Sub Check1_Click()
    Dim s As String
    If Check1.Value = 1 Then
        s = "^Type|^Batch|^Area|<Description|^Source|<Target|<Product|^BarCode|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId|<SortKey"
        Grid1.FormatString = s
        Grid1.ColWidth(0) = 600
        Grid1.ColWidth(1) = 700
        Grid1.ColWidth(2) = 1300
        Grid1.ColWidth(3) = 1000
        Grid1.ColWidth(4) = 1300
        Grid1.ColWidth(5) = 1800
        Grid1.ColWidth(6) = 3000
        Grid1.ColWidth(7) = 1800
        Grid1.ColWidth(8) = 600
        Grid1.ColWidth(9) = 800
        Grid1.ColWidth(10) = 800
        Grid1.ColWidth(11) = 800
        Grid1.ColWidth(12) = 800
        Grid1.ColWidth(13) = 800
        Grid1.ColWidth(14) = 800
        Grid1.ColWidth(15) = 1000
        Grid1.ColWidth(16) = 1400
        Grid1.ColWidth(17) = 1000
        Grid1.ColWidth(18) = 1000
    Else
        s = "^Type|^Batch||||<Target|<Product|^BarCode|^Qty|^Uom|||||^Status||||"
        Grid1.FormatString = s
        Grid1.ColWidth(0) = 600
        Grid1.ColWidth(1) = 700
        Grid1.ColWidth(2) = 1 '1300
        Grid1.ColWidth(3) = 1 '1000
        Grid1.ColWidth(4) = 1 '1300
        Grid1.ColWidth(5) = 2500
        Grid1.ColWidth(6) = 3000
        Grid1.ColWidth(7) = 1800
        Grid1.ColWidth(8) = 800
        Grid1.ColWidth(9) = 800
        Grid1.ColWidth(10) = 1 '800
        Grid1.ColWidth(11) = 1 '800
        Grid1.ColWidth(12) = 1 '800
        Grid1.ColWidth(13) = 1 '800
        Grid1.ColWidth(14) = 1000
        Grid1.ColWidth(15) = 1 '1000
        Grid1.ColWidth(16) = 1 '1400
        Grid1.ColWidth(17) = 1 '1000
        Grid1.ColWidth(18) = 1 '1000
    End If
    
End Sub

Private Sub Combo1_Click()
    trlkey.Caption = Combo1
End Sub

Private Sub Command1_Click()
    If Val(srun) > 0 Then
        If MsgBox("Save changes?", vbYesNo + vbQuestion, "save changes to bill..") = vbYes Then
            'MsgBox "saving"
            Call save_bills(srun)
        'Else
        '    MsgBox "refreshing"
        End If
        srun = "..."
    End If

    If Val(srun) = 0 Then refresh_grid1 (Text1)
End Sub

Private Sub Command2_Click()
    If Val(srun) > 0 Then
        If MsgBox("Save changes?", vbYesNo + vbQuestion, "save changes to bill..") = vbYes Then
            Call save_bills(srun)
        End If
        srun = "..."
    End If
End Sub

Private Sub edunits_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 17)) = 0 Then Exit Sub
    s = InputBox("Units for w/d lot " & Grid1.TextMatrix(i, 10), "1st Lot..", Grid1.TextMatrix(i, 11))
    If Len(s) <> 0 Then Grid1.TextMatrix(i, 11) = s
    If Val(Grid1.TextMatrix(i, 13)) <> 0 Then
        s = InputBox("Units for w/d lot " & Grid1.TextMatrix(i, 12), "2nd Lot..", Grid1.TextMatrix(i, 13))
        If Len(s) <> 0 Then Grid1.TextMatrix(i, 11) = s
    End If
    srun = Grid1.TextMatrix(i, 17)
    Call check_totals(Grid1.TextMatrix(i, 17))
    Grid1.Row = i
End Sub

Private Sub Form_Load()
    'Text1 = Format(Now, "m-d-yyyy")
    'Text1 = Edittrl.sd
    'refresh_grid1
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1500
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid1_RowColChange()
    If Grid1.Redraw = True Then
        If Grid1.TextMatrix(Grid1.Row, 5) > " " Then
            trlkey.Caption = Grid1.TextMatrix(Grid1.Row, 5)
        End If
    End If
End Sub

Private Sub postr12_Click()
    Dim i As Integer, k As Integer, cfile As String, logpath As String
    logpath = Form1.pallogs
    Screen.MousePointer = 11
    Call refresh_grid1(Text1)
    DoEvents
    Call postr12_log(Form1.plantno, Format(Text1, "MMddyyyy"))
    
    cfile = logpath & "bill" & Format(rundate, "MMddyyyy") & ".txt"
    Open cfile For Output As #1
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = "B" Then
            For k = 1 To Grid1.Cols - 3
                Write #1, Grid1.TextMatrix(i, k);
            Next k
            Write #1, Grid1.TextMatrix(i, Grid1.Cols - 2)
        End If
    Next i
    Close #1
    Call refresh_grid1(Text1)
    
    Screen.MousePointer = 0
End Sub

Private Sub prtbill_Click()
    If Val(Grid1.TextMatrix(Grid1.Row, 17)) = 0 Then Exit Sub
    If ycolor.Visible = True Then
        MsgBox ycolor.Caption, vbOKOnly + vbExclamation, "sorry, try again..."
        Exit Sub
    End If
    'Testing = mark as printed
    s = Grid1.TextMatrix(Grid1.Row, 17)
    Grid1.FillStyle = flexFillRepeat
    Grid1.Redraw = False
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 17) = s Then
            If Grid1.TextMatrix(i, 14) <> "POSTED" And Grid1.TextMatrix(i, 14) <> "CANC" Then
                Grid1.TextMatrix(i, 14) = "PRINTED"
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 14: Grid1.ColSel = 14
                Grid1.CellForeColor = gcolor.ForeColor
                Grid1.CellBackColor = gcolor.BackColor
            End If
        End If
    Next i
    Grid1.Redraw = True
    'Live
    Call duplex_bill_log(Grid1.TextMatrix(Grid1.Row, 17))
    DoEvents
    Call save_bills(Grid1.TextMatrix(Grid1.Row, 17))
End Sub

Private Sub rentrl_Click()
    If Val(srun) > 0 Then
        If MsgBox("Save changes?", vbYesNo + vbQuestion, "save changes to bill..") = vbYes Then
            'MsgBox "saving"
            Call save_bills(srun)
        'Else
        '    MsgBox "refreshing"
        End If
        srun = "..."
        Exit Sub
    End If

    If Val(Grid1.TextMatrix(Grid1.Row, 17)) = 0 Then Exit Sub
    Call rename_trailer(Grid1.TextMatrix(Grid1.Row, 17))
End Sub

Private Sub srun_Change()
    rundate = Text1
    If Val(srun) > 0 Then
        Command2.Visible = True
    Else
        Command2.Visible = False
    End If
End Sub

Private Sub tpost_Click()

End Sub

Private Sub trlkey_Change()
    Dim i As Integer, k As Integer, u As Long, j As Integer
    If Val(srun) > 0 Then
        If MsgBox("Save changes?", vbYesNo + vbQuestion, "save changes to bill..") = vbYes Then
            'MsgBox "saving"
            Call save_bills(srun)
        'Else
        '    MsgBox "refreshing"
        End If
        srun = "..."
    End If
    For i = 0 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 5) = trlkey Then
            Grid1.Row = i: Grid1.TopRow = i: Grid1.Col = 1
            j = Grid1.Row
            Exit For
        End If
    Next i
    k = 0: u = 0
    For i = Grid1.Row To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 5) = trlkey Then
            k = k + 1
            u = u + Val(Grid1.TextMatrix(i, 11))
            u = u + Val(Grid1.TextMatrix(i, 13))
        Else
            Exit For
        End If
    Next i
    hcolor.Caption = u & " Units"
    cntlit.Caption = k & " Records"
    Call check_totals(Grid1.TextMatrix(Grid1.Row, 17))
    'Grid1.Row = j: Grid1.TopRow = j: Grid1.Col = 1
    Grid1.Row = j: Grid1.Col = 1
End Sub

