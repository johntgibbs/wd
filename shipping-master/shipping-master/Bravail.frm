VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Bravail 
   Caption         =   "Homepage Updates"
   ClientHeight    =   10965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10785
   LinkTopic       =   "Form2"
   ScaleHeight     =   10965
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Stock Grid Test"
      Height          =   255
      Left            =   7080
      TabIndex        =   17
      Top             =   9240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   9960
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1508
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid grid2 
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   9240
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1296
      _Version        =   327680
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   8520
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Disable Branch Ordering"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enable Branch Ordering"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Out of Stock List "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command5 
         Caption         =   "Refresh"
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
         Left            =   3600
         TabIndex        =   10
         Top             =   8640
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Post To Homepage"
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
         Left            =   240
         TabIndex        =   9
         Top             =   8640
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   14631
         _Version        =   327680
         ForeColor       =   192
      End
   End
   Begin VB.TextBox Text2 
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
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Height          =   285
      Left            =   6000
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
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
      Left            =   6000
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
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
      Left            =   6000
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4320
      Width           =   6735
   End
   Begin VB.Label Label2 
      Caption         =   "Order Date 2:"
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
      Left            =   6000
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Order Date 1:"
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
      Left            =   6000
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Bravail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub outstk_sched()
    Dim i, userid As String, pwd As String, dsn As String, query As String
    Dim q As String, k As Integer, swhs As String
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    If Len(Dir(Form1.srserv & "\wd\bin\gemmodbc.ini")) <= 0 Then
        MsgBox "Gemmodbc.ini File not found in wd\bin directory!", vbOKOnly + vbExclamation, "Request cancelled"
        Exit Sub
    End If
    Open Form1.srserv & "\wd\bin\gemmodbc.ini" For Input As #1
    Line Input #1, dsn
    Line Input #1, userid
    Line Input #1, pwd
    Close #1
    If AllocateODBChEnv(hEnv) <> SQL_SUCCESS Then Exit Sub
    If ConnectToDataSource(hEnv, hdbc, hstmt, dsn, userid, pwd) <> SQL_SUCCESS Then
        i = FreeODBChEnv(hEnv)
        Exit Sub
    End If
    Grid2.Cols = 3: Grid2.Clear
    
    'R12
    q = "select itm.segment1,mtl.organization_code,MIN(plan_start_date)"
    q = q & " from gme_batch_header hdr,mtl_parameters mtl,"
    q = q & "mtl_system_items_b itm ,gme_material_details dtl"
    q = q & " Where trunc(hdr.plan_start_date) >= trunc(SYSDATE - 5)"
    q = q & " and hdr.batch_status in (1,2)"
    q = q & " and hdr.batch_id = dtl.batch_id"
    q = q & " and mtl.organization_id = hdr.organization_id"
    q = q & " and dtl.line_type = 1"
    q = q & " and dtl.inventory_item_id = itm.inventory_item_id"
    q = q & " and itm.organization_id = hdr.organization_id"
    q = q & " and itm.segment1 < '999'"
    q = q & " group by itm.segment1,mtl.organization_code"
    q = q & " order by 1,3"
    
    
    Screen.MousePointer = 11
    i = LoadGrid(Grid2, q, hstmt, 1, "")
    Grid2.FormatString = "^sku|^plant|^date"
    Grid2.ColWidth(0) = 800
    Grid2.ColWidth(1) = 800
    Grid2.ColWidth(2) = 1800
    Screen.MousePointer = 0
    i = DisconnectFromDataSource(hdbc, hstmt)
    i = FreeODBChEnv(hEnv)
    For i = 0 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 1) = "503" Then Grid2.TextMatrix(i, 1) = "500"
    Next i
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 2)) > 0 Then
            If Grid1.TextMatrix(i, 2) = "52" Then swhs = "502"
            If Grid1.TextMatrix(i, 2) = "51" Then swhs = "501"
            If Grid1.TextMatrix(i, 2) = "50" Then swhs = "500"
            For k = 0 To Grid2.Rows - 1
                If Grid2.TextMatrix(k, 0) = Grid1.TextMatrix(i, 0) And Grid2.TextMatrix(k, 1) = swhs Then
                    Grid1.TextMatrix(i, 3) = Format(Grid2.TextMatrix(k, 2), "m-d-yyyy")
                    Exit For
                End If
            Next k
        End If
    Next i
    'In Transit to Sylacauga
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 2)) = 52 Then
            s = "select * from trailers where plant = 50 and branch = 52"
            s = s & " and sku = '" & Grid1.TextMatrix(i, 0) & "'"
            s = s & " order by shipdate"
            Set ds = Sdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                Grid1.TextMatrix(i, 3) = "Transit - " & Format(ds!shipdate, "m-d-yyyy")
            End If
            ds.Close
        End If
    Next i
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "outstk_sched", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " outstk_sched - Error Number: " & eno
        End
    End If
End Sub

Private Sub rebuild_homegrid()
    Dim ds As adodb.Recordset, s As String
    Dim odates As String, scdates As String
    Dim rt As String, rh As String, rf As String
    On Error GoTo vberror
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 2: pgrid.FixedCols = 1
    pgrid.FixedCols = 0
    Set ds = Sdb.Execute("select * from wdstatus")
    If ds.BOF = False Then
        ds.MoveFirst
        odates = ds!orddates
        scdates = ds!schdates
    End If
    ds.Close
    s = "select * from branches where branch > 90 and brnmess > '   ' order by branch desc"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        's = "<img src=" & Chr(34) & "images\lanette.jpg" & Chr(34) & "><BR>" & Chr(9)
        s = "<img src=" & Chr(34) & "images\new.jpg" & Chr(34) & "><BR>" & Chr(9)
        Do Until ds.EOF
            s = s & "<b>" & ds!branchname & ":</b><br>" & ds!brnmess & "<hr>"
            'pgrid.AddItem s & ds!branchname & Chr(9) & ds!brnmess
            's = ""
            ds.MoveNext
        Loop
        pgrid.AddItem s
    End If
    ds.Close
    's = Chr(9)
    s = "<a href=" & Chr(34) & "stock.htm" & Chr(34) & ">"
    s = s & "<img src=" & Chr(34) & "images\bbstock.jpg" & Chr(34) & " Border=0><BR>Out of Stock Listings</a>"
    s = s & Chr(9) & "Last Updated: " & FileDateTime(Form1.webdir & "\stock.htm")
    pgrid.AddItem s
    
    's = Chr(9)
    's = "<img src=" & Chr(34) & "images\men in plant.tif" & Chr(34) & ">" & Chr(9)
    s = "<a href=" & Chr(34) & "discont.htm" & Chr(34) & ">Discontinued Products</a>"
    s = s & Chr(9) & "Last Updated: " & FileDateTime(Form1.webdir & "\discont.htm")
    pgrid.AddItem s
    
    's = Chr(9)
    s = "<a href=" & Chr(34) & "stock\wdstk.htm" & Chr(34) & ">"
    s = s & "<img src=" & Chr(34) & "images\wdstk.jpg" & Chr(34) & " Border=0><BR>Blue Bell Pallet Stacking Patterns</a>"
    s = s & Chr(9) & "Last Updated: " & FileDateTime(Form1.webdir & "\stock\wdstk.htm")
    pgrid.AddItem s
    
    's = Chr(9)
    's = "<img src=" & Chr(34) & "images\realtrail.jpg" & Chr(34) & ">" & Chr(9)
    s = "Branch Orders" & Chr(9)
    If Command2.Enabled = False Then
    'If MsgBox("enable orders", vbYesNo + vbQuestion, "testing....") = vbYes Then
        's = "<img src=" & Chr(34) & "images\greenlite.jpg" & Chr(34) & ">" & Chr(9)
        's = s & "<img src=" & Chr(34) & "images\christi.gif" & Chr(34) & ">"
        s = s & "<img src=" & Chr(34) & "images\realtrail.jpg" & Chr(34) & ">"
        s = s & "<BR>Currently accepting orders for: " & odates & "."
    Else
        's = "<img src=" & Chr(34) & "images\redlite.jpg" & Chr(34) & ">" & Chr(9)
        s = s & "<img src=" & Chr(34) & "images\orderoff.jpg" & Chr(34) & ">"
        s = s & "<BR>Not accepting branch orders at this time..."
    End If
    pgrid.AddItem s
    
    's = "<img src=" & Chr(34) & "images\toytruck.gif" & Chr(34) & ">" & Chr(9)
    's = Chr(9)
    s = "Transport Schedule Requests" & Chr(9)
    s = s & "Currently accepting requests for Week of " & scdates & "."
    pgrid.AddItem s
    
    s = "<a href=" & Chr(34) & "schedule\trnspts.htm" & Chr(34) & ">"
    s = s & "<img src=" & Chr(34) & "images\realtruck.jpg" & Chr(34) & " Border=0><BR>Transport Schedules</a>"
    s = s & Chr(9) & "Last Updated: " & FileDateTime(Form1.webdir & "\schedule\trnspts.htm")
    pgrid.AddItem s
    
    's = Chr(9)
    s = "<a href=" & Chr(34) & "directs\wdirects.htm" & Chr(34) & ">Driving Directions</a>"
    s = s & Chr(9) & "Last Updated: " & FileDateTime(Form1.webdir & "\directs\wdirects.htm")
    pgrid.AddItem s
    
    's = Chr(9)
    s = "<a href=" & Chr(34) & "schedule\trltrks.htm" & Chr(34) & ">Trailer Tracking</a>"
    s = s & Chr(9) & "Last Updated: " & FileDateTime(Form1.webdir & "\schedule\trltrks.htm")
    pgrid.AddItem s
    
    's = Chr(9)
    s = "<a href=" & Chr(34) & "schedule\intrax.htm" & Chr(34) & ">"
    s = s & "<img src=" & Chr(34) & "images\tractors.jpg" & Chr(34) & " Border=0><BR>Tractors in the Yard</a>"
    s = s & Chr(9) & "Last Updated: " & FileDateTime(Form1.webdir & "\schedule\intrax.htm")
    pgrid.AddItem s
    
    's = "<a href=" & Chr(34) & "stock\brancaps.htm" & Chr(34) & ">Branch Storage Capacities</a>"
    's = s & Chr(9) & "Last Updated: " & FileDateTime(Form1.webdir & "\stock\brancaps.htm")
    'pgrid.AddItem s
    
    rt = "Blue Bell Warehousing & Distribution"
    rt = "<img src=" & Chr(34) & "images/wdlogo.gif" & Chr(34) & ">"
    rt = rt & "<body background=" & Chr(34) & "images\wdbkgd.gif" & Chr(34) & ">"
    rh = "<img src=" & Chr(34) & "images/bbcolor.jpg" & Chr(34) & ">"
    'rh = rh & "<br><img src=" & Chr(34) & "images/wdlogo.gif" & Chr(34) & ">"
    rf = "Updated: " & Format(Now, "m-d-yyyy h:mm am/pm")
    pgrid.FormatString = "^|^|^"
    pgrid.ColWidth(0) = 2000
    pgrid.ColWidth(1) = pgrid.Width - 2000
    Call htmlcolorgrid(Me, Form1.webdir & "\bbwd.htm", pgrid, rt, rh, rf, "Linen", "White", "lemonchiffon")
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "rebuild_homegrid", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " rebuild_homegrid - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command2_Click()
    'Enable Branch Ordering
    Dim afile As String
    On Error GoTo vberror
    Sdb.Execute "update wdstatus set ordflag = 'Y'"
    Command2.Enabled = False
    Command3.Enabled = True
    Call rebuild_homegrid
    If Len(Dir(Form1.webdir & "\orderoff.txt")) > 0 Then
        Kill Form1.webdir & "\orderoff.txt"
    End If
    If IsDate(Text1) Then
        'afile = Form1.webdir & "\orders\aord." & DateDiff("d", "01-01-1999", Text1)
        afile = Form1.webdir & "\orders\aord" & DateDiff("d", "01-01-1999", Text1)
        Open afile For Output As #1
        Print #1, Text1
        Close #1
    End If
    If IsDate(Text2) Then
        'afile = Form1.webdir & "\orders\aord." & DateDiff("d", "01-01-1999", Text2)
        afile = Form1.webdir & "\orders\aord" & DateDiff("d", "01-01-1999", Text2)
        Open afile For Output As #1
        Print #1, Text2
        Close #1
    End If
    Open Form1.webdir & "\userlog" For Append As #1
    Print #1, Format(Now, "mm-dd-yyyy hh:mm AM/PM") & " - Enabled Orders: " & Text1 & "; " & Text2 & " ****"
    Close #1
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, Command2.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command2_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command3_Click()
    'Disable Branch Ordering
    On Error Resume Next
    Kill Form1.webdir & "\orders\aord*.*"
    On Error GoTo 0
    On Error GoTo vberror
    Sdb.Execute "update wdstatus set ordflag = 'N'"
    Command3.Enabled = False
    Command2.Enabled = True
    Call rebuild_homegrid
    Open Form1.webdir & "\orderoff.htm" For Output As #1
    Print #1, "<html>"
    Print #1, "<head><title>Branch Order</title></head>"
    Print #1, "<body background=" & Chr(34) & "images/wdbkgd.gif" & Chr(34) & ">"
    Print #1, "<center><img src=" & Chr(34) & "images/wdlogo.gif" & Chr(34) & "></center>"
    Print #1, "<br>"
    Print #1, "<center><b>Sorry!</b></center>"
    Print #1, "<br>"
    Print #1, "<center><b>We are not accepting orders at this time.</b></center>"
    Print #1, "<br>"
    Print #1, "<center><b>Please try again later.</b></center>"
    Print #1, "<br>"
    Print #1, "<center>Posted: " & Format(Now, "dddd m-d-yyyy h:mm am/pm") & " Central Time...</center>"
    Print #1, "</body></html>"
    Close #1
    Open Form1.webdir & "\orderoff.txt" For Output As #1
    Print #1, Format(Now, "mm-dd-yyyy hh:mm AM/PM") & " - Disabled Orders: " & Text1 & "; " & Text2 & " !!!!"
    Close #1
    Open Form1.webdir & "\userlog" For Append As #1
    Print #1, Format(Now, "mm-dd-yyyy hh:mm AM/PM") & " - Disabled Orders: " & Text1 & "; " & Text2 & " !!!!"
    Close #1
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, Command3.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command3_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command4_Click()
    ' Post Out of Stock Grid To Homepage
    Dim i As Integer
    Screen.MousePointer = 11
    outstk_sched
    Open Form1.webdir & "\stock.htm" For Output As #1
    Print #1, "<HTML>"
    Print #1, "<HEAD><TITLE>Out of Stock</TITLE></HEAD>"
    Print #1, "<BODY BGCOLOR=" & Chr(34) & "#C0FFF" & Chr(34) & ">"
    Print #1, "<CENTER><img src=" & Chr(34) & "images/wdlogo.gif" & Chr(34) & "></CENTER>"
    Print #1, "<BR>"
    Print #1, "<FONT FACE=" & Chr(34) & "Arial Black" & Chr(34) & "SIZE=4>Out of Stock Listings [Production Schedule Dates]</FONT>"
    Print #1, "<BR>"
    Print #1, "<FONT FACE=" & Chr(34) & "Arial Black" & Chr(34) & "SIZE=2>Last Updated: " & Format(Now, "m-d-yyyy h:mm am/pm") & "</FONT>"
    For i = 0 To Grid1.Rows - 1
        If i > 0 And Val(Grid1.TextMatrix(i, 0)) = 0 Then
            Print #1, "</FONT>"
            Print #1, "<HR><FONT FACE=" & Chr(34) & "Arial Black" & Chr(34) & "SIZE=2>" & Grid1.TextMatrix(i, 1) & " Plant</FONT><HR>"
            'Print #1, "<FONT FACE=" & Chr(34) & "MS Sans Serif" & Chr(34) & "SIZE=1>"
            Print #1, "<FONT FACE=" & Chr(34) & "Courier New" & Chr(34) & "SIZE=2>"
        Else
            If i > 0 Then
                Print #1, Grid1.TextMatrix(i, 0) & " ";
                Print #1, Grid1.TextMatrix(i, 1) & String(46 - Len(Grid1.TextMatrix(i, 1)), ".");
                Print #1, "  [" & Grid1.TextMatrix(i, 3) & "]<BR>"
            End If
        End If
    Next i
    Print #1, "</FONT>"
    Print #1, "</BODY></HTML>"
    Close #1
    Call rebuild_homegrid
    Screen.MousePointer = 0
End Sub

Private Sub Command5_Click()
    ' Generate Out of Stock List
    Dim ds As adodb.Recordset, ss As adodb.Recordset
    Dim sqlx As String, mplant As Integer
    Screen.MousePointer = 11
    On Error GoTo vberror
    ' Tag Plant Skus for out and low flags.
    Sdb.Execute "update plantskus set lowflag = 'Y', outflag = 'Y'"
    sqlx = "select sku,plant,sum(avail) from whstotals,warehouses"
    sqlx = sqlx & " where whstotals.whs_num = warehouses.whs_num"
    sqlx = sqlx & " and warehouses.whs not in ('OP','DROP','SDRP')"
    sqlx = sqlx & " group by sku,plant"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = "select * from plantskus where plant = " & ds!plant & " and sku = '" & ds!sku & "'"
            Set ss = Sdb.Execute(sqlx)
            If ss.BOF = False Then
                ss.MoveFirst
                If ss!lowstk < ds(2) And ss!lowstk <> 0 Then
                    sqlx = "Update plantskus set lowflag = 'N' where id = " & ss!id
                    Sdb.Execute sqlx
                End If
                If ss!outstk < ds(2) And ss!outstk <> 0 Then
                    sqlx = "Update plantskus set outflag = 'N' where id = " & ss!id
                    Sdb.Execute sqlx
                End If
            End If
            ss.Close
            ds.MoveNext
        Loop
    End If
    ds.Close
    mplant = 0
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4: Grid1.Visible = False
    sqlx = "select plantskus.sku,plantskus.plant,fgunit,fgdesc,plantname"
    sqlx = sqlx & " from plantskus,skumast,plants"
    sqlx = sqlx & " where outflag = 'Y' and outstk > 0"
    sqlx = sqlx & " and plantskus.sku = skumast.sku"
    sqlx = sqlx & " and plantskus.plant = plants.plant"
    sqlx = sqlx & " order by plantskus.plant,fgunit,fgdesc"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds(1) <> mplant Then
                Grid1.AddItem Chr(9) & ds!plantname
                mplant = ds(1)
            End If
            Grid1.AddItem ds(0) & Chr(9) & ds!fgunit & " " & ds!fgdesc & Chr(9) & ds(1) & Chr(9) & "Not scheduled."
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^SKU|<Description|>Plant|>Next Date"
    Grid1.ColWidth(0) = 700
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1
    Grid1.ColWidth(3) = 1
    Grid1.Visible = True
    Screen.MousePointer = 0
    Command4.Enabled = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_wos_ado", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_wos_ado - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command7_Click()
    Dim rt As String, rh As String, rf As String, hf As String
    Screen.MousePointer = 11
    outstk_sched
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 3
    For i = 1 To Grid1.Rows - 1
        s = Grid1.TextMatrix(i, 0) & Chr(9)
        s = s & Grid1.TextMatrix(i, 1) & Chr(9)
        s = s & Grid1.TextMatrix(i, 3)
        pgrid.AddItem s
    Next i
    For i = 0 To pgrid.Rows - 1
        If pgrid.TextMatrix(i, 1) = "Brenham" Then
            'pgrid.TextMatrix(i, 0) = "<img src=" & Chr(34) & "images/texas.jpg" & Chr(34) & ">"
            pgrid.TextMatrix(i, 1) = "<Font size=4><Center>Brenham</Center>"
            'pgrid.TextMatrix(i, 2) = "<img src=" & Chr(34) & "images/texas.jpg" & Chr(34) & ">"
        End If
        If pgrid.TextMatrix(i, 1) = "Broken Arrow" Then
            'pgrid.TextMatrix(i, 0) = "<img src=" & Chr(34) & "images/oklahoma.jpg" & Chr(34) & ">"
            pgrid.TextMatrix(i, 1) = "<Font size=4><Center>Broken Arrow</Center>"
            'pgrid.TextMatrix(i, 2) = "<img src=" & Chr(34) & "images/oklahoma.jpg" & Chr(34) & ">"
        End If
        If pgrid.TextMatrix(i, 1) = "Sylacauga" Then
            'pgrid.TextMatrix(i, 0) = "<img src=" & Chr(34) & "images/alabama.jpg" & Chr(34) & ">"
            pgrid.TextMatrix(i, 1) = "<Font size=4><Center>Sylacauga</Center>"
            'pgrid.TextMatrix(i, 2) = "<img src=" & Chr(34) & "images/alabama.jpg" & Chr(34) & ">"
        End If
    Next i
    pgrid.FormatString = "^SKU|^|^Next Production"
    pgrid.ColWidth(0) = 700
    pgrid.ColWidth(1) = 3000
    pgrid.ColWidth(2) = 1800
    rt = "<CENTER><img src=" & Chr(34) & "images/wdlogo.gif" & Chr(34) & ">"
    rt = rt & "<BR>Out of Stock Listings"
    'rh = "Out of Stock Items"
    rf = rf & "Last updated: " & Format(Now, "m-d-yyyy h:mm am/pm") & " CST"
    hf = Form1.webdir & "\stkgrid.htm"
    Call htmlcolorgrid(Me, hf, pgrid, rt, rh, rf, "linen", "peachpuff", "white")
    Screen.MousePointer = 0
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Bravail.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
            If Form1.FrmGrid.Text = "bravail" Then
                Form1.FrmGrid.Col = 1: Form1.FrmGrid.Text = Bravail.Top
                Form1.FrmGrid.Col = 2: Form1.FrmGrid.Text = Bravail.Left
                Form1.FrmGrid.Col = 3: Form1.FrmGrid.Text = Bravail.Height
                Form1.FrmGrid.Col = 4: Form1.FrmGrid.Text = Bravail.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer, ds As adodb.Recordset
    Dim mplant As Integer
    For i = 1 To Form1.FrmGrid.Rows - 1
        Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
        If Form1.FrmGrid.Text = "bravail" Then
            Form1.FrmGrid.Col = 1: Bravail.Top = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 2: Bravail.Left = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 3: Bravail.Height = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 4: Bravail.Width = Val(Form1.FrmGrid.Text)
            Exit For
        End If
    Next i
    On Error GoTo vberror
    Set ds = Sdb.Execute("select * from wdstatus")
    If ds.BOF = False Then
        ds.MoveFirst
        If ds!ordflag = "Y" Then
            Command3.Enabled = True: Command2.Enabled = False
        Else
            Command3.Enabled = False: Command2.Enabled = True
        End If
        Text1 = Left(ds!orddates, 10)
        Text2 = Right(ds!orddates, 10)
    End If
    ds.Close
    
    ' Generate Out of Stock List
    Screen.MousePointer = 11
    i = 0
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4: Grid1.Visible = False
    sqlx = "select plantskus.sku,plantskus.plant,fgunit,fgdesc,plantname"
    sqlx = sqlx & " from plantskus,skumast,plants"
    sqlx = sqlx & " where outflag <> 'N' and outstk > 0"
    sqlx = sqlx & " and plantskus.sku = skumast.sku"
    sqlx = sqlx & " and plantskus.plant = plants.plant"
    sqlx = sqlx & " order by plantskus.plant,fgunit,fgdesc"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds(1) <> i Then
                Grid1.AddItem Chr(9) & ds!plantname
                i = ds(1)
            End If
            Grid1.AddItem ds(0) & Chr(9) & ds!fgunit & " " & ds!fgdesc & Chr(9) & ds(1) & Chr(9) & "Not scheduled."
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^SKU|<Description|>Plant|>Next Date"
    Grid1.ColWidth(0) = 700
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1
    Grid1.ColWidth(3) = 1
    Grid1.Visible = True
    Screen.MousePointer = 0
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
    pgrid.Width = Me.Width - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Text1_Change()
    Dim daylit(1 To 7) As String
    If IsDate(Text1) Then
        daylit(1) = "Sunday"
        daylit(2) = "Monday"
        daylit(3) = "Tuesday"
        daylit(4) = "Wednesday"
        daylit(5) = "Thursday"
        daylit(6) = "Friday"
        daylit(7) = "Saturday"
        Label4 = daylit(WeekDay(Text1))
    Else
        Label4 = "..."
    End If
End Sub

Private Sub Text2_Change()
    Dim daylit(1 To 7) As String
    If IsDate(Text2) Then
        daylit(1) = "Sunday"
        daylit(2) = "Monday"
        daylit(3) = "Tuesday"
        daylit(4) = "Wednesday"
        daylit(5) = "Thursday"
        daylit(6) = "Friday"
        daylit(7) = "Saturday"
        Label5 = daylit(WeekDay(Text2))
    Else
        Label5 = "..."
    End If
End Sub
