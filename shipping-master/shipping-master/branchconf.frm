VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form branchconf 
   Caption         =   "Branch Configuration and Messages"
   ClientHeight    =   7905
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12270
   ForeColor       =   &H00000040&
   LinkTopic       =   "Form3"
   ScaleHeight     =   7905
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   7320
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "branchconf.frx":0000
      Top             =   6720
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   4320
      TabIndex        =   4
      Top             =   6480
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   4200
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   4575
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1508
      _Version        =   327680
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   4815
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   9375
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6165
      _Version        =   327680
      BackColorFixed  =   12648384
      WordWrap        =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu insrec 
         Caption         =   "Insert New Record"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete Record"
      End
      Begin VB.Menu edrec 
         Caption         =   "Edit Field"
      End
      Begin VB.Menu clrmess 
         Caption         =   "Clear Message"
      End
      Begin VB.Menu copymess 
         Caption         =   "Copy Message"
      End
      Begin VB.Menu pastemess 
         Caption         =   "Paste Message"
      End
      Begin VB.Menu undomess 
         Caption         =   "Undo Message Change"
         Enabled         =   0   'False
      End
      Begin VB.Menu prevhtml 
         Caption         =   "Preview HTML"
      End
   End
End
Attribute VB_Name = "branchconf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim newmess As Boolean
Dim edcell As String
Dim msrow As Integer
Dim mlength As Integer
Private Sub rebuild_homegrid()
    Dim ds As adodb.Recordset, s As String
    Dim odates As String, scdates As String
    Dim rt As String, rh As String, rf As String
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 2: pgrid.FixedCols = 1
    pgrid.FixedCols = 0
    On Error GoTo vberror
    Set ds = Sdb.Execute("select * from wdstatus")      'jv061316
    If ds.BOF = False Then
        ds.MoveFirst
        odates = ds!orddates
        scdates = ds!schdates
    End If
    ds.Close
    s = "select * from branches where branch > 90 and brnmess > '   ' order by branch desc"
    Set ds = Sdb.Execute(s)                             'jv061316
    If ds.BOF = False Then
        ds.MoveFirst
        s = "<img src=" & Chr(34) & "images\new.jpg" & Chr(34) & "><BR>" & Chr(9)
        Do Until ds.EOF
            s = s & "<b>" & ds!branchname & ":</b><br>" & ds!brnmess & "<hr>"
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
    If Len(Dir(Form1.webdir & "\orderoff.txt")) > 0 Then
        s = s & "<img src=" & Chr(34) & "images\orderoff.jpg" & Chr(34) & ">"
        s = s & "<BR>Not accepting branch orders at this time..."
    Else
        s = s & "<img src=" & Chr(34) & "images\realtrail.jpg" & Chr(34) & ">"
        s = s & "<BR>Currently accepting orders for: " & odates & "."
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
    's = "<a href=" & Chr(34) & "schedule\intrax.htm" & Chr(34) & ">"
    's = s & "<img src=" & Chr(34) & "images\tractors.jpg" & Chr(34) & " Border=0><BR>Tractors in the Yard</a>"
    's = s & Chr(9) & "Last Updated: " & FileDateTime(Form1.webdir & "\schedule\intrax.htm")
    'Pgrid.AddItem s
    
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

Private Sub form_memo(memx As String)
    Dim i As Long, k As Long, filx As String
    List1.Clear: List2.Clear
    i = 1: k = 1
    If Len(memx) < 72 Then
        List2.AddItem memx
        Exit Sub
    End If
    Do Until i = 0
        i = InStr(i, memx, " ", vbBinaryCompare)
        If i = 0 Then Exit Do
        List1.AddItem Trim(mid(memx, k, i - k))
        k = i
        i = i + 1
    Loop
    List1.AddItem Trim(mid(memx, k, Len(memx) - k + 1))
    filx = ""
    For i = 0 To List1.ListCount - 1
        If Len(filx & List1.List(i)) > 72 Then
            List2.AddItem filx
            filx = List1.List(i) & " "
        Else
            filx = filx & List1.List(i) & " "
        End If
    Next i
    List2.AddItem filx
End Sub

Private Sub update_item()
    Dim sqlx As String
    On Error GoTo vberror
    sqlx = "Update branches set "
    If edcell = "modem" Then
        Grid1.Text = Left(Grid1.Text, 22)
        sqlx = sqlx & "modem = '" & Grid1.Text & "'"
    End If
    If edcell = "fax" Then
        Grid1.Text = Left(Grid1.Text, 22)
        sqlx = sqlx & "fax = '" & Grid1.Text & "'"
    End If
    If edcell = "ip" Then
        Grid1.Text = Left(Grid1.Text, 30)
        sqlx = sqlx & "ip = '" & Grid1.Text & "'"
    End If
    If edcell = "addr1" Then
        Grid1.Text = Left(Grid1.Text, 25)
        sqlx = sqlx & "addr1 = '" & fixquotes(Grid1.Text) & "'"
    End If
    If edcell = "addr2" Then
        Grid1.Text = Left(Grid1.Text, 25)
        sqlx = sqlx & "addr2 = '" & fixquotes(Grid1.Text) & "'"
    End If
    If edcell = "call_group" Then
        Grid1.Text = Left(Grid1.Text, 1)
        sqlx = sqlx & "call_group = '" & Grid1.Text & "'"
    End If
    If edcell = "gemmsid" Then
        Grid1.Text = Left(Grid1.Text, 6)
        sqlx = sqlx & "gemmsid = '" & Grid1.Text & "'"
    End If
    If edcell = "brphone" Then
        Grid1.Text = Left(Grid1.Text, 25)
        sqlx = sqlx & "brphone = '" & Grid1.Text & "'"
    End If
    If edcell = "brfax" Then
        Grid1.Text = Left(Grid1.Text, 25)
        sqlx = sqlx & "brfax = '" & Grid1.Text & "'"
    End If
    sqlx = sqlx & " where branch = " & Grid1.TextMatrix(Grid1.Row, 0)
    Sdb.Execute sqlx                    'jv061316
    edcell = ""
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

Private Sub refresh_grid1()
    Dim ds As adodb.Recordset, s As String, i As Integer
    On Error GoTo vberror
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    Grid1.FixedCols = 2
    s = "select * from branches order by branch"
    s = "select branch,branchname,brnmess,call_group,modem,fax,ip,gemmsid,addr1,addr2,brphone,brfax from branches order by branch"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!branch & Chr(9)
            s = s & ds!branchname & Chr(9)
            s = s & ds!brnmess & Chr(9)
            s = s & ds!call_group & Chr(9)
            s = s & ds!modem & Chr(9)       'Total Pallet Capacity
            s = s & ds!fax & Chr(9)         'Usable Pallet Space
            s = s & ds!ip & Chr(9)
            s = s & ds!gemmsid & Chr(9)
            s = s & ds!addr1 & Chr(9)
            s = s & ds!addr2 & Chr(9)
            s = s & ds!brphone & Chr(9)
            s = s & ds!brfax
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "^ID|<Name|<Message|^Group|^Total Pallet Cap|^Usable Pallet Space|<IP|^Oracle Storeroom|<Street|<City-Zipcode|<Phone|<Fax"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 2000
    Grid1.ColWidth(2) = 6000
    Grid1.ColWidth(3) = 600
    Grid1.ColWidth(4) = 1400
    Grid1.ColWidth(5) = 1600
    Grid1.ColWidth(6) = 1200
    Grid1.ColWidth(7) = 1200
    Grid1.ColWidth(8) = 2200
    Grid1.ColWidth(9) = 2200
    Grid1.ColWidth(10) = 1500
    Grid1.ColWidth(11) = 1500
    For i = 1 To Grid1.Rows - 1
        Grid1.RowHeight(i) = Grid1.RowHeight(0) * 2
    Next i
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid1", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid1 - Error Number: " & eno
        End
    End If
End Sub

Private Sub clrmess_Click()
    msrow = Grid1.Row
    Text1.Text = "..."
    mlength = Len(Text1)
    newmess = True
    Call Text1_LostFocus
End Sub

Private Sub copymess_Click()
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) > 0 Then
        Text2 = Grid1.TextMatrix(Grid1.Row, 2)
    End If
End Sub

Private Sub delrec_Click()
    Dim sqlx As String
    sqlx = "Ok to delete branch code: " & Grid1.TextMatrix(Grid1.Row, 0)
    sqlx = sqlx & " Branch: " & Grid1.TextMatrix(Grid1.Row, 1)
    If MsgBox(sqlx, vbYesNo + vbQuestion, "Are you sure....") = vbNo Then Exit Sub
    On Error GoTo vberror
    sqlx = "Delete from branches where branch = " & Grid1.TextMatrix(Grid1.Row, 0)
    Sdb.Execute sqlx                        'jv061316
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_grid1
    End If
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

Private Sub edrec_Click()
    Dim s As String, f As String
    If Grid1.Row = 0 Or Grid1.Col < 2 Then Exit Sub
    If Grid1.Col = 2 Then
        msrow = Grid1.Row
        Text1 = Grid1.Text
        Text3 = Grid1.Text
        Text1.Visible = True
        mlength = Len(Text1)
        Text1.SetFocus
        Exit Sub
    End If
    s = Grid1.Text
    f = Grid1.TextMatrix(0, Grid1.Col)
    s = InputBox(f, f, s)
    If Len(s) = 0 Then Exit Sub
    Grid1.Text = s
    On Error GoTo vberror
    s = "Update branches set "
    s = s & "call_group = '" & Grid1.TextMatrix(Grid1.Row, 3) & "'"
    s = s & ",modem = '" & Grid1.TextMatrix(Grid1.Row, 4) & "'"
    s = s & ",fax = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
    s = s & ",ip = '" & Grid1.TextMatrix(Grid1.Row, 6) & "'"
    s = s & ",gemmsid = '" & Grid1.TextMatrix(Grid1.Row, 7) & "'"
    s = s & ",addr1 = '" & fixquotes(Grid1.TextMatrix(Grid1.Row, 8)) & "'"
    s = s & ",addr2 = '" & fixquotes(Grid1.TextMatrix(Grid1.Row, 9)) & "'"
    s = s & ",brphone = '" & Grid1.TextMatrix(Grid1.Row, 10) & "'"
    s = s & ",brfax = '" & Grid1.TextMatrix(Grid1.Row, 11) & "'"
    s = s & " where branch = " & Grid1.TextMatrix(Grid1.Row, 0)
    Sdb.Execute s                           'jv061316
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

Private Sub Form_Load()
    Text2 = ".."
    newmess = False
    refresh_grid1
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    pgrid.Width = Me.Width - 120
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 820
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Len(edcell) > 0 Then Call update_item
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
    If Grid1.Row = 0 Or Grid1.Col < 3 Then Exit Sub
    If Len(edcell) = 0 Then Grid1.Text = ""
    If Grid1.Col = 3 Then edcell = "call_group"
    If Grid1.Col = 4 Then edcell = "modem"
    If Grid1.Col = 5 Then edcell = "fax"
    If Grid1.Col = 6 Then edcell = "ip"
    If Grid1.Col = 7 Then edcell = "gemmsid"
    If Grid1.Col = 8 Then edcell = "addr1"
    If Grid1.Col = 9 Then edcell = "addr2"
    If Grid1.Col = 10 Then edcell = "brphone"
    If Grid1.Col = 11 Then edcell = "brfax"
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

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid1_RowColChange()
    If msrow <> Grid1.Row Then undomess.Enabled = False
End Sub

Private Sub insrec_Click()
    Dim nb As String, i As Integer, s As String
    Dim sqlx As String
    If Len(edcell) > 0 Then Call update_item
    nb = InputBox("New Branch Code: ", "Insert new branch...", "0")
    If Len(nb) = 0 Or Val(nb) = 0 Then Exit Sub
    For i = 0 To Grid1.Rows - 1
        If Val(nb) = Val(Grid1.TextMatrix(i, 0)) Then
            sqlx = "Branch Code " & nb
            sqlx = sqlx & " already in use for "
            sqlx = sqlx & Grid1.TextMatrix(i, 1) & "."
            MsgBox sqlx, vbOKOnly + vbExclamation, "Sorry, cannot add..."
            Exit Sub
        End If
    Next i
    s = InputBox("Branch Name: ", "Branch Name...", "New Branch")
    If Len(s) = 0 Then Exit Sub
    On Error GoTo vberror
    sqlx = "Insert into branches (branch, branchname) Values (" & Val(nb) & ", '" & s & "')"
    Sdb.Execute sqlx
    Call refresh_grid1
    For i = 0 To Grid1.Rows - 1
        If Val(nb) = Val(Grid1.TextMatrix(i, 0)) Then
            Grid1.Row = i: Grid1.TopRow = i
            Exit For
        End If
    Next i
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

Private Sub pastemess_Click()
    msrow = Grid1.Row
    Text1 = Text2
    mlength = Len(Text1)
    newmess = True
    Call Text1_LostFocus
End Sub

Private Sub prevhtml_Click()
    Dim cfile As String, i
    cfile = localAppDataPath & "\mssg.htm" ' U:\mssg.htm
    Open cfile For Output As #1
    Print #1, "<html><body><center>"
    Print #1, Grid1.TextMatrix(Grid1.Row, 2)
    Print #1, "</center></body></html>"
    Close #1
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe u:\mssg.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe u:\mssg.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Text1.Text = RTrim(Text1.Text)
        Text1.Visible = False
        Grid1.SetFocus
    Else
        undomess.Enabled = True
        newmess = True
    End If
End Sub

Private Sub Text1_LostFocus()
    Dim s As String
    On Error GoTo vberror
    If mlength <> Len(Text1) Then newmess = True
    If newmess = True Then
        Screen.MousePointer = 11
        Open Form1.webdir & "\stock\message." & Format(Val(Grid1.TextMatrix(msrow, 0)), "00") For Output As #1
        Print #1, Grid1.TextMatrix(msrow, 1) & " Messages sent: " & Format(Now, "dddd m-d-yyyy h:mm am/pm")
        If Len(Text1.Text) > 72 Then
            Call form_memo(Text1.Text)
            For i = 0 To List2.ListCount - 1
                Print #1, List2.List(i)
            Next i
        Else
            Print #1, Text1.Text
        End If
        Print #1, "<END>"
        Close #1
                
        If Len(Text1) = 0 Then Text1 = " "
        s = fixquotes(Text1)
        's = fixamps(s)
        s = "Update branches set brnmess = '" & s & "' Where branch = " & Grid1.TextMatrix(msrow, 0)
        Sdb.Execute s
        Grid1.TextMatrix(msrow, 2) = Text1
        
        If Val(Grid1.TextMatrix(msrow, 0)) > 90 Then Call rebuild_homegrid
        Screen.MousePointer = 0
    End If
    
    newmess = False
    Text1.Visible = False
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

Private Sub undomess_Click()
    msrow = Grid1.Row
    Text1 = Text3
    mlength = Len(Text1)
    newmess = True
    Call Text1_LostFocus
End Sub
