VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form tstations 
   Caption         =   "Transfer Stations"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Post to Browser"
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
      Left            =   7560
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4800
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6255
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   11033
      _Version        =   327680
      BackColor       =   16777215
      BackColorFixed  =   16777152
      FocusRect       =   0
   End
   Begin VB.ListBox List3 
      Height          =   450
      Left            =   6000
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   6000
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   7095
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
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label dtrig 
      Caption         =   "dtrig"
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "gcolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "bcolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "ycolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "wcolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   8
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Transfer Station:"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "tstations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub post_countsheet()
    Dim psku As String, pdesc As String, ppal As String, pwrap As String, peach As String, pqty As String
    Dim pconv As Integer, wconv As Integer, i As Integer, cfile As String, j As Long
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\counts\tstation.tst"
    Open cfile For Output As #1
    For i = 1 To Grid1.Rows - 1
        psku = Grid1.TextMatrix(i, 0)
        pdesc = StrConv(Grid1.TextMatrix(i, 1), vbProperCase)
        pconv = skurec(Val(psku)).pallet
        wconv = skurec(Val(psku)).wrapunits
        pqty = Val(Grid1.TextMatrix(i, 6))
        ppal = "0"
        pwrap = "0"
        peach = "0"
        j = Val(pqty)
        If j > 0 Then
            If pconv > 1 Then
                If j > pconv Then
                    ppal = Int(j / pconv)
                    j = j - (Val(ppal) * pconv)
                End If
                If j > wconv Then
                    pwrap = Int(j / wconv)
                    j = j - (Val(pwrap) * wconv)
                End If
                peach = j
            Else
                peach = j
            End If
        Else
            pqty = "0"
        End If
        ppal = Format(Val(ppal), "#")
        pwrap = Format(Val(pwrap), "#")
        peach = Format(Val(peach), "#")
        Write #1, psku; psku; pdesc; ppal; pwrap; peach; pqty; "6-17-2018"
    Next i
    Close #1
End Sub

Private Sub refresh_skus(pbranch As String)
    Dim ds As ADODB.Recordset, s As String, pplant As String
    Grid1.Redraw = False
    Grid1.FontBold = True
    Grid1.FontName = "Arial"
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 10: Grid1.FixedCols = 2
    'Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7: Grid1.FixedCols = 2
    pplant = branchrec(Val(pbranch)).supplier
    s = "select sku from bimp where branchwhs = '" & pbranch & "' and plantwhs = '" & pplant & "' order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^SKU|<Product|^Count Date|^Beg Inventory|^Trans In|^Trans Out|^Net|^30 Day Sales|^Days Supply"
    'Grid1.FormatString = "^SKU|<Product|^Count Date|^Beg Inventory|^Trans In|^Trans Out|^Net"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1500 '2000
    Grid1.ColWidth(3) = 1500 '2000
    Grid1.ColWidth(4) = 1500 '2000
    Grid1.ColWidth(5) = 1500 '2000
    Grid1.ColWidth(6) = 1500 '2000
    Grid1.ColWidth(7) = 1500 '2000
    Grid1.ColWidth(8) = 1500 '2000
    Grid1.ColWidth(9) = 0
    Grid1.Redraw = True
End Sub

Private Sub read_countsheet(cfile As String)
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim i As Integer, pflag As Boolean
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7
        pflag = False
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 0) = f1 Then
                Grid1.TextMatrix(i, 2) = f7
                Grid1.TextMatrix(i, 3) = f6
                Grid1.TextMatrix(i, 9) = f0                 'jv072718
                pflag = True
                Exit For
            End If
        Next i
        If pflag = False Then
            s = Trim(f1) & Chr(9)
            s = s & UCase(f2) & Chr(9)
            s = s & f7 & Chr(9)
            s = s & f6
            Grid1.AddItem s
            'MsgBox s
        End If
    Loop
    Close #1
    For i = 1 To Grid1.Rows - 1                                                     'jv072718
        If Val(Grid1.TextMatrix(i, 9)) = 0 Then Grid1.TextMatrix(i, 9) = "99" & Grid1.TextMatrix(i, 0) 'jv072718
    Next i                                                                          'jv072718
    Grid1.Row = 1: Grid1.RowSel = 1
    'Grid1.Col = 0: Grid1.ColSel = 0
    Grid1.Col = 9: Grid1.ColSel = 9                                                 'jv072718
    'Grid1.Sort = 5
    Grid1.Sort = 3                                                                  'jv072718
End Sub

Sub read_startdates()
    Dim s As String, ds As ADODB.Recordset, i As Integer, pflag As Boolean
    s = "select product_no, min(tran_date) from bolinf.inv_adj_input_detail"
    s = s & " where tran_type = '1'"
    s = s & " and branch_no = '062'"
    s = s & " and route_no = '20'"
    s = s & " group by product_no"
    s = s & " order by product_no"
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            pflag = False
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = Trim(ds(0)) Then
                    Grid1.TextMatrix(i, 2) = Format(DateAdd("d", -1, ds(1)), "M-d-yyyy")
                    Grid1.TextMatrix(i, 3) = " "
                    pflag = True
                End If
            Next i
            If pflag = False Then
                s = Trim(ds(0)) & Chr(9)
                s = s & skurec(Val(ds(0))).unit & " " & skurec(Val(ds(0))).desc & Chr(9)
                s = s & Format(DateAdd("d", -1, ds(1)), "M-d-yyyy")
                s = s & " "
                Grid1.AddItem s
                'MsgBox s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.Row = 1: Grid1.RowSel = 1
    Grid1.Col = 0: Grid1.ColSel = 0
    Grid1.Sort = 5
End Sub

Sub refresh_trans(pbranch As String, proute As String, psku As String, pdate As String, prow As Integer)
    Dim ds As ADODB.Recordset, q As String, s As String
    Dim i As Integer, t As Long, j As Long
    
    q = "select tran_qty from bolinf.inv_adj_input_detail"
    q = q & " Where tran_type = '1'"
    q = q & " and tran_date >= TO_DATE('" & Format(pdate, "dd-mmm-yy") & "')"
    'q = q & " and tran_date < TO_DATE('17-Jun-18')"                'build countsheet
    q = q & " and branch_no = '" & pbranch & "'"
    q = q & " and product_no = '" & psku & "'"
    q = q & " and route_no = '" & proute & "'"
        
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!tran_qty > 0 Then
                Grid1.TextMatrix(prow, 4) = Val(Grid1.TextMatrix(prow, 4)) + ds!tran_qty
            Else
                Grid1.TextMatrix(prow, 5) = Val(Grid1.TextMatrix(prow, 5)) + ds!tran_qty
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Sub refresh_sales(pbranch As String, proute As String)
    Dim ds As ADODB.Recordset, q As String, s As String
    Dim i As Integer, t As Long, j As Long
    
    q = "select product_no,sum(tran_qty) from bolinf.inv_adj_input_detail"
    q = q & " Where tran_type = '1'"
    q = q & " and tran_date >= sysdate - 31"
    q = q & " and branch_no = '" & pbranch & "'"
    q = q & " and route_no = '" & proute & "'"
    q = q & " and tran_qty < 0"
    q = q & " group by product_no order by product_no"
        
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = ds(0) Then
                    Grid1.TextMatrix(i, 7) = ds(1) * -1
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    List2.ListIndex = Combo1.ListIndex
    List3.ListIndex = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
    Dim i As Integer, j As Long, pdaze As Integer
    Screen.MousePointer = 11
    Call refresh_skus(List2)
    Call read_countsheet(List1)
    'Call read_startdates                                   'build countsheet
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 2) > " " Then
            Call refresh_trans(List2, List3, Grid1.TextMatrix(i, 0), Grid1.TextMatrix(i, 2), i)
            DoEvents
        End If
        j = Val(Grid1.TextMatrix(i, 3))
        j = j + Val(Grid1.TextMatrix(i, 4))
        j = j + Val(Grid1.TextMatrix(i, 5))
        Grid1.TextMatrix(i, 6) = Format(j, "#")
    Next i
    'post_countsheet                                        'build countsheet
    Call refresh_sales(List2, List3)                        'jv070918
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 7)) > 0 And Val(Grid1.TextMatrix(i, 6)) > 0 Then
            Grid1.TextMatrix(i, 8) = CInt(Val(Grid1.TextMatrix(i, 6)) / Val(Grid1.TextMatrix(i, 7)) * 30)
        End If
    Next i
    Grid1.FillStyle = flexFillRepeat
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 8) > " " Then
            pdaze = Val(Grid1.TextMatrix(i, 8))
        Else
            pdaze = 30
        End If
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
        If pdaze < 14 Then
            Grid1.CellBackColor = wcolor.BackColor
        Else
            If pdaze < 30 Then
                Grid1.CellBackColor = ycolor.BackColor
            Else
                If pdaze > 35 Then
                    Grid1.CellBackColor = gcolor.BackColor
                Else
                    Grid1.CellBackColor = bcolor.BackColor
                End If
            End If
        End If
    Next i
    For i = Grid1.Rows - 1 To 1 Step -1
        If Grid1.Rows = 2 Then
            Grid1.Rows = 2
        Else
            If Val(Grid1.TextMatrix(i, 7)) = 0 And Val(Grid1.TextMatrix(i, 6)) <= 0 Then
                Grid1.RemoveItem i
            End If
        End If
    Next i
    Grid1.Row = 1
    Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
    Dim rt As String, rf As String, rh As String, webfile As String, i As Integer, k As Integer
    webfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\counts\tstation" & List2 & ".csv"
    Open webfile For Output As #1
    For i = 1 To Grid1.Rows - 1
        For k = 0 To Grid1.Cols - 2
            Write #1, Grid1.TextMatrix(i, k);
        Next k
        Write #1, Grid1.TextMatrix(i, Grid1.Cols - 1)
    Next i
    Close #1
    
    
    webfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\counts\tstation" & List2 & ".htm"
    'webfile = "s:\wd\html\counts\test.htm"
    rt = Combo1 & "Transfer Station"
    rh = "Warehouse:  " & Combo1
    rf = "Posted: " & Format(Now, "m-d-yyyy h:mm am/pm")
    'htdc(0) = "seagreen": gndc(0) = Me.Grid1.BackColorFixed
    'htdc(1) = "cyan": gndc(1) = Me.rcolor.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid2.BackColor
    Grid1.Redraw = False
    Grid1.ColWidth(7) = 0
    Grid1.ColWidth(8) = 0
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, webfile, Grid1, rt, rh, rf, "linen", "cyan", "white")
        Grid1.ColWidth(7) = 1500
        Grid1.ColWidth(8) = 1500
        Grid1.Redraw = True
        i = Shell("C:\program files\internet explorer\iexplore.exe " & webfile, vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, webfile, Grid1, rt, rh, rf, "linen", "cyan", "white")
        Grid1.ColWidth(7) = 1500
        Grid1.ColWidth(8) = 1500
        Grid1.Redraw = True
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & webfile, vbNormalFocus)
        Exit Sub
    End If
        
End Sub

Private Sub dtrig_Change()
    Dim rt As String, rf As String, rh As String, webfile As String, i As Integer, k As Integer, j As Integer
    If dtrig = " " Then Exit Sub
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) = dtrig Then
            Combo1.ListIndex = i
            MsgBox "auto " & Combo1
            Command1_Click
            DoEvents
            webfile = "\\BBC-03-FILESVR\SharedGroups\WD\html\counts\tstation" & List2 & ".csv"
            'webfile = "u:\tstation" & List2 & ".csv"
            Open webfile For Output As #1
            For j = 1 To Grid1.Rows - 1
                For k = 0 To Grid1.Cols - 2
                    Write #1, Grid1.TextMatrix(j, k);
                Next k
                Write #1, Grid1.TextMatrix(j, Grid1.Cols - 1)
            Next j
            Close #1
            Exit For
        End If
    Next i
    dtrig = " "
End Sub

Private Sub Form_Load()
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Unload Me
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top

    Combo1.Clear
    List1.Clear
    List2.Clear
    List3.Clear
    Combo1.AddItem "Greenville"
    'List1.AddItem "s:\wd\html\counts\tstation.062"
    List1.AddItem "\\BBC-03-FILESVR\SharedGroups\WD\html\counts\tstation.062"
    'List1.AddItem "s:\wd\html\counts\tstation.tst"
    List2.AddItem "062"
    List3.AddItem "20"
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 200
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 3.5)
End Sub
