VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bimpstkhist 
   Caption         =   "Process Stock History"
   ClientHeight    =   12090
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   18075
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12090
   ScaleWidth      =   18075
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      ForeColor       =   &H000000FF&
      Height          =   9030
      Left            =   14880
      TabIndex        =   31
      Top             =   2040
      Width           =   5295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Update Stock History"
      Height          =   495
      Left            =   9960
      TabIndex        =   29
      Top             =   11400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Post Stock History"
      Height          =   495
      Left            =   1800
      TabIndex        =   28
      Top             =   11400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   12000
      TabIndex        =   27
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   8160
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print"
      Height          =   375
      Left            =   8880
      TabIndex        =   25
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print All Product Listing"
      Height          =   375
      Left            =   11400
      TabIndex        =   24
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "View SKU History"
      Height          =   375
      Left            =   6360
      TabIndex        =   23
      Top             =   600
      Width           =   2295
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   0
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   2520
      TabIndex        =   21
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   855
      Left            =   0
      TabIndex        =   20
      Top             =   4200
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1508
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid4 
      Height          =   6735
      Left            =   7560
      TabIndex        =   19
      Top             =   4320
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   11880
      _Version        =   327680
      BackColor       =   12648447
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   2295
      Left            =   8520
      TabIndex        =   18
      Top             =   1800
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4048
      _Version        =   327680
      Rows            =   8
      Cols            =   3
      BackColorFixed  =   12648447
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   9495
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   16748
      _Version        =   327680
      Appearance      =   0
   End
   Begin VB.TextBox edate 
      Height          =   285
      Left            =   4080
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox sdate 
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   600
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6360
      TabIndex        =   13
      Text            =   "Combo2"
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label postdate 
      Caption         =   "postdate"
      Height          =   255
      Left            =   4680
      TabIndex        =   30
      Top             =   9000
      Width           =   1455
   End
   Begin VB.Label pdesc 
      Caption         =   "pdesc"
      Height          =   255
      Left            =   7560
      TabIndex        =   14
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label psales 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label odaze 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label idaze 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label tdaze 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Loads"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "End Date:"
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Start Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Days Out of Stock"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Days In Stock"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Days"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "SKU:"
      Height          =   255
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Branch Whs:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu g3menu 
      Caption         =   "grid3"
      Visible         =   0   'False
      Begin VB.Menu pabskus 
         Caption         =   "Process All Branch SKUs"
      End
      Begin VB.Menu pallbskus 
         Caption         =   "Process All Branches All SKUs"
      End
      Begin VB.Menu unpamts 
         Caption         =   "Find Un-posted Amounts"
      End
   End
End
Attribute VB_Name = "bimpstkhist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub post_stockhistory(p As stkhist)
    Dim ds As ADODB.Recordset, s As String
    'MsgBox "no post"
    'Exit Sub
    s = "select id from stockhistory where branchwhs = '" & p.branchwhs & "'"
    s = s & " and sku = '" & p.sku & "'"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        p.id = ds!id
        s = "Update stockhistory set startdate = '" & p.startdate & "'"
        s = s & ", enddate = '" & p.enddate & "'"
        s = s & ", postdate = '" & p.postdate & "'"
        s = s & ", totaldays = " & p.totaldays
        s = s & ", daysin = " & p.daysin
        s = s & ", daysout = " & p.daysout
        s = s & ", loads = " & p.loads
        s = s & " Where id = " & p.id
        'MsgBox s
        wdb.Execute s
    Else
        p.id = wd_seq("stockhistory")
        s = "Insert into stockhistory (id, branchwhs, sku, startdate, enddate, postdate, totaldays"
        s = s & ", daysin, daysout, loads) Values (" & p.id
        s = s & ", '" & p.branchwhs & "'"
        s = s & ", '" & p.sku & "'"
        s = s & ", '" & p.startdate & "'"
        s = s & ", '" & p.enddate & "'"
        s = s & ", '" & p.postdate & "'"
        s = s & ", " & p.totaldays
        s = s & ", " & p.daysin
        s = s & ", " & p.daysout
        s = s & ", " & p.loads & ")"
        'MsgBox s
        wdb.Execute s
    End If
    ds.Close
End Sub

Private Sub refresh_grid2()
    Dim cfile As String, s As String, i As Integer, k As Integer ', sdate As String, edate As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, ohstart As Boolean, ohend As Boolean
    Dim ddate As String, ds As ADODB.Recordset, fdate As String, j As Integer, xdaze As Integer
    'Grid2.Redraw = False
    Grid2.FontName = "Callibri"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 7
    Command1.Visible = False
    
    'i = DateDiff("d", sdate, Now)
    i = DateDiff("d", sdate, edate)
    'MsgBox i
    Screen.MousePointer = 11
    ohstart = False
    ohend = False
    ddate = Format(DateAdd("d", 30, Now), "MM-dd-yyyy")
    s = "select discdate from discont where sku = '" & Combo2 & "'"
    s = s & " and plantwhs = '" & branchrec(Val(List1)).supplier & "'"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        ddate = Format(ds!discdate, "MM-dd-yyyy")
    End If
    ds.Close
    
    For k = 0 To i '- 1
        fdate = Format(DateAdd("d", k, sdate), "MM-dd-yyyy")
        j = DateDiff("d", fdate, ddate)
        'MsgBox fdate & " " & ddate & " " & j
        'cfile = "s:\wd\html\stock\stk" & Format(DateAdd("d", k, sdate), "MMddyyyy") & ".csv"
        cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\stock\" & List1 & "\stk" & Format(DateAdd("d", k, sdate), "MMddyyyy") & ".csv"
        'MsgBox cfile
        
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            postdate.Caption = Format(DateAdd("d", k, sdate), "M-dd-yyyy")
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8
                If f0 = List1 And f1 = Combo2 Then
                    'otot = otot + Val(f5)
                    If Val(f5) > 0 Then ohstart = True
                    's = f1 & Chr(9)
                    s = Format(DateAdd("d", k, sdate), "M-d-yyyy") & Chr(9)
                    s = s & f2 & Chr(9)
                    s = s & f3 & Chr(9)
                    s = s & f4 & Chr(9)
                    s = s & f5 & Chr(9)
                    s = s & f7 & Chr(9)
                    s = s & f8
                    
                    'If otot > 0 Then
                    If ohstart = True And ohend = False Then
                        Grid2.AddItem s
                    End If
                    
                    If j < 0 And f7 > " " Then
                        'MsgBox fdate & " ending date reached.."
                        ohend = True
                    End If
                    
                    
                End If
            Loop
            Close #1
        'Else
        '    s = Format(DateAdd("d", k, sdate), "M-d-yyyy") & Chr(9)
        '    s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "In-Stock"
        '    Grid2.AddItem s
        End If
    
    
    Next k
    
    Grid2.Redraw = True
    MsgBox "check"
    
    
    tdaze = "0": idaze = "0": odaze = "0": psales = "0": xdaze = 0
    If Grid2.Rows > 1 Then
        For i = 1 To Grid2.Rows - 1
            tdaze = Val(tdaze) + 1
            If Grid2.TextMatrix(i, 5) > "..." Then
                odaze = Val(odaze) + 1
            Else
                idaze = Val(idaze) + 1
            End If
            If Grid2.TextMatrix(i, 5) = "InActive" Then xdaze = xdaze + 1
            psales = Val(psales) + Val(Grid2.TextMatrix(i, 6))
        Next i
    End If
    
    If Val(psales) = 0 Then
        tdaze = "0": idaze = "0": odaze = "0": psales = "0": xdaze = 0
        Grid2.Rows = 1
    End If
    
    If xdaze > 90 Then                       'Out of rotation?
        'MsgBox "rotation.... " & Combo2 & " " & List1
        For i = Grid2.Rows - 1 To 1 Step -1
            If Grid2.TextMatrix(i, 5) = "InActive" Then
                k = i
                Exit For
            End If
        Next i
        If k < Val(tdaze) Then                  'last row not inactive
            'MsgBox k
            For i = k To 1 Step -1
                Grid2.RemoveItem i
            Next i
            tdaze = "0": idaze = "0": odaze = "0": psales = "0"
            If Grid2.Rows > 1 Then
                For i = 1 To Grid2.Rows - 1
                    tdaze = Val(tdaze) + 1
                    If Grid2.TextMatrix(i, 5) > "..." Then
                        odaze = Val(odaze) + 1
                    Else
                        idaze = Val(idaze) + 1
                    End If
                    psales = Val(psales) + Val(Grid2.TextMatrix(i, 6))
                Next i
            End If
        Else                                    'last row inactive
            MsgBox "last row inactive"
            For i = 1 To Grid2.Rows - 1
                'If Val(Grid2.TextMatrix(i, 6)) > 0 Then k = i
                If Grid2.TextMatrix(i, 5) <> "InActive" Then k = i
            Next i
            'MsgBox k & " clear"
            If k < Val(tdaze) Then
                For i = Grid2.Rows - 1 To k Step -1
                    Grid2.RemoveItem i
                Next i
            End If
            tdaze = "0": idaze = "0": odaze = "0": psales = "0"
            If Grid2.Rows > 1 Then
                For i = 1 To Grid2.Rows - 1
                    tdaze = Val(tdaze) + 1
                    If Grid2.TextMatrix(i, 5) > "..." Then
                        odaze = Val(odaze) + 1
                    Else
                        idaze = Val(idaze) + 1
                    End If
                    psales = Val(psales) + Val(Grid2.TextMatrix(i, 6))
                Next i
            End If
        End If
    End If
    
    
    Screen.MousePointer = 0
    
    Grid2.FormatString = "^Date|^Start|^TransIn|^TransOut|^OnHand|^Status|^Loads"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 800
    Grid2.ColWidth(2) = 900
    Grid2.ColWidth(3) = 900
    Grid2.ColWidth(4) = 800
    Grid2.ColWidth(5) = 900
    Grid2.ColWidth(6) = 1100
    Grid2.Redraw = True
    If Grid2.Rows > 1 Then Command1.Visible = True
End Sub

Private Sub refresh_grid3()
    Dim i As Integer, s As String, ds As ADODB.Recordset
    For i = 1 To Grid3.Rows - 1
        Grid3.TextMatrix(i, 1) = " "
        Grid3.TextMatrix(i, 2) = " "
    Next i
    Grid3.FormatString = "^Field|^Current|^Update"
    Grid3.ColWidth(0) = 1800
    Grid3.ColWidth(1) = 1500
    Grid3.ColWidth(2) = 1500
    Grid3.TextMatrix(1, 0) = "Start Date"
    Grid3.TextMatrix(2, 0) = "End Date"
    Grid3.TextMatrix(3, 0) = "Post Date"
    Grid3.TextMatrix(4, 0) = "Total Days"
    Grid3.TextMatrix(5, 0) = "Days In"
    Grid3.TextMatrix(6, 0) = "Days Out"
    Grid3.TextMatrix(7, 0) = "Loads"
    s = "select * from stockhistory where branchwhs = '" & List1 & "'"
    s = s & " and sku = '" & Combo2 & "'"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Grid3.TextMatrix(1, 1) = ds!startdate
        If IsDate(ds!startdate) Then Grid3.TextMatrix(1, 1) = Format(ds!startdate, "M-d-yyyy")
        Grid3.TextMatrix(2, 1) = ds!enddate
        If IsDate(ds!enddate) Then Grid3.TextMatrix(2, 1) = Format(ds!enddate, "M-d-yyyy")
        Grid3.TextMatrix(3, 1) = ds!postdate
        If IsDate(ds!postdate) Then Grid3.TextMatrix(3, 1) = Format(ds!postdate, "M-d-yyyy")
        Grid3.TextMatrix(4, 1) = ds!totaldays
        Grid3.TextMatrix(5, 1) = ds!daysin
        Grid3.TextMatrix(6, 1) = ds!daysout
        Grid3.TextMatrix(7, 1) = ds!loads
        sdate = ds!postdate
        edate = Format(Now, "M-d-yyyy")
    End If
    ds.Close
End Sub

Private Sub refresh_grid4()
    Dim cfile As String, s As String, i As Integer, k As Integer ', sdate As String, edate As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, ohstart As Boolean, ohend As Boolean
    Dim ddate As String, ds As ADODB.Recordset, fdate As String, j As Integer, xdaze As Integer
    Dim pdate As String, p As stkhist
    Dim date1 As String, date2 As String
    Dim ptdaze As Long, pidaze As Long, podaze As Long, ppsales As Long
    Grid4.Redraw = False
    Grid4.FontName = "Callibri"
    Grid4.FontBold = True
    Grid4.FontSize = 8
    Grid4.Clear: Grid4.Rows = 1: Grid4.Cols = 7
    date1 = Format(DateAdd("d", 1, Grid3.TextMatrix(3, 1)), "M-dd-yyyy")
    date2 = Format(Now, "M-dd-yyyy")
    i = DateDiff("d", date1, date2)
    'MsgBox i
    Screen.MousePointer = 11
    ohstart = False
    ohend = False
    ddate = Format(DateAdd("d", 30, Now), "MM-dd-yyyy")
    s = "select discdate from discont where sku = '" & Combo2 & "'"
    s = s & " and plantwhs = '" & branchrec(Val(List1)).supplier & "'"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        ddate = Format(ds!discdate, "MM-dd-yyyy")
    End If
    ds.Close
    
    For k = 0 To i '- 1
        fdate = Format(DateAdd("d", k, date1), "MM-dd-yyyy")
        j = DateDiff("d", fdate, ddate)
        'MsgBox fdate & " " & ddate & " " & j
        'cfile = "s:\wd\html\stock\stk" & Format(DateAdd("d", k, sdate), "MMddyyyy") & ".csv"
        cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\stock\" & List1 & "\stk" & Format(DateAdd("d", k, date1), "MMddyyyy") & ".csv"
        'MsgBox cfile
        
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            'MsgBox cfile
            pdate = Format(DateAdd("d", k, date1), "M-dd-yyyy")
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8
                If f0 = List1 And f1 = Combo2 Then
                    'MsgBox f0 & " " & f1 & " " & f5
                    'otot = otot + Val(f5)
                    'If Val(f5) > 0 Then ohstart = True
                    If Val(f2) <> 0 Or Val(f3) <> 0 Or Val(f4) <> 0 Or Val(f5) <> 0 Or Val(f8) <> 0 Then ohstart = True
                    's = f1 & Chr(9)
                    s = Format(DateAdd("d", k, date1), "M-d-yyyy") & Chr(9)
                    s = s & f2 & Chr(9)
                    s = s & f3 & Chr(9)
                    s = s & f4 & Chr(9)
                    s = s & f5 & Chr(9)
                    s = s & f7 & Chr(9)
                    s = s & f8
                    
                    'If otot > 0 Then
                    If ohstart = True And ohend = False Then
                        Grid4.AddItem s
                    End If
                    
                    If j < 0 And f7 > " " Then
                        'MsgBox fdate & " ending date reached.."
                        ohend = True
                    End If
                    
                    
                End If
            Loop
            Close #1
        'Else
        '    s = Format(DateAdd("d", k, sdate), "M-d-yyyy") & Chr(9)
        '    s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "In-Stock"
        '    Grid4.AddItem s
        End If
    
    
    Next k
    
    'Grid4.Redraw = True
    'MsgBox "check"
    
    ptdaze = 0: pidaze = 0: podaze = 0: ppsales = 0: xdaze = 0
    If Grid4.Rows > 1 Then
        For i = 1 To Grid4.Rows - 1
            ptdaze = ptdaze + 1
            If Grid4.TextMatrix(i, 5) > "..." Then
                podaze = podaze + 1
            Else
                pidaze = pidaze + 1
            End If
            If Grid4.TextMatrix(i, 5) = "InActive" Then xdaze = xdaze + 1
            ppsales = ppsales + Val(Grid4.TextMatrix(i, 6))
        Next i
    End If
    
    'If ppsales = 0 Then
    '    ptdaze = 0: pidaze = 0: podaze = 0: ppsales = 0: xdaze = 0
    '    Grid4.Rows = 1
    'End If
    
    'If xdaze > 45 Then                       'Out of rotation?
    If xdaze >= ptdaze Then                       'Out of rotation?
        'MsgBox "rotation.... " & Combo2 & " " & List1
        For i = Grid4.Rows - 1 To 1 Step -1
            If Grid4.TextMatrix(i, 5) = "InActive" Then
                k = i
                Exit For
            End If
        Next i
        If k < ptdaze Then                  'last row not inactive
            'MsgBox k
            For i = k To 1 Step -1
                Grid4.RemoveItem i
            Next i
            ptdaze = 0: pidaze = 0: podaze = 0: ppsales = 0
            If Grid4.Rows > 1 Then
                For i = 1 To Grid4.Rows - 1
                    ptdaze = ptdaze + 1
                    If Grid4.TextMatrix(i, 5) > "..." Then
                        podaze = podaze + 1
                    Else
                        pidaze = pidaze + 1
                    End If
                    ppsales = ppsales + Val(Grid4.TextMatrix(i, 6))
                Next i
            End If
        Else                                    'last row inactive
            For i = 1 To Grid4.Rows - 1
                'If Val(Grid4.TextMatrix(i, 6)) > 0 Then k = i
                If Grid4.TextMatrix(i, 5) <> "InActive" Then k = i
            Next i
            'MsgBox k & " clear"
            If k < ptdaze Then
                For i = Grid4.Rows - 1 To k Step -1
                    Grid4.RemoveItem i
                Next i
            End If
            ptdaze = 0: pidaze = 0: podaze = 0: ppsales = "0"
            If Grid4.Rows > 1 Then
                For i = 1 To Grid4.Rows - 1
                    ptdaze = ptdaze + 1
                    If Grid4.TextMatrix(i, 5) > "..." Then
                        podaze = podaze + 1
                    Else
                        pidaze = pidaze + 1
                    End If
                    ppsales = ppsales + Val(Grid4.TextMatrix(i, 6))
                Next i
            End If
        End If
    End If
    
    
    Screen.MousePointer = 0
    
    Grid4.FormatString = "^Date|^Start|^TransIn|^TransOut|^OnHand|^Status|^Loads"
    Grid4.ColWidth(0) = 1000
    Grid4.ColWidth(1) = 800
    Grid4.ColWidth(2) = 900
    Grid4.ColWidth(3) = 900
    Grid4.ColWidth(4) = 800
    Grid4.ColWidth(5) = 900
    Grid4.ColWidth(6) = 1100
    Grid4.Redraw = True
End Sub

Private Sub refresh_lists()
    Dim i As Integer
    For i = 1 To 99
        If branchrec(i).oraloc > " " Then
            Combo1.AddItem Format(branchrec(i).branchno, "000") & "-" & branchrec(i).branchname
            List1.AddItem Format(branchrec(i).branchno, "000")
        End If
    Next i
    Combo1.ListIndex = 0
    'For i = 1 To 9999
    '    If skurec(i).wrapunits > 0 Then
    '        Combo2.AddItem i
    '        List2.AddItem skurec(i).unit & " " & skurec(i).desc
    '    End If
    'Next i
    'Combo2.ListIndex = 0
    refresh_skus
End Sub

Private Sub refresh_skus()
    Dim ds As ADODB.Recordset, s As String, i As Integer
    Combo2.Clear: List2.Clear
    s = "select distinct sku from bimp where branchwhs = '" & List1 & "'"
    s = s & " and plantwhs in ('A10', 'K10', 'T10') order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = Val(ds!sku)
            Combo2.AddItem ds!sku
            List2.AddItem skurec(i).unit & " " & skurec(i).desc
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    refresh_skus
End Sub

Private Sub Combo2_Click()
    List2.ListIndex = Combo2.ListIndex
    Grid2.Rows = 1
    Command1.Visible = False
    tdaze = "0": idaze = "0": odaze = "0": psales = "0"
    Command6_Click
End Sub

Private Sub Command1_Click()
    Dim p As stkhist
    p.id = 0
    p.branchwhs = List1
    p.sku = Combo2
    If Grid2.Rows > 2 Then
        p.startdate = Format(Grid2.TextMatrix(1, 0), "M-d-yyyy")
        p.enddate = Format(Grid2.TextMatrix(Grid2.Rows - 1, 0), "M-d-yyyy")
    Else
        p.startdate = "N/A"
        p.enddate = "N/A"
    End If
    p.postdate = postdate.Caption
    p.totaldays = Val(tdaze)
    p.daysin = Val(idaze)
    p.daysout = Val(odaze)
    p.loads = Val(psales)
    Call post_stockhistory(p)
End Sub

Private Sub Command2_Click()
    sdate = Grid3.TextMatrix(1, 1)
    sdate = InputBox("Start Date:", "Starting date...", sdate)
    If Len(sdate) = 0 Then Exit Sub
    edate = InputBox("End Date:", "Ending date...", edate)
    If Len(edate) = 0 Then Exit Sub
    refresh_grid2
End Sub

Private Sub Command3_Click()
    Dim i As Integer, s As String, k As Integer
    Dim rt As String, rh As String, rf As String, hfile As String
    pgrid.FontName = "Callibri"
    pgrid.FontBold = True
    pgrid.FontSize = 8
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 9
    For i = 0 To Combo2.ListCount - 1
        Combo2.ListIndex = i
        refresh_grid2
        DoEvents
        k = Val(Combo2)
        s = Combo2 & Chr(9)
        s = s & skurec(k).unit & " "
        s = s & skurec(k).desc & Chr(9)
        If Grid2.Rows > 2 Then
            s = s & Grid2.TextMatrix(1, 0) & Chr(9)
            s = s & Grid2.TextMatrix(Grid2.Rows - 1, 0) & Chr(9)
        Else
            s = s & "N/A" & Chr(9) & "N/A" & Chr(9)
        End If
        s = s & tdaze & Chr(9)
        s = s & idaze & Chr(9)
        s = s & odaze & Chr(9)
        s = s & psales & Chr(9)
        If psales > 0 And idaze > 0 Then
            s = s & Format((psales / idaze) * odaze, "0")
        Else
            s = s & "."
        End If
        pgrid.AddItem s
        DoEvents
    Next i
    pgrid.RowSel = pgrid.Row
    pgrid.Col = 1: pgrid.ColSel = 1
    pgrid.Sort = 5
    'pgrid.FormatString = "^SKU|^Unit|<Flavor|^Start Date|^Total Days|^Days In-Stock|^Days Out-of-Stock|^Loads|^Lost Sales"
    pgrid.FormatString = "^SKU|<Product|^Start Date|^End Date|^Total Days|^Days In-Stock|^Days Out-of-Stock|^Loads|^Lost Sales"
    pgrid.ColWidth(0) = 1000
    pgrid.ColWidth(1) = 3800
    pgrid.ColWidth(2) = 1800
    pgrid.ColWidth(3) = 1800
    pgrid.ColWidth(4) = 1800
    pgrid.ColWidth(5) = 1800
    pgrid.ColWidth(6) = 1800
    pgrid.ColWidth(7) = 1800
    pgrid.ColWidth(8) = 1800
    'rt = Me.Caption
    rt = "SKU Stock History"
    rt = rt & "<br>" & Combo1
    rh = sdate & " Thru " & edate
    rf = "Printed:  " & Format(Now, "M-d-yyyy h:mm:ss am/pm")
    Exit Sub
    'Excel
    hfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\stock\" & List1 & "\ostk" & List1 & ".xls"
    Call htmlcolorgrid(Me, hfile, pgrid, rt, rh, rf, "lemonchiffon", "linen", "white")
    'HTML
    hfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\stock\" & List1 & "\ostk" & List1 & ".htm"
    Call htmlcolorgrid(Me, hfile, pgrid, rt, rh, rf, "lemonchiffon", "linen", "white")
    Form1.WebBrowser1.Navigate hfile
    'Unload Me
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe " & hfile, vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & hfile, vbNormalFocus)
        Exit Sub
    End If
    
End Sub

Private Sub Command4_Click()
    Dim rt As String, rh As String, rf As String, hfile As String
    rt = Me.Caption & "  " & Combo1 & "<br>" & sdate & " thru " & edate
    rt = rt & "<br>" & Combo2 & " " & List2
    rh = "Total Days:  " & tdaze
    rh = rh & "<br>Days In-Stock:  " & idaze
    rh = rh & "   Days Out-of-Stock:  " & odaze
    rh = rh & "<br>Total Loads:  " & psales
    rf = "Printed:  " & Format(Now, "M-d-yyyy h:mm:ss am/pm")
    'EXCEL
    hfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\stock\" & List1 & "\skustk" & List1 & ".xls"
    Call htmlcolorgrid(Me, hfile, Grid2, rt, rh, rf, "lemonchiffon", "linen", "white")
    'HTML
    hfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\stock\" & List1 & "\skustk" & List1 & ".htm"
    Call htmlcolorgrid(Me, hfile, Grid2, rt, rh, rf, "lemonchiffon", "linen", "white")
    Form1.WebBrowser1.Navigate hfile
    'Unload Me
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe " & hfile, vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & hfile, vbNormalFocus)
        Exit Sub
    End If

End Sub

Private Sub Command5_Click()
    Dim i As Integer
    For i = Combo1.ListIndex To Combo1.ListCount - 1
        Combo1.ListIndex = i
        DoEvents
        Command3_Click
        DoEvents
    Next i
End Sub

Private Sub Command6_Click()
    Command7.Visible = False
    refresh_grid3
    DoEvents
    If Grid3.TextMatrix(1, 1) <= " " Then
        Grid4.Rows = 1
        Exit Sub
    End If
    refresh_grid4
    Grid3.TextMatrix(1, 2) = Grid3.TextMatrix(1, 1)
    Grid3.TextMatrix(2, 2) = Grid3.TextMatrix(2, 1)
    Grid3.TextMatrix(3, 2) = Grid3.TextMatrix(3, 1)
    Grid3.TextMatrix(4, 2) = Grid3.TextMatrix(4, 1)
    Grid3.TextMatrix(5, 2) = Grid3.TextMatrix(5, 1)
    Grid3.TextMatrix(6, 2) = Grid3.TextMatrix(6, 1)
    Grid3.TextMatrix(7, 2) = Grid3.TextMatrix(7, 1)
    
    If Grid4.Rows > 1 Then
        For i = 1 To Grid4.Rows - 1
            If Grid4.TextMatrix(i, 5) = "InActive" Then
                Grid3.TextMatrix(3, 2) = Grid4.TextMatrix(i, 0)
            Else
                Grid3.TextMatrix(2, 2) = Grid4.TextMatrix(i, 0)
                Grid3.TextMatrix(3, 2) = Grid4.TextMatrix(i, 0)
                Grid3.TextMatrix(4, 2) = Val(Grid3.TextMatrix(4, 2)) + 1
                If Grid4.TextMatrix(i, 5) = "Out" Then
                    'Grid3.TextMatrix(5, 2) = Grid3.TextMatrix(5, 1)
                    Grid3.TextMatrix(6, 2) = Val(Grid3.TextMatrix(6, 2)) + 1
                Else
                    Grid3.TextMatrix(5, 2) = Val(Grid3.TextMatrix(5, 2)) + 1
                    'Grid3.TextMatrix(6, 2) = Grid3.TextMatrix(6, 1)
                End If
                Grid3.TextMatrix(7, 2) = Val(Grid3.TextMatrix(7, 2)) + Val(Grid4.TextMatrix(i, 6))
            End If
        Next i
        If Grid3.TextMatrix(2, 1) = Grid3.TextMatrix(3, 1) Then Command7.Visible = True
    End If
End Sub

Private Sub Command7_Click()
    Dim p As stkhist
    p.id = 0
    p.branchwhs = List1
    p.sku = Combo2
    p.startdate = Grid3.TextMatrix(1, 2)
    p.enddate = Grid3.TextMatrix(2, 2)
    p.postdate = Grid3.TextMatrix(3, 2)
    p.totaldays = Val(Grid3.TextMatrix(4, 2))
    p.daysin = Val(Grid3.TextMatrix(5, 2))
    p.daysout = Val(Grid3.TextMatrix(6, 2))
    p.loads = Val(Grid3.TextMatrix(7, 2))
    Call post_stockhistory(p)
    DoEvents
    Call Command6_Click
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    refresh_lists
    sdate = "9-01-2017"
    edate = Format(Now, "M-dd-yyyy")
End Sub

Private Sub Form_Resize()
    'If Me.Height > 2000 Then
    '    Grid2.Height = Me.Height - (Combo1.Height * 7)
    'End If
    pgrid.Width = Me.Width - 200
End Sub

Private Sub Grid3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu g3menu
End Sub

Private Sub List2_Click()
    pdesc = List2
End Sub

Private Sub List3_Click()
    Dim i As Integer
    i = Val(Mid(List3, 1, 3))
    Combo2.ListIndex = i
End Sub

Private Sub pabskus_Click()
    Dim i As Integer
    For i = 0 To Combo2.ListCount - 1
        Combo2.ListIndex = i
        DoEvents
        If Command7.Visible = True Then
            'MsgBox "Process"
            Call Command7_Click
            DoEvents
        'Else
        '    MsgBox "Skip"
        End If
    Next i
End Sub

Private Sub pallbskus_Click()
    Dim i As Integer, k As Integer
    For k = 0 To Combo1.ListCount - 1
        Combo1.ListIndex = k
        DoEvents
        For i = 0 To Combo2.ListCount - 1
            Combo2.ListIndex = i
            DoEvents
            If Command7.Visible = True Then
                'MsgBox "Process"
                Call Command7_Click
                DoEvents
            'Else
            '    MsgBox "Skip"
            End If
        Next i
    Next k
End Sub

Private Sub unpamts_Click()
    Dim i As Integer, s As String
    List3.Clear
    For i = Combo2.ListIndex To Combo2.ListCount - 1
        Combo2.ListIndex = i
        DoEvents
        If Grid4.Rows > 1 Then
            s = Format(i, "000") & " " & Combo2.List(i) & " " & List2.List(i)
            List3.AddItem s
            'Exit For
        End If
    Next i
    If List3.ListCount > 0 Then List3.ListIndex = 0
End Sub
