VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form wdbtrkwo 
   Caption         =   "Transport Schedule"
   ClientHeight    =   11265
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   ScaleHeight     =   11265
   ScaleWidth      =   14805
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox sdest 
      Height          =   3375
      Left            =   8760
      TabIndex        =   23
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox sorg 
      Height          =   1230
      Left            =   7680
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   3615
      Left            =   0
      TabIndex        =   20
      Top             =   7200
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6376
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   3015
      Left            =   0
      TabIndex        =   16
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5318
      _Version        =   327680
      FixedCols       =   0
      BackColorFixed  =   12648384
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1335
      Left            =   0
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2355
      _Version        =   327680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox edate 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   13
      Text            =   "edate"
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox sdate 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Text            =   "sdate"
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox scont 
      Height          =   315
      Left            =   11160
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.ComboBox stype 
      Height          =   315
      Left            =   9240
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.ComboBox sdriver 
      Height          =   315
      Left            =   6720
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   240
      Width           =   2535
   End
   Begin VB.ComboBox slocation 
      Height          =   315
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   240
      Width           =   2415
   End
   Begin VB.ComboBox splant 
      Height          =   315
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Posted:"
      Height          =   255
      Left            =   5280
      TabIndex        =   22
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label schdate 
      Caption         =   "Label2"
      Height          =   255
      Left            =   6600
      TabIndex        =   19
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label plantno 
      Caption         =   "Label2"
      Height          =   255
      Left            =   8880
      TabIndex        =   18
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label hcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8880
      TabIndex        =   17
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contents"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   11160
      TabIndex        =   6
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Work Type"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   9240
      TabIndex        =   5
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Driver"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   4
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Destination"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   3
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plant"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "End"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Start"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu prtmenu 
      Caption         =   "&Print"
      Begin VB.Menu prtinb 
         Caption         =   "&In Bound Trailers"
      End
      Begin VB.Menu prtoutb 
         Caption         =   "&Out Bound Trailers"
      End
      Begin VB.Menu prtphone 
         Caption         =   "Plant &Drivers"
      End
      Begin VB.Menu prtlist 
         Caption         =   "Current &List"
      End
   End
   Begin VB.Menu sortmenu 
      Caption         =   "&Sort"
      Begin VB.Menu sortstart 
         Caption         =   "&Start Time"
         Checked         =   -1  'True
      End
      Begin VB.Menu sorttrip 
         Caption         =   "&Trip"
      End
      Begin VB.Menu sortdriver 
         Caption         =   "&Driver"
      End
   End
End
Attribute VB_Name = "wdbtrkwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub location_list()
    Dim i As Integer, k As Integer, mcode As String, mdest As String, sd As String
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 3
    
    For i = 1 To Grid1.Rows - 1
        k = InStr(1, Grid1.TextMatrix(i, 2), ">")
        If k > 0 Then
            mcode = Grid1.TextMatrix(i, 18)
            'morg = Trim(Left(Grid2.TextMatrix(i, 1), k - 1))
            mdest = Trim(Right(Grid1.TextMatrix(i, 2), Len(Grid1.TextMatrix(i, 2)) - k))
            pgrid.AddItem mcode & Chr(9) & mdest
        End If
    Next i
    pgrid.RowSel = pgrid.Row
    pgrid.Col = 1: pgrid.ColSel = 1
    pgrid.Sort = 5
    
    sdest.Clear: slocation.Clear
    sdest.AddItem " ": slocation.AddItem "All"
    sd = "..."
    For i = 1 To pgrid.Rows - 1
        If pgrid.TextMatrix(i, 0) <> sd Then
            sdest.AddItem pgrid.TextMatrix(i, 0)
            slocation.AddItem pgrid.TextMatrix(i, 1)
            pgrid.TextMatrix(i, 2) = "X"
            sd = pgrid.TextMatrix(i, 0)
        End If
    Next i
    
    pgrid.FormatString = "^Code|<Description|^Mark"
    pgrid.ColWidth(0) = 1000
    pgrid.ColWidth(1) = 3000
    pgrid.ColWidth(2) = 1000
    slocation.ListIndex = 0
End Sub

Private Sub refresh_schedule()
    Dim cfile As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String, f5 As String, f6 As String
    Dim f7 As String, f8 As String, f9 As String, f10 As String, f11 As String, f12 As String, f13 As String
    Dim f14 As String, f15 As String, f16 As String, f17 As String, f18 As String
    Dim s As String
    Grid1.Clear
    Grid1.Cols = 19: Grid1.Rows = 1: Grid1.FixedCols = 0
    
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\schedule\truckwo.csv"
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17, f18
        If Val(f0) > 0 Then
            s = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & f4 & Chr(9) & f5 & Chr(9)
            s = s & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9) & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9)
            s = s & f12 & Chr(9) & f13 & Chr(9) & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
            s = s & f17 & Chr(9) & f18
            Grid1.AddItem s
        End If
    Loop
    Close #1
    
    Grid1.FormatString = "^WO|<Date|^Trip|^Comments|^#|^Driver|^Size|^Start|^Hours|^Work Type|^Contents|^Meals|^Status|^Parent|<SortStart|SortTrip|SortDriver|Origin|Destination"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 4000
    Grid1.ColWidth(3) = 3000
    Grid1.ColWidth(4) = 400
    Grid1.ColWidth(5) = 2400
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 700
    Grid1.ColWidth(8) = 900
    Grid1.ColWidth(9) = 1175
    Grid1.ColWidth(10) = 1500
    Grid1.ColWidth(11) = 1500
    Grid1.ColWidth(12) = 1500
    Grid1.ColWidth(13) = 1700
    Grid1.ColWidth(14) = 1700
    Grid1.ColWidth(15) = 1700
    Grid1.ColWidth(16) = 1700
    Grid1.ColWidth(17) = 1700
    Grid1.ColWidth(18) = 1700
    If Grid1.Rows > 1 Then
        sdate = Format(Grid1.TextMatrix(1, 1), "m-d-yyyy")
        edate = Format(Grid1.TextMatrix(Grid1.Rows - 1, 1), "m-d-yyyy")
        If sortstart.Checked = True Then
            Grid1.RowSel = Grid1.Row
            Grid1.Col = 14: Grid1.ColSel = 14
            Grid1.Sort = 5
        End If
        If sorttrip.Checked = True Then
            Grid1.RowSel = Grid1.Row
            Grid1.Col = 12: Grid1.ColSel = 2
            Grid1.Sort = 5
        End If
        If sortdriver.Checked = True Then
            Grid1.RowSel = Grid1.Row
            Grid1.Col = 16: Grid1.ColSel = 16
            Grid1.Sort = 5
        End If
    End If
End Sub

Private Sub refresh_qry()
    Dim i As Integer, s As String, nr As Boolean, c As Boolean
    Dim cfile As String
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\schedule\truckwo.csv"
    If Len(Dir(cfile)) > 0 Then
        If schdate.Caption <> FileDateTime(cfile) Then
            refresh_schedule
            DoEvents
            schdate = FileDateTime(cfile)
        End If
    End If
    
    Grid2.Redraw = False
    Screen.MousePointer = 11
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 14
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            nr = True
            If sorg > " " And nr = True Then
                If Grid1.TextMatrix(i, 17) <> sorg And Grid1.TextMatrix(i, 18) <> sorg Then nr = False
                If sorg = "K10" And Left(Grid1.TextMatrix(i, 17), 1) = "K" And Val(Right(Grid1.TextMatrix(i, 17), Len(Grid1.TextMatrix(i, 17)) - 1)) > 0 Then nr = True
                If sorg = "K10" And Left(Grid1.TextMatrix(i, 18), 1) = "K" And Val(Right(Grid1.TextMatrix(i, 18), Len(Grid1.TextMatrix(i, 18)) - 1)) > 0 Then nr = True
                If sorg = "K10" And Grid1.TextMatrix(i, 17) = "047" Then nr = True
                If sorg = "K10" And Grid1.TextMatrix(i, 18) = "047" Then nr = True
                If sorg = "T10" And Left(Grid1.TextMatrix(i, 17), 1) = "T" And Val(Right(Grid1.TextMatrix(i, 17), Len(Grid1.TextMatrix(i, 17)) - 1)) > 0 Then nr = True
                If sorg = "T10" And Left(Grid1.TextMatrix(i, 18), 1) = "T" And Val(Right(Grid1.TextMatrix(i, 18), Len(Grid1.TextMatrix(i, 18)) - 1)) > 0 Then nr = True
                If sorg = "T10" And Grid1.TextMatrix(i, 17) = "047" Then nr = True
                If sorg = "T10" And Grid1.TextMatrix(i, 18) = "047" Then nr = True
                If sorg = "A10" And Left(Grid1.TextMatrix(i, 17), 1) = "A" And Val(Right(Grid1.TextMatrix(i, 17), Len(Grid1.TextMatrix(i, 17)) - 1)) > 0 Then nr = True
                If sorg = "A10" And Left(Grid1.TextMatrix(i, 18), 1) = "A" And Val(Right(Grid1.TextMatrix(i, 18), Len(Grid1.TextMatrix(i, 18)) - 1)) > 0 Then nr = True
                If sorg = "A10" And Grid1.TextMatrix(i, 17) = "047" Then nr = True
                If sorg = "A10" And Grid1.TextMatrix(i, 18) = "047" Then nr = True
            End If
            If sdest > " " And nr = True Then
                If Grid1.TextMatrix(i, 18) <> sdest Then nr = False
            End If
            If sdriver > " " And nr = True Then
                If Grid1.TextMatrix(i, 5) <> sdriver Then nr = False
            End If
            If stype > " " And nr = True Then
                If Grid1.TextMatrix(i, 9) <> stype Then nr = False
                If stype = "Start" And Grid1.TextMatrix(i, 9) = "SameDay" Then nr = True
                If stype = "SameDay" And Grid1.TextMatrix(i, 9) = "Start" Then nr = True
            End If
            If scont > " " And nr = True Then
                If Grid1.TextMatrix(i, 10) <> scont Then nr = False
            End If
            If nr = True Then
                s = Format(Grid1.TextMatrix(i, 1), "m-d-yyyy") & Chr(9)
                s = s & Grid1.TextMatrix(i, 2) & Chr(9)
                s = s & Grid1.TextMatrix(i, 3) & Chr(9)
                s = s & Grid1.TextMatrix(i, 4) & Chr(9)
                s = s & Grid1.TextMatrix(i, 5) & Chr(9)
                s = s & Grid1.TextMatrix(i, 6) & Chr(9)
                s = s & Grid1.TextMatrix(i, 7) & Chr(9)
                s = s & Grid1.TextMatrix(i, 8) & Chr(9)
                s = s & Grid1.TextMatrix(i, 9) & Chr(9)
                s = s & Grid1.TextMatrix(i, 10) & Chr(9)
                s = s & Chr(9) 'Grid1.TextMatrix(i, 11) & Chr(9)
                s = s & Grid1.TextMatrix(i, 12) & Chr(9)
                s = s & Format(DateAdd("n", Val(Grid1.TextMatrix(i, 8)) * 60, Grid1.TextMatrix(i, 7)), "hh:mm am/pm")
                s = s & Chr(9) & i
                Grid2.AddItem s
            End If
        Next i
    End If
    If Grid2.Rows > 1 Then
        Grid2.FillStyle = flexFillRepeat
        c = True
        For i = 1 To Grid2.Rows - 1
            c = Not c
            If c = True Then
                Grid2.Row = i: Grid2.RowSel = i
                Grid2.Col = 0: Grid2.ColSel = Grid2.Cols - 1
                Grid2.CellBackColor = hcolor.BackColor
            End If
        Next i
        Grid2.Row = 1
    End If
                
    'Grid2.FormatString = "^Date|^Trip|^Comments|^#|^Driver|^Size|^Start|^Hours|^Work Type|^Contents|^Meals|^Status|^End"
    Grid2.FormatString = "^Date|^Trip|^Comments|^#|^Driver|^Size|^Start|^Hours|^Work Type|^Contents|^|^Status|^End"
    Grid2.ColWidth(0) = 1100
    Grid2.ColWidth(1) = 3500
    Grid2.ColWidth(2) = 4000
    Grid2.ColWidth(3) = 300
    Grid2.ColWidth(4) = 1800
    Grid2.ColWidth(5) = 400
    Grid2.ColWidth(6) = 800
    Grid2.ColWidth(7) = 700
    Grid2.ColWidth(8) = 900
    Grid2.ColWidth(9) = 1175
    Grid2.ColWidth(10) = 0 '800
    Grid2.ColWidth(11) = 800
    Grid2.ColWidth(12) = 800
    Grid2.ColWidth(13) = 0 '800
    
    Screen.MousePointer = 0
    Grid2.Redraw = True
End Sub

Private Sub refresh_qlists()
    Dim cfile As String, f0 As String, f1 As String
    'sdest.Clear: sdest.AddItem ""
    sdriver.Clear: sdriver.AddItem ""
    stype.Clear: stype.AddItem ""
    scont.Clear: scont.AddItem ""
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\schedule\qlist.csv"
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1
        'If f0 = "Location" Then sorg.AddItem f1
        'If f0 = "Location" Then sdest.AddItem f1
        If f0 = "Drivers" Then sdriver.AddItem f1
        If f0 = "Work Type" Then stype.AddItem f1
        If f0 = "Contents" Then scont.AddItem f1
    Loop
    Close #1
    sorg.Clear: splant.Clear
    sorg.AddItem "": splant.AddItem "All"
    sorg.AddItem "T10": splant.AddItem "Brenham"
    sorg.AddItem "K10": splant.AddItem "Broken Arrow"
    sorg.AddItem "A10": splant.AddItem " Sylacauga"
    splant.ListIndex = 0
    'sdest.ListIndex = 0
    sdriver.ListIndex = 0
    stype.ListIndex = 0
    scont.ListIndex = 0
End Sub

Private Sub Command1_Click()
    refresh_qry
End Sub

Private Sub Form_Load()
    Dim cfile As String
    refresh_qlists
    refresh_schedule
    DoEvents
    location_list
    refresh_qry
    cfile = "s:\wd\html\schedule\truckwo.csv"
    If Len(Dir(cfile)) > 0 Then
        schdate = FileDateTime(cfile)
    End If
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    Grid2.Width = Me.Width - 100
    pgrid.Width = Me.Width - 100
    If Me.Height > 2000 Then Grid2.Height = Me.Height - 1800
End Sub

Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu sortmenu
End Sub

Private Sub plantno_Change()
    If Me.plantno = "50" Then
        For i = 0 To sorg.ListCount - 1
            If sorg.List(i) = "T10" Then
                splant.ListIndex = i
                Exit For
            End If
        Next i
    End If
    If Me.plantno = "51" Then
        For i = 0 To sorg.ListCount - 1
            If sorg.List(i) = "K10" Then
                splant.ListIndex = i
                Exit For
            End If
        Next i
    End If
    If Me.plantno = "52" Then
        For i = 0 To sorg.ListCount - 1
            If sorg.List(i) = "A10" Then
                splant.ListIndex = i
                Exit For
            End If
        Next i
    End If
    refresh_qry
End Sub

Private Sub prtinb_Click()
    sortstart_Click
    DoEvents
    If sorg = "A10" Then wdbphone.rtype = "inbounda10"
    If sorg = "K10" Then wdbphone.rtype = "inboundk10"
    If sorg = "T10" Or splant = "All" Then wdbphone.rtype = "inboundt10"
    wdbphone.qstr = Val(wdbphone.qstr.Caption) + 1
    wdbphone.Show
    
End Sub

Private Sub prtlist_Click()
    If sortdriver.Checked Then
        wdbphone.rtype = "clist_driver"
        DoEvents
        wdbphone.qstr = Val(wdbphone.qstr.Caption) + 1
        wdbphone.Show
    End If
    If sorttrip.Checked Then
        wdbphone.rtype = "clist_trip"
        DoEvents
        wdbphone.qstr = Val(wdbphone.qstr.Caption) + 1
        wdbphone.Show
    End If
    If sortstart.Checked Then
        wdbphone.rtype = "clist_date"
        DoEvents
        wdbphone.qstr = Val(wdbphone.qstr.Caption) + 1
        wdbphone.Show
    End If
    
    Exit Sub
    

    Dim i As Integer, s As String, morg As String, mdest As String, k As Integer, mtrl As String
    Dim rt As String, rh As String, rf As String
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 4
    For i = 1 To Grid2.Rows - 1
        mdest = " "
        morg = " "
        mtrl = ""
        If Grid2.TextMatrix(i, 8) = "Start" Then mtrl = "#" & Grid2.TextMatrix(i, 3)
        If Grid2.TextMatrix(i, 8) = "SameDay" Then mtrl = "#" & Grid2.TextMatrix(i, 3)
        If Grid2.TextMatrix(i, 8) = "Delivery" Then mtrl = "#" & Grid2.TextMatrix(i, 3)
        
        k = InStr(1, Grid2.TextMatrix(i, 1), ">")
        If k > 0 Then
            morg = Trim(Left(Grid2.TextMatrix(i, 1), k - 1))
            mdest = Trim(Right(Grid2.TextMatrix(i, 1), Len(Grid2.TextMatrix(i, 1)) - k))
        End If
        s = Format(Grid2.TextMatrix(i, 0), "ddd m-d-yy") & Chr(9)
        s = s & Format(Grid2.TextMatrix(i, 6), "h:mm am/pm") & Chr(9)
        If Grid2.TextMatrix(i, 8) = "Return" Then
            's = s & morg & " #" & Grid2.TextMatrix(i, 3) & " Return " & mdest & Chr(9)
            s = s & morg & " Return " & mdest & Chr(9)
        Else
            's = s & mdest & " #" & Grid2.TextMatrix(i, 3) & Chr(9)
            s = s & mdest & " " & mtrl & Chr(9)
        End If
        's = s & Grid2.TextMatrix(i, 1) & " #" & Grid2.TextMatrix(i, 3) & Chr(9)
        s = s & Grid2.TextMatrix(i, 4)
        pgrid.AddItem s
        s = ""
        If morg = "Sylacauga" Then s = "SY->"
        If morg = "Broken Arrow" Then s = "BA->"
        s = s & Chr(9)
        s = s & Grid2.TextMatrix(i, 7) & Chr(9)
        s = s & Grid2.TextMatrix(i, 2) & Chr(9)
        s = s & Grid2.TextMatrix(i, 9)
        pgrid.AddItem s
        pgrid.AddItem " "
    Next i
    pgrid.FormatString = "<Date|^Start|<|<"
    pgrid.ColWidth(0) = 1800
    pgrid.ColWidth(1) = 1500
    pgrid.ColWidth(2) = 5000
    pgrid.ColWidth(3) = 5000
    
    rt = Me.Caption
    rh = "Posted: " & schdate.Caption
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, pgrid, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
    
    
End Sub

Private Sub prtoutb_Click()
    sortstart_Click
    DoEvents
    If sorg = "A10" Then wdbphone.rtype = "outbounda10"
    If sorg = "K10" Then wdbphone.rtype = "outboundk10"
    If sorg = "T10" Or splant = "All" Then wdbphone.rtype = "outboundt10"
    wdbphone.qstr = Val(wdbphone.qstr.Caption) + 1
    wdbphone.Show
End Sub

Private Sub prtphone_Click()
    sortdriver_Click
    DoEvents
    If sorg = "A10" Then wdbphone.rtype = "planttrk_syl"
    If sorg = "K10" Then wdbphone.rtype = "planttrk_ba"
    If sorg = "T10" Or splant = "All" Then wdbphone.rtype = "planttrk_tx"
    wdbphone.qstr = Val(wdbphone.qstr.Caption) + 1
    wdbphone.Show
End Sub

Private Sub slocation_Click()
    sdest.ListIndex = slocation.ListIndex
End Sub

Private Sub sortdriver_Click()
    sortstart.Checked = False
    sorttrip.Checked = False
    sortdriver.Checked = True
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 16: Grid1.ColSel = 16
    Grid1.Sort = 5
    refresh_qry
End Sub

Private Sub sortstart_Click()
    sortstart.Checked = True
    sorttrip.Checked = False
    sortdriver.Checked = False
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 14: Grid1.ColSel = 14
    Grid1.Sort = 5
    refresh_qry
End Sub

Private Sub sorttrip_Click()
    sortstart.Checked = False
    sorttrip.Checked = True
    sortdriver.Checked = False
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 15: Grid1.ColSel = 15
    'Grid1.Col = 2: Grid1.ColSel = 2
    Grid1.Sort = 5
    refresh_qry
End Sub

Private Sub splant_Click()
    sorg.ListIndex = splant.ListIndex
End Sub
