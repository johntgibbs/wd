VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form srlogs 
   Caption         =   "SR Movement Log"
   ClientHeight    =   6885
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13680
   LinkTopic       =   "Form14"
   ScaleHeight     =   6885
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   1215
      Left            =   0
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2143
      _Version        =   327680
   End
   Begin VB.ListBox oplanes 
      Height          =   1620
      Left            =   7080
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox shiplanes 
      Height          =   1620
      Left            =   4440
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8705
      _Version        =   327680
      FixedCols       =   0
      ForeColor       =   64
      BackColorFixed  =   16777152
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label acolor 
      Caption         =   "acolor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10680
      TabIndex        =   15
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label pcolor 
      BackColor       =   &H00C0E0FF&
      Caption         =   "pcolor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   10680
      TabIndex        =   14
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label bcolor 
      BackColor       =   &H00FFFFC0&
      Caption         =   "bcolor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   10680
      TabIndex        =   13
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label ycolor 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ycolor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   255
      Left            =   10680
      TabIndex        =   12
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label gcolor 
      BackColor       =   &H00C0FFC0&
      Caption         =   "gcolor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   10680
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label ccol 
      Caption         =   "ccol"
      Height          =   255
      Left            =   11880
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label cntlit 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   9
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label hcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu prrtmen 
      Caption         =   "Print"
      Begin VB.Menu prtmenu 
         Caption         =   "Current List"
      End
      Begin VB.Menu shiplots 
         Caption         =   "Shipped Lots"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu sortmenu 
      Caption         =   "Sort"
      Begin VB.Menu sortrec 
         Caption         =   "Record Number"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu sorttime 
         Caption         =   "Time"
      End
      Begin VB.Menu sortsku 
         Caption         =   "BarCode"
      End
   End
   Begin VB.Menu findmenu 
      Caption         =   "Find"
      Begin VB.Menu findsku 
         Caption         =   "SKU"
      End
      Begin VB.Menu findcol 
         Caption         =   "Column"
      End
      Begin VB.Menu findbay 
         Caption         =   "Bay Activity"
         Visible         =   0   'False
      End
      Begin VB.Menu findnew 
         Caption         =   "New Bays"
      End
      Begin VB.Menu finddrops 
         Caption         =   "Drops"
         Visible         =   0   'False
      End
      Begin VB.Menu batonhand 
         Caption         =   "View Batch Inventory"
      End
      Begin VB.Menu palhist 
         Caption         =   "View Pallet History"
      End
      Begin VB.Menu findop 
         Caption         =   "Order Pick"
      End
   End
   Begin VB.Menu mockwith 
      Caption         =   "WithDrawal"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "srlogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_srlogs()
    Dim cfile As String, s As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim srpath As String, hr As Boolean
    Screen.MousePointer = 11
    srpath = Form1.logdir
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontSize = 8
    Grid1.FontBold = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 19
            
    If Combo1 = "SR5" Or Combo1 = "All" Then                                                           'jv060117
        cfile = srpath & "SR" & Format(Text1, "mmddyyyy") & ".txt"                                     'jv060117
        If Len(Dir(cfile)) > 0 Then                                                                     'jv060117
            Open cfile For Input Shared As #1                                                           'jv060117
            Do Until EOF(1)                                                                             'jv060117
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16     'jv060117
                s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & Trim(f2) & Chr(9) & Trim(f3) & Chr(9)  'jv060117
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9) 'jv060117
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)         'jv060117
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16                                               'jv060117
                Grid1.AddItem s                                                                         'jv060117
            Loop                                                                                        'jv060117
            Close #1                                                                                    'jv060117
        Else                                                                                            'jv061215
            cfile = srpath & Right(Text1, 4) & "\sr" & Format(Text1, "mmddyyyy") & ".txt"               'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16                                           'jv061215
                    Grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If                                                                                              'jv060117
        
    If Combo1 = "SR1" Or Combo1 = "All" Then                                                           'jv060117
        cfile = srpath & "SR1" & Format(Text1, "mmddyyyy") & ".txt"                                     'jv060117
        If Len(Dir(cfile)) > 0 Then                                                                     'jv060117
            Open cfile For Input Shared As #1                                                           'jv060117
            Do Until EOF(1)                                                                             'jv060117
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16     'jv060117
                s = "SR1" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & Trim(f2) & Chr(9) & Trim(f3) & Chr(9)  'jv060117
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9) 'jv060117
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)         'jv060117
                s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16                                               'jv060117
                Grid1.AddItem s                                                                         'jv060117
            Loop                                                                                        'jv060117
            Close #1                                                                                    'jv060117
        Else                                                                                            'jv061215
            cfile = srpath & Right(Text1, 4) & "\sr1" & Format(Text1, "mmddyyyy") & ".txt"               'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "SR1" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16                                           'jv061215
                    Grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If                                                                                              'jv060117
        
    If Combo1 = "SR2" Or Combo1 = "All" Then                                                           'jv060117
        cfile = srpath & "SR2" & Format(Text1, "mmddyyyy") & ".txt"                                     'jv060117
        If Len(Dir(cfile)) > 0 Then                                                                     'jv060117
            Open cfile For Input Shared As #1                                                           'jv060117
            Do Until EOF(1)                                                                             'jv060117
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16     'jv060117
                s = "SR2" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & Trim(f2) & Chr(9) & Trim(f3) & Chr(9)  'jv060117
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9) 'jv060117
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)         'jv060117
                s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16                                               'jv060117
                Grid1.AddItem s                                                                         'jv060117
            Loop                                                                                        'jv060117
            Close #1                                                                                    'jv060117
        Else                                                                                            'jv061215
            cfile = srpath & Right(Text1, 4) & "\sr2" & Format(Text1, "mmddyyyy") & ".txt"               'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "SR2" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16                                           'jv061215
                    Grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If                                                                                              'jv060117
        
    If Combo1 = "SR3" Or Combo1 = "All" Then                                                           'jv060117
        cfile = srpath & "SR3" & Format(Text1, "mmddyyyy") & ".txt"                                     'jv060117
        If Len(Dir(cfile)) > 0 Then                                                                     'jv060117
            Open cfile For Input Shared As #1                                                           'jv060117
            Do Until EOF(1)                                                                             'jv060117
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16     'jv060117
                s = "SR3" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & Trim(f2) & Chr(9) & Trim(f3) & Chr(9)  'jv060117
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9) 'jv060117
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)         'jv060117
                s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16                                               'jv060117
                Grid1.AddItem s                                                                         'jv060117
            Loop                                                                                        'jv060117
            Close #1                                                                                    'jv060117
        Else                                                                                            'jv061215
            cfile = srpath & Right(Text1, 4) & "\sr3" & Format(Text1, "mmddyyyy") & ".txt"               'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "SR3" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16                                           'jv061215
                    Grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If                                                                                              'jv060117
    
    If Combo1 = "All" Then Call sorttime_Click
           
    hr = True
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            hr = Not hr
            If hr = True Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = hcolor.BackColor
            End If
        Next i
        Grid1.Row = 1
    End If
    
    's = "^Type|^RecId|<Area|<Description|<Source|<Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|<User|<Time|^ReqId"
    s = "^Type||^Area|<Description|^Source|^Target|<Product|^BarCode|||^Lot1|^Units|^Lot2|^Units|^Status|^User|<Time|^Plate"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 0       '600
    Grid1.ColWidth(2) = 900     '1300
    Grid1.ColWidth(3) = 1100
    Grid1.ColWidth(4) = 900     '1300
    Grid1.ColWidth(5) = 900     '1300
    Grid1.ColWidth(6) = 3600
    Grid1.ColWidth(7) = 1800
    Grid1.ColWidth(8) = 0       '600
    Grid1.ColWidth(9) = 0       '800
    Grid1.ColWidth(10) = 800
    Grid1.ColWidth(11) = 800
    Grid1.ColWidth(12) = 800
    Grid1.ColWidth(13) = 800
    Grid1.ColWidth(14) = 1000
    Grid1.ColWidth(15) = 1600
    Grid1.ColWidth(16) = 1400
    Grid1.ColWidth(17) = 1000
    Grid1.ColWidth(18) = 1
    hcolor.Caption = Combo1 & " Logs"
    cntlit.Caption = Grid1.Rows - 1 & " Records"
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub withdrawal()
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String, wsku As String, wlot As String
    sdate = Format(DateAdd("d", -30, Now), "yyyymmdd")
    sdate = InputBox("Start Date (YearMoDa):", "Start Date...", sdate)
    If Len(sdate) = 0 Then Exit Sub
    edate = InputBox("End Date (YearMoDa):", "End Date...", Format(Now, "yyyymmdd"))
    If Len(edate) = 0 Then Exit Sub
    wsku = InputBox("SKU:", "SKU...", "711")
    If Len(sdate) = 0 Then Exit Sub
    wlot = InputBox("Lot:", "W/D Lot...", "All")
    If Len(wlot) = 0 Then Exit Sub
    hcolor.Caption = "Withdrawal"
    Text1 = wsku & " " & wlot
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Cols = 12: Grid1.Rows = 1
    spath = Form1.srserv & "\wd\sr1\bin\sr1*.csv"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(Form1.srserv & "\wd\sr1\bin\" & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open Form1.srserv & "\wd\sr1\bin\" & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                If f2 = wsku Then
                    If wlot = "All" Or f3 = wlot Then
                        rc = rc + 1
                        s = rc & Chr(9)
                        s = s & f0 & Chr(9)
                        s = s & f1 & Chr(9)
                        s = s & f2 & Chr(9)
                        s = s & f3 & Chr(9)
                        f4 = Space(4 - Len(f4)) & f4
                        s = s & f4 & Chr(9)
                        s = s & f5 & Chr(9)
                        s = s & f6 & Chr(9)
                        s = s & f7 & Chr(9)
                        s = s & f8 & Chr(9)
                        s = s & Mid(sdir, 4, 2) & "-" & Mid(sdir, 6, 2) & " "
                        s = s & f9 & Chr(9)
                        s = s & Left(fdate, 4) & Mid(sdir, 4, 2) & "-" & Mid(sdir, 6, 2) & " "
                        s = s & Format(f9, "HH:MM") & Format(rc, "0000")
                        Grid1.AddItem s
                    End If
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    spath = Form1.srserv & "\wd\sr2\bin\sr2*.csv"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(Form1.srserv & "\wd\sr2\bin\" & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open Form1.srserv & "\wd\sr2\bin\" & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                If f2 = wsku Then
                    If wlot = "All" Or f3 = wlot Then
                        rc = rc + 1
                        s = rc & Chr(9)
                        s = s & f0 & Chr(9)
                        s = s & f1 & Chr(9)
                        s = s & f2 & Chr(9)
                        s = s & f3 & Chr(9)
                        f4 = Space(4 - Len(f4)) & f4
                        s = s & f4 & Chr(9)
                        s = s & f5 & Chr(9)
                        s = s & f6 & Chr(9)
                        s = s & f7 & Chr(9)
                        s = s & f8 & Chr(9)
                        s = s & Mid(sdir, 4, 2) & "-" & Mid(sdir, 6, 2) & " "
                        s = s & f9 & Chr(9)
                        s = s & Left(fdate, 4) & Mid(sdir, 4, 2) & "-" & Mid(sdir, 6, 2) & " "
                        s = s & Format(f9, "HH:MM") & Format(rc, "0000")
                        Grid1.AddItem s
                    End If
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    
    spath = Form1.srserv & "\wd\sr3\bin\sr3*.csv"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(Form1.srserv & "\wd\sr3\bin\" & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open Form1.srserv & "\wd\sr3\bin\" & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                If f2 = wsku Then
                    If wlot = "All" Or f3 = wlot Then
                        rc = rc + 1
                        s = rc & Chr(9)
                        s = s & f0 & Chr(9)
                        s = s & f1 & Chr(9)
                        s = s & f2 & Chr(9)
                        s = s & f3 & Chr(9)
                        f4 = Space(4 - Len(f4)) & f4
                        s = s & f4 & Chr(9)
                        s = s & f5 & Chr(9)
                        s = s & f6 & Chr(9)
                        s = s & f7 & Chr(9)
                        s = s & f8 & Chr(9)
                        s = s & Mid(sdir, 4, 2) & "-" & Mid(sdir, 6, 2) & " "
                        s = s & f9 & Chr(9)
                        s = s & Left(fdate, 4) & Mid(sdir, 4, 2) & "-" & Mid(sdir, 6, 2) & " "
                        s = s & Format(f9, "HH:MM") & Format(rc, "0000")
                        Grid1.AddItem s
                    End If
                End If
            Loop
            Close #1
        End If
        DoEvents
        sdir = Dir$
    Loop
    
    s = "^#|^Whs|^Tbl|^SKU|^Lot|^Plt#|<Product|^Func|^From|^To|^Time"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 600
    Grid1.ColWidth(2) = 500
    Grid1.ColWidth(3) = 600
    Grid1.ColWidth(4) = 600
    Grid1.ColWidth(5) = 500
    Grid1.ColWidth(6) = 3000
    Grid1.ColWidth(7) = 900
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 800
    Grid1.ColWidth(10) = 1600
    Grid1.ColWidth(11) = 1 '2000
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 5: Grid1.ColSel = 5
    Grid1.Sort = 5
    Screen.MousePointer = 0
End Sub

Private Sub ship_lots()
    Dim ds As ADODB.Recordset, s As String
    Dim rt As String, rf As String, rh As String
    Dim i As Integer, lflag As Boolean
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 8
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 7) = "Ship" Then
            s = Right(Grid1.TextMatrix(i, 1), 1) & Chr(9)
            s = s & Grid1.TextMatrix(i, 3) & Chr(9)
            s = s & Grid1.TextMatrix(i, 4) & Chr(9)
            s = s & Grid1.TextMatrix(i, 6) & Chr(9)
            s = s & Grid1.TextMatrix(i, 8) & Chr(9)
            s = s & Chr(9)
            s = s & Chr(9)
            s = s & Grid1.TextMatrix(i, 10)
            pgrid.AddItem s
        End If
    Next i
    If pgrid.Rows > 1 Then
        For i = 1 To pgrid.Rows - 1
            lflag = False
            s = "select lock_status from lane where whse_num = " & pgrid.TextMatrix(i, 0)
            s = s & " and vert_loc = " & Left(pgrid.TextMatrix(i, 4), 1)
            s = s & " and horz_loc = " & Mid(pgrid.TextMatrix(i, 4), 3, 2)
            s = s & " and rack_side = '" & Mid(pgrid.TextMatrix(i, 4), 6, 1) & "'"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                If ds!lock_status = 1 Then
                    pgrid.TextMatrix(i, 5) = "LIFO Bay"
                    lflag = True
                End If
            End If
            ds.Close
            If lflag = False Then
                s = "select * from lane where whse_num = " & pgrid.TextMatrix(i, 0)
                s = s & " and sku = '" & pgrid.TextMatrix(i, 1) & "'"
                s = s & " and lane_status < 'B'"
                s = s & " order by lot_num, qty, vert_loc, horz_loc"
                Set ds = Wdb.Execute(s)
                If ds.BOF = False Then
                    ds.MoveFirst
                    pgrid.TextMatrix(i, 5) = ds!lot_num
                    pgrid.TextMatrix(i, 6) = ds!vert_loc & " " & ds!horz_loc & " " & ds!rack_side & " - " & ds!qty
                End If
                ds.Close
            End If
        Next i
    End If
    pgrid.FillStyle = flexFillRepeat
    For i = 1 To pgrid.Rows - 1
        If pgrid.TextMatrix(i, 5) <> "LIFO Bay" And pgrid.TextMatrix(i, 5) > "0" Then
            If pgrid.TextMatrix(i, 5) < pgrid.TextMatrix(i, 2) Then
                pgrid.TextMatrix(i, 5) = "** " & pgrid.TextMatrix(i, 5)
                pgrid.Row = i: pgrid.RowSel = i
                pgrid.Col = 5: pgrid.ColSel = 6
                pgrid.CellBackColor = hcolor.BackColor
            End If
        End If
    Next i
    pgrid.FormatString = "^SR|^SKU|^Lot|<Product|^Position|^Old Lot|^Lane|^Time"
    pgrid.ColWidth(0) = 500
    pgrid.ColWidth(1) = 500
    pgrid.ColWidth(2) = 700
    pgrid.ColWidth(3) = 3000
    pgrid.ColWidth(4) = 1000
    pgrid.ColWidth(5) = 1000
    pgrid.ColWidth(6) = 1000
    pgrid.ColWidth(7) = 1000
        
    rt = "Shipped Lots"
    rh = Text1
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

Private Sub print_pgrid()
    Dim rt As String, rf As String, rh As String
    Dim i As Integer, k As Integer, j As Integer, s As String
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = Grid1.Cols - 1
    For i = 1 To Grid1.Rows - 1
        If Left(Grid1.TextMatrix(i, 11), 3) <> "999" Then
            s = Grid1.TextMatrix(i, 0)
            For k = 1 To Grid1.Cols - 2
                s = s & Chr(9) & Grid1.TextMatrix(i, k)
            Next k
            pgrid.AddItem s
        End If
    Next i
    pgrid.FormatString = Grid1.FormatString
    For i = 0 To 10
        pgrid.ColWidth(i) = Grid1.ColWidth(i)
    Next i
    
    rt = Me.Caption
    rh = hcolor.Caption & "  " & Text1
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

Private Sub refresh_grid()
    Dim cfile As String, f0 As String, f1 As String
    Dim f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, s As String
    Dim rc As Long
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    
    If Combo1 = "All" Then
        For i = 1 To 3
            'cfile = "\\bbc-01-wdmgmt\wd\SR" & i & "\bin\SR" & i & Format(Text1, "mmdd") & ".csv"
            cfile = Form1.srserv & "\wd\SR" & i & "\bin\SR" & i & Format(Text1, "mmdd") & ".csv"
            rc = 0
            If Len(Dir(cfile)) > 0 Then
                Open cfile For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    rc = rc + 1
                    s = rc & Chr(9)
                    s = s & f0 & Chr(9)
                    s = s & f1 & Chr(9)
                    s = s & f2 & Chr(9)
                    s = s & f3 & Chr(9)
                    f4 = Space(4 - Len(f4)) & f4
                    s = s & f4 & Chr(9)
                    s = s & f5 & Chr(9)
                    s = s & f6 & Chr(9)
                    s = s & f7 & Chr(9)
                    s = s & f8 & Chr(9)
                    s = s & f9 & Chr(9)
                    s = s & Format(f9, "HH:MM") & Format(rc, "0000")
                    Grid1.AddItem s
                Loop
                Close #1
            End If
        Next i
    Else
        rc = 0
        'cfile = "c:\" & Combo1 & Format(Text1, "mmdd") & ".csv"
        'cfile = "\\bbc-01-wdmgmt\wd\" & Combo1 & "\bin\" & Combo1 & Format(Text1, "mmdd") & ".csv"
        cfile = Form1.srserv & "\wd\" & Combo1 & "\bin\" & Combo1 & Format(Text1, "mmdd") & ".csv"
        'MsgBox cfile
    
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                rc = rc + 1
                s = rc & Chr(9)
                s = s & Trim(f0) & Chr(9)
                s = s & f1 & Chr(9)
                s = s & f2 & Chr(9)
                s = s & f3 & Chr(9)
                f4 = Space(4 - Len(f4)) & f4
                s = s & f4 & Chr(9)
                s = s & f5 & Chr(9)
                s = s & f6 & Chr(9)
                s = s & f7 & Chr(9)
                s = s & f8 & Chr(9)
                s = s & f9 & Chr(9)
                s = s & Format(f9, "HH:MM") & Format(rc, "0000")
                Grid1.AddItem s
            Loop
            Close #1
        End If
    End If
    If sortsku.Checked = True Then
        If Grid1.Rows > 2 Then
            Grid1.RowSel = Grid1.Row
            Grid1.Col = 3: Grid1.ColSel = 5
            Grid1.Sort = 5
        End If
    End If
    If sortrec.Checked = True Then
        If Grid1.Rows > 2 Then
            Grid1.RowSel = Grid1.Row
            Grid1.Col = 0: Grid1.ColSel = 0
            Grid1.Sort = 3
        End If
    End If
    If sorttime.Checked = True Then
        If Grid1.Rows > 2 Then
            Grid1.RowSel = Grid1.Row
            Grid1.Col = 11: Grid1.ColSel = 11
            Grid1.Sort = 5
        End If
    End If
    
    
    s = "^#|^Whs|^Tbl|^SKU|^Lot|^Plt#|<Product|^Func|^From|^To|^Time"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 600
    Grid1.ColWidth(2) = 500
    Grid1.ColWidth(3) = 600
    Grid1.ColWidth(4) = 600
    Grid1.ColWidth(5) = 500
    Grid1.ColWidth(6) = 3000
    Grid1.ColWidth(7) = 900
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 800
    Grid1.ColWidth(10) = 1000
    Grid1.ColWidth(11) = 1 '2000
    Grid1.FillStyle = flexFillRepeat
    hcolor.Caption = "All Records"
End Sub

Private Sub batonhand_Click()
    Dim wbc As String
    wbc = Grid1.TextMatrix(Grid1.Row, 7)
    wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) '& Mid(wbc, 18, 3)   'undo bc000
    tktonhand.bbarcode = wbc
    tktonhand.bproduct = Grid1.TextMatrix(Grid1.Row, 6)
    tktonhand.Show
End Sub

Private Sub ccol_Change()
    findcol.Caption = ccol.Caption
End Sub

Private Sub Combo1_Click()
    Dim ds As ADODB.Recordset, s As String, w As String
    'refresh_grid
    If Combo1 = "SR1" Then
        hcolor.BackColor = gcolor.BackColor
        hcolor.ForeColor = gcolor.ForeColor
        Grid1.BackColorFixed = gcolor.BackColor
        Grid1.ForeColor = gcolor.ForeColor
    End If
    If Combo1 = "SR2" Then
        hcolor.BackColor = ycolor.BackColor
        hcolor.ForeColor = ycolor.ForeColor
        Grid1.BackColorFixed = ycolor.BackColor
        Grid1.ForeColor = ycolor.ForeColor
    End If
    If Combo1 = "SR3" Then
        hcolor.BackColor = bcolor.BackColor
        hcolor.ForeColor = bcolor.ForeColor
        Grid1.BackColorFixed = bcolor.BackColor
        Grid1.ForeColor = bcolor.ForeColor
    End If
    If Combo1 = "SR5" Then
        hcolor.BackColor = pcolor.BackColor
        hcolor.ForeColor = pcolor.ForeColor
        Grid1.BackColorFixed = pcolor.BackColor
        Grid1.ForeColor = pcolor.ForeColor
    End If
    If Combo1 = "All" Then
        hcolor.BackColor = acolor.BackColor
        hcolor.ForeColor = acolor.ForeColor
        Grid1.BackColorFixed = acolor.BackColor
        Grid1.ForeColor = acolor.ForeColor
    End If
    refresh_srlogs                                              'jv062217
    If shiplanes.ListCount > 2 Then Exit Sub
    w = Right(Combo1, 1)
    oplanes.Clear: shiplanes.Clear
    s = "select * from opbays" ' where whse_num = " & w
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "SR-" & ds!whse_num
            s = s & ds!vert_loc & " "
            s = s & Format(ds!horz_loc, "00") & " "
            s = s & ds!rack_side
            oplanes.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select * from sr_config" ' where whs_num = " & w
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "SR-" & ds!whs_num
            s = s & ds!ship1_lane_vert & " "
            s = s & Format(ds!ship1_lane_horz, "00") & " "
            s = s & ds!ship1_lane_side
            shiplanes.AddItem s
            s = "SR-" & ds!whs_num
            s = s & ds!ship2_lane_vert & " "
            s = s & Format(ds!ship2_lane_horz, "00") & " "
            s = s & ds!ship2_lane_side
            shiplanes.AddItem s
            s = "SR-" & ds!whs_num
            s = s & ds!ship3_lane_vert & " "
            s = s & Format(ds!ship3_lane_horz, "00") & " "
            s = s & ds!ship3_lane_side
            shiplanes.AddItem s
            If ds!ship4_lane_vert > 0 Then
                s = "SR-" & ds!whs_num
                s = s & ds!ship4_lane_vert & " "
                s = s & Format(ds!ship4_lane_horz, "00") & " "
                s = s & ds!ship4_lane_side
                shiplanes.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    'refresh_grid
    refresh_srlogs                                  'jv062217
End Sub

Private Sub Command1_Click()
    'refresh_grid
    refresh_srlogs                                  'jv062217
End Sub

Private Sub findbay_Click()
    Dim i As Integer, s As String, j As String
    s = Left(Grid1.TextMatrix(Grid1.Row, 9), 6)
    s = InputBox("Bay:", "Highlight a bay..", s)
    If Len(s) = 0 Then Exit Sub
    hcolor.Caption = "Bay Activity: " & s
    s = UCase(s)
    j = Grid1.Row
    Grid1.Redraw = False
    For i = 1 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 11) = "9999"
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        Grid1.CellBackColor = Grid1.BackColor
        If Left(Grid1.TextMatrix(i, 8), 6) = s Then
            Grid1.CellBackColor = hcolor.BackColor
            Grid1.TextMatrix(i, 11) = Grid1.TextMatrix(i, 0)
            j = i
        'Else
        '    Grid1.CellBackColor = Grid1.BackColor
        End If
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        If Left(Grid1.TextMatrix(i, 9), 6) = s Then
            Grid1.CellBackColor = hcolor.BackColor
            j = i
            Grid1.TextMatrix(i, 11) = Grid1.TextMatrix(i, 0)
        'Else
        '    Grid1.CellBackColor = Grid1.BackColor
        End If
    Next i
    Grid1.Redraw = True
    'Grid1.TopRow = j
    Grid1.TopRow = 1
    Grid1.Col = 11: Grid1.ColSel = 11
    Grid1.Sort = 3
End Sub

Private Sub findcol_Click()
    Dim i As Integer, s As String, t As String, k As Integer, sc As Integer
    sc = Grid1.Col
    k = 0
    s = Grid1.Text
    s = InputBox(ccol & ": ", "Highlight " & ccol & "...", s)
    If Len(s) = 0 Then Exit Sub
    hcolor.Caption = ccol & ": " & s
    Grid1.Redraw = False
    For i = 1 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 18) = "99999999999"
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        If UCase(Grid1.TextMatrix(i, sc)) = UCase(s) Then
            Grid1.TextMatrix(i, 18) = Grid1.TextMatrix(i, 7)
            Grid1.CellBackColor = hcolor.BackColor
            k = k + 1
        Else
            Grid1.CellBackColor = Grid1.BackColor
        End If
    Next i
    Grid1.Redraw = True
    Grid1.TopRow = 1
    Grid1.Row = 1: Grid1.RowSel = 1
    Grid1.Col = 18: Grid1.ColSel = 18
    Grid1.Sort = 5
    cntlit.Caption = k & " Records"
    Grid1.Col = sc
End Sub

Private Sub finddrops_Click()
    Dim i As Integer, j As Integer
    If shiplanes.ListCount < 2 Then Exit Sub
    hcolor.Caption = "Pallet Drops"
    Grid1.Redraw = False
    For i = 1 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 11) = "9999"
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        Grid1.CellBackColor = Grid1.BackColor
        If Grid1.TextMatrix(i, 7) <> "Ship" Then
            For j = 0 To shiplanes.ListCount - 1
                If Grid1.TextMatrix(i, 1) = Left(shiplanes.List(j), 4) Then
                    If Left(Grid1.TextMatrix(i, 9), 6) = Right(shiplanes.List(j), 6) Then
                        Grid1.TextMatrix(i, 11) = Grid1.TextMatrix(i, 0)
                        Grid1.CellBackColor = hcolor.BackColor
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i
    Grid1.Redraw = True
    Grid1.TopRow = 1
    Grid1.Col = 11: Grid1.ColSel = 11
    Grid1.Sort = 3
End Sub

Private Sub findnew_Click()
    Dim i As Integer, j As Integer, k As Integer
    If oplanes.ListCount < 2 Then Exit Sub
    hcolor.Caption = "New Bays"
    j = 0: k = 0
    Grid1.Redraw = False
    For i = 1 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 18) = "99999999999"
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        Grid1.CellBackColor = Grid1.BackColor
        If Right(Grid1.TextMatrix(i, 5), 1) = "4" Then
            Grid1.TextMatrix(i, 18) = Grid1.TextMatrix(i, 5)
            Grid1.CellBackColor = bcolor.BackColor
            j = j + 1
        End If
        If Right(Grid1.TextMatrix(i, 4), 1) = "4" Then
            Grid1.TextMatrix(i, 18) = Grid1.TextMatrix(i, 5)
            Grid1.CellBackColor = gcolor.BackColor
            k = k + 1
        End If
    Next i
    Grid1.Redraw = True
    Grid1.TopRow = 1
    Grid1.Col = 18: Grid1.ColSel = 18
    Grid1.Sort = 3
    cntlit.Caption = j & " - " & k & " = " & Format(j - k, "0") & " Records"
End Sub

Private Sub findop_Click()
    Dim i As Integer, j As Integer
    If oplanes.ListCount < 2 Then Exit Sub
    hcolor.Caption = "Order Pick"
    Grid1.Redraw = False
    For i = 1 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 18) = "99999999999"
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        Grid1.CellBackColor = Grid1.BackColor
        For j = 0 To oplanes.ListCount - 1
            If Left(Grid1.TextMatrix(i, 2), 4) = Left(oplanes.List(j), 4) Then
                If Left(Grid1.TextMatrix(i, 5), 6) = Right(oplanes.List(j), 6) Then
                    Grid1.TextMatrix(i, 18) = Grid1.TextMatrix(i, 5)
                    Grid1.CellBackColor = hcolor.BackColor
                    Exit For
                End If
            End If
        Next j
    Next i
    Grid1.Redraw = True
    Grid1.TopRow = 1
    Grid1.Col = 18: Grid1.ColSel = 18
    Grid1.Sort = 3
End Sub

Private Sub findsku_Click()
    Dim i As Integer, s As String, t As String, k As Integer, n As Integer
    k = 0
    's = Left(Grid1.TextMatrix(Grid1.Row, 7), 3)
    s = Trim(Left(Grid1.TextMatrix(Grid1.Row, 7), 4))       'jv062916
    s = InputBox("SKU:", "Highlight SKU..", s)
    If Len(s) = 0 Then Exit Sub
    n = Len(s)                                              'jv062916
    hcolor.Caption = "SKU: " & s
    Grid1.Redraw = False
    For i = 1 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 18) = "99999999999"
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        'If Left(Grid1.TextMatrix(i, 7), 3) = s Or Left(Grid1.TextMatrix(i, 6), 3) = s Then
        If Left(Grid1.TextMatrix(i, 7), n) = s Or Left(Grid1.TextMatrix(i, 6), n) = s Then          'jv062916
            Grid1.TextMatrix(i, 18) = Grid1.TextMatrix(i, 7)
            Grid1.CellBackColor = hcolor.BackColor
            k = k + 1
        Else
            Grid1.CellBackColor = Grid1.BackColor
        End If
        'If Left(Grid1.TextMatrix(i, 7), 3) <> Left(Grid1.TextMatrix(i, 6), 3) And Grid1.TextMatrix(i, 7) > "100" Then
        If Left(Grid1.TextMatrix(i, 7), n) <> Left(Grid1.TextMatrix(i, 6), n) And Grid1.TextMatrix(i, 7) > "100" Then       'jv062916
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 6: Grid1.ColSel = 7
            Grid1.CellBackColor = cntlit.BackColor
            Grid1.TextMatrix(i, 18) = Grid1.TextMatrix(i, 7)
        End If
    Next i
    Grid1.Redraw = True
    Grid1.TopRow = 1
    Grid1.Row = 1: Grid1.RowSel = 1
    Grid1.Col = 18: Grid1.ColSel = 18
    Grid1.Sort = 5
    cntlit.Caption = k & " Records"


    'Dim i As Integer, s As String, t As String
    ''s = Grid1.TextMatrix(Grid1.Row, 3)
    's = Trim(Left(Grid1.TextMatrix(Grid1.Row, 7), 4))
    's = InputBox("SKU:", "Highlight SKU..", s)
    'If Len(s) = 0 Then Exit Sub
    'hcolor.Caption = "SKU: " & s
    'Grid1.Redraw = False
    'For i = 1 To Grid1.Rows - 1
    '    Grid1.TextMatrix(i, 11) = "99999999999"
    '    Grid1.Row = i: Grid1.RowSel = i
    '    Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
    '    If Grid1.TextMatrix(i, 3) = s Then
    '        Grid1.CellBackColor = hcolor.BackColor
    '        t = Grid1.TextMatrix(i, 3)
    '        t = t & Format(Val(Grid1.TextMatrix(i, 4)), "00000")
    '        t = t & Format(Val(Grid1.TextMatrix(i, 5)), "000")
    '        Grid1.TextMatrix(i, 11) = t
    '    Else
    '        Grid1.CellBackColor = Grid1.BackColor
    '    End If
    'Next i
    'Grid1.Redraw = True
    ''Grid1.TopRow = j
    'Grid1.TopRow = 1
    'Grid1.Row = 1: Grid1.RowSel = 1
    'Grid1.Col = 11: Grid1.ColSel = 11
    'Grid1.Sort = 5
    
End Sub

Private Sub Form_Load()
    'Text1 = Format(Now, "mm-dd")
    Text1 = Format(Now, "MM-dd-yyyy")
    Combo1.AddItem "SR1"
    Combo1.AddItem "SR2"
    Combo1.AddItem "SR3"
    Combo1.AddItem "SR5"
    Combo1.AddItem "All"
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    pgrid.Width = Me.Width - 100
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1400
End Sub

Private Sub Grid1_Click()
    ccol = Grid1.TextMatrix(0, Grid1.Col)
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu findmenu
End Sub

Private Sub mockwith_Click()
    Call withdrawal
End Sub

Private Sub palhist_Click()
    Dim wbc As String
    wbc = Grid1.TextMatrix(Grid1.Row, 7)
    wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
    palhistory.Show
    palhistory.barkey = wbc
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rf As String, rh As String
    If hcolor.Caption <> "All Records" Then
        Call print_pgrid
        Exit Sub
    End If
    rt = Me.Caption
    rh = Combo1 & "  " & Text1
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
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

Private Sub shiplots_Click()
    'Call ship_lots
    finddrops_Click
    DoEvents
    sortsku_Click
    DoEvents
    prtmenu_Click
End Sub

Private Sub sortrec_Click()
    sortrec.Checked = True
    sortsku.Checked = False
    sorttime.Checked = False
    If Grid1.Rows > 2 Then
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 0: Grid1.ColSel = 0
        Grid1.Sort = 3
    End If
End Sub

Private Sub sortsku_Click()
    sortrec.Checked = False
    sortsku.Checked = True
    sorttime.Checked = False
    If Grid1.Rows > 2 Then
        Grid1.RowSel = Grid1.Row
        'Grid1.Col = 3: Grid1.ColSel = 5
        Grid1.Col = 7: Grid1.ColSel = 7
        Grid1.Sort = 5
    End If
End Sub

Private Sub sorttime_Click()
    sortrec.Checked = False
    sortsku.Checked = False
    sorttime.Checked = True
    If Grid1.Rows > 2 Then
        Grid1.RowSel = Grid1.Row
        'If hcolor.Caption = "Withdrawal" Then
        '    Grid1.Col = 10: Grid1.ColSel = 10
        'Else
        '    Grid1.Col = 11: Grid1.ColSel = 11
        'End If
        Grid1.Col = 16: Grid1.ColSel = 16
        Grid1.Sort = 5
    End If
End Sub
