VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form traffmoves 
   Caption         =   "Pallet Movement"
   ClientHeight    =   8160
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form15"
   ScaleHeight     =   8160
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "View All Fields"
      Height          =   255
      Left            =   10440
      TabIndex        =   8
      Top             =   240
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   240
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   3495
      Left            =   0
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6165
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   6135
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10821
      _Version        =   327680
      BackColorSel    =   32768
      WordWrap        =   -1  'True
      FocusRect       =   0
      FillStyle       =   1
      AllowUserResizing=   3
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3600
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label w5c 
      BackColor       =   &H00800080&
      Caption         =   "SR4 Racks"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   19
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label w4c 
      BackColor       =   &H000080FF&
      Caption         =   "TRI-LEVEL 4"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10680
      TabIndex        =   18
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label w3c 
      BackColor       =   &H0000C000&
      Caption         =   "TRI-LEVEL 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10680
      TabIndex        =   17
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label w2c 
      BackColor       =   &H00FF8080&
      Caption         =   "TRI-LEVEL 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   16
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label w1c 
      BackColor       =   &H008080FF&
      Caption         =   "TRI-LEVEL 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10800
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label sr5c 
      BackColor       =   &H00FFC0FF&
      Caption         =   "SR5"
      Height          =   255
      Left            =   8280
      TabIndex        =   14
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label sr4c 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SR4"
      Height          =   255
      Left            =   8280
      TabIndex        =   13
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label sr3c 
      BackColor       =   &H0080FFFF&
      Caption         =   "SR3"
      Height          =   255
      Left            =   8280
      TabIndex        =   12
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label sr2c 
      BackColor       =   &H00FFFF80&
      Caption         =   "SR2"
      Height          =   255
      Left            =   8400
      TabIndex        =   11
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label sr1c 
      BackColor       =   &H0000FF00&
      Caption         =   "SR1"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8280
      TabIndex        =   10
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label ccol 
      Caption         =   "..."
      Height          =   255
      Left            =   12240
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label cntlit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Date:"
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
      Top             =   240
      Width           =   615
   End
   Begin VB.Label hcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Pallet Moves:"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu sortmenu 
      Caption         =   "Sort"
      Begin VB.Menu sortbc 
         Caption         =   "BarCode"
      End
      Begin VB.Menu sortdt 
         Caption         =   "Date/Time"
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
   End
   Begin VB.Menu usermenu 
      Caption         =   "User"
      Visible         =   0   'False
      Begin VB.Menu emplook 
         Caption         =   "Lookup Employee Name"
      End
   End
   Begin VB.Menu widrpt 
      Caption         =   "Withdrawl"
      Visible         =   0   'False
   End
   Begin VB.Menu addrec 
      Caption         =   "Add Record"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "traffmoves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub withdrawal()
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String, wsku As String, wlot As String
    Dim cfile As String, s As String, bc As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim dl As Long, wbc As String, pname5 As String
    Dim logpath As String
    'If Form1.plantno = 50 Then logpath = "\\bbc-01-wdmgmt\wd\pallogs\"
    'If Form1.plantno = 51 Then logpath = "\\bbc-01-wdmgmt\wd\testlogs\"
    'If Form1.plantno = 51 Then logpath = "\\bbba-02-dc\f\user\waredist\data\pallogs\"
    'If Form1.plantno = 52 Then logpath = "\\bbc-01-wdmgmt\wd\testlogs\"
    'If Form1.plantno = 52 Then logpath = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
    logpath = "\\bbc-01-prodtrk\wd\pallogs\"
    s = grid1.TextMatrix(grid1.Row, 7)
    sdate = Format(Val(Mid(s, 9, 2)) - 2, "00")
    sdate = "20" & sdate & Mid(s, 5, 4)
    
    'sdate = Format(DateAdd("d", -30, Now), "yyyymmdd")
    sdate = InputBox("Start Date (YearMoDa):", "Start Date...", sdate)
    If Len(sdate) = 0 Then Exit Sub
    edate = InputBox("End Date (YearMoDa):", "End Date...", Format(Now, "yyyymmdd"))
    If Len(edate) = 0 Then Exit Sub
    'wsku = Left(Grid1.TextMatrix(Grid1.Row, 7), 3)
    'wsku = InputBox("SKU:", "SKU...", wsku)
    'If Len(wsku) = 0 Then Exit Sub
    'wlot = Grid1.TextMatrix(Grid1.Row, 10)
    'wlot = InputBox("Lot:", "W/D Lot...", wlot)
    'If Len(wlot) = 0 Then Exit Sub
    
    wbc = grid1.TextMatrix(grid1.Row, 7)
    wbc = InputBox("Enter a BarCode for the withdrawal:", "BarCode Example....", wbc)
    If Len(wbc) = 0 Then Exit Sub
    wsku = Trim(Left(wbc, 4))
    wlot = barcode_to_lotnum(wbc)
    If wlot = "01001" Then
        MsgBox "Invalid BarCode example.", vbExclamation + vbOKOnly, "problem with barcode..."
        Exit Sub
    End If
    'MsgBox wsku & " " & wlot
    
    hcolor.Caption = "Withdrawal"
    'Text1 = wsku & " " & wlot
    Screen.MousePointer = 11
    grid1.Clear: grid1.Cols = 19: grid1.Rows = 1
    
    bc = " "
    spath = logpath & "recv*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                'If Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All") Then
                If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All")) Then
                    grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                    pname5 = f5         'SR5
                End If
                
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    
    spath = logpath & "tml*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "TM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                'If Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All") Then
                If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All")) Then
                    grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    
    spath = logpath & "move*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "M" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                'If Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All") Then
                If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All")) Then
                    grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    
    spath = logpath & "sr4rem*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "M" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                'If Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All") Then
                If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All")) Then
                    grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    
    spath = logpath & "ship*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                'If Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All") Then
                '    Grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                'End If
                'If Trim(Left(f6, 4)) = Trim(Left(bc, 4)) And f9 = "..." Then
                'If Left(f6, 10) = Left(bc, 10) And f9 = "..." Then
                If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All")) Then
                    grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                End If
                
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    
    spath = logpath & "wms*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "WM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                'If Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All") Then
                If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All")) Then
                    grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    
    
    spath = logpath & "pick*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "P" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                'If Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All") Then
                If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or f11 = wlot Or wlot = "All")) Then
                    grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    
    If Form1.plantno = 50 Then
        spath = "\\bbc-01-prodtrk\wd\sr1\bin\sr1*.csv"
        'spath = Form1.srserv & "\wd\sr1\bin\sr1*.csv"
        sdir = Dir$(spath)
        Do While sdir <> ""
            'fdate = Format(FileDateTime("\\bbc-01-wdmgmt\wd\sr1\bin\" & sdir), "yyyymmdd")
            fdate = Format(FileDateTime(Form1.srserv & "\wd\sr1\bin\" & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                'Open "\\bbc-01-wdmgmt\wd\sr1\bin\" & sdir For Input Shared As #1
                Open Form1.srserv & "\wd\sr1\bin\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    If f2 = wsku Then
                        If wlot = "All" Or f3 = wlot Then
                            rc = rc + 1
                            s = "SR" & Chr(9)
                            s = s & rc & Chr(9)
                            s = s & f0 & Chr(9)
                            s = s & f6 & Chr(9)
                            s = s & f7 & Chr(9)
                            s = s & f8 & Chr(9)
                            s = s & f2 & " " & f5 & Chr(9)
                            If Len(bc) >= 3 Then
                                s = s & Left(bc, 13) & Format(Val(f4), "000") & Chr(9)
                            Else
                                s = s & f4 & Chr(9)
                            End If
                            s = s & "1" & Chr(9)
                            s = s & "Pallet" & Chr(9)
                            s = s & f3 & Chr(9)
                            s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9)
                            s = s & f0 & Chr(9)
                            's = s & Left(fdate, 4) & Mid(sdir, 4, 2) & "-" & Mid(sdir, 6, 2) & " "
                            s = s & Mid(fdate, 3, 2) & Mid(sdir, 4, 2) & Mid(sdir, 6, 2) & " "
                            's = s & Format(f9, "HH:MM") & Format(rc, "0000")
                            s = s & Format(f9, "HH:MM") & ":99"
                            grid1.AddItem s
                        End If
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    
        'spath = "\\bbc-01-wdmgmt\wd\sr2\bin\sr2*.csv"""
        spath = Form1.srserv & "\wd\sr2\bin\sr2*.csv"""
        sdir = Dir$(spath)
        Do While sdir <> ""
            fdate = Format(FileDateTime(Form1.srserv & "\wd\sr2\bin\" & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                'Open "\\bbc-01-wdmgmt\wd\sr2\bin\" & sdir For Input Shared As #1
                Open Form1.srserv & "\wd\sr2\bin\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    If f2 = wsku Then
                        If wlot = "All" Or f3 = wlot Then
                            rc = rc + 1
                            s = "SR" & Chr(9)
                            s = s & rc & Chr(9)
                            s = s & f0 & Chr(9)
                            s = s & f6 & Chr(9)
                            s = s & f7 & Chr(9)
                            s = s & f8 & Chr(9)
                            s = s & f2 & " " & f5 & Chr(9)
                            If Len(bc) >= 3 Then
                                s = s & Left(bc, 13) & Format(Val(f4), "000") & Chr(9)
                            Else
                                s = s & f4 & Chr(9)
                            End If
                            s = s & "1" & Chr(9)
                            s = s & "Pallet" & Chr(9)
                            s = s & f3 & Chr(9)
                            s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9)
                            s = s & f0 & Chr(9)
                            's = s & Left(fdate, 4) & Mid(sdir, 4, 2) & "-" & Mid(sdir, 6, 2) & " "
                            s = s & Mid(fdate, 3, 2) & Mid(sdir, 4, 2) & Mid(sdir, 6, 2) & " "
                            's = s & Format(f9, "HH:MM") & Format(rc, "0000")
                            s = s & Format(f9, "HH:MM") & ":99"
                            grid1.AddItem s
                        End If
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    
        'spath = "\\bbc-01-wdmgmt\wd\sr3\bin\sr3*.csv"""
        spath = Form1.srserv & "\wd\sr3\bin\sr3*.csv"""
        sdir = Dir$(spath)
        Do While sdir <> ""
            'fdate = Format(FileDateTime("\\bbc-01-wdmgmt\wd\sr3\bin\" & sdir), "yyyymmdd")
            fdate = Format(FileDateTime(Form1.srserv & "\wd\sr3\bin\" & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                'Open "\\bbc-01-wdmgmt\wd\sr3\bin\" & sdir For Input Shared As #1
                Open Form1.srserv & "\wd\sr3\bin\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    If f2 = wsku Then
                        If wlot = "All" Or f3 = wlot Then
                            rc = rc + 1
                            s = "SR" & Chr(9)
                            s = s & rc & Chr(9)
                            s = s & f0 & Chr(9)
                            s = s & f6 & Chr(9)
                            s = s & f7 & Chr(9)
                            s = s & f8 & Chr(9)
                            s = s & f2 & " " & f5 & Chr(9)
                            If Len(bc) >= 3 Then
                                s = s & Left(bc, 13) & Format(Val(f4), "000") & Chr(9)
                            Else
                                s = s & f4 & Chr(9)
                            End If
                            s = s & "1" & Chr(9)
                            s = s & "Pallet" & Chr(9)
                            s = s & f3 & Chr(9)
                            s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9)
                            s = s & f0 & Chr(9)
                            's = s & Left(fdate, 4) & Mid(sdir, 4, 2) & "-" & Mid(sdir, 6, 2) & " "
                            s = s & Mid(fdate, 3, 2) & Mid(sdir, 4, 2) & Mid(sdir, 6, 2) & " "
                            's = s & Format(f9, "HH:MM") & Format(rc, "0000")
                            s = s & Format(f9, "HH:MM") & ":99"
                            grid1.AddItem s
                        End If
                    End If
                Loop
                Close #1
            End If
            DoEvents
            sdir = Dir$
        Loop
        
        'spath = "\\bbc-01-wdmgmt\wd\sr3\bin\sr3*.csv"""
        spath = Form1.srserv & "\wd\sr5\bin\sr5*.csv"""
        sdir = Dir$(spath)
        Do While sdir <> ""
            'fdate = Format(FileDateTime("\\bbc-01-wdmgmt\wd\sr3\bin\" & sdir), "yyyymmdd")
            fdate = Format(FileDateTime(Form1.srserv & "\wd\sr5\bin\" & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                'Open "\\bbc-01-wdmgmt\wd\sr3\bin\" & sdir For Input Shared As #1
                Open Form1.srserv & "\wd\sr5\bin\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    If Trim(f2) = wsku Then
                        If wlot = "All" Or Trim(f3) = wlot Then
                            rc = rc + 1
                            s = "SR" & Chr(9)
                            s = s & rc & Chr(9)
                            s = s & f0 & Chr(9)
                            s = s & f6 & Chr(9)
                            s = s & f7 & Chr(9)
                            s = s & f8 & Chr(9)
                            's = s & f2 & " " & f5 & Chr(9)
                            s = s & pname5 & Chr(9)     'SR5
                            If Len(bc) >= 3 Then
                                s = s & Left(bc, 13) & Format(Val(f4), "000") & Chr(9)
                            Else
                                s = s & f4 & Chr(9)
                            End If
                            s = s & "1" & Chr(9)
                            s = s & "Pallet" & Chr(9)
                            s = s & f3 & Chr(9)
                            s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9)
                            s = s & f0 & Chr(9)
                            's = s & Left(fdate, 4) & Mid(sdir, 4, 2) & "-" & Mid(sdir, 6, 2) & " "
                            s = s & Mid(fdate, 3, 2) & Mid(sdir, 4, 2) & Mid(sdir, 6, 2) & " "
                            's = s & Format(f9, "HH:MM") & Format(rc, "0000")
                            s = s & Format(f9, "HH:MM") & ":99"
                            grid1.AddItem s
                        End If
                    End If
                Loop
                Close #1
            End If
            DoEvents
            sdir = Dir$
        Loop
        
        
    End If
    
    If grid1.Rows > 1 Then
        For i = 1 To grid1.Rows - 1
            grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7) & grid1.TextMatrix(i, 16)
        Next i
    End If
    
    If Check1.Value = 1 Then
        s = "^Type|^RecId|^Area|<Description|^Source|^Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId"
        grid1.FormatString = s
        grid1.ColWidth(0) = 600
        grid1.ColWidth(1) = 1 '600
        grid1.ColWidth(2) = 1300
        grid1.ColWidth(3) = 1000
        grid1.ColWidth(4) = 1300
        grid1.ColWidth(5) = 1300
        grid1.ColWidth(6) = 3000
        grid1.ColWidth(7) = 1800
        grid1.ColWidth(8) = 600
        grid1.ColWidth(9) = 800
        grid1.ColWidth(10) = 800
        grid1.ColWidth(11) = 800
        grid1.ColWidth(12) = 800
        grid1.ColWidth(13) = 800
        grid1.ColWidth(14) = 1 '800
        grid1.ColWidth(15) = 1000
        grid1.ColWidth(16) = 1400
        grid1.ColWidth(17) = 1 '1000
        grid1.ColWidth(18) = 1 '2100
    Else
        s = "^Type|^|^|<Description|^Source|^Target|<Product|^Pallet|^|^|^LotNum|^Units|^LotNum|^Units|^|^|<Time|^"
        grid1.FormatString = s
        grid1.ColWidth(0) = 600
        grid1.ColWidth(1) = 1 '600
        grid1.ColWidth(2) = 1 '300
        grid1.ColWidth(3) = 1000
        grid1.ColWidth(4) = 1300
        grid1.ColWidth(5) = 1300
        grid1.ColWidth(6) = 3000
        grid1.ColWidth(7) = 1800
        grid1.ColWidth(8) = 1 '600
        grid1.ColWidth(9) = 1 '800
        grid1.ColWidth(10) = 800
        grid1.ColWidth(11) = 800
        grid1.ColWidth(12) = 800
        grid1.ColWidth(13) = 800
        grid1.ColWidth(14) = 1 '800
        grid1.ColWidth(15) = 1 '000
        grid1.ColWidth(16) = 1400
        grid1.ColWidth(17) = 1 '1000
        grid1.ColWidth(18) = 1 '2100
    End If
    
    grid1.RowSel = grid1.Row
    'Grid1.col = 7: Grid1.ColSel = 7
    grid1.Col = 18: grid1.ColSel = 18
    grid1.Sort = 5
    cntlit.Caption = grid1.Rows - 1 & " Records"
    Screen.MousePointer = 0
End Sub

Private Sub refresh_grid1()
    Dim cfile As String, s As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim z As Long, c As Boolean
    Dim logpath As String
    
    'If Form1.plantno = 50 Then logpath = "\\bbc-01-wdmgmt\wd\pallogs\"
    'If Form1.plantno = 51 Then logpath = "\\bbc-01-wdmgmt\wd\testlogs\"
    'If Form1.plantno = 51 Then logpath = "\\bbba-02-dc\f\user\waredist\data\pallogs\"
    'If Form1.plantno = 52 Then logpath = "\\bbc-01-wdmgmt\wd\testlogs\"
    'If Form1.plantno = 52 Then logpath = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
    'logpath = "\\bbc-01-wdmgmt\wd\testlogs\"    'jvtcar
    logpath = logdir
    'MsgBox logpath
    'logpath = "\\bbc-01-prodtrk\wd\pallogs\"
    grid1.Redraw = False
    grid1.Clear: grid1.Rows = 1: grid1.Cols = 19
    If Combo1 = "Shipping" Then
        addrec.Enabled = True
    Else
        addrec.Enabled = False
    End If
    
    If Combo1 = "Production" Or Combo1 = "All" Then
        cfile = logpath & "recv" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                If Left(f1, 9) = "TRI-LEVEL" Or Left(f1, 9) = "TRI LEVEL" Then grid1.AddItem s
            Loop
            Close #1
        End If
    End If
    
    If Combo1 = "Shipping" Or Combo1 = "All" Then
        cfile = logpath & "ship" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                grid1.AddItem s
            Loop
            Close #1
        End If
    End If
    
    If Combo1 = "Rack Moves" Or Combo1 = "All" Then
        cfile = logpath & "move" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "M" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                grid1.AddItem s
            Loop
            Close #1
        End If
    End If
    
    If Combo1 = "Picks" Or Combo1 = "All" Then
        cfile = logpath & "pick" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "P" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                grid1.AddItem s
            Loop
            Close #1
        End If
    End If
    
    If Combo1 = "Traffic Master" Or Combo1 = "All" Then
        cfile = logpath & "tml" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "TM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                grid1.AddItem s
            Loop
            Close #1
        End If
    End If
    
    If Combo1 = "WMS" Or Combo1 = "All" Or Combo1 = "Rack Activity" Then
        cfile = logpath & "wms" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "WM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                If Left(f6, 3) <> "ING" Then
                    grid1.AddItem s
                End If
            Loop
            Close #1
        End If
    End If
    
    If Combo1 = "Rack Activity" Or Combo1 = "All" Then
        'cfile = logpath & "sr4rem" & Format(DateAdd("d", -1, Text1), "mmddyyyy") & ".txt"
        ''MsgBox cfile
        'If Len(Dir(cfile)) > 0 Then
        '    Open cfile For Input Shared As #1
        '    Do Until EOF(1)
        '        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
        '        s = "RR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & Trim(f2) & Chr(9) & Trim(f3) & Chr(9)
        '        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
        '        s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
        '        s = s & f14 & Chr(9) & f15 & Chr(9) & f16
        '        Grid1.AddItem s
        '    Loop
        '    Close #1
        'End If
        cfile = logpath & "sr4rem" & Format(Text1, "mmddyyyy") & ".txt"
        'MsgBox cfile
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "RR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & Trim(f2) & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                grid1.AddItem s
            Loop
            Close #1
        End If
        'Check1.Value = 1
        'If Grid1.Rows > 1 Then
        '    Grid1.FillStyle = flexFillRepeat
        '    For i = 1 To Grid1.Rows - 1
        '        If Grid1.TextMatrix(i, 7) <> Left(Grid1.TextMatrix(i, 3), 16) Then
        '            Grid1.Row = i: Grid1.RowSel = i
        '            Grid1.Col = 7: Grid1.ColSel = 7
        '            Grid1.CellBackColor = hcolor.BackColor
        '        End If
        '        If Grid1.TextMatrix(i, 4) <> Right(Grid1.TextMatrix(i, 3), Len(Grid1.TextMatrix(i, 4))) Then
        '            Grid1.Row = i: Grid1.RowSel = i
        '            Grid1.Col = 4: Grid1.ColSel = 4
        '            Grid1.CellBackColor = hcolor.BackColor
        '        End If
        '    Next i
        'End If
        'Grid1.Row = 1: Grid1.Col = 2
    End If
        
    
    If Check1.Value = 1 Then
        s = "^Type|^RecId|^Area|<Description|^Source|^Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId"
        grid1.FormatString = s
        grid1.ColWidth(0) = 600
        grid1.ColWidth(1) = 600
        grid1.ColWidth(2) = 1300
        grid1.ColWidth(3) = 1000
        grid1.ColWidth(4) = 1300
        grid1.ColWidth(5) = 1300
        grid1.ColWidth(6) = 3000
        grid1.ColWidth(7) = 1800
        grid1.ColWidth(8) = 600
        grid1.ColWidth(9) = 800
        grid1.ColWidth(10) = 800
        grid1.ColWidth(11) = 800
        grid1.ColWidth(12) = 800
        grid1.ColWidth(13) = 800
        grid1.ColWidth(14) = 800
        grid1.ColWidth(15) = 1000
        grid1.ColWidth(16) = 1400
        grid1.ColWidth(17) = 1000
        grid1.ColWidth(18) = 1
    Else
        s = "^Type|^RecId|^Area|<Description|^Source|^Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId"
        grid1.FormatString = s
        grid1.ColWidth(0) = 600
        grid1.ColWidth(1) = 1 '600
        grid1.ColWidth(2) = 1 '1300
        grid1.ColWidth(3) = 1 '1000
        grid1.ColWidth(4) = 1300
        grid1.ColWidth(5) = 1300
        grid1.ColWidth(6) = 3000
        grid1.ColWidth(7) = 1800
        grid1.ColWidth(8) = 1 '600
        grid1.ColWidth(9) = 1 '800
        grid1.ColWidth(10) = 800
        grid1.ColWidth(11) = 800
        grid1.ColWidth(12) = 800
        grid1.ColWidth(13) = 800
        grid1.ColWidth(14) = 1 '800
        grid1.ColWidth(15) = 1 '1000
        grid1.ColWidth(16) = 1400
        grid1.ColWidth(17) = 1 '1000
        grid1.ColWidth(18) = 1
    End If
    hcolor.Caption = "All Records"
    cntlit.Caption = grid1.Rows - 1 & " Records"
    If grid1.Rows > 1 Then
        grid1.FillStyle = flexFillRepeat
        c = True
        For i = 1 To grid1.Rows - 1
            c = Not c
            If c = True Then
                grid1.Row = i: grid1.RowSel = i
                grid1.Col = 1: grid1.ColSel = grid1.Cols - 1
                grid1.CellBackColor = cntlit.BackColor
            End If
        Next i
        
        If Combo1 = "Traffic Master" Then
            grid1.FillStyle = flexFillRepeat
            For i = 1 To grid1.Rows - 1
                If grid1.TextMatrix(i, 4) = "TRI-LEVEL 1" Then
                    grid1.Row = i: grid1.RowSel = i
                    grid1.Col = 4: grid1.ColSel = 4
                    grid1.CellBackColor = Me.w1c.BackColor
                    grid1.CellForeColor = Me.w1c.ForeColor
                End If
                If grid1.TextMatrix(i, 4) = "TRI-LEVEL 2" Then
                    grid1.Row = i: grid1.RowSel = i
                    grid1.Col = 4: grid1.ColSel = 4
                    grid1.CellBackColor = Me.w2c.BackColor
                    grid1.CellForeColor = Me.w2c.ForeColor
                End If
                If grid1.TextMatrix(i, 4) = "TRI-LEVEL 3" Then
                    grid1.Row = i: grid1.RowSel = i
                    grid1.Col = 4: grid1.ColSel = 4
                    grid1.CellBackColor = Me.w3c.BackColor
                    grid1.CellForeColor = Me.w3c.ForeColor
                End If
                If grid1.TextMatrix(i, 4) = "TRI-LEVEL 4" Then
                    grid1.Row = i: grid1.RowSel = i
                    grid1.Col = 4: grid1.ColSel = 4
                    grid1.CellBackColor = Me.w4c.BackColor
                    grid1.CellForeColor = Me.w4c.ForeColor
                End If
                If grid1.TextMatrix(i, 4) = "TRI-LEVEL 5" Then
                    grid1.Row = i: grid1.RowSel = i
                    grid1.Col = 4: grid1.ColSel = 4
                    grid1.CellBackColor = Me.w5c.BackColor
                    grid1.CellForeColor = Me.w5c.ForeColor
                End If
                If grid1.TextMatrix(i, 5) = "SR1" Then
                    grid1.Row = i: grid1.RowSel = i
                    grid1.Col = 5: grid1.ColSel = 7
                    grid1.CellBackColor = Me.sr1c.BackColor
                End If
                If grid1.TextMatrix(i, 5) = "SR2" Then
                    grid1.Row = i: grid1.RowSel = i
                    grid1.Col = 5: grid1.ColSel = 7
                    grid1.CellBackColor = Me.sr2c.BackColor
                End If
                If grid1.TextMatrix(i, 5) = "SR3" Then
                    grid1.Row = i: grid1.RowSel = i
                    grid1.Col = 5: grid1.ColSel = 7
                    grid1.CellBackColor = Me.sr3c.BackColor
                End If
                If grid1.TextMatrix(i, 5) = "SR4" Then
                    grid1.Row = i: grid1.RowSel = i
                    grid1.Col = 5: grid1.ColSel = 7
                    grid1.CellBackColor = Me.sr4c.BackColor
                End If
                If grid1.TextMatrix(i, 5) = "SR5" Then
                    grid1.Row = i: grid1.RowSel = i
                    grid1.Col = 5: grid1.ColSel = 7
                    grid1.CellBackColor = Me.sr5c.BackColor
                End If
                    
            Next i
        End If
        grid1.Row = 1
    End If
    grid1.Redraw = True
End Sub

Private Sub addrec_Click()
    Dim mgroup As String, msource As String, mtarget As String, msku As String
    Dim mlot As String, mqty As String, mlot2 As String, mqty2 As String, mbc As String
    Dim i As Integer, s As String, cfile As String
    Dim logpath As String
    'If Form1.plantno = 50 Then logpath = "\\bbc-01-wdmgmt\wd\pallogs\"
    'If Form1.plantno = 51 Then logpath = "\\bbc-01-wdmgmt\wd\testlogs\"
    'If Form1.plantno = 51 Then logpath = "\\bbba-02-dc\f\user\waredist\data\pallogs\"
    'If Form1.plantno = 52 Then logpath = "\\bbc-01-wdmgmt\wd\testlogs\"
    'If Form1.plantno = 52 Then logpath = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
    logpath = logdir
    If grid1.Row = 0 Then Exit Sub
    i = grid1.Row
    mgroup = grid1.TextMatrix(i, 3)
    msource = grid1.TextMatrix(i, 4)
    mtarget = grid1.TextMatrix(i, 5)
    'msku = Trim(Left(Grid1.TextMatrix(i, 6), 4))
    msku = grid1.TextMatrix(i, 6)
    mbc = Left(grid1.TextMatrix(i, 7), 12)
    mlot = grid1.TextMatrix(i, 10)
    mqty = grid1.TextMatrix(i, 11)
    mlot2 = grid1.TextMatrix(i, 12)
    mqty2 = grid1.TextMatrix(i, 13)
    mgroup = InputBox("Shipping Group:", "Shipping Group...", mgroup)
    If Len(mgroup) = 0 Then Exit Sub
    msource = InputBox("Source:", "Source...", msource)
    If Len(msource) = 0 Then Exit Sub
    mtarget = InputBox("Target:", "Target...", mtarget)
    If Len(mtarget) = 0 Then Exit Sub
    msku = InputBox("Product:", "Product...", msku)
    If Len(msku) = 0 Then Exit Sub
    mlot = InputBox("Lot Number 1:", "Lot Number 1...", mlot)
    If Len(mlot) = 0 Then Exit Sub
    mqty = InputBox("Units 1:", "Units 1...", mqty)
    If Len(mqty) = 0 Then Exit Sub
    mlot2 = InputBox("Lot Number 2:", "Lot Number 2...", mlot2)
    If Len(mlot2) = 0 Then Exit Sub
    mqty2 = InputBox("Units 2:", "Units 2...", mqty2)
    If Len(mqty2) = 0 Then Exit Sub
    mbc = UCase(InputBox("BarCode:", "BarCode...", mbc))
    If Len(mbc) = 0 Then Exit Sub
    If Len(mbc) < 16 Then
        MsgBox "Invalid BarCode length: " & mbc, vbOKOnly + vbInformation, "Try again..."
        Exit Sub
    End If
    If Left(mbc, 4) <> Left(msku, 4) Then
        MsgBox "BarCode: " & mbc & " and " & msku & " do not match.", vbOKOnly + vbInformation, "Try again..."
        Exit Sub
    End If
    s = "S" & Chr(9)
    s = s & "0" & Chr(9)
    s = s & "DOCK" & Chr(9)
    s = s & mgroup & Chr(9)
    s = s & msource & Chr(9)
    s = s & mtarget & Chr(9)
    s = s & msku & Chr(9)
    s = s & mbc & Chr(9)
    s = s & "1" & Chr(9)
    s = s & "Pallet" & Chr(9)
    s = s & mlot & Chr(9)
    s = s & mqty & Chr(9)
    s = s & mlot2 & Chr(9)
    s = s & mqty2 & Chr(9)
    s = s & "COMP" & Chr(9)
    s = s & "WMS" & Chr(9)
    s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
    grid1.AddItem s, i
    cfile = logpath & "ship" & Format(Text1, "mmddyyyy") & ".txt"
    'cfile = "c:\jvtest.txt"
    'MsgBox cfile
    Open cfile For Append Shared As #1
    Write #1, "0";
    Write #1, "DOCK";
    Write #1, mgroup;
    Write #1, msource;
    Write #1, mtarget;
    Write #1, msku;
    Write #1, mbc;
    Write #1, "1";
    Write #1, "Pallet";
    Write #1, mlot;
    Write #1, mqty;
    Write #1, mlot2;
    Write #1, mqty2;
    Write #1, "COMP";
    Write #1, "WMS";
    Write #1, Format(Now, "yyMMdd hh:mm:ss");
    Write #1, " "
    Close #1
End Sub

Private Sub ccol_Change()
    findcol.Caption = ccol.Caption
End Sub

Private Sub Check1_Click()
    refresh_grid1
End Sub

Private Sub Combo1_Click()
    refresh_grid1
End Sub

Private Sub emplook_Click()
    Dim db As dao.Database, ds As dao.Recordset, s As String
    If Len(grid1.Text) = 0 Then Exit Sub
    Set db = OpenDatabase("s:\wd\data\wdemp.mdb")
    s = "select bb_num, first_name, last_name, nickname from employees"
    s = s & " where bb_num = '" & grid1.Text & "'"
    Set ds = db.OpenRecordset(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If ds(3) > "0" Then
            s = ds(1) & " '" & ds(3) & "' " & ds(2)
        Else
            s = ds(1) & " " & ds(2)
        End If
    Else
        s = "Employee #: " & grid1.Text & " is not in WdEmp database."
    End If
    ds.Close: db.Close
    MsgBox s, vbOKOnly + vbInformation, "Employee " & grid1.Text & " ...."
End Sub

Private Sub findcol_Click()
    Dim i As Integer, s As String, t As String, K As Integer, sc As Integer, c As Boolean
    sc = grid1.Col
    K = 0
    s = grid1.Text
    s = InputBox(ccol & ": ", "Highlight " & ccol & "...", s)
    If Len(s) = 0 Then Exit Sub
    hcolor.Caption = ccol & ": " & s
    grid1.Redraw = False: grid1.FillStyle = flexFillRepeat
    c = True
    For i = 1 To grid1.Rows - 1
        grid1.TextMatrix(i, 18) = "99999999999"
        grid1.Row = i: grid1.RowSel = i
        grid1.Col = 1: grid1.ColSel = grid1.Cols - 1
        If UCase(grid1.TextMatrix(i, sc)) = UCase(s) Then
            grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7)
            grid1.CellBackColor = hcolor.BackColor
            K = K + 1
        Else
            c = Not c
            If c = True Then
                grid1.Row = i: grid1.RowSel = i
                grid1.Col = 1: grid1.ColSel = grid1.Cols - 1
                grid1.CellBackColor = cntlit.BackColor
            Else
                grid1.CellBackColor = grid1.BackColor
            End If
        End If
    Next i
    grid1.Redraw = True
    'Grid1.TopRow = j
    grid1.TopRow = 1
    grid1.Row = 1: grid1.RowSel = 1
    grid1.Col = 18: grid1.ColSel = 18
    grid1.Sort = 5
    cntlit.Caption = K & " Records"
    grid1.Col = sc
End Sub

Private Sub findsku_Click()
    Dim i As Integer, s As String, t As String, K As Integer, c As Boolean
    K = 0
    s = Left(grid1.TextMatrix(grid1.Row, 7), 3)
    s = InputBox("SKU:", "Highlight SKU..", s)
    If Len(s) = 0 Then Exit Sub
    hcolor.Caption = "SKU: " & s
    grid1.Redraw = False: c = True
    For i = 1 To grid1.Rows - 1
        grid1.TextMatrix(i, 18) = "99999999999"
        grid1.Row = i: grid1.RowSel = i
        grid1.Col = 1: grid1.ColSel = grid1.Cols - 1
        If Left(grid1.TextMatrix(i, 7), 3) = s Or Left(grid1.TextMatrix(i, 6), 3) = s Then
            grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7)
            grid1.CellBackColor = hcolor.BackColor
            K = K + 1
        Else
            c = Not c
            If c = True Then
                grid1.CellBackColor = cntlit.BackColor
            Else
                grid1.CellBackColor = grid1.BackColor
            End If
        End If
        If Left(grid1.TextMatrix(i, 7), 3) <> Left(grid1.TextMatrix(i, 6), 3) And grid1.TextMatrix(i, 7) > "100" Then
            grid1.Row = i: grid1.RowSel = i
            grid1.Col = 6: grid1.ColSel = 7
            grid1.CellBackColor = cntlit.BackColor
            grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7)
        End If
    Next i
    grid1.Redraw = True
    'Grid1.TopRow = j
    grid1.TopRow = 1
    grid1.Row = 1: grid1.RowSel = 1
    grid1.Col = 18: grid1.ColSel = 18
    grid1.Sort = 5
    cntlit.Caption = K & " Records"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Text1 = Format(Now, "mm-dd-yyyy")
    Combo1.Clear
    Combo1.AddItem "Production"
    'Combo1.AddItem "Shipping"
    'Combo1.AddItem "Rack Moves"
    'Combo1.AddItem "Picks"
    Combo1.AddItem "Traffic Master"
    'Combo1.AddItem "WMS"
    'Combo1.AddItem "All"
    Combo1.ListIndex = 0
    'If Form1.plantno = 50 Then
        emplook.Enabled = True
    'Else
    '    emplook.Enabled = False
    'End If
End Sub

Private Sub Form_Resize()
    grid1.Width = Me.Width - 80
    pgrid.Width = Me.Width - 80
    If Me.Height > 2000 Then grid1.Height = Me.Height - 1500
End Sub

Private Sub grid1_Click()
    ccol = grid1.TextMatrix(0, grid1.Col)
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If grid1.TextMatrix(0, grid1.Col) = "User" And emplook.Enabled = True Then
            PopupMenu usermenu
        Else
            PopupMenu findmenu
        End If
    End If
End Sub

Private Sub sortbc_Click()
    grid1.Row = 1: grid1.RowSel = 1
    grid1.Col = 7: grid1.ColSel = 7
    grid1.Sort = 5
End Sub

Private Sub sortdt_Click()
    grid1.Row = 1: grid1.RowSel = 1
    grid1.Col = 16: grid1.ColSel = 16
    grid1.Sort = 5
End Sub

Private Sub widrpt_Click()
    Call withdrawal
End Sub
