VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Edittrl 
   Caption         =   "Edit Trailers"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   8265
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   255
      Left            =   7200
      TabIndex        =   32
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Post to Crane"
      Height          =   375
      Left            =   7320
      TabIndex        =   30
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   1455
      Left            =   3600
      TabIndex        =   27
      Top             =   5760
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2566
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1575
      Left            =   240
      TabIndex        =   26
      Top             =   5640
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2778
      _Version        =   327680
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Blank Bill"
      Height          =   375
      Left            =   8040
      TabIndex        =   25
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Rack Check Off"
      Height          =   375
      Left            =   6240
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.ComboBox sd 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add Product"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel Product"
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Bill"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Product "
      Height          =   3135
      Left            =   6960
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   720
         TabIndex        =   29
         Text            =   "Text5"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Text            =   "Text4"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Text            =   "Text3"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Source"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Units"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Wraps"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Pallets"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.ListBox tid 
      Height          =   645
      Left            =   9240
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox wc 
      Height          =   645
      Left            =   9240
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox pc 
      Height          =   645
      Left            =   9240
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid td 
      Height          =   5655
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9975
      _Version        =   327680
      FixedCols       =   0
      BackColor       =   -2147483638
      BackColorSel    =   128
      BackColorBkg    =   -2147483638
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   9240
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bill of Laden Exists"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6960
      TabIndex        =   33
      Top             =   6840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label gcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "___"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   31
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label plantno 
      Caption         =   "50"
      Height          =   255
      Left            =   7680
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   495
   End
   Begin VB.Label ano 
      Caption         =   "Label7"
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
      Left            =   8280
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label bno 
      Caption         =   "Label7"
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
      Left            =   6960
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Edittrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim srow As Integer
Private Sub post_r12_bill()
    Dim cfile As String, ofile As String, s As String
    Dim f1 As String, f2 As String, f3 As String, f4 As String, f5 As String
    Dim f6 As String, f7 As String, f8 As String, f9 As String, f10 As String
    Dim f11 As String, f12 As String, f13 As String, f14 As String, f15 As String
    Dim f16 As String, f17 As String, i As Integer
    Dim mplant As String, ldate As String, sdate As String, mbatch As String
    Dim pbranch As String, ctest As String, mfile As String, mprod As String
    Dim torg As String, twhs As String, tacct As String, psku As String, plot As String
    Dim ds As adodb.Recordset, sqls As String, tid As Long
    Dim ldate2 As String, cfile2 As String
    On Error GoTo vberror
    mplant = Form1.plantno
    ldate = InputBox("Load Date:", "Trailer Loaded Date....", Format(DateAdd("d", -1, sd), "mm-dd-yyyy"))    'jv022811
    If Len(ldate) = 0 Then Exit Sub         'jv022811
    If Format(ldate, "yyyyMMdd") > Format(Now, "yyyyMMdd") Then
        s = "The load date entered, " & ldate & ", cannot be greater than the current date."
        MsgBox s, vbOKOnly + vbInformation, "sorry, this date cannot be used..."
        Exit Sub                    'jv010313
    End If
    ldate2 = Format(DateAdd("d", 1, ldate), "MMddyyyy")
    ldate = Format(ldate, "MMddyyyy")
    
    sdate = Format(sd, "MMddyyyy")
    cfile = Form1.pallogs & "ship" & ldate & ".txt"
    cfile2 = Form1.pallogs & "ship" & ldate2 & ".txt"
    mfile = Form1.pallogs & "move" & sdate & ".txt"
    ofile = Form1.pallogs & "bill" & sdate & ".txt"
    If Len(Dir(cfile)) = 0 Then Exit Sub
    Open ofile For Append As #1
    s = "select id,runid,trailers.branch,account,trlno,sku,pallets,wraps,units,branchname,shipdate"
    s = s & " from trailers, branches"
    s = s & " where runid = " & Left(List1, Len(List1) - 6)
    s = s & " and branches.branch = trailers.branch"
    s = s & " and plant = " & mplant
    s = s & " and pb_flag = 'N'"
    s = s & " order by runid, sku"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If Left(ds!trlno, 1) = "#" Then i = Val(Right(ds!trlno, 1))
        If Left(ds!trlno, 1) = "B" Then i = 10 - Val(Right(ds!trlno, 1))
        If Val(Right(ds!trlno, 1)) = 0 Then
            mbatch = DateDiff("d", "1-1-2012", sd) & Format(ds(1), "00") & Right(ds!trlno, 1)
        Else
            mbatch = DateDiff("d", "1-1-2012", sd) & Format(ds(1), "00") & i
        End If
        Do Until ds.EOF
            mprod = ds!sku
            tid = 0                     'jv1112
            For i = 0 To td.Rows - 1
                If td.TextMatrix(i, 0) = ds!sku Then
                    mprod = mprod & " " & td.TextMatrix(i, 1)
                    Exit For
                End If
            Next i
            If ds!branch = 15 Or ds!branch = 16 Then
                If ds!pallets > 0 Then
                    mcomm = UCase(ds!branchname) & " " & ds!trlno
                    If mcomm & ds!sku <> ctest Then
                        If Len(Dir(cfile)) > 0 Then
                            Open cfile For Input As #2
                            Do Until EOF(2)
                                Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
                                'If LCase(f3) = LCase(gcode) And f2 = "DOCK" And f5 = mcomm And Left(f6, 3) = ds!sku Then
                                'If LCase(f3) = LCase(gcode) And f2 = "DOCK" And Left(f6, 3) = ds!sku Then
                                If LCase(f3) = LCase(gcode) And f2 = "DOCK" And Trim(Left(f6, 4)) = ds!sku Then 'jv082415
                                    Write #1, ds!account;   'Recid
                                    Write #1, f2;           'Area
                                    Write #1, f3;           'Description
                                    Write #1, f4;           'Source
                                    Write #1, Combo1; 'f5;  'Target
                                    Write #1, f6;           'Product
                                    Write #1, f7;           'Pallet
                                    Write #1, f8;           'Qty
                                    Write #1, f9;           'Uom
                                    Write #1, f10;          'lot
                                    Write #1, f11;          'units
                                    Write #1, f12;          'lot2
                                    Write #1, f13;          'units2
                                    Write #1, "PEND";       'status
                                    Write #1, f15;          'user
                                    Write #1, f16;          'time
                                    Write #1, ds!runid      'reqid
                                    tid = ds!id             'jv1112
                                End If
                            Loop
                            Close #2
                        End If
                        
                        If Len(Dir(cfile2)) > 0 Then
                            Open cfile2 For Input As #2     'After Midnight
                            Do Until EOF(2)
                                Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
                                'If LCase(f3) = LCase(gcode) And f2 = "DOCK" And f5 = mcomm And Left(f6, 3) = ds!sku Then
                                'If LCase(f3) = LCase(gcode) And f2 = "DOCK" And Left(f6, 3) = ds!sku Then
                                If LCase(f3) = LCase(gcode) And f2 = "DOCK" And Trim(Left(f6, 4)) = ds!sku Then 'jv082415
                                    Write #1, ds!account;   'Recid
                                    Write #1, f2;       'Area
                                    Write #1, f3;       'Description
                                    Write #1, f4;       'Source
                                    Write #1, Combo1;   'f5;       'Target
                                    Write #1, f6;       'Product
                                    Write #1, f7;       'Pallet
                                    Write #1, f8;       'Qty
                                    Write #1, f9;       'Uom
                                    Write #1, f10;      'lot
                                    Write #1, f11;      'units
                                    Write #1, f12;      'lot2
                                    Write #1, f13;      'units2
                                    Write #1, "PEND";   'status
                                    Write #1, f15;      'user
                                    Write #1, f16;      'time
                                    Write #1, ds!runid  'reqid
                                    tid = ds!id         'jv1112
                                End If
                            Loop
                            Close #2
                        End If
                    End If
                    ctest = mcomm & ds!sku
                Else
                    Write #1, ds!account;                       'Recid
                    Write #1, "JOBBING";                        'Area
                    Write #1, gcode.Caption;                    'Description
                    Write #1, "ORDER PICK";                     'Source
                    Write #1, Combo1;                           'Target
                    Write #1, mprod;                            'Product
                    Write #1, " ";                              'Palletid
                    Write #1, ds!wraps;                         'qty
                    Write #1, "Wraps";                          'uom
                    Write #1, "LOT1";                           'lot
                    Write #1, ds!units;                         'units
                    Write #1, " ", "0";                         'lot2 & qty2
                    Write #1, "PEND";                           'status
                    Write #1, "wms";                            'user
                    Write #1, Format(Now, "yyMMdd hh:mm:ss");   'time
                    Write #1, ds!runid                          'reqid
                    tid = ds!id                                 'jv1112
                End If
            Else
                If ds!pallets > 0 Then
                    mcomm = UCase(ds!branchname) & " " & ds!trlno
                    If mcomm & ds!sku <> ctest Then
                        If Len(Dir(cfile)) > 0 Then
                            Open cfile For Input As #2
                            Do Until EOF(2)
                                Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
                                'If LCase(f3) = LCase(gcode) And f2 = "DOCK" And f5 = mcomm And Left(f6, 3) = ds!sku Then
                                If LCase(f3) = LCase(gcode) And f2 = "DOCK" And f5 = mcomm And Trim(Left(f6, 4)) = ds!sku Then  'jv082415
                                    Write #1, mbatch;   'Recid
                                    Write #1, f2;       'Area
                                    Write #1, f3;       'Description
                                    Write #1, f4;       'Source
                                    Write #1, f5;       'Target
                                    Write #1, f6;       'Product
                                    Write #1, f7;       'Pallet
                                    Write #1, f8;       'Qty
                                    Write #1, f9;       'Uom
                                    Write #1, f10;      'lot
                                    Write #1, f11;      'units
                                    Write #1, f12;      'lot2
                                    Write #1, f13;      'units2
                                    Write #1, "PEND";      'status
                                    Write #1, f15;      'user
                                    Write #1, f16;      'time
                                    Write #1, ds!runid  'reqid
                                    tid = ds!id         'jv1112
                                End If
                            Loop
                            Close #2
                        End If
                        
                        If Len(Dir(cfile2)) > 0 Then
                            Open cfile2 For Input As #2     'After Midnight
                            Do Until EOF(2)
                                Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
                                'If LCase(f3) = LCase(gcode) And f2 = "DOCK" And f5 = mcomm And Left(f6, 3) = ds!sku Then
                                If LCase(f3) = LCase(gcode) And f2 = "DOCK" And f5 = mcomm And Trim(Left(f6, 4)) = ds!sku Then  'jv082415
                                    Write #1, mbatch;   'Recid
                                    Write #1, f2;       'Area
                                    Write #1, f3;       'Description
                                    Write #1, f4;       'Source
                                    Write #1, f5;       'Target
                                    Write #1, f6;       'Product
                                    Write #1, f7;       'Pallet
                                    Write #1, f8;       'Qty
                                    Write #1, f9;       'Uom
                                    Write #1, f10;      'lot
                                    Write #1, f11;      'units
                                    Write #1, f12;      'lot2
                                    Write #1, f13;      'units2
                                    Write #1, "PEND";   'status
                                    Write #1, f15;      'user
                                    Write #1, f16;      'time
                                    Write #1, ds!runid  'reqid
                                    tid = ds!id         'jv1112
                                End If
                            Loop
                            Close #2
                        End If
                    End If
                    ctest = mcomm & ds!sku
                Else
                    Write #1, mbatch;               'Recid
                    Write #1, "PARTIAL";            'Area
                    Write #1, gcode.Caption;        'Description
                    Write #1, "ORDER PICK";         'Source
                    Write #1, Combo1;               'Target
                    Write #1, mprod;                'Product
                    Write #1, " ";                  'Palletid
                    Write #1, ds!wraps;             'qty
                    Write #1, "Wraps";              'uom
                    Write #1, "LOT1";               'lot
                    Write #1, ds!units;             'units
                    Write #1, " ", "0";             'lot2 & qty2
                    Write #1, "PEND";               'status
                    Write #1, "wms";                'user
                    Write #1, Format(Now, "yyMMdd hh:mm:ss");
                    Write #1, ds!runid              'reqid
                    tid = ds!id                     'jv1112
                End If
            End If
            'turn off for testing
            If tid > 0 Then
                sqls = "update trailers set pb_flag = 'Y'"
                'sqls = sqls & " where runid = " & ds!runid
                sqls = sqls & " where id = " & tid          'jv1112
                Sdb.Execute sqls
            End If
            ds.MoveNext
        Loop
    End If
    Close #1
    ds.Close
    refresh_grid
    Call EdBills.refresh_grid1(sd)
    EdBills.Show
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "post_r12_bill", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " post_r12_bill - Error Number: " & eno
        End
    End If
End Sub

Private Sub ckoffsheet2011()
    Dim prun As String, ss As adodb.Recordset
    Dim sqlx As String, i As Integer, pdesc As String
    Dim ds2 As adodb.Recordset
    Dim tpals As Integer, twrps As Integer
    On Error GoTo vberror
    Screen.MousePointer = 11
    Printer.FontName = "Courier New"
    Printer.FontSize = 12: Printer.FontBold = True
    Printer.Print " "
    Printer.Print " "
    Printer.Print " "
    Printer.Print "Check Off Sheet   "; Combo1; "         Date: "; sd; "     Order #: "; gcode
    Printer.Print " "
    Printer.Print "SKU                                            Pallets      Rack        Wraps"
    Printer.FontUnderline = True
    
    tpals = 0: twrps = 0
    For i = 0 To td.Rows - 1
        If Me.plantno = "52" Then
            If Val(td.TextMatrix(i, 2)) > 0 Then
                If td.TextMatrix(i, 5) = "Crane" Then
                    sqlx = td.TextMatrix(i, 0) & " " & td.TextMatrix(i, 1)
                    sqlx = sqlx & Space(50 - Len(sqlx))
                    sqlx = sqlx & Format(Val(td.TextMatrix(i, 2)), "0")
                    sqlx = sqlx & Space(60 - Len(sqlx))
                    sqlx = sqlx & "Crane"
                    Printer.FontBold = Not Printer.FontBold
                    Printer.Print sqlx
                Else
                    sqlx = td.TextMatrix(i, 0) & " " & td.TextMatrix(i, 1)
                    sqlx = sqlx & Space(50 - Len(sqlx))
                    sqlx = sqlx & Format(Val(td.TextMatrix(i, 2)), "0")
                    sqlx = sqlx & Space(60 - Len(sqlx))
                    sqlx = sqlx & "Racks"
                    Printer.FontBold = Not Printer.FontBold
                    Printer.Print sqlx
                End If
            End If
        Else
            If Val(td.TextMatrix(i, 2)) > 0 Then
                sqlx = td.TextMatrix(i, 0) & " " & td.TextMatrix(i, 1)
                sqlx = sqlx & Space(50 - Len(sqlx))
                sqlx = sqlx & Format(Val(td.TextMatrix(i, 2)), "0")
                sqlx = sqlx & Space(60 - Len(sqlx))
                'sqlx = sqlx & "Racks"
                sqlx = sqlx & "     "
                Printer.FontBold = Not Printer.FontBold
                Printer.Print sqlx
            End If
        End If
        If Val(td.TextMatrix(i, 3)) > 0 Then
            sqlx = td.TextMatrix(i, 0) & " " & td.TextMatrix(i, 1)
            sqlx = sqlx & Space(73 - Len(sqlx))
            sqlx = sqlx & Format(Val(td.TextMatrix(i, 3)), "0")
            Printer.FontBold = Not Printer.FontBold
            Printer.Print sqlx
        End If
        tpals = tpals + Val(td.TextMatrix(i, 2))
        twrps = twrps + Val(td.TextMatrix(i, 3))
    Next i
    Printer.FontUnderline = False
    Printer.FontBold = True
    Printer.Print " "
    sqlx = "Totals" & Space(44) & Format(tpals, "0") & " Pallets" & Space(11) & Format(twrps, "0") & " Wrps"
    Printer.Print sqlx
    Printer.FontBold = False
    Printer.Print "+---------------------------------------------------------------------+"
    Printer.Print "|   SEAL # ___________________    TRAILER # _____________             |"
    Printer.Print "|                                                                     |"
    Printer.Print "|   LOADER ___________________    DRIVER ________________             |"
    Printer.Print "+---------------------------------------------------------------------+"
    Printer.Print "Alternates:"
    sqlx = "select * from brorders where orddate = '" & sd & "'"
    sqlx = sqlx & " and plant = " & plantno
    sqlx = sqlx & " and branch = " & bno
    sqlx = sqlx & " and altflag = 'Y'"
    sqlx = sqlx & " order by sku"
    Set ds2 = Sdb.Execute(sqlx)
    If ds2.BOF = False Then
        ds2.MoveFirst
        Do Until ds2.EOF
            sqlx = "select fgunit, fgdesc from skumast where sku = '" & ds2!sku & "'"
            Set ss = Sdb.Execute(sqlx)
            If ss.BOF = False Then
                pdesc = Trim(ss!fgunit) & " " & Trim(ss!fgdesc)
            Else
                pdesc = "SKU not on file......"
            End If
            ss.Close
            Printer.Print ds2!sku; " "; pdesc
            ds2.MoveNext
        Loop
    Else
        Printer.Print "None specified...."
    End If
    Printer.EndDoc
    ds2.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "ckoffsheet2011", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " ckoffsheet2011 - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String, i As Integer
    On Error GoTo vberror
    List1.ListIndex = Combo1.ListIndex
    td.Rows = 1: pc.Clear: wc.Clear: tid.Clear: Label9.Visible = False
    If Len(List1) = 0 Then Exit Sub
    sqlx = "Select ID,trailers.sku,fgunit,fgdesc,pallets,wraps,units,pallet,"
    sqlx = sqlx & "numwrap,branch,account,plant,trailers.whs_num,groupcode,pb_flag"
    sqlx = sqlx & " from trailers,skumast"
    sqlx = sqlx & " Where runid = " & Left$(List1, Len(List1) - 6)
    sqlx = sqlx & " And trailers.sku = skumast.sku"
    sqlx = sqlx & " Order by trailers.sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        ds.Close
        td.AddItem "*": wc.AddItem "0": pc.AddItem "0": tid.AddItem "0"
        td.Row = 1: Call td_Click
        Exit Sub
    End If
    ds.MoveFirst
    bno = ds!branch: ano = ds!account: plantno = ds!plant
    td.FillStyle = flexFillRepeat
    td.Redraw = False
    Do Until ds.EOF
        sqlx = ds!sku & Chr$(9)
        sqlx = sqlx & " " & ds!fgunit & " " & ds!fgdesc & Chr$(9)
        If ds!pallets > 0 Then sqlx = sqlx & ds!pallets
        sqlx = sqlx & Chr$(9)
        If ds!wraps > 0 Then sqlx = sqlx & ds!wraps
        sqlx = sqlx & Chr$(9)
        If ds!units > 0 Then sqlx = sqlx & ds!units
        sqlx = sqlx & Chr$(9)
        If ds(12) < 4 Then
            sqlx = sqlx & "Crane"
        Else
            sqlx = sqlx & "Rack"
        End If
        td.AddItem sqlx
        If ds!pb_flag = "Y" Then
            td.Row = td.Rows - 1: td.RowSel = td.Row
            td.Col = 0: td.ColSel = td.Cols - 1
            td.CellForeColor = td.BackColorSel
            Label9.Visible = True
        End If
        wc.AddItem ds!numwrap
        pc.AddItem ds!pallet
        tid.AddItem ds!id
        gcode = ds!groupcode
        ds.MoveNext
    Loop
    ds.Close
    td.Redraw = True
    td_Click
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid - Error Number: " & eno
        End
    End If
End Sub

Private Sub sywhs(w As String)
    Dim sqlx As String
    On Error GoTo vberror
    If Val(tid) < 1 Then Exit Sub
    sqlx = "Update trailers set whs_num = "
    If w = "Crane" Then
        sqlx = sqlx & "1"
    Else
        sqlx = sqlx & "15"
    End If
    sqlx = sqlx & " Where ID = " & Val(tid)
    Sdb.Execute sqlx
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "sywhs", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " sywhs - Error Number: " & eno
        End
    End If
End Sub

Private Sub update_trl()
    Dim sqlx As String
    On Error GoTo vberror
    If Val(tid) < 1 Then Exit Sub
    sqlx = "Update Trailers Set Pallets = " & Val(Text2)
    sqlx = sqlx & ", Wraps = " & Val(Text3)
    sqlx = sqlx & ", Units = " & Val(Text4)
    sqlx = sqlx & " Where ID = " & Val(tid)
    Sdb.Execute sqlx
    td.TextMatrix(td.Row, 2) = Format(Val(Text2), "#####")
    td.TextMatrix(td.Row, 3) = Format(Val(Text3), "#####")
    td.TextMatrix(td.Row, 4) = Format(Val(Text4), "#####")
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "update_trl", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " update_trl - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo1_Click()
    Call refresh_grid
    If td.Rows > 1 Then
        td.Row = 1: Call td_Click
    End If
End Sub

Private Sub Command1_Click()        'Print Bill
    Call post_r12_bill
End Sub

Private Sub Command1_GotFocus()
    Command1.FontBold = True
End Sub

Private Sub Command1_LostFocus()
    Command1.FontBold = False
End Sub

Private Sub Command2_Click()                'Cancel Product
    Dim sqlx As String
    On Error GoTo vberror
    If MsgBox("Cancel" & td.TextMatrix(td.Row, 1) & " From Trailer", vbYesNo + vbQuestion, "Are you sure?") = vbYes Then
        sqlx = "Delete From Trailers Where ID = " & tid
        Sdb.Execute sqlx
        If td.Rows > 2 Then
            td.RemoveItem td.Row
            tid.RemoveItem tid.ListIndex
            wc.RemoveItem wc.ListIndex
            pc.RemoveItem pc.ListIndex
        Else
            Call refresh_grid
        End If
        Call td_Click
    End If
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

Private Sub Command2_GotFocus()
    Command2.FontBold = True
End Sub

Private Sub Command2_LostFocus()
    Command2.FontBold = False
End Sub

Private Sub Command3_Click()                'Add Product
    Dim ds As adodb.Recordset, sqlx As String, msku As String, mrun As Long
    Dim mgroup As String, mplant As Integer, mbranch As Integer, maccount As String, pkey As Long
    Dim i As Integer, mtno As String
    On Error GoTo vberror
    msku = InputBox$("Please enter SKU for Product", "New Product", "777")
    If Len(msku) = 0 Then Exit Sub
    sqlx = "select * from skumast where sku = '" & msku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "SKU number not found in list..", vbOKOnly, "Invalid SKU"
        ds.Close ': db.Close
        Exit Sub
    End If
    ds.Close
    tid.ListIndex = td.Row - 1
    sqlx = "select * from trailers where id = " & tid
    Set ds = Sdb.Execute(sqlx)
    ds.MoveFirst
    mrun = ds!runid
    mgroup = ds!groupcode
    mplant = ds!plant
    mbranch = ds!branch
    maccount = ds!account
    mtno = ds!trlno
    
    pkey = wd_seq("trailers", Form1.shipdb)
    sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno, sku"
    sqlx = sqlx & ", pallets, wraps, units, whs_num, pb_flag, ra_flag) Values (" & pkey
    sqlx = sqlx & ", " & mrun
    sqlx = sqlx & ", '" & mgroup & "'"
    sqlx = sqlx & ", '" & mplant & "'"
    sqlx = sqlx & ", '" & mbranch & "'"
    sqlx = sqlx & ", '" & maccount & "'"
    sqlx = sqlx & ", '" & sd & "'"
    sqlx = sqlx & ", '" & mtno & "'"
    sqlx = sqlx & ", '" & msku & "'"
    sqlx = sqlx & ", 0, 0, 0"
    If mplant = "51" Then
        sqlx = sqlx & ", 14"
    Else
        sqlx = sqlx & ", 0"
    End If
    sqlx = sqlx & ", 'N', 'N')"
    Sdb.Execute sqlx
    ds.Close
    Call refresh_grid
    For i = 1 To td.Rows - 1
        If td.TextMatrix(i, 0) = msku Then
            td.Row = i
            If i > (td.Height / 245) Then td.TopRow = i
            Exit For
        End If
    Next i
    Call td_Click
    If Text2.Visible = True Then Text2.SetFocus
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

Private Sub Command3_GotFocus()
    Command3.FontBold = True
End Sub

Private Sub Command3_LostFocus()
    Command3.FontBold = False
End Sub

Private Sub Command4_Click()        'Rack Checkoff
    Call ckoffsheet2011
End Sub

Private Sub Command4_GotFocus()
    Command4.FontBold = True
End Sub

Private Sub Command4_LostFocus()
    Command4.FontBold = False
End Sub

Private Sub Command5_Click()            'Blank Bill
    blnkbill.Show
End Sub

Private Sub Command6_Click()                'Post To Cranes
    Dim cfile As String, i As Integer, s As String
    cfile = localAppDataPath & "\cranereq.txt"  ' C:\cranereq.txt
    Open cfile For Output As #1
    For i = 1 To td.Rows - 1
        If Val(td.TextMatrix(i, 2)) > 0 And td.TextMatrix(i, 5) = "Crane" Then
            s = "Insert into tBBCOrders (sItemID, iQuantity, iPalletType, sLotID, bAutoRelease)"
            s = s & " Values ('" & td.TextMatrix(i, 0) & "', "
            If Val(LotID) = 0 Then
                s = s & Val(pallets) & ", 1, 0, " & Val(AutoReleaseCheck1) & ")"
                's = s & Val(pallets) & ", 1, NULL, " & Val(AutoReleaseCheck1) & ")"
            Else
                s = s & Val(pallets) & ", 1, " & Val(LotID) & ", " & Val(AutoReleaseCheck1) & ")"
            End If
            Print #1, s
        End If
    Next i
    Close #1
    
    MsgBox s
    
    Set Sdb = CreateObject("ADODB.Connection")
    Sdb.Open "Driver={SQL Server};server=bbsy-01-westfalia;DATABASE=BlueBell_WMS;UID=PostToBbcorders;PWD=postorders"
    Sdb.Execute s
    Sdb.Close
    
    'Dim errLoop As Error
      'Dim strError As String

      'i = 1

   ' Process
     StrTmp = StrTmp & vbCrLf & "VB Error # " & Str(Err.Number)
     StrTmp = StrTmp & vbCrLf & "   Generated by " & Err.source
     StrTmp = StrTmp & vbCrLf & "   Description  " & Err.description

   ' Enumerate Errors collection and display properties of
   ' each Error object.
     'Set Errs1 = Sdb.Errors
     'For Each errLoop In Errs1
          'With errLoop
            'StrTmp = StrTmp & vbCrLf & "Error #" & i & ":"
            'StrTmp = StrTmp & vbCrLf & "   ADO Error   #" & .Number
            'StrTmp = StrTmp & vbCrLf & "   Description  " & .description
            'StrTmp = StrTmp & vbCrLf & "   Source       " & .source
            'i = i + 1
       'End With
    'Next

      'MsgBox StrTmp

      ' Clean up Gracefully

      'On Error Resume Next
      'GoTo Done

    
    s = cfile & " has been posted..."
    MsgBox s, vbOKOnly + vbInformation, cfile
End Sub

Private Sub Command7_Click()
    Call post_r12_bill
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Edittrl.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
            If Form1.FrmGrid.Text = "edittrl" Then
                Form1.FrmGrid.Col = 1: Form1.FrmGrid.Text = Me.Top
                Form1.FrmGrid.Col = 2: Form1.FrmGrid.Text = Me.Left
                Form1.FrmGrid.Col = 3: Form1.FrmGrid.Text = Me.Height
                Form1.FrmGrid.Col = 4: Form1.FrmGrid.Text = Me.Width
                Exit For
            End If
        Next i
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Edittrl.ActiveControl.Name = "td" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command3_Click 'insert
        If KeyCode = 46 Or KeyCode = 120 Then Call Command2_Click 'delete
    End If
End Sub

Private Sub Form_Load()
    Dim ds As adodb.Recordset
    Dim i As Integer
    On Error GoTo vberror
    For i = 1 To Form1.FrmGrid.Rows - 1
        Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
        If Form1.FrmGrid.Text = "edittrl" Then
            Form1.FrmGrid.Col = 1: Me.Top = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 2: Me.Left = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 3: Me.Height = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 4: Me.Width = Val(Form1.FrmGrid.Text)
            Exit For
        End If
    Next i
    Set ds = Sdb.Execute("select distinct shipdate from trailers order by shipdate")
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sd.AddItem ds(0)
            ds.MoveNext
        Loop
    End If
    ds.Close
    td.Font = "Arial": td.FontSize = 9: td.FontBold = True
    If Form1.plantno = "52" Then
        td.Cols = 6
        td.FormatString = "^SKU|<Product|^Pallets|^Wraps|^Units|^Source"
        td.ColWidth(0) = 500
        td.ColWidth(1) = 3500: td.ColWidth(2) = 700
        td.ColWidth(3) = 700: td.ColWidth(4) = 700
        td.ColWidth(5) = 700
    Else
        'td.Cols = 5
        td.Cols = 6
        td.FormatString = "^SKU|<Product|^Pallets|^Wraps|^Units"
        td.ColWidth(0) = 500
        td.ColWidth(1) = 3500: td.ColWidth(2) = 800
        td.ColWidth(3) = 800: td.ColWidth(4) = 800
        td.ColWidth(5) = 1
    End If
    If sd.ListCount > 0 Then sd.ListIndex = 0
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
    If Me.Height > 4000 Then
        td.Height = Me.Height - 900 '(sd.Height + 375)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
    If Form1.WindowState = 1 Then End
End Sub

Private Sub plantno_Change()
    If Val(plantno.Caption) = 0 Then Exit Sub
    If Val(plantno.Caption) = 50 Then
        Command4.Visible = False
    Else
        Command4.Visible = True
    End If
    'If Val(plantno.Caption) = 52 Then
    '    Command6.Visible = True
    'Else
    '    Command6.Visible = False
    'End If
End Sub

Private Sub sd_Click()
    Dim ds As adodb.Recordset, js As adodb.Recordset, sqlx As String
    td.Rows = 1: Combo1.Clear: List1.Clear
    On Error GoTo vberror
    sqlx = "Select runid,trailers.branch,account,branchname,trlno,sum(units) from trailers,branches"
    sqlx = sqlx & " Where shipdate = '" & sd & "'"
    sqlx = sqlx & " And trailers.branch = branches.branch"
    sqlx = sqlx & " and trailers.plant = " & Form1.plantno
    sqlx = sqlx & " Group by runid,trailers.branch,account,branchname,trlno"
    sqlx = sqlx & " order by branchname,trlno"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        ds.Close
        MsgBox "No trailers found for selected date..", vbOKOnly, "Schedule"
        Exit Sub
    End If
    ds.MoveFirst
    Do Until ds.EOF
        If ds!account <= "0" Then
            Combo1.AddItem ds!branchname & " " & ds!trlno
            List1.AddItem ds!runid & "......"
        Else
            sqlx = "select * from jobbing where branch = " & ds!branch & " and account = '" & ds!account & "'"
            Set js = Sdb.Execute(sqlx)
            If js.BOF = False Then
                js.MoveFirst
                Combo1.AddItem js!acctdesc
            Else
                Combo1.AddItem "......"
            End If
            js.Close
            List1.AddItem ds!runid & ds!account
        End If
        ds.MoveNext
    Loop
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "sd_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " sd_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub td_Click()
    If wc.ListCount > 0 Then
        wc.ListIndex = td.Row - 1
        pc.ListIndex = td.Row - 1
        tid.ListIndex = td.Row - 1
    End If
    Label1 = Trim$(td.TextMatrix(td.Row, 1))
    Text2 = Val(td.TextMatrix(td.Row, 2))
    Text3 = Val(td.TextMatrix(td.Row, 3))
    Text4 = Val(td.TextMatrix(td.Row, 4))
    Text5 = td.TextMatrix(td.Row, 5)
    Label5 = "@ " & pc
    Label6 = "@ " & wc
End Sub

Private Sub td_GotFocus()
    td.BackColorBkg = &H80000002
    td.BackColor = &H80000005
End Sub

Private Sub td_KeyPress(KeyAscii As Integer)
    If td.Row = 0 Then Exit Sub
    If td.Col = 2 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            Text2 = Text2 & Chr(KeyAscii)
            Call Text2_KeyUp(KeyAscii, 0)
        End If
        If KeyAscii = 8 Then
            If Len(td.Text) > 1 Then
                Text2 = Left(Text2, Len(Text2) - 1)
            Else
                Text2 = ""
            End If
            Call Text2_KeyUp(8, 0)
        End If
    End If
    If td.Col = 3 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            Text3 = Text3 & Chr(KeyAscii)
            Call Text3_KeyUp(KeyAscii, 0)
        End If
        If KeyAscii = 8 Then
            If Len(td.Text) > 1 Then
                Text3 = Left(Text3, Len(Text3) - 1)
            Else
                Text3 = ""
            End If
            Call Text3_KeyUp(8, 0)
        End If
    End If
    If td.Col = 4 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            Text4 = Text4 & Chr(KeyAscii)
            Call Text4_KeyUp(KeyAscii, 0)
        End If
        If KeyAscii = 8 Then
            If Len(td.Text) > 1 Then
                Text4 = Left(Text4, Len(Text4) - 1)
            Else
                Text4 = ""
            End If
            Call Text4_KeyUp(8, 0)
        End If
    End If
    If td.Col = 5 Then
        If Text5 = "Crane" Then
            Text5 = "Rack"
        Else
            Text5 = "Crane"
        End If
        If Me.plantno = "52" Then Call sywhs(Text5)
        td.TextMatrix(td.Row, 5) = Text5
    End If
End Sub

Private Sub td_LostFocus()
    td.BackColorBkg = Edittrl.BackColor
    td.BackColor = Edittrl.BackColor
End Sub

Private Sub td_RowColChange()
    If td.Row <> srow And td.Redraw = True Then
        srow = td.Row
        Call td_Click
    End If
End Sub

Private Sub Text2_GotFocus()
    Text2.SelStart = 0: Text2.SelLength = Len(Text2)
    Text2.FontBold = True
    Label2.FontBold = True: Label5.FontBold = True
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    Text4 = (Val(Text2) * Val(pc)) + (Val(Text3) * Val(wc))
    Call update_trl
End Sub

Private Sub Text2_LostFocus()
    Label2.FontBold = False: Label5.FontBold = False
    Text2.FontBold = False
End Sub

Private Sub Text3_GotFocus()
    Text3.SelStart = 0: Text3.SelLength = Len(Text3)
    Text3.FontBold = True
    Label3.FontBold = True: Label6.FontBold = True
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If
End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)
    Text4 = (Val(Text2) * Val(pc)) + (Val(Text3) * Val(wc))
    Call update_trl
End Sub

Private Sub Text3_LostFocus()
    Text3.FontBold = False
    Label3.FontBold = False: Label6.FontBold = False
End Sub

Private Sub Text4_GotFocus()
    Text4.SelStart = 0: Text4.SelLength = Len(Text4)
    Text4.FontBold = True
    Label4.FontBold = True
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: SendKeys "{TAB}"
    End If
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
    Call update_trl
End Sub

Private Sub Text4_LostFocus()
    Label4.FontBold = False: Text4.FontBold = False
End Sub
