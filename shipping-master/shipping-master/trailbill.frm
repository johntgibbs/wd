VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form trailbill 
   Caption         =   "Edit Trailers"
   ClientHeight    =   12720
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   16440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   12720
   ScaleWidth      =   16440
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox AutoRelease 
      Caption         =   "AutoRelease"
      Height          =   255
      Left            =   12840
      TabIndex        =   53
      Top             =   3120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid addgrid 
      Height          =   1095
      Left            =   0
      TabIndex        =   52
      Top             =   4080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1931
      _Version        =   327680
      BackColorSel    =   192
      FocusRect       =   0
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   10200
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox ldate 
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
      Left            =   1680
      TabIndex        =   47
      Text            =   "ldate"
      Top             =   5400
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid trkgrid 
      Height          =   3135
      Left            =   6960
      TabIndex        =   45
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5530
      _Version        =   327680
      BackColorFixed  =   12648384
      BackColorSel    =   16384
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1815
      Left            =   6840
      TabIndex        =   36
      Top             =   9120
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3201
      _Version        =   327680
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View All Fields"
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
      Left            =   3120
      TabIndex        =   35
      Top             =   5400
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid billgrid 
      Height          =   6855
      Left            =   0
      TabIndex        =   33
      Top             =   5760
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   12091
      _Version        =   327680
      Cols            =   12
      BackColorFixed  =   12648447
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Post to Westfalia"
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
      Left            =   12840
      TabIndex        =   29
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   1455
      Left            =   4320
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2566
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid tmpgrid 
      Height          =   1575
      Left            =   4080
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2778
      _Version        =   327680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Rack Check Off"
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
      Left            =   6960
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.ComboBox sd 
      BackColor       =   &H00FFFFFF&
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
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add Product"
      Height          =   375
      Left            =   9480
      TabIndex        =   12
      Top             =   11640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel Product"
      Height          =   375
      Left            =   9480
      TabIndex        =   13
      Top             =   12120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read Scans"
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
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Product "
      Height          =   2295
      Left            =   9360
      TabIndex        =   7
      Top             =   9240
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   720
         TabIndex        =   28
         Text            =   "Text5"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Text            =   "Text4"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Text            =   "Text3"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Source"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Units"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Wraps"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Pallets"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
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
      Height          =   255
      Left            =   10080
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox wc 
      Height          =   645
      Left            =   9120
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox pc 
      Height          =   255
      Left            =   10080
      TabIndex        =   4
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid td 
      Height          =   3495
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6165
      _Version        =   327680
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16777152
      BackColorSel    =   128
      BackColorBkg    =   -2147483638
      GridColor       =   14737632
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   10440
      TabIndex        =   2
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Label duplicateBarcodes 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Trailer Contains Duplicate Barcodes."
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
      Left            =   12960
      TabIndex        =   54
      Top             =   5400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Oracle Ticket"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9480
      TabIndex        =   51
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Group Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9480
      TabIndex        =   50
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label tempdir 
      Caption         =   "v:\temp"
      Height          =   255
      Left            =   10800
      TabIndex        =   48
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label r12tkt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11160
      TabIndex        =   46
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label postlit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6960
      TabIndex        =   44
      Top             =   3720
      Width           =   5775
   End
   Begin VB.Label logname 
      Caption         =   "logname"
      Height          =   255
      Left            =   9480
      TabIndex        =   43
      Top             =   3960
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label rcolor 
      BackColor       =   &H000000FF&
      Caption         =   "rcolor"
      Height          =   255
      Left            =   10320
      TabIndex        =   42
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label gcolor 
      BackColor       =   &H0000FF00&
      Caption         =   "gcolor"
      Height          =   255
      Left            =   10320
      TabIndex        =   41
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Scanned Units Do Not Match Original Order."
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
      Left            =   8760
      TabIndex        =   40
      Top             =   5400
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label hcolor 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "All Records"
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
      Left            =   5040
      TabIndex        =   39
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label cntlit 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Records"
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
      TabIndex        =   38
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label trlkey 
      Caption         =   "Label10"
      Height          =   255
      Left            =   8040
      TabIndex        =   37
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label pallogs 
      Caption         =   "Label10"
      Height          =   255
      Left            =   1440
      TabIndex        =   34
      Top             =   7320
      Width           =   3855
   End
   Begin VB.Label shipdb 
      Caption         =   "Label10"
      Height          =   255
      Left            =   1320
      TabIndex        =   32
      Top             =   7800
      Width           =   7575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bill of Lading Printed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6960
      TabIndex        =   31
      Top             =   5160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label gcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "___"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11160
      TabIndex        =   30
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label plantno 
      Caption         =   "50"
      Height          =   255
      Left            =   7560
      TabIndex        =   24
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
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
      Left            =   8160
      TabIndex        =   21
      Top             =   6000
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
      Left            =   6840
      TabIndex        =   20
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
      Begin VB.Menu printbill 
         Caption         =   "Bill of Lading"
      End
      Begin VB.Menu prtblank 
         Caption         =   "Blank Bill"
      End
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu edtrls 
         Caption         =   "Trailer Order"
         Begin VB.Menu addtline 
            Caption         =   "Add Line"
         End
         Begin VB.Menu deltline 
            Caption         =   "Delete Line"
         End
      End
      Begin VB.Menu edscans 
         Caption         =   "Scanned Pallets"
         Begin VB.Menu canline 
            Caption         =   "Cancel Line"
         End
         Begin VB.Menu addbc 
            Caption         =   "Add New Barcode"
         End
         Begin VB.Menu addwraps 
            Caption         =   "Add New Product - Wraps"
         End
         Begin VB.Menu edunits 
            Caption         =   "Edit Units"
         End
      End
      Begin VB.Menu edsched 
         Caption         =   "Schedule"
         Begin VB.Menu edtc 
            Caption         =   "Edit Trailer Code"
         End
         Begin VB.Menu edseal 
            Caption         =   "Edit Seal"
         End
         Begin VB.Menu edins 
            Caption         =   "Edit Inspected By"
         End
         Begin VB.Menu edspec 
            Caption         =   "Edit Special Instructions"
         End
         Begin VB.Menu edfr 
            Caption         =   "Edit Freight"
         End
      End
   End
   Begin VB.Menu postmenu 
      Caption         =   "Post"
      Begin VB.Menu postr12 
         Caption         =   "Post to Oracle Batches"
      End
   End
   Begin VB.Menu renmenu 
      Caption         =   "Rename Trailer"
      Begin VB.Menu rentrl 
         Caption         =   "Rename Trailer"
      End
   End
   Begin VB.Menu debmenu 
      Caption         =   "Debug"
      Enabled         =   0   'False
      Begin VB.Menu tpost 
         Caption         =   "Test Post"
      End
   End
End
Attribute VB_Name = "trailbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim srow As Integer

Function bbdriver(dname As String) As String
    Dim i As Integer, s As String
    i = InStr(1, dname, ",")
    If i = 0 Then
        s = dname
    Else
        If UCase(Right(dname, 3)) = "SY " Then dname = Left(dname, Len(dname) - 3)
        If UCase(Right(dname, 3)) = "EP " Then dname = Left(dname, Len(dname) - 3)
        If UCase(Right(dname, 3)) = "CH " Then dname = Left(dname, Len(dname) - 3)
        If UCase(Right(dname, 3)) = "BA " Then dname = Left(dname, Len(dname) - 3)
        If UCase(Right(dname, 4)) = "MOB " Then dname = Left(dname, Len(dname) - 4)
        s = Trim(mid(dname, i + 1, Len(dname) - i))
        s = s & " " & Left(dname, i - 1)
    End If
    bbdriver = s
End Function

Private Sub refresh_addgrid()
    Dim cfile As String, f0 As String, f1 As String, ds As adodb.Recordset
    Dim f2 As String, f3 As String, f4 As String, f5 As String, f6 As String
    Dim s As String
    addgrid.Clear: addgrid.Rows = 1: addgrid.Cols = 5
    cfile = Form1.webdir & "\locflist.csv"
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6
            s = f0 & Chr(9) & f1 & Chr(9) & f2 & ", " & f3 & " " & f4 & Chr(9)
            s = s & f5 & Chr(9) & f6
            addgrid.AddItem s
        Loop
        Close #1
    Else                                                                'jv100815
        s = "select * from branches where gemmsid > '0'"                'jv100815
        Set ds = Sdb.Execute(s)                                         'jv100815
        If ds.BOF = False Then                                          'jv100815
            ds.MoveFirst                                                'jv100815
            Do Until ds.EOF                                             'jv100815
                s = ds!branchname & Chr(9)                              'jv100815
                s = s & ds!addr1 & Chr(9)                               'jv100815
                s = s & ds!addr2 & Chr(9)                               'jv100815
                s = s & ds!brphone & Chr(9)                             'jv100815
                s = s & ds!brfax
                addgrid.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    addgrid.FormatString = "<Ship To|<Address|<City, State Zip|^Phone|^Fax"
    addgrid.ColWidth(0) = 1800
    addgrid.ColWidth(1) = 2300
    addgrid.ColWidth(2) = 2200
    addgrid.ColWidth(3) = 1300
    addgrid.ColWidth(4) = 1300
    If addgrid.Rows > 1 Then
        addgrid.Row = 1: addgrid.RowSel = 1
        addgrid.Col = 0: addgrid.ColSel = 1
        addgrid.Sort = 5
    End If
End Sub

Private Sub save_bp_temp()
    Dim cfile As String, i As Integer, k As Integer
    If billgrid.Rows < 2 Then Exit Sub
    cfile = Me.tempdir & "\bp" & r12tkt.Caption & ".tmp"
    Open cfile For Output As #1
    For i = 1 To billgrid.Rows - 1
        If billgrid.TextMatrix(i, 14) <> "CANC" Then                'jv081915
            For k = 1 To billgrid.Cols - 2
                Write #1, billgrid.TextMatrix(i, k);
            Next k
            Write #1, billgrid.TextMatrix(i, billgrid.Cols - 1)
        End If                                                      'jv081915
    Next i
    Close #1
End Sub

Private Sub refresh_trkgrid(prun As String)
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String
    Dim d1 As Long, pwo As Long
    Dim eno As Long, edesc As String, sorg As String
    On Error GoTo vberror
    trkgrid.Clear: trkgrid.Cols = 2: trkgrid.Rows = 1
    pwo = 0
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.schdb
    s = "select w.wonum, d.driver, w.startime, w.contents, w.trlsize, w.eqnum, w.description, w.drvid, w.sealnum, w.SpecialInstructions,"
    s = s & " w.InspectedBy, w.Freight from truckwo w, drivers d"
    s = s & " where w.r12ticket = '" & prun & "'"
    s = s & " and d.id = w.drvid" ' and w.parentwo = 0" jv081015
    s = s & " and w.wostatus not in ('CANC', 'CLOSE')"                        'jv092415
    s = s & " order by w.parentwo"                      'jv081015
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        pwo = ds(0)
        d1 = ds(7)
        trkgrid.AddItem "WO Number" & Chr(9) & ds(0)
        trkgrid.AddItem "Driver" & Chr(9) & bbdriver(ds(1) & " ")
        trkgrid.AddItem "Start Time" & Chr(9) & ds(2)
        If ds(3) <> "IceCream" Then
            trkgrid.AddItem "Contents" & Chr(9) & ds(3)
            trkgrid.AddItem "Note" & Chr(9) & ds(6)
        End If
        trkgrid.AddItem "Trailer Size" & Chr(9) & ds(4)
        trkgrid.AddItem "Trailer Code" & Chr(9) & ds(5)
        trkgrid.AddItem "Seal #" & Chr(9) & ds(8)
        trkgrid.AddItem "Special Instructions" & Chr(9) & ds(9)
        trkgrid.AddItem "Inspected By" & Chr(9) & ds(10)
        trkgrid.AddItem "Freight" & Chr(9) & ds(11)
    End If
    ds.Close
    If pwo = 0 Then
        If plantno = "50" Then sorg = "T10"
        If plantno = "51" Then sorg = "K10"
        If plantno = "52" Then sorg = "A10"
        If Val(ano) > 0 Then
            s = "select w.wonum, d.driver, w.startime, w.contents, w.trlsize, w.eqnum, w.description, w.drvid, w.sealnum,"
            s = s & " w.SpecialInstructions, w.InspectedBy, w.Freight from truckwo w, drivers d"
            's = s & " where w.origin = '" & sorg & "'"
            s = s & " where w.destination in (select lcode from locations where jobaccount = '" & ano & "')"
            s = s & " and w.wodate = '" & sd & "'"
            s = s & " and d.id = w.drvid" ' and w.parentwo = 0"
        Else
            s = "select w.wonum, d.driver, w.startime, w.contents, w.trlsize, w.eqnum, w.description, w.drvid, w.sealnum, "
            s = s & " w.SpecialInstructions, w.InspectedBy, w.Freight from truckwo w, drivers d"
            s = s & " where w.origin = '" & sorg & "'"
            s = s & " and w.destination = '" & Format(Val(bno), "000") & "'"
            s = s & " and w.trlno = '" & Right(Combo1, 1) & "'"
            s = s & " and w.wodate = '" & sd & "'"
            s = s & " and d.id = w.drvid and w.parentwo = 0"
        End If
        s = s & " and w.wostatus not in ('CANC', 'CLOSE')"                        'jv092415
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            pwo = ds(0)
            d1 = ds(7)
            trkgrid.AddItem "WO Number" & Chr(9) & ds(0)
            'trkgrid.AddItem "Driver" & Chr(9) & ds(1)
            trkgrid.AddItem "Driver" & Chr(9) & bbdriver(ds(1) & " ")
            trkgrid.AddItem "Start Time" & Chr(9) & ds(2)
            If ds(3) <> "IceCream" Then
                trkgrid.AddItem "Contents" & Chr(9) & ds(3)
                trkgrid.AddItem "Note" & Chr(9) & ds(6)
            End If
            trkgrid.AddItem "Trailer Size" & Chr(9) & ds(4)
            trkgrid.AddItem "Trailer Code" & Chr(9) & ds(5)
            trkgrid.AddItem "Seal #" & Chr(9) & ds(8)
            trkgrid.AddItem "Special Instructions" & Chr(9) & ds(9)
            trkgrid.AddItem "Inspected By" & Chr(9) & ds(10)
            trkgrid.AddItem "Freight" & Chr(9) & ds(11)
        End If
        ds.Close
    End If
        
        
        
    If pwo > 0 Then
        'Swap Driver
        s = "select w.drvid, d.driver from truckwo w, drivers d where w.parentwo = " & pwo
        's = s & " and w.wtype = 'Swap'"
        s = s & " and w.wtype <> 'Co-Driver'"
        s = s & " and w.drvid <> " & d1
        s = s & " and d.id = w.drvid"
        s = s & " and w.wostatus not in ('CANC', 'CLOSE')"                        'jv092415
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            trkgrid.AddItem "2nd Driver" & Chr(9) & bbdriver(ds(1) & " ")
        End If
        ds.Close
    End If
    db.Close
    If Val(ano.Caption) > 0 Then trkgrid.AddItem "Jobbing Account" & Chr(9) & ano.Caption
    'trkgrid.AddItem "Inspected By"
    'trkgrid.AddItem "Special Instructions"
    'trkgrid.AddItem "Freight"
    trkgrid.FormatString = "<Field|<Schedule Value"
    trkgrid.ColWidth(0) = 1500
    trkgrid.ColWidth(1) = 3500
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "refresh_trkgrid(" & prun & ")", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_trkgrid(" & prun & ") - Error Number: " & eno
        End
    End If
End Sub

Private Sub check_totals(runno As String)
    Dim ds As adodb.Recordset, s As String, adflag As Boolean
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 5
    Grid2.Visible = False
    ycolor.Visible = False
    duplicateBarcodes.Visible = False
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
    billgrid.Redraw = False
    billgrid.FillStyle = flexFillRepeat
    Grid2.FillStyle = flexFillRepeat
    If Grid2.Rows > 1 Then
        For i = 1 To billgrid.Rows - 1
            If billgrid.TextMatrix(i, 17) = Grid2.TextMatrix(1, 0) And billgrid.TextMatrix(i, 14) <> "CANC" Then
                adflag = True
                For k = 1 To Grid2.Rows - 1
                    If Grid2.TextMatrix(k, 1) = Trim(Left(billgrid.TextMatrix(i, 6), 4)) Then                       'jv093015
                        Grid2.TextMatrix(k, 3) = Val(Grid2.TextMatrix(k, 3)) + Val(billgrid.TextMatrix(i, 11))
                        Grid2.TextMatrix(k, 3) = Val(Grid2.TextMatrix(k, 3)) + Val(billgrid.TextMatrix(i, 13))
                        adflag = False
                        Exit For
                    End If
                Next k
                If adflag = True Then
                    s = Grid2.TextMatrix(1, 0) & Chr(9)
                    s = s & Trim(Left(billgrid.TextMatrix(i, 6), 4)) & Chr(9)                                       'jv093015
                    s = s & "0" & Chr(9)
                    k = Val(billgrid.TextMatrix(i, 11))
                    k = k + Val(billgrid.TextMatrix(i, 13))
                    s = s & k
                    Grid2.AddItem s
                End If
                billgrid.Row = i: billgrid.RowSel = i
                billgrid.Col = 6: billgrid.ColSel = 13
                billgrid.CellBackColor = billgrid.BackColor
                billgrid.Col = 1
            End If
        Next i
                
        For i = 1 To Grid2.Rows - 1
            Grid2.TextMatrix(i, 4) = Val(Grid2.TextMatrix(i, 3)) - Val(Grid2.TextMatrix(i, 2))
            If Val(Grid2.TextMatrix(i, 4)) <> 0 Then
                For k = 1 To billgrid.Rows - 1
                    If billgrid.TextMatrix(k, 17) = Grid2.TextMatrix(i, 0) And Trim(Left(billgrid.TextMatrix(k, 6), 4)) = Grid2.TextMatrix(i, 1) Then  'jv093015
                        billgrid.Row = k: billgrid.RowSel = k
                        billgrid.Col = 6: billgrid.ColSel = 13
                        billgrid.CellBackColor = ycolor.BackColor
                        'Grid2.Visible = True
                        'ycolor.Visible = True
                        billgrid.Col = 1
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
    
    ' Check for duplicate barcodes on this bill
    If billgrid.Rows > 1 Then
        For i = 1 To billgrid.Rows - 1
            Dim currentBarcode As String
            currentBarcode = billgrid.TextMatrix(i, 7)
            Dim currentCount As Integer
            currentCount = 0
            
            ' Ignore "LOT1" for jobbing/OP
            If billgrid.TextMatrix(i, 10) <> "LOT1" Then
                ' Do second loop to search barcodes for the one in outer loop
                For j = 1 To billgrid.Rows - 1
                    If billgrid.TextMatrix(j, 7) = currentBarcode Then
                        ' If found, increment count
                        currentCount = currentCount + 1
                        ' Check for 2 because it will find itself every time so 1 is expected
                        If currentCount >= 2 Then
                            ' Duplicate found, set visibility of error message
                            duplicateBarcodes.Visible = True
                        End If
                    End If
                Next j
            End If
        Next i
    End If
    
    billgrid.Redraw = True
    Grid2.FormatString = "^RunId|^SKU|^Ordered|^Scanned|^Diff"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 1000
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "check_totals", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " check_totals - Error Number: " & eno
        End
    End If
End Sub

Private Sub postro_bill(mplant As String, sdate As String)
    Dim ofile As String, s As String, rfile As String, tktno As Long
    Dim i As Integer, k As Integer, addfile As Boolean, ftpexe As String
    Dim ds As adodb.Recordset
    Dim x, eno As Long, edesc As String
    'On Error GoTo vberror
    ofile = Me.pallogs & "bill" & sdate & ".txt"
    Open ofile For Append As #1
    For i = 1 To billgrid.Rows - 1
        If billgrid.TextMatrix(i, 0) = "B" Then
            Write #1, billgrid.TextMatrix(i, 1);        'Batch
            Write #1, billgrid.TextMatrix(i, 2);        'Area
            Write #1, billgrid.TextMatrix(i, 3);        'Group
            Write #1, billgrid.TextMatrix(i, 4);        'Source
            Write #1, billgrid.TextMatrix(i, 5);        'Target
            Write #1, billgrid.TextMatrix(i, 6);        'Product
            Write #1, billgrid.TextMatrix(i, 7);        'Barcode
            Write #1, billgrid.TextMatrix(i, 8);        'Qty
            Write #1, billgrid.TextMatrix(i, 9);        'Uom
            Write #1, billgrid.TextMatrix(i, 10);        'Lot
            Write #1, billgrid.TextMatrix(i, 11);        'Units
            Write #1, billgrid.TextMatrix(i, 12);        'Lot2
            Write #1, billgrid.TextMatrix(i, 13);        'Units
            Write #1, "POSTED";                         'Status
            Write #1, billgrid.TextMatrix(i, 15);        'User
            Write #1, billgrid.TextMatrix(i, 16);        'DateTime
            Write #1, billgrid.TextMatrix(i, 17)         'Ticket
        End If
    Next i
    Close #1
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
    tktno = 0
    s = "select * from trailers where runid = " & r12tkt.Caption
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        rfile = Me.pallogs & "RO" & ds!runid & ".txt"
        tktno = ds!runid
        Open rfile For Append As #5
        For i = 1 To billgrid.Rows - 1
            If billgrid.TextMatrix(i, 0) = "B" Then
                If billgrid.TextMatrix(i, 7) > "00" Then  'barcode indicates pallet
                    Write #5, ds!runid & "P";
                    Write #5, morg;
                    Write #5, mwhs;
                    Write #5, "FLOOR" & mwhs;
                    If ds!branch = 47 Then
                        Write #5, "501"; "K10"; "FLOORK10";
                    Else
                        If ds!branch = 52 Then
                            Write #5, "502"; "A10"; "FLOORA10";
                        Else
                            If ds!branch = 1 Then
                                Write #5, "500"; "T10"; "FLOORT10";
                            Else
                                Write #5, Format(ds!branch, "000");
                                Write #5, Format(ds!branch, "000");
                                Write #5, "FLOOR" & Format(ds!branch, "000");
                            End If
                        End If
                    End If
                    Write #5, ds!account;
                    Write #5, Trim(Left(billgrid.TextMatrix(i, 7), 4));
                    Write #5, Trim(mid(billgrid.TextMatrix(i, 7), 5, 9));            'jv052515     'lot
                    Write #5, Format(Val(billgrid.TextMatrix(i, 11)), "0");
                    Write #5, "EACH";
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                    Write #5, StrConv(billgrid.TextMatrix(i, 5), vbProperCase) & " " & Right(billgrid.TextMatrix(i, 7), 3);
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy");
                    If Left(ds!trlno, 1) = "B" Or ds!branch = 15 Or ds!branch = 16 Then
                        Write #5, "Y"
                    Else
                        Write #5, "N"
                    End If
                        
                    If Val(billgrid.TextMatrix(i, 12)) > 0 Then   '2nd lot
                        f7 = billgrid.TextMatrix(i, 7)
                        f10 = billgrid.TextMatrix(i, 10)
                        f12 = billgrid.TextMatrix(i, 12)
                        s = r12_lot(Left(f12, 5), mid(f12, 6, 3))        'jv011916  'here figure out what r12_lot does and what values it's being sent
                        Write #5, ds!runid & "P";
                        Write #5, morg;
                        Write #5, mwhs;
                        Write #5, "FLOOR" & mwhs;
                        If ds!branch = 47 Then
                            Write #5, "501"; "K10"; "FLOORK10";
                        Else
                            If ds!branch = 52 Then
                                Write #5, "502"; "A10"; "FLOORA10";
                            Else
                                If ds!branch = 1 Then
                                    Write #5, "500"; "T10"; "FLOORT10";
                                Else
                                    Write #5, Format(ds!branch, "000");
                                    Write #5, Format(ds!branch, "000");
                                    Write #5, "FLOOR" & Format(ds!branch, "000");
                                End If
                            End If
                        End If
                        Write #5, ds!account;
                        Write #5, Trim(Left(billgrid.TextMatrix(i, 7), 4));
                        Write #5, StringReplace(s, " ", "");
                        Write #5, Format(Val(billgrid.TextMatrix(i, 13)), "0");
                        Write #5, "EACH";
                        Write #5, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                        Write #5, StrConv(billgrid.TextMatrix(i, 5), vbProperCase) & " " & Right(f7, 3);
                        Write #5, Format(ds!shipdate, "MM-dd-yyyy");
                        If Left(ds!trlno, 1) = "B" Or ds!branch = 15 Or ds!branch = 16 Then
                            Write #5, "Y"
                        Else
                            Write #5, "N"
                        End If
                    End If
                Else
                    Write #5, ds!runid & "W";
                    If mplant = "50" Then Write #5, "001"; "001"; "FLOOR001";
                    If mplant = "51" Then Write #5, "047"; "047"; "FLOOR047";
                    If mplant = "52" Then Write #5, "052"; "052"; "FLOOR052";
                    Write #5, Format(ds!branch, "000");
                    Write #5, Format(ds!branch, "000");
                    Write #5, "FLOOR" & Format(ds!branch, "000");
                    Write #5, ds!account;
                    Write #5, Trim(Left(billgrid.TextMatrix(i, 6), 4));
                    Write #5, "LOT1"; 'billgrid.TextMatrix(i, 10);                                'jv082915
                    Write #5, Format(Val(billgrid.TextMatrix(i, 11)), "0");
                    Write #5, "EACH";
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                    Write #5, StrConv(billgrid.TextMatrix(i, 5), vbProperCase);
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy");
                    If Left(ds!trlno, 1) = "B" Or ds!branch = 15 Or ds!branch = 16 Then
                        Write #5, "Y"
                    Else
                        Write #5, "N"
                    End If
                End If
            End If
        Next i
        Close #5
        ds.Close
    End If
    
    'Exit Sub
    addfile = False
    If tktno > 0 Then
        ofile = Form1.pallogs & "r12trls.win"
        Open ofile For Output As #1
        Print #1, "open pbelle.bluebell.com"
        Print #1, "infbbcri"
        'Print #1, "infbbcri"
        Print #1, "welcome@2023"
        Print #1, "BINARY"
        'Print #1, "cd /interface/infbbcri/PBELLE/incoming"
        Print #1, "cd PBELLE/incoming"
        Print #1, "lcd "; Left(Form1.pallogs, Len(Form1.pallogs) - 1)
        rfile = Form1.pallogs & "RO" & tktno & ".txt"
        If Len(Dir(rfile)) > 0 Then
            s = "put RO" & tktno & ".txt RO" & tktno & ".txt"
            Print #1, s
            addfile = True
        End If
            
        s = "Update trailers set pb_flag = 'Y', ra_flag = 'Y' where runid = " & tktno
        Sdb.Execute s
        Print #1, "close"
        Print #1, "bye"
        Close #1
    End If
    If addfile = True Then
        ftpexe = "c:\windows\system32\ftp.exe"
        x = Shell(ftpexe & " -s:" & ofile, vbNormalFocus)
        'MsgBox ftpexe & " -s:" & ofile
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "postro_bill(" & mplant & ", " & sdate & ")", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " postro_bill(" & mplant & ", " & sdate & ") - Error Number: " & eno
        End
    End If
End Sub

Private Sub testro_bill(mplant As String, sdate As String)
    Dim ofile As String, s As String, rfile As String, tktno As Long
    Dim i As Integer, k As Integer, addfile As Boolean, ftpexe As String
    Dim ds As adodb.Recordset
    Dim x, eno As Long, edesc As String
    ofile = "c:\jvwork\billtest.txt"
    Open ofile For Output As #1
    For i = 1 To billgrid.Rows - 1
        If billgrid.TextMatrix(i, 0) = "B" Then
            Write #1, billgrid.TextMatrix(i, 1);        'Batch
            Write #1, billgrid.TextMatrix(i, 2);        'Area
            Write #1, billgrid.TextMatrix(i, 3);        'Group
            Write #1, billgrid.TextMatrix(i, 4);        'Source
            Write #1, billgrid.TextMatrix(i, 5);        'Target
            Write #1, billgrid.TextMatrix(i, 6);        'Product
            Write #1, billgrid.TextMatrix(i, 7);        'Barcode
            Write #1, billgrid.TextMatrix(i, 8);        'Qty
            Write #1, billgrid.TextMatrix(i, 9);        'Uom
            Write #1, billgrid.TextMatrix(i, 10);        'Lot
            Write #1, billgrid.TextMatrix(i, 11);        'Units
            Write #1, billgrid.TextMatrix(i, 12);        'Lot2
            Write #1, billgrid.TextMatrix(i, 13);        'Units
            Write #1, "POSTED";                         'Status
            Write #1, billgrid.TextMatrix(i, 15);        'User
            Write #1, billgrid.TextMatrix(i, 16);        'DateTime
            Write #1, billgrid.TextMatrix(i, 17)         'Ticket
        End If
    Next i
    Close #1
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
    tktno = 0
    s = "select * from trailers where runid = " & r12tkt.Caption
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        rfile = "c:\jvwork\RO" & ds!runid & ".txt"
        tktno = ds!runid
        Open rfile For Append As #5
        For i = 1 To billgrid.Rows - 1
            If billgrid.TextMatrix(i, 0) = "B" Then
                If billgrid.TextMatrix(i, 7) > "00" Then  'barcode indicates pallet
                    Write #5, ds!runid & "P";
                    Write #5, morg;
                    Write #5, mwhs;
                    Write #5, "FLOOR" & mwhs;
                    If ds!branch = 47 Then
                        Write #5, "501"; "K10"; "FLOORK10";
                    Else
                        If ds!branch = 52 Then
                            Write #5, "502"; "A10"; "FLOORA10";
                        Else
                            If ds!branch = 1 Then
                                Write #5, "500"; "T10"; "FLOORT10";
                            Else
                                Write #5, Format(ds!branch, "000");
                                Write #5, Format(ds!branch, "000");
                                Write #5, "FLOOR" & Format(ds!branch, "000");
                            End If
                        End If
                    End If
                    Write #5, ds!account;
                    Write #5, Trim(Left(billgrid.TextMatrix(i, 7), 4));
                    Write #5, Trim(mid(billgrid.TextMatrix(i, 7), 5, 9));            'jv052515     'lot
                    Write #5, Format(Val(billgrid.TextMatrix(i, 11)), "0");
                    Write #5, "EACH";
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                    Write #5, StrConv(billgrid.TextMatrix(i, 5), vbProperCase) & " " & Right(billgrid.TextMatrix(i, 7), 3);
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy");
                    If Left(ds!trlno, 1) = "B" Or ds!branch = 15 Or ds!branch = 16 Then
                        Write #5, "Y"
                    Else
                        Write #5, "N"
                    End If
                        
                    If Val(billgrid.TextMatrix(i, 12)) > 0 Then   '2nd lot
                        f7 = billgrid.TextMatrix(i, 7)
                        f10 = billgrid.TextMatrix(i, 10)
                        f12 = billgrid.TextMatrix(i, 12)
                        s = mid(f7, 5, 2) & "-" & mid(f7, 7, 2) & "-20" & mid(f7, 9, 2)
                        s = Format(DateAdd("d", Val(Left(f12, 5)) - Val(f10), s), "MMddyy")     'jv081415
                        's = s & RTrim(mid(f7, 11, 3))                           'jv052515
                        s = s & RTrim(mid(f12, 6, 3))                           'jv083115
                        Write #5, ds!runid & "P";
                        Write #5, morg;
                        Write #5, mwhs;
                        Write #5, "FLOOR" & mwhs;
                        If ds!branch = 47 Then
                            Write #5, "501"; "K10"; "FLOORK10";
                        Else
                            If ds!branch = 52 Then
                                Write #5, "502"; "A10"; "FLOORA10";
                            Else
                                If ds!branch = 1 Then
                                    Write #5, "500"; "T10"; "FLOORT10";
                                Else
                                    Write #5, Format(ds!branch, "000");
                                    Write #5, Format(ds!branch, "000");
                                    Write #5, "FLOOR" & Format(ds!branch, "000");
                                End If
                            End If
                        End If
                        Write #5, ds!account;
                        Write #5, Trim(Left(billgrid.TextMatrix(i, 7), 4));
                        Write #5, s;
                        Write #5, Format(Val(billgrid.TextMatrix(i, 13)), "0");
                        Write #5, "EACH";
                        Write #5, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                        Write #5, StrConv(billgrid.TextMatrix(i, 5), vbProperCase) & " " & Right(f7, 3);
                        Write #5, Format(ds!shipdate, "MM-dd-yyyy");
                        If Left(ds!trlno, 1) = "B" Or ds!branch = 15 Or ds!branch = 16 Then
                            Write #5, "Y"
                        Else
                            Write #5, "N"
                        End If
                    End If
                Else
                    Write #5, ds!runid & "W";
                    If mplant = "50" Then Write #5, "001"; "001"; "FLOOR001";
                    If mplant = "51" Then Write #5, "047"; "047"; "FLOOR047";
                    If mplant = "52" Then Write #5, "052"; "052"; "FLOOR052";
                    Write #5, Format(ds!branch, "000");
                    Write #5, Format(ds!branch, "000");
                    Write #5, "FLOOR" & Format(ds!branch, "000");
                    Write #5, ds!account;
                    Write #5, Trim(Left(billgrid.TextMatrix(i, 6), 4));
                    Write #5, "LOT1"; 'billgrid.TextMatrix(i, 10);                                'jv082915
                    Write #5, Format(Val(billgrid.TextMatrix(i, 11)), "0");
                    Write #5, "EACH";
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy"); 'sdate;
                    Write #5, StrConv(billgrid.TextMatrix(i, 5), vbProperCase);
                    Write #5, Format(ds!shipdate, "MM-dd-yyyy");
                    If Left(ds!trlno, 1) = "B" Or ds!branch = 15 Or ds!branch = 16 Then
                        Write #5, "Y"
                    Else
                        Write #5, "N"
                    End If
                End If
            End If
        Next i
        Close #5
        ds.Close
    End If
    MsgBox "check: " & rfile
End Sub

Private Sub rename_trailer(runno As String)
    Dim ds As adodb.Recordset, s As String
    Dim obranch As String, otrlno As String, odate As String
    Dim nbranch As String, ntrlno As String, ndate As String, nbatch As String
    Dim bname As String, cfile As String, buildrun As Boolean, newrun As String
    Dim eno As Long, edesc As String, rundate As String
    On Error GoTo vberror
    rundate = Format(sd, "MM-dd-yyyy")                                          'jv111815
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
    
    s = "Ok to rename " & rundate & " " & billgrid.TextMatrix(billgrid.Row, 5)
    s = s & " to " & ndate & " " & bname & " " & ntrlno & "?"
    If MsgBox(s, vbYesNo + vbQuestion, "are you sure...") = vbNo Then Exit Sub
    
    If Form1.plantno <> 50 Then                                                                                 'jv111815
        gcode = "T" & mid(Format(ndate, "MM-dd-yyyy"), 4, 2) & Format(Val(nbranch), "00") & Right(ntrlno, 1)    'jv111815
    End If                                                                                                      'jv111815
    'Exit Sub
    
    nbatch = DateDiff("d", "1-1-2012", ndate) & Format(Val(nbranch), "00") & Right(ntrlno, 1)
    For i = 1 To billgrid.Rows - 1
        If billgrid.TextMatrix(i, 1) = nbatch Then
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
        s = "Update trailers set branch = " & Val(nbranch)
        s = s & ", account = '......'"
        s = s & ", shipdate = '" & ndate & "'"
        s = s & ", trlno = '" & ntrlno & "'"
        s = s & ", groupcode = '" & gcode.Caption & "'"                     'jv111815
        s = s & ", pb_flag = 'N'"
        s = s & ", ra_flag = 'N'"
        s = s & " Where runid = " & runno
        Sdb.Execute s
    Else
        buildrun = True
    End If
    ds.Close
    
    If buildrun = True Then
        s = "select * from runs where loaded = '" & Me.plantno & "'"
        s = s & " and destination = '" & nbranch & "'"
        s = s & " and trlno = '" & ntrlno & "'"
        s = s & " and startime = #" & ndate & "#"
        Set ds = Sdb.OpenRecordset(s)
        If ds.BOF = False Then
            ds.MoveFirst
            newrun = ds!id
        Else
            newrun = wd_seq("Oratkt", Form1.schdb)
            s = "Insert into runs (id, loaded, destination, locname, trlno, trlsize, trldate, startime, pickup, oc)"
            s = s & " Values (" & newrun
            s = s & ", " & Me.plantno
            s = s & ", " & nbranch
            s = s & ", '" & bname & "'"
            s = s & ", '" & ntrlno & "'"
            s = s & ", 32"
            s = s & ", '" & ndate & "'"
            s = s & ", '12:00 PM'"
            s = s & ", 'Swapped-" & billgrid.TextMatrix(billgrid.Row, 5) & "'"
            s = s & ", ' ')"
            Sdb.Execute s
        End If
        ds.Close
        For i = 1 To billgrid.Rows - 1
            If billgrid.TextMatrix(i, 17) = runno Then
                s = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno, sku"
                s = s & ", pallets, wraps, units, whs_num, pb_flag, ra_flag) Values (" & wd_seq("Trailers", Me.shipdb)
                s = s & ", " & newrun
                's = s & ", '" & billgrid.TextMatrix(i, 3) & "'"
                s = s & ", '" & gcode.Caption & "'"                             'jv111815
                s = s & ", " & Me.plantno
                s = s & ", " & nbranch
                s = s & ", '......'"
                s = s & ", '" & ndate & "'"
                s = s & ", '" & ntrlno & "'"
                s = s & ", '" & Left(billgrid.TextMatrix(i, 6), 3) & "'"
                If billgrid.TextMatrix(i, 7) > "00" Then
                    s = s & ", 1, 0"
                Else
                    s = s & ", 0, 0"
                End If
                s = s & ", " & Val(billgrid.TextMatrix(i, 11)) + Val(billgrid.TextMatrix(i, 13))
                s = s & ", 4"
                s = s & ", 'Y', 'N')"
                Sdb.Execute s
            End If
        Next i
    End If
    
    If newrun <> runno Then
        cfile = Form1.pallogs & "bill" & Format(ndate, "MMddyyyy") & ".txt"
        Open cfile For Append As #1
        For i = 1 To billgrid.Rows - 1
            If billgrid.TextMatrix(i, 17) = runno Then
                billgrid.TextMatrix(i, 3) = gcode.Caption                   'jv111815
                billgrid.TextMatrix(i, 5) = bname & " " & ntrlno             'jv111815
                Write #1, nbatch;
                Write #1, billgrid.TextMatrix(i, 2);
                Write #1, billgrid.TextMatrix(i, 3);
                Write #1, billgrid.TextMatrix(i, 4);
                Write #1, bname & " " & ntrlno;
                Write #1, billgrid.TextMatrix(i, 6);
                Write #1, billgrid.TextMatrix(i, 7);
                Write #1, billgrid.TextMatrix(i, 8);
                Write #1, billgrid.TextMatrix(i, 9);
                Write #1, billgrid.TextMatrix(i, 10);
                Write #1, billgrid.TextMatrix(i, 11);
                Write #1, billgrid.TextMatrix(i, 12);
                Write #1, billgrid.TextMatrix(i, 13);
                Write #1, "PEND";
                Write #1, billgrid.TextMatrix(i, 15);
                Write #1, billgrid.TextMatrix(i, 16);
                Write #1, newrun
            End If
        Next i
        Close #1
    Else
        For i = 1 To billgrid.Rows - 1
            If billgrid.TextMatrix(i, 17) = runno Then
                s = "B" & Chr(9)
                s = s & nbatch & Chr(9)
                s = s & billgrid.TextMatrix(i, 2) & Chr(9)
                's = s & billgrid.TextMatrix(i, 3) & Chr(9)
                s = s & gcode.Caption & Chr(9)                              'jv111815
                s = s & billgrid.TextMatrix(i, 4) & Chr(9)
                s = s & bname & " " & ntrlno & Chr(9)
                s = s & billgrid.TextMatrix(i, 6) & Chr(9)
                s = s & billgrid.TextMatrix(i, 7) & Chr(9)
                s = s & billgrid.TextMatrix(i, 8) & Chr(9)
                s = s & billgrid.TextMatrix(i, 9) & Chr(9)
                s = s & billgrid.TextMatrix(i, 10) & Chr(9)
                s = s & billgrid.TextMatrix(i, 11) & Chr(9)
                s = s & billgrid.TextMatrix(i, 12) & Chr(9)
                s = s & billgrid.TextMatrix(i, 13) & Chr(9)
                s = s & "PEND" & Chr(9)
                s = s & billgrid.TextMatrix(i, 15) & Chr(9)
                s = s & billgrid.TextMatrix(i, 16) & Chr(9)
                s = s & newrun
                billgrid.AddItem s
                cntlit.Caption = billgrid.Rows - 1 & " Records"
            End If
        Next i
        For i = 1 To billgrid.Rows - 1
            If billgrid.TextMatrix(i, 17) = runno And billgrid.TextMatrix(i, 1) <> nbatch Then
                billgrid.TextMatrix(i, 14) = "CANC"
            End If
        Next i
    End If
    Call save_bp_temp                                                       'jv111815
    'Call save_bills(runno)
    DoEvents                                                                'jv111815
    sd_Click                                                                'jv111815
    'Call refresh_billgrid(Text1)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "rename_trailer(" & runno & ")", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " rename_trailer(" & runno & ") - Error Number: " & eno
        End
    End If
End Sub

Private Sub duplex_bill_log(runno As String)
    Dim ds As adodb.Recordset, sqlx As String, s As String
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
    Dim eno As Long, edesc As String
    Dim sealno As String, dname1 As String, dname2 As String
    Dim insBy As String, specIns As String, freight As String
    Dim printerProperty As String
    
    On Error GoTo vberror
    pno = 1: jobtrail = False: pflag = False: wflag = False
    tc = "OC": sealno = 0: dname1 = " ": dname2 = " "
    For i = 1 To trkgrid.Rows - 1
        If trkgrid.TextMatrix(i, 0) = "Trailer Code" Then
            If trkgrid.TextMatrix(i, 1) > "0" Then tc = trkgrid.TextMatrix(i, 1)
        End If
        If trkgrid.TextMatrix(i, 0) = "Seal #" Then sealno = Val(trkgrid.TextMatrix(i, 1))
        If trkgrid.TextMatrix(i, 0) = "Driver" Then dname1 = trkgrid.TextMatrix(i, 1)
        If trkgrid.TextMatrix(i, 0) = "2nd Driver" Then dname2 = trkgrid.TextMatrix(i, 1)
        If trkgrid.TextMatrix(i, 0) = "Inspected By" Then insBy = trkgrid.TextMatrix(i, 1)
        If trkgrid.TextMatrix(i, 0) = "Special Instructions" Then specIns = trkgrid.TextMatrix(i, 1)
        If trkgrid.TextMatrix(i, 0) = "Freight" Then freight = trkgrid.TextMatrix(i, 1)
    Next i
    If tc = "OC" Then
        s = "No trailer code was entered.  OC will be used on the bill."
        If MsgBox(s, vbYesNo + vbQuestion, "specify Outside carrier???") = vbNo Then Exit Sub
    End If
    
    If Val(runno) = 0 Then Exit Sub
    Screen.MousePointer = 11
    
    printerProperty = "Duplex = 3"
    Printer.Duplex = 3
    printerProperty = "Orientation = 1"
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
    printerProperty = "FontName"
    Printer.FontName = "Arial"
    printerProperty = "FontSize"
    Printer.FontSize = 14
    printerProperty = "FontBold"
    Printer.FontBold = True
    printerProperty = "Print"
    Printer.Print Tab(32); " " '"B i l l   O f   L a d i n g"
    Printer.FontSize = 10
    Printer.FontBold = True
    printerProperty = "CurrentX"
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
        Printer.Print Format(bno, "00"); " "; ds!branchname; " "; Right(billgrid.TextMatrix(billgrid.Row, 5), 2)
        ltarget = ds!branchname & " " & Right(billgrid.TextMatrix(billgrid.Row, 5), 2)      'jv022811
    End If
    
    Printer.CurrentX = 1440 * 1.5: Printer.Print oaddr1; '"1101 S. Blue Bell Road";
    Printer.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        Printer.Print j1
    Else
        'Printer.Print ds!addr1
        Printer.Print addgrid.TextMatrix(addgrid.Row, 1)                'jv093015
    End If
    Printer.CurrentX = 1440 * 1.5: Printer.Print oaddr2; '"Brenham, Texas  77834-1807";
    Printer.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        Printer.Print j2
    Else
        'Printer.Print ds!addr2
        Printer.Print addgrid.TextMatrix(addgrid.Row, 2)                'jv093015
    End If
    Printer.CurrentX = 1440 * 1.5: Printer.Print ophone; '"(979) 836-7977";
    Printer.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        Printer.Print j3
    Else
        'Printer.Print ds!brphone
        Printer.Print addgrid.TextMatrix(addgrid.Row, 3)                'jv093015
    End If
    Printer.CurrentX = 1440 * 1.5: Printer.Print "Fax: " & ofax; '"Fax: (979) 830-7398";
    Printer.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        Printer.Print j4
    Else
        'Printer.Print "Fax: " & ds!brfax
        Printer.Print addgrid.TextMatrix(addgrid.Row, 4)                'jv093015
    End If
    Printer.Print String(130, "_")
    ds.Close
    tu = 0: tw = 0: tp = 0
    tmpgrid.Clear: tmpgrid.Rows = 1: tmpgrid.Cols = 5
    ppflag = False
    For i = 1 To billgrid.Rows - 1
        If billgrid.TextMatrix(i, 17) = runno And billgrid.TextMatrix(i, 14) <> "CANC" Then
            f1 = billgrid.TextMatrix(i, 1)
            f2 = billgrid.TextMatrix(i, 2)
            f3 = billgrid.TextMatrix(i, 3)
            f4 = billgrid.TextMatrix(i, 4)
            f5 = billgrid.TextMatrix(i, 5)
            f6 = billgrid.TextMatrix(i, 6)
            f7 = billgrid.TextMatrix(i, 7)
            f8 = billgrid.TextMatrix(i, 8)
            f9 = billgrid.TextMatrix(i, 9)
            f10 = billgrid.TextMatrix(i, 10)
            f11 = billgrid.TextMatrix(i, 11)
            f12 = billgrid.TextMatrix(i, 12)
            f13 = billgrid.TextMatrix(i, 13)
            f14 = billgrid.TextMatrix(i, 14)
            f15 = billgrid.TextMatrix(i, 15)
            f16 = billgrid.TextMatrix(i, 16)
            f17 = billgrid.TextMatrix(i, 17)
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
            If billgrid.TextMatrix(i, 14) <> "POSTED" Then billgrid.TextMatrix(i, 14) = "PRINTED"
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
            printerProperty = "TextWidth"
            Printer.CurrentX = p1 - Printer.TextWidth("Wraps")
            Printer.Print "Wraps";
        Else                                                            'jv100609
            printerProperty = "TextWidth"
            Printer.CurrentX = p1 - Printer.TextWidth("Pallets")        'jv100609
            Printer.Print "Pallets";                                    'jv100609
        End If
        printerProperty = "TextWidth"
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
                printerProperty = "NewPage"
                Printer.NewPage
                pno = pno + 1
                Printer.Print "Page "; pno;
                Printer.CurrentX = 8600: Printer.Print "Policy Number ";
                Printer.FontBold = True
                printerProperty = "FontUnderline"
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
                printerProperty = "NewPage"
                Printer.NewPage
                pno = pno + 1
                Printer.Print "Page "; pno;
                Printer.CurrentX = 8600: Printer.Print "Policy Number ";
                Printer.FontBold = True
                printerProperty = "FontUnderline"
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
    printerProperty = "CurrentY"
    Printer.CurrentY = 1440 * 9
    'For i = lc To 50 '54 '45 '50 '57
    '    printer.Print " "
    'Next i
    
    ' Old bill of lading format, pre 4/3/2020
    'Printer.CurrentX = 720: Printer.Print "Ship Date:";
    'Printer.CurrentX = 1440 * 1.5: Printer.Print sd; 'Text1; 'Edittrl.sd;
    'If tc = "OC" Then                                                       'jv080615
    '    Printer.CurrentX = 1440 * 2.5: Printer.Print "OC: " & dname1;       'jv080615
    'Else                                                                    'jv080615
    '    Printer.CurrentX = 1440 * 2.5: Printer.Print "Trailer #:";          'jv080615
    '    Printer.CurrentX = 1440 * 3.25: Printer.Print tc;                   'jv080615
    'End If                                                                  'jv080615
    'Printer.CurrentX = 1440 * 4.25: Printer.Print "Total Pallets:";         'jv080615
    'Printer.CurrentX = 1440 * 5.25: Printer.Print tp;  'Int(tp + 0.8)       'jv080615
    'Printer.CurrentX = 1440 * 6: Printer.Print ltarget                      'jv080615
    'Printer.Print " "
    'Printer.CurrentX = 720: Printer.Print "Inspected By:";                  'jv082415
    'Printer.CurrentX = 1440 * 1.5: Printer.Print "_____________________________";
    'Printer.CurrentX = 1440 * 4: Printer.Print "Completed By:";
    'Printer.CurrentX = 1440 * 5: Printer.Print "_____________________________"
    'Printer.Print " "
    'Printer.CurrentX = 720: Printer.Print "Seal #:";
    'Printer.CurrentX = 1440 * 1.5
    'If sealno > 0 Then
    '    Printer.Print sealno;
    'Else
    '    Printer.Print "_____________________________";
    'End If
    'Printer.CurrentX = 1440 * 4: Printer.Print "Sealed By:";
    'Printer.CurrentX = 1440 * 5: Printer.Print "_____________________________"
    'Printer.Print " "
    'If tc = "OC" Then
    '    Printer.CurrentX = 720: Printer.Print "Attention " & Trim(dname1) & " Driver: ";
    '    Printer.Print "Cargo must be kept at -20F."
    'Else
    '    Printer.CurrentX = 720: Printer.Print "Driver:";
    '    Printer.CurrentX = 1440 * 1.5
    '    If dname1 > " " Then
    '        Printer.Print dname1;
    '    Else
    '        Printer.Print "_____________________________";
    '    End If
    '    Printer.CurrentX = 1440 * 4: Printer.Print "Freight:";
    '    Printer.CurrentX = 1440 * 5: Printer.Print "_____________________________"
    'End If
    'Printer.Print " "
    'Printer.CurrentX = 720: Printer.Print "Special Instructions:";
    'Printer.CurrentX = 1440 * 2: Printer.Print "____________________________________________________________________"
    
    
    ' New bill of lading format
    ' Line 1: Destination @ 1440*5.5, Driver @ 1440*6
    ' NOTE FOR FUTURE DEVS: A Semicolon after the print tells it to not terminate the line after printing.
    '       Not having the semicolon will break to a newline on the next print operation.
    Printer.CurrentX = 720: Printer.Print ltarget;
    Printer.CurrentX = 1440 * 4: Printer.Print "Driver:";
    If dname1 > " " Then
        Printer.CurrentX = 1440 * 5.5: Printer.Print dname1;
    Else
        Printer.CurrentX = 1440 * 5.5: Printer.Print "_________________________";
    End If
    Printer.Print " "
    
    ' Line 2: Trailer @ 720, Ship Date @ 1440*6
    If tc = "OC" Then
        Printer.CurrentX = 720: Printer.Print "OC: "
        Printer.CurrentX = 1440 * 1.5: Printer.Print dname1;
    Else
        Printer.CurrentX = 720: Printer.Print "Trailer #: ";
        Printer.CurrentX = 1440 * 1.5: Printer.Print tc;
    End If
    Printer.CurrentX = 1440 * 4: Printer.Print "Ship Date:";
    Printer.CurrentX = 1440 * 5.5: Printer.Print sd;
    Printer.Print " "
    
    ' Line 3: Seal @ 720, Inspected By @ 1440*6
    Printer.CurrentX = 720: Printer.Print "Seal #:";
    If sealno > 0 Then
        Printer.CurrentX = 1440 * 1.5: Printer.Print sealno;
    Else
        Printer.CurrentX = 1440 * 1.5: Printer.Print "_________________________";
    End If
    Printer.CurrentX = 1440 * 4: Printer.Print "Inspected By:";
    If insBy > " " Then
        Printer.CurrentX = 1440 * 5.5: Printer.Print insBy;
    Else
        Printer.CurrentX = 1440 * 5.5: Printer.Print "_________________________";
    End If
    Printer.Print " "
    
    ' Line 4: Total Pallets @720, Special Instruction
    Printer.CurrentX = 720: Printer.Print "Total Pallets:";
    Printer.CurrentX = 1440 * 1.5: Printer.Print tp;
    Printer.CurrentX = 1440 * 4: Printer.Print "Special Instructions:";
    If specIns > " " Then
        Printer.CurrentX = 1440 * 5.5: Printer.Print specIns;
    Else
        Printer.CurrentX = 1440 * 5.5: Printer.Print "_________________________";
    End If
    Printer.Print " "
    
    ' Line 5: Freight @720 full width
    Printer.CurrentX = 720: Printer.Print "Freight:";
    If freight > " " Then
        Printer.CurrentX = 1440 * 1.5: Printer.Print freight;
    Else
        Printer.CurrentX = 1440 * 1.5: Printer.Print "____________________________________________________________________"
    End If
    Printer.Print " "
    ' End of new Bill of Lading format
    
    printerProperty = "NewPage"
    Printer.NewPage
    Call prtpage2(Printer, pflag, wflag, dname1, dname2, sealno)
    printerProperty = "EndDoc"
    Printer.EndDoc
    printerProperty = "Duplex = 1"
    Printer.Duplex = 1
        
    'Screen.MousePointer = 0
    'Exit Sub
    'Turn off for testing   jv022811
    'sqlx = "Update trailers set pb_flag = 'Y' where runid = " & runno
    'Sdb.Execute sqlx
    Open Form1.tempdir & "/billtrl.prn" For Append As #1
    If bno = "16" Or bno = "15" Then
        Write #1, sccode; Right(billgrid.TextMatrix(billgrid.Row, 5), 2); sd; tc
    Else
        Write #1, Format(bno, "00"); Right(billgrid.TextMatrix(billgrid.Row, 5), 2); sd; tc
    End If
    Close #1
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "duplex_bill_log(" & runno & ")", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " duplex_bill_log(" & runno & ") - Error Number: " & eno & vbLf & "Printer Property: " & printerProperty
        End
    End If
End Sub

Private Sub prtpage2(pd As Control, pallets As Boolean, wraps As Boolean, d01 As String, d02 As String, sno As String)
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
    pd.CurrentX = 1440 * 0.5: pd.Print "Driver Name";
    pd.CurrentX = 1440 * 2.5: pd.Print d01;
    pd.CurrentX = 1440 * 4.5: pd.Print d02
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Starting Location";
    pd.CurrentX = 1440 * 2.5
    If Me.plantno = "50" Then pd.Print "Brenham"
    If Me.plantno = "51" Then pd.Print "Broken Arrow"
    If Me.plantno = "52" Then pd.Print "Sylacauga"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Date";
    pd.CurrentX = 1440 * 2.5: pd.Print Me.sd
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Destination";
    pd.CurrentX = 1440 * 2.5: pd.Print Left(Combo1, Len(Combo1) - 2)
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Seal #";
    pd.CurrentX = 1440 * 2.5: pd.Print sno
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Depart temp."
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Mid trip temp."             'jv080615
    pd.Print " "                                                    'jv080615
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)                     'jv080615
    pd.Print " "                                                    'jv080615
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
        pd.Print billgrid.TextMatrix(billgrid.Row, 17) & "P " & billgrid.TextMatrix(billgrid.Row, 17) & "W"
    Else
        If pallets = True Then
            pd.Print billgrid.TextMatrix(billgrid.Row, 17) & "P"
        Else
            If wraps = True Then
                pd.Print billgrid.TextMatrix(billgrid.Row, 17) & "W"
            Else
                pd.Print billgrid.TextMatrix(billgrid.Row, 17)
            End If
        End If
    End If
End Sub

Private Sub fetch_r12_bill()
    Dim cfile As String, ofile As String, s As String
    Dim f1 As String, f2 As String, f3 As String, f4 As String, f5 As String
    Dim f6 As String, f7 As String, f8 As String, f9 As String, f10 As String
    Dim f11 As String, f12 As String, f13 As String, f14 As String, f15 As String
    Dim f16 As String, f17 As String, i As Integer
    Dim mplant As String, ldate1 As String, sdate As String, mbatch As String
    Dim pbranch As String, ctest As String, mfile As String, mprod As String
    Dim torg As String, twhs As String, tacct As String, psku As String, plot As String
    Dim ds As adodb.Recordset, sqls As String, tid As Long
    Dim ldate2 As String, cfile2 As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    mplant = Me.plantno
    ldate1 = ldate
    'ldate = InputBox("Load Date:", "Trailer Loaded Date....", Format(DateAdd("d", -1, sd), "mm-dd-yyyy"))    'jv022811
    If Len(ldate1) = 0 Then Exit Sub         'jv022811
    If Format(ldate1, "yyyyMMdd") > Format(Now, "yyyyMMdd") Then
        s = "The load date entered, " & ldate1 & ", cannot be greater than the current date."
        MsgBox s, vbOKOnly + vbInformation, "sorry, this date cannot be used..."
        Exit Sub                    'jv010313
    End If
    'Exit Sub
    ldate2 = Format(DateAdd("d", 1, ldate), "MMddyyyy")
    ldate1 = Format(ldate1, "MMddyyyy")
    
    sdate = Format(sd, "MMddyyyy")
    cfile = Me.tempdir & "\bp" & r12tkt.Caption & ".tmp"
    'MsgBox cfile
    If Len(Dir(cfile)) > 0 Then
        cfile2 = "none"
    Else
        cfile = Me.pallogs & "ship" & ldate1 & ".txt"
        cfile2 = Me.pallogs & "ship" & ldate2 & ".txt"
    End If
    mfile = Me.pallogs & "move" & sdate & ".txt"
    ofile = Me.pallogs & "bill" & sdate & ".txt"
    'If Len(Dir(cfile)) = 0 Then Exit Sub                       jv091415
    
    billgrid.Clear: billgrid.Rows = 1: billgrid.Cols = 18
    s = "select id,runid,trailers.branch,account,trlno,sku,pallets,wraps,units,branchname,shipdate"
    s = s & " from trailers, branches"
    s = s & " where runid = " & Left(List1, Len(List1) - 6)
    s = s & " and branches.branch = trailers.branch"
    s = s & " and plant = " & mplant
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
                                If LCase(f3) = LCase(gcode) And f2 = "DOCK" And Trim(Left(f6, 4)) = ds!sku Then         'jv082415
                                    s = "B" & Chr(9)              'Type
                                    s = s & ds!account & Chr(9)   'Recid
                                    s = s & f2 & Chr(9)           'Area
                                    s = s & f3 & Chr(9)           'Description
                                    s = s & f4 & Chr(9)           'Source
                                    s = s & Combo1 & Chr(9)       'Target
                                    s = s & f6 & Chr(9)           'Product
                                    s = s & f7 & Chr(9)           'Pallet
                                    s = s & f8 & Chr(9)           'Qty
                                    s = s & f9 & Chr(9)           'Uom
                                    s = s & f10 & Chr(9)          'lot
                                    s = s & f11 & Chr(9)          'units
                                    s = s & f12 & Chr(9)          'lot2
                                    s = s & f13 & Chr(9)          'units2
                                    's = s & "PEND" & Chr(9)       'status
                                    s = s & f14 & Chr(9)       'status
                                    s = s & f15 & Chr(9)          'user
                                    s = s & f16 & Chr(9)          'time
                                    s = s & ds!runid              'reqid
                                    billgrid.AddItem s
                                    cntlit.Caption = billgrid.Rows - 1 & " Records"
                                    logname.Caption = cfile
                                    tid = ds!id             'jv1112
                                End If
                            Loop
                            Close #2
                        End If
                        
                        If Len(Dir(cfile2)) > 0 Then
                            Open cfile2 For Input As #2     'After Midnight
                            Do Until EOF(2)
                                Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
                                If LCase(f3) = LCase(gcode) And f2 = "DOCK" And Trim(Left(f6, 4)) = ds!sku Then         'jv082415
                                    s = "B" & Chr(9)              'Type
                                    s = s & ds!account & Chr(9)   'Recid
                                    s = s & f2 & Chr(9)           'Area
                                    s = s & f3 & Chr(9)           'Description
                                    s = s & f4 & Chr(9)           'Source
                                    s = s & Combo1 & Chr(9)       'Target
                                    s = s & f6 & Chr(9)           'Product
                                    s = s & f7 & Chr(9)           'Pallet
                                    s = s & f8 & Chr(9)           'Qty
                                    s = s & f9 & Chr(9)           'Uom
                                    s = s & f10 & Chr(9)          'lot
                                    s = s & f11 & Chr(9)          'units
                                    s = s & f12 & Chr(9)          'lot2
                                    s = s & f13 & Chr(9)          'units2
                                    's = s & "PEND" & Chr(9)       'status
                                    s = s & f14 & Chr(9)       'status
                                    s = s & f15 & Chr(9)          'user
                                    s = s & f16 & Chr(9)          'time
                                    s = s & ds!runid              'reqid
                                    billgrid.AddItem s
                                    cntlit.Caption = billgrid.Rows - 1 & " Records"
                                    logname.Caption = cfile
                                    tid = ds!id             'jv1112
                                End If
                            Loop
                            Close #2
                        End If
                    End If
                    ctest = mcomm & ds!sku
                Else
                    s = "B" & Chr(9)                                  'Type
                    s = s & ds!account & Chr(9)                       'Recid
                    s = s & "JOBBING" & Chr(9)                        'Area
                    s = s & gcode.Caption & Chr(9)                    'Description
                    s = s & "ORDER PICK" & Chr(9)                     'Source
                    s = s & Combo1 & Chr(9)                           'Target
                    s = s & mprod & Chr(9)                            'Product
                    s = s & " " & Chr(9)                              'Palletid
                    s = s & ds!wraps & Chr(9)                         'qty
                    s = s & "Wraps" & Chr(9)                          'uom
                    s = s & "LOT1" & Chr(9)                           'lot
                    s = s & ds!units & Chr(9)                         'units
                    s = s & " " & Chr(9) & "0" & Chr(9)               'lot2 & qty2
                    s = s & "PEND" & Chr(9)                           'status
                    s = s & "wms" & Chr(9)                            'user
                    s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)   'time
                    s = s & ds!runid & Chr(9)                         'reqid
                    billgrid.AddItem s
                    cntlit.Caption = billgrid.Rows - 1 & " Records"
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
                                If LCase(f3) = LCase(gcode) And f2 = "DOCK" And LCase(f5) = LCase(mcomm) And Trim(Left(f6, 4)) = ds!sku Then  'jv111815
                                    s = "B" & Chr(9)          'Type
                                    s = s & mbatch & Chr(9)   'Recid
                                    s = s & f2 & Chr(9)       'Area
                                    s = s & f3 & Chr(9)       'Description
                                    s = s & f4 & Chr(9)       'Source
                                    s = s & f5 & Chr(9)       'Target
                                    s = s & f6 & Chr(9)       'Product
                                    s = s & f7 & Chr(9)       'Pallet
                                    s = s & f8 & Chr(9)       'Qty
                                    s = s & f9 & Chr(9)       'Uom
                                    s = s & f10 & Chr(9)      'lot
                                    s = s & f11 & Chr(9)      'units
                                    s = s & f12 & Chr(9)      'lot2
                                    s = s & f13 & Chr(9)      'units2
                                    's = s & "PEND" & Chr(9)   'status
                                    s = s & f14 & Chr(9)   'status
                                    s = s & f15 & Chr(9)      'user
                                    s = s & f16 & Chr(9)      'time
                                    s = s & ds!runid          'reqid
                                    billgrid.AddItem s
                                    cntlit.Caption = billgrid.Rows - 1 & " Records"
                                    logname.Caption = cfile
                                    tid = ds!id         'jv1112
                                End If
                            Loop
                            Close #2
                        End If
                        
                        If Len(Dir(cfile2)) > 0 Then
                            Open cfile2 For Input As #2     'After Midnight
                            Do Until EOF(2)
                                Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
                                If LCase(f3) = LCase(gcode) And f2 = "DOCK" And LCase(f5) = LCase(mcomm) And Trim(Left(f6, 4)) = ds!sku Then  'jv111815
                                    s = "B" & Chr(9)          'Type
                                    s = s & mbatch & Chr(9)   'Recid
                                    s = s & f2 & Chr(9)       'Area
                                    s = s & f3 & Chr(9)       'Description
                                    s = s & f4 & Chr(9)       'Source
                                    s = s & f5 & Chr(9)       'Target
                                    s = s & f6 & Chr(9)       'Product
                                    s = s & f7 & Chr(9)       'Pallet
                                    s = s & f8 & Chr(9)       'Qty
                                    s = s & f9 & Chr(9)       'Uom
                                    s = s & f10 & Chr(9)      'lot
                                    s = s & f11 & Chr(9)      'units
                                    s = s & f12 & Chr(9)      'lot2
                                    s = s & f13 & Chr(9)      'units2
                                    's = s & "PEND" & Chr(9)   'status
                                    s = s & f14 & Chr(9)   'status
                                    s = s & f15 & Chr(9)      'user
                                    s = s & f16 & Chr(9)      'time
                                    s = s & ds!runid          'reqid
                                    billgrid.AddItem s
                                    cntlit.Caption = billgrid.Rows - 1 & " Records"
                                    logname.Caption = cfile2
                                    tid = ds!id         'jv1112
                                End If
                            Loop
                            Close #2
                        End If
                    End If
                    ctest = mcomm & ds!sku
                Else
                    s = "B" & Chr(9)                                'Type
                    s = s & mbatch & Chr(9)                         'Recid
                    s = s & "PARTIAL" & Chr(9)                      'Area
                    s = s & gcode.Caption & Chr(9)                  'Description
                    s = s & "ORDER PICK" & Chr(9)                   'Source
                    s = s & Combo1 & Chr(9)                         'Target
                    s = s & mprod & Chr(9)                          'Product
                    s = s & " " & Chr(9)                            'Palletid
                    s = s & ds!wraps & Chr(9)                       'qty
                    s = s & "Wraps" & Chr(9)                        'uom
                    s = s & "LOT1" & Chr(9)                         'lot
                    s = s & ds!units & Chr(9)                       'units
                    s = s & " " & Chr(9) & "0" & Chr(9)             'lot2 & qty2
                    s = s & "PEND" & Chr(9)                         'status
                    s = s & "wms" & Chr(9)                          'user
                    s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
                    s = s & ds!runid                                'reqid
                    billgrid.AddItem s
                    cntlit.Caption = billgrid.Rows - 1 & " Records"
                    tid = ds!id                     'jv1112
                End If
            End If
            'turn off for testing
            If tid > 0 Then
                sqls = "update trailers set pb_flag = 'Y'"
                'sqls = sqls & " where runid = " & ds!runid
                sqls = sqls & " where id = " & tid          'jv1112
                'sdb.Execute sqls
            End If
            ds.MoveNext
        Loop
    End If
    Close #1
    ds.Close
    Call Check1_Click
    billgrid.RowSel = billgrid.Row
    billgrid.Col = 16: billgrid.ColSel = 16
    billgrid.Sort = 5
    refresh_grid
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "fetch_r12_bill", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " fetch_r12_bill - Error Number: " & eno
        End
    End If
End Sub

Private Sub ckoffsheet2011()
    Dim prun As String, ss As adodb.Recordset
    Dim sqlx As String, i As Integer, pdesc As String
    Dim ds2 As adodb.Recordset
    Dim tpals As Integer, twrps As Integer
    Dim eno As Long, edesc As String
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
            pdesc = skurec(Val(ds2!sku)).unit & " " & skurec(Val(ds2!sku)).desc
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
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "ckoffsheet2011", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " ckoffsheet2011 - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String, i As Integer, pfile As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    td.Rows = 1: pc.Clear: wc.Clear: tid.Clear: Label9.Visible = False
    If Val(r12tkt) = 0 Then Exit Sub
    sqlx = "Select ID,trailers.sku,fgunit,fgdesc,pallets,wraps,units,pallet,"
    sqlx = sqlx & "numwrap,branch,account,plant,trailers.whs_num,groupcode,pb_flag"
    sqlx = sqlx & " from trailers,skumast"
    sqlx = sqlx & " Where runid = " & r12tkt 'Left$(List1, Len(List1) - 6)
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
    refresh_trkgrid (r12tkt)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid - Error Number: " & eno
        End
    End If
    Exit Sub
End Sub

Private Sub sywhs(w As String)
    Dim sqlx As String
    Dim eno As Long, edesc As String
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
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "sywhs(" & w & ")", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " sywhs(" & w & ") - Error Number: " & eno
        End
    End If
End Sub

Private Sub update_trl()
    Dim sqlx As String
    Dim eno As Long, edesc As String
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
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "update_trl", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " update_trl - Error Number: " & eno
        End
    End If
End Sub

Private Sub addbc_Click()
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String, i As Integer
    Dim ds As adodb.Recordset
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
    logpath = Me.pallogs
    If Val(billgrid.TextMatrix(billgrid.Row, 17)) < 1 Then Exit Sub
    wbc = billgrid.TextMatrix(billgrid.Row, 7)
    wbc = InputBox("Enter a BarCode to search for:", "BarCode Example....", wbc)
    If Len(wbc) = 0 Then Exit Sub
    wbc = UCase(wbc)
    For i = 1 To billgrid.Rows - 1
        If wbc = billgrid.TextMatrix(i, 7) And billgrid.TextMatrix(i, 14) <> "CANC" Then
            MsgBox "BarCode is already on this bill.", vbOKOnly + vbInformation, "Duplicate barcode..."
            Exit Sub
        End If
    Next i
    Screen.MousePointer = 11
    t10 = "0"
    On Error GoTo vberror
    s = "Select * from pallets where barcode = '" & wbc & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        t5 = ds!sku
        t6 = ds!BarCode
        t7 = "1"
        t8 = "Pallet"
        t9 = ds!lot1
        t10 = ds!qty1
        t11 = ds!lot2
        t12 = ds!qty2
        MsgBox t6 & " found in pallet table.."
        ds.Close
        s = "select description, uom_type from sku_config where sku = '" & t5 & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            t5 = t5 & " " & ds!uom_type & " " & ds!description
        End If
    End If
    ds.Close
    
    sdate = Format(DateAdd("d", -1, sd), "MM-dd-yyyy")
    edate = Format(sd, "MM-dd-yyyy")
    sdate = Format(sdate, "yyyymmdd")
    edate = Format(edate, "yyyymmdd")
    If Val(t10) = 0 Then
        'Look for barcode in movement log
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
    End If
    
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
        i = billgrid.Row
        s = "B" & Chr(9)
        s = s & billgrid.TextMatrix(i, 1) & Chr(9)
        s = s & billgrid.TextMatrix(i, 2) & Chr(9)
        s = s & billgrid.TextMatrix(i, 3) & Chr(9)
        s = s & "ADD" & Chr(9) 'billgrid.TextMatrix(i, 4) & Chr(9)
        s = s & billgrid.TextMatrix(i, 5) & Chr(9)
        s = s & t5 & Chr(9)
        s = s & t6 & Chr(9)
        s = s & t7 & Chr(9)
        s = s & t8 & Chr(9)
        s = s & t9 & Chr(9)
        s = s & t10 & Chr(9)
        s = s & t11 & Chr(9)
        s = s & t12 & Chr(9)
        s = s & "PEND" & Chr(9) 'billgrid.TextMatrix(i, 14) & Chr(9)
        's = s & "wms" & Chr(9) 'billgrid.TextMatrix(i, 15) & Chr(9)
        s = s & Form1.userid & Chr(9) 'billgrid.TextMatrix(i, 15) & Chr(9)
        s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9) 'billgrid.TextMatrix(i, 16) & Chr(9)
        s = s & billgrid.TextMatrix(i, 17) & Chr(9)
        billgrid.AddItem s, i
        cntlit.Caption = billgrid.Rows - 1 & " Records"
        srun = billgrid.TextMatrix(i, 17)
        Call check_totals(r12tkt)
        billgrid.Row = i
        cfile = Me.pallogs & "wms" & Format(ldate.Text, "MMddyyyy") & ".txt"            'jv062615
        Open cfile For Append As #1                                                     'jv062615
        For k = 1 To 16                                                                 'jv062615
            Write #1, billgrid.TextMatrix(i, k);                                        'jv062615
        Next k                                                                          'jv062615
        Write #1, billgrid.TextMatrix(i, 17)                                            'jv062615
        Close #1                                                                        'jv062615
        save_bp_temp
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "addbc_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " addbc_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub addtline_Click()
    Command3_Click
End Sub

Private Sub addwraps_Click()
    Dim ds As adodb.Recordset, s As String
    Dim wqty As Integer, uqty As Integer
    Dim sprod As String, wconv As Integer, srun As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    If Val(billgrid.TextMatrix(billgrid.Row, 17)) = 0 Then Exit Sub
    wconv = 0
    s = Trim(Left(billgrid.TextMatrix(billgrid.Row, 6), 4))
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
    s = billgrid.TextMatrix(billgrid.Row, 0)
    s = s & Chr(9) & billgrid.TextMatrix(billgrid.Row, 1)
    s = s & Chr(9) & billgrid.TextMatrix(billgrid.Row, 2)
    s = s & Chr(9) & billgrid.TextMatrix(billgrid.Row, 3)
    s = s & Chr(9) & "ORDER PICK"
    s = s & Chr(9) & billgrid.TextMatrix(billgrid.Row, 5)
    s = s & Chr(9) & sprod & Chr(9) & " " & Chr(9)
    s = s & wqty & Chr(9)
    s = s & "Wraps" & Chr(9)
    s = s & "LOT1" & Chr(9)
    s = s & Format(wconv * wqty, "0") & Chr(9)
    s = s & ".." & Chr(9) & "0" & Chr(9) & "PEND" & Chr(9)
    's = s & "wms" & Chr(9) & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
    s = s & Form1.userid & Chr(9) & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
    s = s & billgrid.TextMatrix(billgrid.Row, 17)
    i = billgrid.Row
    billgrid.AddItem s, billgrid.Row
    cntlit.Caption = billgrid.Rows - 1 & " Records"
    srun = billgrid.TextMatrix(billgrid.Row, 17)
    Call check_totals(srun)
    billgrid.Row = i
    save_bp_temp
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "addwraps_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " addwraps_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub billgrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edscans
End Sub

Private Sub canline_Click()
    Dim i As Integer, srun As String
    If billgrid.TextMatrix(billgrid.Row, 14) = "POSTED" Then      'JV010313
        MsgBox "This line has been POSTED.", vbOKOnly + vbInformation, "Cancel is denied.."
        Exit Sub
    End If
    i = billgrid.Row
    billgrid.TextMatrix(billgrid.Row, 8) = "-" & billgrid.TextMatrix(billgrid.Row, 8)           'jv081915
    billgrid.TextMatrix(billgrid.Row, 11) = "-" & billgrid.TextMatrix(billgrid.Row, 11)         'jv081915
    If Val(billgrid.TextMatrix(billgrid.Row, 13)) > 0 Then                                      'jv081915
        billgrid.TextMatrix(billgrid.Row, 13) = "-" & billgrid.TextMatrix(billgrid.Row, 13)     'jv081915
    End If                                                                                      'jv081915
    billgrid.TextMatrix(billgrid.Row, 14) = "CANC"
    billgrid.TextMatrix(billgrid.Row, 15) = Form1.userid
    srun = billgrid.TextMatrix(billgrid.Row, 17)
    Call check_totals(srun)
    billgrid.Row = i
    cfile = Me.pallogs & "wms" & Format(ldate.Text, "MMddyyyy") & ".txt"            'jv062615
    Open cfile For Append As #1                                                     'jv062615
    For k = 1 To 16                                                                 'jv062615
        Write #1, billgrid.TextMatrix(i, k);                                        'jv062615
    Next k                                                                          'jv062615
    Write #1, billgrid.TextMatrix(i, 17)                                            'jv062615
    Close #1                                                                        'jv062615
    save_bp_temp
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        s = "^Type|^Batch|^Area|^Group|^Source|<Target|<Product|^BarCode|^Qty|^Uom|^Lot|^Units|^Lot|^Units|^Status|^User|^DateTime|^Ticket"
        billgrid.FormatString = s
        billgrid.ColWidth(0) = 600
        billgrid.ColWidth(1) = 1200
        billgrid.ColWidth(2) = 1000
        billgrid.ColWidth(3) = 1000
        billgrid.ColWidth(4) = 1300
        billgrid.ColWidth(5) = 3000
        billgrid.ColWidth(6) = 3000
        billgrid.ColWidth(7) = 1800
        billgrid.ColWidth(8) = 800
        billgrid.ColWidth(9) = 800
        billgrid.ColWidth(10) = 800
        billgrid.ColWidth(11) = 800
        billgrid.ColWidth(12) = 800
        billgrid.ColWidth(13) = 800
        billgrid.ColWidth(14) = 1000
        billgrid.ColWidth(15) = 1000
        billgrid.ColWidth(16) = 1800
        billgrid.ColWidth(17) = 1000
    Else
        s = "^Type||||^Source||<Product|^BarCode|^Qty|^Uom|^Lot|^Units|^Lot|^Units||^User|^DateTime|"
        billgrid.FormatString = s
        billgrid.ColWidth(0) = 600
        billgrid.ColWidth(1) = 0 '1000
        billgrid.ColWidth(2) = 0 '1000
        billgrid.ColWidth(3) = 0 '1000
        billgrid.ColWidth(4) = 1300
        billgrid.ColWidth(5) = 0 '3000
        billgrid.ColWidth(6) = 3000
        billgrid.ColWidth(7) = 1600
        billgrid.ColWidth(8) = 800
        billgrid.ColWidth(9) = 800
        billgrid.ColWidth(10) = 800
        billgrid.ColWidth(11) = 800
        billgrid.ColWidth(12) = 800
        billgrid.ColWidth(13) = 800
        billgrid.ColWidth(14) = 0 '1000
        billgrid.ColWidth(15) = 800
        billgrid.ColWidth(16) = 1600
        billgrid.ColWidth(17) = 0 '1000
    End If
End Sub

Private Sub Combo1_Click()
    Dim i As Integer, t As String
    List1.ListIndex = Combo1.ListIndex
    addgrid.Redraw = False
    For i = 1 To addgrid.Rows - 1
        t = UCase(Trim(Left(Combo1, Len(Combo1) - 2)))
        If UCase(addgrid.TextMatrix(i, 0)) = t Then
            addgrid.Row = i '- 1
            addgrid.TopRow = addgrid.Row
            Exit For
        End If
    Next i
    addgrid.Redraw = True
End Sub

Private Sub Command1_Click()
    Grid2.Visible = False: ycolor.Visible = False
    Call fetch_r12_bill
    DoEvents
End Sub

Private Sub Command1_GotFocus()
    Command1.FontBold = True
End Sub

Private Sub Command1_LostFocus()
    Command1.FontBold = False
End Sub

Private Sub Command2_Click()
    Dim sqlx As String
    Dim eno As Long, edesc As String
    If postlit.Visible = True Then
        Command6.Visible = False                                            'jv071918
        AutoRelease.Visible = False
        MsgBox postlit.Caption, vbOKOnly + vbInformation, "sorry, edits are disabled..."
        Exit Sub
    End If
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
        If billgrid.Rows > 1 Then check_totals (r12tkt)
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command2_click", Form1.userid)
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

Private Sub Command3_Click()
    Dim ds As adodb.Recordset, sqlx As String, msku As String, mrun As Long
    Dim mgroup As String, mplant As Integer, mbranch As Integer, maccount As String
    Dim i As Integer, mtno As String, zid As Long
    Dim eno As Long, edesc As String
    If postlit.Visible = True Then
        Command6.Visible = False                                            'jv071918
        AutoRelease.Visible = False
        MsgBox postlit.Caption, vbOKOnly + vbInformation, "sorry, edits are disabled..."
        Exit Sub
    End If
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
    tid.ListIndex = td.Row - 1 'ListIndex - 1
    sqlx = "select * from trailers where id = " & tid
    Set ds = Sdb.Execute(sqlx)
    ds.MoveFirst
    mrun = ds!runid
    mgroup = ds!groupcode
    mplant = ds!plant
    mbranch = ds!branch
    maccount = ds!account
    mtno = ds!trlno
    zid = wd_seq("Trailers", Me.shipdb)
    sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno, sku, pallets"
    sqlx = sqlx & ",wraps, units, whs_num, pb_flag, ra_flag) Values (" & zid
    sqlx = sqlx & ", " & mrun
    sqlx = sqlx & ", '" & mgroup & "'"
    sqlx = sqlx & ", " & mplant
    sqlx = sqlx & ", " & mbranch
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
    sqlx = sqlx & ",'N', 'N')"
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
    If billgrid.Rows > 1 Then check_totals (r12tkt)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command3_click", Form1.userid)
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

Private Sub Command4_Click()
    Call ckoffsheet2011
End Sub

Private Sub Command4_GotFocus()
    Command4.FontBold = True
End Sub

Private Sub Command4_LostFocus()
    Command4.FontBold = False
End Sub

Private Sub Command5_Click()
    'save_bp_temp
    'addbc_Click
    'canline_Click
    'addwraps_Click
    'edunits_Click
    Dim i As Integer, k As Integer, cfile As String, logpath As String
    If ycolor.Visible = True Then
        MsgBox ycolor.Caption, vbOKOnly + vbExclamation, "sorry, try again..."
        Exit Sub
    End If
    logpath = Me.pallogs
    'MsgBox "stop"
    'Exit Sub
    Screen.MousePointer = 11
    Call fetch_r12_bill
    DoEvents
    Call postro_bill(Form1.plantno, Format(sd, "MMddyyyy"))
    Screen.MousePointer = 0
End Sub

Private Sub Command6_Click()
    Dim cfile As String, i As Integer, s As String, answer As String
    Dim db5 As adodb.Connection, ds As adodb.Recordset
    Set db5 = CreateObject("ADODB.Connection")
    db5.Open "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=sywms;PWD=!Sylacauga_WMS1907"
    s = "select * from tBBCOrders where sBBCOrderNo = '" & gcode.Caption & "'"
    'MsgBox s
    Set ds = db5.Execute(s)
    If ds.BOF = False Then 'check if entry has already been posted
        ds.MoveFirst
        MsgBox "This order has already been posted.", vbOKOnly + vbExclamation, gcode.Caption
    Else
        For i = 1 To td.Rows - 1
            If Val(td.TextMatrix(i, 2)) > 0 And td.TextMatrix(i, 5) = "Crane" Then
                s = "Insert into tBBCOrders (sItemID, iQuantity, iPalletType, sLotID, sBBCOrderNo, bAutoRelease)" 'format insert statement
                s = s & " Values ('" & td.TextMatrix(i, 0) & "', "
                s = s & Val(td.TextMatrix(i, 2)) & ", 1, ' ', '" & gcode.Caption & "', " & Val(AutoRelease) & ")"
                db5.Execute s 'exectue insert statement
                'MsgBox s
            End If
        Next i
        MsgBox "Orders have been posted."
    End If
    ds.Close: db5.Close
End Sub

Private Sub deltline_Click()
    Command2_Click
End Sub
Private Sub edins_Click()
    Dim ib As String
    Dim tkwo As Long
    ib = ""
    On Error GoTo vberror
    For i = 0 To trkgrid.Rows - 1
        If trkgrid.TextMatrix(i, 0) = "WO Number" Then tkwo = Val(trkgrid.TextMatrix(i, 1))
        If trkgrid.TextMatrix(i, 0) = "Inspected By" Then ib = trkgrid.TextMatrix(i, 1)
    Next i
    ib = InputBox("Inspected By", "Inspected By...", ib)
    
    For i = 0 To trkgrid.Rows - 1
        If trkgrid.TextMatrix(i, 0) = "Inspected By" Then trkgrid.TextMatrix(i, 1) = ib
    Next i
    
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.schdb
    s = "update truckwo set InspectedBy = '" & ib & "' where wonum = " & tkwo
    db.Execute s
    db.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "edins_Click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " edins_Click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edfr_Click()
    Dim freight As String
    Dim tkwo As Long
    freight = ""
    On Error GoTo vberror
    For i = 0 To trkgrid.Rows - 1
        If trkgrid.TextMatrix(i, 0) = "Freight" Then freight = trkgrid.TextMatrix(i, 1)
    Next i
    
    freight = InputBox("Freight", "Freight...", freight)
    
    For i = 0 To trkgrid.Rows - 1
        If trkgrid.TextMatrix(i, 0) = "WO Number" Then tkwo = Val(trkgrid.TextMatrix(i, 1))
        If trkgrid.TextMatrix(i, 0) = "Freight" Then trkgrid.TextMatrix(i, 1) = freight
    Next i
    
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.schdb
    s = "update truckwo set Freight = '" & freight & "' where wonum = " & tkwo
    db.Execute s
    db.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "edfr_Click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " edfr_Click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edspec_Click()
    Dim specialIns As String
    Dim tkwo As Long
    specialIns = ""
    On Error GoTo vberror
    For i = 0 To trkgrid.Rows - 1
        If trkgrid.TextMatrix(i, 0) = "Special Instructions" Then specialIns = trkgrid.TextMatrix(i, 1)
    Next i
    
    specialIns = InputBox("Special Instructions", "Special Instructions...", specialIns)
    
    For i = 0 To trkgrid.Rows - 1
        If trkgrid.TextMatrix(i, 0) = "WO Number" Then tkwo = Val(trkgrid.TextMatrix(i, 1))
        If trkgrid.TextMatrix(i, 0) = "Special Instructions" Then trkgrid.TextMatrix(i, 1) = specialIns
    Next i
    
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.schdb
    s = "update truckwo set SpecialInstructions = '" & specialIns & "' where wonum = " & tkwo
    db.Execute s
    db.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "edspec_Click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " edspec_Click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edseal_Click()
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String, tc As String
    Dim i As Integer, tkwo As Long
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    tc = 0: tkwo = 0
    For i = 0 To trkgrid.Rows - 1
        If trkgrid.TextMatrix(i, 0) = "WO Number" Then tkwo = Val(trkgrid.TextMatrix(i, 1))
        If trkgrid.TextMatrix(i, 0) = "Seal #" Then tc = Val(trkgrid.TextMatrix(i, 1))
    Next i

    If tkwo = 0 Then Exit Sub
    tc = InputBox("BlueBell Seal #:", "Seal #....", tc)
    If Len(tc) = 0 Then Exit Sub
    
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.schdb
    s = "update truckwo set sealnum = " & Val(tc) & " where wonum = " & tkwo
    db.Execute s
    db.Close
    refresh_trkgrid (r12tkt)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "edseal_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " edseal_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edtc_Click()
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String, tc As String
    Dim i As Integer, tkwo As Long
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    tc = "OC": tkwo = 0
    For i = 0 To trkgrid.Rows - 1
        If trkgrid.TextMatrix(i, 0) = "WO Number" Then tkwo = Val(trkgrid.TextMatrix(i, 1))
        If trkgrid.TextMatrix(i, 0) = "Trailer Code" Then tc = trkgrid.TextMatrix(i, 1)
    Next i
    If tc = "OC" Then Exit Sub
    tc = InputBox("BlueBell Trailer Code or 'OC' for Outside Carrier:", "Trailer Code....", tc)
    If Len(tc) = 0 Then Exit Sub
    
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.schdb
    s = "Select listdisplay from valuelists where listname = 'trlcode' and listreturn = '" & tc & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update truckwo set eqnum = '" & tc & "' Where wonum = " & tkwo & " or parentwo = " & tkwo
        db.Execute s
        If bno <> "16" And bno <> "15" Then
            s = "Update trailertrack set lastwo = " & tkwo
            s = s & ", loaddate = '" & sd & "'"
            If plantno = "50" Then s = s & ", loadorigin = 'T10'"
            If plantno = "51" Then s = s & ", loadorigin = 'K10'"
            If plantno = "52" Then s = s & ", loadorigin = 'A10'"
            s = s & ", loaddestination = '" & Format(Val(bno), "000") & "'"
            s = s & " Where bbtrlnum = '" & tc & "'"
            db.Execute s
        End If
    Else
        MsgBox "Invalid trailer code: " & tc, vbOKOnly + vbExclamation, "Invalid code..."
    End If
    ds.Close: db.Close
    refresh_trkgrid (r12tkt)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "edtc_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " edtc_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edunits_Click()
    Dim s As String, i As Integer
    i = billgrid.Row
    If Val(billgrid.TextMatrix(i, 17)) = 0 Then Exit Sub
    s = InputBox("Units for w/d lot " & billgrid.TextMatrix(i, 10), "1st Lot..", billgrid.TextMatrix(i, 11))
    If Len(s) <> 0 Then billgrid.TextMatrix(i, 11) = s
    If Val(billgrid.TextMatrix(i, 13)) <> 0 Then
        s = InputBox("Units for w/d lot " & billgrid.TextMatrix(i, 12), "2nd Lot..", billgrid.TextMatrix(i, 13))
        If Len(s) <> 0 Then billgrid.TextMatrix(i, 11) = s
    End If
    srun = billgrid.TextMatrix(i, 17)
    Call check_totals(billgrid.TextMatrix(i, 17))
    billgrid.Row = i
    save_bp_temp
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If trailbill.ActiveControl.Name = "td" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command3_Click 'insert
        If KeyCode = 46 Or KeyCode = 120 Then Call Command2_Click 'delete
    End If
End Sub

Private Sub Form_Load()
    Dim ds As adodb.Recordset
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    Me.shipdb = Form1.shipdb
    Me.pallogs = Form1.pallogs
    Me.tempdir = Form1.tempdir
    Me.plantno = Form1.plantno
    If Form1.userid = "jvierus" Then debmenu.Enabled = True
    
    Set ds = Sdb.Execute("select distinct shipdate from trailers order by shipdate")
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sd.AddItem ds(0)
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Me.plantno = "52" Then
        td.Cols = 6
        td.FormatString = "^SKU|^Product|^Pallets|^Wraps|^Units|^Source"
        td.ColWidth(0) = 500
        td.ColWidth(1) = 3500: td.ColWidth(2) = 600
        td.ColWidth(3) = 600: td.ColWidth(4) = 600
        td.ColWidth(5) = 600
    Else
        td.Cols = 6: td.FixedCols = 1
        td.FormatString = "^SKU|^Product|^Pallets|^Wraps|^Units"
        td.ColWidth(0) = 600
        td.ColWidth(1) = 3500: td.ColWidth(2) = 800
        td.ColWidth(3) = 800: td.ColWidth(4) = 800
        td.ColWidth(5) = 1
    End If
    If sd.ListCount > 0 Then sd.ListIndex = 0
    refresh_addgrid
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "form_load", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " form_load - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Resize()
    billgrid.Width = Me.Width - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If UCase(Right(Form1.Caption, 6)) = "ETONLY" Then End
End Sub

Private Sub List1_Click()
    r12tkt = Left(List1, Len(List1) - 6)
    DoEvents
    Call refresh_grid
    If td.Rows > 1 Then
        td.Row = 1: Call td_Click
    End If
    Command1_Click
    DoEvents
    Call check_totals(r12tkt)
End Sub

Private Sub plantno_Change()
    If Val(plantno.Caption) = 0 Then Exit Sub
    If Val(plantno.Caption) = 50 Then
        Command4.Visible = False
    Else
        Command4.Visible = True
    End If
End Sub

Private Sub postr12_Click()
    Dim i As Integer, k As Integer, cfile As String, logpath As String
    If ycolor.Visible = True Then
        MsgBox ycolor.Caption, vbOKOnly + vbExclamation, "sorry, try again..."
        Exit Sub
    End If
    logpath = Me.pallogs
    'MsgBox "stop"
    'Exit Sub
    Screen.MousePointer = 11
    Call fetch_r12_bill
    DoEvents
    Call postro_bill(Form1.plantno, Format(sd, "MMddyyyy"))
    
    save_bp_temp
    Call r12tkt_Change
    List1_Click
    Screen.MousePointer = 0
End Sub

Private Sub printbill_Click()
    Dim onHold() As String
    ReDim onHold(0) As String
    Dim msg As String
    If Val(billgrid.TextMatrix(billgrid.Row, 17)) = 0 Then Exit Sub
    If ycolor.Visible = True Then
        MsgBox ycolor.Caption, vbOKOnly + vbExclamation, "sorry, try again..."
        Exit Sub
    End If
    
    ' Exit if bill contains duplicate barcodes. This flag set in check_totals function.
    If duplicateBarcodes.Visible = True Then
        MsgBox duplicateBarcodes.Caption, vbOKOnly + vbExclamation, "sorry, try again..."
        Exit Sub
    End If
    
    ' Do final hold check on contents of the trailer
    Call TrailerHoldCheck(onHold)
    If UBound(onHold) > 0 Then
        msg = "The following pallets are on HOLD: "
        For i = 1 To UBound(onHold)
            msg = msg & vbCrLf & onHold(i)
        Next i
        msg = msg & vbCrLf & "This trailer cannot be shipped, so BOL will not be printed."
        MsgBox msg, vbOKOnly, "PALLETS ON HOLD"
        Exit Sub
    End If
    
    'Testing = mark as printed
    s = billgrid.TextMatrix(billgrid.Row, 17)
    billgrid.FillStyle = flexFillRepeat
    billgrid.Redraw = False
    For i = 1 To billgrid.Rows - 1
        If billgrid.TextMatrix(i, 17) = s Then
            If billgrid.TextMatrix(i, 14) <> "POSTED" And billgrid.TextMatrix(i, 14) <> "CANC" Then
                billgrid.TextMatrix(i, 14) = "PRINTED"
                billgrid.Row = i: billgrid.RowSel = i
                billgrid.Col = 14: billgrid.ColSel = 14
                billgrid.CellForeColor = gcolor.ForeColor
                billgrid.CellBackColor = gcolor.BackColor
            End If
        End If
    Next i
    billgrid.Redraw = True
    'Live
    Call duplex_bill_log(billgrid.TextMatrix(billgrid.Row, 17))
End Sub

' Reece added this as a final hold check before the BOL is printed.
Private Sub TrailerHoldCheck(ByRef onHold() As String)
    Dim BarCode, LotOne, LotTwo, query As String ' Declare holding variables
    Dim ds As adodb.Recordset ' Declare recordset variable
    Dim ctr As Integer ' Declare counter for how many items have been found that are on hold
    ctr = 0 ' Initialize counter
    
    ' Loop over all items on the trailer
    For i = 1 To billgrid.Rows - 1
        ' Get the relevant values
        BarCode = billgrid.TextMatrix(i, 7)
        
        ' Build the query
        query = "EXEC cspTrailerHoldCheck '" & BarCode & "'"
        ' Run the query
        Set ds = Wdb.Execute(query)
        
        If ds.BOF = False Then
            ' Pallet is on hold, add to array
            ctr = ctr + 1
            ReDim Preserve onHold(ctr)
            onHold(ctr) = BarCode
        End If
    Next i
End Sub

Private Sub prtblank_Click()
    blnkbill.Show
End Sub

Private Sub r12tkt_Change()
    postlit.Visible = False
    edtrls.Enabled = True
    edscans.Enabled = True
    postmenu.Enabled = True
    renmenu.Enabled = True
    addtline.Enabled = True
    deltline.Enabled = True
    canline.Enabled = True
    addbc.Enabled = True
    addwraps.Enabled = True
    edunits.Enabled = True
    If Form1.plantno = "52" Then
        Command6.Visible = True
        AutoRelease.Visible = True  'only show post savannah button and autorelease checkbox if user is in sylacauga
    End If
    td.Rows = 1: pc.Clear: wc.Clear: tid.Clear: Label9.Visible = False
    pfile = Me.pallogs & "RO" & r12tkt & ".txt"
    If Len(Dir(pfile)) > 0 Then
        postlit = "Posted to R12 ticket: RO" & r12tkt
        postlit.Visible = True
        edtrls.Enabled = False
        edscans.Enabled = False
        postmenu.Enabled = False
        renmenu.Enabled = False
        addtline.Enabled = False
        deltline.Enabled = False
        canline.Enabled = False
        addbc.Enabled = False
        addwraps.Enabled = False
        edunits.Enabled = False
        Command6.Visible = False        'dont show post button and autorelease checkbox if pfile is null. I think this means it's already been posted.
        AutoRelease.Visible = False
    End If
    refresh_trkgrid (r12tkt)
    billgrid.Rows = 1
End Sub

Private Sub rentrl_Click()
    If Val(billgrid.TextMatrix(billgrid.Row, 17)) = 0 Then Exit Sub
    Call rename_trailer(billgrid.TextMatrix(billgrid.Row, 17))
End Sub

Private Sub sd_Click()
    Dim ds As adodb.Recordset, js As adodb.Recordset, sqlx As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    td.Rows = 1: Combo1.Clear: List1.Clear
    sqlx = "Select t.runid,t.branch,t.account,b.branchname,t.trlno,sum(t.units) from trailers t,branches b"
    sqlx = sqlx & " Where t.shipdate = '" & sd & "'"
    sqlx = sqlx & " And t.branch = b.branch"
    sqlx = sqlx & " and t.plant = " & Me.plantno
    sqlx = sqlx & " Group by t.runid,t.branch,t.account,b.branchname,t.trlno"
    sqlx = sqlx & " order by b.branchname,t.trlno"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        ds.Close ': db.Close
        MsgBox "No trailers found for selected date..", vbOKOnly, "Schedule"
        Exit Sub
    End If
    ds.MoveFirst
    Do Until ds.EOF
        If ds!account <= "0" Then
            Combo1.AddItem ds!branchname & " " & ds!trlno
            List1.AddItem ds!runid & "......"
        Else
            sqlx = "select acctdesc from jobbing where branch = " & ds!branch
            sqlx = sqlx & " and account = '" & ds!account & "'"
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
    
    ldate = Format(DateAdd("d", -1, sd), "mm-dd-yyyy")
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
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

Private Sub td_KeyPress(KeyAscii As Integer)
    If td.Row = 0 Then Exit Sub
    If postlit.Visible = True Then
        Command6.Visible = False                                            'jv071918
        AutoRelease.Visible = False
        MsgBox postlit.Caption, vbOKOnly + vbInformation, "sorry, edits are disabled...."
        Exit Sub
    End If
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
    If billgrid.Rows > 1 Then check_totals (r12tkt)
End Sub

Private Sub td_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edtrls
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

Private Sub tpost_Click()
    Dim i As Integer, k As Integer, cfile As String, logpath As String
    If ycolor.Visible = True Then
        MsgBox ycolor.Caption, vbOKOnly + vbExclamation, "sorry, try again..."
        Exit Sub
    End If
    logpath = Me.pallogs
    Screen.MousePointer = 11
    Call fetch_r12_bill
    DoEvents
    Call testro_bill(Form1.plantno, Format(sd, "MMddyyyy"))
    Screen.MousePointer = 0
End Sub

Private Sub trkgrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edsched
End Sub

Private Sub trlkey_Change()
    Dim i As Integer, k As Integer, u As Long, j As Integer
    k = 0: u = 0
    For i = 1 To billgrid.Rows - 1
        If billgrid.TextMatrix(i, 5) = trlkey Then
            k = k + 1
            u = u + Val(billgrid.TextMatrix(i, 11))
            u = u + Val(billgrid.TextMatrix(i, 13))
        Else
            Exit For
        End If
    Next i
    hcolor.Caption = u & " Units"
    cntlit.Caption = k & " Records"
    Call check_totals(billgrid.TextMatrix(billgrid.Row, 17))
End Sub

