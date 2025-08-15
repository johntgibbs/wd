VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form cranetraffic 
   Caption         =   "Tri-Level Crane Traffic"
   ClientHeight    =   9390
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   15765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9390
   ScaleWidth      =   15765
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Crane Conveyors Online "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CheckBox Check5 
         BackColor       =   &H0000FF00&
         Caption         =   "SR-5"
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
         Left            =   3480
         TabIndex        =   27
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H0000FF00&
         Caption         =   "SR-4"
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
         Left            =   2640
         TabIndex        =   26
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0000FF00&
         Caption         =   "SR-3"
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
         Left            =   1800
         TabIndex        =   25
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000FF00&
         Caption         =   "SR-2"
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
         Left            =   960
         TabIndex        =   24
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000FF00&
         Caption         =   "SR-1"
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
         TabIndex        =   23
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group By "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12840
      TabIndex        =   18
      Top             =   0
      Width           =   2175
      Begin VB.OptionButton Option2 
         Caption         =   "Crane"
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
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Wrapper"
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
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox timelog 
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox logsize1 
      Height          =   285
      Left            =   5640
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox logfile1 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   7695
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
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4260
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   11880
      _Version        =   327680
      BackColorSel    =   12632256
      FocusRect       =   0
   End
   Begin VB.Label Label2 
      Caption         =   "0925832"
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
      Left            =   15120
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label pcount 
      Caption         =   "pcount"
      Height          =   255
      Left            =   15960
      TabIndex        =   17
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Poll Count:"
      Height          =   255
      Left            =   15120
      TabIndex        =   16
      Top             =   360
      Width           =   975
   End
   Begin VB.Label w5c 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR4 Racks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   15
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label w4c 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wrapper 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   11
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label w3c 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wrapper 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label w2c 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wrapper 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label w1c 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wrapper 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label sr5c 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR-5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label sr4c 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR-4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label sr3c 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR-3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label sr2c 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR-2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label sr1c 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu switchd 
         Caption         =   "Switch Destination"
      End
      Begin VB.Menu switchsr4 
         Caption         =   "Switch SR4 to Crane"
      End
      Begin VB.Menu retwrap 
         Caption         =   "Return to Wrapper"
      End
      Begin VB.Menu cwq 
         Caption         =   "Change Wrap Qty"
      End
      Begin VB.Menu changeq 
         Caption         =   "Change Queue Number"
      End
   End
End
Attribute VB_Name = "cranetraffic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub poll_logs()
    Dim sdate As String, t As String
    Do While True
        DoEvents
        'If Form1.scanlogs.Value = 0 Then Exit Do
        sdate = Format(Now, "MMddyyyy")
        Me.logfile1 = "v:\pallogs\recv" & sdate & ".txt"
        If Len(Dir(Me.logfile1)) > 0 Then
            Me.logsize1 = FileLen(Me.logfile1)
            DoEvents
            'pcount.Caption = Val(pcount.Caption) + 1
        
        End If
        'Form1.logfile2 = "v:\pallogs\move" & sdate & ".txt"
        'If Len(Dir(Form1.logfile2)) > 0 Then
        '    Form1.logsize2 = FileLen(Form1.logfile2)
        '    DoEvents
        'End If
        t = Format(Now, "hh:mm:ss")
        If Right(t, 1) = "0" Or Right(t, 1) = "5" Then
            Me.timelog = Format(Now, "h:mm:ss am/pm")
            'pcount.Caption = Val(pcount.Caption) + 1
        End If
        'If Len(Dir(Form1.shipordfile)) > 0 Then
        '    Form1.shipordtime = Format(FileDateTime(Form1.shipordfile), "MM-dd-yyyy hh:mm:ss am/pm")
        '    DoEvents
        'End If
        'pcount.Caption = Val(pcount.Caption) + 1
    Loop
End Sub

Sub refresh_sr_logs()
    Dim cfile As String, sdate As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim f10 As String, f11 As String, f12 As String, f13 As String, f14 As String
    Dim f15 As String, f16 As String, f17 As String, f18 As String, f19 As String
    'Dim db As ADODB.Connection, s As String, p As ptask, ds As ADODB.Recordset
    Dim s As String, p As ptask, ds As ADODB.Recordset
    
    'Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr

    
    Grid2.Redraw = False
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 7
    
    s = "select source, product, palletid, reqid from paltasks where area = 'TRAFFIC MASTER'"
    'Set ds = db.Execute(s)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If UCase(Left(ds!source, 3)) = "TRI" Then
                s = ds!source & Chr(9)
            Else
                s = "SR4 5" & Chr(9)
            End If
            s = s & ds!product & Chr(9)
            s = s & ds!palletid & Chr(9)
            s = s & ds!reqid
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close ': db.Close
    
    'sdate = Format(Now, "MMddyyyy")
    'For i = 1 To 3
    '    cfile = "\\bbc-01-prodtrk\wd\pallogs\recv" & sdate & ".txt"
    '    If Len(Dir(cfile)) > 0 Then
    '        Open cfile For Input As #1
    '        Do Until EOF(1)
    '            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
    '            'If f1 = "TRI-LEVEL 1" Or f1 = "TRI-LEVEL 2" Or f1 = "TRI-LEVEL 3" Or f1 = "TRI-LEVEL 4" Then
    '            If f4 = "TRAFFIC MASTER" Then
    '                s = f1 & Chr(9)         'area
    '                s = s & f5 & Chr(9)     'sku
    '                s = s & f6 & Chr(9)     'barcode
    '                's = s & f4 & Chr(9)     'pallet #
    '                's = s & f6 & Chr(9)     'function
    '                's = s & f8 & Chr(9)     'dest lane
    '                's = s & Format(f9, "hh:mm")  'time
    '                Grid2.AddItem s
    '            End If
    '        Loop
    '        Close #1
    '    End If
    'Next i
    If Grid2.Rows > 1 Then
        Grid2.RowSel = Grid2.Row
        Grid2.Col = 6: Grid2.ColSel = 6
        Grid2.Sort = 5
    End If
    Grid2.FormatString = "^Wrapper|^SKU|^Barcode"
    Grid2.ColWidth(0) = 2000
    Grid2.ColWidth(1) = 3000
    Grid2.ColWidth(2) = 2000
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 1000
    Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 1000
    Grid2.Redraw = True
End Sub

Sub wrap12_switch(pkey As Long, pswtch As String, pno As String)
    Dim s As String, p As ptask, ds As ADODB.Recordset
    'Dim db As ADODB.Connection, s As String, p As ptask, ds As ADODB.Recordset
    'Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    'MsgBox pswtch
    If pswtch = "1>2" Then
        s = "update queue_infc set whse_num = 2 where id = " & pkey
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
    End If
    If pswtch = "1>3" Then
        s = "update queue_infc set whse_num = 3 where id = " & pkey
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
        If Val(pno) > 0 Then Call send_dai_request(pkey, "ADD", pno)
    End If
    If pswtch = "1>4" Or pswtch = "2>4" Or pswtch = "3>4" Or pswtch = "5>4" Then
        s = "select * from queue_infc where id = " & pkey
        'Set ds = db.Execute(s)
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            p.qty = ds!rack_num
            p.uom = "Wraps"
            p.lotnum = ds!lot_num
            p.units = ds!units
            p.lotnum2 = ds!lot_num2
            p.units2 = ds!units2
            p.palletid = ds!palletid
        End If
        ds.Close
        s = "update queue_infc set queue_num = 0 where id = " & pkey
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
        p.area = "FORKLIFT"
        p.description = " "
        p.source = "TRI LEVEL"
        p.target = "RACKS"
        p.product = Grid1.TextMatrix(Grid1.Row, 4)
        p.status = "PEND"
        p.userid = " "
        p.trandate = Format(Now, "yyMMdd hh:mm:ss")
        p.reqid = pno
        Call insert_trans(p)
    End If
    If pswtch = "1>5" Then
        s = "update queue_infc set whse_num = 5 where id = " & pkey
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
        If Val(pno) > 0 Then Call send_dai_request(pkey, "ADD", pno)
    End If
    If pswtch = "2>1" Then
        s = "update queue_infc set whse_num = 1 where id = " & pkey
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
    End If
    If pswtch = "2>3" Then
        s = "update queue_infc set whse_num = 3 where id = " & pkey
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
        If Val(pno) > 0 Then Call send_dai_request(pkey, "ADD", pno)
    End If
    If pswtch = "2>5" Then
        s = "update queue_infc set whse_num = 5 where id = " & pkey
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
        If Val(pno) > 0 Then Call send_dai_request(pkey, "ADD", pno)
    End If
    
    
    If pswtch = "3>1" Or pswtch = "5>1" Then
        If Val(pno) > 0 Then Call send_dai_request(pkey, "DELETE", pno)
        DoEvents
        s = "update queue_infc set whse_num = 1 where id = " & pkey
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
    End If
    If pswtch = "3>2" Or pswtch = "5>2" Then
        If Val(pno) > 0 Then Call send_dai_request(pkey, "DELETE", pno)
        DoEvents
        s = "update queue_infc set whse_num = 2 where id = " & pkey
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
    End If
    If pswtch = "3>5" Then
        If Val(pno) > 0 Then Call send_dai_request(pkey, "DELETE", pno)
        DoEvents
        s = "update queue_infc set whse_num = 5 where id = " & pkey
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
        If Val(pno) > 0 Then Call send_dai_request(pkey, "ADD", pno)
    End If
    If pswtch = "5>3" Then
        If Val(pno) > 0 Then Call send_dai_request(pkey, "DELETE", pno)
        DoEvents
        s = "update queue_infc set whse_num = 3 where id = " & pkey
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
        If Val(pno) > 0 Then Call send_dai_request(pkey, "ADD", pno)
    End If
    'db.Close
End Sub

Sub refresh_crane_conveyors()
    Dim ds As Recordset, s As String
    'Dim db As ADODB.Connection, ds As Recordset, s As String
    Dim bc As String, psku As String, plot As String, pplt As String
    Dim i As Integer, K As Integer, pwhs As String, y As String, x As Integer
    Dim fc1 As Integer, fc2 As Integer, fc3 As Integer, fc4 As Integer, fc5 As Integer      'jv020116
    'Screen.MousePointer = 11
    y = Grid1.TextMatrix(Grid1.Row, 0): x = Grid1.Col
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 9
    Grid1.Redraw = False
    'Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    s = "select id,whse_num,queue_num,rack_num,palletid"
    s = s & " from queue_infc where queue_num > 0"
    s = s & " and source = 'TML'"
    's = s & " order by whse_num, queue_num"
    s = s & " order by queue_num desc"
    'Set ds = db.Execute(s)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & " " & Chr(9)
            s = s & ds!whse_num & Chr(9)
            s = s & ds!queue_num & Chr(9)
            s = s & " " & Chr(9)
            s = s & ds!rack_num & Chr(9)
            s = s & ds!palletid & Chr(9)
            's = s & ds!trandate & Chr(9)
            's = s & ds!lotnum
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select * from paltasks where area = 'FORKLIFT'"
    s = s & " and source in ('TRI LEVEL', 'TRI-LEVEL') and status = 'PEND' and userid < '0'"
    'Set ds = db.Execute(s)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & "5" & Chr(9)
            s = s & "4" & Chr(9)
            s = s & "0" & Chr(9)
            s = s & ds!product & Chr(9)
            s = s & ds!qty & Chr(9)
            s = s & ds!palletid & Chr(9)
            s = s & ds!reqid
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
        
    ds.Close                                                    'jv020116
    s = "select * from wrx_conveyor_status where id = 1"        'jv020116
    Set ds = Wdb.Execute(s)                                     'jv020116
    If ds.BOF = False Then                                      'jv020116
        ds.MoveFirst                                            'jv020116
        fc1 = ds!conveyorsr1                                    'jv020116
        fc2 = ds!conveyorsr2                                    'jv020116
        fc3 = ds!conveyorsr3                                    'jv020116
        fc4 = ds!conveyorsr4                                    'jv020116
        fc5 = ds!conveyorsr5                                    'jv020116
    End If                                                      'jv020116
    
    ds.Close ': db.Close
    
    If fc1 <> Check1.Value Then Check1.Value = fc1              'jv020116
    If fc2 <> Check2.Value Then Check2.Value = fc2              'jv020116
    If fc3 <> Check3.Value Then Check3.Value = fc3              'jv020116
    If fc4 <> Check4.Value Then Check4.Value = fc4              'jv020116
    If fc5 <> Check5.Value Then Check5.Value = fc5              'jv020116
    
    refresh_sr_logs
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            For K = 1 To Grid2.Rows - 1
                If Grid1.TextMatrix(i, 6) = Grid2.TextMatrix(K, 2) Then
                    Grid1.TextMatrix(i, 1) = Right(Grid2.TextMatrix(K, 0), 1)
                    Grid1.TextMatrix(i, 4) = Grid2.TextMatrix(K, 1)
                    Grid1.TextMatrix(i, 7) = Grid2.TextMatrix(K, 3)
                    Exit For
                End If
            Next K
        Next i
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 1: Grid1.ColSel = 1
            If Grid1.TextMatrix(i, 1) = "1" Then
                Grid1.CellBackColor = w1c.BackColor
                Grid1.CellForeColor = w1c.ForeColor
            End If
            If Grid1.TextMatrix(i, 1) = "2" Then
                Grid1.CellBackColor = w2c.BackColor
                Grid1.CellForeColor = w2c.ForeColor
            End If
            If Grid1.TextMatrix(i, 1) = "3" Then
                Grid1.CellBackColor = w3c.BackColor
                Grid1.CellForeColor = w3c.ForeColor
            End If
            If Grid1.TextMatrix(i, 1) = "4" Then
                Grid1.CellBackColor = w4c.BackColor
                Grid1.CellForeColor = w4c.ForeColor
            End If
            If Grid1.TextMatrix(i, 1) = "5" Then
                Grid1.CellBackColor = w5c.BackColor
                Grid1.CellForeColor = w5c.ForeColor
            End If
        Next i
        For i = 1 To Grid1.Rows - 1
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 2: Grid1.ColSel = Grid1.Cols - 1
            If Grid1.TextMatrix(i, 2) = "1" Then Grid1.CellBackColor = sr1c.BackColor
            If Grid1.TextMatrix(i, 2) = "2" Then Grid1.CellBackColor = sr2c.BackColor
            If Grid1.TextMatrix(i, 2) = "3" Then Grid1.CellBackColor = sr3c.BackColor
            If Grid1.TextMatrix(i, 2) = "4" Then Grid1.CellBackColor = sr4c.BackColor
            If Grid1.TextMatrix(i, 2) = "5" Then Grid1.CellBackColor = sr5c.BackColor
            If Option1.Value = True Then
                Grid1.TextMatrix(i, 8) = Grid1.TextMatrix(i, 1) & Format(Val(Grid1.TextMatrix(i, 3)), "000000")
            Else
                Grid1.TextMatrix(i, 8) = Grid1.TextMatrix(i, 2) & Format(Val(Grid1.TextMatrix(i, 3)), "000000")
            End If
        Next i
        For i = Grid1.Rows - 1 To 1 Step -1
            If Option1.Value = True And Grid1.TextMatrix(i, 1) < "1" Then
                If Grid1.Rows > 2 Then
                    Grid1.RemoveItem i
                Else
                    Grid1.Rows = 1
                End If
            End If
        Next i
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 8: Grid1.ColSel = 8
        Grid1.Sort = 5
        Grid1.Row = 0: Grid1.Col = 1
    End If
    s = "^ID|^Wrapper|^SR|^Queue|<Product|^Wraps|^BarCode|^Plate"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 5000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 2400
    Grid1.ColWidth(7) = 1200
    Grid1.ColWidth(8) = 1
    For i = 0 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = y Then
            Grid1.Row = i
            Exit For
        End If
    Next i
    Grid1.Col = x
    'If y < Grid1.Rows Then
    '    Grid1.Row = y
    '    Grid1.Col = x
    'End If
    Grid1.Redraw = True
    Screen.MousePointer = 0
    pcount.Caption = Val(pcount.Caption) + 1
End Sub

Private Sub changeq_Click()
    Dim pkey As Long, pno As String, pqueue As String
    'Dim db As ADODB.Connection, ds As ADODB.Recordset, s As String
    Dim ds As ADODB.Recordset, s As String
    pqueue = InputBox("Queue Number:", "Change Queue Number...", Grid1.TextMatrix(Grid1.Row, 3))
    If Len(pqueue) = 0 Then Exit Sub
    If Val(pqueue) < 1 Then Exit Sub
    'Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    'db.Open "odbc;database=wdracks;uid=bbcwd500;pwd=brenham500;dsn=wdsql500"
    s = "select * from queue_infc where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    s = s & " and queue_num > 0 and whse_num = " & Grid1.TextMatrix(Grid1.Row, 2)
    'Set ds = db.Execute(s)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "update queue_infc set queue_num = " & pqueue
        s = s & " where id = " & ds!id
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
    End If
    ds.Close ': db.Close
    refresh_crane_conveyors
End Sub

Private Sub Check1_Click()
    Dim s As String
    If Check1.Value = 0 Then
        Check1.BackColor = w1c.BackColor
        Check1.ForeColor = w1c.ForeColor
    Else
        Check1.BackColor = sr1c.BackColor
        Check1.ForeColor = sr1c.ForeColor
    End If
    s = "Update Wrx_conveyor_status set conveyorsr1 = " & Check1.Value & " Where id = 1"
    'MsgBox s
    Wdb.Execute s
End Sub

Private Sub Check2_Click()
    Dim s As String
    If Check2.Value = 0 Then
        Check2.BackColor = w1c.BackColor
        Check2.ForeColor = w1c.ForeColor
    Else
        Check2.BackColor = sr1c.BackColor
        Check2.ForeColor = sr1c.ForeColor
    End If
    s = "Update Wrx_conveyor_status set conveyorsr2 = " & Check2.Value & " Where id = 1"
    'MsgBox s
    Wdb.Execute s
End Sub

Private Sub Check3_Click()
    Dim s As String
    If Check3.Value = 0 Then
        Check3.BackColor = w1c.BackColor
        Check3.ForeColor = w1c.ForeColor
    Else
        Check3.BackColor = sr1c.BackColor
        Check3.ForeColor = sr1c.ForeColor
    End If
    s = "Update Wrx_conveyor_status set conveyorsr3 = " & Check3.Value & " Where id = 1"
    'MsgBox s
    Wdb.Execute s
End Sub

Private Sub Check4_Click()
    Dim s As String
    If Check4.Value = 0 Then
        Check4.BackColor = w1c.BackColor
        Check4.ForeColor = w1c.ForeColor
    Else
        Check4.BackColor = sr1c.BackColor
        Check4.ForeColor = sr1c.ForeColor
    End If
    s = "Update Wrx_conveyor_status set conveyorsr4 = " & Check4.Value & " Where id = 1"
    'MsgBox s
    Wdb.Execute s
End Sub

Private Sub Check5_Click()
    Dim s As String
    If Check5.Value = 0 Then
        Check5.BackColor = w1c.BackColor
        Check5.ForeColor = w1c.ForeColor
    Else
        Check5.BackColor = sr1c.BackColor
        Check5.ForeColor = sr1c.ForeColor
    End If
    s = "Update Wrx_conveyor_status set conveyorsr5 = " & Check5.Value & " Where id = 1"
    'MsgBox s
    Wdb.Execute s
End Sub

Private Sub Command1_Click()
    refresh_crane_conveyors
End Sub

Private Sub cwq_Click()
    Dim s As String, p As ptask, ds As ADODB.Recordset
    'Dim db As ADODB.Connection, s As String, p As ptask, ds As ADODB.Recordset
    Dim wc As Integer, nwraps As Integer, psku As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    s = InputBox("Wrap Qty:", "Wrap Qty...", Grid1.TextMatrix(Grid1.Row, 5))
    If Len(s) = 0 Then Exit Sub
    If Val(s) <= 0 Then Exit Sub
    If Val(s) = Val(Grid1.TextMatrix(Grid1.Row, 5)) Then Exit Sub
    nwraps = Val(s)
    'Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    If Val(Grid1.TextMatrix(Grid1.Row, 3)) > 0 Then
        s = "select id, units2 from paltasks where area = 'TRAFFIC MASTER'"
        s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 6) & "'"
        s = s & " and status = 'PEND'"
    Else
        s = "select id, units2 from paltasks where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 6) & "'"
        s = s & " and status = 'PEND'"
    End If
    'Set ds = db.Execute(s)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        p = masterec(ds!id)
        If ds!units2 <> 0 Then
            MsgBox "Edit is not allowed for a mixed lot pallet.", vbOKOnly + vbInformation, "Use return to wrapper.."
            ds.Close ': db.Close
            Exit Sub
        End If
    Else
        MsgBox "Task is no longer available for edit.", vbOKOnly + vbInformation, "task is not pending..."
        ds.Close ': db.Close
        Exit Sub
    End If
    ds.Close
    psku = Trim(Left(p.palletid, 4))
    p.qty = Format(Val(p.qty) * -1)
    p.units = Format(Val(p.units) * -1)
    p.units2 = Format(Val(p.units2) * -1)
    Call post_recv_trans(p)
    wc = sku_info(psku, "units")
    wc = wc / sku_info(psku, "wraps")
    If nwraps = sku_info(psku, "wraps") Then
        p.product = psku & " " & sku_info(psku, "desc")
    Else
        p.product = psku & " " & nwraps & " Wraps! " & StrConv(sku_info(psku, "desc"), vbProperCase)
    End If
    p.qty = nwraps
    p.units = Format(wc * nwraps, "0")
    p.units2 = "0"
    Call post_recv_trans(p)
    If Val(Grid1.TextMatrix(Grid1.Row, 3)) > 0 Then
        s = "update queue_infc set rack_num = " & nwraps
        s = s & ", units = " & p.units
        s = s & " where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
    End If
    s = "update paltasks set qty = " & nwraps
    s = s & ", units = " & p.units
    s = s & ", units2 = 0"
    s = s & ", product = '" & p.product & "'"
    s = s & " where id = " & p.id
    'MsgBox s
    'db.Execute s
    Wdb.Execute s
    'db.Close
    Call refresh_crane_conveyors
    MsgBox "Wrap qty has been updated.", vbOKOnly + vbInformation, "Wrap qty changed..."
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Grid1.Font = "Arial"
    Grid1.FontSize = 12
    Grid1.FontBold = True
    refresh_crane_conveyors
    DoEvents
    poll_logs
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 200
    If Me.Height > 2000 Then
        Grid1.Height = Me.Height - (380 + Command1.Height)
    End If
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid1_RowColChange()
    If Val(Grid1.TextMatrix(Grid1.Row, 1)) < 1 Or Val(Grid1.TextMatrix(Grid1.Row, 1)) > 4 Then
        retwrap.Visible = False
    Else
        retwrap.Visible = True
    End If
    If Val(Grid1.TextMatrix(Grid1.Row, 3)) > 0 Then
        changeq.Visible = True
        switchd.Enabled = True
    Else
        changeq.Visible = False
        switchd.Enabled = False
    End If
    If Val(Grid1.TextMatrix(Grid1.Row, 2)) = 4 Then
        switchsr4.Visible = True
    Else
        switchsr4.Visible = False
    End If
End Sub

Private Sub logsize1_Change()
    refresh_crane_conveyors
End Sub

Private Sub Option1_Click()
    refresh_crane_conveyors
End Sub

Private Sub Option2_Click()
    refresh_crane_conveyors
End Sub

Private Sub retwrap_Click()
    Dim bc As String, parea As String, preq As String, emess As String
    parea = "TRI LEVEL " & Grid1.TextMatrix(Grid1.Row, 1)
    bc = Grid1.TextMatrix(Grid1.Row, 6)
    preq = Format(Val(Grid1.TextMatrix(Grid1.Row, 7)), "000000")
    emess = return_to_wrapper(bc, "TMaster", parea, preq)
    MsgBox emess, vbOKOnly + vbInformation, "Traffic....."
    refresh_crane_conveyors
End Sub

Private Sub switchd_Click()
    Dim pkey As Long, pno As String, pswtch As String
    'Dim db As ADODB.Connection, ds As ADODB.Recordset, s As String
    Dim ds As ADODB.Recordset, s As String
    pswtch = InputBox("Crane destination (1-5):", "Destination...", Grid1.TextMatrix(Grid1.Row, 2))
    If Len(pswtch) = 0 Then Exit Sub
    If Val(pswtch) < 1 Then Exit Sub
    If Val(pwstch) > 5 Then Exit Sub
    'Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    'db.Open "odbc;database=wdracks;uid=bbcwd500;pwd=brenham500;dsn=wdsql500"
    s = "select id from prodrcv where sku = '" & Trim(Left(Grid1.TextMatrix(Grid1.Row, 4), 4)) & "'"
    s = s & " and lot_num = '" & barcode_to_lotnum(Grid1.TextMatrix(Grid1.Row, 6)) & "'"
    If pswtch <> "4" Then s = s & " and sr" & pswtch & " > 0"
    'MsgBox s
    'Set ds = db.Execute(s)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update prodrcv set sr" & Grid1.TextMatrix(Grid1.Row, 2) & " = sr" & Grid1.TextMatrix(Grid1.Row, 2) & " + 1,"
        s = s & " sr" & pswtch & " = sr" & pswtch & " - 1"
        s = s & " where id = " & ds!id
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
        pno = Grid1.TextMatrix(Grid1.Row, 7)
        s = "Update paltasks set target = 'SR" & pswtch & "'"
        s = s & " where area = 'TRAFFIC MASTER' and reqid = '" & pno & "'"
        'MsgBox s
        'db.Execute s
        Wdb.Execute s
        s = "Update pallets set target = 'SR" & pswtch & "'"                        'jv072514
        s = s & " where barcode = '" & Grid1.TextMatrix(Grid1.Row, 6) & "'"         'jv072514
        'db.Execute s                                                                'jv072514
        Wdb.Execute s                                                                'jv072514
        'MsgBox s
        pswtch = Grid1.TextMatrix(Grid1.Row, 2) & ">" & pswtch
        pkey = Val(Grid1.TextMatrix(Grid1.Row, 0))
        If Grid1.TextMatrix(Grid1.Row, 1) = "1" Then
            Call wrap12_switch(pkey, pswtch, pno)
        End If
        If Grid1.TextMatrix(Grid1.Row, 1) = "2" Then
            Call wrap12_switch(pkey, pswtch, pno)
        End If
        refresh_crane_conveyors
    Else
        MsgBox "Product is not assigned to SR-" & pswtch, vbOKOnly + vbExclamation, "sorry, try again..."
    End If
    ds.Close ': db.Close
End Sub

Private Sub switchsr4_Click()
    Dim pkey As Long, pno As String, pswtch As String, p As ptask
    'Dim db As ADODB.Connection, ds As ADODB.Recordset, s As String, i As Long, q As Long
    Dim ds As ADODB.Recordset, s As String, i As Long, q As Long
    pswtch = InputBox("Crane destination (1-5):", "Destination...", Grid1.TextMatrix(Grid1.Row, 2))
    If Len(pswtch) = 0 Then Exit Sub
    If Val(pswtch) = 4 Then Exit Sub
    If Val(pswtch) < 1 Then Exit Sub
    If Val(pwstch) > 5 Then Exit Sub
    'Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    'db.Open "odbc;database=wdracks;uid=bbcwd500;pwd=brenham500;dsn=wdsql500"
    's = "select id from prodrcv where sku = '" & Trim(Left(Grid1.TextMatrix(Grid1.Row, 4), 4)) & "'"
    's = s & " and lot_num = '" & barcode_to_lotnum(Grid1.TextMatrix(Grid1.Row, 6)) & "'"
    's = s & " and sr" & pswtch & " > 0"
    s = "select id from prodrcv where sku = '777'"
    MsgBox s
    'Set ds = db.Execute(s)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        's = "Update prodrcv set sr" & Grid1.TextMatrix(Grid1.Row, 2) & " = sr" & Grid1.TextMatrix(Grid1.Row, 2) & " + 1,"
        's = s & " sr" & pswtch & " = sr" & pswtch & " - 1"
        's = s & " where id = " & ds!id
        s = "Update prodrcv set sr" & pswtch & " = sr" & pswtch & " - 1"
        s = s & " where id = " & ds!id
        MsgBox s
        'wdb.Execute s
        
        's = "Update paltasks set target = 'SR" & pswtch & "'"
        's = s & " where area = 'TRAFFIC MASTER' and reqid = '" & pno & "'"
        'wdb.Execute s
        
        pno = Grid1.TextMatrix(Grid1.Row, 7)
        pswtch = Grid1.TextMatrix(Grid1.Row, 2) & ">" & pswtch
        pkey = Val(Grid1.TextMatrix(Grid1.Row, 0))
        
        'Build task structure with forklift task
        p = masterec(pkey)
        
        'Mark forklift task complete
        s = "Update paltasks set status = 'COMP' where id = " & pkey
        MsgBox s
        'wdb.Execute s
        
        'Check if Traffic Master is available for update
        s = "select id from paltasks where area = 'TRAFFIC MASTER' and status = 'PEND'"
        s = s & " and reqid = '" & pno & "'"
        'Set ds = db.Execute(s)
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            s = "Update paltasks set target = 'SR" & pswtch & "'"
            s = s & " where id = " & ds!id
            MsgBox s
            'wdb.Execute s
        Else
            'Build with forklift info
            p.area = "TRAFFIC MASTER"
            p.source = "TRI-LEVEL 1"
            p.target = "SR" & Right(pswtch, 1)
            p.status = "PEND"
            p.userid = " "
            MsgBox "i = insert_trans(p)"
            'i = insert_trans(p)
        End If
                
        'Insert crane queue
        q = new_pallet_queue
        s = "Update queue_infc set whse_num = " & Right(pswtch, 1)
        s = s & ",sku='" & Trim(Left(p.product, 4)) & "'"
        s = s & ",lot_num = '" & p.lotnum & "'"
        s = s & ",drop_flag = ''"
        s = s & ",rack_num = " & p.qty
        s = s & ",units = " & p.units
        s = s & ",lot_num2 = '" & p.lotnum2 & "'"
        s = s & ",units2 = " & p.units2
        s = s & ",palletid = '" & p.palletid & "'"
        s = s & ",source = 'TML'"
        s = s & " where id = " & q
        MsgBox s
        'wdb.Execute s
                
                
        'Send expected receipt to 3 and 5
        If pswtch = "4>3" Or pswtch = "4>5" Then
            MsgBox "Call send_dai_request(" & q & ", 'ADD, " & pno & ")"
            'Call send_dai_request(q, "ADD", pno)
        End If
                
                
                
            
        
        
        refresh_crane_conveyors
    Else
        MsgBox "Product is not assigned to SR-" & pswtch, vbOKOnly + vbExclamation, "sorry, try again..."
    End If
    ds.Close ': db.Close

End Sub

Private Sub timelog_Change()
    refresh_crane_conveyors
End Sub
