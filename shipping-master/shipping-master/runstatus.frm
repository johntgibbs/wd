VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form runstatus 
   Caption         =   "Oracle Ticket Status"
   ClientHeight    =   12840
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12585
   LinkTopic       =   "Form2"
   ScaleHeight     =   12840
   ScaleWidth      =   12585
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid7 
      Height          =   1815
      Left            =   0
      TabIndex        =   14
      Top             =   240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3201
      _Version        =   327680
      BackColorFixed  =   8454143
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid6 
      Height          =   1815
      Left            =   0
      TabIndex        =   5
      Top             =   10560
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3201
      _Version        =   327680
      BackColorFixed  =   16777088
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid5 
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   8520
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3201
      _Version        =   327680
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid4 
      Height          =   1815
      Left            =   0
      TabIndex        =   3
      Top             =   6480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3201
      _Version        =   327680
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   4440
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3201
      _Version        =   327680
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   3360
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1508
      _Version        =   327680
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1508
      _Version        =   327680
      AllowUserResizing=   3
   End
   Begin VB.Label wokey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
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
      Left            =   10800
      TabIndex        =   16
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Transport Schedule:"
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
      TabIndex        =   15
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label6 
      Caption         =   "Oracle Trailer Ticket"
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
      TabIndex        =   13
      Top             =   10320
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   "Remote Plant Trailers"
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
      TabIndex        =   12
      Top             =   8280
      Width           =   5895
   End
   Begin VB.Label plantkey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   11
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label runkey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   10
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Shipping.Trailers"
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
      TabIndex        =   9
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Shipping.GroupItems"
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
      TabIndex        =   8
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Shipping.Trgroups"
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
      TabIndex        =   7
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Shipping.Runs"
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
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Menu tktmenu 
      Caption         =   "Ticket"
      Begin VB.Menu clrtkt 
         Caption         =   "Clear Ticket"
      End
   End
End
Attribute VB_Name = "runstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim rdb As adodb.Connection

Private Sub refresh_grid1(oratkt As String)
    Dim s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    Dim ds As adodb.Recordset
    s = "select * from runs where id = " & oratkt
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!loaded & Chr(9)
            s = s & ds!Destination & Chr(9)
            s = s & ds!locname & Chr(9)
            s = s & ds!trlno & Chr(9)
            s = s & ds!trlsize & Chr(9)
            s = s & Format(ds!trldate, "MM-dd-yyyy") & Chr(9)
            s = s & Format(ds!startime, "h:mm Am/Pm") & Chr(9)
            s = s & ds!pickup & Chr(9)
            s = s & ds!oc & Chr(9)
            s = s & ds!yardnote & Chr(9)
            s = s & ds!loadnote
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "^Id|^Loaded|^Destination|^LocName|^TrlNo|^TrlSize|^Date|^StartTime|<DriverNote|^OC|<YardNote|<LoadNote"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 3000
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 1000
    Grid1.ColWidth(11) = 1000
    Call refresh_grid2(oratkt)
End Sub

Private Sub refresh_grid2(oratkt As String)
    Dim s As String, gc As String, scol As Integer
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 5
    Dim ds As adodb.Recordset
    s = "select * from trgroups where run1 = " & oratkt & " or run2 = " & oratkt
    s = s & " or run3 = " & oratkt & " or run4 = " & oratkt
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            gc = ds!groupcode
            If ds!run1 = oratkt Then scol = 1
            If ds!run2 = oratkt Then scol = 2
            If ds!run3 = oratkt Then scol = 3
            If ds!run4 = oratkt Then scol = 4
            s = ds!groupcode & Chr(9)
            s = s & ds!run1 & Chr(9)
            s = s & ds!run2 & Chr(9)
            s = s & ds!run3 & Chr(9)
            s = s & ds!run4
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid2.Rows > 1 Then
        Grid2.RowSel = Grid2.Row
        Grid2.Col = scol: Grid2.ColSel = scol
        Grid2.FillStyle = flexFillRepeat
        Grid2.CellBackColor = Label2.BackColor
    End If
    s = "^Group|^Run1|^Run2|^Run3|^Run4"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 1000
    Call refresh_grid3(oratkt, gc, scol)
End Sub

Private Sub refresh_grid3(oratkt As String, gcode As String, scol As Integer)
    Dim s As String, i As Integer, ds As adodb.Recordset
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 13
    If gcode > " " And scol > 0 Then
        s = "select * from groupitems where groupcode = '" & gcode & "'"
        s = s & " and qty" & scol & " > 0 and whs" & scol & " > 0"
        Set ds = Sdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds!id & Chr(9)
                s = s & ds!groupcode & Chr(9)
                s = s & ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                s = s & ds!qty1 & Chr(9)
                s = s & ds!whs1 & Chr(9)
                s = s & ds!qty2 & Chr(9)
                s = s & ds!whs2 & Chr(9)
                s = s & ds!qty3 & Chr(9)
                s = s & ds!whs3 & Chr(9)
                s = s & ds!qty4 & Chr(9)
                s = s & ds!whs4 & Chr(9)
                s = s & ds!grank
                Grid3.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    If Grid3.Rows > 1 Then
        Grid3.FillStyle = flexFillRepeat
        For i = 1 To Grid3.Rows - 1
            Grid3.Row = i: Grid3.RowSel = i
            If scol = 1 Then
                Grid3.Col = 4
                Grid3.ColSel = 5
            End If
            If scol = 2 Then
                Grid3.Col = 6
                Grid3.ColSel = 7
            End If
            If scol = 3 Then
                Grid3.Col = 8
                Grid3.ColSel = 9
            End If
            If scol = 4 Then
                Grid3.Col = 10
                Grid3.ColSel = 11
            End If
            Grid3.CellBackColor = Label3.BackColor
        Next i
        Grid3.Row = 1
    End If
    s = "^Id|^Group|^SKU|<Product|^Qty1|^Whs1|^Qty2|^Whs2|^Qty3|^Whs3|^Qty4|^Whs4|^Rank"
    Grid3.FormatString = s
    Grid3.ColWidth(0) = 1000
    Grid3.ColWidth(1) = 1000
    Grid3.ColWidth(2) = 1000
    Grid3.ColWidth(3) = 3000
    Grid3.ColWidth(4) = 1000
    Grid3.ColWidth(5) = 1000
    Grid3.ColWidth(6) = 1000
    Grid3.ColWidth(7) = 1000
    Grid3.ColWidth(8) = 1000
    Grid3.ColWidth(9) = 1000
    Grid3.ColWidth(10) = 1000
    Grid3.ColWidth(11) = 1000
    Grid3.ColWidth(12) = 1000
    Call refresh_grid4(oratkt)
End Sub

Private Sub refresh_grid4(oratkt As String)
    Dim s As String, i As Integer
    Grid4.Clear: Grid4.Rows = 1: Grid4.Cols = 16
    Dim ds As adodb.Recordset
    s = "select * from trailers where runid = " & oratkt
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!runid & Chr(9)
            s = s & ds!groupcode & Chr(9)
            s = s & ds!plant & Chr(9)
            s = s & ds!branch & Chr(9)
            s = s & ds!account & Chr(9)
            s = s & Format(ds!shipdate, "MM-dd-yyyy") & Chr(9)
            s = s & ds!trlno & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!pallets & Chr(9)
            s = s & ds!wraps & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!whs_num & Chr(9)
            s = s & ds!pb_flag & Chr(9)
            s = s & ds!ra_flag
            Grid4.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "^Id|^RunId|^Group|^Plant|^Branch|^Account|^ShipDate|^TrlNo|^SKU|<Product|^Pallets|^Wraps|^Units|^Whs|^PBill|^Post"
    Grid4.FormatString = s
    Grid4.ColWidth(0) = 1000
    Grid4.ColWidth(1) = 1000
    Grid4.ColWidth(2) = 1000
    Grid4.ColWidth(3) = 1000
    Grid4.ColWidth(4) = 1000
    Grid4.ColWidth(5) = 1000
    Grid4.ColWidth(6) = 1000
    Grid4.ColWidth(7) = 1000
    Grid4.ColWidth(8) = 1000
    Grid4.ColWidth(9) = 3000
    Grid4.ColWidth(10) = 1000
    Grid4.ColWidth(11) = 1000
    Grid4.ColWidth(12) = 1000
    Grid4.ColWidth(13) = 1000
    Grid4.ColWidth(14) = 1000
    Grid4.ColWidth(15) = 1000
    Call refresh_grid5(oratkt)
End Sub

Private Sub refresh_grid5(oratkt As String)
    Dim s As String, i As Integer
    Grid5.Clear: Grid5.Rows = 1: Grid5.Cols = 16: Label5.Caption = "..."
    If Me.plantkey = "T10" Or Me.plantkey = "50" Then Exit Sub
    Dim db As adodb.Connection, ds As adodb.Recordset
    Set db = CreateObject("ADODB.Connection")
    If Me.plantkey = "A10" Or Me.plantkey = "52" Then           'jv091115
        db.Open Form1.syship
        Label5.Caption = "Sylacauga Shipping.Trailers"
    End If
    If Me.plantkey = "K10" Or Me.plantkey = "51" Then           'jv091115
        Label5.Caption = "Broken Arrow Shipping.Trailers"
        db.Open Form1.baship
    End If
    s = "select * from trailers where runid = " & oratkt
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!runid & Chr(9)
            s = s & ds!groupcode & Chr(9)
            s = s & ds!plant & Chr(9)
            s = s & ds!branch & Chr(9)
            s = s & ds!account & Chr(9)
            s = s & Format(ds!shipdate, "MM-dd-yyyy") & Chr(9)
            s = s & ds!trlno & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!pallets & Chr(9)
            s = s & ds!wraps & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!whs_num & Chr(9)
            s = s & ds!pb_flag & Chr(9)
            s = s & ds!ra_flag
            Grid5.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    s = "^Id|^RunId|^Group|^Plant|^Branch|^Account|^ShipDate|^TrlNo|^SKU|<Product|^Pallets|^Wraps|^Units|^Whs|^PBill|^Post"
    Grid5.FormatString = s
    Grid5.ColWidth(0) = 1000
    Grid5.ColWidth(1) = 1000
    Grid5.ColWidth(2) = 1000
    Grid5.ColWidth(3) = 1000
    Grid5.ColWidth(4) = 1000
    Grid5.ColWidth(5) = 1000
    Grid5.ColWidth(6) = 1000
    Grid5.ColWidth(7) = 1000
    Grid5.ColWidth(8) = 1000
    Grid5.ColWidth(9) = 3000
    Grid5.ColWidth(10) = 1000
    Grid5.ColWidth(11) = 1000
    Grid5.ColWidth(12) = 1000
    Grid5.ColWidth(13) = 1000
    Grid5.ColWidth(14) = 1000
    Grid5.ColWidth(15) = 1000
End Sub

Private Sub refresh_grid6(oratkt)
    Dim cfile As String, s As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim f10 As String, f11 As String, f12 As String, f13 As String, f14 As String, f15 As String
    Grid6.Clear: Grid6.Rows = 1: Grid6.Cols = 16
    If Me.plantkey = "T10" Or Me.plantkey = "50" Then
        cfile = "\\bbc-01-prodtrk\wd\pallogs\ro" & oratkt & ".txt"
    End If
    If Me.plantkey = "A10" Or Me.plantkey = "52" Then
        cfile = "\\bbsy-02-dc\f\user\waredist\data\pallogs\ro" & oratkt & ".txt"
    End If
    If Me.plantkey = "K10" Or Me.plantkey = "51" Then
        cfile = "\\bbba-03-dc\f\user\waredist\data\pallogs\ro" & oratkt & ".txt"
    End If
    'MsgBox cfile
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15
            s = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & f4 & Chr(9) & f5 & Chr(9)
            s = s & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9) & f9 & Chr(9) & f10 & Chr(9)
            s = s & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9) & f14 & Chr(9) & f15
            Grid6.AddItem s
        Loop
        Close #1
    End If
    s = "^Ticket|^FromOrg|^FromSub|^FromLoc|^ToOrg|^ToSub|^ToLoc|^Account|^SKU|^LotNum|^Units|^UOM|^ShipDate|<Comment|^EarlyDate|^PFlag"
    Grid6.FormatString = s
    Grid6.ColWidth(0) = 1000
    Grid6.ColWidth(1) = 1000
    Grid6.ColWidth(2) = 1000
    Grid6.ColWidth(3) = 1000
    Grid6.ColWidth(4) = 1000
    Grid6.ColWidth(5) = 1000
    Grid6.ColWidth(6) = 1000
    Grid6.ColWidth(7) = 1000
    Grid6.ColWidth(8) = 1000
    Grid6.ColWidth(9) = 1000
    Grid6.ColWidth(10) = 1000
    Grid6.ColWidth(11) = 1000
    Grid6.ColWidth(12) = 1000
    Grid6.ColWidth(13) = 2000
    Grid6.ColWidth(14) = 1000
    Grid6.ColWidth(15) = 1000
End Sub

Private Sub refresh_grid7()
    Dim s As String, i As Integer
    Dim db As adodb.Connection, ds As adodb.Recordset
    Set db = CreateObject("ADODB.Connection")
    Grid7.Clear: Grid7.Rows = 1: Grid7.Cols = 23
    If Val(Me.wokey) = 0 Then Exit Sub
    db.Open Form1.schdb
    s = "select * from truckwo where wonum = " & Me.wokey & " or parentwo = " & Me.wokey
    s = s & " order by parentwo, wodate, startime"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!wonum & Chr(9)
            s = s & Format(ds!wodate, "MM-dd-yyyy") & Chr(9)
            s = s & ds!origin & Chr(9)
            s = s & ds!Destination & Chr(9)
            s = s & ds!ethours & Chr(9)
            s = s & Format(ds!startime, "h:mm Am/Pm") & Chr(9)
            s = s & ds!drvid & Chr(9)
            s = s & ds!wtype & Chr(9)
            s = s & ds!jpnum & Chr(9)
            s = s & ds!parentwo & Chr(9)
            s = s & ds!linkwo & Chr(9)
            s = s & ds!description & Chr(9)
            s = s & ds!trlsize & Chr(9)
            s = s & ds!trlno & Chr(9)
            s = s & ds!eqnum & Chr(9)
            s = s & ds!mealpay & Chr(9)
            s = s & ds!r12ticket & Chr(9)
            s = s & ds!contents & Chr(9)
            s = s & ds!drvpool & Chr(9)
            s = s & ds!wostatus & Chr(9)
            s = s & ds!updatedby & Chr(9)
            s = s & ds!lastchange & Chr(9)
            s = s & ds!sealnum
            Grid7.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    s = "^WoNum|^WoDate|^Origin|^Destination|^ethours|^StartTime|^drvid|^wtype|<jpnum|<parentwo|^linkwo|<description|^trlsize|^trlno|^eqnum|^mealpay|^r12ticket|<contents|<drvpool|^wostatus|<Updatedby|<lastchange|^Sealnum"
    Grid7.FormatString = s
    Grid7.ColWidth(0) = 1000
    Grid7.ColWidth(1) = 1000
    Grid7.ColWidth(2) = 1000
    Grid7.ColWidth(3) = 1000
    Grid7.ColWidth(4) = 1000
    Grid7.ColWidth(5) = 1000
    Grid7.ColWidth(6) = 1000
    Grid7.ColWidth(7) = 1000
    Grid7.ColWidth(8) = 1600
    Grid7.ColWidth(9) = 1000
    Grid7.ColWidth(10) = 1000
    Grid7.ColWidth(11) = 3000
    Grid7.ColWidth(12) = 1000
    Grid7.ColWidth(13) = 1000
    Grid7.ColWidth(14) = 1000
    Grid7.ColWidth(15) = 1000
    Grid7.ColWidth(16) = 1000
    Grid7.ColWidth(17) = 1000
    Grid7.ColWidth(18) = 1000
    Grid7.ColWidth(19) = 1000
    Grid7.ColWidth(20) = 1000
    Grid7.ColWidth(21) = 1000
    Grid7.ColWidth(22) = 1000
End Sub

Private Sub clrtkt_Click()
    Dim s As String, i As Integer, db As adodb.Connection
    If Grid6.Rows > 1 Then
        MsgBox "This ticket has already been scanned and posted.", vbOKOnly + vbInformation, "Cannot clear...."
        Exit Sub
    End If
    If Val(wokey) > 0 Then
        'Set db = CreateObject("ADODB.Connection")
        'db.Open Form1.schdb
        'db.Execute s
        'db.Close
        s = "Update truckwo set r12ticket = ' ' where wonum = " & wokey.Caption
        MsgBox s
    End If
    If Grid1.Rows > 1 Then
        s = "Delete from runs where id = " & runkey
        MsgBox s
        'sdb.Execute s
    End If
    If Grid2.Rows > 1 Then
        s = "Update trgroups "
        If Grid2.TextMatrix(1, 1) = runkey Then s = s & "set run1 = 0"
        If Grid2.TextMatrix(1, 2) = runkey Then s = s & "set run2 = 0"
        If Grid2.TextMatrix(1, 3) = runkey Then s = s & "set run3 = 0"
        If Grid2.TextMatrix(1, 4) = runkey Then s = s & "set run4 = 0"
        s = s & " Where groupcode = '" & Grid2.TextMatrix(1, 0) & "'"
        MsgBox s
        'sdb.Execute s
        If Grid3.Rows > 1 Then
            For i = 1 To Grid3.Rows - 1
                s = "update groupitems "
                If Grid2.TextMatrix(1, 1) = runkey Then s = s & "set qty1 = 0, whs1 = 0"
                If Grid2.TextMatrix(1, 2) = runkey Then s = s & "set qty2 = 0, whs2 = 0"
                If Grid2.TextMatrix(1, 3) = runkey Then s = s & "set qty3 = 0, whs3 = 0"
                If Grid2.TextMatrix(1, 4) = runkey Then s = s & "set qty4 = 0, whs4 = 0"
                s = s & " where id = " & Grid3.TextMatrix(i, 0)
                MsgBox s
                'sdb.Execute s
            Next i
        End If
    End If
    s = "Delete from trailers where runid = " & runkey
    MsgBox s
    'sdb.Execute s
    
    If Grid5.Rows > 1 Then
        'Set db = CreateObject("ADODB.Connection")
        If Me.plantkey = "A10" Then
        '    db.Open Form1.syship
            s = "Delete from trailers where runid = " & runkey
        '    db.Execute s
            MsgBox s, vbOKOnly + vbInformation, "Sylacauga"
        End If
        If Me.plantkey = "K10" Then
        '    db.Open Form1.baship
            s = "Delete from trailers where runid = " & runkey
        '    db.Execute s
            MsgBox s, vbOKOnly + vbInformation, "Broken Arrow"
        End If
        'db.Close
    End If
    
    If Val(wokey) > 0 Then
        trucknotes.Combo2.Enabled = False
        trucknotes.Command2.Enabled = False
        s = "The r12ticket has been cleared from the shipping tables.  Do not refresh the trailer schedule until"
        s = s & " the scheduled work order #" & wokey & " has been modified by the scheduling department."
        MsgBox s, vbOKOnly + vbInformation, "ticket cleared...."
    End If
End Sub

Private Sub Form_Load()
    Call build_skumast
    'Call refresh_grid1("288231")
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    Grid2.Width = Me.Width - 100
    Grid3.Width = Me.Width - 100
    Grid4.Width = Me.Width - 100
    Grid5.Width = Me.Width - 100
    Grid6.Width = Me.Width - 100
    Grid7.Width = Me.Width - 100
End Sub

Private Sub runkey_Change()
    Screen.MousePointer = 11
    Call refresh_grid7
    Call refresh_grid1(runkey.Caption)
    Call refresh_grid6(runkey.Caption)
    Screen.MousePointer = 0
End Sub

