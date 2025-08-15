VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form batpalpos 
   Caption         =   "Crane Pallets"
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11685
   LinkTopic       =   "Form14"
   ScaleHeight     =   8400
   ScaleWidth      =   11685
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4335
      Left            =   7440
      TabIndex        =   6
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   7646
      _Version        =   327680
      BackColorFixed  =   16777152
   End
   Begin VB.TextBox mylabels 
      Height          =   285
      Left            =   6720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5741
      _Version        =   327680
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "---"
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
      Left            =   7440
      TabIndex        =   7
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label ycolor 
      BackColor       =   &H0080FFFF&
      Caption         =   "ycolor"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label gcolor 
      BackColor       =   &H00C0E0FF&
      Caption         =   "gcolor"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Production Dates:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu prtmenu 
      Caption         =   "&Print"
   End
End
Attribute VB_Name = "batpalpos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_dates()
    Dim t As String, f0 As String, f1 As String, f2 As String
    Dim f3 As String, f4 As String, f5 As String, f6 As String
    Dim f7 As String, fsx As Long
    Combo1.Clear
    If Len(Dir(Me.mylabels)) = 0 Then
        Combo1.AddItem Format(Now, "m-dd-yyyy")
    Else
        fsx = FileLen(Me.mylabels)
        If fsx = 0 Then
            MsgBox "FileSize: " & Me.mylabels & "=" & fsx, vbOKOnly + vbInformation, "Current Label File Size Zero..."
            Combo1.AddItem Format(Now, "m-dd-yyyy")
        Else
            t = ">>"
            Open Me.mylabels For Input As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7
                If f1 <> t Then
                    Combo1.AddItem f1
                    t = f1
                End If
            Loop
            Close #1
        End If
    End If
    Combo1.ListIndex = 0
End Sub

Private Sub refresh_grid1()
    Dim s As String, f0 As String, f1 As String, f2 As String
    Dim f3 As String, f4 As String, f5 As String, f6 As String
    Dim f7 As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    If Len(Dir(Me.mylabels)) = 0 Then Exit Sub
    Open Me.mylabels For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7
        's = f0 & Chr(9)         'Plant
        's = s & f1 & Chr(9)     'Production Date
        'If f1 = Combo1 Then
            s = f1 & Chr(9)
            s = s & f2 & Chr(9)         'SKU
            s = s & f3 & Chr(9)     'Description
            s = s & f4 & Chr(9)     'Pallet Qty
            s = s & f5 & " "        'Code Date
            s = s & f6 & Chr(9)     'OP Code
            If f1 = Combo1 Then s = s & "*"
            Grid1.AddItem s
        'End If
    Loop
    Close #1
    s = "^Date|^SKU|<Description|^Planned|^Code Date|"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 600
    Grid1.ColWidth(2) = 2800
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 200
    ycolor.Visible = False
    Grid1.FillStyle = flexFillRepeat
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 5) = "*" Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = 5
                Grid1.CellBackColor = gcolor.BackColor
            End If
        Next i
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 5) = "*" Then
                Grid1.Row = i
                Grid1.TopRow = i
                Grid1_Click
                Exit For
            End If
        Next i
    End If
End Sub
Private Sub refresh_grid2()
    Dim ds As adodb.Recordset, s As String
    Dim msku As String, mlot As String
    Dim mdate1 As String, mdate2 As String
    Dim mt1 As Integer, mt2 As Integer, mt3 As Integer, mw As Integer, mt4 As Integer
    msku = Grid1.TextMatrix(Grid1.Row, 1)
    mdate1 = "1-1-" & Right(Grid1.TextMatrix(Grid1.Row, 0), 4)
    mdate2 = Grid1.TextMatrix(Grid1.Row, 0)
    mlot = Right(Grid1.TextMatrix(Grid1.Row, 0), 2)
    mlot = mlot & Format(DateDiff("d", mdate1, mdate2) + 1, "000")
    s = Grid1.TextMatrix(Grid1.Row, 1) & " "
    s = s & Grid1.TextMatrix(Grid1.Row, 2) & " "
    s = s & Grid1.TextMatrix(Grid1.Row, 4)
    Label2.Caption = s
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 3
    
    'Brenham Cranes
    s = "select lane.whse_num,lane.vert_loc,lane.horz_loc,lane.rack_side"
    s = s & ",position.posn_num,position.barcode,position.recv_date"
    s = s & " from lane,position"
    s = s & " where lane.id = position.laneno"
    s = s & " and position.sku = '" & msku & "'"
    s = s & " and position.lot_num = '" & mlot & "'"
    s = s & " order by lane.whse_num,position.barcode,lane.vert_loc,lane.horz_loc,lane.rack_side,position.posn_num"
    mt1 = 0: mt2 = 0: mt3 = 0: mw = 0
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds(0) <> mw Then
                Grid2.AddItem "SR-" & ds(0)
                mw = ds(0)
            End If
            If ds(0) = 1 Then mt1 = mt1 + 1
            If ds(0) = 2 Then mt2 = mt2 + 1
            If ds(0) = 3 Then mt3 = mt3 + 1
            s = ds(1) & "-" & ds(2) & "-" & ds(3) & " " & ds(4) & Chr(9)
            s = s & Right(ds(5), 5) & Chr(9)
            s = s & Format(ds(6), "m-d-yyyy")
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    mt4 = 0
    s = "select racks.aisle,racks.rack,rackpos.barcode,rackpos.recv_date"
    s = s & " from racks,rackpos"
    s = s & " where racks.id = rackpos.rackno"
    s = s & " and rackpos.sku = '" & msku & "'"
    s = s & " and rackpos.lot_num = '" & mlot & "'"
    s = s & " order by rackpos.barcode"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Grid2.AddItem "Racks"
        Do Until ds.EOF
            mt4 = mt4 + 1
            s = Trim(ds(0)) & "-" & Trim(ds(1)) & Chr(9)
            s = s & Right(ds(2), 5) & Chr(9)
            s = s & Format(ds(3), "m-d-yyyy")
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    
    ds.Close
    For i = 1 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 0) = "SR-1" Then Grid2.TextMatrix(i, 0) = "SR-1 " & mt1 & " pallets"
        If Grid2.TextMatrix(i, 0) = "SR-2" Then Grid2.TextMatrix(i, 0) = "SR-2 " & mt2 & " pallets"
        If Grid2.TextMatrix(i, 0) = "SR-3" Then Grid2.TextMatrix(i, 0) = "SR-3 " & mt3 & " pallets"
        If Grid2.TextMatrix(i, 0) = "Racks" Then Grid2.TextMatrix(i, 0) = "Racks " & mt4 & " pallets"
    Next i
    s = "Pallets: " & Format(mt1 + mt2 + mt3 + mt4, "0")
    Grid2.AddItem s
    s = "^Crane Position|^Pallet #|^Recv Date"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 1400
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1200
End Sub

Private Sub Combo1_Click()
    refresh_grid1
End Sub

Private Sub Form_Load()
    Dim cfile As String
    'cfile = "\\bbc-01-imgsvr\sharedgroups$\wd\bin\pallabels.ini"
    'cfile = "s:\wd\bin\pallabels.ini"
    cfile = "pallabels.ini"
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Line Input #1, s
            s = LCase(s)
            'If Left(s, 5) = "plant" Then Me.plantno = Right(s, Len(s) - 6)
            'If Left(s, 7) = "opcodes" Then Me.opcodes = Right(s, Len(s) - 8)
            If Left(s, 8) = "mylabels" Then Me.mylabels = Right(s, Len(s) - 9)
            'If Left(s, 9) = "oracledsn" Then Me.oradsn = Right(s, Len(s) - 10)
            'If Left(s, 10) = "oracleuser" Then Me.orauser = Right(s, Len(s) - 11)
            'If Left(s, 9) = "oraclepwd" Then Me.orapwd = Right(s, Len(s) - 10)
            'If Left(s, 11) = "productlist" Then Me.gemmies = Right(s, Len(s) - 12)
        Loop
        Close #1
    End If
    refresh_dates
    'refresh_skugrid
    'refresh_grid1
End Sub

Private Sub Form_Resize()
    If Me.Height > 2000 Then
        Grid1.Height = Me.Height - (gcolor.Height + 720)
        Grid2.Height = Grid1.Height
    End If
End Sub

Private Sub Grid1_Click()
    ycolor.Caption = Grid1.Row
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rf As String, rh As String
    rt = Me.Caption
    rh = Label2.Caption
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Call printflexgrid(Printer, Grid2, rt, rh, rf)
End Sub

Private Sub ycolor_Change()
    If Val(ycolor.Caption) > 0 Then refresh_grid2
End Sub

