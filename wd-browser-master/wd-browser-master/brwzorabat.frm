VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brwzorabat 
   Caption         =   "Oracle Production Batches"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form13"
   ScaleHeight     =   7350
   ScaleWidth      =   10890
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid skugrid 
      Height          =   2055
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3625
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4683
      _Version        =   327680
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Paste"
      Height          =   255
      Left            =   8640
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5280
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Plant:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "End Date:"
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Start Date:"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "brwzorabat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function pallet_conv(msku As String) As Integer
    Dim i As Integer
    pallet_conv = 1
    For i = 1 To skugrid.Rows - 1
        If skugrid.TextMatrix(i, 0) = msku Then
            pallet_conv = Val(skugrid.TextMatrix(i, 2))
            Exit For
        End If
    Next i
End Function

Private Sub refresh_skugrid()
    Dim s As String, f0 As String, f1 As String
    Dim f2 As String, f3 As String
    If Len(Dir(brwzplana.gemmies)) = 0 Then Exit Sub
    skugrid.Clear: skugrid.Rows = 1: skugrid.Cols = 4
    Open brwzplana.gemmies For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3
        s = f0 & Chr(9)
        s = s & f1 & Chr(9)
        s = s & f2 & Chr(9)
        s = s & f3
        skugrid.AddItem s
    Loop
    Close #1
    skugrid.FormatString = "^SKU|<Description|^PalConv|^WrpConv"
    skugrid.ColWidth(0) = 800
    skugrid.ColWidth(1) = 2800
    skugrid.ColWidth(0) = 1200
    skugrid.ColWidth(0) = 1200
    
End Sub

Private Sub refresh_tickets()
    Dim q As String, i As Integer
    Dim dsn As String, UserId As String, pwd As String
    Screen.MousePointer = 11
    dsn = brwzplana.oradsn       'dsn = "pbbcri"
    UserId = brwzplana.orauser   'UserId = "bbcgmd"
    pwd = brwzplana.orapwd       'pwd = "gmd0207"
    If AllocateODBChEnv(hEnv) <> SQL_SUCCESS Then Exit Sub
    If ConnectToDataSource(hEnv, hdbc, hstmt, dsn, UserId, pwd) <> SQL_SUCCESS Then
        i = FreeODBChEnv(hEnv)
        Exit Sub
    End If
    
    Grid1.Visible = False
    q = "select h.batch_no,h.batch_id,TO_CHAR(h.plan_start_date,'MM-DD-YYYY'),h.batch_status,"
    q = q & "h.attribute1,d.item_id,i.item_no,i.item_desc1,d.plan_qty,"
    q = q & "d.item_um"
    q = q & " from apps.gme_batch_header h, apps.gme_material_details d, apps.ic_item_mst_b i"
    If Val(Text3) = 500 Then
        q = q & " where h.plant_code in ('500', '503')"
    Else
        q = q & " where h.plant_code = '" & Format(Val(Text3), "000") & "'"
    End If
    q = q & " and h.plan_start_date >= TO_DATE('" & Format(Text1, "DD-MMM-YYYY") & "')"
    q = q & " and h.plan_start_date <= TO_DATE('" & Format(DateAdd("d", 1, Text2), "DD-MMM-YYYY") & "')"
    q = q & " and h.delete_mark = 0"
    q = q & " and h.batch_status in (1, 2)"
    q = q & " and h.batch_id = d.batch_id"
    q = q & " and d.line_type = 1"
    q = q & " and d.item_id = i.item_id"
    q = q & " and i.item_no > '000'"
    q = q & " and i.item_no < '999'"
    q = q & " order by 3, i.item_no, d.plan_qty desc, h.attribute1"
    'MsgBox q
    i = LoadGrid(Grid1, q, hstmt, 1, "")
    i = DisconnectFromDataSource(hdbc, hstmt)
    i = FreeODBChEnv(hEnv)
    Screen.MousePointer = 0
    'Grid1.FormatString = "^Batch No|^Batch ID|^Plan Start|^Status|<Location|^Item|^SKU|<Description|^Plan Qty|^UOM"
    Grid1.FormatString = "^Batch No||^Plan Start|^Status|<Location||^SKU|<Description|^Plan Qty|^UOM"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 1 '800
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 700
    Grid1.ColWidth(4) = 1800
    Grid1.ColWidth(5) = 1 '600
    Grid1.ColWidth(6) = 600
    Grid1.ColWidth(7) = 2000
    Grid1.ColWidth(8) = 900
    Grid1.ColWidth(9) = 650
    'Grid1.FillStyle = flexFillRepeat
    For i = 1 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 2) = Format(Grid1.TextMatrix(i, 2), "m-dd-yyyy")
        If Grid1.TextMatrix(i, 3) = "1" Then Grid1.TextMatrix(i, 3) = "PEND"
        If Grid1.TextMatrix(i, 3) = "2" Then Grid1.TextMatrix(i, 3) = "WIP"
        If Grid1.TextMatrix(i, 3) = "3" Then Grid1.TextMatrix(i, 3) = "CERT"
        If Grid1.TextMatrix(i, 3) = "4" Then Grid1.TextMatrix(i, 3) = "Closed"
    Next i
    'refresh_hroute
    DoEvents
    Grid1.Row = 1: Grid1.Col = 1
    Grid1.Visible = True
End Sub


Private Sub Command1_Click()
    refresh_tickets
End Sub

Private Sub Command2_Click()
    Dim i As Integer, k As Integer, j As Integer
    brwzplana.schlit.Caption = "Production Batches: " & Text1 & " thru " & Text2
    If brwzplana.Grid1.Cols = 11 Then
        brwzplana.Grid1.Cols = 17
        'brwzplana.Grid1.FormatString = "^SKU|<Description|^Plant Units|^Branch Units|^Total Units|^Branch Orders|^Sales Last 30|^Units Diff|^Pallet Diff|||^Batch Units|^Batch Pals|^Net Units|^Net Pals|^Adj Pals|^Adj Net Pals"
        brwzplana.Grid1.FormatString = "^SKU|<Description|^Plant Qty|^Branches|^Total|^Orders|^Sales|^Diff|^Pal Diff|||^Batches|^Batch Pals|^Net Units|^Net Pals|^Adj Pals|^Adj Net"
        brwzplana.Grid1.ColWidth(0) = 500
        brwzplana.Grid1.ColWidth(1) = 2800
        brwzplana.Grid1.ColWidth(2) = 900
        brwzplana.Grid1.ColWidth(3) = 900
        brwzplana.Grid1.ColWidth(4) = 800
        brwzplana.Grid1.ColWidth(5) = 700
        brwzplana.Grid1.ColWidth(6) = 700
        brwzplana.Grid1.ColWidth(7) = 700
        brwzplana.Grid1.ColWidth(8) = 700
        brwzplana.Grid1.ColWidth(9) = 1
        brwzplana.Grid1.ColWidth(10) = 1
        brwzplana.Grid1.ColWidth(11) = 800
        brwzplana.Grid1.ColWidth(12) = 1000
        brwzplana.Grid1.ColWidth(13) = 900
        brwzplana.Grid1.ColWidth(14) = 800
        brwzplana.Grid1.ColWidth(15) = 800
        brwzplana.Grid1.ColWidth(16) = 800
        
        
    End If
    For i = 1 To brwzplana.Grid1.Rows - 1
        brwzplana.Grid1.TextMatrix(i, 11) = " "
        brwzplana.Grid1.TextMatrix(i, 12) = " "
        brwzplana.Grid1.TextMatrix(i, 13) = " "
        brwzplana.Grid1.TextMatrix(i, 14) = " "
        brwzplana.Grid1.TextMatrix(i, 15) = " "
        brwzplana.Grid1.TextMatrix(i, 16) = " "
        
    Next i
    For i = 1 To Grid1.Rows - 1
        'Pallets
        j = CInt(Val(Grid1.TextMatrix(i, 8)) / pallet_conv(Grid1.TextMatrix(i, 6)))
        'Units
        'j = Val(Grid1.TextMatrix(i, 8))
        For k = 1 To brwzplana.Grid1.Rows - 1
            If Grid1.TextMatrix(i, 6) = brwzplana.Grid1.TextMatrix(k, 0) Then
                brwzplana.Grid1.TextMatrix(k, 11) = Val(brwzplana.Grid1.TextMatrix(k, 11)) + Val(Grid1.TextMatrix(i, 8))
                brwzplana.Grid1.TextMatrix(k, 12) = Val(brwzplana.Grid1.TextMatrix(k, 12)) + j
                Exit For
            End If
        Next k
    Next i
    For i = 1 To brwzplana.Grid1.Rows - 1
        brwzplana.Grid1.TextMatrix(i, 13) = Val(brwzplana.Grid1.TextMatrix(i, 11)) + Val(brwzplana.Grid1.TextMatrix(i, 7))
        brwzplana.Grid1.TextMatrix(i, 14) = Val(brwzplana.Grid1.TextMatrix(i, 12)) + Val(brwzplana.Grid1.TextMatrix(i, 8))
        brwzplana.Grid1.TextMatrix(i, 16) = Val(brwzplana.Grid1.TextMatrix(i, 12)) + Val(brwzplana.Grid1.TextMatrix(i, 8))
    Next i
    brwzplana.Grid1.FillStyle = flexFillRepeat
    For i = 1 To brwzplana.Grid1.Rows - 1
        brwzplana.Grid1.Row = i: brwzplana.Grid1.RowSel = i: brwzplana.Grid1.Col = 13: brwzplana.Grid1.ColSel = 14
        If Val(brwzplana.Grid1.TextMatrix(i, 14)) < 0 Then
            If brwzplana.Grid1.TextMatrix(i, 9) = "W" Then brwzplana.Grid1.CellBackColor = brwzplana.wcolor.BackColor
            If brwzplana.Grid1.TextMatrix(i, 9) = "B" Then brwzplana.Grid1.CellBackColor = brwzplana.bcolor.BackColor
            If brwzplana.Grid1.TextMatrix(i, 9) = "G" Then brwzplana.Grid1.CellBackColor = brwzplana.gcolor.BackColor
            If brwzplana.Grid1.TextMatrix(i, 9) = "Y" Then brwzplana.Grid1.CellBackColor = brwzplana.ycolor.BackColor
        End If
        If Val(brwzplana.Grid1.TextMatrix(i, 14)) = 0 Then brwzplana.Grid1.CellBackColor = brwzplana.bcolor.BackColor
        If Val(brwzplana.Grid1.TextMatrix(i, 14)) > 0 Then brwzplana.Grid1.CellBackColor = brwzplana.gcolor.BackColor
    Next i
    For i = 1 To brwzplana.Grid1.Rows - 1
        brwzplana.Grid1.Row = i: brwzplana.Grid1.RowSel = i: brwzplana.Grid1.Col = 16: brwzplana.Grid1.ColSel = 16
        If Val(brwzplana.Grid1.TextMatrix(i, 14)) < 0 Then
            If brwzplana.Grid1.TextMatrix(i, 9) = "W" Then brwzplana.Grid1.CellBackColor = brwzplana.wcolor.BackColor
            If brwzplana.Grid1.TextMatrix(i, 9) = "B" Then brwzplana.Grid1.CellBackColor = brwzplana.bcolor.BackColor
            If brwzplana.Grid1.TextMatrix(i, 9) = "G" Then brwzplana.Grid1.CellBackColor = brwzplana.gcolor.BackColor
            If brwzplana.Grid1.TextMatrix(i, 9) = "Y" Then brwzplana.Grid1.CellBackColor = brwzplana.ycolor.BackColor
        End If
        If Val(brwzplana.Grid1.TextMatrix(i, 14)) = 0 Then brwzplana.Grid1.CellBackColor = brwzplana.bcolor.BackColor
        If Val(brwzplana.Grid1.TextMatrix(i, 14)) > 0 Then brwzplana.Grid1.CellBackColor = brwzplana.gcolor.BackColor
    Next i
    
    brwzplana.Grid1.Row = 1
    Unload Me
    
End Sub

Private Sub Form_Load()
    Text1 = Format(Now, "m-d-yyyy")
    Text2 = Format(DateAdd("d", 7, Now), "m-d-yyyy")
    Text3 = brwzplana.brcode
    refresh_skugrid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 880
End Sub

