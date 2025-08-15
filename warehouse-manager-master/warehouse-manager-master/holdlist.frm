VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form holdlist 
   Caption         =   "Hold Product Listing"
   ClientHeight    =   10410
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13770
   LinkTopic       =   "Form19"
   ScaleHeight     =   10410
   ScaleWidth      =   13770
   StartUpPosition =   3  'Windows Default
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
      Left            =   7560
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   0
      Width           =   3735
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   2295
      Left            =   5760
      TabIndex        =   9
      Top             =   8040
      Visible         =   0   'False
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4048
      _Version        =   327680
      BackColorFixed  =   12648447
   End
   Begin MSFlexGridLib.MSFlexGrid skulist 
      Height          =   2895
      Left            =   9600
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5106
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   2055
      Left            =   0
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   3625
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2295
      Left            =   5760
      TabIndex        =   4
      Top             =   8040
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4048
      _Version        =   327680
      BackColorFixed  =   16777152
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2295
      Left            =   0
      TabIndex        =   3
      Top             =   8040
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4048
      _Version        =   327680
      BackColorFixed  =   16777152
      FocusRect       =   0
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
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid5 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   13150
      _Version        =   327680
      BackColorFixed  =   16777152
      BackColorSel    =   255
      FocusRect       =   0
   End
   Begin VB.Label bcolor 
      BackColor       =   &H00FF0000&
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7680
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label mcolor 
      BackColor       =   &H000000C0&
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7680
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Sources:"
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
      Left            =   6600
      TabIndex        =   11
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "..."
      Height          =   255
      Left            =   12120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label rcolor 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label1"
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label hcolor 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
      Begin VB.Menu printlist 
         Caption         =   "All Product List"
      End
      Begin VB.Menu prtwhslist 
         Caption         =   "All Product Warehouse List"
      End
      Begin VB.Menu prtdetails 
         Caption         =   "All Product Rack Details List"
      End
      Begin VB.Menu printracks 
         Caption         =   "Item Rack List"
      End
      Begin VB.Menu printlanes 
         Caption         =   "Item Lane List"
      End
      Begin VB.Menu ppaltots 
         Caption         =   "Pallet Totals"
      End
   End
   Begin VB.Menu holdmenu 
      Caption         =   "E&dit"
      Begin VB.Menu delhold 
         Caption         =   "Remove From Hold"
      End
      Begin VB.Menu addholdprod 
         Caption         =   "Add Hold Product"
      End
      Begin VB.Menu edspallet 
         Caption         =   "Edit Start Pallet"
      End
      Begin VB.Menu edepallet 
         Caption         =   "Edit End Pallet"
      End
      Begin VB.Menu edsource 
         Caption         =   "Edit Source"
      End
      Begin VB.Menu batonhand 
         Caption         =   "View Batch Inventory"
      End
      Begin VB.Menu hshist 
         Caption         =   "View Hold Status History"
      End
      Begin VB.Menu edrackmenu 
         Caption         =   "Racks"
         Visible         =   0   'False
         Begin VB.Menu markrackhold 
            Caption         =   "Mark Rack - On Hold"
         End
         Begin VB.Menu remrackhold 
            Caption         =   "Remove Rack Hold"
         End
      End
      Begin VB.Menu edsrmenu 
         Caption         =   "Cranes"
         Visible         =   0   'False
         Begin VB.Menu marksrhold 
            Caption         =   "Mark Lane - On Hold"
         End
         Begin VB.Menu remsrhold 
            Caption         =   "Remove Lane Hold"
         End
      End
   End
   Begin VB.Menu schmenu 
      Caption         =   "Schedule"
      Begin VB.Menu psched 
         Caption         =   "Production Schedule"
      End
   End
End
Attribute VB_Name = "holdlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim srrefresh As Boolean                            'jv120415

Private Sub build_hold_hist()
    Dim i As Integer, zid As Long
    For i = 1 To Grid5.Rows - 1
        Grid5.Row = i
        zid = Val(Grid5.TextMatrix(i, 0))                           'jv010616
        Call post_hold_log(zid, "Current")                           'jv010616
    Next i
End Sub

Sub refresh_sources()
    Dim ds As ADODB.Recordset, s As String
    Combo1.Clear
    Combo1.AddItem "ALL"
    s = "select hsource, count(*) from holdlist group by hsource order by hsource"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If LCase(ds(0)) = "schedule" Then           'jv121815
                Combo1.AddItem "TEST HOLD"              'jv121815
            Else                                        'jv121815
                Combo1.AddItem ds(0)
            End If                                      'jv121815
            ds.MoveNext
        Loop
    End If
    ds.Close
    Combo1.ListIndex = 0
End Sub

Sub refresh_holdlist()                                  'jv040615
    Dim ds As ADODB.Recordset, s As String, t As String
    Dim ss As ADODB.Recordset, pdesc As String, hrow As Boolean, crow As Boolean
    Screen.MousePointer = 11
    Grid5.Redraw = False
    Grid5.Clear: Grid5.Rows = 1: Grid5.Cols = 14
    If Combo1 = "ALL" Then
        s = "select * from holdlist order by  sku, lot_num, opcode, epallet"
    Else
        If Combo1 = "TEST HOLD" Then
            s = "select * from holdlist where hsource in ('" & Combo1 & "', 'Schedule') order by  sku, lot_num, opcode, epallet"
        Else
            s = "select * from holdlist where hsource = '" & Combo1 & "' order by  sku, lot_num, opcode, epallet"
        End If
    End If
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            pdesc = skurec(Val(ds!sku)).prodname
            s = ds!id & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & pdesc & Chr(9)
            s = s & ds!lot_num & Chr(9)
            s = s & ds!opcode & Chr(9)
            s = s & ds!spallet & Chr(9)
            s = s & ds!epallet & Chr(9)
            If LCase(ds!hsource) = "schedule" Then                  'jv121815
                s = s & "TEST HOLD" & Chr(9)                        'jv121815
            Else                                                    'jv121815
                s = s & ds!hsource & Chr(9)
            End If                                                  'jv121815
            's = s & r12lot(ds!lot_num) & " " & ds!opcode
            If ds!lot_num = "16060" Then                            'jv030416
                s = s & "022918" & ds!opcode                        'jv030416
            Else
                s = s & r12lot(ds!lot_num) & ds!opcode                      'jv052515
            End If
            s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9)
            s = s & ds!userid & Chr(9)
            s = s & ds!holddate
            Grid5.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    'If Grid5.Rows > 1 Then
    '    For i = 1 To Grid5.Rows - 1
    '        If Grid5.TextMatrix(i, 7) = "Schedule" Then
    '            s = "select * from prodrcv where id = " & Grid5.TextMatrix(i, 0)
    '            Set ds = db.OpenRecordset(s)
    '            If ds.BOF = False Then
    '                ds.MoveFirst
    '                's = Format(DateAdd("yyyy", 2, ds!proddate), "MMddyy") & " " & ds!sp_flag
    '                'Grid5.TextMatrix(i, 8) = s
    '                Grid5.TextMatrix(i, 9) = Format(ds!recdate1, "M-dd-yyyy")
    '                Grid5.TextMatrix(i, 10) = Format(ds!recdate2, "M-dd-yyyy")
    '                Grid5.TextMatrix(i, 11) = Format(ds!recdate3, "M-dd-yyyy")
    '            End If
    '            ds.Close
    '        End If
    '    Next i
    'End If
    'db.Close
        
    Grid5.FillStyle = flexFillRepeat
    If Grid5.Rows > 2 Then
        's = Grid5.TextMatrix(1, 1)
        s = Grid5.TextMatrix(1, 1) & Grid5.TextMatrix(1, 3) & Grid5.TextMatrix(1, 4)        'jv010816
        t = Grid5.TextMatrix(1, 1)
        For i = 1 To Grid5.Rows - 1
            If Grid5.TextMatrix(i, 1) <> t Then
                crow = Not crow
                t = Grid5.TextMatrix(i, 1)
            End If
            If Grid5.TextMatrix(i, 1) & Grid5.TextMatrix(i, 3) & Grid5.TextMatrix(i, 4) <> s Then   'jv010816
                hrow = Not hrow
                's = Grid5.TextMatrix(i, 1)
                s = Grid5.TextMatrix(i, 1) & Grid5.TextMatrix(i, 3) & Grid5.TextMatrix(i, 4)    'jv010816
            End If
            If crow = True Then
                Grid5.Row = i: Grid5.RowSel = i
                Grid5.Col = 1: Grid5.ColSel = Grid5.Cols - 1
                Grid5.CellBackColor = hcolor.BackColor
                'Grid5.CellFontBold = True
                'Grid5.CellBackColor = Grid5.BackColorFixed
            End If
            If hrow = True Then
                Grid5.Row = i: Grid5.RowSel = i
                Grid5.Col = 3: Grid5.ColSel = 7
                Grid5.CellForeColor = mcolor.BackColor
                'Grid5.CellFontBold = True
            Else
                Grid5.Row = i: Grid5.RowSel = i
                Grid5.Col = 3: Grid5.ColSel = 7
                Grid5.CellForeColor = bcolor.BackColor
                'Grid5.CellFontBold = True
            End If
            Grid5.Row = i: Grid5.RowSel = i
            Grid5.Col = 3: Grid5.ColSel = 7
            'Grid5.CellFontBold = True
        Next i
        Grid5.Row = 1
        Grid5_Click
    End If
    
    's = "^ID|^SKU|<Product|^Lot|^OpCode|^Start|^End|^Source|^R12Lot|^Date1|^Date2|^Date3|<Userid|<DateTime"
    s = "^ID|^SKU|<Product|^Lot|^OpCode|^Start|^End|^Source|^R12Lot||||<Userid|<DateTime"
    Grid5.FormatString = s
    Grid5.ColWidth(0) = 1000
    Grid5.ColWidth(1) = 800
    Grid5.ColWidth(2) = 3000
    Grid5.ColWidth(3) = 800
    Grid5.ColWidth(4) = 800
    Grid5.ColWidth(5) = 800
    Grid5.ColWidth(6) = 800
    Grid5.ColWidth(7) = 1600
    Grid5.ColWidth(8) = 1000
    Grid5.ColWidth(9) = 0 '900
    Grid5.ColWidth(10) = 0 '900
    Grid5.ColWidth(11) = 0 '900
    Grid5.ColWidth(12) = 1400
    Grid5.ColWidth(13) = 1400
    Grid5.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Function wd_lotnum(pdate As String) As String
    Dim sdate As String, s As String
    pdate = Format(pdate, "m-d-yyyy")
    sdate = "1-1-" & Right(pdate, 4)
    s = Format(DateDiff("d", sdate, pdate) + 1, "000")
    s = Right(pdate, 2) & s
    wd_lotnum = s
End Function

Private Sub refresh_cs5_locks()
    Dim db5 As ADODB.Connection, ds5 As ADODB.Recordset, s As String, bsku As String, blot As String
    Dim wlotdate As String                                                      'jv091615
    Dim bcode As String
    bsku = Grid5.TextMatrix(Grid5.Row, 1)
    blot = Grid5.TextMatrix(Grid5.Row, 3)
    bcode = Grid5.TextMatrix(Grid5.Row, 4)
    Set db5 = CreateObject("ADODB.Connection")
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 7
    'db5.Open "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
    db5.Open "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=sywms;PWD=!Sylacauga_WMS1907"
    s = "SELECT tLocationData.sLocationID, "
    s = s & "tLaneData.iLevel, tLaneData.iRow, tLaneData.iBlock, "
    s = s & "tContainerLocationData.iLocationID, "
    s = s & "tInventoryData.nQuantity, "
    's = s & "tLotData.dtProduction, "
    s = s & "tLotData.dtExpiration, "                                           'jv091615
    s = s & "tItemMaster.sItemID, tItemMaster.sItemDescription, tLotData.sQAHoldReason,"
    s = s & "tLaneLock.sDescription, tlaneLock.iLocked, count(*) "
    s = s & "FROM tLocationData, tLaneData, tContainerLocationData, tInventoryData, "
    s = s & "tLotData, tItemMaster, tLaneLock"
    s = s & " WHERE tLaneData.iLocationID = tLocationData.iLocationID"
    s = s & " AND tContainerLocationData.iLocationID = tLaneData.iLocationID"
    s = s & " AND tLaneLock.iLaneSysID = tLaneData.iLocationID"
    s = s & " AND tInventoryData.iContainerDataSysID = tContainerLocationData.iContainerDataSysID"
    s = s & " AND tLotData.iLotDataSysID = tInventoryData.iLotDataSysID"
    s = s & " AND tItemMaster.iItemMasterSysID = tLotData.iItemMasterSysID"
    s = s & " and tItemMaster.sItemID = '" & bsku & "-" & bcode & "'"
    s = s & " GROUP BY tLocationData.sLocationID, "
    s = s & "tLaneData.iLevel, tLaneData.iRow, tLaneData.iBlock, "
    s = s & "tContainerLocationData.iLocationID, "
    s = s & "tInventoryData.nQuantity, "
    's = s & "tLotData.dtProduction, "
    s = s & "tLotData.dtExpiration, "                                           'jv091615
    s = s & "tItemMaster.sItemID, tItemMaster.sItemDescription, tLotData.sQAHoldReason,"
    s = s & " tLaneLock.sDescription, tLaneLock.iLocked"
    s = s & " ORDER BY tLocationData.sLocationID " ', tContainerLocationData.iPosition"
    Set ds5 = db5.Execute(s)
    If ds5.BOF = False Then
        ds5.MoveFirst
        Do Until ds5.EOF
            wlotdate = Format(DateAdd("yyyy", -2, ds5(6)), "M-d-yyyy")          'jv091615
            s = "CS5" & Chr(9)
            s = s & Trim(ds5(0)) & Chr(9)
            's = s & wd_lotnum(ds5(6)) & Chr(9)
            s = s & wd_lotnum(wlotdate) & Chr(9)                                'jv091615
            's = s & Format(ds5(6), "m-d-yyyy") & Chr(9)
            s = s & wlotdate & Chr(9)                                           'jv091615
            s = s & ds5(12) & Chr(9)
            s = s & ds5(5) & Chr(9)
            If ds5(11) = 0 Then
                s = s & "Unlocked"
            Else
                s = s & ds5(10)
            End If
            '& ds5(9)
            'If Val(blot) = 0 Or blot = wd_lotnum(ds5(6)) Then
            If Val(blot) = 0 Or blot = wd_lotnum(wlotdate) Then                 'jv091615
                Grid3.AddItem s
            End If
            ds5.MoveNext
        Loop
    End If
    ds5.Close
    If Grid3.Rows > 1 Then
        Grid3.FillStyle = flexFillRepeat
        For i = 1 To Grid3.Rows - 1
            If Grid3.TextMatrix(i, 6) = "Unlocked" Then
                Grid3.Row = i: Grid3.RowSel = i
                Grid3.Col = 1: Grid3.ColSel = 6
                Grid3.CellBackColor = rcolor.BackColor
            End If
        Next i
        Grid3.Row = 1
    End If
    db5.Close
    Grid3.FormatString = "^Whs|^Location|^Lot|^Date|^Pallets|^Size|^Lock Status"
    Grid3.ColWidth(0) = 600 '1000
    Grid3.ColWidth(1) = 1000
    Grid3.ColWidth(2) = 800 '1000
    Grid3.ColWidth(3) = 1000
    Grid3.ColWidth(4) = 1000
    Grid3.ColWidth(5) = 800 '1000
    Grid3.ColWidth(6) = 2000 '1000
End Sub

Private Sub refresh_cs5()
    Dim db5 As ADODB.Connection, ds5 As ADODB.Recordset, s As String, bsku As String, blot As String
    Dim bcode As String
    bsku = Grid5.TextMatrix(Grid5.Row, 1)
    blot = Grid5.TextMatrix(Grid5.Row, 3)
    bcode = Grid5.TextMatrix(Grid5.Row, 4)
    Set db5 = CreateObject("ADODB.Connection")
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 7
    'db5.Open "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
    db5.Open "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=sywms;PWD=!Sylacauga_WMS1907"
    s = "SELECT tLocationData.sLocationID, "
    s = s & "tLaneData.iLevel, tLaneData.iRow, tLaneData.iBlock, "
    s = s & "tContainerLocationData.iLocationID, "
    s = s & "tInventoryData.nQuantity, "
    s = s & "tLotData.dtProduction, "
    s = s & "tItemMaster.sItemID, tItemMaster.sItemDescription, tLotData.sQAHoldReason, count(*) "
    s = s & "FROM tLocationData, tLaneData, tContainerLocationData, tInventoryData, "
    s = s & "tLotData, tItemMaster"
    s = s & " WHERE tLaneData.iLocationID = tLocationData.iLocationID"
    s = s & " AND tContainerLocationData.iLocationID = tLaneData.iLocationID"
    s = s & " AND tInventoryData.iContainerDataSysID = tContainerLocationData.iContainerDataSysID"
    s = s & " AND tLotData.iLotDataSysID = tInventoryData.iLotDataSysID"
    s = s & " AND tItemMaster.iItemMasterSysID = tLotData.iItemMasterSysID"
        
    's = s & " and tItemMaster.sItemID >= '" & bsku & "'"
    's = s & " and tItemMaster.sItemID < '" & bsku & "ZZZZ'"
    s = s & " and tItemMaster.sItemID = '" & bsku & "-" & bcode & "'"
    s = s & " GROUP BY tLocationData.sLocationID, "
    s = s & "tLaneData.iLevel, tLaneData.iRow, tLaneData.iBlock, "
    s = s & "tContainerLocationData.iLocationID, "
    s = s & "tInventoryData.nQuantity, "
    s = s & "tLotData.dtProduction, "
    s = s & "tItemMaster.sItemID, tItemMaster.sItemDescription, tLotData.sQAHoldReason"
    s = s & " ORDER BY tLocationData.sLocationID " ', tContainerLocationData.iPosition"
    Set ds5 = db5.Execute(s)
    If ds5.BOF = False Then
        ds5.MoveFirst
        Do Until ds5.EOF
            s = "CS5" & Chr(9)
            s = s & Trim(ds5(0)) & Chr(9)
            s = s & wd_lotnum(ds5(6)) & Chr(9)
            s = s & Format(ds5(6), "m-d-yyyy") & Chr(9)
            s = s & ds5(10) & Chr(9)
            s = s & ds5(5) & Chr(9) & ds5(9)
            If Val(blot) = 0 Or blot = wd_lotnum(ds5(6)) Then
                Grid3.AddItem s
            End If
            ds5.MoveNext
        Loop
    End If
    ds5.Close
    If Grid3.Rows > 1 Then
        For i = 1 To Grid3.Rows - 1
            s = "select * from vAllLanes_1033 where location = '" & Grid3.TextMatrix(i, 1) & "'"
            Set ds5 = db5.Execute(s)
            If ds5.BOF = False Then
                ds5.MoveFirst
                'If ds5(9) = "Unlocked" Then
                '    Grid3.TextMatrix(i, 6) = "Unlocked"
                'Else
                    Grid3.TextMatrix(i, 6) = Trim(ds5(9))
                'End If
            End If
            ds5.Close
        Next i
        Grid3.FillStyle = flexFillRepeat
        For i = 1 To Grid3.Rows - 1
            If Grid3.TextMatrix(i, 6) = "Unlocked" Then
                Grid3.Row = i: Grid3.RowSel = i
                Grid3.Col = 1: Grid3.ColSel = 6
                Grid3.CellBackColor = rcolor.BackColor
            End If
        Next i
        Grid3.Row = 1
    End If
    db5.Close
    Grid3.FormatString = "^Whs|^Location|^Lot|^Date|^Pallets|^Size|^Lock Status"
    Grid3.ColWidth(0) = 600 '1000
    Grid3.ColWidth(1) = 1000
    Grid3.ColWidth(2) = 800 '1000
    Grid3.ColWidth(3) = 1000
    Grid3.ColWidth(4) = 1000
    Grid3.ColWidth(5) = 800 '1000
    Grid3.ColWidth(6) = 2000 '1000
End Sub

Private Sub refresh_cs5_new()
    Dim db5 As ADODB.Connection, ds5 As ADODB.Recordset, s As String, bsku As String, blot As String
    Dim wlotdate As String, slot As String                                                      'jv091615
    Dim bcode As String
    bsku = Grid5.TextMatrix(Grid5.Row, 1)
    blot = Grid5.TextMatrix(Grid5.Row, 8)
    bcode = Grid5.TextMatrix(Grid5.Row, 4)
    slot = bsku & "-" & bcode & "-" & Left(blot, 6)
    Set db5 = CreateObject("ADODB.Connection")
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 7
    db5.Open "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=sywms;PWD=!Sylacauga_WMS1907"
    s = "select * from vAllInventory_1033 where lot = '" & slot & "' order by lpn"
    'MsgBox s
    Set ds5 = db5.Execute(s)
    If ds5.BOF = False Then
        ds5.MoveFirst
        Do Until ds5.EOF = True
            s = "CS5" & Chr(9)
            s = s & ds5!location & Chr(9)
            's = s & ds5!lot & Chr(9)
            s = s & ds5!lpn & Chr(9)
            s = s & ds5(14) & Chr(9) '!lotproduction & Chr(9)
            s = s & "1" & Chr(9)
            s = s & ds5!quantity & Chr(9)
            s = s & ds5!lock
            Grid3.AddItem s
            ds5.MoveNext
        Loop
    End If
    ds5.Close
    If Grid3.Rows > 1 Then
        Grid3.FillStyle = flexFillRepeat
        For i = 1 To Grid3.Rows - 1
            'If Grid3.TextMatrix(i, 6) = "Unlocked" Then
            If Grid3.TextMatrix(i, 6) < "1" Then
                Grid3.Row = i: Grid3.RowSel = i
                Grid3.Col = 1: Grid3.ColSel = 6
                Grid3.CellBackColor = rcolor.BackColor
            End If
        Next i
        Grid3.Row = 1
    End If
    db5.Close
    Grid3.FormatString = "^Whs|^Location|^BarCode|^Date|^Pallets|^Size|^Lock Status"
    Grid3.ColWidth(0) = 600 '1000
    Grid3.ColWidth(1) = 1000
    Grid3.ColWidth(2) = 1900 '800 '1000
    Grid3.ColWidth(3) = 1000
    Grid3.ColWidth(4) = 1000
    Grid3.ColWidth(5) = 800 '1000
    Grid3.ColWidth(6) = 1000
End Sub
Private Sub Refresh_racks()
    Dim ds As ADODB.Recordset, s As String, q As Integer, ss As ADODB.Recordset
    Dim i As Integer, psku As String, pcode As String, ps As String, pe As String
    Dim bc1 As String, bc2 As String
    Screen.MousePointer = 11
    i = Grid5.Row
    markrackhold.Enabled = False
    remrackhold.Enabled = False
    psku = Grid5.TextMatrix(i, 1)
    s = Grid5.TextMatrix(i, 8)                              'jv052515
    If Len(s) = 7 Then                                      'jv052515
        pdate = Left(s, 6) & " " & Right(s, 1) & " "        'jv052515
    Else                                                    'jv052515
        If Len(s) = 8 Then                                  'jv052515
            pdate = Left(s, 6) & " " & Right(s, 2)          'jv052515
        Else                                                'jv052515
            pdate = s                                       'jv052515
        End If                                              'jv052515
    End If                                                  'jv052515
    If Len(psku) = 3 Then
        ps = psku & " " & pdate & Grid5.TextMatrix(i, 5)        'jv052515
        pe = psku & " " & pdate & Grid5.TextMatrix(i, 6)        'jv052515
    Else
        ps = psku & pdate & Grid5.TextMatrix(i, 5)          'jv082415
        pe = psku & pdate & Grid5.TextMatrix(i, 6)          'jv082415
    End If
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    s = "select id, aisle, rack, hold from racks where id in "
    s = s & "(select rackno from rackpos where barcode >= '" & ps & "'"
    s = s & " and barcode <= '" & pe & "')"
    s = s & " order by aisle, rack"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            q = 0: bc1 = "": bc2 = ""
            s = "select rackno, barcode,count(*) from rackpos where rackno = " & ds!id & ""
            s = s & " and barcode >= '" & ps & "' and barcode <= '" & pe & "'"
            s = s & " group by rackno, barcode order by barcode"
            Set ss = Wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                bc1 = Right(ss(1), 3)
                Do Until ss.EOF
                    q = q + ss(2)
                    bc2 = Right(ss(1), 3)
                    ss.MoveNext
                Loop
            End If
            ss.Close
            s = ds!id & Chr(9) & Trim(ds!aisle) & Chr(9) & Trim(ds!rack) & Chr(9) & q & Chr(9) & ds!hold
            If bc2 = bc1 Then
                s = s & Chr(9) & bc1
            Else
                s = s & Chr(9) & bc1 & " .. " & bc2
            End If
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 4) <> "1" Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = rcolor.BackColor
            End If
        Next i
        Grid1.Row = 1
    End If
    Grid1.FormatString = "^ID|^Aisle|^Rack|^Qty|^Hold Flag|^Pallet Labels"
    Grid1.ColWidth(0) = 700
    Grid1.ColWidth(1) = 700
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 700
    Grid1.ColWidth(4) = 900
    Grid1.ColWidth(5) = 1200
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_cranes()
    Dim ds As ADODB.Recordset, s As String, q As Integer, ss As ADODB.Recordset
    Dim i As Integer, psku As String, pcode As String, ps As String, pe As String
    Dim bcs As String
    'If Form1.plantno = "52" Then refresh_cs5_locks
    If Form1.plantno = "52" Then refresh_cs5_new
    If Form1.plantno <> "50" Then Exit Sub
    Screen.MousePointer = 11
    i = Grid5.Row
    marksrhold.Enabled = False
    remsrhold.Enabled = False
    psku = Grid5.TextMatrix(i, 1)
    s = Grid5.TextMatrix(i, 8)                              'jv052515
    If Len(s) = 7 Then                                      'jv052515
        pdate = Left(s, 6) & " " & Right(s, 1) & " "        'jv052515
    Else                                                    'jv052515
        If Len(s) = 8 Then                                  'jv052515
            pdate = Left(s, 6) & " " & Right(s, 2)          'jv052515
        Else                                                'jv052515
            pdate = s                                       'jv052515
        End If                                              'jv052515
    End If                                                  'jv052515
    If Len(psku) = 3 Then
        ps = psku & " " & pdate & Grid5.TextMatrix(i, 5)        'jv052515
        pe = psku & " " & pdate & Grid5.TextMatrix(i, 6)        'jv052515
    Else
        ps = psku & pdate & Grid5.TextMatrix(i, 5)        'jv082415
        pe = psku & pdate & Grid5.TextMatrix(i, 6)        'jv082415
    End If
    Grid2.Redraw = False
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 8
    s = "select id, whse_num, zone_num, vert_loc, horz_loc, rack_side, qty, lane_status from lane"
    s = s & " where id in (select laneno from position where barcode >= '" & ps & "'"
    s = s & " and barcode <= '" & pe & "')"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            q = 0: bcs = ""
            s = "select laneno, barcode, count(*) from position where laneno = " & ds!id & ""
            s = s & " and barcode >= '" & ps & "' and barcode <= '" & pe & "'"
            s = s & " group by laneno, barcode order by barcode"
            Set ss = Wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                Do Until ss.EOF
                    q = q + ss(2)
                    If bcs > " " Then
                        bcs = bcs & "  " & Right(ss(1), 3)
                    Else
                        bcs = Right(ss(1), 3)
                    End If
                    ss.MoveNext
                Loop
            End If
            ss.Close
        
            If ds!whse_num < 5 Then
                s = ds!id & Chr(9) & ds!whse_num & Chr(9)
            Else
                s = ds!id & Chr(9) & ds!zone_num & Chr(9)
            End If
            s = s & ds!vert_loc & Chr(9) & ds!horz_loc & Chr(9) & ds!rack_side & Chr(9)
            's = s & ds!qty & Chr(9) & ds!lane_status
            s = s & q & Chr(9) & ds!lane_status & Chr(9) & bcs
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid2.Rows > 1 Then
        Grid2.FillStyle = flexFillRepeat
        For i = 1 To Grid2.Rows - 1
            If Grid2.TextMatrix(i, 6) <> "H" Then
                Grid2.Row = i: Grid2.RowSel = i
                Grid2.Col = 1: Grid2.ColSel = Grid2.Cols - 1
                Grid2.CellBackColor = rcolor.BackColor
            End If
        Next i
        Grid2.Row = 1
    End If
    Grid2.FormatString = "^ID|^SR|^Vert|^Horz|^Side|^Qty|^Status|^Pallet Labels"
    Grid2.ColWidth(0) = 800
    Grid2.ColWidth(1) = 800
    Grid2.ColWidth(2) = 800
    Grid2.ColWidth(3) = 800
    Grid2.ColWidth(4) = 800
    Grid2.ColWidth(5) = 800
    Grid2.ColWidth(6) = 800
    Grid2.ColWidth(7) = 1600
    Grid2.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub addholdprod_Click()
    Dim psku As String, plot As String, pcode As String, psrc As String
    Dim sp As String, ep As String
    Dim s As String, i As Long
    i = Grid5.Row
    If Val(Grid5.TextMatrix(i, 0)) > 0 Then
        psku = Grid5.TextMatrix(i, 1)
        plot = Grid5.TextMatrix(i, 3)
        pcode = Grid5.TextMatrix(i, 4)
    End If
    psku = InputBox("SKU:", "SKU Number....", psku)
    If Len(psku) = 0 Then Exit Sub
    plot = InputBox("Lot:", "Lot Number....", plot)
    If Len(plot) = 0 Then Exit Sub
    pcode = InputBox("OP Code:", "OP code..", pcode)
    If Len(pcode) = 0 Then Exit Sub
    
    sp = InputBox("Start Pallet:", "Starting Pallet...", "001")         'jv121815
    sp = UCase(sp)                                                      'jv121815
    If Len(sp) = 0 Then Exit Sub                                        'jv121815
    If Val(sp) < 0 Or Val(sp) > 999 Then Exit Sub                       'jv121815
    If Val(sp) = 0 And sp <> "EOR" Then Exit Sub                        'jv121815
    If sp <> "EOR" Then sp = Format(Val(sp), "000")                     'jv121815
    
    If sp = "EOR" Then                                                  'jv121815
        ep = "EOR"                                                      'jv121815
    Else                                                                'jv121815
        ep = InputBox("End Pallet:", "Ending Pallet...", sp)            'jv121815
        If Len(ep) = 0 Then Exit Sub                                    'jv121815
        ep = UCase(ep)                                                  'jv121815
        If ep <> "EOR" And Val(ep) < Val(sp) Then Exit Sub              'jv121815
        If Val(ep) = 0 And ep <> "EOR" Then Exit Sub                    'jv121815
        If ep <> "EOR" Then ep = Format(Val(ep), "000")                 'jv121815
    End If                                                              'jv121815
    
    MsgBox sp & "..." & ep, vbOKOnly + vbInformation, "Pallet Numbers to be used..."
    
    psrc = "HoldEdit"                                           'jv120715
    psrc = InputBox("Source:", "Source..", psrc)                'jv120715
    If Len(psrc) = 0 Then Exit Sub                              'jv120715
    If Len(psrc) > 20 Then psrc = Left(psrc, 20)                'jv120715
    
    i = wd_seq("HoldList")                                      'jv042015
    s = "Insert into holdlist (id, sku, lot_num, opcode, spallet, epallet, hsource, userid, holddate) values (" & i
    s = s & ", '" & psku & "', '" & plot & "', '" & pcode & "', '" & sp & "', '" & ep & "', '" & psrc & "', '" & WDUserId & "', '" & Format(Now, "yyMMdd hh:mm:ss") & "')"
    Wdb.Execute s
    refresh_holdlist
    Call post_hold_log(i, "Add")                        'jv010616
End Sub

Private Sub batonhand_Click()
    Dim s As String
    s = Grid5.TextMatrix(Grid5.Row, 1)
    If Len(s) = 3 Then s = s & " "
    s = s & Left(Grid5.TextMatrix(Grid5.Row, 8), 6)
    If Len(Grid5.TextMatrix(Grid5.Row, 4)) = 1 Then
        s = s & " " & Grid5.TextMatrix(Grid5.Row, 4) & " "
    Else
        s = s & Grid5.TextMatrix(Grid5.Row, 4)
    End If
    tktonhand.bbarcode = s
    tktonhand.bproduct = Grid5.TextMatrix(Grid5.Row, 2)
    tktonhand.Show
End Sub

Private Sub Combo1_Click()
    refresh_holdlist
End Sub

Private Sub Command1_Click()
    'refresh_holdlist
    refresh_sources                                             'jv121515
End Sub

Private Sub delhold_Click()                                     'jv040615
    Dim s As String, i As Integer, k As Integer, ps As String, wrxflag As Boolean, wrxfile As String
    Dim d As daimessagerec, zid As Long, ds As ADODB.Recordset
    i = Grid5.Row
    If Val(Grid5.TextMatrix(i, 0)) = 0 Then Exit Sub
    
    ' Check if row is on QC Hold, display additional confirmation before proceeding.
    If UCase(Grid5.TextMatrix(i, 7)) = "QC HOLD" Then
        Dim confirmResponse As Integer
        Dim confirmMessage As String
        confirmMessage = "You are about to remove an item on QC HOLD. Are you sure?" & vbCrLf & vbCrLf & _
            "Sku: " & Grid5.TextMatrix(i, 1) & vbCrLf & _
            "R12 Lot: " & Grid5.TextMatrix(i, 8) & vbCrLf & _
            "Start: " & Grid5.TextMatrix(i, 5) & vbCrLf & _
            "End: " & Grid5.TextMatrix(i, 6) & vbCrLf
        
        confirmResponse = MsgBox(confirmMessage, vbYesNo + vbDefaultButton2 + vbApplicationModal + vbMsgBoxSetForeground, "QC HOLD")
        If confirmResponse = 7 Then
            Exit Sub
        End If
    End If
    zid = Val(Grid5.TextMatrix(i, 0))                           'jv010616
    Call post_hold_log(zid, "Delete")                           'jv010616
    'Clear Rack Hold Flag
    If Grid1.Rows > 1 Then
        For k = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(k, 4) = "1" Then
                ps = "Remove Hold Flag from rack " & Grid1.TextMatrix(k, 1) & " " & Grid1.TextMatrix(k, 2) & "?"
                If MsgBox(ps, vbYesNo + vbQuestion, "remove rack on hold flag.....") = vbYes Then
                    s = "update racks set hold = '0' where id = " & Grid1.TextMatrix(k, 0)
                    Wdb.Execute s
                End If
            End If
        Next k
    End If
    'Clear Lane Hold Status
    If Grid2.Rows > 1 Then
        srrefresh = False                                           'jv120415
        For k = 1 To Grid2.Rows - 1                                 'jv050615
            Grid2.Row = k                                           'jv050615
            remsrhold_Click                                         'jv050615
        Next k                                                      'jv050615
        'refresh_cranes                                              'jv120415
        srrefresh = True                                            'jv120415
    End If
    s = "delete from holdlist where id = " & Grid5.TextMatrix(i, 0)
    Wdb.Execute s
    refresh_holdlist
    DoEvents
    If i < Grid5.Rows - 1 Then
        Grid5.Row = i
        Grid5_Click                         'Refresh Racks and Cranes
    End If
End Sub

Private Sub edepallet_Click()
    Dim s As String, ds As ADODB.Recordset, i As Integer, pno As String, wrxflag As Boolean, wrxfile As String
    Dim d As daimessagerec, zid As Long, k As Integer
    i = Grid5.Row
    If Val(Grid5.TextMatrix(i, 0)) = 0 Then Exit Sub
    pno = Grid5.TextMatrix(i, 6)
    pno = "EOR"
    pno = InputBox("End Pallet #:", "End Pallet #.....", pno)
    If Len(pno) = 0 Then Exit Sub
    If Val(pno) > 0 Then
        pno = Format(Val(pno), "000")
    Else
        pno = "EOR"
    End If
    If Len(pno) <> 3 Then Exit Sub
    If pno < Grid5.TextMatrix(i, 5) Then Exit Sub
    zid = Val(Grid5.TextMatrix(i, 0))                           'jv010616
    Call post_hold_log(zid, "End Pallet")                       'jv010616
    s = "Update holdlist set epallet = '" & pno & "'"
    If Grid5.TextMatrix(i, 12) <> "WMS" Then
        s = s & ", userid = '" & WDUserId & "', holddate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
        Grid5.TextMatrix(i, 12) = WDUserId
        Grid5.TextMatrix(i, 13) = Format(Now, "yyMMdd hh:mm:ss")
    End If
    s = s & " Where id = " & Grid5.TextMatrix(i, 0)
    Wdb.Execute s
    'Clear Lane Hold Status to previous range
    If Grid2.Rows > 1 Then
        srrefresh = False                                           'jv120415
        For k = 1 To Grid2.Rows - 1                                 'jv050615
            Grid2.Row = k                                           'jv050615
            remsrhold_Click                                         'jv050615
        Next k                                                      'jv050615
        'refresh_cranes                                              'jv120415
        srrefresh = True                                            'jv120415
    End If
    Grid5.TextMatrix(i, 6) = pno
    Grid5_Click                                 'Refresh Racks and Cranes
    'Add Lane Hold Status to new range
    If Grid2.Rows > 1 Then
        srrefresh = False                                           'jv120415
        For k = 1 To Grid2.Rows - 1                                 'jv050615
            Grid2.Row = k                                           'jv050615
            marksrhold_Click                                        'jv050615
        Next k                                                      'jv050615
        'refresh_cranes                                              'jv120415
        srrefresh = True                                            'jv120415
    End If
    Grid5_Click                                 'Refresh Racks and Cranes
    Call post_hold_log(zid, "End Pallet")                       'jv010616
End Sub

Private Sub edsource_Click()
    Dim psource As String, s As String, zid As Long
    If Val(Grid5.TextMatrix(Grid5.Row, 0)) = 0 Then Exit Sub
    psource = Grid5.TextMatrix(Grid5.Row, 7)
    psource = InputBox("Source:", "Source.....", psource)
    If Len(psource) = 0 Then Exit Sub
    zid = Val(Grid5.TextMatrix(Grid5.Row, 0))                   'jv010616
    Call post_hold_log(zid, "Source")                           'jv010616
    If Len(psource) > 20 Then psource = Left(psource, 20)
    s = "Update holdlist set hsource = '" & psource & "' where id = " & Grid5.TextMatrix(Grid5.Row, 0)
    Grid5.TextMatrix(Grid5.Row, 7) = psource
    Wdb.Execute s
    Call post_hold_log(zid, "Source")                           'jv010616
End Sub

Private Sub edspallet_Click()
    Dim s As String, ds As ADODB.Recordset, i As Integer, pno As String, wrxflag As Boolean, wrxfile As String
    Dim d As daimessagerec, zid As Long, k As Integer
    i = Grid5.Row
    If Val(Grid5.TextMatrix(i, 0)) = 0 Then Exit Sub
    pno = Grid5.TextMatrix(i, 5)
    pno = InputBox("Start Pallet #:", "Start Pallet #.....", pno)
    If Len(pno) = 0 Then Exit Sub
    If Val(pno) > 0 Then
        pno = Format(Val(pno), "000")
    Else
        pno = "EOR"
    End If
    If Len(pno) <> 3 Then Exit Sub
    If pno > Grid5.TextMatrix(i, 6) Then Exit Sub
    zid = Val(Grid5.TextMatrix(i, 0))                           'jv010616
    Call post_hold_log(zid, "Start Pallet")                     'jv010616
    s = "Update holdlist set spallet = '" & pno & "'"
    If Grid5.TextMatrix(i, 12) <> "WMS" Then
        s = s & ", userid = '" & WDUserId & "', holddate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
        Grid5.TextMatrix(i, 12) = WDUserId
        Grid5.TextMatrix(i, 13) = Format(Now, "yyMMdd hh:mm:ss")
    End If
    s = s & " Where id = " & Grid5.TextMatrix(i, 0)
    Wdb.Execute s
    'Clear Lane Hold Status to previous range
    If Grid2.Rows > 1 Then
        srrefresh = False                                           'jv120415
        For k = 1 To Grid2.Rows - 1                                 'jv050615
            Grid2.Row = k                                           'jv050615
            remsrhold_Click                                         'jv050615
        Next k                                                      'jv050615
        'refresh_cranes                                              'jv120416
        srrefresh = True                                            'jv120415
    End If
    Grid5.TextMatrix(i, 5) = pno
    Grid5_Click                         'Refresh Racks and Cranes
    'Add Lane Hold Status to new range
    If Grid2.Rows > 1 Then
        srrefresh = False                                           'jv120415
        For k = 1 To Grid2.Rows - 1                                 'jv050615
            Grid2.Row = k                                           'jv050615
            marksrhold_Click                                        'jv050615
        Next k                                                      'jv050615
        'refresh_cranes                                              'jv120414
        srrefresh = True                                            'jv120415
    End If
    Grid5_Click                         'Refresh Racks and Cranes
    Call post_hold_log(zid, "Start Pallet")                     'jv010616
End Sub

Private Sub Form_Load()
    'If Form1.plantno <> "50" Then hshist.Visible = False        'jv010716
    If Form1.userid <> "jvierus" Then hshist.Visible = False        'jv010716
    If Form1.plantno = "50" Then Grid2.Visible = True
    If Form1.plantno = "52" Then Grid3.Visible = True
    If Form1.plantno = "51" Then printlanes.Visible = False
    'refresh_holdlist
    refresh_sources                                             'jv121515
    srrefresh = True                                            'jv120415
End Sub

Private Sub Form_Resize()
    Grid5.Width = Me.Width - 120
    If Me.Height > 8000 Then Grid5.Height = Me.Height - (Grid1.Height + 1200)
    Grid1.Top = Grid5.Top + Grid5.Height
    Grid2.Top = Grid1.Top
    Grid3.Top = Grid1.Top
End Sub

Private Sub Grid1_Click()
    Grid1_RowColChange
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Grid1_RowColChange
    If Button = 2 Then PopupMenu edrackmenu
End Sub

Private Sub Grid1_RowColChange()
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then
        markrackhold.Enabled = False
        remrackhold.Enabled = False
        Exit Sub
    End If
    If Grid1.TextMatrix(Grid1.Row, 4) = "1" Then
        markrackhold.Enabled = False
        remrackhold.Enabled = True
    Else
        markrackhold.Enabled = True
        remrackhold.Enabled = False
    End If
End Sub

Private Sub Grid2_Click()
    Grid2_RowColChange
End Sub

Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Grid2_RowColChange
    If Button = 2 Then PopupMenu edsrmenu
End Sub

Private Sub Grid2_RowColChange()
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then
        marksrhold.Enabled = False
        remsrhold.Enabled = False
        Exit Sub
    End If
    If Grid2.TextMatrix(Grid2.Row, 6) = "H" Then
        marksrhold.Enabled = False
        remsrhold.Enabled = True
    Else
        marksrhold.Enabled = True
        remsrhold.Enabled = False
    End If
End Sub

Private Sub Grid5_Click()
    'If Grid5.TextMatrix(Grid5.Row, 7) = "TERM" Then
    '    delhold.Enabled = False
    'Else
        delhold.Enabled = True
    'End If
    Label1.Caption = Grid5.TextMatrix(Grid5.Row, 0)
    Refresh_racks
    refresh_cranes
End Sub

Private Sub Grid5_DblClick()
    edepallet_Click
End Sub

Private Sub Grid5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu holdmenu
    If Button = 1 Then Grid5_Click
End Sub

Private Sub hshist_Click()
    holdhist.hsku = Grid5.TextMatrix(Grid5.Row, 1)
    holdhist.hprod = Grid5.TextMatrix(Grid5.Row, 2)
    holdhist.Show
End Sub

Private Sub markrackhold_Click()
    Dim s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    s = "Update racks set hold = '1' where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Wdb.Execute s
    Refresh_racks
End Sub

Private Sub marksrhold_Click()
    Dim ds As ADODB.Recordset, s As String, d As daimessagerec, zid As Long
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    s = "Update lane set lane_status = 'H' where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    Wdb.Execute s
    If Grid2.TextMatrix(Grid2.Row, 1) >= "5" Then
        zid = wd_seq("BBC_HostToWrx")
        d.dhostmodifytime = Format(Now, "yyMMdd hh:mm:ss")
        d.imessagesequence = zid
        d.smessageidentifier = "InventoryHoldMessage"
        d.smessage = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
        d.smessage = d.smessage & "<!DOCTYPE InventoryHoldMessage SYSTEM " & Chr(34) & "wrxj.dtd" & Chr(34) & ">"
        d.smessage = d.smessage & "<InventoryHoldMessage>"
        s = "select sku, lot_num, barcode from position where laneno = " & Grid2.TextMatrix(Grid2.Row, 0)
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                d.bbcidentity = ds!barcode
                d.smessage = d.smessage & "<InventoryHold>"
                d.smessage = d.smessage & "<sItem>" & ds!sku & "</sItem>"
                'd.smessage = d.smessage & "<sLot>" & ds!lot_num & Mid(ds!barcode, 12, 1) & Mid(ds!barcode, 14, 3) & "</sLot>"
                d.smessage = d.smessage & "<sLot>" & ds!lot_num & Mid(ds!barcode, 11, 3) & Mid(ds!barcode, 14, 3) & "</sLot>"
                d.smessage = d.smessage & "<sHoldReason>PC</sHoldReason>"
                d.smessage = d.smessage & "</InventoryHold>"
                ds.MoveNext
            Loop
        End If
        ds.Close
        d.smessage = d.smessage & "</InventoryHoldMessage>"
        s = "Insert into BBC_HostToWrx(dhostmodifytime, imessagesequence, smessageidentifier, smessage"
        s = s & ", bbcidentity, bbcstatus) Values ('" & d.dhostmodifytime & "'"
        s = s & ", " & d.imessagesequence
        s = s & ", '" & d.smessageidentifier & "'"
        s = s & ", '" & d.smessage & "'"
        s = s & ", '" & d.bbcidentity & "'"
        s = s & ", 'PEND')"
        Wdb.Execute s
    End If
    If srrefresh = True Then refresh_cranes                                     'jv120415
End Sub

Private Sub ppaltots_Click()                                            'jv052517
    Dim i As Integer, prow As Integer, k As Integer, c As Integer
    Dim rt As String, rh As String, rf As String
    'pgrid.Visible = True
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 7
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) <> "ALL" Then
            pgrid.AddItem Combo1.List(i)
        End If
    Next i
    For i = 1 To Grid5.Rows - 1
        For k = 1 To pgrid.Rows - 1
            If pgrid.TextMatrix(k, 0) = Grid5.TextMatrix(i, 7) Then
                prow = k
                Exit For
            End If
        Next k
        Grid5.Row = i
        Grid5_Click
        DoEvents
        If Grid1.Rows > 1 Then
            For k = 1 To Grid1.Rows - 1
                pgrid.TextMatrix(prow, 4) = Val(pgrid.TextMatrix(prow, 4)) + Val(Grid1.TextMatrix(k, 3))
                pgrid.TextMatrix(prow, 6) = Val(pgrid.TextMatrix(prow, 6)) + Val(Grid1.TextMatrix(k, 3))
            Next k
        End If
        If Grid2.Visible And Grid2.Rows > 1 Then
            For k = 1 To Grid2.Rows - 1
                c = Val(Grid2.TextMatrix(k, 1))
                If c > 5 Then c = 5
                pgrid.TextMatrix(prow, c) = Val(pgrid.TextMatrix(prow, c)) + Val(Grid2.TextMatrix(k, 5))
                pgrid.TextMatrix(prow, 6) = Val(pgrid.TextMatrix(prow, 6)) + Val(Grid2.TextMatrix(k, 5))
            Next k
        End If
        If Grid3.Visible And Grid3.Rows > 1 Then
            For k = 1 To Grid3.Rows - 1
                pgrid.TextMatrix(prow, 5) = Val(pgrid.TextMatrix(prow, 5)) + Val(Grid3.TextMatrix(k, 4))
                pgrid.TextMatrix(prow, 6) = Val(pgrid.TextMatrix(prow, 6)) + Val(Grid3.TextMatrix(k, 4))
            Next k
        End If
    Next i
    For i = pgrid.Rows - 1 To 1 Step -1
        If Val(pgrid.TextMatrix(i, 6)) = 0 Then
            If pgrid.Rows > 2 Then
                pgrid.RemoveItem i
            Else
                pgrid.Rows = 1
            End If
        End If
    Next i
    pgrid.FormatString = "<Hold Reason|^SR1|^SR2|^SR3|^Racks|^SR5|^Total"
    pgrid.ColWidth(0) = 2400
    pgrid.ColWidth(1) = 1200
    pgrid.ColWidth(2) = 1200
    pgrid.ColWidth(3) = 1200
    pgrid.ColWidth(4) = 1200
    pgrid.ColWidth(5) = 1200
    pgrid.ColWidth(6) = 1200
    If Form1.plantno = "50" Then rt = "Brenham "
    If Form1.plantno = "51" Then rt = "Broken Arrow "
    If Form1.plantno = "52" Then rt = "Sylacauga "
    rt = rt & "On-Hold Pallet Totals"
    rh = "Warehouse Totals"
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub printlanes_Click()
    Dim rt As String, rh As String, rf As String, i As Integer
    If Grid5.Row > 0 Then i = Grid5.Row
    rt = "Hold Product Lanes - " & Grid5.TextMatrix(i, 7) & "<br>"              'jv121515
    rt = rt & Grid5.TextMatrix(i, 2) & " - " & Grid5.TextMatrix(i, 8)
    rh = "Pallet Barcodes: "
    rh = rh & Grid5.TextMatrix(i, 1) & " " & Grid5.TextMatrix(i, 8) & " " & Grid5.TextMatrix(i, 5) & " - "
    rh = rh & Grid5.TextMatrix(i, 1) & " " & Grid5.TextMatrix(i, 8) & " " & Grid5.TextMatrix(i, 6)
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    'If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
    '    Call printflexgrid(Printer, Grid2, rt, rh, rf)
    'Else
        If Grid2.Visible Then
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid2, rt, rh, rf, "linen", "lemonchiffon", "white")
        End If
        If Grid3.Visible Then
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid3, rt, rh, rf, "linen", "lemonchiffon", "white")
        End If
        
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    'End If
End Sub

Private Sub printlist_Click()
    Dim rt As String, rh As String, rf As String
    rt = "Hold Product Listing"
    rh = rt
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    Grid5.Redraw = False
    'If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
    '    Call printflexgrid(Printer, Grid5, rt, rh, rf)
    'Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid5, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    'End If
    Grid5.Redraw = True
End Sub

Private Sub printracks_Click()
    Dim rt As String, rh As String, rf As String, i As Integer
    If Grid5.Row > 0 Then i = Grid5.Row
    rt = "Hold Product Racks - " & Grid5.TextMatrix(i, 7) & "<br>"              'jv121515
    'rt = rt & Grid5.TextMatrix(i, 1) & " " & Grid5.TextMatrix(i, 2) & " - " & Grid5.TextMatrix(i, 8)
    rt = rt & Grid5.TextMatrix(i, 2) & " - " & Grid5.TextMatrix(i, 8)
    rh = "Pallet Barcodes: "
    rh = rh & Grid5.TextMatrix(i, 1) & " " & Grid5.TextMatrix(i, 8) & " " & Grid5.TextMatrix(i, 5) & " - "
    rh = rh & Grid5.TextMatrix(i, 1) & " " & Grid5.TextMatrix(i, 8) & " " & Grid5.TextMatrix(i, 6)
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    'If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
    '    Call printflexgrid(Printer, Grid1, rt, rh, rf)
    'Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    'End If
End Sub

Private Sub prtdetails_Click()
    Dim s As String, i As Integer, psku As String, ptot As Integer, k As Integer, plot As String
    Dim rt As String, rh As String, rf As String, psrc As String, pbc As String
    pgrid.Clear: pgrid.Cols = 6: pgrid.Rows = 1: psku = " ": ptot = 0: plot = " "
    For i = 1 To Grid5.Rows - 1
        'If Grid5.TextMatrix(i, 7) <> "TERM" Then
            psrc = Grid5.TextMatrix(i, 7)                                           'jv121515
            Grid5.Row = i
            Grid5_Click
            DoEvents
            If Grid5.TextMatrix(i, 1) <> psku Or Grid5.TextMatrix(i, 8) <> plot Then
                If psku > " " Then
                    s = psku & Chr(9) & "Total Pallets" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ptot
                    pgrid.AddItem s
                    pgrid.AddItem " "
                End If
                psku = Grid5.TextMatrix(i, 1)
                ptot = 0
                plot = Grid5.TextMatrix(i, 8)
                s = Grid5.TextMatrix(i, 1) & Chr(9) & Grid5.TextMatrix(i, 2) & Chr(9) & Grid5.TextMatrix(i, 8)
                pgrid.AddItem s
            End If
            If Grid1.Rows > 1 Then
                For k = 1 To Grid1.Rows - 1
                    pbc = Grid1.TextMatrix(k, 5)                    'jv121515
                    s = Chr(9) & pbc & Chr(9) & psrc & Chr(9) & "Racks" & Chr(9) & Grid1.TextMatrix(k, 1) & "-"   'jv121515
                    s = s & Grid1.TextMatrix(k, 2) & Chr(9) & Grid1.TextMatrix(k, 3)
                    pgrid.AddItem s
                    ptot = ptot + Val(Grid1.TextMatrix(k, 3))
                Next k
            End If
            If Grid2.Visible And Grid2.Rows > 1 Then
                For k = 1 To Grid2.Rows - 1
                    pbc = Grid2.TextMatrix(k, 7)                    'jv121515
                    s = Chr(9) & pbc & Chr(9) & psrc & Chr(9) & "SR-" & Grid2.TextMatrix(k, 1) & Chr(9)   'jv121515
                    s = s & Grid2.TextMatrix(k, 2) & "-" & Grid2.TextMatrix(k, 3) & "-" & Grid2.TextMatrix(k, 4)
                    s = s & Chr(9) & Grid2.TextMatrix(k, 5)
                    pgrid.AddItem s
                    ptot = ptot + Val(Grid2.TextMatrix(k, 5))
                Next k
            End If
            If Grid3.Visible And Grid3.Rows > 1 Then
                For k = 1 To Grid3.Rows - 1
                    pbc = Grid3.TextMatrix(k, 6)                    'jv121515
                    s = Chr(9) & pbc & Chr(9) & psrc & Chr(9) & Grid3.TextMatrix(k, 0) & Chr(9)   'jv121515
                    s = s & Grid3.TextMatrix(k, 1)
                    s = s & Chr(9) & Grid3.TextMatrix(k, 4)
                    pgrid.AddItem s
                    ptot = ptot + Val(Grid3.TextMatrix(k, 4))
                Next k
            End If
        'End If
    Next i
    s = psku & Chr(9) & "Total Pallets" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ptot
    pgrid.AddItem s
    
    pgrid.FillStyle = flexFillRepeat
    If pgrid.Rows > 1 Then
        For i = 1 To pgrid.Rows - 1
            If pgrid.TextMatrix(i, 0) > " " Then
                pgrid.Row = i: pgrid.RowSel = i
                If pgrid.TextMatrix(i, 2) > " " Then
                    pgrid.Col = 0: pgrid.ColSel = 2
                Else
                    pgrid.Col = 0: pgrid.ColSel = 5
                End If
                pgrid.CellBackColor = rcolor.BackColor
            End If
        Next i
    End If
    pgrid.FormatString = "^SKU|<Product|^Lot Code|^Whs|^Location|^Pallets"
    pgrid.ColWidth(0) = 1000
    pgrid.ColWidth(1) = 3000
    pgrid.ColWidth(2) = 1400
    pgrid.ColWidth(3) = 1000
    pgrid.ColWidth(4) = 1200
    pgrid.ColWidth(5) = 1200
    
    pgrid.Width = Me.Width - 100
    rt = "Hold Product Rack Details"
    rh = "Hold Product"
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "Yellow": gndc(0) = rcolor.BackColor
    Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub prtwhslist_Click()
    Dim s As String, i As Integer, psku As String, ptot As Integer, k As Integer, plot As String
    Dim rt As String, rh As String, rf As String, psrc As String
    Dim t1 As Integer, t2 As Integer, t3 As Integer, t5 As Integer
    Dim h1 As Integer, h2 As Integer, h3 As Integer, h5 As Integer
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid2.Redraw = False
    pgrid.Clear: pgrid.Cols = 6: pgrid.Rows = 1: psku = " ": ptot = 0: plot = " "
    For i = 1 To Grid5.Rows - 1
        If Grid5.TextMatrix(i, 7) <> "TERM" Then
            psrc = Grid5.TextMatrix(i, 7)                                           'jv121515
            'If psrc = "Schedule" Then psrc = "TEST HOLD"                            'jv121515
            Grid5.Row = i
            Grid5_Click
            DoEvents
            If Grid5.TextMatrix(i, 1) <> psku Or Grid5.TextMatrix(i, 8) <> plot Then
                If psku > " " Then
                    s = psku & Chr(9) & "Total Pallets" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ptot
                    pgrid.AddItem s
                    pgrid.AddItem " "
                End If
                psku = Grid5.TextMatrix(i, 1)
                ptot = 0
                plot = Grid5.TextMatrix(i, 8)
                s = Grid5.TextMatrix(i, 1) & Chr(9) & Grid5.TextMatrix(i, 2) & Chr(9) & Grid5.TextMatrix(i, 8)
                pgrid.AddItem s
            End If
            If Grid1.Rows > 1 Then
                For k = 1 To Grid1.Rows - 1
                    s = Chr(9) & Chr(9) & psrc & Chr(9) & "Racks" & Chr(9) & Grid1.TextMatrix(k, 1) & "-"   'jv121515
                    s = s & Grid1.TextMatrix(k, 2) & Chr(9) & Grid1.TextMatrix(k, 3)
                    pgrid.AddItem s
                    ptot = ptot + Val(Grid1.TextMatrix(k, 3))
                Next k
            End If
            If Grid2.Visible And Grid2.Rows > 1 Then
                t1 = 0: t2 = 0: t3 = 0: t5 = 0
                h1 = 0: h2 = 0: h3 = 0: h5 = 0
                For k = 1 To Grid2.Rows - 1
                    If Grid2.TextMatrix(k, 6) = "H" Then
                        If Grid2.TextMatrix(k, 1) = "1" Then h1 = h1 + Val(Grid2.TextMatrix(k, 5))
                        If Grid2.TextMatrix(k, 1) = "2" Then h2 = h2 + Val(Grid2.TextMatrix(k, 5))
                        If Grid2.TextMatrix(k, 1) = "3" Then h3 = h3 + Val(Grid2.TextMatrix(k, 5))
                        If Grid2.TextMatrix(k, 1) >= "5" Then h5 = h5 + Val(Grid2.TextMatrix(k, 5))
                    Else
                        If Grid2.TextMatrix(k, 1) = "1" Then t1 = t1 + Val(Grid2.TextMatrix(k, 5))
                        If Grid2.TextMatrix(k, 1) = "2" Then t2 = t2 + Val(Grid2.TextMatrix(k, 5))
                        If Grid2.TextMatrix(k, 1) = "3" Then t3 = t3 + Val(Grid2.TextMatrix(k, 5))
                        If Grid2.TextMatrix(k, 1) >= "5" Then t5 = t5 + Val(Grid2.TextMatrix(k, 5))
                    End If
                    ptot = ptot + Val(Grid2.TextMatrix(k, 5))
                Next k
                If t1 > 0 Then
                    s = Chr(9) & Chr(9) & psrc & Chr(9) & "SR-1" & Chr(9)       'jv121515
                    s = s & "Available Lanes"
                    s = s & Chr(9) & t1
                    pgrid.AddItem s
                End If
                If h1 > 0 Then
                    s = Chr(9) & Chr(9) & psrc & Chr(9) & "SR-1" & Chr(9)       'jv121515
                    s = s & "Marked On-Hold"
                    s = s & Chr(9) & h1
                    pgrid.AddItem s
                End If
                If t2 > 0 Then
                    s = Chr(9) & Chr(9) & psrc & Chr(9) & "SR-2" & Chr(9)       'jv121515
                    s = s & "Available Lanes"
                    s = s & Chr(9) & t2
                    pgrid.AddItem s
                End If
                If h2 > 0 Then
                    s = Chr(9) & Chr(9) & psrc & Chr(9) & "SR-2" & Chr(9)       'jv121515
                    s = s & "Marked On-Hold"
                    s = s & Chr(9) & h2
                    pgrid.AddItem s
                End If
                If t3 > 0 Then
                    s = Chr(9) & Chr(9) & psrc & Chr(9) & "SR-3" & Chr(9)       'jv121515
                    s = s & "Available Lanes"
                    s = s & Chr(9) & t3
                    pgrid.AddItem s
                End If
                If h3 > 0 Then
                    s = Chr(9) & Chr(9) & psrc & Chr(9) & "SR-3" & Chr(9)       'jv121515
                    s = s & "Marked On-Hold"
                    s = s & Chr(9) & h3
                    pgrid.AddItem s
                End If
                If t5 > 0 Then
                    s = Chr(9) & Chr(9) & psrc & Chr(9) & "SR-5" & Chr(9)       'jv121515
                    s = s & "Available Lanes"
                    s = s & Chr(9) & t5
                    pgrid.AddItem s
                End If
                If h5 > 0 Then
                    s = Chr(9) & Chr(9) & psrc & Chr(9) & "SR-5" & Chr(9)       'jv121515
                    s = s & "Marked On-Hold"
                    s = s & Chr(9) & h5
                    pgrid.AddItem s
                End If
            End If
            If Grid3.Visible And Grid3.Rows > 1 Then                            'jv121515
                t5 = 0
                For k = 1 To Grid3.Rows - 1
                    t5 = t5 + Val(Grid3.TextMatrix(k, 4))
                    ptot = ptot + Val(Grid3.TextMatrix(k, 4))
                Next k
                If t5 > 0 Then
                    s = Chr(9) & Chr(9) & psrc & Chr(9) & "CS5" & Chr(9)       'jv121515
                    s = s & "Lanes"
                    s = s & Chr(9) & t5
                    pgrid.AddItem s
                End If
            End If
        End If
    Next i
    s = psku & Chr(9) & "Total Pallets" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ptot
    pgrid.AddItem s
    
    pgrid.FillStyle = flexFillRepeat
    If pgrid.Rows > 1 Then
        For i = 1 To pgrid.Rows - 1
            If pgrid.TextMatrix(i, 0) > " " Then
                pgrid.Row = i: pgrid.RowSel = i
                If pgrid.TextMatrix(i, 2) > " " Then
                    pgrid.Col = 0: pgrid.ColSel = 2
                Else
                    pgrid.Col = 0: pgrid.ColSel = 5
                End If
                pgrid.CellBackColor = rcolor.BackColor
            End If
        Next i
    End If
    pgrid.FormatString = "^SKU|<Product|^Lot Code|^Whs|^Location|^Pallets"
    pgrid.ColWidth(0) = 1000
    pgrid.ColWidth(1) = 3000
    pgrid.ColWidth(2) = 1400
    pgrid.ColWidth(3) = 1000
    pgrid.ColWidth(4) = 1200
    pgrid.ColWidth(5) = 1200
    
    pgrid.Width = Me.Width - 100
    rt = "Hold Product Warehouses"
    rh = "Hold Product"
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    Grid1.Redraw = True: Grid2.Redraw = True
    Screen.MousePointer = 0
    htdc(0) = "Yellow": gndc(0) = rcolor.BackColor
    Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub psched_Click()
    Form21.Show
End Sub

Private Sub remrackhold_Click()
    Dim s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    s = "Update racks set hold = '0' where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Wdb.Execute s
    Refresh_racks
End Sub

Private Sub remsrhold_Click()
    Dim ds As ADODB.Recordset, s As String, d As daimessagerec, zid As Long
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    s = "Update lane set lane_status = ' ' where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    Wdb.Execute s
    If Grid2.TextMatrix(Grid2.Row, 1) >= "5" Then
        zid = wd_seq("BBC_HostToWrx")
        d.dhostmodifytime = Format(Now, "yyMMdd hh:mm:ss")
        d.imessagesequence = zid
        d.smessageidentifier = "InventoryHoldMessage"
        d.smessage = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
        d.smessage = d.smessage & "<!DOCTYPE InventoryHoldMessage SYSTEM " & Chr(34) & "wrxj.dtd" & Chr(34) & ">"
        d.smessage = d.smessage & "<InventoryHoldMessage>"
        s = "select sku, lot_num, barcode from position where laneno = " & Grid2.TextMatrix(Grid2.Row, 0)
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                d.bbcidentity = ds!barcode
                d.smessage = d.smessage & "<InventoryHold>"
                d.smessage = d.smessage & "<sItem>" & ds!sku & "</sItem>"
                'd.smessage = d.smessage & "<sLot>" & ds!lot_num & Mid(ds!barcode, 12, 1) & Mid(ds!barcode, 14, 3) & "</sLot>"
                d.smessage = d.smessage & "<sLot>" & ds!lot_num & Mid(ds!barcode, 11, 3) & Mid(ds!barcode, 14, 3) & "</sLot>"
                d.smessage = d.smessage & "<sHoldReason/>"
                d.smessage = d.smessage & "</InventoryHold>"
                ds.MoveNext
            Loop
        End If
        ds.Close
        d.smessage = d.smessage & "</InventoryHoldMessage>"
        s = "Insert into BBC_HostToWrx(dhostmodifytime, imessagesequence, smessageidentifier, smessage"
        s = s & ", bbcidentity, bbcstatus) Values ('" & d.dhostmodifytime & "'"
        s = s & ", " & d.imessagesequence
        s = s & ", '" & d.smessageidentifier & "'"
        s = s & ", '" & d.smessage & "'"
        s = s & ", '" & d.bbcidentity & "'"
        s = s & ", 'PEND')"
        Wdb.Execute s
    End If
    If srrefresh = True Then refresh_cranes                             'jv120415
End Sub
