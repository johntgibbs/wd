VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form13 
   Caption         =   "Sylacuaga CS5 and Rack Totals"
   ClientHeight    =   12450
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12375
   LinkTopic       =   "Form13"
   ScaleHeight     =   12450
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   6735
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   11880
      _Version        =   327680
      BackColorFixed  =   16777152
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   10800
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2990
      _Version        =   327680
      BackColor       =   12648447
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   8880
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   2990
      _Version        =   327680
      BackColor       =   16777152
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   6960
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2990
      _Version        =   327680
      BackColor       =   12648447
      FocusRect       =   0
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rack Pallets"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   10560
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CS5 Lane Locks"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   8640
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CS5 Pallet Locations"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Menu refmenu 
      Caption         =   "Refresh"
      Begin VB.Menu refdata 
         Caption         =   "Refresh Data"
      End
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
      Begin VB.Menu prtlpt 
         Caption         =   "HTML"
      End
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim db5 As ADODB.Connection, ds5 As ADODB.Recordset

Private Sub refresh_pgrid()
    Dim i As Integer, k As Integer, psku As String, newrow As Boolean, s As String, punits As Long
    Dim rphold As Long, ruhold As Long, rskudesc As String
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 8
    For i = 1 To Grid1.Rows - 1
        newrow = True
        psku = Trim(Grid1.TextMatrix(i, 10))                    'jv082415
        If Len(psku) = 8 Then                                   'jv082415       1320-228
            psku = Left(psku, 4)                                'jv082415
        Else                                                    'jv082415
            If Len(psku) = 7 Then                               'jv082415       777-228
                psku = Left(psku, 3)                            'jv082415
            End If                                              'jv082415
        End If                                                  'jv082415
        'psku = Left(Grid1.TextMatrix(i, 10), 3)
        punits = Val(Grid1.TextMatrix(i, 12))
        If Val(psku) > 0 Then
            For k = 0 To pgrid.Rows - 1
                If pgrid.TextMatrix(k, 0) = psku Then
                    pgrid.TextMatrix(k, 2) = Val(pgrid.TextMatrix(k, 2)) + 1
                    pgrid.TextMatrix(k, 5) = Val(pgrid.TextMatrix(k, 5)) + punits
                    newrow = False
                    Exit For
                End If
            Next k
            If newrow = True Then
                s = psku & Chr(9) & Grid1.TextMatrix(i, 11) & Chr(9) & "1" & Chr(9) & Chr(9) & Chr(9) & punits
                pgrid.AddItem s
                'MsgBox "check " & psku
            End If
        End If
    Next i
    For i = 1 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 10) <> "0" Then
            psku = Trim(Grid2.TextMatrix(i, 7))                     'jv082415
            If Len(psku) = 8 Then                                   'jv082415       1320-228
                psku = Left(psku, 4)                                'jv082415
            Else                                                    'jv082415
                If Len(psku) = 7 Then                               'jv082415       777-228
                    psku = Left(psku, 3)                            'jv082415
                End If                                              'jv082415
            End If                                                  'jv082415
            'psku = Left(Grid2.TextMatrix(i, 7), 3)
            punits = Val(Grid2.TextMatrix(i, 5)) * Val(Grid2.TextMatrix(i, 9))
            For k = 0 To pgrid.Rows - 1
                If pgrid.TextMatrix(k, 0) = psku Then
                    pgrid.TextMatrix(k, 3) = Val(pgrid.TextMatrix(k, 3)) + Val(Grid2.TextMatrix(i, 9))
                    pgrid.TextMatrix(k, 6) = Val(pgrid.TextMatrix(k, 6)) + punits
                    newrow = False
                    Exit For
                End If
            Next k
        End If
    Next i
    For i = 1 To Grid3.Rows - 1
        newrow = True
        psku = Grid3.TextMatrix(i, 0)
        punits = (Val(Grid3.TextMatrix(i, 1)) + Val(Grid3.TextMatrix(i, 2))) '* Val(Grid3.TextMatrix(i, 4))
        For k = 0 To pgrid.Rows - 1
            If pgrid.TextMatrix(k, 0) = psku Then
                pgrid.TextMatrix(k, 2) = Val(pgrid.TextMatrix(k, 2)) + 1 'Val(Grid3.TextMatrix(i, 4))
                pgrid.TextMatrix(k, 5) = Val(pgrid.TextMatrix(k, 5)) + punits
                If Grid3.TextMatrix(i, 3) = "Y" Or Grid3.TextMatrix(i, 4) = "1" Then
                    pgrid.TextMatrix(k, 3) = Val(pgrid.TextMatrix(k, 3)) + 1 'Val(Grid3.TextMatrix(i, 4))
                    pgrid.TextMatrix(k, 6) = Val(pgrid.TextMatrix(k, 6)) + punits
                End If
                newrow = False
                Exit For
            End If
        Next k
        If newrow = True Then
            If skurec(Val(psku)).sku = psku Then
                rskudesc = skurec(Val(psku)).prodname
            Else
                rskudesc = "Invalid SKU"
            End If
                
            If Grid3.TextMatrix(i, 3) = "Y" Or Grid3.TextMatrix(i, 4) = "1" Then
                rphold = 1 'Val(Grid3.TextMatrix(i, 4))
                ruhold = punits
            Else
                rphold = 0
                ruhold = 0
            End If
            s = psku & Chr(9) & rskudesc & Chr(9)
            s = s & Grid3.TextMatrix(i, 4) & Chr(9)
            s = s & Format(rphold, "#") & Chr(9) & Chr(9)
            s = s & punits & Chr(9) & Format(ruhold, "#")
            pgrid.AddItem s
            'MsgBox "check " & psku & " " & s & " " & rphold & " " & ruhold
        End If
    Next i
    If pgrid.Rows > 1 Then
        For i = 1 To pgrid.Rows - 1
            pgrid.TextMatrix(i, 4) = Val(pgrid.TextMatrix(i, 2)) - Val(pgrid.TextMatrix(i, 3))
            pgrid.TextMatrix(i, 7) = Val(pgrid.TextMatrix(i, 5)) - Val(pgrid.TextMatrix(i, 6))
        Next i
        pgrid.RowSel = pgrid.Row
        pgrid.Col = 0: pgrid.ColSel = 1
        pgrid.Sort = 5
    End If
    pgrid.FormatString = "^SKU|<Product|^Pallet Qty|^On HOLD|^Net|^Unit Qty|^On HOLD|^Net"
    pgrid.ColWidth(0) = 1000
    pgrid.ColWidth(1) = 3500
    pgrid.ColWidth(2) = 1000
    pgrid.ColWidth(3) = 1000
    pgrid.ColWidth(4) = 1000
    pgrid.ColWidth(5) = 1000
    pgrid.ColWidth(6) = 1000
    pgrid.ColWidth(7) = 1000
End Sub

Private Sub refresh_grid1_ado()
    'Dim db5 As ADODB.Connection, ds5 As ADODB.Recordset
    Dim s As String, t1 As Long, t2 As Long, i As Integer
    Dim cfile As String, psku As String, p1 As Currency, p2 As Currency
    Me.Caption = "Sylacauga CS5"
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 22
    'Set Wdb = CreateObject("ADODB.Connection")
    'Wdb.Open "ODBC;DATABASE=SYRacks;UID=bbcwd502;PWD=alabama502;DSN=wdsql502"
    Set db5 = CreateObject("ADODB.Connection")
    db5.Open "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=sywms;PWD=!Sylacauga_WMS1907"
    's = "Select * from vContainerLocation_1033"
    s = "Select * from vAllInventory_1033"                          'westfalia update
    Set ds5 = db5.Execute(s)
    If ds5.BOF = False Then
        ds5.MoveFirst
        Do Until ds5.EOF
            s = ds5(0)
            'For i = 1 To 21
            '    s = s & Chr(9) & Trim(ds5(i))
            'Next i
            
            s = s & Chr(9)                                          'jv051018   westfalia update
            s = s & Mid(ds5(7), 2, 1)                               'jv051018
            s = s & Mid(ds5(7), 4, 2)                               'jv051018
            s = s & Mid(ds5(7), 7, 2)                               'jv051018
            s = s & Chr(9) & " "                                    'jv051018
            s = s & Chr(9) & " "                                    'jv051018
            s = s & Chr(9) & " "                                    'jv051018
            s = s & Chr(9) & " "                                    'jv051018
            s = s & Chr(9) & " "                                    'jv051018
            s = s & Chr(9) & Trim(ds5(4))                           'jv051018
            s = s & Chr(9) & Trim(ds5(7)) & "." & ds5(8)            'jv051018
            s = s & Chr(9) & Trim(ds5(8))                           'jv051018
            s = s & Chr(9) & Trim(ds5(9))                           'jv051018
            s = s & Chr(9) & Trim(ds5(10))                          'jv051018
            s = s & Chr(9) & Trim(ds5(11))                          'jv051018
            s = s & Chr(9) & Trim(ds5(22))                          'jv051018
            s = s & Chr(9) & Trim(ds5(15))                          'jv051018
            s = s & Chr(9) & " "                                    'jv051018
            s = s & Chr(9) & Trim(ds5(16))                          'jv051018
            If ds5(16) > "0" Then
                s = s & Chr(9) & Trim(ds5(17)) & "Locked"                          'jv051018
            Else
                s = s & Chr(9) & Trim(ds5(17))                          'jv051018
            End If
            s = s & Chr(9) & Trim(ds5(5))                           'jv051018
            s = s & Chr(9) & " "                                    'jv051018
            s = s & Chr(9) & " "                                    'jv051018
            s = s & Chr(9) & " "                                    'jv051018
            
            
            Grid1.AddItem s
            ds5.MoveNext
        Loop
    End If
    ds5.Close ': db5.Close
    Grid1.TextMatrix(1, 0) = Right(Grid1.TextMatrix(1, 0), 5)
    s = "^PalKey|^LocKey|^|^|^|^|^|<PalType|<Location|^Position|<SKU|<ProdDesc|^Units|^RecTime|^LotDate|^|^QAHold|<Reason|^Plate|^Disposed|^PalNum|^"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 600
    Grid1.ColWidth(11) = 2000
    Grid1.ColWidth(12) = 1000
    Grid1.ColWidth(13) = 1000
    Grid1.ColWidth(14) = 1000
    Grid1.ColWidth(15) = 1000
    Grid1.ColWidth(16) = 1000
    Grid1.ColWidth(17) = 1000
    Grid1.ColWidth(18) = 1000
    Grid1.ColWidth(19) = 1000
    Grid1.ColWidth(20) = 1000
    Grid1.ColWidth(21) = 1000
    
    Grid1.Row = 1: Grid1.RowSel = 1
    Grid1.Col = 8: Grid1.ColSel = 8
    Grid1.Sort = 5
    
    refresh_grid2_ado
    db5.Close
End Sub

Private Sub refresh_grid2_ado()
    Dim s As String ', ds5 As ADODB.Recordset
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 12
    s = "SELECT tLocationData.sLocationID, "
    s = s & "tLaneData.iLevel, tLaneData.iRow, tLaneData.iBlock, "
    s = s & "tContainerLocationData.iLocationID, "
    s = s & "tInventoryData.nQuantity, "
    s = s & "tLotData.dtProduction, "
    s = s & "tItemMaster.sItemID, tItemMaster.sItemDescription, tLaneLock.iLocked,"
    s = s & "tLaneLock.sDescription, count(*) "
    s = s & "FROM tLocationData, tLaneData, tContainerLocationData, tInventoryData, "
    s = s & "tLotData, tItemMaster, tLaneLock"
    s = s & " WHERE tLaneData.iLocationID = tLocationData.iLocationID"
    s = s & " AND tContainerLocationData.iLocationID = tLaneData.iLocationID"
    s = s & " AND tLaneLock.iLaneSysID = tLaneData.iLocationID"
    s = s & " AND tInventoryData.iContainerDataSysID = tContainerLocationData.iContainerDataSysID"
    s = s & " AND tLotData.iLotDataSysID = tInventoryData.iLotDataSysID"
    s = s & " AND tItemMaster.iItemMasterSysID = tLotData.iItemMasterSysID"
    s = s & " GROUP BY tLocationData.sLocationID, "
    s = s & "tLaneData.iLevel, tLaneData.iRow, tLaneData.iBlock, "
    s = s & "tContainerLocationData.iLocationID, "
    s = s & "tInventoryData.nQuantity, "
    s = s & "tLotData.dtProduction, "
    s = s & "tItemMaster.sItemID, tItemMaster.sItemDescription, tLaneLock.iLocked, tLaneLock.sDescription"
    s = s & " ORDER BY tLocationData.sLocationID " ', tContainerLocationData.iPosition"
    Set ds5 = db5.Execute(s)
    If ds5.BOF = False Then
        ds5.MoveFirst
        Do Until ds5.EOF
            s = Trim(ds5(0)) & Chr(9)
            s = s & Trim(ds5(1)) & Chr(9)
            s = s & Trim(ds5(2)) & Chr(9)
            s = s & Trim(ds5(3)) & Chr(9)
            s = s & Trim(ds5(4)) & Chr(9)
            s = s & Trim(ds5(5)) & Chr(9)
            s = s & Trim(ds5(6)) & Chr(9)
            s = s & Trim(ds5(7)) & Chr(9)
            s = s & Trim(ds5(8)) & Chr(9)
            s = s & Trim(ds5(11)) & Chr(9)
            s = s & Trim(ds5(9)) & Chr(9)
            s = s & Trim(ds5(10)) & Chr(9)
            Grid2.AddItem s
            ds5.MoveNext
        Loop
    End If
    ds5.Close
    s = "^Location|^Level|^Row|^Block|^ContLoc|^Qty|^LotDate|^SKU|<Description|^Pallets|^Lock|<Reason"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 800
    Grid2.ColWidth(2) = 800
    Grid2.ColWidth(3) = 800
    Grid2.ColWidth(4) = 900
    Grid2.ColWidth(5) = 800
    Grid2.ColWidth(6) = 1400
    Grid2.ColWidth(7) = 800
    Grid2.ColWidth(8) = 3000
    Grid2.ColWidth(9) = 1000
    Grid2.ColWidth(10) = 600
    Grid2.ColWidth(11) = 2000
End Sub

Private Sub refresh_grid3_ado()
    Dim rdb As ADODB.Connection, rds As ADODB.Recordset, s As String
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 8
    Set rdb = CreateObject("ADODB.Connection")
    rdb.Open "Driver={SQL Server};Server=bbsy-01-wdsql;DATABASE=SYRacks;UID=bbcwd502;PWD=alabama502;DSN=wdsql502"
    's = "select p.sku, p.count_qty, p.qty2, r.hold, count(*) from rackpos p, racks r"
    's = s & " where p.sku > '000' and p.sku < '999'"
    's = s & " and r.id = p.rackno"
    's = s & " group by p.sku, p.count_qty, p.qty2, r.hold"
    'Set rds = rdb.Execute(s)
    'If rds.BOF = False Then
    '    rds.MoveFirst
    '    Do Until rds.EOF
    '        s = rds(0) & Chr(9)
    '        s = s & rds(1) & Chr(9)
    '        s = s & rds(2) & Chr(9)
    '        s = s & rds(3) & Chr(9)
    '        s = s & rds(4)
    '        Grid3.AddItem s
    '        rds.MoveNext
    '    Loop
    'End If
    s = "select p.sku, p.count_qty, p.qty2, p.hold, r.hold, r.aisle, r.rack, p.posn_num from rackpos p, racks r"
    s = s & " where p.sku > '000' and p.sku < '999'"
    s = s & " and r.id = p.rackno"
    s = s & " order by p.sku, r.aisle, r.rack, p.posn_num"
    Set rds = rdb.Execute(s)
    If rds.BOF = False Then
        rds.MoveFirst
        Do Until rds.EOF
            s = rds(0) & Chr(9)
            s = s & rds(1) & Chr(9)
            s = s & rds(2) & Chr(9)
            s = s & rds(3) & Chr(9)
            s = s & rds(4) & Chr(9)
            s = s & rds(5) & Chr(9)
            s = s & rds(6) & Chr(9)
            s = s & rds(7)
            Grid3.AddItem s
            rds.MoveNext
        Loop
    End If
    
    rds.Close: rdb.Close
    Grid3.FormatString = "^SKU|^PalUnits|^PalUnits2|^p.Hold|^r.Hold|^Aisle|^Rack|^Pos"
    Grid3.ColWidth(0) = 1000
    Grid3.ColWidth(1) = 1000
    Grid3.ColWidth(2) = 1000
    Grid3.ColWidth(3) = 1000
    Grid3.ColWidth(4) = 1000
    Grid3.ColWidth(5) = 1000
    Grid3.ColWidth(6) = 1000
    Grid3.ColWidth(7) = 1000
End Sub

Private Sub refresh_grid2()
    Dim s As String, cfile As String, psku As String
    cfile = "\\bbsy-02-dc\f\shared\general\All Pallets.xls"
    If Len(Dir(cfile)) = 0 Then Exit Sub
    Screen.MousePointer = 11
    s = Format(FileDateTime(cfile), "m-d-yyyy h:mm am/pm")
    Me.Caption = "Sylacauga CS5  Last updated: " & s
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 22
    Open cfile For Input As #1
    Line Input #1, s
    psku = Left(s, 7)
    Grid2.AddItem s
    Grid2.TextMatrix(1, 0) = Trim(psku)
    Do Until EOF(1)
        Line Input #1, s
        psku = Left(s, 5)
        Grid2.AddItem s
        Grid2.TextMatrix(Grid2.Rows - 1, 0) = Trim(psku)
    Loop
    Close #1
    Grid2.TextMatrix(1, 0) = Right(Grid2.TextMatrix(1, 0), 5)
    s = "^PalKey|^LocKey|^|^|^|^|^|<PalType|<Location|^Position|<SKU|<ProdDesc|^Units|^Rectime|^LotDate|^|^Locked|<Reason|^|^Hold|^PalNum|^"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 800
    Grid2.ColWidth(1) = 800
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 1000
    Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 1000
    Grid2.ColWidth(7) = 1000
    Grid2.ColWidth(8) = 1000
    Grid2.ColWidth(9) = 1000
    Grid2.ColWidth(10) = 600
    Grid2.ColWidth(11) = 2000
    Grid2.ColWidth(12) = 1000
    Grid2.ColWidth(13) = 1000
    Grid2.ColWidth(14) = 1000
    Grid2.ColWidth(15) = 1000
    Grid2.ColWidth(16) = 1000
    Grid2.ColWidth(17) = 1000
    Grid2.ColWidth(18) = 1000
    Grid2.ColWidth(19) = 1000
    Grid2.ColWidth(20) = 1000
    Grid2.ColWidth(21) = 1000
    Screen.MousePointer = 0
End Sub

Private Sub refresh_grid()
    Dim s As String, tp As Long, tu As Long, i As Integer
    Dim cfile As String, psku As String, th As Long, tn As Long
    refresh_grid2
    cfile = "\\bbsy-02-dc\f\shared\general\inv compact.xls"
    If Len(Dir(cfile)) = 0 Then Exit Sub
    s = Format(FileDateTime(cfile), "m-d-yyyy h:mm am/pm")
    Me.Caption = "Sylacauga CS5  Last updated: " & s
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    
    Open cfile For Input As #1
    Line Input #1, s
    psku = Left(s, 5)
    Grid1.AddItem s
    Grid1.TextMatrix(1, 0) = Trim(psku)
    Do Until EOF(1)
        Line Input #1, s
        'MsgBox s
        psku = Left(s, 3)
        Grid1.AddItem s
        Grid1.TextMatrix(Grid1.Rows - 1, 0) = Trim(psku)
    Loop
    Close #1
    
    For i = 1 To Grid2.Rows - 1
        'If Grid2.TextMatrix(i, 16) = "True" Or Grid2.TextMatrix(i, 19) = "True" Then
        If Grid2.TextMatrix(i, 16) = "True" Then
            For k = 1 To Grid1.Rows - 1
                If Val(Grid1.TextMatrix(k, 0)) = Val(Grid2.TextMatrix(i, 10)) Then
                    Grid1.TextMatrix(k, 4) = Val(Grid1.TextMatrix(k, 4)) + 1
                    'Grid1.TextMatrix(k, 5) = Val(Grid1.TextMatrix(k, 2)) - Val(Grid1.TextMatrix(k, 4))
                    Exit For
                End If
            Next k
        End If
    Next i
    
    tp = 0: tu = 0: th = 0: tn = 0
    For i = 0 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 3) = Format(Val(Grid1.TextMatrix(i, 3)), "0")
        Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 2)) - Val(Grid1.TextMatrix(i, 4))
        tp = tp + Val(Grid1.TextMatrix(i, 2))
        tu = tu + Val(Grid1.TextMatrix(i, 3))
        th = th + Val(Grid1.TextMatrix(i, 4))
        tn = tn + Val(Grid1.TextMatrix(i, 5))
    Next i
    s = Chr(9) & "Totals" & Chr(9) & tp & Chr(9) & tu & Chr(9) & th & Chr(9) & tn
    Grid1.AddItem s
    Grid1.TextMatrix(1, 0) = Right(Grid1.TextMatrix(1, 0), 3)
    Grid1.FormatString = "^SKU|<Product|^Pallets|^Units|^OnHold|^Net"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    refresh_grid1_ado
    'refresh_grid2_ado
    refresh_grid3_ado
    refresh_pgrid
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    'If Me.Height > 2000 Then Grid1.Height = Me.Height - 780
    Grid2.Width = Me.Width - 100
    Grid3.Width = Me.Width - 100
    pgrid.Width = Me.Width - 100
    Label1.Width = Me.Width - 100
    Label2.Width = Me.Width - 100
    Label3.Width = Me.Width - 100
End Sub

Private Sub prtlpt_Click()
    Dim rt As String, rh As String, rf As String
    rt = "Sylacauga CS5 and Rack Totals"
    rh = Format(Now, "mmmm d, yyyy h:mm AM/PM")
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, pgrid, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, "s:\wd\html\sycs5.htm", pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe s:\wd\html\sycs5.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe s:\\wd\html\sycs5.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
    
End Sub

Private Sub refdata_Click()
    Screen.MousePointer = 11
    refresh_grid1_ado
    'refresh_grid2_ado
    refresh_grid3_ado
    refresh_pgrid
    Screen.MousePointer = 0

    'refresh_grid
End Sub
