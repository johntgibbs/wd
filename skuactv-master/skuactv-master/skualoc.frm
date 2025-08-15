VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Pallet Locations"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form4"
   ScaleHeight     =   3600
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4683
      _Version        =   327680
      Cols            =   6
      BackColorFixed  =   12648447
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.Label blot 
      Caption         =   "blot"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label bwhs 
      Caption         =   "bwhs"
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
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Label bprod 
      Caption         =   "bprod"
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
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label bsku 
      Caption         =   "bsku"
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
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function calc_date(lotcode As String) As String
    Dim seed As String
    'seed = "12-31-19" & Val(Left(lotcode, 2)) - 1
    If Val(Left(lotcode, 2)) > 90 Then
        seed = "19" & Left(lotcode, 2)
    Else
        seed = "20" & Left(lotcode, 2)
    End If
    seed = "12-31-" & Val(seed) - 1
    calc_date = Format(DateAdd("d", Val(Right(lotcode, 3)), seed), "m-d-yyyy")
End Function

Private Function wd_lotnum(pdate As String) As String
    Dim sdate As String, s As String
    'pdate = Format(pdate, "m-d-yyyy")
    pdate = Format(DateAdd("yyyy", -2, pdate), "m-d-yyyy")
    sdate = "1-1-" & Right(pdate, 4)
    s = Format(DateDiff("d", sdate, pdate) + 1, "000")
    s = Right(pdate, 2) & s
    wd_lotnum = s
End Function

Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, sqlx As String, r As String
    Dim ps As ADODB.Recordset, ds5 As ADODB.Recordset, s As String, psz As String
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1
    If blot = "blot" Then Exit Sub
    If bsku = "bsku" Then Exit Sub
    If bwhs = "bwhs" Then Exit Sub
    If Val(bwhs) <> 4 And Form1.plantno = "50" Then
        psz = "??"
        sqlx = "select uom_per_pallet from sku_config where sku = '" & bsku & "'"
        Set ps = Form1.wdb.Execute(sqlx)
        If ps.BOF = False Then
            ps.MoveFirst
            psz = ps(0)
        End If
        ps.Close
        sqlx = "select laneno,lot_num,count(*) from position"
        sqlx = sqlx & " where sku = '" & bsku & "'"
        If Val(bwhs) > 0 And Val(bwhs) < 4 Then
            sqlx = sqlx & " and whse_num = " & bwhs
        End If
        If Val(blot) > 0 Then
            sqlx = sqlx & " and (lot_num = '" & blot & "'"
            sqlx = sqlx & " or lot2 = '" & blot & "')"
        End If
        sqlx = sqlx & " group by laneno,lot_num"
        Set ps = Form1.wdb.Execute(sqlx)
        If ps.BOF = False Then
            ps.MoveFirst
            Do Until ps.EOF
                sqlx = "select whse_num,vert_loc,horz_loc,rack_side,gmasize from lane"
                sqlx = sqlx & " where id = " & ps(0)
                Set ds = Form1.wdb.Execute(sqlx)
                If ds.BOF = False Then
                    ds.MoveFirst
                    If ds!gmasize > 0 Then                  'jv082813
                        sqlx = "6" & Chr(9)
                    Else
                        sqlx = ds!whse_num & Chr(9)
                    End If
                    sqlx = sqlx & ds!vert_loc & " "
                    sqlx = sqlx & ds!horz_loc & " "
                    sqlx = sqlx & ds!rack_side & Chr(9)
                    sqlx = sqlx & ps!lot_num & Chr(9)
                    sqlx = sqlx & calc_date(ps!lot_num) & Chr(9)
                    If ds!gmasize > 0 Then                          'jv082813
                        sqlx = sqlx & ps(2) & Chr(9) & ds!gmasize
                    Else
                        sqlx = sqlx & ps(2) & Chr(9) & psz
                    End If
                    Grid1.AddItem sqlx
                End If
                ds.Close
                ps.MoveNext
            Loop
        End If
        ps.Close
    End If
    If Val(bwhs) <> 4 And Form1.plantno = "52" Then
        s = "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
        'Set db5 = OpenDatabase(mysqldev, dbcdrivernoprompt, True, s)
        
        s = "SELECT tLocationData.sLocationID, "
        s = s & "tLaneData.iLevel, tLaneData.iRow, tLaneData.iBlock, "
        s = s & "tContainerLocationData.iLocationID, "
        s = s & "tInventoryData.nQuantity, "
        's = s & "tLotData.dtProduction, "
        s = s & "tLotData.dtExpiration, "                       'jv081115
        s = s & "tItemMaster.sItemID, tItemMaster.sItemDescription, count(*) "
        s = s & "FROM tLocationData, tLaneData, tContainerLocationData, tInventoryData, "
        s = s & "tLotData, tItemMaster"
        s = s & " WHERE tLaneData.iLocationID = tLocationData.iLocationID"
        s = s & " AND tContainerLocationData.iLocationID = tLaneData.iLocationID"
        s = s & " AND tInventoryData.iContainerDataSysID = tContainerLocationData.iContainerDataSysID"
        s = s & " AND tLotData.iLotDataSysID = tInventoryData.iLotDataSysID"
        s = s & " AND tItemMaster.iItemMasterSysID = tLotData.iItemMasterSysID"
        
        s = s & " and tItemMaster.sItemID >= '" & bsku & "'"
        s = s & " and tItemMaster.sItemID < '" & bsku & "ZZZZ'"
        s = s & " GROUP BY tLocationData.sLocationID, "
        s = s & "tLaneData.iLevel, tLaneData.iRow, tLaneData.iBlock, "
        s = s & "tContainerLocationData.iLocationID, "
        s = s & "tInventoryData.nQuantity, "
        's = s & "tLotData.dtProduction, "
        s = s & "tLotData.dtExpiration, "                       'jv081115
        s = s & "tItemMaster.sItemID, tItemMaster.sItemDescription"
        s = s & " ORDER BY tLocationData.sLocationID " ', tContainerLocationData.iPosition"
        'Set ds5 = db5.OpenRecordset(s, dbOpenSnapshot, dbSeeChanges, dbReadOnly)
        Set ds5 = Form1.db5.Execute(s)
        If ds5.BOF = False Then
            ds5.MoveFirst
            Do Until ds5.EOF
                s = "CS5" & Chr(9)
                s = s & Trim(ds5(0)) & Chr(9)
                s = s & wd_lotnum(ds5(6)) & Chr(9)
                's = s & Format(ds5(6), "m-d-yyyy") & Chr(9)
                s = s & Format(DateAdd("yyyy", -2, ds5(6)), "m-d-yyyy") & Chr(9)    'jv081115
                s = s & ds5(9) & Chr(9)
                s = s & ds5(5)
                If Val(blot) = 0 Or blot = wd_lotnum(ds5(6)) Then
                    Grid1.AddItem s
                End If
                ds5.MoveNext
            Loop
        End If
        ds5.Close ': db5.Close
    End If
    'If Val(bwhs) < 1 Or Val(bwhs) > 3 Then
    If Val(bwhs) = 0 Or Val(bwhs) = 4 Then
        sqlx = "select rackno,lot_num,count_qty+qty2,count(*) from rackpos"
        sqlx = sqlx & " where sku = '" & bsku & "'"
        If Val(blot) > 0 Then
            sqlx = sqlx & " and (lot_num = '" & blot & "'"
            sqlx = sqlx & " or lot2 = '" & blot & "')"
        End If
        sqlx = sqlx & " group by rackno,lot_num,count_qty+qty2"
        Set ps = Form1.wdb.Execute(sqlx)
        If ps.BOF = False Then
            ps.MoveFirst
            Do Until ps.EOF
                sqlx = "select aisle,rack,fo,hold from racks"
                sqlx = sqlx & " where id = " & ps(0)
                Set ds = Form1.wdb.Execute(sqlx)
                If ds.BOF = False Then
                    ds.MoveFirst
                    sqlx = "4" & Chr(9)
                    r = Trim(ds!aisle & "-" & ds!rack)
                    sqlx = sqlx & r & Chr(9)
                    If ds!fo = "1" And ds!aisle <> "M" Then
                        sqlx = sqlx & "1stOut" & Chr(9)
                    Else
                        If ds!hold = "1" And ds!aisle <> "M" Then
                            sqlx = sqlx & "OnHold" & Chr(9)
                        Else
                            sqlx = sqlx & ps!lot_num & Chr(9)
                        End If
                    End If
                    If Val(ps!lot_num) > 0 Then
                        sqlx = sqlx & calc_date(ps!lot_num) & Chr(9)
                    Else
                        sqlx = sqlx & "????" & Chr(9)
                    End If
                    sqlx = sqlx & ps(3) & Chr(9)
                    sqlx = sqlx & ps(2)
                    Grid1.AddItem sqlx
                End If
                ds.Close
                ps.MoveNext
            Loop
        End If
        ps.Close
    End If
    'db.Close
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 1
    Grid1.Sort = 5
    Grid1.FormatString = "^Whs|^Location|^Lot#|^Date|^Pallets|^Size"
    Grid1.ColWidth(0) = 600: Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 900: Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 900: Grid1.ColWidth(5) = 800
    Grid1.Redraw = True
End Sub

Private Sub blot_Change()
    Dim llit As String, wlit As String
    llit = "-All Lots"
    wlit = "-All Warehouses"
    If Val(blot) > 0 Then llit = "-Lot " & blot
    If Val(bwhs) > 0 And Val(bwhs) < 5 Then wlit = "-SR" & bwhs
    Form4.Caption = "Locations" & llit & wlit & " " & Form1.plantdesc
    Call refresh_grid
End Sub

Private Sub bsku_Change()
    Dim llit As String, wlit As String
    llit = "-All Lots"
    wlit = "-All Warehouses"
    If Val(blot) > 0 Then llit = "-Lot " & blot
    If Val(bwhs) > 0 And Val(bwhs) < 5 Then wlit = "-SR" & bwhs
    Form4.Caption = "Locations" & llit & wlit & " " & Form1.plantdesc
    Call refresh_grid
End Sub

Private Sub bwhs_Change()
    Dim llit As String, wlit As String
    llit = "-All Lots"
    wlit = "-All Warehouses"
    If Val(blot) > 0 Then llit = "-Lot " & blot
    If Val(bwhs) > 0 And Val(bwhs) < 5 Then wlit = "-SR" & bwhs
    Form4.Caption = "Locations" & llit & wlit & " " & Form1.plantdesc
    Call refresh_grid
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Form4.Caption = Form4.Caption & " " & Form1.plantdesc
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form4" Then
            Form4.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form4.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form4.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form4.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Resize()
    If Form4.Height > 3540 Then Grid1.Height = Form4.Height - 885
    Grid1.Width = Me.Width - 100
End Sub

Private Sub Form_Terminate()
    Dim i As Integer
    If Form4.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form4" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form4.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form4.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form4.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form4.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Terminate
End Sub

