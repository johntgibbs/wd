VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form19 
   Caption         =   "Form19"
   ClientHeight    =   13110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13200
   LinkTopic       =   "Form19"
   ScaleHeight     =   13110
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Run Script"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   6960
      Width           =   1935
   End
   Begin VB.ListBox List1 
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
      Height          =   5325
      Left            =   0
      TabIndex        =   3
      Top             =   7320
      Width           =   11535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Build Script"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh Records"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10186
      _Version        =   327680
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      Caption         =   "..."
      Height          =   375
      Left            =   10200
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function wd_lotnum(pdate As String) As String
    Dim sdate As String, s As String
    pdate = Format(pdate, "m-d-yyyy")
    sdate = "1-1-" & Right(pdate, 4)
    s = Format(DateDiff("d", sdate, pdate) + 1, "000")
    s = Right(pdate, 2) & s
    wd_lotnum = s
End Function

Private Sub refresh_grid()
        Dim ds As ADODB.Recordset, s As String, i As Long, aflag As Boolean
        Dim db5 As ADODB.Connection
        Screen.MousePointer = 11
        Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 9
        
        's = "select * from holdlist order by sku, lot_num, opcode"
        'Set ds = Wdb.Execute(s)
        'If ds.BOF = False Then
        '    ds.MoveFirst
        '    Do Until ds.EOF
        '        s = ds!id & Chr(9)
        '        s = s & ds!sku & Chr(9)
        '        s = s & ds!lot_num & Chr(9)
        '        s = s & ds!opcode & Chr(9)
        '        s = s & ds!spallet & Chr(9)
        '        s = s & ds!epallet & Chr(9)
        '        s = s & ds!hsource & Chr(9)
        '        s = s & ds!userid & Chr(9)
        '        s = s & ds!holddate
        '        Grid1.AddItem s
        '        ds.MoveNext
        '    Loop
        'End If
        'ds.Close
        
        s = "select * from rackpos where sku < '9999' and count_qty > 0 and lot_num < '15107' order by sku, lot_num"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            'MsgBox ds!sku & " " & ds!lot_num & " " & Mid(ds!barcode, 12, 1)
            Do Until ds.EOF
                aflag = True
                For i = 0 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 1) = ds!sku And Grid1.TextMatrix(i, 2) = ds!lot_num And Grid1.TextMatrix(i, 3) = Mid(ds!barcode, 12, 1) Then
                        aflag = False
                        Exit For
                    End If
                Next i
                If aflag = True Then
                    s = ds!id & Chr(9) & ds!sku & Chr(9)
                    s = s & ds!lot_num & Chr(9)
                    s = s & Mid(ds!barcode, 12, 1) & Chr(9)
                    s = s & "001" & Chr(9) & "001" & Chr(9) & "TERM" & Chr(9) & "WMS" & Chr(9)
                    s = s & Format(Now, "yyMMdd hh:mm:ss")
                    Grid1.AddItem s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
        
        s = "select * from position where count_qty > 0 and lot_num < '15107' order by sku, lot_num"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            'MsgBox ds!sku & " " & ds!lot_num & " " & Mid(ds!barcode, 12, 1)
            Do Until ds.EOF
                aflag = True
                For i = 0 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 1) = ds!sku And Grid1.TextMatrix(i, 2) = ds!lot_num And Grid1.TextMatrix(i, 3) = Mid(ds!barcode, 12, 1) Then
                        aflag = False
                        Exit For
                    End If
                Next i
                If aflag = True Then
                    s = ds!id & Chr(9) & ds!sku & Chr(9)
                    s = s & ds!lot_num & Chr(9)
                    s = s & Mid(ds!barcode, 12, 1) & Chr(9)
                    s = s & "001" & Chr(9) & "001" & Chr(9) & "TERM" & Chr(9) & "WMS" & Chr(9)
                    s = s & Format(Now, "yyMMdd hh:mm:ss")
                    Grid1.AddItem s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
        
        If Form1.plantno = "52" Then
            Set db5 = CreateObject("ADODB.Connection")
            'db5.Open "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
            db5.Open "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=sywms;PWD=!Sylacauga_WMS1907"
            s = "SELECT tLocationData.sLocationID, "
            s = s & "tLaneData.iLevel, tLaneData.iRow, tLaneData.iBlock, "
            s = s & "tContainerLocationData.iLocationID, "
            s = s & "tInventoryData.nQuantity, "
            s = s & "tLotData.dtProduction, "
            s = s & "tItemMaster.sItemID, tItemMaster.sItemDescription, count(*) "
            s = s & "FROM tLocationData, tLaneData, tContainerLocationData, tInventoryData, "
            s = s & "tLotData, tItemMaster"
            s = s & " WHERE tLaneData.iLocationID = tLocationData.iLocationID"
            s = s & " AND tContainerLocationData.iLocationID = tLaneData.iLocationID"
            s = s & " AND tInventoryData.iContainerDataSysID = tContainerLocationData.iContainerDataSysID"
            s = s & " AND tLotData.iLotDataSysID = tInventoryData.iLotDataSysID"
            s = s & " AND tItemMaster.iItemMasterSysID = tLotData.iItemMasterSysID"
        
            s = s & " and tItemMaster.sItemID >= '100'"
            s = s & " and tItemMaster.sItemID < '999ZZZZ'"
            s = s & " GROUP BY tLocationData.sLocationID, "
            s = s & "tLaneData.iLevel, tLaneData.iRow, tLaneData.iBlock, "
            s = s & "tContainerLocationData.iLocationID, "
            s = s & "tInventoryData.nQuantity, "
            s = s & "tLotData.dtProduction, "
            s = s & "tItemMaster.sItemID, tItemMaster.sItemDescription"
            s = s & " ORDER BY tItemMaster.sItemID, tLocationData.sLocationID " ', tContainerLocationData.iPosition"
            Set ds = db5.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                Do Until ds.EOF
                    aflag = True
                    For i = 0 To Grid1.Rows - 1
                        If Grid1.TextMatrix(i, 1) = Trim(ds(7)) And Grid1.TextMatrix(i, 2) = wd_lotnum(ds(6)) Then
                            aflag = False
                            Exit For
                        End If
                    Next i
                    If aflag = True Then
                        s = "CS5" & Chr(9)
                        s = s & Trim(ds(7)) & Chr(9)
                        s = s & wd_lotnum(ds(6)) & Chr(9)
                        's = s & "X" & Chr(9)
                        s = s & " " & Chr(9)
                        s = s & "001" & Chr(9)
                        s = s & "EOR" & Chr(9)
                        s = s & "TERM" & Chr(9)
                        s = s & "WMS" & Chr(9)
                        s = s & Format(Now, "yyMMdd hh:mm:ss")
                        'If wd_lotnum(ds(6)) < "15107" Then
                            Grid1.AddItem s
                        'End If
                    End If
                    ds.MoveNext
                Loop
            End If
            ds.Close: db5.Close
        End If
        
        
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 1: Grid1.ColSel = 3
        Grid1.Sort = 5
        
        Grid1.FormatString = "^ID|^SKU|^Lot|^OP|^SPallet|^EPallet|^Source|^User|<Date"
        Grid1.ColWidth(0) = 1000
        Grid1.ColWidth(1) = 1000
        Grid1.ColWidth(2) = 1000
        Grid1.ColWidth(3) = 1000
        Grid1.ColWidth(4) = 1000
        Grid1.ColWidth(5) = 1000
        Grid1.ColWidth(6) = 1000
        Grid1.ColWidth(7) = 1000
        Grid1.ColWidth(8) = 1600
        Label1.Caption = Grid1.Rows - 1 & " Records"
        Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub Command2_Click()
    Dim ds As ADODB.Recordset, s As String, i As Long, zid As Long
    Screen.MousePointer = 11
    List1.Clear
    s = "select sequence_id from sequences where seq = 'HoldList'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "delete from holdlist where id <= " & ds(0)
        MsgBox s
        Wdb.Execute s
        s = "Update sequences set sequence_id = 0 where seq = 'HoldList'"
        MsgBox s
        Wdb.Execute s
    End If
    ds.Close
    'Exit Sub
    For i = 1 To Grid1.Rows - 1
        zid = wd_seq("HoldList")
        s = "Insert into HoldList (id, sku, lot_num, opcode, spallet, epallet, hsource, userid, holddate)"
        s = s & " Values (" & zid
        s = s & ", '" & Grid1.TextMatrix(i, 1) & "'"
        s = s & ", '" & Grid1.TextMatrix(i, 2) & "'"
        s = s & ", '" & Grid1.TextMatrix(i, 3) & "'"
        s = s & ", '" & Grid1.TextMatrix(i, 4) & "'"
        s = s & ", '" & Grid1.TextMatrix(i, 5) & "'"
        s = s & ", '" & Grid1.TextMatrix(i, 6) & "'"
        s = s & ", '" & Grid1.TextMatrix(i, 7) & "'"
        s = s & ", '" & Grid1.TextMatrix(i, 8) & "')"
        List1.AddItem s
    Next i
    Screen.MousePointer = 0
End Sub

Private Sub Command3_Click()
    Dim i As Integer
    Screen.MousePointer = 11
    For i = 0 To List1.ListCount - 1 ' 10
        Wdb.Execute List1.List(i)
    Next i
    Screen.MousePointer = 0
End Sub

