VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Lot Totals"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   LinkTopic       =   "Form3"
   ScaleHeight     =   3660
   ScaleWidth      =   3870
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
      _Version        =   327680
      Cols            =   4
      BackColorFixed  =   12648447
      Appearance      =   0
   End
   Begin VB.Label lwhs 
      Caption         =   "lwhs"
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lprod 
      Caption         =   "lprod"
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
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lsku 
      Caption         =   "lsku"
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
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form3"
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
    pdate = Format(DateAdd("yyyy", -2, pdate), "m-d-yyyy")              'jv081115
    sdate = "1-1-" & Right(pdate, 4)
    s = Format(DateDiff("d", sdate, pdate) + 1, "000")
    s = Right(pdate, 2) & s
    wd_lotnum = s
End Function

Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, sqlx As String
    Dim ds5 As ADODB.Recordset
    'Dim db As Database, ds As Recordset, sqlx As String
    'Dim db5 As Database, ds5 As Recordset
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1
    If lwhs = "lwhs" Then Exit Sub
    If lsku = "lsku" Then Exit Sub
    's = "ODBC;DATABASE=WDRacks;DSN=wdracks"
    's = "ODBC;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500;DSN=wdsql500"
    'Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, True, Form1.bbsr)
    'If Val(lwhs) <> 4 And Form1.Check1.Value = 1 Then
    If Val(lwhs) <> 4 And Form1.plantno = "50" Then
        sqlx = "select whse_num,lot_num,gmasize,sum(qty) from lane" 'jv082813
        sqlx = sqlx & " where sku = '" & lsku & "'"
        If Val(lwhs) > 0 And Val(lwhs) < 6 Then
            sqlx = sqlx & " and whse_num = " & lwhs
        End If
        'sqlx = sqlx & " and gmasize < 1"
        sqlx = sqlx & " group by whse_num,lot_num,gmasize"      'jv082813
        sqlx = sqlx & " order by gmasize,lot_num"               'jv082813
        Set ds = Form1.wdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ds!gmasize > 0 Then                  'jv082813
                    sqlx = "6" & Chr(9)
                Else
                    sqlx = ds!whse_num & Chr(9)
                End If
                sqlx = sqlx & ds!lot_num & Chr(9)
                sqlx = sqlx & calc_date(ds!lot_num) & Chr(9)
                sqlx = sqlx & ds(3)
                Grid1.AddItem sqlx
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    
    If Val(bwhs) <> 4 And Form1.plantno = "52" Then
        s = "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
        'Set db5 = OpenDatabase(mysqldev, dbcdrivernoprompt, True, s)
        's = "SELECT tLotData.dtProduction, count(*)"
        s = "SELECT tLotData.dtExpiration, count(*)"                    'jv081115
        s = s & " FROM tLocationData, tLaneData, tContainerLocationData, tInventoryData, "
        s = s & "tLotData, tItemMaster"
        s = s & " WHERE tLaneData.iLocationID = tLocationData.iLocationID"
        s = s & " AND tContainerLocationData.iLocationID = tLaneData.iLocationID"
        s = s & " AND tInventoryData.iContainerDataSysID = tContainerLocationData.iContainerDataSysID"
        s = s & " AND tLotData.iLotDataSysID = tInventoryData.iLotDataSysID"
        s = s & " AND tItemMaster.iItemMasterSysID = tLotData.iItemMasterSysID"
        's = s & " and tItemMaster.sItemID = '" & lsku & "'"
        s = s & " and tItemMaster.sItemID >= '" & lsku & "'"                    'jv073015
        s = s & " and tItemMaster.sItemID < '" & lsku & "ZZZZZ'"                'jv073015
        's = s & " GROUP BY tLotData.dtProduction"
        s = s & " GROUP BY tLotData.dtExpiration"                       'jv081115
        'MsgBox s
        'MsgBox lsku & " " & lwhs
        'Set ds5 = db5.OpenRecordset(s, dbOpenSnapshot, dbSeeChanges, dbReadOnly)
        Set ds5 = Form1.db5.Execute(s)
        If ds5.BOF = False Then
            ds5.MoveFirst
            Do Until ds5.EOF
                s = "CS5" & Chr(9)
                s = s & wd_lotnum(ds5(0)) & Chr(9)
                's = s & Format(Trim(ds5(0)), "m-d-yyyy") & Chr(9)
                s = s & Format(DateAdd("yyyy", -2, Trim(ds5(0))), "m-d-yyyy") & Chr(9)      'jv081115
                s = s & ds5(1)
                Grid1.AddItem s
                ds5.MoveNext
            Loop
        End If
        ds5.Close ': db5.Close
    End If
    
    
    If Val(lwhs) = 0 Or Val(lwhs) = 4 Then
        sqlx = "select lot_num,count(*) from rackpos"
        sqlx = sqlx & " where sku = '" & lsku & "'"
        sqlx = sqlx & " group by lot_num order by lot_num"
        Set ds = Form1.wdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                sqlx = "4" & Chr(9)
                sqlx = sqlx & ds!lot_num & Chr(9)
                If Val(ds!lot_num) > 0 Then
                    sqlx = sqlx & calc_date(ds!lot_num) & Chr(9)
                Else
                    sqlx = sqlx & "????" & Chr(9)
                End If
                sqlx = sqlx & ds(1)
                Grid1.AddItem sqlx
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 1: Grid1.ColSel = 1
    Grid1.Sort = 5
    Grid1.FormatString = "^Whs|^Lot#|^Date|^Qty"
    Grid1.ColWidth(0) = 600: Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 1200: Grid1.ColWidth(3) = 800
    Grid1.Redraw = True
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Form3.Caption = Form3.Caption & " " & Form1.plantdesc
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form3" Then
            Form3.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form3.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form3.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form3.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Resize()
    If Form3.Height > 3540 Then Grid1.Height = Form3.Height - 885
    Grid1.Width = Me.Width - 100
End Sub

Private Sub Form_Terminate()
    Dim i As Integer
    If Form3.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form3" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form3.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form3.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form3.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form3.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Terminate
End Sub

Private Sub Grid1_Click()
    If Grid1.Row > 0 Then
        Form4.bwhs = Val(Grid1.TextMatrix(Grid1.Row, 0))
        Form4.blot = Grid1.TextMatrix(Grid1.Row, 1)
    End If
End Sub

Private Sub Grid1_DblClick()
    'MsgBox wd_lotnum(Grid1.TextMatrix(Grid1.Row, 2))
End Sub

Private Sub lsku_Change()
    Call refresh_grid
End Sub

Private Sub lwhs_Change()
    If Val(lwhs) > 0 And Val(lwhs) < 4 Then
        Form3.Caption = "Lot Totals - SR" & lwhs & " " & Form1.plantdesc
    Else
        Form3.Caption = "Lot Totals - All Warehouses " & Form1.plantdesc
    End If
    Call refresh_grid
End Sub

