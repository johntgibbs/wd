VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form r12pickorders 
   Caption         =   "R12 Pick Orders"
   ClientHeight    =   11175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13785
   LinkTopic       =   "Form2"
   ScaleHeight     =   11175
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Post to Pick Tasks"
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
      Left            =   5160
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox sdate 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   3855
      Left            =   0
      TabIndex        =   2
      Top             =   7560
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6800
      _Version        =   327680
      BackColorFixed  =   12648447
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   7011
      _Version        =   327680
      BackColorFixed  =   12648384
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4260
      _Version        =   327680
      BackColorFixed  =   12648447
      FocusRect       =   0
   End
   Begin VB.Label bbsr 
      Caption         =   "..."
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
      Left            =   5880
      TabIndex        =   8
      Top             =   360
      Width           =   7815
   End
   Begin VB.Label tkkey 
      Caption         =   "..."
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
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Shipped Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "r12pickorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function fixquotes(s As String) As String
    Dim i As Integer, k As Integer, rs As String
    rs = ""
    i = 1: k = 1
    Do Until i = 0
        i = InStr(k, s, "'")
        If i = 0 Then
            rs = rs & Mid(s, k, Len(s))
            Exit Do
        Else
            rs = rs & Mid(s, k, i - k) & "''"
            k = i + 1
        End If
    Loop
    fixquotes = rs
End Function

Public Function fixamps(s As String) As String
    Dim i As Integer, k As Integer, rs As String
    rs = ""
    i = 1: k = 1
    Do Until i = 0
        i = InStr(k, s, "&")
        If i = 0 Then
            rs = rs & Mid(s, k, Len(s))
            Exit Do
        Else
            rs = rs & Mid(s, k, i - k) & "&&"
            k = i + 1
        End If
    Loop
    fixamps = rs
End Function

Function wd_seq(tbname As String, dbname As String) As Long
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String, i As Long
    'On Error GoTo vberror
    i = 1
    Set db = CreateObject("ADODB.Connection")
    db.Open dbname
    s = "select sequence_id from sequences where seq = '" & tbname & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        i = ds!sequence_id + 1
        s = "update sequences set sequence_id = " & i
        s = s & " where seq = '" & tbname & "'"
        db.Execute s
    End If
    ds.Close: db.Close
    wd_seq = i
End Function

Function jobbing_account(bno As String, jname As String) As String
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String, acct As String
    Set db = CreateObject("ADODB.Connection")
    If Form1.Combo1 = "500" Then db.Open "ODBC;DATABASE=WDship;UID=bbcship500;PWD=brenham500;DSN=wdship500"
    If Form1.Combo1 = "501" Then db.Open "ODBC;DATABASE=BAship;UID=bbcship501;PWD=Barrow501;DSN=wdship501"
    If Form1.Combo1 = "502" Then db.Open "ODBC;DATABASE=SYship;UID=bbcship502;PWD=Alabama502;DSN=wdship502"
    'Set db = OpenDatabase("v:\data\shipping.mdb")
    s = "select account from jobbing where branch = " & bno
    s = s & " and acctdesc = '" & fixquotes(jname) & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        acct = ds!account
    Else
        acct = "000000"
    End If
    ds.Close: db.Close
    jobbing_account = acct
End Function

Private Sub post_sae()
    Dim cfile As String, i As Integer, k As Integer
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String
    Dim bno As Integer
    Dim palid As String, zid As Long
    'On Error GoTo vberror
    Screen.MousePointer = 11
    Set db = CreateObject("ADODB.Connection")
    db.Open Me.bbsr
    
    
    If Grid3.TextMatrix(1, 1) = "16" Or Grid3.TextMatrix(1, 1) = "15" Then
        bno = Val(Grid3.TextMatrix(1, 1))
        s = "select id,status,userid from picktasks where branch = " & Grid3.TextMatrix(1, 1)
        s = s & " and brname = '" & fixquotes(Grid3.TextMatrix(1, 2)) & "'"
        s = s & " and shipdate = '" & Grid3.TextMatrix(1, 3) & "'"
        's = s & " and palnum = " & Grid3.TextMatrix(1, 4)
        palid = Grid3.TextMatrix(1, 11)
    Else
        bno = Val(Grid3.TextMatrix(1, 1))
        s = "select id,status,userid from picktasks where branch = " & Grid3.TextMatrix(1, 1)
        s = s & " and shipdate = '" & Grid3.TextMatrix(1, 3) & "'"
        's = s & " and palnum = " & Grid3.TextMatrix(1, 4)
        palid = Grid3.TextMatrix(1, 11)
    End If
    'MsgBox s
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        s = "This pallet order already exists in pick tasks."
        If MsgBox(s, vbYesNo + vbQuestion, "do you want to continue....") = vbNo Then
            ds.Close: db.Close
            Screen.MousePointer = 0
            Exit Sub
        End If
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update picktasks set status = 'COMP', userid = ' ' where id = " & ds!id
            MsgBox s
            db.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    For i = 1 To Grid3.Rows - 1
        s = "select * from picktasks where status in ('SHIPPED', 'COMP') order by id"
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            s = "update picktasks set branch = " & Grid3.TextMatrix(i, 1)
            s = s & ", brname = '" & fixquotes(Grid3.TextMatrix(i, 2)) & "'"
            s = s & ", shipdate = '" & Grid3.TextMatrix(i, 3) & "'"
            s = s & ", palnum = '" & Grid3.TextMatrix(i, 4) & "'"
            s = s & ", opseq = " & Grid3.TextMatrix(i, 5)
            s = s & ", sku = '" & Grid3.TextMatrix(i, 6) & "'"
            s = s & ", lotnum = '...'"
            s = s & ", qty = " & Val(Grid3.TextMatrix(i, 8))
            s = s & ", uom = 'Wraps'"
            s = s & ", units = " & Val(Grid3.TextMatrix(i, 10))
            s = s & ", palletid = '" & palid & "'"
            s = s & ", status = 'PEND'"
            s = s & ", userid = '.'"
            s = s & ", location = 'ORDER PICK'"
            s = s & ", reqid = '" & Grid3.TextMatrix(i, 7) & "'"
            s = s & " Where id = " & ds!id
            'MsgBox s
            db.Execute s
        Else
            zid = wd_seq("PickTasks", Me.bbsr)
            'zid = i
            s = "INSERT INTO PickTasks (ID, Branch, BrName, ShipDate, PalNum, OPSeq,"
            s = s & " SKU, LotNum, Qty, Uom, Units, PalletID, Status, UserID, Location,"
            s = s & " ReqID) VALUES (" & zid & ","
            s = s & bno & ","
            s = s & "'" & fixquotes(Grid3.TextMatrix(i, 2)) & "',"
            s = s & "'" & Grid3.TextMatrix(i, 3) & "',"
            s = s & Grid3.TextMatrix(i, 4) & ","
            s = s & Grid3.TextMatrix(i, 5) & ","
            s = s & "'" & Grid3.TextMatrix(i, 6) & "',"
            s = s & "'...',"
            s = s & Val(Grid3.TextMatrix(i, 8)) & ","
            s = s & "'Wraps',"
            s = s & Val(Grid3.TextMatrix(i, 10)) & ","
            s = s & "'" & palid & "',"
            s = s & "'PEND',"
            s = s & "'.',"
            s = s & "'ORDER PICK',"
            s = s & "'" & Grid3.TextMatrix(i, 7) & "')"
            'MsgBox s
            db.Execute s
        End If
        ds.Close
    Next i
    db.Close
    Screen.MousePointer = 0
    'Exit Sub
'vberror:
    'eno = Err.Number: edesc = Err.Description: Err.Clear
    'Call vb_elog(eno, edesc, Me.Name, "post_sae", Form1.UserId)
    'If eno = -2147467259 Then
    '    Resume
    'Else
    '    MsgBox edesc, vbOKOnly, Me.Name & " post_sae - Error Number: " & eno
    '    End
    'End If
End Sub

Public Sub refresh_grid1(sd As String)
    Dim q As String, i As Integer, k As Integer, s As String
    Dim dsn As String, UserId As String, pwd As String
    Screen.MousePointer = 11
    dsn = "pbelle"                            'R12Test
    UserId = "Apps"                             'R12Test
    pwd = "pb3113tx"                             'R12Test
    s = DateDiff("d", sdate, Now)
    If Form1.Combo1 = "500" Then s = "001"
    If Form1.Combo1 = "501" Then s = "047"
    If Form1.Combo1 = "502" Then s = "052"
    'MsgBox s
    If AllocateODBChEnv(hEnv) <> SQL_SUCCESS Then Exit Sub
    If ConnectToDataSource(hEnv, hdbc, hstmt, dsn, UserId, pwd) <> SQL_SUCCESS Then
        i = FreeODBChEnv(hEnv)
        Exit Sub
    End If

    q = "  "
    q = "select shipment_number, transaction_date, transaction_reference, sum(transaction_quantity)"
    q = q & " From mtl_material_transactions"
    'q = q & " where subinventory_code = '001'"
    q = q & " where subinventory_code = '" & s & "'"
    q = q & " and transaction_date = TO_DATE('" & Format(sdate, "dd-mmm-yyyy") & "')"
    q = q & " and transfer_subinventory > '001'"
    q = q & " and transaction_source_name = 'TRAILER TRANSFER'"
    q = q & " group by shipment_number, transaction_date, transaction_reference"
    q = q & " order by shipment_number"
    'MsgBox q
    
    i = LoadGrid(Grid1, q, hstmt, 1, "")
    i = DisconnectFromDataSource(hdbc, hstmt)
    i = FreeODBChEnv(hEnv)
    Screen.MousePointer = 0
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Rows - 1
        For i = 1 To Grid1.Rows - 1
            s = Grid1.TextMatrix(i, 1)
            Grid1.TextMatrix(i, 1) = Mid(s, 6, 5) & "-" & Mid(s, 1, 4)
        Next i
    End If
    Grid1.FormatString = "^Ticket|^Date|<Reference|^Qty"
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 2200
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 1200
    
    If Grid1.Rows > 1 Then
        Grid1.Row = 1
        Call refresh_grid2(Grid1.TextMatrix(Grid1.Row, 0))
    End If
End Sub

Public Sub refresh_grid2(tkt As String)
    Dim q As String, i As Integer, k As Integer, s As String
    Dim dsn As String, UserId As String, pwd As String
    Screen.MousePointer = 11
    dsn = "pbelle"                            'R12Test
    UserId = "Apps"                             'R12Test
    pwd = "pb3113tx"                             'R12Test
    If AllocateODBChEnv(hEnv) <> SQL_SUCCESS Then Exit Sub
    If ConnectToDataSource(hEnv, hdbc, hstmt, dsn, UserId, pwd) <> SQL_SUCCESS Then
        i = FreeODBChEnv(hEnv)
        Exit Sub
    End If

    q = "select t.subinventory_code, transfer_subinventory, i.segment1, i.description, t.transaction_quantity,"
    q = q & " t.transaction_reference , t.transaction_date, t.shipment_number"
    q = q & " from mtl_material_transactions t, mtl_system_items_b i"
    q = q & " where t.shipment_number = '" & tkt & "'"
    'q = q & " and t.transfer_subinventory > '001'"
    q = q & " and t.transaction_source_name = 'TRAILER TRANSFER'"
    q = q & " and i.inventory_item_id = t.inventory_item_id"
    q = q & " and i.organization_id = t.organization_id"
    q = q & " order by t.shipment_number, i.segment1, t.subinventory_code"
    'MsgBox q
    
    i = LoadGrid(Grid2, q, hstmt, 1, "")
    i = DisconnectFromDataSource(hdbc, hstmt)
    i = FreeODBChEnv(hEnv)
    Screen.MousePointer = 0
    If Grid2.Rows > 2 Then
        Grid2.RemoveItem Grid2.Rows - 1
        For i = 1 To Grid2.Rows - 1
            s = Grid2.TextMatrix(i, 6)
            Grid2.TextMatrix(i, 6) = Mid(s, 6, 5) & "-" & Mid(s, 1, 4)
        Next i
    End If
    Grid2.FormatString = "^SubInv|^TranInv|^SKU|<Product|^Qty|<Reference|^Date|^Ticket"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 2200
    Grid2.ColWidth(4) = 1200
    Grid2.ColWidth(5) = 2400
    Grid2.ColWidth(6) = 1200
    Grid2.ColWidth(7) = 1200
    Command2.Visible = False
    If Grid2.TextMatrix(1, 7) = tkt And tkt > "0" Then Call refresh_grid3(tkt)
End Sub

Private Sub refresh_grid3(tkt As String)
    Dim s As String, i As Integer, pid As String, wc As Integer
    Dim wdb As adodb.Connection, ds As adodb.Recordset
    Screen.MousePointer = 11
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 16
    For i = 1 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 7) = tkt Then
            s = "0" & Chr(9)
            s = s & Val(Grid2.TextMatrix(i, 1)) & Chr(9)
            If Grid2.TextMatrix(i, 1) = "016" Or Grid2.TextMatrix(i, 1) = "015" Then
                s = s & UCase(Grid2.TextMatrix(i, 5)) & Chr(9)
                pid = Format(Val(Grid2.TextMatrix(i, 1)), "00")
                'pid = pid & "......"
                pid = pid & jobbing_account(Val(Grid2.TextMatrix(i, 1)), UCase(Grid2.TextMatrix(i, 5)))
                pid = pid & "01"
                pid = pid & Format(Grid2.TextMatrix(i, 6), "MMddyy")
            Else
                s = s & Left(Grid2.TextMatrix(i, 5), Len(Grid2.TextMatrix(i, 5)) - 3) & Chr(9)
                pid = Grid2.TextMatrix(i, 1)
                pid = pid & " " & Format(Grid2.TextMatrix(i, 6), "MMddyy")
                pid = pid & " B 001"
            End If
            s = s & Grid2.TextMatrix(i, 6) & Chr(9)
            s = s & "1" & Chr(9)
            s = s & "0" & Chr(9)
            s = s & Grid2.TextMatrix(i, 2) & Chr(9)
            s = s & "..." & Chr(9)
            s = s & "0" & Chr(9)
            s = s & "Wraps" & Chr(9)
            s = s & Format(Val(Grid2.TextMatrix(i, 4)) * -1, "0") & Chr(9)
            s = s & pid & Chr(9)
            s = s & "PEND" & Chr(9)
            s = s & " " & Chr(9)
            s = s & "ORDER PICK" & Chr(9)
            s = s & tkt
            Grid3.AddItem s
        End If
    Next i
    
    Set wdb = CreateObject("ADODB.Connection")
    wdb.Open Me.bbsr
    For i = 1 To Grid3.Rows - 1
        s = "select opseq from oplist where sku = '" & Grid3.TextMatrix(i, 6) & "'"
        Set ds = wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Grid3.TextMatrix(i, 5) = ds!opseq
        End If
        ds.Close
        wc = 1
        s = "select uom_per_pallet, qty_per_pallet from sku_config where sku = '" & Grid3.TextMatrix(i, 6) & "'"
        Set ds = wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            If ds(0) > 0 And ds(1) > 0 Then wc = ds(0) / ds(1)
        End If
        ds.Close
        Grid3.TextMatrix(i, 8) = Format(Val(Grid3.TextMatrix(i, 10)) / wc, "0")
    Next i
    wdb.Close
    
    Screen.MousePointer = 0
    Grid3.FormatString = "^ID|^Branch|<Location|^Date|^Pallet #|^OPSeq|^SKU|^Lot|^Qty|^Uom|^Units|^Barcode|^Status|^User|^Source|^Reqid"
    Grid3.ColWidth(0) = 800
    Grid3.ColWidth(1) = 800
    Grid3.ColWidth(2) = 2000
    Grid3.ColWidth(3) = 1000
    Grid3.ColWidth(4) = 1000
    Grid3.ColWidth(5) = 800
    Grid3.ColWidth(6) = 800
    Grid3.ColWidth(7) = 800
    Grid3.ColWidth(8) = 800
    Grid3.ColWidth(9) = 800
    Grid3.ColWidth(10) = 800
    Grid3.ColWidth(11) = 1800
    Grid3.ColWidth(12) = 800
    Grid3.ColWidth(13) = 800
    Grid3.ColWidth(14) = 1400
    Grid3.ColWidth(15) = 1000
    Command2.Visible = True
End Sub

Private Sub Command1_Click()
    Call refresh_grid1(sdate)
End Sub

Private Sub Command2_Click()
    Call post_sae
End Sub

Private Sub Form_Load()
    sdate = Format(Now, "MM-dd-yyyy")
    'If Form1.Combo1 = "500" Then Me.bbsr = "ODBC;DATABASE=WDRacks;DSN=wdracks"
    If Form1.Combo1 = "500" Then Me.bbsr = "ODBC;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500;DSN=wdsql500"
    If Form1.Combo1 = "501" Then Me.bbsr = "ODBC;DATABASE=BARacks;UID=bbcwd501;PWD=barrow501;DSN=wdsql501"
    If Form1.Combo1 = "502" Then Me.bbsr = "ODBC;DATABASE=SYRacks;UID=bbcwd502;PWD=alabama502;DSN=wdsql502"
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    Grid2.Width = Me.Width - 100
    Grid3.Width = Me.Width - 100
End Sub

Private Sub Grid1_RowColChange()
    If Val(Grid1.TextMatrix(Grid1.Row, 3)) < 0 Then tkkey.Caption = Grid1.TextMatrix(Grid1.Row, 0)
End Sub

Private Sub tkkey_Change()
    Call refresh_grid2(tkkey.Caption)
End Sub

