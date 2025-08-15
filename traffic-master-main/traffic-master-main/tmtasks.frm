VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form tmtasks 
   Caption         =   "Tri Level Tasks"
   ClientHeight    =   9015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13050
   LinkTopic       =   "Form2"
   ScaleHeight     =   9015
   ScaleWidth      =   13050
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   7080
      Left            =   9480
      TabIndex        =   7
      Top             =   1800
      Width           =   3495
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4455
      Left            =   0
      TabIndex        =   3
      Top             =   3960
      Width           =   9375
      ExtentX         =   16536
      ExtentY         =   7858
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "tmtasks.frx":0000
      Top             =   2160
      Width           =   4935
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2990
      _Version        =   327680
      BackColorFixed  =   12648447
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3201
      _Version        =   327680
      AllowUserResizing=   3
   End
   Begin VB.Label bc2 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label bc1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label ttrig 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Menu refmenu 
      Caption         =   "Refresh Tasks"
   End
End
Attribute VB_Name = "tmtasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w1cap As Integer
Dim w2cap As Integer
Dim w3cap As Integer
Dim w4cap As Integer
Dim w5cap As Integer

Sub post_tm_log(pwhs As Integer, i As Integer, dtask As String)
    Dim cfile As String
    cfile = "\\BBC-01-PRODTRK\wd\testlogs\tml" & Format(Now, "MMddyyyy") & ".txt"
    Open cfile For Append As #1
    Write #1, Grid1.TextMatrix(i, 0);
    Write #1, "TRAFFIC MASTER";
    Write #1, " ";
    Write #1, Grid1.TextMatrix(i, 3);
    Write #1, "SR" & pwhs;
    Write #1, Grid1.TextMatrix(i, 5);
    Write #1, Grid1.TextMatrix(i, 6);
    Write #1, Grid1.TextMatrix(i, 7);
    Write #1, Grid1.TextMatrix(i, 8);
    Write #1, Grid1.TextMatrix(i, 9);
    Write #1, Grid1.TextMatrix(i, 10);
    Write #1, Grid1.TextMatrix(i, 11);
    Write #1, Grid1.TextMatrix(i, 12);
    Write #1, Grid1.TextMatrix(i, 13);
    Write #1, Grid1.TextMatrix(i, 14);
    Write #1, Grid1.TextMatrix(i, 15);
    Write #1, dtask
    Close #1
End Sub

Private Sub record_pallet(i As Integer)
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Dim pid As Long, pstat As String, psku As String, recid As Long
    Screen.MousePointer = 11
    'pstat = "Wrapper"
    pstat = "Warehouse"
    recid = 0
    Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    db.Open Form1.tbbsr
    
    s = "select * from pallets where barcode = '" & Grid1.TextMatrix(i, 6) & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        recid = ds!id
    Else
        ds.Close
        s = "select * from pallets where status in ('Shipped','Order Pick')"
        s = s & " order by trandate"
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            recid = ds!id
        End If
    End If
    ds.Close
    If recid > 0 Then
        s = "Update pallets set plateno = '" & Grid1.TextMatrix(i, 16) & "'"
        s = s & ",barcode = '" & Grid1.TextMatrix(i, 6) & "'"
        s = s & ",qty1 = " & Val(Grid1.TextMatrix(i, 10))
        s = s & ",lot1 = '" & Grid1.TextMatrix(i, 9) & "'"
        s = s & ",qty2 = " & Val(Grid1.TextMatrix(i, 12))
        s = s & ",lot2 = '" & Grid1.TextMatrix(i, 11) & "'"
        s = s & ",source = '" & Grid1.TextMatrix(i, 3) & "'"
        's = s & ",target = '" & Grid1.TextMatrix(i, 4) & "'"
        s = s & ",target = 'SR" & Grid2.TextMatrix(1, 0) & "'"
        s = s & ",bbc = 'Y'"
        s = s & ",status = '" & pstat & "'"
        s = s & ",trandate = '" & Grid1.TextMatrix(i, 15) & "'"
        s = s & ",sku = '" & psku & "'"
        s = s & " Where id = " & recid
        'MsgBox s
        db.Execute s
    Else
        pid = wd_seq("Pallets")
        s = "Insert Into pallets Values (" & pid
        s = s & ",'" & Grid1.TextMatrix(i, 16) & "'"
        s = s & ",'" & Grid1.TextMatrix(i, 6) & "'"
        s = s & "," & Val(Grid1.TextMatrix(i, 10))
        s = s & ",'" & Grid1.TextMatrix(i, 9) & "'"
        s = s & "," & Val(Grid1.TextMatrix(i, 12))
        s = s & ",'" & Grid1.TextMatrix(i, 11) & "'"
        s = s & ",'" & Grid1.TextMatrix(i, 3) & "'"
        's = s & ",'" & Grid1.TextMatrix(i, 4) & "'"
        s = s & ",'SR" & Grid2.TextMatrix(1, 0) & "'"
        s = s & ",'Y'"
        If Grid1.TextMatrix(i, 4) = "ORDER PICK" Then
            s = s & ",'Order Pick'"
        Else
            s = s & ",'" & pstat & "'"
        End If
        s = s & ",'" & Grid1.TextMatrix(i, 15) & "'"
        s = s & ",'" & psku & "')"
        db.Execute s
    End If
    db.Close
    Screen.MousePointer = 0
End Sub

Sub post_route_to_kep(wrapper As Integer, whs As Integer, palid As Integer)
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Dim croute As Integer
    If whs = 3 Then
        croute = 5
    Else
        croute = whs
    End If
    Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    db.Open Form1.tbbsr
    s = "update wrapper_config set pallet_id = " & palid
    s = s & ", sr_destination = " & whs
    s = s & ", conv_route = " & croute
    s = s & " where wrapper_id = " & wrapper
    'MsgBox s
    db.Execute s
    db.Close
End Sub

Sub post_queue_to_sr(whs As Integer, i As Integer)
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Dim qid As Long, hf As Boolean, nque As Integer
    hf = False
    Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    db.Open Form1.tbbsr
    'Process Queue
    s = "select * from queue_infc where id = " & new_pallet_queue(False)
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "update queue_infc set whse_num = " & whs
        s = s & ",sku = '" & Trim(Left(Grid1.TextMatrix(i, 5), 4)) & "'"
        s = s & ",lot_num = '" & Grid1.TextMatrix(i, 9) & "'"
        If hf Then
            s = s & ",drop_flag = 'H'"
        Else
            s = s & ",drop_flag = ' '"
        End If
        s = s & ",rack_num = " & Val(Grid1.TextMatrix(i, 7))
        s = s & ",units = " & Val(Grid1.TextMatrix(i, 10))
        s = s & ",lot_num2 = '" & Grid1.TextMatrix(i, 11) & "'"
        s = s & ",units2 = " & Val(Grid1.TextMatrix(i, 12))
        s = s & ",palletid = '" & Grid1.TextMatrix(i, 6) & "'"
        s = s & ",source = 'TML'"
        s = s & " where id = " & ds!id
        db.Execute s
    Else
        nque = 50
        qid = wd_seq("Queue_Infc")
        s = "INSERT INTO Queue_Infc (ID, Whse_Num, SKU, Lot_Num, Drop_Flag, Queue_Num,"
        s = s & " Rack_Num, Units, Lot_Num2, Units2, PalletID, Source)"
        s = s & " VALUES (" & qid & ","
        s = s & whs & ","
        s = s & "'" & Trim(Left(Grid1.TextMatrix(i, 5), 4)) & "',"
        s = s & "'" & Grid1.TextMatrix(i, 9) & "',"
        If hf Then
            s = s & "'H',"
        Else
            s = s & "' ',"
        End If
        s = s & nque & ","
        s = s & Val(Grid1.TextMatrix(i, 7)) & ","
        s = s & Val(Grid1.TextMatrix(i, 10)) & ","
        s = s & "'" & Grid1.TextMatrix(i, 11) & "',"
        s = s & Val(Grid1.TextMatrix(i, 12)) & ","
        s = s & "'" & Grid1.TextMatrix(i, 6) & "',"
        s = s & "'TML')"
        db.Execute s
        'ds.AddNew
    End If
    ds.Close
End Sub

Sub post_dai_exp_rcpt()
    Dim d As daiexprct, rkey As Long
    d.action = "ADD"
    d.sOrderID = Grid1.TextMatrix(Grid1.Row, 16)
    d.dExpectedDate = Format(Now, "MM/dd/yyyy hh:mm:ss")
    d.sItem = Left(Grid1.TextMatrix(Grid1.Row, 6), 3)
    d.sLot = Grid1.TextMatrix(Grid1.Row, 9)
    d.fExpectedQuantity = Val(Grid1.TextMatrix(Grid1.Row, 10)) + Val(Grid1.TextMatrix(Grid1.Row, 12))
    d.sStoreDestination = Grid2.TextMatrix(1, 0)
    If d.sStoreDestination = "1" Then d.sStoreDestination = "5"         'jv0521 test
    If d.sStoreDestination = "2" Or d.sStoreDestination = "3" Or d.sStoreDestination = "5" Then
        Text1.Text = Dai_expected_receipt(d)
        Open "c:\jvwork\daiExpectedReceiptMessage.xml" For Output As #1
        Print #1, Text1.Text
        Close #1
        DoEvents
        rkey = wd_seq("DAIRequests")
        'Call write_oracle_request("ExpectedReceiptMessage", Val(d.sOrderID))
        Call write_oracle_request("ExpectedReceiptMessage", rkey)
        WebBrowser1.Navigate2 "c:\jvwork\daiExpectedReceiptMessage.xml"
    End If
End Sub

Function part_pallet_whs(psku As String) As String
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    s = "select sku from opbays where sku = '" & psku & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        part_pallet_whs = "1"
    Else
        part_pallet_whs = "4"
    End If
    ds.Close: db.Close
End Function

Function bbpallet_units(psku As String) As Long
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    s = "select uom_per_pallet from sku_config where sku = '" & psku & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        bbpallet_units = ds(0)
    Else
        bbpallet_units = 1
    End If
    ds.Close: db.Close
End Function

Sub refresh_grid1()
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 17
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    s = "select * from paltasks"
    's = s & " where area = 'TRAFFIC MASTER'"
    's = s & " where target = 'TRAFFIC MASTER'"
    s = s & " where area in ('TRI-LEVEL 1','TRI-LEVEL 2','TRI-LEVEL 3','SNACK PLANT WRAPPER','ROBOT ZERO')"
    s = s & " and units > 0"
    's = s & " and status = 'PEND' order by trandate"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!area & Chr(9)
            s = s & ds!description & Chr(9)
            s = s & ds!source & Chr(9)
            s = s & ds!target & Chr(9)
            s = s & ds!product & Chr(9)
            s = s & ds!palletid & Chr(9)
            s = s & ds!qty & Chr(9)
            s = s & ds!uom & Chr(9)
            s = s & ds!lotnum & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!lotnum2 & Chr(9)
            s = s & ds!units2 & Chr(9)
            s = s & ds!status & Chr(9)
            s = s & ds!userid & Chr(9)
            s = s & ds!trandate & Chr(9)
            s = s & ds!reqid
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    'ds.Close
    's = "select * from paltasks"
    's = s & " where target = 'CRANE 3'"
    's = s & " and status = 'PEND' order by trandate"
    'Set ds = db.Execute(s)
    'If ds.BOF = False Then
    '    ds.MoveFirst
    '    Do Until ds.EOF
    '        s = ds!id & Chr(9)
    '        s = s & ds!area & Chr(9)
    '        s = s & ds!description & Chr(9)
    '        s = s & ds!source & Chr(9)
    '        s = s & ds!target & Chr(9)
    '        s = s & ds!product & Chr(9)
    '        s = s & ds!palletid & Chr(9)
    '        s = s & ds!qty & Chr(9)
    '        s = s & ds!uom & Chr(9)
    '        s = s & ds!lotnum & Chr(9)
    '        s = s & ds!units & Chr(9)
    '        s = s & ds!lotnum2 & Chr(9)
    '        s = s & ds!units2 & Chr(9)
    '        s = s & ds!status & Chr(9)
    '        s = s & ds!userid & Chr(9)
    '        s = s & ds!trandate & Chr(9)
    '        s = s & ds!reqid
    '        Grid1.AddItem s
    '        ds.MoveNext
    '    Loop
    'End If
    ds.Close: db.Close
    s = "^Id|<Area|<Desc|<Source|<Target|<Product|<BarCode|^Qty|^Uom|^Lot|^Units|^Lot2|^Units|^Status|^User|^Time|<Reqid"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 1800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 1800
    Grid1.ColWidth(4) = 1800
    Grid1.ColWidth(5) = 1800
    Grid1.ColWidth(6) = 1800
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 800
    Grid1.ColWidth(10) = 1600
    Grid1.ColWidth(11) = 800
    Grid1.ColWidth(12) = 800
    Grid1.ColWidth(13) = 800
    Grid1.ColWidth(14) = 800
    Grid1.ColWidth(15) = 800
    Grid1.ColWidth(16) = 800
End Sub

Private Sub Form_Load()
    Dim i As Integer
    w1cap = 4
    w2cap = 6
    w3cap = 8
    w4cap = 8
    w5cap = 12
    refresh_grid1
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 1) = "TRI-LEVEL 1" Then bc1 = Grid1.TextMatrix(i, 6)
        If Grid1.TextMatrix(i, 1) = "TRI-LEVEL 2" Then bc2 = Grid1.TextMatrix(i, 6)
    Next i
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
End Sub

Private Sub Grid1_DblClick()
    Dim psku As String, plot As String, i As Integer, s As String
    Dim j As Long, k As Long
    If List1.ListCount > 0 Then                 'test for repeats 5-17-13
        For i = 0 To List1.ListCount - 1
            s = Mid(List1.List(i), 7, 16)
            If Trim(Grid1.TextMatrix(Grid1.Row, 6)) = Trim(s) Then Exit Sub
        Next i
    End If
    psku = Left(Grid1.TextMatrix(Grid1.Row, 6), 3)
    plot = Grid1.TextMatrix(Grid1.Row, 9)
    k = bbpallet_units(psku)
    j = Val(Grid1.TextMatrix(Grid1.Row, 10)) + Val(Grid1.TextMatrix(Grid1.Row, 12))
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 6
    'MsgBox psku & " j=" & j & " k=" & k
    
    'If InStr(1, Grid1.TextMatrix(Grid1.Row, 5), "Wraps") Then
    If j < k Then       'Partial
        s = part_pallet_whs(psku) & Chr(9)
        Grid2.AddItem s
        'MsgBox psku & " j=" & j & " k=" & k
    Else
        queue_infc.trigkey = Val(queue_infc.trigkey) + 1
        DoEvents
        allocations.Label1 = Val(allocations.Label1) + 1
        DoEvents
        If Form1.srstat1.Value = 1 Then
            i = allocations.sku_alloc(psku, plot, 1)
            'If i > 0 Then Grid2.AddItem "1" & Chr(9) & queue_infc.ques1 & Chr(9) & i
            If i > 0 Then
                s = "1" & Chr(9) & w1cap & Chr(9)
                s = s & queue_infc.ques1 & Chr(9)
                s = s & queue_infc.sr_single_sku("1", psku) & Chr(9)
                s = s & Format(w1cap - Val(queue_infc.ques1), "0") & Chr(9) & i
                Grid2.AddItem s
            End If
        End If
        If Form1.srstat2.Value = 1 Then
            i = allocations.sku_alloc(psku, plot, 2)
            'If i > 0 Then Grid2.AddItem "2" & Chr(9) & queue_infc.ques2 & Chr(9) & i
            If i > 0 Then
                s = "2" & Chr(9) & w2cap & Chr(9)
                s = s & queue_infc.ques2 & Chr(9)
                s = s & queue_infc.sr_single_sku("2", psku) & Chr(9)
                s = s & Format(w2cap - Val(queue_infc.ques2), "0") & Chr(9) & i
                Grid2.AddItem s
            End If
        End If
        If Form1.srstat3.Value = 1 Then
            i = allocations.sku_alloc(psku, plot, 3)
            'If i > 0 Then Grid2.AddItem "3" & Chr(9) & queue_infc.ques3 & Chr(9) & i
            If i > 0 Then
                s = "3" & Chr(9) & w3cap & Chr(9)
                s = s & queue_infc.ques3 & Chr(9)
                s = s & queue_infc.sr_single_sku("3", psku) & Chr(9)
                s = s & Format(w3cap - Val(queue_infc.ques3), "0") & Chr(9) & i
                Grid2.AddItem s
            End If
        End If
        If Form1.srstat4.Value = 1 Then
            i = allocations.sku_alloc(psku, plot, 4)
            'If i > 0 Then Grid2.AddItem "3" & Chr(9) & queue_infc.ques3 & Chr(9) & i
            'If i > 0 Then
                s = "4" & Chr(9) & w4cap & Chr(9)
                s = s & "0" & Chr(9)
                If i > 0 Then
                    s = s & "1" & Chr(9)
                Else
                    s = s & "-1" & Chr(9)
                End If
                s = s & i & Chr(9) & i
                Grid2.AddItem s
            'End If
        End If
    End If
    
    Grid2.RowSel = Grid2.Row
    Grid2.Col = 3: Grid2.ColSel = 5
    Grid2.Sort = 4
    'Grid2.Sort = 5
    Grid2.FormatString = "^SR|^Length|^Queues|^Single|^Capacity|^ResvPals"
    Grid2.ColWidth(0) = 500
    Grid2.ColWidth(1) = 800
    Grid2.ColWidth(2) = 800
    Grid2.ColWidth(3) = 800
    Grid2.ColWidth(4) = 800
    Grid2.ColWidth(5) = 800
    If Grid2.Rows > 1 Then
        If Grid1.TextMatrix(Grid1.Row, 3) = "TRI-LEVEL 1" Then
            'Form1.ws1.Text = Val(Form1.ws1.Text) + 1
            'Grid1.TextMatrix(Grid1.Row, 16) = Form1.ws1.Text
            'Form1.ws1.Text = Val(Form1.ws1.Text) + 1
            Grid1.TextMatrix(Grid1.Row, 16) = wd_seq("TLW1Barcode")
            Form1.ws1.Text = Grid1.TextMatrix(Grid1.Row, 16)
        End If
        If Grid1.TextMatrix(Grid1.Row, 3) = "TRI-LEVEL 2" Then
            'Form1.ws2.Text = Val(Form1.ws2.Text) + 1
            'Grid1.TextMatrix(Grid1.Row, 16) = Form1.ws2.Text
            'Form1.ws2.Text = Val(Form1.ws2.Text) + 1
            Grid1.TextMatrix(Grid1.Row, 16) = wd_seq("TLW2Barcode")
            Form1.ws2.Text = Grid1.TextMatrix(Grid1.Row, 16)
        End If
        
        Call post_dai_exp_rcpt
        'saerequests.rowkey = Grid1.Row
        'saerequests.barkey = Grid1.TextMatrix(Grid1.Row, 6)
        If Grid1.TextMatrix(Grid1.Row, 3) = "TRI-LEVEL 1" Then bc1 = Grid1.TextMatrix(Grid1.Row, 6)
        If Grid1.TextMatrix(Grid1.Row, 3) = "TRI-LEVEL 2" Then bc2 = Grid1.TextMatrix(Grid1.Row, 6)
        List1.AddItem Right(Grid1.TextMatrix(Grid1.Row, 3), 1) & "  " & Grid2.TextMatrix(1, 0) & "  " & Grid1.TextMatrix(Grid1.Row, 6) & " " & Grid1.TextMatrix(Grid1.Row, 16)
        Call post_route_to_kep(Right(Grid1.TextMatrix(Grid1.Row, 3), 1), Grid2.TextMatrix(1, 0), Grid1.TextMatrix(Grid1.Row, 16))
        Call post_tm_log(Grid2.TextMatrix(1, 0), Grid1.Row, Grid1.TextMatrix(Grid1.Row, 16))
        Call record_pallet(Grid1.Row)
        If Val(Grid2.TextMatrix(1, 0)) > 0 And Val(Grid2.TextMatrix(1, 0)) < 4 Then
            Call post_queue_to_sr(Val(Grid2.TextMatrix(1, 0)), Grid1.Row)
        End If
    End If
End Sub

Private Sub refmenu_Click()
    refresh_grid1
End Sub

Private Sub ttrig_Change()
    refresh_grid1
    DoEvents
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 3) = "TRI-LEVEL 1" Then
                If Grid1.TextMatrix(i, 6) <> bc1 Then
                    Grid1.Row = i
                    'Grid1.TextMatrix(i, 16) = Form1.ws1.Text
                    'Form1.ws1.Text = Val(Form1.ws1.Text) + 1
                    Grid1_DblClick
                    DoEvents
                End If
            End If
            If Grid1.TextMatrix(i, 3) = "TRI-LEVEL 2" Then
                If Grid1.TextMatrix(i, 6) <> bc2 Then
                    Grid1.Row = i
                    'Grid1.TextMatrix(i, 16) = Form1.ws2.Text
                    'Form1.ws2.Text = Val(Form1.ws2.Text) + 1
                    Grid1_DblClick
                    DoEvents
                End If
            End If
        Next i
        'Grid1.Row = 1
        'Grid1_DblClick
    End If
End Sub

