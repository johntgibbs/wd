VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form r12drytickets 
   Caption         =   "Dry Goods Tickets"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8820
   LinkTopic       =   "Form2"
   ScaleHeight     =   8610
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   1935
      Left            =   0
      TabIndex        =   4
      Top             =   6600
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3413
      _Version        =   327680
      BackColorFixed  =   12648384
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   4080
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4048
      _Version        =   327680
      BackColorFixed  =   12632319
      BackColorSel    =   255
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      _Version        =   327680
      BackColorFixed  =   12648447
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "R12 Material Transactions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   6360
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Posted to RFGen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Branch Orders:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label seqkey 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "r12drytickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.oradb
    db.Open "odbc;database=pbelle;uid=Apps;pwd=pb3113tx;dsn=pbelle"
    s = "select shipment_number, source_warehouse, receiving_warehouse, order_date, count(*)"
    s = s & " from bbc_bimp_shipments"
    s = s & " where trunc(order_date) > trunc(sysdate - 10)"
    If Form1.Combo1 = "501" Then s = s & " and source_warehouse = 'K01'"
    If Form1.Combo1 = "502" Then s = s & " and source_warehouse = 'A01'"
    s = s & " group by shipment_number, source_warehouse, receiving_warehouse, order_date"
    s = s & " order by shipment_number"
    'If Option1 = True Then s = s & " where upload_flag = 2"
    'If Option2 = True Then s = s & " where upload_flag = 0"
    'If Option3 = True Then s = s & " where upload_flag = 1"
    'If Option4 = True Then s = s & " and type = 'MIXER'"
    'If Option5 = True Then s = s & " and type = 'FG'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!shipment_number & Chr(9)
            s = s & ds!source_warehouse & Chr(9)
            s = s & ds!receiving_warehouse & Chr(9)
            s = s & ds!order_date
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    s = "^Shipment_number|^Source_warehouse|^Receiving_warehouse|^Order_date"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 2000
    Grid1.ColWidth(1) = 2000
    Grid1.ColWidth(2) = 2000
    Grid1.ColWidth(3) = 1500
    Call Grid1_RowColChange
End Sub

Private Sub refresh_grid2()
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 14
    Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.oradb
    db.Open "odbc;database=pbelle;uid=Apps;pwd=pb3113tx;dsn=pbelle"
    s = "select * from bbc_bimp_shipments where shipment_number = '" & seqkey.Caption & "'"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!source_warehouse & Chr(9)
            s = s & ds!receiving_warehouse & Chr(9)
            s = s & ds!inventory_item_id & Chr(9)
            s = s & ds!item_number & Chr(9)
            s = s & ds!Description & Chr(9)
            s = s & ds!order_date & Chr(9)
            s = s & ds!qty & Chr(9)
            s = s & ds!uom & Chr(9)
            s = s & ds!shipment_number & Chr(9)
            s = s & ds!line_number & Chr(9)
            s = s & ds!attribute1 & Chr(9)
            s = s & ds!attribute2 & Chr(9)
            s = s & ds!attribute3 & Chr(9)
            s = s & ds!attribute4
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    s = "<Source|<Recv|<ItemId|<ItemNo|<Description|<Date|<Qty|<Uom|<ShipNum|<Line|<Attr1|<Attr2|<Attr3|<Attr4"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 2000
    Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 1000
    Grid2.ColWidth(7) = 1000
    Grid2.ColWidth(8) = 1000
    Grid2.ColWidth(9) = 1000
    Grid2.ColWidth(10) = 1000
    Grid2.ColWidth(11) = 1000
    Grid2.ColWidth(12) = 1000
    Grid2.ColWidth(13) = 1000
    If Grid2.Rows > 1 Then Call refresh_grid3(Grid2.TextMatrix(1, 8))
End Sub

Private Sub refresh_grid3(rid As String)
    Dim q As String, i As Integer, k As Integer
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
    
    q = "select t.subinventory_code, i.segment1, i.description, t.transaction_quantity," & _
        " t.transaction_uom, t.source_code, t.transaction_reference, t.transaction_date" & _
        " from mtl_material_transactions t, mtl_system_items_b i" & _
        " where t.shipment_number in ('" & rid & "', 'S" & rid & "')" & _
        " and i.inventory_item_id = t.inventory_item_id" & _
        " and i.organization_id = t.organization_id" & _
        " order by t.source_code, i.segment1, t.subinventory_code"
        
    'MsgBox q
    i = LoadGrid(Grid3, q, hstmt, 1, "")
    i = DisconnectFromDataSource(hdbc, hstmt)
    i = FreeODBChEnv(hEnv)
    Screen.MousePointer = 0
    Grid3.FormatString = "^SubInv|^SKU|<Product|^Qty|^UOM|^SourceCode|<Reference|^Date/Time"
    Grid3.ColWidth(0) = 800
    Grid3.ColWidth(1) = 800
    Grid3.ColWidth(2) = 2000
    Grid3.ColWidth(3) = 800
    Grid3.ColWidth(4) = 600
    Grid3.ColWidth(5) = 1800
    Grid3.ColWidth(6) = 2600
    Grid3.ColWidth(7) = 1800
    If Grid3.Rows > 2 Then Grid3.RemoveItem Grid3.Rows - 1
End Sub

Private Sub Command1_Click()
    refresh_grid1
End Sub

Private Sub Form_Load()
    refresh_grid1
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    Grid2.Width = Me.Width - 100
    Grid3.Width = Me.Width - 100
End Sub

Private Sub Grid1_RowColChange()
    seqkey.Caption = Grid1.TextMatrix(Grid1.Row, 0)
End Sub

Private Sub seqkey_Change()
    refresh_grid2
End Sub

