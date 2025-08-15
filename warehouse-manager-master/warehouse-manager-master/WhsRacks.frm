VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Rack Inventory"
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   1425
   ClientWidth     =   10965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   7920
   ScaleWidth      =   10965
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9600
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid5 
      Height          =   975
      Left            =   0
      TabIndex        =   6
      Top             =   7560
      Visible         =   0   'False
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   1720
      _Version        =   327680
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2990
      _Version        =   327680
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "F2:Edit Rack"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "^N:On Hold"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "^F:First Out"
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
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox PAisle 
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid RGrid 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8916
      _Version        =   327680
      Cols            =   13
      BackColorFixed  =   12648447
      BackColorSel    =   128
      FocusRect       =   2
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label bcolor 
      BackColor       =   &H00FFFFC0&
      Caption         =   "bcolor"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   6840
      Width           =   855
   End
   Begin VB.Menu qmenu 
      Caption         =   "&Query"
      Begin VB.Menu qs 
         Caption         =   "Selected Aisle"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu qs 
         Caption         =   "Rack For SKU"
         Index           =   1
      End
      Begin VB.Menu qs 
         Caption         =   "Racks On Hold"
         Index           =   2
      End
      Begin VB.Menu qs 
         Caption         =   "1st Out Racks"
         Index           =   3
      End
      Begin VB.Menu qs 
         Caption         =   "Reserved Racks"
         Index           =   4
      End
      Begin VB.Menu qs 
         Caption         =   "4Way Pallets"
         Index           =   5
      End
      Begin VB.Menu qs 
         Caption         =   "Empty Slots"
         Index           =   6
      End
      Begin VB.Menu qs 
         Caption         =   "BarCode"
         Index           =   7
      End
   End
   Begin VB.Menu prtmenu 
      Caption         =   "&Print"
      Begin VB.Menu cntsheet 
         Caption         =   "Count Sheet Format"
      End
      Begin VB.Menu bbpalship 
         Caption         =   "Shipping Pallet Sheet"
      End
      Begin VB.Menu bbpalprt 
         Caption         =   "BB Pallets Format"
      End
      Begin VB.Menu barclist 
         Caption         =   "BarCode Listing"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub syl_floor_pallets()
    Dim s As String
    Dim rid As Long, pid As Long
    Dim i As Integer
    Screen.MousePointer = 11
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'R', 'FLOOR', 1, 100, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Wdb.Execute s
    For i = 1 To 100
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '2-1-2014', 'Y', ' ', ' ', 0, 'Y', 'N')"
        db.Execute s
    Next i
    Screen.MousePointer = 0
End Sub

Sub t10_a_d_aisles()
    Dim s As String
    Dim rid As Long, pid As Long
    Dim i As Integer, k As Integer
    Dim cfile As String
    cfile = "U:\jvwork.txt"
    Open cfile For Output As #1
    Screen.MousePointer = 11
    
    'Aisle A
    s = "Delete from rackpos where rackno in (select id from racks where aisle = 'A')"
    Print #1, s
    Wdb.Execute s
    s = "Delete from racks where aisle = 'A'"
    Print #1, s
    Wdb.Execute s
    'A1 - A15
    For k = 1 To 15
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'A', '" & k & "', " & k & ", 16, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 16
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'A16 - A30
    For k = 16 To 30
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'A', '" & k & "', " & k & ", 12, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 12
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    
    
    'Aisle B
    s = "Delete from rackpos where rackno in (select id from racks where aisle = 'B')"
    Print #1, s
    Wdb.Execute s
    s = "Delete from racks where aisle = 'B'"
    Print #1, s
    Wdb.Execute s
    'B1 - B4
    For k = 1 To 4
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'B', '" & k & "', " & k & ", 16, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 16
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'B5
    k = 5
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'B', '" & k & "', " & k & ", 13, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 13
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    'B6 - B12
    For k = 6 To 12
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'B', '" & k & "', " & k & ", 16, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 16
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'B13
    k = 13
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'B', '" & k & "', " & k & ", 13, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 13
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    'B14 - B15
    For k = 14 To 15
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'B', '" & k & "', " & k & ", 16, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 16
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'B16 - B17
    For k = 16 To 17
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'B', '" & k & "', " & k & ", 12, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 12
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'B18
    k = 18
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'B', '" & k & "', " & k & ", 8, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 8
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    'B19 - B25
    For k = 19 To 25
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'B', '" & k & "', " & k & ", 12, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 12
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'B26
    k = 26
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'B', '" & k & "', " & k & ", 8, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 8
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    'B27 - B30
    For k = 27 To 30
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'B', '" & k & "', " & k & ", 12, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 12
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    
    'Aisle C
    s = "Delete from rackpos where rackno in (select id from racks where aisle = 'C')"
    Print #1, s
    Wdb.Execute s
    s = "Delete from racks where aisle = 'C'"
    Print #1, s
    Wdb.Execute s
    'C1 - C4
    For k = 1 To 4
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'C', '" & k & "', " & k & ", 20, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 20
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'C5
    k = 5
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'C', '" & k & "', " & k & ", 8, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 8
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    'C6 - C12
    For k = 6 To 12
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'C', '" & k & "', " & k & ", 20, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 20
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'C13
    k = 13
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'C', '" & k & "', " & k & ", 8, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 8
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    'C14 - C15
    For k = 14 To 15
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'C', '" & k & "', " & k & ", 20, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 20
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'C16 - C30
    For k = 16 To 30
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'C', '" & k & "', " & k & ", 16, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 16
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    
    'Aisle D
    s = "Delete from rackpos where rackno in (select id from racks where aisle = 'D')"
    Print #1, s
    Wdb.Execute s
    s = "Delete from racks where aisle = 'D'"
    Print #1, s
    Wdb.Execute s
    'D1 - D4
    For k = 1 To 4
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'D', '" & k & "', " & k & ", 16, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 16
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'D5
    k = 5
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'D', '" & k & "', " & k & ", 20, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 20
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    'D6 - D12
    For k = 6 To 12
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'D', '" & k & "', " & k & ", 16, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 16
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'D13
    k = 13
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'D', '" & k & "', " & k & ", 20, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 20
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    'D14 - D15
    For k = 14 To 15
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'D', '" & k & "', " & k & ", 16, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 16
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'D16 - D27
    For k = 16 To 27
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'D', '" & k & "', " & k & ", 20, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 20
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'D28 - D31
    For k = 28 To 31
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'D', '" & k & "', " & k & ", 15, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 15
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'D32 - D41
    For k = 32 To 41
        rid = wd_seq("Racks")
        s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
        s = s & ", resv_lot, fo, hold) Values (" & rid
        s = s & ", '1', 'D', '" & k & "', " & k & ", 20, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
        Print #1, s
        Wdb.Execute s
        For i = 1 To 20
            pid = wd_seq("RackPos")
            s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
            s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
            s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
            Print #1, s
            Wdb.Execute s
        Next i
    Next k
    'D42
    k = 42
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'D', '" & k & "', " & k & ", 15, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 15
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    'D43
    k = 43
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'D', '" & k & "', " & k & ", 10, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 10
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    'D-Wall
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'D', 'WALL', 0, 26, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 26
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    
    'M-Walls
    'WALL-N
    s = "Delete from rackpos where rackno in (select id from racks where aisle = 'M' and rack = 'WALL-N')"
    Print #1, s
    Wdb.Execute s
    s = "Delete from racks where aisle = 'M' and rack = 'WALL-N'"
    Print #1, s
    Wdb.Execute s
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'M', 'WALL-N', 0, 36, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 36
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    'WALL-S
    s = "Delete from rackpos where rackno in (select id from racks where aisle = 'M' and rack = 'WALL-S')"
    Print #1, s
    Wdb.Execute s
    s = "Delete from racks where aisle = 'M' and rack = 'WALL-S'"
    Print #1, s
    Wdb.Execute s
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'M', 'WALL-S', 0, 18, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Print #1, s
    Wdb.Execute s
    For i = 1 To 18
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '5-23-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Print #1, s
        Wdb.Execute s
    Next i
    
    Close #1
    Screen.MousePointer = 0
End Sub

Sub bill_aisles()
    Dim s As String
    Dim rid As Long, pid As Long
    Dim i As Integer
    Screen.MousePointer = 11
    'Clear Old Racks
    s = "Delete from rackpos where rackno in (select id from racks where aisle = '1' and rack = '1')"
    Wdb.Execute s
    s = "Delete from racks where aisle = '1' and rack = '1'"
    Wdb.Execute s
    'E-17
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'E', '17', 17, 4, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Wdb.Execute s
    For i = 1 To 4
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '9-16-2015', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Wdb.Execute s
    Next i
    Screen.MousePointer = 0
End Sub

Sub bae18e19()
    Dim s As String, ds As ADODB.Recordset
    Dim rid As Long, pid As Long
    Screen.MousePointer = 11
    'E-18
    s = "select * from racks where aisle = 'E' and rack = '18'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        rid = ds!id
        s = "Update racks set capacity = 5 where id = " & rid
        Wdb.Execute s
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", 5, ' ', ' ', ' ', 0, '8-24-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Wdb.Execute s
    End If
    ds.Close
    'E-19
    s = "select * from racks where aisle = 'E' and rack = '19'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        rid = ds!id
        s = "Update racks set capacity = 4 where id = " & rid
        Wdb.Execute s
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", 4, ' ', ' ', ' ', 0, '8-24-2016', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Wdb.Execute s
    End If
    ds.Close
    Screen.MousePointer = 0
End Sub

Sub ba_Easisle_racks()
    Dim s As String
    Dim rid As Long, pid As Long
    Dim i As Integer
    Screen.MousePointer = 11
    'E-12
    s = "Update racks set capacity = 2 where aisle = 'E' and rack = '12'"
    Wdb.Execute s
    s = "Delete from rackpos where rackno in (select id from racks where aisle = 'E' and rack = '12')"
    s = s & " and posn_num in (3, 4)"
    Wdb.Execute s
    'E-17
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'E', '17', 17, 4, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Wdb.Execute s
    For i = 1 To 4
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '9-16-2015', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Wdb.Execute s
    Next i
    'E-18
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'E', '18', 18, 4, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Wdb.Execute s
    For i = 1 To 4
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '9-16-2015', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Wdb.Execute s
    Next i
    'E-19
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'E', '19', 19, 3, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Wdb.Execute s
    For i = 1 To 3
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '9-16-2015', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Wdb.Execute s
    Next i
    'E-20
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'E', '20', 20, 4, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Wdb.Execute s
    For i = 1 To 4
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '9-16-2015', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Wdb.Execute s
    Next i
    'E-21
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'E', '21', 21, 2, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Wdb.Execute s
    For i = 1 To 2
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '9-16-2015', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Wdb.Execute s
    Next i
    'E-22
    rid = wd_seq("Racks")
    s = "Insert into racks (id, room, aisle, rack, slot, capacity, sku, lot_num, qty, qty4, resv_sku"
    s = s & ", resv_lot, fo, hold) Values (" & rid
    s = s & ", '1', 'E', '22', 22, 2, ' ', ' ', 0, 0, ' ', ' ', 0, 0)"
    Wdb.Execute s
    For i = 1 To 2
        pid = wd_seq("RackPos")
        s = "Insert into rackpos (id, rackno, posn_num, sku, lot_num, pallet_num, count_qty, recv_date,"
        s = s & " bbc, barcode, lot2, qty2, wrapped, hold) Values (" & pid
        s = s & ", " & rid & ", " & i & ", ' ', ' ', ' ', 0, '9-16-2015', 'Y', ' ', ' ', 0, 'Y', 'N')"
        Wdb.Execute s
    Next i
    
    Screen.MousePointer = 0
End Sub

Private Function calc_date(lotcode As String) As String
    Dim seed As String
    If lotcode > "0" Then
        If Left(lotcode, 2) = "00" Then
            seed = "12-31-1999"
        Else
            If Val(lotcode) > 90000 Then
                seed = "12-31-19" & Val(Left(lotcode, 2)) - 1
            Else
                seed = "12-31-20" & Format(Val(Left(lotcode, 2)) - 1, "00")
            End If
        End If
        calc_date = Format(DateAdd("d", Val(Right(lotcode, 3)), seed), "m-d-yyyy")
    Else
        calc_date = Now
    End If
End Function

Private Sub outgoing_pallets()
    Dim ds As ADODB.Recordset, sqlx As String, rs As ADODB.Recordset
    Dim i As Integer, psku As String, s As String
    Dim t As String, tbold As Boolean, trow As Integer
    Dim ordqty As Integer, rkqty As Integer
    trow = RGrid.TopRow
    
    Screen.MousePointer = 11
    RGrid.Visible = False: RGrid.Clear
    RGrid.Rows = 1: RGrid.Cols = 13
    RGrid.FixedCols = 2
    
    s = "select product,count(*) from paltasks where area = 'FORKLIFT' and target = 'STAGING'"
    s = s & " and description > '  ' and status = 'PEND'"
    s = s & " group by product order by product"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            psku = Trim(Left(ds(0), 4))                             'jv082415
            ordqty = ds(1) + 2
            rkqty = 0
            sqlx = "select r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
            'sqlx = "select r.aisle,r.rack,p.sku,r.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
            sqlx = sqlx & " from racks r, rackpos p"
            sqlx = sqlx & " where p.sku = '" & psku & "' and r.aisle <> 'M' and p.bbc = 'Y'"
            sqlx = sqlx & " and r.hold = 0"
            sqlx = sqlx & " and r.resv_sku <> 'ALL'"
            sqlx = sqlx & " and p.rackno = r.id"
            sqlx = sqlx & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
            'sqlx = sqlx & " group by r.aisle,r.rack,p.sku,r.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
            sqlx = sqlx & " order by r.fo desc, p.lot_num, r.aisle, r.slot"
            'sqlx = sqlx & " order by r.fo desc, r.lot_num"
            Set rs = Wdb.Execute(sqlx)
            If rs.BOF = False Then
                rs.MoveFirst
                Do Until rs.EOF
                    If ordqty >= rkqty Then
                        s = " " & Chr(9)                              'jv062111
                        s = s & rs(0) & "-" & Trim(rs(1)) & Chr(9)    'jv062111
                        's = s & " " & Chr(9)                          'jv062111
                        s = s & ordqty & Chr(9)                          'jv062111
                        s = s & rs(2) & Chr(9)                        'jv062111
                        s = s & rs(3) & Chr(9)                        'jv062111
                        s = s & rs(10) & Chr(9)                       'jv062111
                        If rs(8) = "N" Then s = s & "+"               'jv062111
                        s = s & Chr(9)                                'jv062111
                        s = s & rs(4) & Chr(9)                        'jv062111
                        s = s & rs(5) & Chr(9)                        'jv062111
                        s = s & rs(6) & Chr(9)                        'jv062111
                        s = s & rs(7) & Chr(9)                        'jv062111
                        s = s & Chr(9)                                'jv062111
                        s = s & rs(9)                                 'jv062111
                        RGrid.AddItem s
                        rkqty = rkqty + rs(10)
                    End If
                    rs.MoveNext
                Loop
            End If
            rs.Close
            ds.MoveNext
        Loop
    End If
    RGrid.FillStyle = flexFillRepeat                    'jv062111
    t = RGrid.TextMatrix(0, 1)
    For i = 1 To RGrid.Rows - 1
        RGrid.Row = i: RGrid.Col = 3: psku = " "
        If Val(RGrid.Text) > 0 Then
            psku = Trim(RGrid.Text)
        Else
            RGrid.Col = 7
            If Val(RGrid.Text) > 0 Then psku = Trim(RGrid.Text)
        End If
        If Val(psku) > 0 Then
            If skurec(Val(psku)).sku <> psku Then
                RGrid.TextMatrix(i, 11) = "Invalid SKU"
            Else
                RGrid.TextMatrix(i, 11) = " " & skurec(Val(psku)).prodname
            End If
        End If
        If t <> RGrid.TextMatrix(i, 1) Then
            tbold = Not tbold
            t = RGrid.TextMatrix(i, 1)
        End If
        RGrid.Row = i: RGrid.RowSel = i
        RGrid.Col = 1: RGrid.ColSel = 2 'RGrid.Cols - 1
        RGrid.CellFontBold = tbold
        
        If RGrid.TextMatrix(i, 3) > "0" Then                'jv062111
            RGrid.Row = i: RGrid.RowSel = i                 'jv062111
            RGrid.Col = 1: RGrid.ColSel = RGrid.Cols - 1    'jv062111
            If tbold = False Then
            RGrid.CellBackColor = RGrid.BackColorFixed      'jv062111
            Else
                RGrid.CellBackColor = bcolor.BackColor
            End If
        End If                                              'jv062111
    Next i
    
    RGrid.FormatString = "#|^Rack|^Cap|^SKU|^Lot|^Qty|^4W|^Resv|^Lot|^FO|^Hold|<Description|^Slot"
    RGrid.ColWidth(0) = 1
    RGrid.ColWidth(1) = 1150: RGrid.ColWidth(2) = 500
    RGrid.ColWidth(3) = 600: RGrid.ColWidth(4) = 850
    RGrid.ColWidth(5) = 500: RGrid.ColWidth(6) = 500
    RGrid.ColWidth(7) = 600: RGrid.ColWidth(8) = 700
    RGrid.ColWidth(9) = 400: RGrid.ColWidth(10) = 550
    RGrid.ColWidth(11) = 3500: RGrid.ColWidth(12) = 500
    Screen.MousePointer = 0
    RGrid.Visible = True
    If lr <> 1 Then
        RGrid.TopRow = trow
    End If
    If RGrid.Rows > 1 Then
        RGrid.Row = lr
        Call RGrid_Click
    End If
End Sub

Sub Refresh_racks(ByRef lr As Integer)
    Dim ds As ADODB.Recordset, sqlx As String
    Dim i As Integer, psku As String
    Dim t As String, tbold As Boolean, trow As Integer
    trow = RGrid.TopRow
    psku = Trim(RGrid.TextMatrix(RGrid.Row, 3))
    If Len(psku) = 0 Then psku = Trim(RGrid.TextMatrix(RGrid.Row, 7))
    If Len(psku) = 0 Then psku = "507"
    If qs(0).Checked = True Then                                'jv062111 changed all queries
        Form4.Caption = "Rack Inventory - Aisle " & PAisle
        If PAisle = "M" Then
            sqlx = "select aisle,rack,sku,lot_num,resv_sku,resv_lot,fo,hold,id,slot,qty"
            sqlx = sqlx & " from racks where aisle = 'M' order by slot"
        Else
            sqlx = "select r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
            sqlx = sqlx & " from racks r, rackpos p"
            sqlx = sqlx & " where r.aisle = '" & PAisle & "' and p.rackno = r.id"
            sqlx = sqlx & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
            sqlx = sqlx & " order by r.slot,p.sku desc"
        End If
    End If
    If qs(1).Checked = True Then
        psku = InputBox("SKU:", "Which SKU", psku)
        If Len(psku) = 0 Then Exit Sub
        Form4.Caption = "Rack Inventory - SKU " & psku
        sqlx = "select r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
        sqlx = sqlx & " from racks r, rackpos p"
        sqlx = sqlx & " where p.sku = '" & psku & "' and p.rackno = r.id"
        sqlx = sqlx & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        sqlx = sqlx & " order by r.aisle,r.slot,p.lot_num"
    End If
    If qs(2).Checked = True Then
        Form4.Caption = "Rack Inventory - On Hold"
        sqlx = "select r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
        sqlx = sqlx & " from racks r, rackpos p"
        sqlx = sqlx & " where r.hold = 1 and p.rackno = r.id"
        sqlx = sqlx & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        sqlx = sqlx & " order by r.aisle, r.slot, p.sku desc"
    End If
    If qs(3).Checked = True Then
        Form4.Caption = "Rack Inventory - 1st Out"
        sqlx = "select r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
        sqlx = sqlx & " from racks r, rackpos p"
        sqlx = sqlx & " where r.fo = 1 and p.rackno = r.id"
        sqlx = sqlx & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        sqlx = sqlx & " order by r.aisle, r.slot, p.sku desc"
    End If
    If qs(4).Checked = True Then
        Form4.Caption = "Rack Inventory - Reserved Racks"
        sqlx = "select r.aisle,r.rack,r.resv_sku,r.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
        sqlx = sqlx & " from racks r, rackpos p"
        sqlx = sqlx & " where r.resv_sku > ' ' and p.rackno = r.id"
        sqlx = sqlx & " group by r.aisle,r.rack,r.resv_sku,r.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        sqlx = sqlx & " order by r.aisle, r.slot, r.resv_sku desc"
    End If
    If qs(5).Checked = True Then
        Form4.Caption = "Rack Inventory - 4 Way Pallets"
        sqlx = "select r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
        sqlx = sqlx & " from racks r, rackpos p"
        sqlx = sqlx & " where p.bbc = 'N' and p.rackno = r.id"
        sqlx = sqlx & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        sqlx = sqlx & " order by r.aisle, r.slot, p.sku desc"
    End If
    If qs(6).Checked = True Then
        Form4.Caption = "Rack Inventory - Empty Slots"
        sqlx = "select r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
        sqlx = sqlx & " from racks r, rackpos p"
        sqlx = sqlx & " where r.qty = 0 and r.qty4 = 0 and p.rackno = r.id"
        sqlx = sqlx & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        sqlx = sqlx & " order by r.aisle, r.slot, p.sku desc"
    End If
    If qs(7).Checked = True Then
        psku = InputBox("BarCode:", "Which BarCode", psku)
        If Len(psku) = 0 Then Exit Sub
        If Len(psku) < 16 Then
            MsgBox "Insuffecient barcode info..", vbOKOnly + vbInformation, "Not a barcode..."
            Exit Sub
        End If
        psku = UCase(psku)
        Form4.Caption = "Rack Inventory - BarCode " & psku
        sqlx = "select r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
        sqlx = sqlx & " from racks r, rackpos p"
        sqlx = sqlx & " where p.barcode in ('" & psku & "',"
        psku = Left(psku, 11) & "_" & Right(psku, 4)
        sqlx = sqlx & "'" & psku & "') and p.rackno = r.id"
        sqlx = sqlx & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        sqlx = sqlx & " order by r.aisle, r.slot, p.sku desc"
    End If
    Screen.MousePointer = 11
    RGrid.Redraw = False
    RGrid.FontName = "Arial"
    RGrid.FontBold = True
    RGrid.FontSize = 8
    RGrid.Clear: RGrid.Rows = 1: RGrid.Cols = 13
    RGrid.FixedCols = 2
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = " " & Chr(9)                                 'jv062111
            sqlx = sqlx & ds(0) & "-" & Trim(ds(1)) & Chr(9)    'jv062111
            sqlx = sqlx & " " & Chr(9)                          'jv062111
            sqlx = sqlx & ds(2) & Chr(9)                        'jv062111
            sqlx = sqlx & ds(3) & Chr(9)                        'jv062111
            sqlx = sqlx & ds(10) & Chr(9)                       'jv062111
            If ds(8) = "N" Then sqlx = sqlx & "+"               'jv062111
            sqlx = sqlx & Chr(9)                                'jv062111
            sqlx = sqlx & ds(4) & Chr(9)                        'jv062111
            sqlx = sqlx & ds(5) & Chr(9)                        'jv062111
            sqlx = sqlx & ds(6) & Chr(9)                        'jv062111
            sqlx = sqlx & ds(7) & Chr(9)                        'jv062111
            sqlx = sqlx & Chr(9)                                'jv062111
            sqlx = sqlx & ds(9)                                 'jv062111
            RGrid.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    RGrid.FillStyle = flexFillRepeat                    'jv062111
    t = RGrid.TextMatrix(0, 1)
    For i = 1 To RGrid.Rows - 1
        RGrid.Row = i: RGrid.Col = 3: psku = " "
        If Val(RGrid.Text) > 0 Then
            psku = Trim(RGrid.Text)
        Else
            RGrid.Col = 7
            If Val(RGrid.Text) > 0 Then psku = Trim(RGrid.Text)
        End If
        If Val(psku) > 0 Then
            If skurec(Val(psku)).sku <> psku Then
                RGrid.TextMatrix(i, 11) = "Invalid SKU"
            Else
                RGrid.TextMatrix(i, 11) = " " & skurec(Val(psku)).prodname
            End If
        End If
        If t <> RGrid.TextMatrix(i, 1) Then
            tbold = Not tbold
            t = RGrid.TextMatrix(i, 1)
        End If
        RGrid.Row = i: RGrid.RowSel = i
        RGrid.Col = 1: RGrid.ColSel = 2
        RGrid.CellFontBold = tbold
        
        If RGrid.TextMatrix(i, 3) > "0" Then                'jv062111
            RGrid.Row = i: RGrid.RowSel = i                 'jv062111
            RGrid.Col = 1: RGrid.ColSel = RGrid.Cols - 1    'jv062111
            If tbold = False Then
            RGrid.CellBackColor = RGrid.BackColorFixed      'jv062111
            Else
                RGrid.CellBackColor = bcolor.BackColor
            End If
        End If                                              'jv062111
    Next i
    RGrid.FormatString = "#|^Rack|^Cap|^SKU|^Lot|^Qty|^4W|^Resv|^Lot|^FO|^Hold|<Description|^Slot"
    RGrid.ColWidth(0) = 1
    RGrid.ColWidth(1) = 1150: RGrid.ColWidth(2) = 500
    RGrid.ColWidth(3) = 600: RGrid.ColWidth(4) = 850
    RGrid.ColWidth(5) = 500: RGrid.ColWidth(6) = 500
    RGrid.ColWidth(7) = 600: RGrid.ColWidth(8) = 700
    RGrid.ColWidth(9) = 400: RGrid.ColWidth(10) = 550
    RGrid.ColWidth(11) = 3500: RGrid.ColWidth(12) = 500
    Screen.MousePointer = 0
    RGrid.Redraw = True
    If lr <> 1 Then
        RGrid.TopRow = trow
    End If
    If RGrid.Rows > 1 Then
        RGrid.Row = lr
        Call RGrid_Click
    End If
End Sub

Private Sub barclist_Click()
    Dim ds As ADODB.Recordset, s As String, rs As ADODB.Recordset, i, mprod As String
    Dim mrack As String, msku As String, rt As String, rf As String, rh As String
    mrack = RGrid.TextMatrix(RGrid.Row, 1)
    msku = RGrid.TextMatrix(RGrid.Row, 3)
    mrack = InputBox("Rack # or All", "Specify Rack or All Racks....", mrack)
    If Len(mrack) = 0 Then Exit Sub
    msku = InputBox("SKU # or All", "Specify SKU or All SKUs....", msku)
    If Len(msku) = 0 Then Exit Sub
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 10
    s = "select * from rackpos where barcode > ' '"
    If LCase(msku) <> "all" Then s = s & " and sku = '" & msku & "'"
    If LCase(mrack) <> "all" Then
        s = s & " and rackno in (select id from racks where aisle = '"
        s = s & Left(mrack, 1) & "' and rack = '" & Right(mrack, Len(mrack) - 2) & "')"
    End If
    s = s & " order by barcode"
    Screen.MousePointer = 11
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If LCase(msku) <> "all" Then
            If skurec(Val(msku)).sku = msku Then
                mprod = StrConv(skurec(Val(msku)).prodname, vbProperCase)
            Else
                mprod = "Undefined SKU"
            End If
        End If
        Do Until ds.EOF
            If LCase(msku) = "all" Then
                If skurec(Val(ds!sku)).sku = ds!sku Then
                    mprod = StrConv(skurec(Val(ds!sku)).prodname, vbProperCase)
                Else
                    mprod = "Undefined SKU"
                End If
            End If
            If LCase(mrack) = "all" Then
                s = "select aisle, rack from racks where id = " & ds!rackno
                Set rs = Wdb.Execute(s)
                If rs.BOF = False Then
                    rs.MoveFirst
                    s = rs!aisle & "-" & rs!rack
                Else
                    s = "???"
                End If
                rs.Close
            Else
                s = mrack
            End If
            s = s & Chr(9) & ds!sku & Chr(9)
            s = s & mprod & Chr(9)
            s = s & ds!barcode & Chr(9)
            s = s & ds!bbc & Chr(9)
            s = s & ds!lot_num & Chr(9)
            s = s & ds!count_qty & Chr(9)
            s = s & ds!lot2 & Chr(9)
            s = s & ds!qty2 & Chr(9)
            s = s & Format(ds!recv_date, "m-d-yyyy")
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^Location|^SKU|<Prouct|^BarCode|^BB|^LotNum|^Units|^Lot2|^Qty|^Recv Date"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 700
    Grid1.ColWidth(2) = 2600
    Grid1.ColWidth(3) = 1600
    Grid1.ColWidth(4) = 600
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 1200
    Screen.MousePointer = 0
    
    rt = "Barcode Listing"
    rh = "SKU " & UCase(msku) & "  Rack " & UCase(mrack) & "  " & Grid1.Rows - 1 & " Pallets"
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
    
End Sub

Private Sub bbpalprt_Click()
    If qs(0).Checked = True Then
        'Form12.qstr = "where aisle = '" & PAisle & "'"
        'Form12.qsort = "order by slot"
        Form12.qstr = "where r.aisle = '" & PAisle & "' and p.sku > ' '"
        Form12.qsort = "order by r.slot"
    End If
    If qs(1).Checked = True Then
        'Form12.qstr = "where sku = '" & Right$(Form4.Caption, 3) & "'"
        'Form12.qsort = "order by lot_num,aisle,slot"
        Form12.qstr = "where p.sku = '" & Trim(Right$(Form4.Caption, 4)) & "'"
        'Form12.qsort = "order by p.lot_num,r.aisle,r.slot"
        Form12.qsort = "order by r.aisle,r.slot,p.lot_num"
    End If
    If qs(2).Checked = True Then
        'Form12.qstr = "where hold = 1"
        'Form12.qsort = "order by aisle,slot"
        Form12.qstr = "where r.hold = 1 and p.sku > ' '"
        Form12.qsort = "order by r.aisle,r.slot"
    End If
    If qs(3).Checked = True Then
        'Form12.qstr = "where fo = 1"
        'Form12.qsort = "order by aisle,slot"
        Form12.qstr = "where r.fo = 1 and p.sku > ' '"
        Form12.qsort = "order by r.aisle,r.slot"
    End If
    If qs(4).Checked = True Then
        'Form12.qstr = "where resv_sku > '0'"
        'Form12.qsort = "order by aisle,slot"
        Form12.qstr = "where r.resv_sku > '0'"
        Form12.qsort = "order by r.aisle,r.slot"
    End If
    If qs(5).Checked = True Then
        'Form12.qstr = "where qty4 > 0"
        'Form12.qsort = "order by aisle,slot"
        Form12.qstr = "where p.bbc = 'N'"
        Form12.qsort = "order by r.aisle,r.slot"
    End If
    If qs(6).Checked = True Then
        'Form12.qstr = "where qty = 0 and qty4 = 0"
        'Form12.qsort = "order by aisle,slot"
        'Form12.qstr = "where p.count_qty = 0"
        Form12.qstr = "where r.qty = 0 and r.qty4 = 0"
        Form12.qsort = "order by r.aisle,r.slot"
    End If
    Form12.Show
End Sub

Private Sub bbpalship_Click()
    Call outgoing_pallets
    DoEvents
    Form12.qstr = "outgoing"
    Form12.Show
End Sub

Private Sub cntsheet_Click()
    Dim ds As ADODB.Recordset, s As String, w As String, i As Double
    Dim rh As String, rt As String, rf As String, ss As ADODB.Recordset
    Dim tbb As Long, t4 As Long
    tbb = 0: t4 = 0
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    s = "select * from racks "
    s = "select r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
    s = s & " from racks r, rackpos p"
    
    If qs(0).Checked = True Then
        'w = "where aisle = '" & PAisle & "' order by slot,sku desc"
        w = " where r.aisle = '" & PAisle & "' and p.rackno = r.id"
        w = w & " and (p.sku > ' ' or r.resv_sku > ' ')"
        w = w & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        w = w & " order by r.slot,p.sku desc"
        rh = "Aisle " & PAisle
    End If
    If qs(1).Checked = True Then
        'w = "where sku = '" & Right$(Form4.Caption, 3) & "' order by lot_num,aisle,slot"
        w = " where p.sku = '" & Trim(Right$(Form4.Caption, 4)) & "' and p.rackno = r.id"
        w = w & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        w = w & " order by p.lot_num,r.aisle,r.slot"
        rh = "SKU " & Trim(Right$(Form4.Caption, 4))
    End If
    If qs(2).Checked = True Then
        'w = "where hold = 1 order by aisle,slot"
        w = " where r.hold = 1 and p.rackno = r.id"
        w = w & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        w = w & " order by r.aisle,r.slot"
        rh = "On Hold Racks"
    End If
    If qs(3).Checked = True Then
        'w = "where fo = 1 order by aisle,slot"
        w = " where r.fo = 1 and p.rackno = r.id"
        w = w & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        w = w & " order by r.aisle,r.slot"
        rh = "1st Out Racks"
    End If
    If qs(4).Checked = True Then
        'w = "where resv_sku > ' ' order by aisle,slot"
        w = " where r.resv_sku > ' ' and p.rackno = r.id"
        w = w & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        w = w & " order by r.aisle,r.slot,p.sku desc"
        rh = "Reserved Racks"
    End If
    If qs(5).Checked = True Then
        'w = "where qty4 > 0 order by sku,aisle,slot"
        w = " where p.bbc = 'N' and p.rackno = r.id"
        w = w & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        w = w & " order by r.aisle,r.slot,p.sku desc"
        rh = "4 Way Pallets"
    End If
    If qs(6).Checked = True Then
        'w = "where qty = 0 and qty4 = 0 order by aisle,slot"
        w = " where p.count_qty = 0 and p.rackno = r.id"
        w = w & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
        w = w & " order by r.aisle,r.slot"
        rh = "Empty Racks"
    End If
    
    s = s & w
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(0) & Chr(9)
            s = s & ds(1) & Chr(9)
            s = s & ds(3) & Chr(9)
            s = s & ds(2) & Chr(9)
            If ds(2) > "0" Then
                If skurec(Val(ds(2))).sku = ds(2) Then
                    s = s & skurec(Val(ds(2))).prodname
                Else
                    s = s & ".."
                End If
            End If
            s = s & Chr(9)
            If ds(8) = "Y" Then
                s = s & Format(ds(10), "#")
                tbb = tbb + ds(10)
            End If
            s = s & Chr(9)
            If ds(8) = "N" Then
                s = s & Format(ds(10), "#")
                t4 = t4 + ds(10)
            End If
            s = s & Chr(9)
            If ds(5) > "." Then
                s = s & "<" & ds(5)
            Else
                If ds(7) <> 0 Then
                    s = s & "On Hold"
                Else
                    If ds(6) <> 0 Then
                        s = s & "1st Out"
                    End If
                End If
            End If
            Grid1.AddItem s
            
            ds.MoveNext
        Loop
    End If
    s = "." & Chr(9) & "." & Chr(9) & "." & Chr(9) & "." & Chr(9) & "." & Chr(9)
    If tbb > 0 Then
        s = s & tbb & Chr(9)
    Else
        s = s & "." & Chr(9)
    End If
    If t4 > 0 Then
        s = s & t4 & Chr(9)
    Else
        s = s & "." & Chr(9)
    End If
    s = s & "."
    Grid1.AddItem s
    
    ds.Close
    Grid1.FormatString = "^Aisle|^Rack|^Lot #|^SKU|<Contents|^BB|^4 Way|^Status"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 3600
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1200
    
    rt = "Rack Count Sheet"
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
End Sub

Private Sub Command1_Click()
    Dim sqlx As String, pkey As Long, ds As ADODB.Recordset, i As Integer
    Dim maisle As String, mrack As String                               'jv062111
    If Val(RGrid.TextMatrix(RGrid.Row, 12)) = 0 Then Exit Sub           'jv062111
    'pkey = Val(RGrid.TextMatrix(RGrid.Row, 0))
    'If pkey = 0 Then Exit Sub
    If RGrid.TextMatrix(RGrid.Row, 9) = "1" Then
        RGrid.TextMatrix(RGrid.Row, 9) = "0"
    Else
        RGrid.TextMatrix(RGrid.Row, 9) = "1"
    End If
    maisle = Left(RGrid.TextMatrix(RGrid.Row, 1), 1)                    'jv062111
    mrack = Right(RGrid.TextMatrix(RGrid.Row, 1), Len(RGrid.TextMatrix(RGrid.Row, 1)) - 2)
    
    sqlx = "update racks set fo = " & RGrid.TextMatrix(RGrid.Row, 9)    'jv062111
    sqlx = sqlx & " where aisle = '" & maisle & "'"                     'jv062111
    sqlx = sqlx & " and rack = '" & mrack & "'"                         'jv062111
    Wdb.Execute sqlx                                                     'jv062111
    i = RGrid.Row
    Call Refresh_racks(i)
    DoEvents
    RGrid.Row = i
    Call RGrid_Click
End Sub

Private Sub Command2_Click()
    Dim sqlx As String, pkey As Long, ds As ADODB.Recordset, i As Integer
    Dim p As ptask, cfile As String                                     'jv111314
    Dim zid As Long, psku As String, plot As String, pcode As String    'jv040715
    Dim ps As ADODB.Recordset, ppal As String                                 'jv040715
    Dim maisle As String, mrack As String                               'jv062111
    If Val(RGrid.TextMatrix(RGrid.Row, 12)) = 0 Then Exit Sub           'jv062111
    If Left(RGrid.TextMatrix(RGrid.Row, 1), 1) = "M" Then Exit Sub      'jv111314
    If RGrid.TextMatrix(RGrid.Row, 10) = "1" Then
        RGrid.TextMatrix(RGrid.Row, 10) = "0"
    Else
        RGrid.TextMatrix(RGrid.Row, 10) = "1"
    End If
    maisle = Left(RGrid.TextMatrix(RGrid.Row, 1), 1)                    'jv062111
    mrack = Right(RGrid.TextMatrix(RGrid.Row, 1), Len(RGrid.TextMatrix(RGrid.Row, 1)) - 2)
    
    sqlx = "Update racks Set hold = " & RGrid.TextMatrix(RGrid.Row, 10)
    sqlx = sqlx & " Where aisle = '" & maisle & "'"
    sqlx = sqlx & " and rack = '" & mrack & "'"
    Wdb.Execute sqlx
    
    cfile = Form1.logdir & "move" & Format(Now, "mmddyyyy") & ".txt"            'jv111314
    Open cfile For Append Shared As #1                                          'jv111314
    sqlx = "select * from rackpos where rackno in ("
    sqlx = sqlx & "select id from racks where aisle = '" & maisle & "'"
    sqlx = sqlx & " and rack = '" & mrack & "')"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            p.id = ds!id                                                        'jv111314
            p.area = "HOLD"                                                     'jv111314
            p.description = " "                                                 'jv111314
            If RGrid.TextMatrix(RGrid.Row, 10) = "1" Then
                sqlx = "Update rackpos set hold = 'Y' Where id = " & ds!id
                Wdb.Execute sqlx
                p.source = maisle & "-" & mrack                                 'jv111314
                p.target = "HOLD"                                               'jv111314
            Else
                sqlx = "Update rackpos set hold = 'N' Where id = " & ds!id
                Wdb.Execute sqlx
                p.source = "HOLD"                                               'jv111314
                p.target = maisle & "-" & mrack                                 'jv111314
            End If
            If ds!sku > "0" Then
                p.product = ds!sku                                              'jv111314
                For i = 1 To RGrid.Rows - 1                                     'jv111314
                    If RGrid.TextMatrix(i, 3) = ds!sku Then                     'jv111314
                        p.product = ds!sku & " " & RGrid.TextMatrix(i, 11)      'jv111314
                        Exit For                                                'jv111314
                    End If                                                      'jv111314
                Next i                                                          'jv111314
                p.palletid = ds!barcode                                         'jv111314
                p.qty = "1"                                                     'jv111314
                p.uom = "Pallet"                                                'jv111314
                p.lotnum = ds!lot_num                                           'jv111314
                p.units = ds!count_qty                                          'jv111314
                p.lotnum2 = ds!lot2                                             'jv111314
                p.units2 = ds!qty2                                              'jv111314
                p.status = "COMP"                                               'jv111314
                p.userid = Form1.userid                                         'jv111314
                p.trandate = Format(Now, "yyMMdd hh:mm:ss")                     'jv111314
                p.reqid = ".."                                                  'jv111314
                Write #1, p.id                                                  'jv111314
                Write #1, p.area;                                               'jv111314
                Write #1, p.description;                                        'jv111314
                Write #1, p.source;                                             'jv111314
                Write #1, p.target;                                             'jv111314
                Write #1, p.product;                                            'jv111314
                Write #1, p.palletid;                                           'jv111314
                Write #1, p.qty;                                                'jv111314
                Write #1, p.uom;                                                'jv111314
                Write #1, p.lotnum;                                             'jv111314
                Write #1, p.units;                                              'jv111314
                Write #1, p.lotnum2;                                            'jv111314
                Write #1, p.units2;                                             'jv111314
                Write #1, p.status;                                             'jv111314
                Write #1, p.userid;                                             'jv111314
                Write #1, p.trandate;                                           'jv111314
                Write #1, p.reqid                                               'jv111314
                'zid = p.id                                                                      'jv040715
                zid = wd_seq("HoldList")                    'jv042015
                psku = ds!sku                                                                   'jv040715
                plot = ds!lot_num                                                               'jv040715
                'pcode = Mid(ds!barcode, 12, 1)                                                  'jv040715
                pcode = Trim(Mid(ds!barcode, 11, 3))                                       'jv052515
                ppal = Mid(ds!barcode, 14, 3)                                                   'jv040715
                'If psku > "000" And psku < "999" And plot > "00000" And plot < "99999" Then     'jv040915
                If psku > "000" And psku < "9999" And plot > "00000" And plot < "99999" Then     'jv082415
                    's = "select id from pallets where barcode = '" & ds!barcode & "'"           'jv040715
                    'Set ps = db.OpenRecordset(s)                                                                'jv040715
                    'If ps.BOF = False Then                                                                      'jv040715
                    '    ps.MoveFirst                                                                            'jv040715
                    '    zid = ps!id                                                                             'jv040715
                    'End If                                                                                      'jv040715
                    'ps.Close                                                                                    'jv040715
                    'If ds!hold = "Y" Then                                                        'jv040715
                    If RGrid.TextMatrix(RGrid.Row, 10) = "1" Then
                        s = "Insert into holdlist (id, sku, lot_num, opcode, spallet, epallet, hsource, userid, holddate)"
                        s = s & " values (" & zid  'jv040715
                        s = s & ", '" & psku & "', '" & plot & "', '" & pcode & "', '" & ppal & "', '" & ppal & "', 'Racks'"
                        s = s & ", '" & WDUserId & "', '" & Format(Now, "yyMMdd hh:mm:ss") & "')"
                        Wdb.Execute s                                                        'jv040715
                    Else                                                                    'jv040715
                        s = "delete from holdlist where sku = '" & psku & "'"               'jv040715
                        s = s & " and lot_num = '" & plot & "'"                             'jv040715
                        s = s & " and opcode = '" & pcode & "'"                             'jv040715
                        s = s & " and spallet = '" & ppal & "'"                             'jv040715
                        s = s & " and epallet = '" & ppal & "'"                             'jv040715
                        Wdb.Execute s                                                        'jv040715
                    End If                                                                  'jv040715
                End If
            End If                                                              'jv111314
            
            ds.MoveNext
        Loop
    End If
    Close #1
    ds.Close
    i = RGrid.Row                       'jv062111
    Call Refresh_racks(i)               'jv062111
    DoEvents                            'jv062111
    RGrid.Row = i                       'jv062111
    Call RGrid_Click
End Sub

Private Sub Command4_Click()
    Dim y As Integer
    y = RGrid.Row
    Form5.Caption = "Rack " & Trim$(RGrid.TextMatrix(y, 1))
    'Form5.Text7 = Trim$(RGrid.TextMatrix(y, 2))
    DoEvents
    
    'Form5.RKey = Val(RGrid.TextMatrix(y, 0))
    Form5.RKey = RGrid.TextMatrix(y, 1)
    Form5.Caption = "Rack " & Trim$(RGrid.TextMatrix(y, 1))
    Form5.Text1 = Trim$(RGrid.TextMatrix(y, 3))
    Form5.Text2 = Trim$(RGrid.TextMatrix(y, 4))
    Form5.Text3 = Trim$(RGrid.TextMatrix(y, 5))
    Form5.Text4 = Trim$(RGrid.TextMatrix(y, 6))
    Form5.Text5 = Trim$(RGrid.TextMatrix(y, 7))
    Form5.Text6 = Trim$(RGrid.TextMatrix(y, 8))
    'Form5.Text7 = Trim$(RGrid.TextMatrix(y, 2))
    Form5.Text8 = Trim$(RGrid.TextMatrix(y, 12))
    Form5.Check1.Value = Val(RGrid.TextMatrix(y, 9))
    Form5.Check2.Value = Val(RGrid.TextMatrix(y, 10))
    RGrid.SetFocus
    Form5.Show
End Sub

Private Sub Command8_Click()
    'Call syl_floor_pallets
    'Call ba_Easisle_racks
    'Call t10_a_d_aisles
    Call bae18e19
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Form4.WindowState = 0 Then
        For i = 1 To Form1.Frmgrid.Rows - 1
            If Form1.Frmgrid.TextMatrix(i, 0) = "form4" Then
                Form1.Frmgrid.TextMatrix(i, 1) = Form4.Top
                Form1.Frmgrid.TextMatrix(i, 2) = Form4.Left
                Form1.Frmgrid.TextMatrix(i, 3) = Form4.Height
                Form1.Frmgrid.TextMatrix(i, 4) = Form4.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then Call Command4_Click 'F2 - Edit Rack
    If KeyCode = 70 And Shift = 2 Then Call Command1_Click 'Ctrl-F First Out
    If KeyCode = 78 And Shift = 2 Then Call Command2_Click 'Ctrl-N Tag Hold Rack
End Sub

Private Sub Form_Load()
    Dim ds As ADODB.Recordset, sqlx As String, i As Integer
    For i = 1 To Form1.Frmgrid.Rows - 1
        If Form1.Frmgrid.TextMatrix(i, 0) = "form4" Then
            Form4.Top = Val(Form1.Frmgrid.TextMatrix(i, 1))
            Form4.Left = Val(Form1.Frmgrid.TextMatrix(i, 2))
            Form4.Height = Val(Form1.Frmgrid.TextMatrix(i, 3))
            Form4.Width = Val(Form1.Frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    sqlx = "select distinct aisle from racks order by aisle"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If Len(ds!aisle) > 0 Then PAisle.AddItem ds!aisle
            ds.MoveNext
        Loop
    End If
    ds.Close
    PAisle.ListIndex = 0
End Sub

Private Sub Form_Resize()
    RGrid.Width = Me.Width - 100
    If Me.Height > 2175 Then
        RGrid.Height = Me.Height - 1180
    End If
    Grid1.Width = Me.Width - 80
    Grid5.Width = Me.Width - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub PAisle_Click()
    Dim i As Integer
    For i = 0 To 7
        qs(i).Checked = False
    Next i
    qs(0).Checked = True
    Call Refresh_racks(1)
    If RGrid.Visible = True Then RGrid.SetFocus
End Sub

Private Sub qs_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 7
        qs(i).Checked = False
    Next i
    qs(Index).Checked = True
    Call Refresh_racks(1)
    RGrid.SetFocus
End Sub

Private Sub RGrid_Click()
    Dim y As Integer
    RGrid.Col = 1: RGrid.ColSel = RGrid.Cols - 1
    If Form5.Visible = True Then
        y = RGrid.Row
        Form5.Caption = "Rack " & Trim$(RGrid.TextMatrix(y, 1))
        'Form5.Text7 = Trim$(RGrid.TextMatrix(y, 2))
        DoEvents
        'Form5.RKey = Val(RGrid.TextMatrix(y, 0))
        Form5.RKey = RGrid.TextMatrix(y, 1)
        Form5.Caption = "Rack " & Trim$(RGrid.TextMatrix(y, 1))
        Form5.Text1 = Trim$(RGrid.TextMatrix(y, 3))
        Form5.Text2 = Trim$(RGrid.TextMatrix(y, 4))
        Form5.Text3 = Trim$(RGrid.TextMatrix(y, 5))
        Form5.Text4 = Trim$(RGrid.TextMatrix(y, 6))
        Form5.Text5 = Trim$(RGrid.TextMatrix(y, 7))
        Form5.Text6 = Trim$(RGrid.TextMatrix(y, 8))
        'Form5.Text7 = Trim$(RGrid.TextMatrix(y, 2))
        Form5.Text8 = Trim$(RGrid.TextMatrix(y, 12))
        Form5.Check1.Value = Val(RGrid.TextMatrix(y, 9))
        Form5.Check2.Value = Val(RGrid.TextMatrix(y, 10))
    End If
End Sub

Private Sub RGrid_GotFocus()
    Call RGrid_Click
End Sub

Private Sub RGrid_LostFocus()
    Call RGrid_Click
End Sub

Private Sub RGrid_RowColChange()
    RGrid.ColSel = RGrid.Cols - 1
End Sub
