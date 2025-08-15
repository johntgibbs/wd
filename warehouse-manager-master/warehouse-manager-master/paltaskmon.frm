VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form8 
   Caption         =   "Pallet Task Monitor"
   ClientHeight    =   7140
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11130
   LinkTopic       =   "Form8"
   ScaleHeight     =   7140
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   960
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7646
      _Version        =   327680
      BackColorSel    =   128
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Completed"
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
      Left            =   6000
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Area:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu mtc 
         Caption         =   "Mark Task - Complete"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mtp 
         Caption         =   "Mark Task - Pending"
      End
      Begin VB.Menu edsrc 
         Caption         =   "Change Source"
         Enabled         =   0   'False
      End
      Begin VB.Menu edtar 
         Caption         =   "Change Target"
         Enabled         =   0   'False
      End
      Begin VB.Menu edbc 
         Caption         =   "Change BarCode"
      End
      Begin VB.Menu insque 
         Caption         =   "Insert Queue Record"
      End
      Begin VB.Menu swapque 
         Caption         =   "Swap Queue Warehouse "
      End
      Begin VB.Menu edque 
         Caption         =   "Edit Queue Field"
      End
      Begin VB.Menu cu 
         Caption         =   "Clear User"
      End
      Begin VB.Menu palhist 
         Caption         =   "View Pallet History"
      End
      Begin VB.Menu batonhand 
         Caption         =   "View Batch Inventory"
      End
   End
   Begin VB.Menu usermenu 
      Caption         =   "User"
      Visible         =   0   'False
      Begin VB.Menu emplook 
         Caption         =   "Lookup Employee Name"
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub post_wms(rk As Integer)
    Dim cfile As String, p As ptask
    If LCase(Form1.userid) = "jvierus" Then
        Exit Sub
    End If
    p.id = Grid1.TextMatrix(rk, 0)
    p.area = Grid1.TextMatrix(rk, 1)
    p.description = Grid1.TextMatrix(rk, 15)
    p.source = Grid1.TextMatrix(rk, 2)
    p.target = Grid1.TextMatrix(rk, 3)
    p.product = Grid1.TextMatrix(rk, 4)
    p.palletid = Grid1.TextMatrix(rk, 5)
    p.qty = Grid1.TextMatrix(rk, 6)
    p.uom = Grid1.TextMatrix(rk, 7)
    p.lotnum = Grid1.TextMatrix(rk, 9)
    p.units = Grid1.TextMatrix(rk, 8)
    p.lotnum2 = Grid1.TextMatrix(rk, 11)
    p.units2 = Grid1.TextMatrix(rk, 10)
    p.status = Grid1.TextMatrix(rk, 12)
    p.userid = Grid1.TextMatrix(rk, 13)
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")
    p.reqid = ".."
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"
    Open cfile For Append Shared As #1
    Write #1, p.id;
    Write #1, p.area;
    Write #1, p.description;
    Write #1, p.source;
    Write #1, p.target;
    Write #1, p.product;
    Write #1, p.palletid;
    Write #1, p.qty;
    Write #1, p.uom;
    Write #1, p.lotnum;
    Write #1, p.units;
    Write #1, p.lotnum2;
    Write #1, p.units2;
    Write #1, p.status;
    Write #1, p.userid;
    Write #1, p.trandate;
    Write #1, p.reqid
    Close #1
End Sub

Private Sub refresh_queues()                    'Crane Pallet Queues
    Dim ds As adodb.Recordset, s As String, i As Integer
    Dim q As String
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    s = "select * from queue_infc order by whse_num, source, queue_num"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!whse_num & Chr(9)
            i = Val(ds!sku)
            If skurec(i).sku = ds!sku Then
                s = s & ds!sku & " " & StrConv(skurec(i).prodname, vbProperCase) & Chr(9)
            Else
                s = s & ds!sku & " Invalid SKU!!!!!" & Chr(9)
            End If
            s = s & ds!lot_num & Chr(9)
            s = s & ds!drop_flag & Chr(9)
            s = s & ds!queue_num & Chr(9)
            s = s & ds!rack_num & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!lot_num2 & Chr(9)
            s = s & ds!units2 & Chr(9)
            s = s & ds!palletid & Chr(9)
            s = s & ds!source
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    ycolor.Visible = False
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 5) = "0" Then 'Or Grid1.TextMatrix(i, 13) > "." Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ycolor.BackColor
                ycolor.Visible = True
            End If
        Next i
        Grid1.Row = 1: Grid1.Col = 1
    End If
    s = "^ID|^SR|<SKU|^Lot|^Drop|^Queue|^Wraps|^Units|^Lot2|^Units2|<BarCode|^Source"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 600
    Grid1.ColWidth(2) = 2800
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 600
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 1900
    Grid1.ColWidth(11) = 800
    'Grid1.ColWidth(12) = 800
    'Grid1.ColWidth(13) = 1200
    'Grid1.ColWidth(14) = 1400 '800
    mtp.Enabled = False
    edbc.Enabled = False
    edque.Enabled = True
    swapque.Enabled = True
    If ycolor.Visible = False Then insque.Enabled = True
    Grid1_RowColChange
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_grid1()
    Dim ds As adodb.Recordset, s As String, i As Integer
    If Combo1 = "Crane Queues" Then
        refresh_queues
        Exit Sub
    End If
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 16
    Grid1.Redraw = False
    If Combo1 = "MISC, SR Drops" Then
        s = "select * from paltasks where description = '2 Step Request'"
        s = s & " or (description >= 'DROP' and description < 'DROPZZZZ')"
        s = s & " or area in ('GROUP-COMP', 'NONE')"
    Else
        If Combo1 = "DOCK-All" Then
            s = "select * from paltasks where area = 'DOCK'"
        Else
            If Combo1 = "DOCK-Active" Then
                s = "select * from paltasks where area = 'DOCK'"
                s = s & " and source <> 'ALT'"
                s = s & " and status = 'PEND'"
                s = s & " and description in (select left(product, 6) from paltasks where area = 'GROUP' and status = 'ACTV')"
            Else
                s = "select * from paltasks where area = '" & Combo1 & "'"
            End If
        End If
    End If
    's = s & " order by status DESC,userid DESC,palletid DESC,trandate,id"
    s = s & " order by status DESC,id"
    If Combo1 = "Crane Output" Then
        s = "select * from paltasks where area = 'DOCK' and status = 'PEND'"
        s = s & " and userid < '0' and lotnum > '0'"
        s = s & " and palletid > '0' and source in ('SR1','SR2','SR3','SR5')"
        s = s & " order by trandate"
    End If
    If Combo1 = "Snack Plant Trailer" Then
        s = "select * from paltasks where area = 'DOCK' and status = 'PEND'"
        s = s & " and palletid > '0' and source in ('1405','1406','1731','SNACK PLANT')"
        s = s & " and description < '0'"
        s = s & " order by trandate"
    End If
    
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!area & Chr(9)
            s = s & Trim(ds!source) & Chr(9)
            s = s & Trim(ds!target) & Chr(9)
            s = s & ds!product & Chr(9)
            s = s & ds!palletid & Chr(9)
            s = s & ds!qty & Chr(9)
            s = s & ds!uom & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!lotnum & Chr(9)
            s = s & ds!units2 & Chr(9)
            s = s & ds!lotnum2 & Chr(9)
            s = s & ds!status & Chr(9)
            s = s & ds!userid & Chr(9)
            s = s & ds!trandate & Chr(9)
            s = s & ds!description
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    ycolor.Visible = False
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 12) <> "PEND" Then 'Or Grid1.TextMatrix(i, 13) > "." Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ycolor.BackColor
                ycolor.Visible = True
            End If
        Next i
        Grid1.Row = 1: Grid1.Col = 1
    End If
    's = "^ID|^Area|^Source|^Target|<Product|^BarCode|^Qty|^UOM|^Units|^Lot|^Units2|^Lot2|^Status|^User"
    s = "^ID|<Area|<Source|<Target|<Product|^BarCode|^Qty|^UOM|||||^Status|^User|^Time|<Group"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 1400
    Grid1.ColWidth(2) = 1400
    Grid1.ColWidth(3) = 1400
    Grid1.ColWidth(4) = 3800
    Grid1.ColWidth(5) = 1800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 1 '800
    Grid1.ColWidth(9) = 1 '800
    Grid1.ColWidth(10) = 1 '800
    Grid1.ColWidth(11) = 1 '800
    Grid1.ColWidth(12) = 800
    Grid1.ColWidth(13) = 1200
    Grid1.ColWidth(14) = 1400 '800
    Grid1.ColWidth(15) = 1200
    mtp.Enabled = True
    edbc.Enabled = True
    edque.Enabled = False
    swapque.Enabled = False
    insque.Enabled = False
    Grid1.Redraw = True
    Grid1_RowColChange
    Screen.MousePointer = 0
End Sub

Private Sub batonhand_Click()                               'jv051717
    Dim s As String, d As String
    If Grid1.TextMatrix(0, 5) = "Queue" Then
        s = Left(Grid1.TextMatrix(Grid1.Row, 10), 13)
        d = Grid1.TextMatrix(Grid1.Row, 2)
    Else
        s = Left(Grid1.TextMatrix(Grid1.Row, 5), 13)
        d = Grid1.TextMatrix(Grid1.Row, 4)
    End If
    'MsgBox s
    If s > "0" Then
        tktonhand.bbarcode = s
        tktonhand.bproduct = d
        tktonhand.Show
    End If
End Sub

Private Sub Combo1_Click()
    refresh_grid1
End Sub

Private Sub cu_Click()                  'Clear UserId from Task
    Dim i As Long, ds As adodb.Recordset, s As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    s = "select userid from paltasks where id = " & i
    s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set userid = ' ' Where id = " & i
        Wdb.Execute s
        Grid1.TextMatrix(Grid1.Row, 13) = " "
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
        If Grid1.TextMatrix(Grid1.Row, 12) = "PEND" Then
            Grid1.CellBackColor = Grid1.BackColor
        Else
            Grid1.CellBackColor = ycolor.BackColor
        End If
        Grid1.Col = 1
    End If
    ds.Close
End Sub

Private Sub edbc_Click()                'Edit Pallet BarCode
    Dim i As Long, ds As adodb.Recordset, s As String, ns As String, uqty As Integer
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 12) <> "PEND" Then
        MsgBox "Task is not Pending.", vbOKOnly + vbExclamation, "Sorry, not this task..."
        Exit Sub
    End If
    If Grid1.TextMatrix(Grid1.Row, 2) = "ALT" Then
        MsgBox "This is an ALTERNATE.", vbOKOnly + vbExclamation, "Sorry, not this task..."
        Exit Sub
    End If
    ns = Grid1.TextMatrix(Grid1.Row, 5)
    If ns = "..." Then
        If Len(ns) = 4 Then                                                                     'jv082415
            ns = ns & Format(DateAdd("yyyy", 2, Now), "MMddyy") & "000001"                      'jv082415
        Else
            ns = Left(Grid1.TextMatrix(Grid1.Row, 4), 3) & " 000000 X 001"
        End If
    End If
    ns = InputBox("New BarCode:", "Change BarCode...", ns)
    If Len(ns) = 0 Then Exit Sub
    If ns = Grid1.TextMatrix(Grid1.Row, 5) Then Exit Sub
    ns = UCase(ns)
    uqty = 0
    s = Trim(Left(ns, 4))
    If skurec(Val(s)).sku = s Then
        uqty = skurec(Val(s)).uom_per_pallet
    Else
        MsgBox ns & " unrecognized SKU..."
    End If
    If uqty > 0 Then
        s = "select palletid from paltasks where id = " & i
        s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Grid1.TextMatrix(Grid1.Row, 5) = ns
            s = "Update paltasks set palletid = '" & ns & "' Where id = " & i
            Wdb.Execute s
        End If
        ds.Close
    End If
End Sub

Private Sub edque_Click()               'Edit Crane Queue Task Fields
    Dim i As Long, ds As adodb.Recordset, s As String, nq As String, nf As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    nf = Grid1.TextMatrix(0, Grid1.Col)
    nq = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
    nq = InputBox("New " & nf & ":", "Change " & nf & "...", nq)
    If Len(nq) = 0 Then Exit Sub
    If nq = Grid1.Text Then Exit Sub
    
    s = "select * from queue_infc where id = " & i
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update Queue_infc "
        If nf = "SR" Then s = s & "Set whse_num = " & Val(nq)
        If nf = "SKU" Then s = s & "Set sku = '" & Val(nq) & "'"
        If nf = "Lot" Then s = s & "Set lot_num = '" & nq & "'"
        If nf = "Drop" Then
            If UCase(nq) = "H" Then                 'jv040215
                s = s & "Set drop_flag = 'H'"
                nq = "H"
            Else
                s = s & "Set drop_flag = ' '"
                nq = " "
            End If
        End If
        If nf = "Queue" Then s = s & "Set queue_num = " & Val(nq)
        If nf = "Wraps" Then s = s & "Set rack_num = " & Val(nq)
        If nf = "Units" Then s = s & "Set units = " & Val(nq)
        If nf = "Lot2" Then s = s & "Set lot_num2 = '" & nq & "'"
        If nf = "Units2" Then s = s & "Set units2 = " & Val(nq)
        If nf = "BarCode" Then s = s & "Set palletid = '" & nq & "'"
        If nf = "Source" Then s = s & "Set source = '" & nq & "'"
        s = s & " Where id = " & i
        Wdb.Execute s
        
        Grid1.Text = nq
    End If
    ds.Close
End Sub

Private Sub edsrc_Click()                   'Edit Pallet Source
    Dim i As Long, ds As adodb.Recordset, s As String, ns As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 12) <> "PEND" Then Exit Sub
    ns = Grid1.TextMatrix(Grid1.Row, 2)
    ns = InputBox("New Source:", "Change source...", ns)
    If Len(ns) = 0 Then Exit Sub
    If ns = Grid1.TextMatrix(Grid1.Row, 2) Then Exit Sub
    ns = UCase(ns)
    s = "select source from paltasks where id = " & i
    s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set source = '" & ns & "' Where id = " & i
        Wdb.Execute s
        Grid1.TextMatrix(Grid1.Row, 2) = ns
    End If
    ds.Close
End Sub

Private Sub edtar_Click()                   'Edit Task Target
    Dim i As Long, ds As adodb.Recordset, s As String, nt As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 12) <> "PEND" Then Exit Sub
    nt = Grid1.TextMatrix(Grid1.Row, 3)
    nt = InputBox("New Target:", "Change target...", nt)
    If Len(nt) = 0 Then Exit Sub
    If nt = Grid1.TextMatrix(Grid1.Row, 3) Then Exit Sub
    nt = UCase(nt)
    s = "select target from paltasks where id = " & i
    s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set target = '" & nt & "' Where id = " & i
        Wdb.Execute s
        Grid1.TextMatrix(Grid1.Row, 3) = nt
    End If
    ds.Close
End Sub

Private Sub emplook_Click()                 'Look Up Employee
    Dim ds As adodb.Recordset, s As String
    If Len(Grid1.Text) = 0 Then Exit Sub
    'SQL Database - bbsr
    s = "select * from valuelists where listname = 'wdempid'"
    s = s & " and listreturn = '" & Grid1.Text & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = ds!listdisplay
    Else
        s = "Employee #: " & Grid1.Text & " is not in WdEmp database."
    End If
    ds.Close
    MsgBox s, vbOKOnly + vbInformation, "WMS SQL Employee " & Grid1.Text & " ...."
End Sub

Private Sub Form_Load()
    Combo1.Clear
    If Form1.plantno = 50 Then
        Combo1.AddItem "DOCK-All"
        Combo1.AddItem "DOCK-Active"
        Combo1.AddItem "EFLMove"
        Combo1.AddItem "FORKLIFT"
        Combo1.AddItem "GROUP"
        Combo1.AddItem "MISC, SR Drops"  '"MISC"
        Combo1.AddItem "ROBOT ZERO"
        Combo1.AddItem "ROLLER BED"
        Combo1.AddItem "SNACK PLANT WRAPPER"
        Combo1.AddItem "TRAFFIC MASTER"
        Combo1.AddItem "TRI-LEVEL 1"
        Combo1.AddItem "TRI-LEVEL 2"
        Combo1.AddItem "TRI-LEVEL 3"
        Combo1.AddItem "TRI-LEVEL 4"
        Combo1.AddItem "Crane Queues"
        Combo1.AddItem "Crane Output"
        Combo1.AddItem "Snack Plant Trailer"
        emplook.Enabled = True
    End If
    If Form1.plantno = 51 Then
        Combo1.AddItem "DOCK-All"
        Combo1.AddItem "DOCK-Active"
        Combo1.AddItem "EFLMove"
        Combo1.AddItem "FORKLIFT"
        Combo1.AddItem "GROUP"
        Combo1.AddItem "MISC, SR Drops"  '"MISC"
        Combo1.AddItem "WRAPPER"
        emplook.Enabled = False
    End If
    If Form1.plantno = 52 Then
        Combo1.AddItem "DOCK-All"
        Combo1.AddItem "DOCK-Active"
        Combo1.AddItem "EFLMove"
        Combo1.AddItem "FORKLIFT"
        Combo1.AddItem "GROUP"
        Combo1.AddItem "MISC, SR Drops"  '"MISC"
        Combo1.AddItem "TRAFFIC MASTER"         'jvtcar
        Combo1.AddItem "WRAPPER"
        emplook.Enabled = False
    End If
    Combo1.ListIndex = 4
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1240 '1020
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Grid1.TextMatrix(0, Grid1.Col) = "User" And emplook.Enabled = True Then
            PopupMenu usermenu
        Else
            PopupMenu edmenu
        End If
    End If
End Sub

Private Sub Grid1_RowColChange()
    If Grid1.Row = 0 Then
        edmenu.Enabled = False
    Else
        edmenu.Enabled = True
    End If
    If Grid1.Rows = 1 Then
        edmenu.Enabled = False
    Else
        edmenu.Enabled = True
    End If
    If Grid1.Cols = 12 Then
        mtc.Enabled = True
        mtp.Enabled = False
        cu.Enabled = False
        edsrc.Enabled = False
        edtar.Enabled = False
        Exit Sub
    Else
        cu.Enabled = True
    End If
    If Grid1.TextMatrix(Grid1.Row, 12) = "COMP" Or Grid1.TextMatrix(Grid1.Row, 12) = "ACTV" Then
        mtc.Enabled = False
        mtp.Enabled = True
        edsrc.Enabled = False
        edtar.Enabled = False
    End If
    If Grid1.TextMatrix(Grid1.Row, 12) = "PEND" Then
        mtp.Enabled = False
        mtc.Enabled = True
        edsrc.Enabled = False
        edtar.Enabled = False
        If Combo1 = "DOCK" Then
            edsrc.Enabled = True
            edtar.Enabled = True
        End If
        If Combo1 = "FORKLIFT" Then
            edsrc.Enabled = True
            edtar.Enabled = True
        End If
        If Combo1 = "ROLLER BED" Then
            edsrc.Enabled = True
            edtar.Enabled = True
        End If
        If Combo1 = "EFLMove" Then
            edsrc.Enabled = True
            edtar.Enabled = True
        End If
        If Combo1 = "GROUP" Then
            edsrc.Enabled = True
            edtar.Enabled = True
        End If
    End If
End Sub

Private Sub insque_Click()                  'New Crane Queue Record
    Dim i As Long, s As String, qid As Long
    Dim psku As String, pwhs As String, psrc As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    pwhs = Grid1.TextMatrix(Grid1.Row, 1)
    psku = Trim(Left(Grid1.TextMatrix(Grid1.Row, 2), 4))                        'jv082415
    psrc = Grid1.TextMatrix(Grid1.Row, 11)
    qid = wd_seq("Queue_Infc")
    s = "INSERT INTO Queue_Infc (ID, Whse_Num, SKU, Lot_Num, Drop_Flag, Queue_Num,"
    s = s & " Rack_Num, Units, Lot_Num2, Units2, PalletID, Source)"
    s = s & " VALUES (" & qid & "," & pwhs & ",'" & psku & "','.',' ',0,0,0,' ',0,'...','" & psrc & "')"
    Wdb.Execute s
    DoEvents
    refresh_queues
End Sub

Private Sub mtc_Click()                     'Mark Task Completed
    Dim i As Long, ds As adodb.Recordset, s As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    If LCase(Grid1.TextMatrix(0, 5)) = "queue" Then
        s = "select queue_num from queue_infc where id = " & i
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            s = "Update queue_infc Set queue_num = 0 Where id = " & i
            Wdb.Execute s
            Grid1.TextMatrix(Grid1.Row, 5) = "0"
        End If
    Else
        s = "select userid,status from paltasks where id = " & i
        s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            s = "Update paltasks set status = 'COMP', userid = ' ' Where id = " & i
            Wdb.Execute s
            Grid1.TextMatrix(Grid1.Row, 12) = "COMP"
            Grid1.TextMatrix(Grid1.Row, 13) = Form1.userid
            Call post_wms(Grid1.Row)
        End If
    End If
    ds.Close
    i = Grid1.Col
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
    Grid1.CellBackColor = ycolor.BackColor
    Grid1.Col = i
    If Grid1.Row <> Grid1.Rows - 1 Then Grid1.Row = Grid1.Row + 1
End Sub

Private Sub mtp_Click()                     'Mark Task Pending
    Dim i As Long, ds As adodb.Recordset, s As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    s = "select userid,status from paltasks where id = " & i
    s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set status = 'PEND', userid = ' ' Where id = " & i
        Wdb.Execute s
        Grid1.TextMatrix(Grid1.Row, 12) = "PEND"
        Grid1.TextMatrix(Grid1.Row, 13) = " "
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
        Grid1.CellBackColor = Grid1.BackColor
        Grid1.Col = 1
    End If
    ds.Close
End Sub

Private Sub palhist_Click()
    palhistory.Show
    If Grid1.TextMatrix(0, 5) = "Queue" Then
        palhistory.barkey = Grid1.TextMatrix(Grid1.Row, 10)
    Else
        palhistory.barkey = Grid1.TextMatrix(Grid1.Row, 5)
    End If
End Sub

Private Sub swapque_Click()             'Swap Queue Warehouse
    Dim pwhs As String, psrc As String, sok As Boolean
    Dim i As Long, ds As adodb.Recordset, s As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    pwhs = Grid1.TextMatrix(Grid1.Row, 1)
    psrc = Grid1.TextMatrix(Grid1.Row, 11)
    pwhs = InputBox("Warehouse:", "Crane Warehouse..", pwhs)
    If Len(pwhs) = 0 Then Exit Sub
    If Val(pwhs) < 1 Or Val(pwhs) > 5 Then Exit Sub
    If Val(pwhs) = 4 Then Exit Sub
    psrc = InputBox("Source:", "Conveyor", psrc)
    If Len(psrc) = 0 Then Exit Sub
    sok = False
    If psrc = "TML" Then sok = True
    If psrc = "FG1" Then sok = True
    If psrc = "FG2" Then sok = True
    If psrc = "FG3" Then sok = True
    If psrc = "FG5" Then sok = True
    If sok = False Then
        MsgBox "Source: " & psrc & " is not valid.", vbExclamation + vbOKOnly, "try again.."
        Exit Sub
    End If
    s = "select * from queue_infc where id = " & i
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update queue_infc set whse_num = " & Val(pwhs) & ", source = '" & psrc & "' Where id = " & i
        Wdb.Execute s
        Grid1.TextMatrix(Grid1.Row, 1) = Val(pwhs)
        Grid1.TextMatrix(Grid1.Row, 11) = psrc
    End If
    ds.Close
End Sub
