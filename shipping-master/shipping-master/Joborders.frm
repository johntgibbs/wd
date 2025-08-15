VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Joborders 
   Caption         =   "Process Jobbing Pallets"
   ClientHeight    =   9690
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14220
   LinkTopic       =   "Form3"
   ScaleHeight     =   9690
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1950
      Left            =   11280
      TabIndex        =   14
      Top             =   4800
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   4545
      Left            =   8400
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   6240
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   2895
      Left            =   0
      TabIndex        =   10
      Top             =   6720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5106
      _Version        =   327680
      BackColorFixed  =   12648384
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   4800
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2990
      _Version        =   327680
      BackColorFixed  =   16777152
      BackColorSel    =   255
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4680
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   120
      Width           =   7935
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   6588
      _Version        =   327680
      BackColorSel    =   33023
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Jobbing Groups:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11280
      TabIndex        =   15
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Pallet Tasks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label gcode 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Group Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Ship Date:"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Account:"
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
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pallet Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jobbing Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6735
   End
   Begin VB.Menu savemenu 
      Caption         =   "Save"
      Begin VB.Menu saveorder 
         Caption         =   "Order"
      End
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu edord 
         Caption         =   "Order"
         Begin VB.Menu addosku 
            Caption         =   "Add SKU"
         End
         Begin VB.Menu canosku 
            Caption         =   "Cancel SKU"
         End
         Begin VB.Menu edofield 
            Caption         =   "Edit Order Field"
         End
         Begin VB.Menu posttogroup 
            Caption         =   "Post Pallets to Group"
         End
         Begin VB.Menu posttotrailers 
            Caption         =   "Post Order to Trailers"
            Visible         =   0   'False
         End
         Begin VB.Menu csdate 
            Caption         =   "Change Ship Date"
         End
      End
      Begin VB.Menu edgrp 
         Caption         =   "Group"
         Begin VB.Menu edgcansku 
            Caption         =   "Cancel SKU"
         End
         Begin VB.Menu edgwhs 
            Caption         =   "Change Warehouse"
         End
         Begin VB.Menu edgqty 
            Caption         =   "Change Pallet Qty"
         End
      End
      Begin VB.Menu edptasks 
         Caption         =   "Pallet Tasks"
         Begin VB.Menu edpsrc 
            Caption         =   "Change Source"
         End
         Begin VB.Menu edptarg 
            Caption         =   "Change Target"
         End
         Begin VB.Menu edpstatus 
            Caption         =   "Mark Complete"
         End
      End
   End
End
Attribute VB_Name = "Joborders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function clean_account(jacct As String) As String
    Dim s As String, i As Integer, c As String
    For i = 1 To Len(jacct)
        c = mid(jacct, i, 1)
        If c = "&" Then
            s = s & "+"
        Else
            If c = "'" Then
                s = s & "`"
            Else
                If c = "\" Or c = "/" Then
                    s = s & "-"
                Else
                    s = s & c
                End If
            End If
        End If
    Next i
    clean_account = s
End Function

Private Sub refresh_accounts()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    Combo2.Clear: List1.Clear: List2.Clear
    s = "select jobbing.branch,jobbing.account,jobbing.acctdesc from jobbing"
    s = s & " where jobbing.account in (select account from trailers where shipdate = '" & Combo1 & "')" ' where ra_flag = false)"
    s = s & " order by jobbing.acctdesc, jobbing.branch, jobbing.account"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            List1.AddItem ds!branch
            List2.AddItem ds!account
            Combo2.AddItem clean_account(ds!acctdesc)
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_accounts", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_accounts - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_groups()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    List3.Clear
    s = "select distinct groupcode from trailers where plant = " & Form1.plantno
    s = s & " and branch in (15, 16) and groupcode <> 'jobAdd'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            List3.AddItem ds(0)
            ds.MoveNext
        Loop
    End If
    ds.Close
    If List3.ListCount > 0 Then List3.ListIndex = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_groups", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_groups - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_dates()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    Combo1.Clear
    s = "select distinct shipdate from trailers"
    s = s & " where branch in (15, 16)"
    s = s & " order by shipdate"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem Format(ds(0), "MM-dd-yyyy")
            ds.MoveNext
        Loop
    Else
        Combo1.AddItem Format(Now, "MM-dd-yyyy")
    End If
    Combo1.ListIndex = 0
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_dates", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_dates - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_grid1()
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim f10 As String, f11 As String, f12 As String, f13 As String
    Dim cfile As String
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 15
    If Form1.plantno = "50" Then
        cfile = Form1.srserv & "\wd\jobbing\jo" & List2 & Format(Combo1, "mmddyy") & ".txt"
    Else
        cfile = Form1.srserv & "\f\user\waredist\data\jobbing\jo" & List2 & Format(Combo1, "mmddyyy") & ".txt"
    End If
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input Shared As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13
            pdesc = "...."
            s = "select fgunit,fgdesc from skumast where sku = '" & f3 & "'"
            Set ds = Sdb.Execute(s)
            If ds.BOF = False Then
                pdesc = StrConv(ds!fgunit & " " & ds!fgdesc, vbProperCase)
            End If
            ds.Close
            s = f0 & Chr(9)                 'Plant
            s = s & f1 & Chr(9)             'Branch
            s = s & f2 & Chr(9)             'Account
            s = s & f3 & Chr(9)             'SKU
            s = s & pdesc & Chr(9)          'Product
            s = s & f4 & Chr(9)             'Unit Total
            s = s & f5 & Chr(9)             'Pallet Qty
            s = s & f6 & Chr(9)             'Pallet Size
            s = s & f7 & Chr(9)             'Wrap Qty
            s = s & f8 & Chr(9)             'Wrap Size
            s = s & f9 & Chr(9)             'Unit Qty
            s = s & f10 & Chr(9)            'Net Qty
            s = s & f11 & Chr(9)            'Group Code
            gcode.Caption = f11
            s = s & f12 & Chr(9)            'Order Date
            s = s & f13                     'Ship Date
            Grid1.AddItem s
        Loop
        Close #1
    End If
    's = "^Plant|^Branch|^Account|^SKU|<Product|^TotalUnits|^Pallets|^PalSize|^Wraps|^WrapSize|^Units|^Net|^Group|^Date|^Ship"
    's = "|||^SKU|<Product|^Total Units|^Pallets|^PalSize|^Wraps|^WrapSize|^Units|^Net|||"
    s = "|||^SKU|<Product|^Total Units|^Pallets|^PalSize|^Wraps|^WrapSize|||||"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1 '800
    Grid1.ColWidth(1) = 1 '800
    Grid1.ColWidth(2) = 1 '800
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 4700
    Grid1.ColWidth(5) = 1400
    Grid1.ColWidth(6) = 1200
    Grid1.ColWidth(7) = 1200
    Grid1.ColWidth(8) = 1200
    Grid1.ColWidth(9) = 1200
    Grid1.ColWidth(10) = 1 '1200
    Grid1.ColWidth(11) = 1 '1200
    Grid1.ColWidth(12) = 1 '800
    Grid1.ColWidth(13) = 1 '800
    Grid1.ColWidth(14) = 1 '800
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid1", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid1 - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_grid2()
    Dim ds As adodb.Recordset, s As String
    Dim i As Integer, k As Integer
    On Error GoTo vberror
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 9
    If gcode > "..." Then
        s = "select * from ship_infc where order_num = '" & gcode & "'"
        s = s & " order by sku"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds!id & Chr(9)
                s = s & ds!to_whse_num & Chr(9)
                s = s & ds!sku & Chr(9)
                s = s & "..." & Chr(9)
                s = s & ds!order_qty & Chr(9)
                s = s & ds!ship_plt_qty & Chr(9)
                s = s & ds!order_qty - ds!ship_plt_qty & Chr(9)
                s = s & ds!ship_status & Chr(9)
                s = s & ds!gmasize
                Grid2.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
        If Grid2.Rows > 1 Then
            For i = 1 To Grid2.Rows - 1
                For k = 0 To Grid1.Rows - 1
                    If Grid1.TextMatrix(k, 3) = Grid2.TextMatrix(i, 2) Then
                        Grid2.TextMatrix(i, 3) = Grid1.TextMatrix(k, 4)
                        Exit For
                    End If
                Next k
            Next i
        End If
    End If
    s = "^ID|^SR|^SKU|<Product|^Ordered|^Shipped|^Net|^Status|^4Way Size"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 2500
    Grid2.ColWidth(4) = 1000
    Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 1000
    Grid2.ColWidth(7) = 1000
    Grid2.ColWidth(8) = 1000
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid2", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid2 - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_grid3()
    Dim ds As adodb.Recordset, s As String
    Dim i As Integer, k As Integer
    On Error GoTo vberror
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 10
    If gcode > "..." Then
        s = "select * from paltasks where description >= '" & gcode & "'"
        s = s & " and description < '" & gcode & "ZZZZZZ'"
        s = s & " and area in ('DOCK','FORKLIFT')"
        s = s & " order by status desc, product"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds!id & Chr(9)
                s = s & ds!area & Chr(9)
                s = s & ds!source & Chr(9)
                s = s & ds!target & Chr(9)
                s = s & ds!product & Chr(9)
                s = s & ds!palletid & Chr(9)
                's = s & ds!qty & Chr(9)
                's = s & ds!uom & Chr(9)
                s = s & ds!units & Chr(9)
                s = s & ds!units2 & Chr(9)
                s = s & ds!status & Chr(9)
                s = s & ds!userid
                Grid3.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    's = "^ID|<Area|<Source|<Target|<Product|^BarCode|^Qty|^UOM|^Status|^User"
    s = "^ID|<Area|<Source|<Target|<Product|^BarCode|^Units|^Units2|^Status|^User"
    Grid3.FormatString = s
    Grid3.ColWidth(0) = 800
    Grid3.ColWidth(1) = 1200
    Grid3.ColWidth(2) = 1200
    Grid3.ColWidth(3) = 2500
    Grid3.ColWidth(4) = 2500
    Grid3.ColWidth(5) = 1500
    Grid3.ColWidth(6) = 1000
    Grid3.ColWidth(7) = 1000
    Grid3.ColWidth(8) = 1000
    Grid3.ColWidth(9) = 1000
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid3", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid3 - Error Number: " & eno
        End
    End If
End Sub

Private Sub addosku_Click()
    Dim psku As String, ds As adodb.Recordset, s As String, uqty As Integer, wqty As String
    On Error GoTo vberror
    psku = InputBox("SKU:", "Add SKU to order....", "507")
    If Len(psku) = 0 Then Exit Sub
    wqty = InputBox("Total Wraps:", "Total Wrap Quantity...", "0")
    If Len(wqty) = 0 Then Exit Sub
    If Val(wqty) = 0 Then Exit Sub
    s = "select * from skumast where sku = '" & psku & "'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        uqty = ds!numwrap * Val(wqty)
        s = Form1.plantno & Chr(9)
        s = s & List1 & Chr(9)
        s = s & List2 & Chr(9)
        s = s & psku & Chr(9)
        s = s & StrConv(ds!fgunit & " " & ds!fgdesc, vbProperCase) & Chr(9)
        s = s & uqty & Chr(9)
        s = s & " " & Chr(9)
        s = s & ds!pallet & Chr(9)
        s = s & wqty & Chr(9)
        s = s & ds!numwrap & Chr(9)
        s = s & " " & Chr(9)
        's = s & uqty & Chr(9)
        s = s & "0" & Chr(9)
        s = s & " " & Chr(9)
        s = s & Format(Now, "MM-dd-yyyy") & Chr(9)
        s = s & Combo1
        Grid1.AddItem s
    End If
    ds.Close
    saveorder_Click
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "addosku_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " addosku_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub canosku_Click()
    Dim s As String
    If Grid1.Row < 1 Then Exit Sub
    s = Grid1.TextMatrix(Grid1.Row, 3) & " " & Grid1.TextMatrix(Grid1.Row, 4)
    If MsgBox("Ok to cancel " & s & " from this order?", vbYesNo + vbQuestion, "are you sure...") = vbNo Then Exit Sub
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Grid1.Rows = 1
    End If
    saveorder_Click
End Sub

Private Sub Combo1_Click()
    refresh_accounts
End Sub

Private Sub Combo2_Click()
    gcode.Caption = "..."
    List1.ListIndex = Combo2.ListIndex
    List2.ListIndex = Combo2.ListIndex
    refresh_grid1
    DoEvents
    refresh_grid2
    DoEvents
    refresh_grid3
End Sub

Private Sub csdate_Click()
    Dim ds As adodb.Recordset, s As String
    Dim odate As String, ndate As String
    Dim ofile As String, nfile As String
    On Error GoTo vberror
    ndate = Format(Now, "MM-dd-yyyy")
    odate = Combo1
    ndate = InputBox("New Ship date:", "Ship Date....", ndate)
    If Len(ndate) = 0 Then Exit Sub
    If IsDate(ndate) = False Then Exit Sub
    Screen.MousePointer = 11
    s = "select * from trailers where branch = " & List1
    s = s & " and account = '" & List2 & "'"
    s = s & " and shipdate = '" & odate & "'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update trailers set shipdate = '" & ndate & "' where id = " & ds!id
            Sdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    If Form1.plantno = "50" Then
        ofile = Form1.srserv & "\wd\jobbing\jo" & List2 & Format(odate, "mmddyy") & ".txt"
        nfile = Form1.srserv & "\wd\jobbing\jo" & List2 & Format(ndate, "mmddyy") & ".txt"
    Else
        ofile = Form1.srserv & "\f\user\waredist\data\jobbing\jo" & List2 & Format(odate, "mmddyy") & ".txt"
        nfile = Form1.srserv & "\f\user\waredist\data\jobbing\jo" & List2 & Format(ndate, "mmddyy") & ".txt"
    End If
    'MsgBox "name " & ofile & " as " & nfile
    Name ofile As nfile
    refresh_dates
    refresh_accounts
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "csdate_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " csdate_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edgcansku_Click()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    s = "select ship_status from ship_infc where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update ship_infc set ship_status = 'CANC' Where id = " & ds!id
        Wdb.Execute s
        Grid2.TextMatrix(Grid2.Row, 7) = "CANC"
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "edgcansku_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " edgcansku_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edgqty_Click()
    Dim ds As adodb.Recordset, s As String, pq As String
    On Error GoTo vberror
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    pq = Grid2.TextMatrix(Grid2.Row, 4)
    pq = InputBox("Order Qty:", "Pallet Order Qty.....", pq)
    If Len(pq) = 0 Then Exit Sub
    If Val(pq) < 0 Then Exit Sub
    s = "select * from ship_infc where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update ship_Infc set order_qty = " & Val(pq) & ", ship_status = "
        If ds!order_qty <= ds!ship_plt_qty Then
            s = s & "'DONE'"
        Else
            s = s & "'NEW'"
        End If
        s = s & " Where id = " & ds!id
        Wdb.Execute s
        Grid2.TextMatrix(Grid2.Row, 4) = pq
        Grid2.TextMatrix(Grid2.Row, 6) = Val(pq) - Val(Grid2.TextMatrix(Grid2.Row, 5))
        Grid2.TextMatrix(Grid2.Row, 7) = ds!ship_status
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "edgqty_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " edgqty_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edgwhs_Click()
    Dim ds As adodb.Recordset, s As String, pw As String
    On Error GoTo vberror
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    pw = Grid2.TextMatrix(Grid2.Row, 1)
    pw = InputBox("Warehouse:", "Warehouse 1 - 4 .....", pw)
    If Len(pw) = 0 Then Exit Sub
    If Val(pw) < 1 Or Val(pw) > 4 Then Exit Sub
    s = "select * from ship_infc where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If pw = "1" Then
            s = "Update ship_infc set to_whse_num = 1, to_vert_loc = 2, to_horz_loc = 18, to_rack_side = 'L'"
        End If
        If pw = "2" Then
            s = "Update ship_infc set to_whse_num = 2, to_vert_loc = 2, to_horz_loc = 22, to_rack_side = 'R'"
        End If
        If pw = "3" Then
            s = "Update ship_infc set to_whse_num = 3, to_vert_loc = 2, to_horz_loc = 43, to_rack_side = 'R'"
        End If
        If pw = "4" Then
            s = "Update ship_infc set to_whse_num = 4, to_vert_loc = 0, to_horz_loc = 0, to_rack_side = 'R'"
        End If
        s = s & " Where id = " & ds!id
        Wdb.Execute s
        Grid2.TextMatrix(Grid2.Row, 1) = pw
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "edgwhs_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " edgwhs_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edofield_Click()
    Dim pqty As String, s As String, k As Integer, su As Long, tu As Long
    Dim pu As Long
    If Grid1.Row < 1 Then Exit Sub
    pqty = Grid1.Text
    s = Grid1.TextMatrix(0, Grid1.Col)
    pqty = InputBox(s & ":", "Change " & s & "....", pqty)
    If Len(pqty) = 0 Then Exit Sub
    Grid1.Text = pqty
    k = Grid1.Row
    tu = Val(Grid1.TextMatrix(k, 5))
    
    pu = Val(Grid1.TextMatrix(k, 6)) * Val(Grid1.TextMatrix(k, 7))
    Grid1.TextMatrix(k, 8) = CInt((tu - pu) / Val(Grid1.TextMatrix(k, 9)))
    
    su = Val(Grid1.TextMatrix(k, 6)) * Val(Grid1.TextMatrix(k, 7))
    su = su + (Val(Grid1.TextMatrix(k, 8)) * Val(Grid1.TextMatrix(k, 9)))
    su = su + Val(Grid1.TextMatrix(k, 10))
    Grid1.TextMatrix(k, 11) = Format(tu - su, "0")
    saveorder_Click
End Sub

Private Sub edpsrc_Click()
    Dim ds As adodb.Recordset, s As String, psrc As String
    On Error GoTo vberror
    If Val(Grid3.TextMatrix(Grid3.Row, 0)) = 0 Then Exit Sub
    psrc = Grid3.TextMatrix(Grid3.Row, 2)
    psrc = InputBox("Pallet Source:", "Source.....", psrc)
    If Len(psrc) = 0 Then Exit Sub
    s = "select source from paltasks where id = " & Grid3.TextMatrix(Grid3.Row, 0)
    s = s & " and palletid = '" & Grid3.TextMatrix(Grid3.Row, 5) & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set source = '" & psrc & "' Where id = " & Grid3.TextMatrix(Grid3.Row, 0)
        Wdb.Execute s
        Grid3.TextMatrix(Grid3.Row, 2) = psrc
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "edpsrc_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " edpsrc_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edpstatus_Click()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    If Val(Grid3.TextMatrix(Grid3.Row, 0)) = 0 Then Exit Sub
    s = "select status from paltasks where id = " & Grid3.TextMatrix(Grid3.Row, 0)
    s = s & " and palletid = '" & Grid3.TextMatrix(Grid3.Row, 5) & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set status = 'COMP' Where id = " & Grid3.TextMatrix(Grid3.Row, 0)
        Wdb.Execute s
        Grid3.TextMatrix(Grid3.Row, 8) = "COMP"
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "edpstatus_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " edpstatus_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edptarg_Click()
    Dim ds As adodb.Recordset, s As String, ptar As String
    On Error GoTo vberror
    If Val(Grid3.TextMatrix(Grid3.Row, 0)) = 0 Then Exit Sub
    ptar = Grid3.TextMatrix(Grid3.Row, 3)
    ptar = InputBox("Pallet Target:", "Target.....", ptar)
    If Len(ptar) = 0 Then Exit Sub
    s = "select target from paltasks where id = " & Grid3.TextMatrix(Grid3.Row, 0)
    s = s & " and palletid = '" & Grid3.TextMatrix(Grid3.Row, 5) & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set target = '" & ptar & "' Where id = " & Grid3.TextMatrix(Grid3.Row, 0)
        Wdb.Execute s
        Grid3.TextMatrix(Grid3.Row, 3) = ptar
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "edptarg_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " edptarg_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Load()
    refresh_dates
    If Combo1.ListCount < 1 Then
        MsgBox "No Jobbing Orders exist.", vbOKOnly + vbInformation, "no Jobbing orders..."
        Exit Sub
    End If
    refresh_accounts
    refresh_groups
    If Form1.plantno <> "50" Then Grid2.Visible = False
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    Grid2.Width = Me.Width - 80
    Grid3.Width = Me.Width - 80
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edord
End Sub

Private Sub Grid1_RowColChange()
    If Grid1.Col >= 5 And Grid1.Col <= 9 Then
        edofield.Caption = "Edit " & Grid1.TextMatrix(0, Grid1.Col)
    Else
        edofield.Caption = "..."
    End If
End Sub

Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edgrp
End Sub

Private Sub Grid3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edptasks
End Sub

Private Sub posttogroup_Click()
    Dim ds As adodb.Recordset, s As String, p As ptask, sflag As Boolean
    Dim i As Integer, p4 As Boolean, pwhs As Integer, psku As String, pqty As Integer, psize As Integer
    Dim k As Integer, sg As String, gflag As Boolean, tcode As String, tno As String, oratkt As Long
    Dim sb As DAO.Database, ss As DAO.Recordset
    On Error GoTo vberror
    gflag = False
    If Grid1.Rows < 2 Then Exit Sub
    sg = gcode.Caption
    sg = InputBox("Group Code:", "group code....", sg)
    If Len(sg) = 0 Then Exit Sub
    If Len(sg) > 6 Then sg = Left(sg, 6)
    If sg <> gcode Then
        gcode = sg
        For i = 1 To Grid1.Rows - 1
            Grid1.TextMatrix(i, 12) = sg
        Next i
    End If
    s = "select id, groupcode, runid from trailers where branch = " & List1         'jv021815
    s = s & " and account = '" & List2 & "'"
    s = s & " and shipdate = '" & Combo1 & "'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        oratkt = ds!runid                                                           'jv021815
        Do Until ds.EOF
            s = "Update trailers set groupcode = '" & sg & "' Where id = " & ds!id
            Sdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Form1.plantno = "50" Then                                        'jv02082012
        s = "select * from ship_infc where order_num = '" & gcode & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = "Update ship_infc set order_qty = 0, ship_plt_qty = 0, ship_status = 'DONE', gmasize = 0"
                s = s & " Where id = " & ds!id
                Wdb.Execute s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    s = "select * from paltasks where area = 'DOCK' and description >= '" & gcode & "'"
    s = s & " and description < '" & gcode & "ZZZZZ'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            tcode = Trim(Left(ds!description, 6))           'jv100313
            If gcode = tcode Then                           'jv100313
                s = "Update paltasks set area = 'DONE', description = ' ', status = 'COMP', userid = ' '"
                s = s & " Where id = " & ds!id
                Wdb.Execute s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select * from paltasks where area = 'FORKLIFT'"
    s = s & " and description >= '" & gcode & "'"
    s = s & " and description < '" & gcode & "ZZZZZ'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            tcode = Trim(Left(ds!description, 6))           'jv100313
            If gcode = tcode Then                           'jv100313
                s = "Update paltasks set area = 'DONE', description = ' ', status = 'COMP', userid = ' '"
                s = s & " Where id = " & ds!id
                Wdb.Execute s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select * from paltasks where area = 'GROUP'"
    s = s & " and product >= '" & gcode & "'"
    s = s & " and product < '" & gcode & "ZZZZZ'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            tcode = Trim(Left(ds!product, 6))               'jv100313
            If gcode = tcode Then                           'jv100313
                s = "Update paltasks set area = 'GROUP-DONE', status = 'COMP', userid = ' '"
                s = s & " Where id = " & ds!id
                Wdb.Execute s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 6)) > 0 Then
            gflag = True
            s = "Are the pallets for " & Grid1.TextMatrix(i, 4) & " going to be re-stacked?"
            If MsgBox(s, vbYesNo + vbQuestion, "re-stacking question.....") = vbNo Then
                sflag = True
                s = "select * from trailers where branch = " & List1        'R12 1204
                s = s & " and account = '" & List2 & "'"                    'R12 1204
                s = s & " and shipdate = '" & Combo1 & "'"                  'R12 1204
                s = s & " and sku = '" & Grid1.TextMatrix(i, 3) & "'"       'R12 1204
                Set ds = Sdb.Execute(s)                               'R12 1204
                If ds.BOF = False Then                                      'R12 1204
                    ds.MoveFirst                                            'R12 1204
                    oratkt = ds!runid                                       'jv021815
                    s = "Update trailers set pallets = " & Val(Grid1.TextMatrix(i, 6))
                    If Val(Grid1.TextMatrix(i, 8)) > 0 Then                 'R12 1204
                        s = s & ", wraps = " & Val(Grid1.TextMatrix(i, 8))              'R12 1204
                    Else                                                    'R12 1204
                        s = s & ", wraps = 0"                                        'R12 1204
                    End If                                                  'R12 1204
                    s = s & ", units = " & Val(Grid1.TextMatrix(i, 5))
                    s = s & " Where id = " & ds!id
                    Sdb.Execute s
                End If                                                      'R12 1204
                ds.Close
            Else
                sflag = False
            End If
            p4 = False: pwhs = 4: psize = 0
            psku = Grid1.TextMatrix(i, 3)
            pqty = Val(Grid1.TextMatrix(i, 6))
            s = "select uom_per_pallet from sku_config where sku = '" & psku & "'"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                If Val(Grid1.TextMatrix(i, 7)) > ds!uom_per_pallet Then
                    p4 = True
                    psize = Val(Grid1.TextMatrix(i, 7))
                End If
            End If
            ds.Close
            If Form1.plantno = "50" Then                                    'jv02082012
                If p4 = False Then
                    s = "select whse_num,sum(qty) from lane where sku = '" & psku & "'"
                    s = s & " and lane_status not in ('B','H')"
                    s = s & " and gmasize = 0"                      'jv090413
                    s = s & " group by whse_num"
                    s = s & " having sum(qty) > 0"
                    s = s & " order by sum(qty) desc"
                    Set ds = Wdb.Execute(s)
                    If ds.BOF = False Then
                        ds.MoveFirst
                        pwhs = ds!whse_num
                    End If
                    ds.Close
                Else
                    s = "select whse_num,sum(qty) from lane where sku = '" & psku & "'"
                    s = s & " and lane_status not in ('B','H')"
                    s = s & " and gmasize > 0"                      'jv090413
                    s = s & " group by whse_num"
                    s = s & " having sum(qty) > 0"
                    s = s & " order by sum(qty) desc"
                    Set ds = Wdb.Execute(s)
                    If ds.BOF = False Then
                        ds.MoveFirst
                        pwhs = ds!whse_num
                    End If
                    ds.Close
                End If
                If pwhs < 4 Then
                    Call insert_ship_infc(gcode, psku, pwhs, pqty, psize)
                End If
                s = "select * from trailers where branch = " & List1    'jv071613
                s = s & " and account = '" & List2 & "'"
                s = s & " and shipdate = '" & Combo1 & "'"
                s = s & " and sku = '" & psku & "'"
                s = s & " and groupcode = '" & gcode & "'"
                Set ds = Sdb.Execute(s)
                If ds.BOF = False Then
                    ds.MoveFirst
                    oratkt = ds!runid                                   'jv021815
                    Do Until ds.EOF
                        s = "Update trailers set whs_num = " & pwhs & " Where id = " & ds!id
                        Sdb.Execute s
                        ds.MoveNext
                    Loop
                End If
                ds.Close
            End If
            If Form1.plantno = "52" Then                                    'jv02082012
                s = "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
                Set sb = OpenDatabase(mysqldev, dbcdrivernoprompt, True, s)
                s = "select dtExpiration from tLotData where bQAHold = 0"
                s = s & " and iItemMasterSysID in (select iItemMasterSysId from tItemMaster"
                s = s & " where sItemId > '" & psku & "'"
                s = s & " and sItemId < '" & psku & "ZZZZ'"
                s = s & " and nDefaultQuantity = " & Val(Grid1.TextMatrix(i, 7)) & ")"
                Set ss = sb.OpenRecordset(s, dbOpenSnapshot, dbSeeChanges, dbReadOnly)
                If ss.BOF = False Then
                    ss.MoveFirst
                    pwhs = 1
                End If
                ss.Close: sb.Close
            End If
            
            For k = 1 To pqty
                If pwhs = 4 Then
                    p.area = "FORKLIFT"
                    p.description = gcode & Space(8 - Len(gcode)) & Combo2 & " #1"
                    If p4 = True Then
                        p.source = "4WAY"
                    Else
                        p.source = "RACKS"
                    End If
                    If sflag = True Then
                        p.target = "STAGING"
                    Else
                        p.target = "ORDER PICK"
                    End If
                    p.product = psku & " " & UCase(Grid1.TextMatrix(i, 4))
                    p.palletid = "..."
                    p.qty = "1"
                    p.uom = "Pallet"
                    p.lotnum = " "
                    p.units = Val(Grid1.TextMatrix(i, 7))
                    p.lotnum2 = " "
                    p.units2 = 0
                    p.status = "PEND"
                    p.userid = " "
                    p.trandate = Format(Now, "yyMMdd hh:mm:ss")
                    p.reqid = " "
                    Call insert_trans(p)
                    If sflag = True Then
                        p.area = "DOCK"
                        p.description = gcode '& " #1"
                        p.source = "STAGING"
                        p.target = Combo2 & " " & Grid1.TextMatrix(i, 13) ' #1"
                        tno = Grid1.TextMatrix(i, 13)               'jv050714
                        p.product = psku & " " & UCase(Grid1.TextMatrix(i, 4))
                        p.palletid = "..."
                        p.qty = "1"
                        p.uom = "Pallet"
                        p.lotnum = " "
                        p.units = Val(Grid1.TextMatrix(i, 7))  '0   jv02082012
                        p.lotnum2 = " "
                        p.units2 = 0
                        p.status = "PEND"
                        p.userid = " "
                        p.trandate = Format(Now, "yyMMdd hh:mm:ss")
                        p.reqid = " "
                        Call insert_trans(p)
                    End If
                Else
                    p.area = "DOCK"
                    p.description = gcode
                    If Form1.plantno = "50" And pwhs = 5 And p4 = True Then     'jv082113
                        p.source = "SR6"                                        'jv082113
                    Else
                        p.source = "SR" & pwhs
                    End If
                    If sflag = True Then
                        p.target = Combo2 & " " & Grid1.TextMatrix(i, 13) ' #1"
                    Else
                        p.target = "STAGING " & gcode
                    End If
                    tno = Grid1.TextMatrix(i, 13)                   'jv050714
                    p.product = psku & " " & UCase(Grid1.TextMatrix(i, 4))
                    p.palletid = "..."
                    If Form1.plantno = "52" Then p.palletid = psku & " ...... . ..."  'jv02082012
                    If Form1.plantno = "50" And pwhs = 5 Then   'jv071513
                        If Len(psku) = 4 Then                                   'jv082415
                            p.palletid = psku & "...... . ..."                  'jv082415
                        Else                                                    'jv082415
                            p.palletid = psku & " ...... . ..."
                        End If                                                  'jv082415
                    End If
                    p.qty = "1"
                    p.uom = "Pallet"
                    p.lotnum = " "
                    p.units = Val(Grid1.TextMatrix(i, 7))   '0    jv02082012
                    p.lotnum2 = " "
                    p.units2 = 0
                    p.status = "PEND"
                    p.userid = " "
                    p.trandate = Format(Now, "yyMMdd hh:mm:ss")
                    p.reqid = " "
                    Call insert_trans(p)
                End If
            Next k
        End If
    Next i
    If gflag = True Then
        p.area = "GROUP"
        p.description = " "
        p.source = Combo2 & " " & tno   '" #1"      'jv050714
        p.target = "..."
        p.product = gcode & Space(8 - Len(gcode)) & Combo2 & " " & tno      '" #1"      'jv050714
        p.palletid = "..."
        p.qty = "0"
        p.uom = " "
        p.lotnum = " "
        p.units = "0"
        p.lotnum2 = " "
        p.units2 = "0"
        p.status = "PEND"
        p.userid = " "
        p.trandate = Format(Now, "yyMMdd hh:mm:ss")
        p.reqid = oratkt                                            'jv021815
        Call insert_trans(p)
    End If
    refresh_grid2
    DoEvents
    refresh_grid3
    DoEvents
    saveorder_Click
    DoEvents
    refresh_groups
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "posttogroup_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " posttogroup_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub posttotrailers_Click()
    Dim ds As adodb.Recordset, s As String, jrun As Long
    Dim i As Integer, wc As Integer, sdate As String, sgrp As String, pkey As Long
    On Error GoTo vberror
    If Grid1.Rows < 2 Then Exit Sub
    sdate = Format(Now, "mm-dd-yyyy")
    sdate = InputBox("Ship Date:", "ship date....", sdate)
    If Len(sdate) = 0 Then Exit Sub
    sgrp = "jobADD"
    sgrp = InputBox("Group Code:", "group code...", sgrp)
    If Len(sgrp) = 0 Then Exit Sub
    If Len(sgrp) > 6 Then sgrp = Left(sgrp, 6)
    pkey = wd_seq("Oratkt", Form1.schdb)
    s = "Insert into runs (id, loaded, destination, locname, trlno, trlsize, trldate, startime, pickup, oc)"
    s = s & " Values (" & pkey
    s = s & ", " & Form1.plantno
    s = s & ", " & List1
    If Len(Combo2) > 30 Then
        s = s & ", '" & Left(Combo2, 30) & "'"
    Else
        s = s & ", '" & Combo2 & "'"
    End If
    s = s & ", '#1'"
    s = s & ", 0"
    s = s & ", '" & sdate & "'"
    s = s & ", '12:00 PM'"
    s = s & ", 'Added for jobbing..'"
    s = s & ", '*')"
    Sdb.Execute s
    jrun = pkey
    For i = 1 To Grid1.Rows - 1
        wc = 1
        s = "select numwrap from skumast where sku = '" & Grid1.TextMatrix(i, 3) & "'"
        Set ds = Sdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            wc = ds!numwrap
        End If
        ds.Close
        pkey = wd_seq("trailers", Form1.shipdb)
        s = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno, sku"
        s = s & ", pallets, wraps, units, whs_num, pb_flag, ra_flag) Values (" & pkey
        s = s & ", " & jrun
        s = s & ", '" & sgrp & "'"
        s = s & ", " & Form1.plantno
        s = s & ", " & List1
        s = s & ", '" & List2 & "'"
        s = s & ", '" & sdate & "'"
        s = s & ", '#1'"
        s = s & ", '" & Grid1.TextMatrix(i, 3) & "'"
        s = s & ", 0"
        s = s & ", " & Val(Grid1.TextMatrix(i, 5)) / wc
        s = s & ", " & Val(Grid1.TextMatrix(i, 5))
        s = s & ", 0, 'N', 'N')"
        Sdb.Execute s
        gcode = sgrp
        Grid1.TextMatrix(i, 12) = sgrp
        Grid1.TextMatrix(i, 13) = Format(Now, "MM-dd-yyyy")
        Grid1.TextMatrix(i, 14) = Combo1
        ds.Close
    Next i
    saveorder_Click
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "posttotrailers_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " posttotrailers_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub saveorder_Click()
    Dim cfile As String, i As Integer
    If Grid1.Rows < 2 Then Exit Sub
    If Form1.plantno = "50" Then
        cfile = Form1.srserv & "\wd\jobbing\jo" & List2 & Format(Combo1, "mmddyy") & ".txt"
    Else
        cfile = Form1.srserv & "\f\user\waredist\data\jobbing\jo" & List2 & Format(Combo1, "mmddyy") & ".txt"
    End If
    Open cfile For Output As #1
    For i = 1 To Grid1.Rows - 1
        Write #1, Grid1.TextMatrix(i, 0);       'Plant
        Write #1, Grid1.TextMatrix(i, 1);       'branch
        Write #1, Grid1.TextMatrix(i, 2);       'account
        Write #1, Grid1.TextMatrix(i, 3);       'sku
        Write #1, Grid1.TextMatrix(i, 5);       'total units
        Write #1, Grid1.TextMatrix(i, 6);       'pallets
        Write #1, Grid1.TextMatrix(i, 7);       'palsize
        Write #1, Grid1.TextMatrix(i, 8);       'wraps
        Write #1, Grid1.TextMatrix(i, 9);       'wrapsize
        Write #1, Grid1.TextMatrix(i, 10);      'units
        Write #1, Grid1.TextMatrix(i, 11);      'net
        Write #1, Grid1.TextMatrix(i, 12);      'group
        Write #1, Grid1.TextMatrix(i, 13);      'order date
        Write #1, Grid1.TextMatrix(i, 14)       'ship date
    Next i
    Close #1
End Sub
