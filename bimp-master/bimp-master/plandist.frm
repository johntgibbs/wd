VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form plandist 
   Caption         =   "Planned Distribution"
   ClientHeight    =   11265
   ClientLeft      =   3495
   ClientTop       =   2655
   ClientWidth     =   13845
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11265
   ScaleWidth      =   13845
   Begin MSFlexGridLib.MSFlexGrid hgrid 
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   11280
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2415
      Left            =   0
      TabIndex        =   9
      Top             =   8160
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4260
      _Version        =   327680
      BackColorFixed  =   12648384
      BackColorSel    =   12583104
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   6360
      TabIndex        =   8
      Text            =   "Combo3"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3840
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   12938
      _Version        =   327680
      Cols            =   8
      FixedCols       =   4
      BackColorFixed  =   14737632
      BackColorSel    =   0
      WordWrap        =   -1  'True
      FocusRect       =   0
      GridLines       =   2
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   10920
      TabIndex        =   18
      Top             =   7920
      Width           =   3975
   End
   Begin VB.Label nend 
      Caption         =   "nend"
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   10800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label nstart 
      Caption         =   "nstart"
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   10800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label cend 
      Caption         =   "cend"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   10800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label cstart 
      Caption         =   "cstart"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   10800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Surplus"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12360
      TabIndex        =   13
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Month Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12360
      TabIndex        =   12
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "< 2 Week Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10560
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Transport Schedule"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   7920
      Width           =   13695
   End
   Begin VB.Label Label3 
      Caption         =   "SKU:"
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Plant:"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Branch:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2 Week Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10560
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
      Begin VB.Menu prtgrd 
         Caption         =   "Grid Listing"
      End
   End
   Begin VB.Menu impmenu 
      Caption         =   "Import"
      Begin VB.Menu proclr 
         Caption         =   "Process Last Receipt"
      End
      Begin VB.Menu procdisc 
         Caption         =   "Process Discontinued"
      End
      Begin VB.Menu impbranch 
         Caption         =   "Process Branch Whs"
      End
      Begin VB.Menu impr12 
         Caption         =   "R12 OnHand Qtys - All Whs"
      End
      Begin VB.Menu impsales 
         Caption         =   "Sales - All Whs"
      End
      Begin VB.Menu csrteloads 
         Caption         =   "Countsheet Route Loads"
      End
      Begin VB.Menu stkmenu 
         Caption         =   "Stock History"
         Begin VB.Menu stkimport 
            Caption         =   "Import R12"
         End
         Begin VB.Menu stkpost 
            Caption         =   "Post Stock History"
         End
      End
      Begin VB.Menu imptbcs 
         Caption         =   "Ticket BarCodes"
      End
   End
   Begin VB.Menu confmenu 
      Caption         =   "Configure"
      Begin VB.Menu edpq 
         Caption         =   "Product Quotas"
      End
      Begin VB.Menu edpbs 
         Caption         =   "Plant Branches"
      End
      Begin VB.Menu edps 
         Caption         =   "Plant SKUs"
      End
      Begin VB.Menu edbs 
         Caption         =   "Branch SKUs"
      End
      Begin VB.Menu hublists 
         Caption         =   "3 Gallon Hubs"
      End
   End
   Begin VB.Menu prodmenu 
      Caption         =   "Product Info"
      Begin VB.Menu pbts 
         Caption         =   "Production Batch Tickets"
      End
      Begin VB.Menu batrels 
         Caption         =   "Batch Releases"
      End
      Begin VB.Menu plantproddates 
         Caption         =   "Plant Production Dates"
      End
      Begin VB.Menu newrelease 
         Caption         =   "New Product Releases"
      End
      Begin VB.Menu missbimp 
         Caption         =   "Missing BIMP SKU Report"
      End
      Begin VB.Menu bimpdels 
         Caption         =   "BIMP Deletion Log"
      End
   End
   Begin VB.Menu trlmenu 
      Caption         =   "Trailers"
      Begin VB.Menu brruns 
         Caption         =   "Active Trailers"
      End
      Begin VB.Menu planships 
         Caption         =   "Planned Shipments"
      End
      Begin VB.Menu planvsact 
         Caption         =   "Planned vs Actual"
      End
      Begin VB.Menu expplanttrailers 
         Caption         =   "Send Trailers to Plant"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu servmenu 
      Caption         =   "Server Status"
   End
End
Attribute VB_Name = "plandist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub export_greenville()
    Dim i As Integer, j As Long
    Dim rt As String, rf As String, rh As String, webfile As String
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    Screen.MousePointer = 11
    Call refresh_tstation_skus("062")
    Call read_countsheet("\\BBC-03-FILESVR\SharedGroups\WD\html\counts\tstation.062")
    'Call read_countsheet("S:\wd\html\counts\tstation.062")
    For i = 1 To hgrid.Rows - 1
        If hgrid.TextMatrix(i, 2) > " " Then
            Call refresh_trans("062", "20", hgrid.TextMatrix(i, 0), hgrid.TextMatrix(i, 2), i)
            DoEvents
        End If
        j = Val(hgrid.TextMatrix(i, 3))
        j = j + Val(hgrid.TextMatrix(i, 4))
        j = j + Val(hgrid.TextMatrix(i, 5))
        hgrid.TextMatrix(i, 6) = Format(j, "#")
    Next i
    
    
    webfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\counts\tstation062.htm"
    rt = "Greenville Transfer Station"
    rh = "Warehouse:  Greenville"
    rf = "Posted: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "seagreen": gndc(0) = Me.hgrid.BackColorFixed
    hgrid.Redraw = False
    Call htmlcolorgrid(Me, webfile, hgrid, rt, rh, rf, "linen", "cyan", "white")
    hgrid.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_tstation_skus(pbranch As String)
    Dim ds As ADODB.Recordset, s As String, pplant As String
    hgrid.Redraw = False
    hgrid.FontBold = True
    hgrid.FontName = "Arial"
    hgrid.Clear: hgrid.Rows = 1: hgrid.Cols = 7: hgrid.FixedCols = 2
    pplant = branchrec(Val(pbranch)).supplier
    s = "select sku from bimp where branchwhs = '" & pbranch & "' and plantwhs = '" & pplant & "' order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            hgrid.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    hgrid.FormatString = "^SKU|<Product|^Count Date|^Beg Inventory|^Trans In|^Trans Out|^Net"
    hgrid.ColWidth(0) = 800
    hgrid.ColWidth(1) = 4000
    hgrid.ColWidth(2) = 2000
    hgrid.ColWidth(3) = 2000
    hgrid.ColWidth(4) = 2000
    hgrid.ColWidth(5) = 2000
    hgrid.ColWidth(6) = 2000
    hgrid.Redraw = True
End Sub

Sub refresh_trans(pbranch As String, proute As String, psku As String, pdate As String, prow As Integer)
    Dim ds As ADODB.Recordset, q As String, s As String
    Dim i As Integer, t As Long, j As Long
    
    q = "select tran_qty from bolinf.inv_adj_input_detail"
    q = q & " Where tran_type = '1'"
    q = q & " and tran_date >= TO_DATE('" & Format(pdate, "dd-mmm-yy") & "')"
    'q = q & " and tran_date < TO_DATE('17-Jun-18')"                'build countsheet
    q = q & " and branch_no = '" & pbranch & "'"
    q = q & " and product_no = '" & psku & "'"
    q = q & " and route_no = '" & proute & "'"
        
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!tran_qty > 0 Then
                hgrid.TextMatrix(prow, 4) = Val(hgrid.TextMatrix(prow, 4)) + ds!tran_qty
            Else
                hgrid.TextMatrix(prow, 5) = Val(hgrid.TextMatrix(prow, 5)) + ds!tran_qty
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Private Sub read_countsheet(cfile As String)
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim i As Integer, pflag As Boolean
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7
        pflag = False
        For i = 1 To hgrid.Rows - 1
            If hgrid.TextMatrix(i, 0) = f1 Then
                hgrid.TextMatrix(i, 2) = f7
                hgrid.TextMatrix(i, 3) = f6
                pflag = True
                Exit For
            End If
        Next i
        If pflag = False Then
            s = Trim(f1) & Chr(9)
            s = s & UCase(f2) & Chr(9)
            s = s & f7 & Chr(9)
            s = s & f6
            hgrid.AddItem s
            'MsgBox s
        End If
    Loop
    Close #1
    hgrid.Row = 1: hgrid.RowSel = 1
    hgrid.Col = 0: hgrid.ColSel = 0
    hgrid.Sort = 5
End Sub

Private Sub refresh_branches()
    Dim i As Integer
    Combo1.Clear
    Combo1.AddItem "All-All Branches"
    For i = 1 To 99
        If branchrec(i).oraloc > " " Then Combo1.AddItem Format(branchrec(i).branchno, "000") & "-" & branchrec(i).branchname
    Next i
End Sub

Private Sub refresh_skus()
    Dim ds As ADODB.Recordset, s As String
    Combo3.Clear
    Combo3.AddItem "ALL"
    s = "select distinct sku from bimp order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo3.AddItem ds!sku
            ds.MoveNext
        Loop
    End If
    ds.Close
    Combo3.ListIndex = 0
End Sub

Private Sub refresh_grid1()
    Dim s As String, i As Integer, ds As ADODB.Recordset
    Dim tt As Integer, nt As Integer, np As Integer, cws As Integer, nws As Integer, qpt As Single
    Label5.Caption = bimp_status_time                                                        'jv022316
    If Label5.Caption > " " Then Label5.Caption = "Last R12 import @ " & Label5.Caption      'jv022316
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    'Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 17
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 18                    'jv021116
    s = "select id, plantwhs, branchwhs, sku, onhand, onorder, sales, plantpool, quotapct, thiswknewpals,"
    's = s & " nextwknewpals, poolsched, bimpstatus from bimp where sku > ' '"
    s = s & " nextwknewpals, poolsched, bimpstatus, ohpct from bimp where sku > ' '"    'jv021116
    's = "select * from bimp where sku > ' '"
    If Combo2 <> "ALL" Then s = s & " and plantwhs = '" & Combo2 & "'"
    If Combo3 <> "ALL" Then s = s & " and sku = '" & Combo3 & "'"
    If Left(Combo1, 3) <> "All" Then s = s & " and branchwhs = '" & Left(Combo1, 3) & "'"
    s = s & " and plantwhs <> 'DRY'"
    s = s & " order by sku, branchwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!plantwhs & Chr(9)
            s = s & ds!branchwhs & "-" & branchrec(Val(ds!branchwhs)).branchname & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!onhand & Chr(9)
            s = s & ds!onorder & Chr(9)
            s = s & ds!sales & Chr(9)
            s = s & ds!plantpool & Chr(9)
            s = s & Format(ds!quotapct, "0.000") & Chr(9)
            s = s & Format((ds!quotapct / 100) * ds!plantpool, "0") & Chr(9)
            s = s & Chr(9)
            s = s & Format(ds!thiswknewpals, "#") & Chr(9)
            s = s & Format(ds!nextwknewpals, "#") & Chr(9)
            s = s & Format((ds!quotapct / 100) * ds!poolsched, "0") & Chr(9)
            s = s & Chr(9)
            s = s & ds!bimpstatus & Chr(9)
            s = s & ds!id
            If ds!ohpct > 0 Then                                'jv021116
                s = s & Chr(9) & Format(ds!ohpct * 30, "0")     'jv021116
            End If                                              'jv021116
            'If ds!ohpct > 0 And ds!ohpct < 0.5 Then
            '    s = s & "W"
            'Else
            '    If ds!paldiff = 0 Then
            '        s = s & "B"
            '    Else
            '        If ds!paldiff > 0 Then
            '            s = s & "G"
            '        Else
            '            s = s & "Y"
            '        End If
            '    End If
            'End If
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        tt = 0: nt = 0: np = 0: cws = 0: nws = 0: qpt = 0
        For i = 1 To Grid1.Rows - 1
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 2:: Grid1.ColSel = 6
            If Grid1.TextMatrix(i, 15) = "W" Then Grid1.CellBackColor = wcolor.BackColor
            If Grid1.TextMatrix(i, 15) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
            If Grid1.TextMatrix(i, 15) = "B" Then Grid1.CellBackColor = bcolor.BackColor
            If Grid1.TextMatrix(i, 15) = "G" Then Grid1.CellBackColor = gcolor.BackColor
            If skurec(Val(Grid1.TextMatrix(i, 2))).pallet > 0 Then
                Grid1.TextMatrix(i, 10) = Format(Val(Grid1.TextMatrix(i, 9)) / skurec(Val(Grid1.TextMatrix(i, 2))).pallet, "0")
            Else
                Grid1.TextMatrix(i, 10) = "0"
            End If
            Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 10)) + Val(Grid1.TextMatrix(i, 13))  'Pallets + New Pool
            Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 14)) - Val(Grid1.TextMatrix(i, 11))   'less curr week pallets
            Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 14)) - Val(Grid1.TextMatrix(i, 12))   'less next week pallets
            tt = tt + Val(Grid1.TextMatrix(i, 10))
            cws = cws + Val(Grid1.TextMatrix(i, 11))
            nws = nws + Val(Grid1.TextMatrix(i, 12))
            np = np + Val(Grid1.TextMatrix(i, 13))
            nt = nt + Val(Grid1.TextMatrix(i, 14))
            qpt = qpt + Val(Grid1.TextMatrix(i, 8))             'jv120215
        Next i
        s = " " & Chr(9)
        s = s & " " & Chr(9)
        s = s & " " & Chr(9)
        s = s & "Totals" & Chr(9)
        s = s & " " & Chr(9)
        s = s & " " & Chr(9)
        s = s & " " & Chr(9)
        s = s & " " & Chr(9)
        If Combo3 = "ALL" Then                                  'jv120215
            s = s & " " & Chr(9)
        Else
            s = s & Format(qpt, "0.000") & Chr(9)               'jv120215
        End If
        s = s & " " & Chr(9)
        s = s & tt & Chr(9) & cws & Chr(9) & nws & Chr(9) & np & Chr(9) & nt
        Grid1.AddItem s
        Grid1.Row = 1
        Grid1.RowHeight(0) = Grid1.RowHeight(1) * 2.5
    End If
    s = "^Plant|<Branch|^SKU|<Product|^On Hand Units|^On Order Units|^Sales|^Plant Units|^Quota %|^Quota Units|^Quota Pallets|^This Week Orders|^Next Week Orders|^New Pool Pallets|^Transport"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    If UCase(Left(Combo1, 3)) = "ALL" Then
        Grid1.ColWidth(1) = 2000
    Else
        Grid1.ColWidth(1) = 0
    End If
    If Combo3 = "ALL" Then
        Grid1.ColWidth(2) = 600
    Else
        Grid1.ColWidth(2) = 0
    End If
    Grid1.ColWidth(3) = 3000
    Grid1.ColWidth(4) = 1100
    Grid1.ColWidth(5) = 1100
    Grid1.ColWidth(6) = 1100
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1100
    Grid1.ColWidth(9) = 1100
    Grid1.ColWidth(10) = 1400
    Grid1.ColWidth(11) = 1100
    Grid1.ColWidth(12) = 1100
    Grid1.ColWidth(13) = 1100
    Grid1.ColWidth(14) = 1100
    Grid1.ColWidth(15) = 0 '1100
    Grid1.ColWidth(16) = 0 '1000
    Grid1.ColWidth(17) = 0 '1000
    Grid1.Redraw = True
    refresh_grid2
End Sub

Private Sub refresh_grid2()
    Dim ss As ADODB.Recordset, s As String, cp As Integer, np As Integer
    Dim cs As String, ce As String, ns As String, ne As String, cc As Integer, nc As Integer
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 10
    s = "select * from truckwo where wodate >= '" & Format(Now, "yyyy-MM-dd") & "'"
    s = s & " and wodate <= '" & Format(nend, "yyyy-MM-dd") & "'"
    's = s & " and r12ticket > '0' and wostatus not in ('CANC', 'CLOSE')"
    s = s & " and wtype in ('Start', 'SameDay') and wostatus not in ('CANC', 'CLOSE')"
    If Left(Combo1, 3) <> "All" Then s = s & " and destination = '" & Left(Combo1, 3) & "'"
    If Combo2 = "ALL" Then
        s = s & " and origin in ('A10', 'K10', 'T10')"
    Else
        s = s & " and origin = '" & Combo2 & "'"
    End If
    s = s & " order by wodate, startime"
    Set ss = tsb.Execute(s)
    'MsgBox s
    If ss.BOF = False Then
        ss.MoveFirst
        Do Until ss.EOF
            If Val(ss!destination) > 0 And Val(ss!destination) < 100 Then
                s = ss!wonum & Chr(9)
                s = s & ss!r12ticket & Chr(9)
                s = s & ss!origin & Chr(9)
                s = s & ss!destination & Chr(9)
                s = s & ss!wodate & Chr(9)
                s = s & ss!startime & Chr(9)
                's = s & ss!Description & Chr(9)
                s = s & branchrec(Val(ss!destination)).branchname & Chr(9)
                s = s & ss!trlno & Chr(9)
                s = s & ss!wostatus & Chr(9)
                s = s & ss!trlsize
                Grid2.AddItem s
            End If
            ss.MoveNext
        Loop
    End If
    ss.Close
    cp = 0: np = 0: cc = 0: nc = 0
    cs = Format(cstart, "yyyyMMdd")
    ce = Format(cend, "yyyyMMdd")
    ns = Format(nstart, "yyyyMMdd")
    ne = Format(nend, "yyyyMMdd")
    
    If Grid2.Rows > 1 Then
        For i = 1 To Grid2.Rows - 1
            s = Format(Grid2.TextMatrix(i, 4), "yyyyMMdd")
            If s >= cs And s <= ce Then
                cp = cp + Val(Grid2.TextMatrix(i, 9))
                cc = cc + 1
            End If
            If s >= ns And s <= ne Then
                np = np + Val(Grid2.TextMatrix(i, 9))
                nc = nc + 1
            End If
        Next i
    End If
            
    's = "^WONum|^Ticket|^Origin|^Destination|^Date|^Start|<Note|^#|^Status|^Pallets"
    s = "^WONum|^Ticket|^Origin|^Destination|^Date|^Start|<Location|^#|^Status|^Pallets"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 1200
    Grid2.ColWidth(4) = 1200
    Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 2000 '6000
    Grid2.ColWidth(7) = 600
    Grid2.ColWidth(8) = 1000
    Grid2.ColWidth(9) = 800
    's = "Transport Schedule: " & Grid2.Rows - 1 & " Trailer(s) "
    s = "Transport Schedule: "
    If cp > 0 Then s = s & "      This Week: " & cc & " Trailer(s) " & cp & " Pallets"
    If np > 0 Then s = s & "      Next Week: " & nc & " Trailer(s) " & np & " Pallets"
    Label4.Caption = s '"Transport Schedule: " & Grid2.Rows - 1
    Grid2.Redraw = True
End Sub

Private Sub batrels_Click()
    batchreleases.Show
End Sub

Private Sub bimpdels_Click()
    Dim rt As String, rh As String, rf As String
    Dim cfile As String, s As String, f0 As String, f1 As String, f2 As String, f3 As String
    cfile = "\\BBC-01-PRODTRK\wd\data\bimpdels.csv"
    bimpbanner.pgrid.Clear: bimpbanner.pgrid.Rows = 1: bimpbanner.pgrid.Cols = 4
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3
        s = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3
        bimpbanner.pgrid.AddItem s
    Loop
    Close #1
    bimpbanner.pgrid.FormatString = "^Date/Time|^User|<Application|<SQL"
    bimpbanner.pgrid.ColWidth(0) = 1600
    bimpbanner.pgrid.ColWidth(1) = 1200
    bimpbanner.pgrid.ColWidth(2) = 2000
    bimpbanner.pgrid.ColWidth(3) = 4000
    Screen.MousePointer = 0
    rt = "BIMP SKU Deletions"
    rh = " "
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    'htdc(0) = "cyan": gndc(0) = Me.Grid1.BackColorFixed
    'htdc(1) = "yellow": gndc(1) = Me.Grid1.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid1.BackColor
    bimpbanner.pgrid.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(bimpbanner, "c:\htmlgrid.htm", bimpbanner.pgrid, rt, rh, rf, "linen", "khaki", "white")
        bimpbanner.pgrid.Redraw = True
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(bimpbanner, "c:\htmlgrid.htm", bimpbanner.pgrid, rt, rh, rf, "linen", "khaki", "white")
        bimpbanner.pgrid.Redraw = True
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub brruns_Click()
    branchruns.Show
End Sub

Private Sub Combo1_Click()
    refresh_grid1
End Sub

Private Sub Combo2_Click()
    refresh_grid1
End Sub

Private Sub Combo3_Click()
    refresh_grid1
End Sub

Private Sub Command1_Click()
    refresh_grid1
End Sub

Private Sub csrteloads_Click()
    'Call import_r12_route_loads
    bimprtloads.Show
End Sub

Private Sub edbs_Click()
    branchdsku.Show
End Sub

Private Sub edpbs_Click()
    plantdbranch.Show
End Sub

Private Sub edpq_Click()
    prodquotas.Show
End Sub

Private Sub edps_Click()
    plantdsku.Show
End Sub

Private Sub expplanttrailers_Click()
    exptrailers.Show
End Sub

Private Sub Form_Load()
    cstart = Format(Now, "ddd")
    If cstart = "Sun" Then pday = 0
    If cstart = "Mon" Then pday = -1
    If cstart = "Tue" Then pday = -2
    If cstart = "Wed" Then pday = -3
    If cstart = "Thu" Then pday = -4
    If cstart = "Fri" Then pday = -5
    If cstart = "Sat" Then pday = -6
    cstart = Format(DateAdd("d", pday, Now), "M-dd-yyyy")
    cend = Format(DateAdd("d", 6, cstart), "M-dd-yyyy")
    nstart = Format(DateAdd("d", 1, cend), "M-dd-yyyy")
    nend = Format(DateAdd("d", 6, nstart), "M-dd-yyyy")
    'cstart = Format(cstart, "yyyyMMdd")
    'cend = Format(cend, "yyyyMMdd")
    'nstart = Format(nstart, "yyyyMMdd")
    'nend = Format(nend, "yyyyMMdd")
    refresh_branches
    refresh_skus
    Combo2.Clear
    Combo2.AddItem "A10"
    Combo2.AddItem "K10"
    Combo2.AddItem "T10"
    Combo2.AddItem "VENDOR"
    'Combo2.AddItem "DRY"
    Combo2.AddItem "ALL"
    Combo2.ListIndex = 0
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    Combo1.ListIndex = 1
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    Grid2.Width = Me.Width - 180
    Label4.Width = Me.Width - 300
    hgrid.Width = Me.Width - 180                              'jv070618
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim i As Integer, cp As Long, np As Long, pp As Long
    If Val(Grid1.TextMatrix(Grid1.Row, 16)) = 0 Then Exit Sub
    'If Grid1.Col = 11 Or Grid1.Col = 12 Or Grid1.Col = 13 Then
    If Grid1.Col = 11 Or Grid1.Col = 12 Then
        If edcol = True Then
            Grid1.Text = ""
            edcol = False
        End If
        If KeyAscii = 8 Then
            If Len(Grid1.Text) > 1 Then
                Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
            Else
                Grid1.Text = ""
            End If
        End If
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            Grid1.Text = Grid1.Text & Chr(KeyAscii)
        End If
        If Grid1.Col = 11 Then
            i = Grid1.Row
            Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 10)) + Val(Grid1.TextMatrix(i, 13))  'Pallets + New Pool
            Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 14)) - Val(Grid1.TextMatrix(i, 11))   'less curr week pallets
            Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 14)) - Val(Grid1.TextMatrix(i, 12))   'less next week pallets
            s = "Update bimp set thiswknewpals = " & Val(Grid1.Text) & " where id = " & Grid1.TextMatrix(Grid1.Row, 16)
            'MsgBox s
            wdb.Execute s
        End If
        If Grid1.Col = 12 Then
            i = Grid1.Row
            Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 10)) + Val(Grid1.TextMatrix(i, 13))  'Pallets + New Pool
            Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 14)) - Val(Grid1.TextMatrix(i, 11))   'less curr week pallets
            Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 14)) - Val(Grid1.TextMatrix(i, 12))   'less next week pallets
            s = "Update bimp set nextwknewpals = " & Val(Grid1.Text) & " where id = " & Grid1.TextMatrix(Grid1.Row, 16)
            'MsgBox s
            wdb.Execute s
        End If
        'If Grid1.Col = 13 Then
        '    s = "Update bimp set poolsched = " & val(Grid1.Text) & " where id = " & Grid1.TextMatrix(Grid1.Row, 16)
        '    MsgBox s
        '    'wdb.Execute s
        'End If
        cp = 0: np = 0: pp = 0
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 16)) > 0 Then
                cp = cp + Val(Grid1.TextMatrix(i, 11))
                np = np + Val(Grid1.TextMatrix(i, 12))
                'pp = pp + Val(Grid1.TextMatrix(i, 13))
            End If
        Next i
        Grid1.TextMatrix(Grid1.Rows - 1, 11) = Format(cp, "#")
        Grid1.TextMatrix(Grid1.Rows - 1, 12) = Format(np, "#")
        'Grid1.TextMatrix(Grid1.Rows - 1, 13) = Format(pp, "#")
    End If
End Sub

Private Sub Grid1_RowColChange()
    Dim i As Integer, s As String, c As Integer
    Grid1.ToolTipText = ""                  'jv021116
    'If Val(Grid1.TextMatrix(Grid1.Row, 17)) > 0 Then
    '    Grid1.ToolTipText = "  " & Grid1.TextMatrix(Grid1.Row, 17) & " days supply. "
    'End If
    i = Grid1.Row: c = Grid1.Col
    If c = 4 Or c = 5 Or c = 6 Or c = 7 Then            'Unit Columns
        If Val(Grid1.TextMatrix(i, c)) <> 0 Then
            s = Format(Val(Grid1.TextMatrix(i, c)) / skurec(Val(Grid1.TextMatrix(i, 2))).pallet, "0")
            If Val(s) > 1 Then
                Grid1.ToolTipText = "  " & s & " Pallets.."
            Else
                Grid1.ToolTipText = "  " & s & " Pallet.."
            End If
        End If
    Else
        If Val(Grid1.TextMatrix(Grid1.Row, 17)) > 0 Then
            Grid1.ToolTipText = "  " & Grid1.TextMatrix(Grid1.Row, 17) & " days supply. "
        End If
    End If
        
            
    'If Grid1.Col = 4 Then
    '    If Val(Grid1.TextMatrix(Grid1.Row, 4)) <> 0 Then
    '        s = Format(Val(Grid1.TextMatrix(Grid1.Row, 4)) / skurec(Val(Grid1.TextMatrix(Grid1.Row, 2))).pallet, "0")
    '        If Val(s) > 1 Then
    '            Grid1.ToolTipText = "  " & s & " Pallets.."
    '        Else
    '            Grid1.ToolTipText = "  " & s & " Pallet.."
    '        End If
    '    End If
    'End If
End Sub

Private Sub hublists_Click()
    bimpvallists.vkey = "hubnames"
    bimpvallists.Show
    DoEvents
    'bimpvallists.Combo1.AddItem "hubnames"
    bimpvallists.Combo1.AddItem "hubbranches"
    bimpvallists.Combo1.AddItem "hubskus"
End Sub

Private Sub impbranch_Click()
    Dim s As String, i As Integer
    s = InputBox("Branch Warehouse Code:", "Branch Warehouse Code......", "001")
    If Len(s) = 0 Then Exit Sub
    If Val(s) < 1 Or Val(s) > 99 Then Exit Sub
    i = Val(s)
    s = branchrec(i).oraloc
    If s > "000" Then
        s = Format(Val(i), "000")
        Call import_r12_branch_qty(s)
        Call import_r12_branch_sales(s)
        imptbcs_Click                                       'jv021417
        MsgBox "Warehouse: " & s & " has been updated."
    Else
        MsgBox "Invalid Branch Code!", vbOKOnly + vbExclamation, "sorry, try again...."
    End If
End Sub

Private Sub impr12_Click()
    import_r12_qtys
    DoEvents                        'jv062416
    refresh_grid1                   'jv062416
    DoEvents
    browserpage.refkey = Val(browserpage.refkey) + 1        'jv031116
    imptbcs_Click                                           'jv021417
End Sub

Private Sub impsales_Click()
    import_r12_sales
    'export_greenville                                       'jv070617
    tstations.dtrig = "Greenville" 'Val(dtrig) + 1          'jv092118
End Sub

Private Sub imptbcs_Click()                                 'jv021417
    export_branchbarcodes_ships
    'export_branchbarcodes_bills
End Sub

Private Sub missbimp_Click()
    bimp_missing_skus
End Sub

Private Sub newrelease_Click()
    skurelease.Show
End Sub

Private Sub pbts_Click()
    prodbatches.Show
End Sub

Private Sub planships_Click()
    planskuship.Show
End Sub

Private Sub plantproddates_Click()
    bimpvallists.vkey = "bimpproddates"
    bimpvallists.Show
End Sub

Private Sub planvsact_Click()
    branchtrailers.Show
End Sub

Private Sub procdisc_Click()
    process_bimp_discontinued
End Sub

Private Sub proclr_Click()
    process_bimp_lastreceipt
    'process_bimp_lastissue                              'jv032118
End Sub

Private Sub prtgrd_Click()
    Dim rt As String, rf As String, rh As String
    Dim i As Integer
    On Error Resume Next
    rt = Me.Caption
    If Combo2 = "T10" Then rh = "Brenham --> "
    If Combo2 = "K10" Then rh = "Broken Arrow --> "
    If Combo2 = "A10" Then rh = "Sylacauga --> "
    rh = rh & "Branch: " & Combo1 & " SKU: " & Combo3
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    'htdc(0) = "Yellow": gndc(0) = Grid1.BackColorFixed
    'htdc(0) = "Pink": gndc(0) = Grid1.BackColorFixed
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Grid1.Redraw = False
        Call htmlcolorgrid(Me, "c:\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        Grid1.Redraw = True
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
End Sub

Private Sub servmenu_Click()
    'tstations.dtrig = "Greenville" 'Val(dtrig) + 1
    'MsgBox "ready"
    'Exit Sub

    bimpvallists.vkey.Caption = "wdserverstatus"
    bimpvallists.Show
End Sub

Private Sub stkimport_Click()
    skuoutsk.Show
End Sub

Private Sub stkpost_Click()
    bimpstkhist.Show
End Sub
