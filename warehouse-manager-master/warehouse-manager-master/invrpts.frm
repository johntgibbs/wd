VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form invrpts 
   Caption         =   "Rack Inventory Reports"
   ClientHeight    =   6000
   ClientLeft      =   570
   ClientTop       =   2520
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   6000
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1815
      Left            =   0
      TabIndex        =   38
      Top             =   6120
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3201
      _Version        =   327680
   End
   Begin VB.Frame Frame3 
      Caption         =   "Crane Count Sheet Selections "
      Enabled         =   0   'False
      Height          =   2055
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   5775
      Begin VB.CheckBox Check3 
         Caption         =   "Left"
         Height          =   255
         Left            =   2760
         TabIndex        =   37
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Right"
         Height          =   255
         Left            =   1920
         TabIndex        =   36
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sort By Level"
         Height          =   255
         Left            =   3480
         TabIndex        =   35
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Lev 
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   34
         Text            =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Lev 
         Height          =   285
         Index           =   0
         Left            =   2640
         TabIndex        =   33
         Text            =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox bay 
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   32
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox bay 
         Height          =   285
         Index           =   0
         Left            =   2640
         TabIndex        =   31
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Print Crane Count Sheet"
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         TabIndex        =   30
         Top             =   1440
         Width           =   3975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "SR-3"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "SR-2"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "SR-1"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Thru"
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   29
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Thru"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   28
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Level"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   27
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Bay"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   26
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Additional Reports "
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   4920
      Width           =   5775
      Begin VB.CommandButton Command4 
         Caption         =   "SKU Totals - Whs"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Rack Reservations"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "BB Pallets"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rack Count Sheet Selections "
      Enabled         =   0   'False
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.CommandButton Command1 
         Caption         =   "Print Rack Count Sheet"
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   1920
         Width           =   3975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Aisles"
         Height          =   255
         Index           =   15
         Left            =   4200
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle E - H"
         Height          =   255
         Index           =   14
         Left            =   4200
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle A - D"
         Height          =   255
         Index           =   13
         Left            =   4200
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle I"
         Height          =   255
         Index           =   12
         Left            =   4200
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle H"
         Height          =   255
         Index           =   11
         Left            =   3000
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle G"
         Height          =   255
         Index           =   10
         Left            =   3000
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle F"
         Height          =   255
         Index           =   9
         Left            =   3000
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle E"
         Height          =   255
         Index           =   8
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle D"
         Height          =   255
         Index           =   7
         Left            =   1680
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle C"
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle B"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle A"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle 4"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle 3"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle 2"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aisle 1"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
End
Attribute VB_Name = "invrpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()                    'Rack Count Sheet
    Dim ds As ADODB.Recordset, s As String, w As String, i As Double
    Dim rh As String, rt As String, rf As String
    Dim tbb As Long, t4 As Long, PAisle
    tbb = 0: t4 = 0
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    s = "select r.aisle,r.rack,p.sku,r.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
    s = s & " from racks r, rackpos p"
    For i = 0 To 12
        If Option1(i) = True Then
            w = " where r.aisle = '" & Right(Option1(i).Caption, 1) & "'"
            rh = Option1(i).Caption
        End If
    Next i
    If LCase(Right(Form1.Caption, 7)) = "brenham" Then
        If Option1(12) = True Then
            w = " where r.aisle >= '1' And r.aisle <= '4'"
            rh = Option1(12).Caption
        End If
    End If
    If Option1(13) = True Then
        w = " where r.aisle >= 'A' And r.aisle <= 'D'"
        rh = Option1(13).Caption
    End If
    If Option1(14) = True Then
        w = " where r.aisle >= 'E' And r.aisle <= 'H'"
        rh = Option1(14).Caption
    End If
    If Option1(15) = True Then
        w = " where r.aisle >= '1' And r.aisle < 'M'"
        rh = Option1(15).Caption
    End If
    s = s & w & " and p.rackno = r.id"
    s = s & " and (p.count_qty > 0"
    s = s & " or r.qty + r.qty4 = 0)"
    s = s & " group by r.aisle,r.rack,p.sku,r.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
    s = s & " order by r.aisle,r.slot,p.sku desc"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        'PAisle = ds!aisle
        PAisle = ds(0)
        Do Until ds.EOF
            If ds(0) <> PAisle Then
                Grid1.AddItem "..."
                PAisle = ds(0)
            End If
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
            If ds(2) > "0" Then
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
            Else
                s = s & Chr(9) & Chr(9)
            End If
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

Private Sub Command2_Click()
    Form12.qstr = "BB Pallets"
    Form12.Show
End Sub

Private Sub Command3_Click()                    'Rack Reservations
    Dim ds As ADODB.Recordset, s As String
    Dim rt As String, rh As String, rf As String, i As Double
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    s = "select resv_lot,resv_sku,aisle,rack"
    s = s & " from racks"
    s = s & " where resv_sku > '.' or resv_lot > '.'"
    s = s & " order by resv_sku,resv_lot"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!resv_lot & Chr(9)
            s = s & ds!resv_sku & Chr(9)
            If ds!resv_sku > " " Then
                s = s & skurec(Val(ds!resv_sku)).prodname
            End If
            s = s & Chr(9)
            s = s & ds!aisle & " " & ds!rack
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^Lot #|^SKU|<Product|^Rack"
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3800
    Grid1.ColWidth(3) = 1800
    rt = "Rack Reservations"
    rh = Format(Now, "mmmm d, yyyy")
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

Private Sub Command4_Click()                'SKU Totals - Warehouse
    Dim ds As ADODB.Recordset, s As String, ss As ADODB.Recordset
    Dim rt As String, rh As String, rf As String, pc As Integer, i As Double
    Dim msku As String, p1 As Long, p2 As Long, p3 As Long, pt As Long
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 10
    s = "select * from sku_config"
    s = s & " where sku > '0'"
    s = s & " order by uom_type,description,sku"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            p1 = 0: p2 = 0: p3 = 0: pt = 0
            pc = ds!uom_per_pallet
            s = "select whse_num,sum(qty) from lane"
            s = s & " where sku = '" & ds!sku & "'"
            s = s & " group by whse_num"
            Set ss = Wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                Do Until ss.EOF
                    If ss!whse_num = 1 Then p1 = p1 + ss(1)
                    If ss!whse_num = 2 Then p2 = p2 + ss(1)
                    If ss!whse_num = 3 Then p3 = p3 + ss(1)
                    pt = pt + ss(1)
                    ss.MoveNext
                Loop
            End If
            ss.Close
            If pt > 0 Then
                s = ds!sku & Chr(9)
                s = s & ds!uom_type & " " & StrConv(ds!description, vbProperCase) & Chr(9)
                s = s & Format(p1, "#,###") & Chr(9)
                s = s & Format(p2, "#,###") & Chr(9)
                s = s & Format(p3, "#,###") & Chr(9)
                s = s & Format(pt, "#,##0") & Chr(9)
                s = s & Format(p1 * pc, "#,###") & Chr(9)
                s = s & Format(p2 * pc, "#,###") & Chr(9)
                s = s & Format(p3 * pc, "#,###") & Chr(9)
                s = s & Format(pt * pc, "#,##0")
                Grid1.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^SKU|<Product|^SR-1|^SR-2|^SR-3|^Total|^SR-1|^SR-2|^SR-3|^Total"
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 3600
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1000
    Screen.MousePointer = 0
    rt = "Crane SKU Totals"
    rh = Format(Now, "mmmm d, yyyy")
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

Private Sub Command5_Click()                'Crane Count Sheet
    Dim ds As ADODB.Recordset, s As String
    Dim rt As String, rh As String, rf As String, i As Double
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    s = "select sku,vert_loc,horz_loc,rack_side,zone_num,qty,lot_num"
    s = s & " from lane"
    If Option2(0) = True Then
        s = s & " where whse_num = 1"
        rt = "SR-1 Countsheet"
    End If
    If Option2(1) = True Then
        s = s & " where whse_num = 2"
        rt = "SR-2 Countsheet"
    End If
    If Option2(2) = True Then
        s = s & " where whse_num = 3"
        rt = "SR-3 Countsheet"
    End If
    s = s & " and horz_loc >= " & bay(0)
    s = s & " and horz_loc <= " & bay(1)
    s = s & " and vert_loc >= " & Lev(0)
    s = s & " and vert_loc <= " & Lev(1)
    If Check2 = Check3 Then
        x = 0
    Else
        If Check2 = 1 Then
            s = s & " and rack_side = 'R'"
        Else
            s = s & " and rack_side = 'L'"
        End If
    End If
    If Check1 = 1 Then
        s = s & " order by whse_num,vert_loc,horz_loc,rack_side"
    Else
        s = s & " order by whse_num,horz_loc,vert_loc,rack_side"
    End If
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!vert_loc & " " & ds!horz_loc & " " & ds!rack_side & Chr(9)
            s = s & ds!zone_num & Chr(9)
            s = s & ds!qty & Chr(9)
            If ds!sku > "0" Then s = s & ds!sku
            s = s & Chr(9)
            If ds!lot_num > "0" Then s = s & ds!lot_num
            s = s & Chr(9)
            If ds!sku > " " Then
                s = s & skurec(Val(ds!sku)).prodname
            End If
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    Grid1.FormatString = "^Bay|^Zone|^Qty|^SKU|^Lot #|<Product"
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 3000
    
    'rt = "Crane " & Whs & " Count Sheet"
    rh = Format(Now, "mmmm d, yyyy")
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

Private Sub Form_Deactivate()
    Dim i As Integer
    If invrpts.WindowState = 0 Then
        For i = 1 To Form1.Frmgrid.Rows - 1
            If Form1.Frmgrid.TextMatrix(i, 0) = "invrpts" Then
                Form1.Frmgrid.TextMatrix(i, 1) = invrpts.Top
                Form1.Frmgrid.TextMatrix(i, 2) = invrpts.Left
                Form1.Frmgrid.TextMatrix(i, 3) = invrpts.Height
                Form1.Frmgrid.TextMatrix(i, 4) = invrpts.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.Frmgrid.Rows - 1
        If Form1.Frmgrid.TextMatrix(i, 0) = "invrpts" Then
            invrpts.Top = Val(Form1.Frmgrid.TextMatrix(i, 1))
            invrpts.Left = Val(Form1.Frmgrid.TextMatrix(i, 2))
            invrpts.Height = Val(Form1.Frmgrid.TextMatrix(i, 3))
            invrpts.Width = Val(Form1.Frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    If Form1.edlane.Enabled = True Then
        Command5.Enabled = True
        Command4.Enabled = True
        Frame3.Enabled = True
    End If
    If Form1.edracks.Enabled = True Then
        Command1.Enabled = True
        Command2.Enabled = True
        Command3.Enabled = True
        Frame1.Enabled = True
    End If
    If LCase(Right(Form1.Caption, 5)) = "arrow" Then
        Option1(0).Caption = "Aisle A"
        Option1(1).Caption = "Aisle B"
        Option1(2).Caption = "Aisle C"
        Option1(3).Caption = "Aisle D"
        Option1(4).Caption = "Aisle E"
        Option1(5).Caption = "Aisle F"
        Option1(6).Caption = "Aisle G"
        Option1(7).Caption = "Aisle H"
        Option1(8).Caption = "Aisle I"
        Option1(9).Visible = False
        Option1(10).Visible = False
        Option1(11).Visible = False
        Option1(12).Visible = False
        Option1(13).Caption = "Aisle A-D"
        Option1(14).Caption = "Aisle E-H"
        Option1(15).Caption = "All Aisles"
    End If
    If LCase(Right(Form1.Caption, 9)) = "sylacauga" Then
        Option1(0).Caption = "Aisle A"
        Option1(1).Caption = "Aisle B"
        Option1(2).Caption = "Aisle C"
        Option1(3).Caption = "Aisle D"
        Option1(4).Caption = "Aisle E"
        Option1(5).Caption = "Aisle F"
        Option1(6).Caption = "Aisle G"
        Option1(7).Caption = "Aisle H"
        Option1(8).Caption = "Aisle I"
        Option1(9).Caption = "Aisle J"
        Option1(10).Caption = "Aisle K"
        Option1(11).Caption = "Aisle L"
        Option1(12).Caption = "Aisle M"
        Option1(13).Visible = False
        Option1(14).Visible = False
        Option1(15).Caption = "All Aisles"
    End If
    If LCase(Right(Form1.Caption, 7)) = "brenham" Then
        Option1(0).Caption = "Aisle 1"
        Option1(1).Caption = "Aisle 2"
        Option1(2).Caption = "Aisle 3"
        Option1(3).Caption = "Aisle 4"
        Option1(4).Caption = "Aisle A"
        Option1(5).Caption = "Aisle B"
        Option1(6).Caption = "Aisle C"
        Option1(7).Caption = "Aisle D"
        Option1(8).Caption = "Aisle E"
        Option1(9).Caption = "Aisle F"
        Option1(10).Caption = "Aisle G"
        Option1(11).Caption = "Aisle H"
        Option1(12).Caption = "Aisle 1-4"
        Option1(13).Caption = "Aisle A-D"
        Option1(14).Caption = "Aisle E-H"
        Option1(15).Caption = "All Aisles"
    End If
    Option1(15) = True
    Option2(0) = True
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

