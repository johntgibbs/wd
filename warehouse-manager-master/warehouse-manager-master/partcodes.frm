VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form18 
   Caption         =   "Partial Pallet Code Dates"
   ClientHeight    =   10095
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14865
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form18"
   ScaleHeight     =   10095
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   20
      Text            =   "Text3"
      Top             =   8040
      Width           =   2535
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Sort by SKU"
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
      Left            =   11400
      TabIndex        =   18
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
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
      Left            =   12480
      TabIndex        =   17
      Top             =   120
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   16
      Top             =   9600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   7200
      TabIndex        =   14
      Top             =   9120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mark as Picked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   12
      Top             =   9240
      Width           =   2415
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7080
      TabIndex        =   11
      Text            =   "Combo2"
      Top             =   8520
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   8520
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   8040
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   7455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   12091
      _Version        =   327680
      BackColorSel    =   32768
      FocusRect       =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Units"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   5400
      TabIndex        =   19
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label wconv 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   15
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Partials"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OP Lots:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   10
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   7560
      Width           =   9135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lot Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wraps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SKU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Marked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9600
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu mtc 
         Caption         =   "Mark Task Complete"
      End
      Begin VB.Menu mtp 
         Caption         =   "Mark Task Pending"
      End
      Begin VB.Menu mac 
         Caption         =   "Mark All Tasks - Complete"
      End
      Begin VB.Menu map 
         Caption         =   "Mark All Tasks - Pending"
      End
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function check_lot(plot As String) As Boolean
    Dim s As String, eflag As Boolean
    Dim maxYear As String
    maxYear = Trim(Right(Date, 2) + 2)
    eflag = True
    plot = UCase(plot)
    If Len(plot) = 8 Then
        s = Mid(plot, 1, 2)
        If s > "12" Or s < "01" Then eflag = False
        s = Mid(plot, 3, 2)
        If s > "31" Or s < "01" Then eflag = False
        s = Mid(plot, 5, 2)
        If s < "11" Or s > maxYear Then eflag = False
        s = Mid(plot, 8, 1)
        If s < "A" Or s > "Z" Then eflag = False
    End If
    
    If Len(plot) = 9 Then                                  'jv112015
        s = Mid(plot, 1, 2)
        If s > "12" Or s < "01" Then eflag = False
        s = Mid(plot, 3, 2)
        If s > "31" Or s < "01" Then eflag = False
        s = Mid(plot, 5, 2)
        If s < "11" Or s > maxYear Then eflag = False
        s = Mid(plot, 7, 3)                                 'jv112015
        If s < "100" Or s > "999" Then eflag = False        'jv112015
    End If
    check_lot = eflag
End Function

Private Sub refresh_lotcodes()
    Dim ds As ADODB.Recordset, s As String, plot As String, pcode As String
    Combo2.Clear: List2.Clear: plot = " ": pcode = " "
    s = "select lot1, barcode from pallets where sku = '" & Grid1.TextMatrix(Grid1.Row, 6) & "'"
    s = s & " order by lot1 desc, barcode"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!lot1 <> plot Or pcode <> Mid(ds!barcode, 11, 3) Then      'jv010516
                List2.AddItem ds!lot1
                Combo2.AddItem Trim(Mid(ds!barcode, 5, 9))                  'jv112015
                plot = ds!lot1
                pcode = Mid(ds!barcode, 11, 3)                              'jv010516
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Combo2.AddItem Grid1.TextMatrix(Grid1.Row, 11)
    s = Grid1.TextMatrix(Grid1.Row, 6) & " " & Grid1.TextMatrix(Grid1.Row, 11) & " 001"
    List2.AddItem barcode_to_lotnum(s)
    If Text2 > "0" Then
        For i = 0 To Combo2.ListCount - 1
            If Combo2.List(i) = Text2 Then
                Combo2.ListIndex = i
                Exit For
            End If
        Next i
    Else
        Combo2.ListIndex = 0
    End If
End Sub

Private Sub refresh_partials()
    Dim ds As ADODB.Recordset, s As String, ts As ADODB.Recordset
    Screen.MousePointer = 11
    Combo1.Clear: List1.Clear
    s = "select brname,palnum,shipdate,count(*) from picktasks"
    's = s & " where status = 'PEND'"
    s = s & " where status in ('PEND', 'PICKED')"
    s = s & " group by brname,palnum,shipdate"
    s = s & " order by shipdate,brname,palnum"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds(3) > 0 Then
                Combo1.AddItem ds!brname & " " & ds!palnum & " " & Format(ds!shipdate, "mm-dd-yyyy")
                's = "select * from picktasks where brname = '" & ds!brname & "'"
                s = "where brname = '" & fixquotes(ds!brname) & "'"
                s = s & " and shipdate = '" & Format(ds!shipdate, "mm-dd-yyyy") & "'"
                s = s & " and palnum = " & ds!palnum
                List1.AddItem s
            End If
            ds.MoveNext
        Loop
    Else
        Combo1.AddItem "...."
        List1.AddItem 0
    End If
    ds.Close
    Screen.MousePointer = 0
    If Combo1.ListCount > 1 Then Combo1.ListIndex = 0
End Sub

Private Sub refresh_picks()
    Dim ds As ADODB.Recordset, s As String, i As Integer, pdesc As String
    Dim wc As Integer
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 18: Grid1.Redraw = False
    Check1.Value = 1
    s = "select * from picktasks " & List1
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            wc = 1
            i = Val(ds!sku)
            If skurec(i).sku = ds!sku Then
                pdesc = skurec(i).prodname
                If skurec(i).uom_per_pallet > 0 And skurec(i).qty_per_pallet > 0 Then
                    wc = skurec(i).uom_per_pallet / skurec(i).qty_per_pallet
                End If
            Else
                pdesc = "------"
            End If
            s = ds!id & Chr(9)
            s = s & ds!branch & Chr(9)
            s = s & ds!brname & Chr(9)
            s = s & ds!shipdate & Chr(9)
            s = s & ds!palnum & Chr(9)
            s = s & ds!opseq & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & pdesc & Chr(9)
            s = s & ds!qty & Chr(9)
            s = s & ds!uom & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & Left(ds!lotnum, 9) & Chr(9)                 'jv010516
            s = s & ds!palletid & Chr(9)
            s = s & ds!status & Chr(9)
            s = s & ds!userid & Chr(9)
            s = s & ds!location & Chr(9)
            s = s & ds!reqid & Chr(9)
            s = s & wc
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    ycolor.Visible = False
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 13) <> "PEND" Or Grid1.TextMatrix(i, 14) > "." Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ycolor.BackColor
                ycolor.Visible = True
            End If
        Next i
        Grid1.Row = 1: Grid1.Col = 1
    End If
    's = "^ID|^Branch|<Name|^Date|^Tag #|^OPSeq|^SKU|<Product|^Qty|^UOM|^units|^Lot|^PalletID|^Status|^User|^Location|^ReqId"
    s = "^ID||||||^SKU|<Product|^Qty|^UOM|^Units|^Lot||^Status|^User||"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 0 '800
    Grid1.ColWidth(2) = 0 '3000
    Grid1.ColWidth(3) = 0 '1000
    Grid1.ColWidth(4) = 0 '800
    Grid1.ColWidth(5) = 0 '800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 3600
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 1000
    Grid1.ColWidth(11) = 1400
    Grid1.ColWidth(12) = 0 '800
    Grid1.ColWidth(13) = 1000
    Grid1.ColWidth(14) = 1000
    Grid1.ColWidth(15) = 0 '1400
    Grid1.ColWidth(16) = 0 '800
    Grid1.ColWidth(17) = 0 '800
    Check1.Value = 0
    If Grid1.Rows > 1 Then
        If Check2.Value = 1 Then
            Grid1.RowSel = Grid1.Row
            Grid1.Col = 6: Grid1.ColSel = 6
            Grid1.Sort = 5
        End If
        Grid1.Row = 1: Grid1.Col = 7
        Call Grid1_RowColChange
        For i = 0 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 13) = "PEND" Then
                Grid1.Row = i: Grid1.Col = 7
                Call Grid1_RowColChange
                Exit For
            End If
        Next i
    End If
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub Check2_Click()
    refresh_picks
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    If Left(List1, 5) = "where" Then refresh_picks
End Sub

Private Sub Combo2_Click()
    Text2 = Combo2
    List2.ListIndex = Combo2.ListIndex
End Sub

Private Sub Command1_Click()
    Dim p As ptask, i As Integer, ds As ADODB.Recordset, wdiff As String, pid As Long, udiff As String
    If Grid1.Row = 0 Then Exit Sub
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 0)) = 0 Then Exit Sub
    If check_lot(Text2) = False Then
        MsgBox "Invalid Code Date:  " & Text2, vbOKOnly + vbInformation, "sorry, try again...."
        Text2 = "..."
        Exit Sub
    End If
    wdiff = Val(Grid1.TextMatrix(i, 8)) - Val(Text1)
    udiff = Val(Grid1.TextMatrix(i, 10)) - Val(Text3)                           'jv122115
    If Val(udiff) > 0 Then                                                      'jv122115
        Text1 = Format((Val(Text3) / Val(wconv)) - 0.5, "0")                    'jv122115
        wdiff = Val(Grid1.TextMatrix(i, 8)) - Val(Text1)                        'jv122115
        If Val(wdiff) < Val(wconv) Then wdiff = 0                               'jv122115
    End If                                                                      'jv122115
    'If wdiff > 0 Then
    If Val(udiff) > 0 Then                                                      'jv122115
        s = "Update picktasks set qty = " & Val(Text1)
        's = s & ", units = " & Format(Val(Text1) * Val(wconv), "0")
        s = s & ", units = " & Val(Text3)                                       'jv122115
        s = s & ", lotnum = '" & Text2 & "'"
        s = s & ", status = 'PICKED'"
        s = s & ", userid = '" & WDUserId & "'"
        s = s & " Where id = " & Grid1.TextMatrix(i, 0)
    Else
        s = "Update picktasks set status = 'PICKED'"
        s = s & ", lotnum = '" & Text2 & "'"
        s = s & ", userid = '" & WDUserId & "'"
        s = s & " Where id = " & Grid1.TextMatrix(i, 0)
    End If
    'MsgBox s
    Wdb.Execute s
    s = "select * from picktasks where id = " & Grid1.TextMatrix(i, 0)
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        'If wdiff > 0 Then
        If Val(udiff) > 0 Then                                              'jv122115
            pid = wd_seq("PickTasks")
            s = "Insert into picktasks (id, branch, brname, shipdate, palnum, opseq, sku"
            s = s & ", lotnum, qty, uom, units, palletid, status, userid, location, reqid) Values ("
            s = s & pid
            s = s & ", " & ds!branch
            s = s & ", '" & fixquotes(ds!brname) & "'"
            s = s & ", '" & ds!shipdate & "'"
            s = s & ", " & ds!palnum
            s = s & ", " & ds!opseq
            s = s & ", '" & ds!sku & "'"
            s = s & ", '...'"
            s = s & ", " & wdiff
            s = s & ", 'Wraps'"
            's = s & ", " & Format(wdiff * Val(wconv), "0")
            s = s & ", " & udiff                                            'jv122115
            s = s & ", '" & ds!palletid & "'"
            s = s & ", 'PEND'"
            s = s & ", '.'"
            s = s & ", '" & ds!location & "'"
            s = s & ", '" & ds!reqid & "')"
            'MsgBox s
            Wdb.Execute s
        End If
        p.id = Grid1.TextMatrix(i, 0)
        p.area = "PICK ORDER"
        p.description = ds!branch & " " & ds!brname
        p.source = "OP-" & ds(5)
        p.target = ds!palletid
        p.product = Label2
        If Len(Text2) = 8 Then                                      'jv112015
            If Grid1.TextMatrix(i, 15) = "RE WORK" Then             'jv032916
                p.palletid = Left(Label2, 4) & Text2 & " REW"       'jv032916
            Else                                                    'jv032916
                p.palletid = Left(Label2, 4) & Text2 & " PAR"
            End If
        Else
            If Grid1.TextMatrix(i, 15) = "RE WORK" Then             'jv032916
                p.palletid = Left(Label2, 4) & Text2 & "REW"        'jv032916
            Else                                                    'jv032916
                p.palletid = Left(Label2, 4) & Text2 & "PAR"        'jv112015
            End If
        End If
        p.qty = ds!qty 'Text1
        p.uom = ds!uom '"Wraps"
        p.units = ds!units 'Grid1.TextMatrix(i, 8)
        'MsgBox "X" & p.palletid & "X Len=" & Len(p.palletid)
        p.lotnum = barcode_to_lotnum(p.palletid) 'List2
        p.units2 = "0"
        p.lotnum2 = ""
        p.status = "PICKED"
        p.userid = WDUserId
        p.trandate = Format(Now, "yyMMdd HH:mm:ss")
        p.reqid = ""
        Call post_pick_trans(p, Format(ds!shipdate, "MMddyyyy"))
    End If
    ds.Close
    refresh_picks
    DoEvents
    'Text2.SetFocus
    Combo2.SetFocus
End Sub

Private Sub Command2_Click()
    refresh_partials
End Sub

Private Sub Form_Load()
    refresh_partials
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid1_RowColChange()
    If Check1.Value = 1 Then Exit Sub
    mtc.Enabled = False: mtp.Enabled = False
    'mac.Enabled = False: map.Enabled = False
    If Grid1.Row = 0 Then Exit Sub
    'Label2 = Grid1.TextMatrix(Grid1.Row, 6) & " " & Grid1.TextMatrix(Grid1.Row, 7)
    Text1 = Val(Grid1.TextMatrix(Grid1.Row, 8))
    Text2 = Grid1.TextMatrix(Grid1.Row, 11)
    Text3 = Val(Grid1.TextMatrix(Grid1.Row, 10))                                    'jv122115
    wconv = Grid1.TextMatrix(Grid1.Row, 17)
    If Grid1.TextMatrix(Grid1.Row, 13) <> "COMP" Then mtc.Enabled = True
    If Grid1.TextMatrix(Grid1.Row, 13) <> "PEND" Then mtp.Enabled = True
    Label2 = Grid1.TextMatrix(Grid1.Row, 6) & " " & Grid1.TextMatrix(Grid1.Row, 7)
End Sub

Private Sub Label2_Change()
    refresh_lotcodes
End Sub

Private Sub mac_Click()
    Dim i As Long, s As String
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 0)) > 0 Then
            s = "Update picktasks set status = 'COMP', userid = ' ' Where id = " & Grid1.TextMatrix(i, 0)
            Wdb.Execute s
        End If
    Next i
    refresh_picks
    Grid1.Row = 1
End Sub

Private Sub map_Click()
    Dim i As Long, s As String
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 0)) > 0 Then
            s = "Update picktasks set status = 'PEND', userid = ' ' Where id = " & Grid1.TextMatrix(i, 0)
            Wdb.Execute s
        End If
    Next i
    refresh_picks
    Grid1.Row = 1
End Sub

Private Sub mtc_Click()
    Dim i As Long, s As String
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 0)) = 0 Then Exit Sub
    s = "Update picktasks set status = 'COMP', userid = ' ' Where id = " & Grid1.TextMatrix(i, 0)
    Wdb.Execute s
    refresh_picks
    Grid1.Row = i
End Sub

Private Sub mtp_Click()
    Dim i As Long, s As String
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 0)) = 0 Then Exit Sub
    s = "Update picktasks set status = 'PEND', userid = ' ' Where id = " & Grid1.TextMatrix(i, 0)
    Wdb.Execute s
    refresh_picks
    Grid1.Row = i
End Sub

Private Sub Text2_Change()
    If check_lot(Text2) = False Then
        Command1.Enabled = False
    Else
        If Grid1.TextMatrix(Grid1.Row, 13) = "PICKED" Then
            Command1.Enabled = False
        Else
            Command1.Enabled = True
        End If
    End If
End Sub
