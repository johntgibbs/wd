VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form tktonhand 
   Caption         =   "Batch Inventory"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12135
   LinkTopic       =   "Form23"
   ScaleHeight     =   9585
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Print List"
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
      Left            =   10200
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   15690
      _Version        =   327680
      ForeColor       =   12582912
      BackColorFixed  =   16777152
   End
   Begin VB.Label dupflag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Duplicate BarCodes Found"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label bunits 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "bunits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label bproduct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label bbarcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "bbarcode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WMS Units"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BarCode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "tktonhand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, s As String, i As Integer, wdlot As String, wdsku As String
    Dim cb As ADODB.Connection, cs As ADODB.Recordset, t As String
    Dim aqty As Long, hqty As Long
    Screen.MousePointer = 11
    dupflag.Visible = False
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 10
    bunits = "0"
    wdlot = barcode_to_lotnum(bbarcode & "EOR")
    wdlot = wdlot & Right(bbarcode, 3)
    wdsku = Trim(Left(bbarcode, 4))                     'jv051917
    
    If Form1.plantno = "50" Then                'Cranes
        s = "select l.whse_num, l.zone_num, l.vert_loc, l.horz_loc, l.rack_side, p.posn_num, p.barcode,"
        's = s & " p.lot_num, p.count_qty, p.lot2, p.qty2"
        s = s & " p.lot_num, p.count_qty, p.lot2, p.qty2, p.sku"                    'jv051917
        s = s & " from lane l, position p"
        s = s & " where ((p.barcode >= '" & bbarcode & "'"
        's = s & " and p.barcode <= '" & bbarcode & "EOR') or p.lot2 = '" & wdlot & "')"
        s = s & " and p.barcode <= '" & bbarcode & "EOR')"                          'jv051917
        s = s & " or (p.lot2 = '" & wdlot & "' and p.sku = '" & wdsku & "'))"       'jv051917
        s = s & " and l.id = p.laneno"
        'MsgBox s
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds(0) & Chr(9)
                If ds(0) = 5 Then
                    s = s & ds(1) & " " & ds(2) & "-" & ds(3) & "-" & ds(4) & Chr(9)
                Else
                    s = s & ds(2) & "-" & ds(3) & "-" & ds(4) & " " & ds(5) & Chr(9)
                End If
                s = s & ds(6) & Chr(9)
                s = s & ds(7) & Chr(9)
                s = s & ds(8) & Chr(9)
                s = s & ds(9) & Chr(9)
                s = s & ds(10)
                Grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    
    If Form1.plantno = "52" Then                    'CS5
        Set cb = CreateObject("ADODB.Connection")
        cb.Open "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=sywms;PWD=!Sylacauga_WMS1907"
        t = Mid(bbarcode, 5, 2) & "/" & Mid(bbarcode, 7, 2) & "/20" & Mid(bbarcode, 9, 2)
        t = Format(t, "M/d/yyyy")
        s = "EXEC bb_get_pallet_locations '" & Trim(Left(bbarcode, 4))
        s = s & "-" & Right(bbarcode, 3) & "', '" & t & "'"
        'MsgBox s
        Set cs = cb.Execute(s)
        If cs.BOF = False Then
            cs.MoveFirst
            Do Until cs.EOF
                s = "CS5" & Chr(9) & Left(cs(0), 8) & Chr(9) '& Trim(cs(18))
                t = "select barcode, lot1, qty1, lot2, qty2 from pallets where barcode = '" & Trim(cs(1)) & "'"
                Set ds = Wdb.Execute(t)
                If ds.BOF = False Then
                    ds.MoveFirst
                    s = s & ds!barcode & Chr(9)
                    s = s & ds!lot1 & Chr(9)
                    s = s & ds!qty1 & Chr(9)
                    s = s & ds!lot2 & Chr(9)
                    s = s & ds!qty2 & Chr(9)
                End If
                ds.Close
                Grid1.AddItem s
                cs.MoveNext
            Loop
        End If
        cs.Close: cb.Close
    End If
    
    's = "select r.aisle, r.rack, p.barcode, p.lot_num, p.count_qty, p.lot2, p.qty2 from racks r, rackpos p"
    s = "select r.aisle, r.rack, p.barcode, p.lot_num, p.count_qty, p.lot2, p.qty2, p.sku from racks r, rackpos p"  'jv051917
    s = s & " Where ((p.barcode >= '" & bbarcode & "'"
    's = s & " and p.barcode <= '" & bbarcode & "EOR') or p.lot2 = '" & wdlot & "')"
    s = s & " and p.barcode <= '" & bbarcode & "EOR')"                          'jv051917
    s = s & " or (p.lot2 = '" & wdlot & "' and p.sku = '" & wdsku & "'))"       'jv051917
    s = s & " and r.id = p.rackno"
    'MsgBox s
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "4" & Chr(9)
            s = s & Trim(ds(0)) & "-" & Trim(ds(1)) & Chr(9)
            s = s & ds(2) & Chr(9)
            s = s & ds(3) & Chr(9)
            s = s & ds(4) & Chr(9)
            s = s & ds(5) & Chr(9)
            s = s & ds(6) & Chr(9)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 2: Grid1.ColSel = 2
        Grid1.Sort = 5
        'MsgBox wdlot
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 5) > " " Then            '2nd lot
                If Grid1.TextMatrix(i, 5) = wdlot Then
                    s = "select hsource from holdlist where sku = '" & Trim(Left(bbarcode, 4)) & "'" 'jv122815
                    s = s & " and lot_num = '" & Grid1.TextMatrix(i, 3) & "'"
                    s = s & " and opcode = '" & Trim(Mid(Grid1.TextMatrix(i, 2), 11, 3)) & "'"
                    s = s & " and spallet <= '" & Right(Grid1.TextMatrix(i, 2), 3) & "'"
                    s = s & " and epallet >= '" & Right(Grid1.TextMatrix(i, 2), 3) & "'"
                    'MsgBox s
                    Set ds = Wdb.Execute(s)
                    If ds.BOF = False Then
                        ds.MoveFirst
                        Grid1.TextMatrix(i, 7) = "Yes"
                        Grid1.TextMatrix(i, 8) = Grid1.TextMatrix(i, 6)
                        If LCase(ds!hsource) = "schedule" Then              'jv122815
                            Grid1.TextMatrix(i, 9) = "TEST HOLD"            'jv122815
                        Else                                                'jv122815
                            Grid1.TextMatrix(i, 9) = ds!hsource             'jv122815
                        End If                                              'jv122815
                    End If
                    ds.Close
                Else
                     s = "select hsource from holdlist where sku = '" & Trim(Left(bbarcode, 4)) & "'"
                     s = s & " and lot_num = '" & Left(Grid1.TextMatrix(i, 5), 5) & "'"
                     s = s & " and opcode = '" & Trim(Right(Grid1.TextMatrix(i, 5), 3)) & "'"
                     'MsgBox s
                     Set ds = Wdb.Execute(s)
                     If ds.BOF = False Then
                        ds.MoveFirst
                        Grid1.TextMatrix(i, 7) = "Yes"
                        Grid1.TextMatrix(i, 8) = Grid1.TextMatrix(i, 4)
                        If LCase(ds!hsource) = "schedule" Then              'jv122815
                            Grid1.TextMatrix(i, 9) = "TEST HOLD"            'jv122815
                        Else                                                'jv122815
                            Grid1.TextMatrix(i, 9) = ds!hsource             'jv122815
                        End If                                              'jv122815
                    End If
                    ds.Close
                End If
            Else                                                                                    'jv122815
                'MsgBox wdlot
                If Grid1.TextMatrix(i, 3) = Left(wdlot, 5) Then                                     'jv122815
                    s = "select hsource from holdlist where sku = '" & Trim(Left(bbarcode, 4)) & "'"   'jv122815
                    s = s & " and lot_num = '" & Grid1.TextMatrix(i, 3) & "'"                       'jv122815
                    s = s & " and opcode = '" & Trim(Mid(Grid1.TextMatrix(i, 2), 11, 3)) & "'"      'jv122815
                    s = s & " and spallet <= '" & Right(Grid1.TextMatrix(i, 2), 3) & "'"            'jv122815
                    s = s & " and epallet >= '" & Right(Grid1.TextMatrix(i, 2), 3) & "'"            'jv122815
                    'MsgBox s
                    Set ds = Wdb.Execute(s)                                                         'jv122815
                    If ds.BOF = False Then                                                          'jv122815
                        ds.MoveFirst                                                                'jv122815
                        If LCase(ds!hsource) = "schedule" Then              'jv122815
                            Grid1.TextMatrix(i, 9) = "TEST HOLD"            'jv122815
                        Else                                                'jv122815
                            Grid1.TextMatrix(i, 9) = ds!hsource             'jv122815
                        End If                                              'jv122815
                        
                    End If                                                                          'jv122815
                    ds.Close                                                                        'jv122815
                End If                                                                              'jv122815
            End If
        Next i
        aqty = 0: hqty = 0
        
        If Grid1.Rows > 2 Then                                                  'jv030718
            For i = 1 To Grid1.Rows - 2                                         'jv030718
                If Grid1.TextMatrix(i, 2) = Grid1.TextMatrix(i + 1, 2) Then     'jv030718
                    Grid1.Row = i: Grid1.RowSel = i                             'jv030718
                    Grid1.Col = 2: Grid1.ColSel = 2                             'jv030718
                    Grid1.CellBackColor = dupflag.BackColor                     'jv030718
                    Grid1.CellForeColor = dupflag.ForeColor                     'jv030718
                    Grid1.Row = i + 1: Grid1.RowSel = i + 1                     'jv030718
                    Grid1.Col = 2: Grid1.ColSel = 2                             'jv030718
                    Grid1.CellBackColor = dupflag.BackColor                     'jv030718
                    Grid1.CellForeColor = dupflag.ForeColor                     'jv030718
                    dupflag.Visible = True                                      'jv030718
                End If                                                          'jv030718
            Next i                                                              'jv030718
        End If                                                                  'jv030718
        
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 5) > " " Then
                Grid1.Row = i: Grid1.RowSel = i
                If Grid1.TextMatrix(i, 5) = wdlot Then
                    If Grid1.TextMatrix(i, 7) > " " Then
                        hqty = hqty + Val(Grid1.TextMatrix(i, 8))
                    Else
                        aqty = aqty + Val(Grid1.TextMatrix(i, 6))
                    End If
                    Grid1.Col = 5: Grid1.ColSel = 6
                Else
                    If Grid1.TextMatrix(i, 7) > " " Then
                        hqty = hqty + Val(Grid1.TextMatrix(i, 8))
                    Else
                        aqty = aqty + Val(Grid1.TextMatrix(i, 4))
                    End If
                    Grid1.Col = 3: Grid1.ColSel = 4
                End If
                Grid1.CellBackColor = ycolor.BackColor
            Else
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 3: Grid1.ColSel = 4
                Grid1.CellBackColor = ycolor.BackColor
                aqty = aqty + Val(Grid1.TextMatrix(i, 4))
            End If
            If Len(Grid1.TextMatrix(i, 2)) = 16 Then Grid1.TextMatrix(i, 2) = bc000(Grid1.TextMatrix(i, 2))
            If Grid1.TextMatrix(i, 9) > " " Then                'jv122815
                Grid1.Row = i: Grid1.RowSel = i                 'jv122815
                Grid1.Col = 1: Grid1.ColSel = 2                 'jv122815
                Grid1.CellForeColor = ycolor.ForeColor          'jv122815
                Grid1.Row = i: Grid1.RowSel = i                 'jv122815
                Grid1.Col = 9: Grid1.ColSel = 9                 'jv122815
                Grid1.CellForeColor = ycolor.ForeColor          'jv122815
            End If                                              'jv122815
        Next i
        bunits = Format(aqty + hqty, "0")
        s = "All" & Chr(9)
        s = s & " " & Chr(9)
        s = s & "Totals" & Chr(9) & Chr(9) & aqty & Chr(9)
        s = s & Chr(9) & Chr(9) & Chr(9) & hqty
        Grid1.AddItem s
        Grid1.Row = 1
    End If
    Grid1.FormatString = "^SR|^Rack|^BarCode|^Lot|^Units|^Lot2|^Units|^2nd Lot OnHold|^Hold Qty|^Status"
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 1900
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1600
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1400
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub bbarcode_Change()
    refresh_grid
End Sub


Private Sub Command1_Click()                'Print List
    Dim rt As String, rh As String, rf As String
    rt = "Batch Inventory - " & bbarcode.Caption
    rh = bproduct.Caption
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    Grid1.Redraw = False
    'If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
    '    Call printflexgrid(Printer, Grid1, rt, rh, rf)
    'Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    'End If
    Grid1.Redraw = True
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Label1.Height * 4)
End Sub
