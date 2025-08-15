VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form batchtktinv 
   Caption         =   "Batch Ticket Inventory"
   ClientHeight    =   10365
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   10365
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   16960
      _Version        =   327680
      ForeColor       =   12582912
      BackColorFixed  =   16777152
   End
   Begin VB.Label bunits 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WMS Units:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label bproduct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "bproduct"
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
      Left            =   7080
      TabIndex        =   6
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Product:"
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
      Left            =   6000
      TabIndex        =   5
      Top             =   0
      Width           =   1215
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
      Left            =   4200
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BarCode:"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label batchno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "batchno"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Batch Ticket:"
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
   End
   Begin VB.Menu pastemenu 
      Caption         =   "Paste"
      Begin VB.Menu paste2bat 
         Caption         =   "Paste to Production Batches"
      End
   End
End
Attribute VB_Name = "batchtktinv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function format_bc(bc As String) As String
    Dim s As String
    s = Trim(Mid(bc, 1, 4)) & " "
    s = s & Mid(bc, 5, 6) & "  "
    s = s & Mid(bc, 11, 3) & "  "
    s = s & Mid(bc, 14, 3)
    format_bc = s
End Function

Private Sub refresh_grid()
    Dim db As ADODB.Connection, ds As ADODB.Recordset, s As String, i As Integer, wdlot As String
    Dim cb As ADODB.Connection, cs As ADODB.Recordset, t As String, wdsku As String
    Dim aqty As Long, hqty As Long
    If plant_server_status(prodbatches.Combo1) = False Then                             'jv010417
        s = "Sorry, The server for Warehouse " & prodbatches.Combo1 & " has been flagged to be offline."
        MsgBox s, vbOKOnly + vbInformation, "sorry, try again later..."                 'jv010417
        Exit Sub                                                                        'jv010417
    End If                                                                              'jv010417
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 10
    bunits = "0"
    wdlot = barcode_to_lotnum(bbarcode & "EOR")
    wdlot = wdlot & Right(bbarcode, 3)
    wdsku = Trim(Left(bbarcode, 4))                     'jv051917
    Set db = CreateObject("ADODB.Connection")
    If prodbatches.Combo1 = "T10" Then db.Open t10bbsr
    If prodbatches.Combo1 = "K10" Then db.Open k10bbsr
    If prodbatches.Combo1 = "A10" Then db.Open a10bbsr
    
    If prodbatches.Combo1 = "T10" Then          'T10 Cranes
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
        Set ds = db.Execute(s)
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
    
    If prodbatches.Combo1 = "A10" Then
        Set cb = CreateObject("ADODB.Connection")
        cb.Open cs5db
        t = Format(DateAdd("yyyy", 2, prodbatches.Grid1.TextMatrix(prodbatches.Grid1.Row, 1)), "M/d/yyyy")
        ''s = "Select * from vContainerLocation_1033 where item = '" & Trim(Left(bbarcode, 4))
        's = "Select location, [Pal ID] from vContainerLocation_1033 where item = '" & Trim(Left(bbarcode, 4))
        's = s & "-" & Right(bbarcode, 3) & "' and expiration = '" & t & "'"
        s = "Select location, LPN from vAllInventory_1033 where item = '" & Trim(Left(bbarcode, 4)) 'Westfalia
        s = s & "-" & Right(bbarcode, 3) & "' and [Lot Expiration] = '" & t & "'"                   'Upgrade
        'MsgBox s
        Set cs = cb.Execute(s)
        If cs.BOF = False Then
            cs.MoveFirst
            Do Until cs.EOF
                's = "CS5" & Chr(9) & Left(cs(8), 8) & Chr(9) '& Trim(cs(18))
                't = "select barcode, lot1, qty1, lot2, qty2 from pallets where plateno = '" & Trim(cs(18)) & "'"
                s = "CS5" & Chr(9) & Left(cs(0), 8) & Chr(9) '& Trim(cs(18))
                t = "select barcode, lot1, qty1, lot2, qty2 from pallets where plateno = '" & Trim(cs(1)) & "'"
                t = t & " or barcode = '" & Trim(cs(1)) & "'"               'jv010417
                'MsgBox t
                Set ds = db.Execute(t)
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
    Set ds = db.Execute(s)
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
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 5) > " " Then            '2nd lot
                If Grid1.TextMatrix(i, 5) = wdlot Then
                    s = "select hsource from holdlist where sku = '" & Trim(Left(bbarcode, 4)) & "'"    'jv122815
                    s = s & " and lot_num = '" & Grid1.TextMatrix(i, 3) & "'"
                    s = s & " and opcode = '" & Mid(Grid1.TextMatrix(i, 2), 11, 3) & "'"
                    s = s & " and spallet <= '" & Right(Grid1.TextMatrix(i, 2), 3) & "'"
                    s = s & " and epallet >= '" & Right(Grid1.TextMatrix(i, 2), 3) & "'"
                    'MsgBox s
                    Set ds = db.Execute(s)
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
                     s = "select hsource from holdlist where sku = '" & Trim(Left(bbarcode, 4)) & "'"   'jv122815
                     s = s & " and lot_num = '" & Left(Grid1.TextMatrix(i, 5), 5) & "'"
                     s = s & " and opcode = '" & Right(Grid1.TextMatrix(i, 5), 3) & "'"
                     'MsgBox s
                     Set ds = db.Execute(s)
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
                    Set ds = db.Execute(s)                                                         'jv122815
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
            If Len(Grid1.TextMatrix(i, 2)) = 16 Then Grid1.TextMatrix(i, 2) = format_bc(Grid1.TextMatrix(i, 2))
            If Grid1.TextMatrix(i, 9) > " " Then                'jv122815
                Grid1.Row = i: Grid1.RowSel = i                 'jv122815
                Grid1.Col = 1: Grid1.ColSel = 2                 'jv122815
                Grid1.CellForeColor = ycolor.ForeColor          'jv122815
                Grid1.Row = i: Grid1.RowSel = i                 'jv122815
                Grid1.Col = 9: Grid1.ColSel = 9                 'jv122815
                Grid1.CellForeColor = ycolor.ForeColor          'jv122815
            End If                                              'jv12281
        Next i
        bunits = Format(aqty + hqty, "0")
        s = "All" & Chr(9)
        s = s & " " & Chr(9)
        s = s & "Totals" & Chr(9) & Chr(9) & aqty & Chr(9)
        s = s & Chr(9) & Chr(9) & Chr(9) & hqty
        Grid1.AddItem s
        Grid1.Row = 1
    End If
    db.Close
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

Private Sub Form_Load()
    Me.Height = prodbatches.Height
    Me.Top = prodbatches.Top
    Me.Left = prodbatches.Width - Me.Width
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Label1.Height * 5)
End Sub

Sub paste2bat_Click()
    Dim i As Integer, pdate As String, psku As String, k As Integer
    Dim tbc As String, palcnt As Integer
    Dim t7 As Long, t8 As Long, t10 As Long, t11 As Long
    Dim g7 As Long, g8 As Long, g10 As Long, g11 As Long
    t7 = 0: t8 = 0: t10 = 0: t11 = 0: palcnt = 0
    g7 = 0: g8 = 0: g10 = 0: g11 = 0
    If Val(bunits) = 0 Then Exit Sub
    If bunits = prodbatches.Grid1.TextMatrix(prodbatches.Grid1.Row, 7) Then Exit Sub
    tbc = Left(bbarcode, 10) & "  " & Mid(bbarcode, 11, 3)
    For i = 1 To Grid1.Rows - 1
        If Left(Grid1.TextMatrix(i, 2), 15) = tbc Then
            palcnt = palcnt + 1
        End If
    Next i
    For i = 1 To prodbatches.Grid1.Rows - 1
        If prodbatches.Grid1.TextMatrix(i, 0) = Me.batchno Then
            pdate = prodbatches.Grid1.TextMatrix(i, 1)
            psku = prodbatches.Grid1.TextMatrix(i, 4)
            prodbatches.Grid1.TextMatrix(i, 7) = Me.bunits
            prodbatches.Grid1.TextMatrix(i, 8) = Val(prodbatches.Grid1.TextMatrix(i, 7)) - Val(prodbatches.Grid1.TextMatrix(i, 6))
            prodbatches.Grid1.TextMatrix(i, 10) = palcnt 'Grid1.Rows - 2
            prodbatches.Grid1.TextMatrix(i, 11) = Val(prodbatches.Grid1.TextMatrix(i, 10)) - Val(prodbatches.Grid1.TextMatrix(i, 9))
            prodbatches.Grid1.TextMatrix(0, 7) = "Received"
            If prodbatches.sortdate.Checked = True Then
                For k = 1 To prodbatches.Grid1.Rows - 1
                    If Val(prodbatches.Grid1.TextMatrix(k, 0)) > 0 Then
                        g7 = g7 + Val(prodbatches.Grid1.TextMatrix(k, 7))
                        g8 = g8 + Val(prodbatches.Grid1.TextMatrix(k, 8))
                        g10 = g10 + Val(prodbatches.Grid1.TextMatrix(k, 10))
                        g11 = g11 + Val(prodbatches.Grid1.TextMatrix(k, 11))
                    End If
                    If prodbatches.Grid1.TextMatrix(k, 1) = pdate Then
                        t7 = t7 + Val(prodbatches.Grid1.TextMatrix(k, 7))
                        t8 = t8 + Val(prodbatches.Grid1.TextMatrix(k, 8))
                        t10 = t10 + Val(prodbatches.Grid1.TextMatrix(k, 10))
                        t11 = t11 + Val(prodbatches.Grid1.TextMatrix(k, 11))
                    End If
                Next k
                For k = 1 To prodbatches.Grid1.Rows - 1
                    If prodbatches.Grid1.TextMatrix(k, 5) = "Daily Total " & pdate Then
                        prodbatches.Grid1.TextMatrix(k, 7) = t7
                        prodbatches.Grid1.TextMatrix(k, 8) = t8
                        prodbatches.Grid1.TextMatrix(k, 10) = t10
                        prodbatches.Grid1.TextMatrix(k, 11) = t11
                    End If
                Next k
                k = prodbatches.Grid1.Rows - 1
                prodbatches.Grid1.TextMatrix(k, 7) = g7
                prodbatches.Grid1.TextMatrix(k, 8) = g8
                prodbatches.Grid1.TextMatrix(k, 10) = g10
                prodbatches.Grid1.TextMatrix(k, 11) = g11
            End If
            If prodbatches.sortsku.Checked = True Then
                For k = 1 To prodbatches.Grid1.Rows - 1
                    If Val(prodbatches.Grid1.TextMatrix(k, 0)) > 0 Then
                        g7 = g7 + Val(prodbatches.Grid1.TextMatrix(k, 7))
                        g8 = g8 + Val(prodbatches.Grid1.TextMatrix(k, 8))
                        g10 = g10 + Val(prodbatches.Grid1.TextMatrix(k, 10))
                        g11 = g11 + Val(prodbatches.Grid1.TextMatrix(k, 11))
                    End If
                    If prodbatches.Grid1.TextMatrix(k, 4) = psku Then
                        t7 = t7 + Val(prodbatches.Grid1.TextMatrix(k, 7))
                        t8 = t8 + Val(prodbatches.Grid1.TextMatrix(k, 8))
                        t10 = t10 + Val(prodbatches.Grid1.TextMatrix(k, 10))
                        t11 = t11 + Val(prodbatches.Grid1.TextMatrix(k, 11))
                    End If
                Next k
                For k = 1 To prodbatches.Grid1.Rows - 1
                    If prodbatches.Grid1.TextMatrix(k, 5) = "SKU Total - " & psku Then
                        prodbatches.Grid1.TextMatrix(k, 7) = t7
                        prodbatches.Grid1.TextMatrix(k, 8) = t8
                        prodbatches.Grid1.TextMatrix(k, 10) = t10
                        prodbatches.Grid1.TextMatrix(k, 11) = t11
                    End If
                Next k
                k = prodbatches.Grid1.Rows - 1
                prodbatches.Grid1.TextMatrix(k, 7) = g7
                prodbatches.Grid1.TextMatrix(k, 8) = g8
                prodbatches.Grid1.TextMatrix(k, 10) = g10
                prodbatches.Grid1.TextMatrix(k, 11) = g11
            End If
        End If
    Next i
    Unload Me
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rf As String, rh As String
    rt = "Batch Ticket: " & Me.batchno & "   BarCode: " & Me.bbarcode
    rh = Me.bproduct & "  Total Units: " & bunits
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "cyan": gndc(0) = Me.Grid1.BackColorFixed
    htdc(1) = "yellow": gndc(1) = Me.ycolor.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid1.BackColor
    Grid1.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub
