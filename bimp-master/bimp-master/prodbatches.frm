VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form prodbatches 
   Caption         =   "Production Batch Tickets"
   ClientHeight    =   11175
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   ScaleHeight     =   11175
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ticket Inventory"
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
      Left            =   12360
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   3135
      Left            =   0
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   5530
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8775
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   15478
      _Version        =   327680
      ForeColor       =   4194368
      BackColorFixed  =   12648384
      BackColorSel    =   8388736
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
      Left            =   10440
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   10560
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   8760
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
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
      Left            =   840
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label rcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hold Quantity Adjusted"
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
      Left            =   2160
      TabIndex        =   12
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "thru"
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
      Left            =   7800
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Production Dates:"
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
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "..."
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
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Plant:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu mnq 
         Caption         =   "Mark New Qty"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu sortmenu 
      Caption         =   "Sort"
      Begin VB.Menu sortdate 
         Caption         =   "Date"
         Checked         =   -1  'True
      End
      Begin VB.Menu sortsku 
         Caption         =   "SKU"
      End
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
      Begin VB.Menu prtgrid 
         Caption         =   "Print Grid"
      End
   End
   Begin VB.Menu impmenu 
      Caption         =   "Import"
      Visible         =   0   'False
      Begin VB.Menu impwms 
         Caption         =   "Paste WMS Batch Units to Released"
      End
   End
End
Attribute VB_Name = "prodbatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim q As String, i As Integer
    'Dim dsn As String, userid As String, pwd As String
    Dim db As ADODB.Connection, ds As ADODB.Recordset, s As String, hs As ADODB.Recordset
    Dim t6 As Long, t7 As Long, t8 As Long                      'jv121415
    Dim t9 As Long, t10 As Long, t11 As Long, t13 As Long, t14 As Long, t15 As Long
    Dim k As Long, nl As Boolean, wdlot As String, pflag As String
    Dim psku As String, plot As String, sp As String, ep As String
    
    rcolor.Visible = False
    wdlot = Right(Text1, 2)
    wdlot = wdlot & Format(DateDiff("d", "1-1-" & Right(Text1, 4), Text1) + 1, "000")
    'MsgBox wdlot
    
    If plant_server_status(prodbatches.Combo1) = False Then                             'jv010417
        s = "Sorry, The server for Warehouse " & prodbatches.Combo1 & " has been flagged to be offline."
        MsgBox s, vbOKOnly + vbInformation, "sorry, try again later..."                 'jv010417
        Exit Sub                                                                        'jv010417
    End If                                                                              'jv010417
    
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    
    'On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 17
    'q = "select h.batch_no,TO_CHAR(h.plan_start_date,'MM-DD-YYYY'),h.batch_status,"
    q = "select h.batch_no,TO_CHAR(h.plan_start_date,'YYYY-MM-DD'),h.batch_status,"         'jv010516
    q = q & "h.attribute1,i.segment1,i.description,d.plan_qty,"
    q = q & "d.actual_qty"
    q = q & " from apps.gme_batch_header h, apps.gme_material_details d, apps.mtl_system_items_b i"
    If Val(List1) = 500 Then
        q = q & " where h.organization_id in (select organization_id from mtl_parameters where organization_code in ('500','503'))"
    Else
        q = q & " where h.organization_id in (select organization_id from mtl_parameters where organization_code in ('" & Format(Val(List1), "000") & "'))"
    End If
    q = q & " and h.plan_start_date >= TO_DATE('" & Format(Text1, "DD-MMM-YYYY") & "')"
    q = q & " and h.plan_start_date <= TO_DATE('" & Format(DateAdd("d", 1, Text2), "DD-MMM-YYYY") & "')"
    q = q & " and h.delete_mark = 0"
    q = q & " and h.batch_id = d.batch_id"
    q = q & " and h.batch_status in (1, 2, 3, 4)"
    q = q & " and d.line_type = 1"
    q = q & " and i.organization_id = d.organization_id"
    q = q & " and i.inventory_item_id = d.inventory_item_id"
    q = q & " and i.segment1 >= '100' and i.segment1 <= '9999'"             'jv082415
    If sortsku.Checked Then
        q = q & " order by i.segment1, 2, d.plan_qty desc, h.attribute1"
    Else
        q = q & " order by 2, i.segment1, d.plan_qty desc, h.attribute1"
    End If
    'MsgBox q
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(0) & Chr(9)                              'Batch No
            s = s & Format(ds(1), "M-dd-yyyy") & Chr(9)     'Date
            If ds(2) = 1 Then s = s & "PEND" & Chr(9)       'Status
            If ds(2) = 2 Then s = s & "WIP" & Chr(9)
            If ds(2) = 3 Then s = s & "CERT" & Chr(9)
            If ds(2) = 4 Then s = s & "Closed" & Chr(9)
            s = s & ds(3) & Chr(9)                          'Location
            s = s & ds(4) & Chr(9)                          'SKU
            s = s & ds(5) & Chr(9)                          'Product Name
            s = s & ds(6) & Chr(9)                          'Planned Qty
            s = s & ds(7) & Chr(9)                          'Actual Qty
            s = s & Format(ds(7) - ds(6), "0")              'Qty Diff
            pflag = Trim(ds(4))
            If Len(pflag) = 3 Then pflag = pflag & " "
            If Format(ds(1), "M-dd-yyyy") = "2-29-2016" Then                    'jv030416
                pflag = pflag & "022918"                                        'jv030416
            ElseIf Format(ds(1), "M-dd-yyyy") = "2-29-2020" Then
                pflag = pflag & "022922"
            ElseIf Format(ds(1), "M-dd-yyyy") = "2-29-2024" Then
                pflag = pflag & "022926"
            Else                                                                'jv030416
                pflag = pflag & Format(DateAdd("yyyy", 2, Format(ds(1), "M-dd-yyyy")), "MMddyy")
            End If
            pflag = pflag & Right(ds(3), 3)
            s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9) & pflag
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    'Grid1.Redraw = True
    
    'Pool Schedule Pallets
    s = "select * from poolschedule where plantwhs = '" & Combo1 & "'"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = ds!batchno Then
                    Grid1.TextMatrix(i, 16) = ds!palqty ' New Pool
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.Redraw = True
    
    'MsgBox "racks"
    Set db = CreateObject("ADODB.Connection")
    If Combo1 = "A10" Then db.Open a10bbsr
    If Combo1 = "K10" Then db.Open k10bbsr
    If Combo1 = "T10" Then db.Open t10bbsr
    ' This query selects all pallets in the warehouse with the specified lot number that are not Order Pick
    s = "select sku, lot_num, barcode from rackpos where lot_num >= '" & wdlot & "' and barcode < '9999'"
    s = s & " and rackno not in (select id from racks where rack = 'OP')"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                ' This adds up all the matching pallets based on barcode minus pallet number and keeps track in column 13 (on hand)
                If Grid1.TextMatrix(i, 12) = Left(ds!barcode, 13) Then ' Barcode
                    Grid1.TextMatrix(i, 13) = Val(Grid1.TextMatrix(i, 13)) + 1 ' WMS Quantity
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    'Rack Hold List
    s = "select * from holdlist where lot_num >= '" & wdlot & "'"
    Set hs = db.Execute(s)
    If hs.BOF = False Then
        hs.MoveFirst
        Do Until hs.EOF
            If hs!lot_num = "16060" Then
                If Len(hs!sku) = 3 Then
                    sp = hs!sku & " " & "022918" & hs!opcode & hs!spallet
                    ep = hs!sku & " " & "022918" & hs!opcode & hs!epallet
                Else
                    sp = hs!sku & "022918" & hs!opcode & hs!spallet
                    ep = hs!sku & "022918" & hs!opcode & hs!epallet
                End If
            Else
                If Len(hs!sku) = 3 Then
                    sp = hs!sku & " " & r12lot(hs!lot_num) & hs!opcode & hs!spallet ' Start Pallet on hold
                    ep = hs!sku & " " & r12lot(hs!lot_num) & hs!opcode & hs!epallet ' End Pallet on hold
                Else
                    sp = hs!sku & r12lot(hs!lot_num) & hs!opcode & hs!spallet
                    ep = hs!sku & r12lot(hs!lot_num) & hs!opcode & hs!epallet
                End If
            End If
            ' Get all pallets in storage that are on hold between the start and end pallets defined in the above conditional
            s = "select sku, count(*) from rackpos where barcode >= '" & sp & "'"
            s = s & " and barcode <= '" & ep & "' group by sku"
            'MsgBox s
            Set ds = db.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                'MsgBox ds(1), vbOKOnly, sp
                For i = 1 To Grid1.Rows - 1
                    ' This adds up all the matching pallets based on barcode minus pallet number and keeps track in column 14 (hold)
                    If Grid1.TextMatrix(i, 12) = Left(sp, 13) Then ' Barcode/Flag
                        Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 14)) + ds(1) ' Hold
                        Exit For
                    End If
                Next i
            End If
            ds.Close
            hs.MoveNext
        Loop
    End If
    hs.Close
    Grid1.Redraw = True
    
    If Combo1 = "T10" Then
        's = "select barcode, pallet_status from position where lot_num >= '" & wdlot & "' and barcode < '9999'"
        s = "select p.barcode, l.lane_status from position p, lane l where l.id = p.laneno"
        s = s & " and p.lot_num >= '" & wdlot & "' and p.barcode < '9999'"
        'MsgBox s
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                For i = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 12) = Left(ds!barcode, 13) Then ' Barcode/Flag
                        Grid1.TextMatrix(i, 13) = Val(Grid1.TextMatrix(i, 13)) + 1 ' WMS Qty
                        'If ds!pallet_status = "H" Then
                        If ds(1) = "H" Then
                            Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 14)) + 1 ' Hold
                        End If
                        Exit For
                    End If
                Next i
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    db.Close
    Grid1.Redraw = True
    If Combo1 = "A10" Then
        'MsgBox "cs5"
        db.Open cs5db
        
        ' This gets all the pallets that individually marked as being on hold within Westfalia.
        s = "Select item, [Lot Expiration], Lock, LPN from vAllInventory_1033"       'Westfalia Upgrade
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                'If Len(ds(10)) > 1 Then
                '    psku = Trim(ds(10))
                '    plot = Trim(ds(14))
                If Len(ds(0)) > 1 Then
                    psku = Trim(ds(0))
                    plot = Trim(ds(1))
                    'If Val(Mid(psku, 4, 1)) > 0 Then
                    If Mid(psku, 4, 1) <> "-" Then                      'jv032118
                        s = Left(psku, 4)
                    Else
                        s = Left(psku, 3) & " "
                    End If
                    s = s & Format(plot, "MMddyy")
                    s = s & Right(psku, 3)
                    s = ds(3)                                           'jv071818
                    For i = 1 To Grid1.Rows - 1
                        If Grid1.TextMatrix(i, 12) = Left(s, 13) Then ' Barcode/Flag
                            ' This adds up all the matching pallets in westfalia based on barcode minus pallet number and keeps track in columns 13 and 14
                            Grid1.TextMatrix(i, 13) = Val(Grid1.TextMatrix(i, 13)) + 1 ' WMS Qty
                            'If Trim(ds(16)) = "True" Then
                            'If Trim(ds(2)) = "True" Then
                            If ds(2) <> 0 Then                                  'jv072618
                                Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 14)) + 1 ' Hold
                                'MsgBox psku & " " & plot & ": " & Grid1.TextMatrix(i, 14), vbOKOnly, "vAllInventory"
                            End If
                            Exit For
                        End If
                    Next i
                    'MsgBox s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close                                                                'jv071818
        
        ' This gets all of the entire lots that are on hold in Westfalia, and subtracts the pallets that were already accounted for above.
        s = "SELECT l.Quantity - ISNULL(i.PalletHoldQuantity, 0) AS LotHoldQuantity, l.[Lot ID] FROM vLotData_1033 l OUTER APPLY (SELECT SUM(Quantity) AS PalletHoldQuantity FROM vAllInventory_1033 WHERE Lot = l.[Lot ID] AND Lock = 1) i WHERE l.Lock = 1"
        Set ds = db.Execute(s)                                                  'jv071818
        If ds.BOF = False Then                                                  'jv071818
            ds.MoveFirst
            Do Until ds.EOF
                If Mid(ds(1), 4, 1) = "-" Then
                    s = Mid(ds(1), 1, 3) & " "
                    s = s & Mid(ds(1), 9, 6)
                    s = s & Mid(ds(1), 5, 3)
                    psku = Mid(ds(1), 1, 3)
                Else
                    s = Mid(ds(1), 1, 4)
                    s = s & Mid(ds(1), 10, 6)
                    s = s & Mid(ds(1), 6, 3)
                    psku = Mid(ds(1), 1, 4)
                End If
                For i = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 12) = s Then ' Barcode/Flag
                        Grid1.TextMatrix(i, 14) = Format(Val(Grid1.TextMatrix(i, 14)) + CInt(ds(0) / skurec(Val(psku)).pallet), "0") ' Hold
                        'MsgBox psku & " " & plot & ": " & Grid1.TextMatrix(i, 14), vbOKOnly, "vLotData"
                        Exit For
                    End If
                Next i
                ds.MoveNext
            Loop
        End If
        ds.Close: db.Close
    End If
    Screen.MousePointer = 0
    
    'If Grid1.Rows > 2 Then Grid1.RemoveItem Grid1.Rows - 1
    If Combo1 = "A10" Then
        Grid1.FormatString = "^Batch No|^Plan Start|^Status|<Location|^SKU|<Description|^Planned|^Released|^Diff|^PalPlan|^PalAct|^PalDiff|^Flag|^WMS Qty|^Hold|^Avail|^New Pool"
    Else
        Grid1.FormatString = "^Batch No|^Plan Start|^Status|<Location|^SKU|<Description|^Planned|^Released|^Diff|^PalPlan|^PalAct|^PalDiff|^BarCode|^WMS Qty|^Hold|^Avail|^New Pool"
    End If
    Grid1.ColWidth(0) = 900
    Grid1.ColWidth(1) = 1100
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 2000
    Grid1.ColWidth(4) = 700
    Grid1.ColWidth(5) = 2200
    Grid1.ColWidth(6) = 900
    Grid1.ColWidth(7) = 900
    Grid1.ColWidth(8) = 900
    Grid1.ColWidth(9) = 800
    Grid1.ColWidth(10) = 800
    Grid1.ColWidth(11) = 800
    Grid1.ColWidth(12) = 1400
    Grid1.ColWidth(13) = 900
    Grid1.ColWidth(14) = 900
    Grid1.ColWidth(15) = 900
    Grid1.ColWidth(16) = 1100
    Grid1.FillStyle = flexFillRepeat
    ' Grid2 is for calculating totals/subtotals
    'Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 7
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 10                'jv121415
    t6 = 0: t7 = 0: t8 = 0                                      'jv121415
    t9 = 0: t10 = 0: t11 = 0: t13 = 0: t14 = 0: t15 = 0
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 13)) > Val(Grid1.TextMatrix(i, 14)) Then
            Grid1.TextMatrix(i, 15) = Val(Grid1.TextMatrix(i, 13)) - Val(Grid1.TextMatrix(i, 14))
        End If
        k = Val(Grid1.TextMatrix(i, 4))
        If skurec(k).sku = Grid1.TextMatrix(i, 4) Then
            Grid1.TextMatrix(i, 9) = CInt(Val(Grid1.TextMatrix(i, 6)) / skurec(k).pallet)
            Grid1.TextMatrix(i, 10) = CInt(Val(Grid1.TextMatrix(i, 7)) / skurec(k).pallet)
            Grid1.TextMatrix(i, 11) = Format(Val(Grid1.TextMatrix(i, 10)) - Val(Grid1.TextMatrix(i, 9)), "#")
        End If
        t6 = t6 + Val(Grid1.TextMatrix(i, 6))                   'jv121415
        t7 = t7 + Val(Grid1.TextMatrix(i, 7))                   'jv121415
        t8 = t8 + Val(Grid1.TextMatrix(i, 8))                   'jv121415
        t9 = t9 + Val(Grid1.TextMatrix(i, 9))
        t10 = t10 + Val(Grid1.TextMatrix(i, 10))
        t11 = t11 + Val(Grid1.TextMatrix(i, 11))
        t13 = t13 + Val(Grid1.TextMatrix(i, 13))
        t14 = t14 + Val(Grid1.TextMatrix(i, 14))
        t15 = t15 + Val(Grid1.TextMatrix(i, 15))
        nl = True
        If sortdate.Checked = True Then
            For k = 0 To Grid2.Rows - 1
                If Grid2.TextMatrix(k, 0) = Grid1.TextMatrix(i, 1) Then
                    Grid2.TextMatrix(k, 1) = Val(Grid2.TextMatrix(k, 1)) + Val(Grid1.TextMatrix(i, 6))
                    Grid2.TextMatrix(k, 2) = Val(Grid2.TextMatrix(k, 2)) + Val(Grid1.TextMatrix(i, 7))
                    Grid2.TextMatrix(k, 3) = Val(Grid2.TextMatrix(k, 3)) + Val(Grid1.TextMatrix(i, 8))
                    Grid2.TextMatrix(k, 4) = Val(Grid2.TextMatrix(k, 4)) + Val(Grid1.TextMatrix(i, 9))
                    Grid2.TextMatrix(k, 5) = Val(Grid2.TextMatrix(k, 5)) + Val(Grid1.TextMatrix(i, 10))
                    Grid2.TextMatrix(k, 6) = Val(Grid2.TextMatrix(k, 6)) + Val(Grid1.TextMatrix(i, 11))
                    Grid2.TextMatrix(k, 7) = Val(Grid2.TextMatrix(k, 7)) + Val(Grid1.TextMatrix(i, 13))
                    Grid2.TextMatrix(k, 8) = Val(Grid2.TextMatrix(k, 8)) + Val(Grid1.TextMatrix(i, 14))
                    Grid2.TextMatrix(k, 9) = Val(Grid2.TextMatrix(k, 9)) + Val(Grid1.TextMatrix(i, 15))
                    nl = False
                    Exit For
                End If
            Next k
            If nl = True Then
                s = Grid1.TextMatrix(i, 1) & Chr(9)
                s = s & Grid1.TextMatrix(i, 6) & Chr(9)             'jv121415
                s = s & Grid1.TextMatrix(i, 7) & Chr(9)             'jv121415
                s = s & Grid1.TextMatrix(i, 8) & Chr(9)             'jv121415
                s = s & Grid1.TextMatrix(i, 9) & Chr(9)
                s = s & Grid1.TextMatrix(i, 10) & Chr(9)
                s = s & Grid1.TextMatrix(i, 11) & Chr(9)
                s = s & Grid1.TextMatrix(i, 13) & Chr(9)
                s = s & Grid1.TextMatrix(i, 14) & Chr(9)
                s = s & Grid1.TextMatrix(i, 15)
                Grid2.AddItem s
            End If
        Else
            For k = 0 To Grid2.Rows - 1
                If Grid2.TextMatrix(k, 0) = Grid1.TextMatrix(i, 4) Then
                    Grid2.TextMatrix(k, 1) = Val(Grid2.TextMatrix(k, 1)) + Val(Grid1.TextMatrix(i, 6))
                    Grid2.TextMatrix(k, 2) = Val(Grid2.TextMatrix(k, 2)) + Val(Grid1.TextMatrix(i, 7))
                    Grid2.TextMatrix(k, 3) = Val(Grid2.TextMatrix(k, 3)) + Val(Grid1.TextMatrix(i, 8))
                    Grid2.TextMatrix(k, 4) = Val(Grid2.TextMatrix(k, 4)) + Val(Grid1.TextMatrix(i, 9))
                    Grid2.TextMatrix(k, 5) = Val(Grid2.TextMatrix(k, 5)) + Val(Grid1.TextMatrix(i, 10))
                    Grid2.TextMatrix(k, 6) = Val(Grid2.TextMatrix(k, 6)) + Val(Grid1.TextMatrix(i, 11))
                    Grid2.TextMatrix(k, 7) = Val(Grid2.TextMatrix(k, 7)) + Val(Grid1.TextMatrix(i, 13))
                    Grid2.TextMatrix(k, 8) = Val(Grid2.TextMatrix(k, 8)) + Val(Grid1.TextMatrix(i, 14))
                    Grid2.TextMatrix(k, 9) = Val(Grid2.TextMatrix(k, 9)) + Val(Grid1.TextMatrix(i, 15))
                    nl = False
                    Exit For
                End If
            Next k
            If nl = True Then
                s = Grid1.TextMatrix(i, 4) & Chr(9)
                s = s & Grid1.TextMatrix(i, 6) & Chr(9)             'jv121415
                s = s & Grid1.TextMatrix(i, 7) & Chr(9)             'jv121415
                s = s & Grid1.TextMatrix(i, 8) & Chr(9)             'jv121415
                s = s & Grid1.TextMatrix(i, 9) & Chr(9)
                s = s & Grid1.TextMatrix(i, 10) & Chr(9)
                s = s & Grid1.TextMatrix(i, 11) & Chr(9)
                s = s & Grid1.TextMatrix(i, 13) & Chr(9)
                s = s & Grid1.TextMatrix(i, 14) & Chr(9)
                s = s & Grid1.TextMatrix(i, 15)
                Grid2.AddItem s
            End If
        End If
        
        
        
    Next i
    If Grid2.Rows > 1 Then
        For i = 1 To Grid2.Rows - 1
            For k = Grid1.Rows - 1 To 1 Step -1
                If sortdate.Checked = True Then
                    If Grid1.TextMatrix(k, 1) = Grid2.TextMatrix(i, 0) Then
                        If k = Grid1.Rows - 1 Then
                            Grid1.AddItem " "
                        Else
                            Grid1.AddItem " ", k + 1
                        End If
                        Grid1.TextMatrix(k + 1, 5) = "Daily Total " & Grid2.TextMatrix(i, 0)
                        Grid1.TextMatrix(k + 1, 6) = Grid2.TextMatrix(i, 1)
                        Grid1.TextMatrix(k + 1, 7) = Grid2.TextMatrix(i, 2)
                        Grid1.TextMatrix(k + 1, 8) = Grid2.TextMatrix(i, 3)
                        Grid1.TextMatrix(k + 1, 9) = Grid2.TextMatrix(i, 4)
                        Grid1.TextMatrix(k + 1, 10) = Grid2.TextMatrix(i, 5)
                        Grid1.TextMatrix(k + 1, 11) = Grid2.TextMatrix(i, 6)
                        Grid1.TextMatrix(k + 1, 13) = Grid2.TextMatrix(i, 7)
                        Grid1.TextMatrix(k + 1, 14) = Grid2.TextMatrix(i, 8)
                        Grid1.TextMatrix(k + 1, 15) = Grid2.TextMatrix(i, 9)
                        Exit For
                    End If
                Else
                    If Grid1.TextMatrix(k, 4) = Grid2.TextMatrix(i, 0) Then
                        If k = Grid1.Rows - 1 Then
                            Grid1.AddItem " "
                        Else
                            Grid1.AddItem " ", k + 1
                        End If
                        Grid1.TextMatrix(k + 1, 5) = "SKU Total - " & Grid2.TextMatrix(i, 0)
                        Grid1.TextMatrix(k + 1, 6) = Grid2.TextMatrix(i, 1)
                        Grid1.TextMatrix(k + 1, 7) = Grid2.TextMatrix(i, 2)
                        Grid1.TextMatrix(k + 1, 8) = Grid2.TextMatrix(i, 3)
                        Grid1.TextMatrix(k + 1, 9) = Grid2.TextMatrix(i, 4)
                        Grid1.TextMatrix(k + 1, 10) = Grid2.TextMatrix(i, 5)
                        Grid1.TextMatrix(k + 1, 11) = Grid2.TextMatrix(i, 6)
                        Grid1.TextMatrix(k + 1, 13) = Grid2.TextMatrix(i, 7)
                        Grid1.TextMatrix(k + 1, 14) = Grid2.TextMatrix(i, 8)
                        Grid1.TextMatrix(k + 1, 15) = Grid2.TextMatrix(i, 9)
                        Exit For
                    End If
                End If
            Next k
        Next i
    End If
                        
                        
    Grid1.AddItem " "
    Grid1.AddItem " "
    Grid1.TextMatrix(Grid1.Rows - 1, 5) = "Pallet Summary"
    Grid1.TextMatrix(Grid1.Rows - 1, 6) = t6
    Grid1.TextMatrix(Grid1.Rows - 1, 7) = t7
    Grid1.TextMatrix(Grid1.Rows - 1, 8) = t8
    Grid1.TextMatrix(Grid1.Rows - 1, 9) = t9
    Grid1.TextMatrix(Grid1.Rows - 1, 10) = t10
    Grid1.TextMatrix(Grid1.Rows - 1, 11) = t11
    Grid1.TextMatrix(Grid1.Rows - 1, 13) = t13
    Grid1.TextMatrix(Grid1.Rows - 1, 14) = t14
    Grid1.TextMatrix(Grid1.Rows - 1, 15) = t15
    For i = 1 To Grid1.Rows - 1
        If Left(Grid1.TextMatrix(i, 5), 11) = "Daily Total" Or Grid1.TextMatrix(i, 5) = "Pallet Summary" Or Left(Grid1.TextMatrix(i, 5), 9) = "SKU Total" Then
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellBackColor = Grid1.BackColorFixed
        End If
        ' If Hold > WMS Qty Then...
        If Val(Grid1.TextMatrix(i, 14)) > Val(Grid1.TextMatrix(i, 13)) Then
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 14: Grid1.ColSel = 14
            Grid1.CellBackColor = rcolor.BackColor
            Grid1.CellForeColor = rcolor.ForeColor
            Grid1.TextMatrix(i, 14) = Grid1.TextMatrix(i, 13)
            rcolor.Visible = True
        End If
    Next i
    DoEvents
    Grid1.Row = 1: Grid1.Col = 1
    Grid1.Redraw = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.Description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_tickets", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_tickets - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_vlists()
    Combo1.Clear: List1.Clear: List2.Clear
    Combo1.AddItem "T10": List1.AddItem "500": List2.AddItem "Brenham"
    Combo1.AddItem "K10": List1.AddItem "501": List2.AddItem "Broken Arrow"
    Combo1.AddItem "A10": List1.AddItem "502": List2.AddItem "Sylacauga"
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    List2.ListIndex = Combo1.ListIndex
    Label2 = List2
    Grid1_RowColChange
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub Command2_Click()
    batchtktinv.batchno = Grid1.TextMatrix(Grid1.Row, 0)
    batchtktinv.bbarcode = Grid1.TextMatrix(Grid1.Row, 12)
    batchtktinv.bproduct = Grid1.TextMatrix(Grid1.Row, 5)
    batchtktinv.Show
End Sub

Private Sub Form_Load()
    refresh_vlists
    Combo1.ListIndex = 0
    Text1 = Format(DateAdd("d", -14, Now), "M-d-yyyy")
    Text2 = Format(Now, "M-d-yyyy")
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 17
    If Combo1 = "A10" Then
        Grid1.FormatString = "^Batch No|^Plan Start|^Status|<Location|^SKU|<Description|^Planned|^Released|^Diff|^PalPlan|^PalAct|^PalDiff|^Flag|^WMS Qty|^Hold|^Avail|^New Pool"
    Else
        Grid1.FormatString = "^Batch No|^Plan Start|^Status|<Location|^SKU|<Description|^Planned|^Released|^Diff|^PalPlan|^PalAct|^PalDiff|^BarCode|^WMS Qty|^Hold|^Avail|^New Pool"
    End If
    Grid1.ColWidth(0) = 900
    Grid1.ColWidth(1) = 1100
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 2000
    Grid1.ColWidth(4) = 700
    Grid1.ColWidth(5) = 2200
    Grid1.ColWidth(6) = 900
    Grid1.ColWidth(7) = 900
    Grid1.ColWidth(8) = 900
    Grid1.ColWidth(9) = 800
    Grid1.ColWidth(10) = 800
    Grid1.ColWidth(11) = 800
    Grid1.ColWidth(12) = 1400
    Grid1.ColWidth(13) = 900
    Grid1.ColWidth(14) = 900
    Grid1.ColWidth(15) = 900
    Grid1.ColWidth(16) = 1100
    rcolor.Visible = False
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Command1.Height * 3.5)
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid1_RowColChange()
    Command2.Visible = False
    impmenu.Visible = False
    mnq.Enabled = False
    'If Combo1 = "T10" And Val(Grid1.TextMatrix(Grid1.Row, 0)) > 0 Then
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) > 0 Then
        If Combo1 = "T10" And Left(Grid1.TextMatrix(Grid1.Row, 3), 3) = "TX." Then
            Command2.Visible = True
            impmenu.Visible = True
            mnq.Enabled = True
        End If
        If Combo1 = "K10" And Left(Grid1.TextMatrix(Grid1.Row, 3), 3) = "OK." Then
            Command2.Visible = True
            impmenu.Visible = True
            mnq.Enabled = True
        End If
        If Combo1 = "A10" And Left(Grid1.TextMatrix(Grid1.Row, 3), 3) = "AL." Then
            Command2.Visible = True
            impmenu.Visible = True
            mnq.Enabled = True
        End If
    End If
End Sub

Private Sub impwms_Click()
    Dim i As Integer
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 0)) > 0 And Val(Grid1.TextMatrix(i, 7)) = 0 Then
            Grid1.Row = i
            Command2_Click
            'DoEvents
            Call batchtktinv.paste2bat_Click
        End If
    Next i
End Sub

Private Sub mnq_Click()
    Dim p As String, psku As String, ptot As String, ds As ADODB.Recordset, i As Integer, s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then
        If Grid1.Row < Grid1.Rows - 1 Then Grid1.Row = Grid1.Row + 1
        Exit Sub
    End If
    If Val(Grid1.TextMatrix(Grid1.Row, 15)) > 0 Then
        s = Grid1.TextMatrix(Grid1.Row, 15)                         'WMS Avail
    Else
        If Val(Grid1.TextMatrix(Grid1.Row, 13)) > 0 Then
            s = Grid1.TextMatrix(Grid1.Row, 13)                     'WMS Qty
        Else
            If Val(Grid1.TextMatrix(Grid1.Row, 10)) > 0 Then
                s = Grid1.TextMatrix(Grid1.Row, 10)                 'PalAct
            Else
                s = Grid1.TextMatrix(Grid1.Row, 9)                  'PalPlan
            End If
        End If
    End If
    p = InputBox("New Available Pallets:", "Mark New Pallets Available...", s)
    If Len(p) = 0 Then Exit Sub
    Grid1.TextMatrix(Grid1.Row, 16) = Val(p)
    psku = Grid1.TextMatrix(Grid1.Row, 4)
    'ptot = 0
    'For i = 1 To Grid1.Rows - 1
    '    If Grid1.TextMatrix(i, 4) = psku Then
    '        ptot = ptot + Val(Grid1.TextMatrix(i, 16))
    '        'MsgBox Grid1.TextMatrix(i, 0) & " " & ptot
    '    End If
    'Next i
    's = "Update bimp set poolsched = " & ptot
    's = s & " Where plantwhs = '" & Combo1 & "' and sku = '" & psku & "'"
    ''MsgBox s
    'wdb.Execute s
    s = "Select * from poolschedule where batchno = " & Grid1.TextMatrix(Grid1.Row, 0)
    s = s & " and plantwhs = '" & Combo1 & "'"                  'jv020816
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update poolschedule set palqty = " & Grid1.TextMatrix(Grid1.Row, 16)
        s = s & " Where batchno = " & Grid1.TextMatrix(Grid1.Row, 0)
        s = s & " and plantwhs = '" & Combo1 & "'"
        'MsgBox s
        wdb.Execute s
    Else
        s = "Insert into poolschedule (batchno, plantwhs, sku, proddate, location, barcode, palqty)"
        s = s & " Values (" & Grid1.TextMatrix(Grid1.Row, 0)                'batchno
        s = s & ", '" & Combo1 & "'"                                        'plantwhs
        s = s & ", '" & Grid1.TextMatrix(Grid1.Row, 4) & "'"                'sku
        s = s & ", '" & Grid1.TextMatrix(Grid1.Row, 1) & "'"                'proddate
        s = s & ", '" & Grid1.TextMatrix(Grid1.Row, 3) & "'"                'location
        s = s & ", '" & Grid1.TextMatrix(Grid1.Row, 12) & "'"               'barcode
        s = s & ", " & p & ")"                                              'palqty
        'MsgBox s
        wdb.Execute s
    End If
    ds.Close
    ptot = 0
    s = "select palqty from poolschedule where plantwhs = '" & Combo1 & "'"
    s = s & " and sku = '" & psku & "'"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            ptot = ptot + ds!palqty
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "Update bimp set poolsched = " & ptot
    s = s & " Where plantwhs = '" & Combo1 & "' and sku = '" & psku & "'"
    'MsgBox s
    wdb.Execute s
    s = "delete from poolschedule where palqty = 0"
    'MsgBox s
    wdb.Execute s
    If Grid1.Row < Grid1.Rows - 1 Then Grid1.Row = Grid1.Row + 1
End Sub

Private Sub prtgrid_Click()
    Dim rt As String, rf As String, rh As String
    Dim i As Integer, c6 As Long, c7 As Long, c8 As Long, c12 As Long, c16 As Long
    rt = Me.Caption & " " & Combo1 & " " & Label2
    rh = Text1 & " thru " & Text2
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    'htdc(0) = "Yellow": gndc(0) = Grid1.BackColorFixed
    htdc(0) = "Pink": gndc(0) = Grid1.BackColorFixed
    c6 = Grid1.ColWidth(6)
    c7 = Grid1.ColWidth(7)
    c8 = Grid1.ColWidth(8)
    If MsgBox("Print pallet columns only?", vbYesNo + vbQuestion, "Pallet totals...") = vbYes Then
        Grid1.ColWidth(6) = 0
        Grid1.ColWidth(7) = 0
        Grid1.ColWidth(8) = 0
    End If
    c12 = Grid1.ColWidth(12): Grid1.ColWidth(12) = 0
    s = "c6=" & c6 & vbCrLf & "c7=" & c7 & vbCrLf & "c8=" & c8 & vbCrLf & "c12=" & c12 & vbCrLf
    'MsgBox s, vbOKOnly + vbInformation, "Column sizes...."
    On Error Resume Next
    'c16 = Grid1.ColWidth(16): Grid1.ColWidth(16) = 0
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
        Grid1.ColWidth(6) = c6
        Grid1.ColWidth(7) = c7
        Grid1.ColWidth(8) = c8
        Grid1.ColWidth(12) = c12
        'Grid1.ColWidth(16) = c16
    Else
        Grid1.Redraw = False
        Call htmlcolorgrid(Me, "u:\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        Grid1.ColWidth(6) = c6
        Grid1.ColWidth(7) = c7
        Grid1.ColWidth(8) = c8
        Grid1.ColWidth(12) = c12
        'Grid1.ColWidth(16) = c16
        Grid1.Redraw = True
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe u:\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe u:\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
End Sub

Private Sub sortdate_Click()
    sortdate.Checked = True
    sortsku.Checked = False
End Sub

Private Sub sortsku_Click()
    sortdate.Checked = False
    sortsku.Checked = True
End Sub
