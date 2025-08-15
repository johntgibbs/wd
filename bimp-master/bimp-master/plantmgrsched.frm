VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form plantmgrsched 
   Caption         =   "Planned Production"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12091
      _Version        =   327680
      ForeColor       =   16711680
      BackColorFixed  =   12640511
      FocusRect       =   0
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Days of Supply Onhand:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8520
      TabIndex        =   11
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sales - Last 30 Days:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Available Units:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label daysupply 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "daysupply"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11280
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label sales30 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "sales30"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label unitsoh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "unitsoh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Product:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plant Code:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label prodname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "prodname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   3
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label prodcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "prodcode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.Label plantcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "plantcode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "plantmgrsched"
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
    Dim toh As Long, tpct As Currency
    
    If plant_server_status(plantcode) = False Then                                      'jv010417
        s = "Sorry, The server for Warehouse " & prodbatches.Combo1 & " has been flagged to be offline."
        MsgBox s, vbOKOnly + vbInformation, "sorry, try again later..."                 'jv010417
        Exit Sub                                                                        'jv010417
    End If                                                                              'jv010417
    
    
    startdate = Format(DateAdd("d", -14, Now), "MM-dd-yyyy")
    enddate = Format(DateAdd("d", 5, Now), "MM-dd-yyyy")
    
    
    'rcolor.Visible = False
    wdlot = Right(startdate, 2)
    wdlot = wdlot & Format(DateDiff("d", "1-1-" & Right(startdate, 4), startdate) + 1, "000")
    'MsgBox wdlot
    
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    
    'On Error GoTo vberror
    Screen.MousePointer = 11
    'Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 10
    'q = "select h.batch_no,TO_CHAR(h.plan_start_date,'MM-DD-YYYY'),h.batch_status,"
    q = "select h.batch_no,TO_CHAR(h.plan_start_date,'YYYY-MM-DD'),h.batch_status,"         'jv010516
    q = q & "h.attribute1,i.segment1,i.description,d.plan_qty,"
    q = q & "d.actual_qty"
    q = q & " from apps.gme_batch_header h, apps.gme_material_details d, apps.mtl_system_items_b i"
    If plantcode = "T10" Then
        q = q & " where h.organization_id in (select organization_id from mtl_parameters where organization_code in ('500','503'))"
    Else
        If plantcode = "K10" Then
            q = q & " where h.organization_id in (select organization_id from mtl_parameters where organization_code in ('501'))"
        Else
            q = q & " where h.organization_id in (select organization_id from mtl_parameters where organization_code in ('502'))"
        End If
    End If
    q = q & " and h.plan_start_date >= TO_DATE('" & Format(startdate, "DD-MMM-YYYY") & "')"
    q = q & " and h.plan_start_date <= TO_DATE('" & Format(DateAdd("d", 1, enddate), "DD-MMM-YYYY") & "')"
    q = q & " and h.delete_mark = 0"
    q = q & " and h.batch_id = d.batch_id"
    q = q & " and h.batch_status in (1, 2, 3, 4)"
    q = q & " and d.line_type = 1"
    q = q & " and d.actual_qty = 0"
    q = q & " and i.organization_id = d.organization_id"
    q = q & " and i.inventory_item_id = d.inventory_item_id"
    'q = q & " and i.segment1 >= '100' and i.segment1 <= '9999'"             'jv082415
    q = q & " and i.segment1 >= '" & prodcode.Caption & "' and i.segment1 <= '" & prodcode.Caption & "x'"
    'If sortsku.Checked Then
        q = q & " order by i.segment1, 2, d.plan_qty desc, h.attribute1"
    'Else
    '    q = q & " order by 2, i.segment1, d.plan_qty desc, h.attribute1"
    'End If
    'MsgBox q
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            pflag = Trim(ds(4))
            If Len(pflag) = 3 Then pflag = pflag & " "
            pflag = pflag & Format(DateAdd("yyyy", 2, Format(ds(1), "M-dd-yyyy")), "MMddyy")
            pflag = pflag & Right(ds(3), 3)
            s = ds(0) & Chr(9)                              'Batch No
            s = s & Format(ds(1), "M-dd-yyyy") & Chr(9)     'Date
            If ds(2) = 1 Then s = s & "PEND" & Chr(9)       'Status
            If ds(2) = 2 Then s = s & "WIP" & Chr(9)
            If ds(2) = 3 Then s = s & "CERT" & Chr(9)
            If ds(2) = 4 Then s = s & "Closed" & Chr(9)
            s = s & ds(3) & Chr(9)                          'Location
            s = s & ds(4) & Chr(9)                          'SKU
            s = s & ds(5) & Chr(9)                          'Product Name
            s = s & pflag & Chr(9)                          'BarCode
            s = s & ds(6) & Chr(9)                          'Planned Qty
            s = s & ds(7) & Chr(9)                          'Actual Qty
            'If Trim(ds(4)) = prodcode.Caption Then Grid1.AddItem s
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    
    'MsgBox "racks"
    Set db = CreateObject("ADODB.Connection")
    If plantcode = "A10" Then db.Open a10bbsr
    If plantcode = "K10" Then db.Open k10bbsr
    If plantcode = "T10" Then db.Open t10bbsr
    s = "select sku, lot_num, barcode, count_qty from rackpos where lot_num >= '" & wdlot & "' and barcode < '9999'"
    s = s & " and rackno not in (select id from racks where rack = 'OP')"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 6) = Left(ds!barcode, 13) Then
                    Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 8)) + ds!count_qty
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    ''Rack Hold List
    's = "select * from holdlist where lot_num >= '" & wdlot & "'"
    'Set hs = db.Execute(s)
    'If hs.BOF = False Then
    '    hs.MoveFirst
    '    Do Until hs.EOF
    '        sp = hs!sku & " " & r12lot(hs!lot_num) & hs!opcode & hs!spallet
    '        ep = hs!sku & " " & r12lot(hs!lot_num) & hs!opcode & hs!epallet
    '        s = "select sku, count(*) from rackpos where barcode >= '" & sp & "'"
    '        s = s & " and barcode <= '" & ep & "' group by sku"
    '        'MsgBox s
    '        Set ds = db.Execute(s)
    '        If ds.BOF = False Then
    '            ds.MoveFirst
    '            'MsgBox ds(1), vbOKOnly, sp
    '            For i = 1 To Grid1.Rows - 1
    '                If Grid1.TextMatrix(i, 12) = Left(sp, 13) Then
    '                    Grid1.TextMatrix(i, 14) = Val(Grid1.TextMatrix(i, 14)) + ds(1)
    '                    Exit For
    '                End If
    '            Next i
    '        End If
    '        ds.Close
    '        hs.MoveNext
    '    Loop
    'End If
    'hs.Close
    
    If plantcode = "T10" Then
        s = "select barcode, count_qty from position where lot_num >= '" & wdlot & "' and barcode < '9999'"
        'MsgBox s
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                For i = 1 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 6) = Left(ds!barcode, 13) Then
                        Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 8)) + ds!count_qty
                        Exit For
                    End If
                Next i
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    db.Close
    If plantcode = "A10" Then
        db.Open cs5db
        's = "Select item, expiration, quantity from vContainerLocation_1033"
        s = "Select item, [Lot Expiration], quantity from vAllInventory_1033"       'Westfalia Upgrade
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If Len(ds(0)) > 1 Then
                    psku = Trim(ds(0))
                    plot = Trim(ds(1))
                    If Val(Mid(psku, 4, 1)) > 0 Then
                        s = Left(psku, 4)
                    Else
                        s = Left(psku, 3) & " "
                    End If
                    s = s & Format(plot, "MMddyy")
                    s = s & Right(psku, 3)
                    For i = 1 To Grid1.Rows - 1
                        If Grid1.TextMatrix(i, 6) = Left(s, 13) Then
                            Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 8)) + ds(2)
                            Exit For
                        End If
                    Next i
                    'MsgBox s
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close: db.Close
    End If
    Screen.MousePointer = 0
    
    If Grid1.Rows > 1 And Val(sales30) > 0 Then
        toh = Val(unitsoh)
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 8)) > 0 Then
                toh = toh + Val(Grid1.TextMatrix(i, 8))
                tpct = toh / Val(sales30)
                Grid1.TextMatrix(i, 9) = Format(tpct * 30, "0")
            Else
                toh = toh + Val(Grid1.TextMatrix(i, 7))
                tpct = toh / Val(sales30)
                Grid1.TextMatrix(i, 9) = Format(tpct * 30, "0")
            End If
        Next i
    End If
    If plantcode = "A10" Then
        Grid1.FormatString = "^Batch No|^Plan Start|^Status|<Location|^SKU|<Description|^Flag|^Planned|^Test Hold|^Days"
    Else
        Grid1.FormatString = "^Batch No|^Plan Start|^Status|<Location|^SKU|<Description|^BarCode|^Planned|^Test Hold|^Days"
    End If
    Grid1.ColWidth(0) = 900
    Grid1.ColWidth(1) = 1100
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 2000
    Grid1.ColWidth(4) = 700
    Grid1.ColWidth(5) = 2200
    Grid1.ColWidth(6) = 1500
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1100
    Grid1.ColWidth(9) = 900
    Grid1.Redraw = True
End Sub

Private Sub Form_Load()
    'Me.Height = whssales.Height
    Me.Top = whssales.Top
    Me.Left = whssales.Width - Me.Width
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Label1.Height * 4)
End Sub

Private Sub prodcode_Change()
    refresh_grid
End Sub

