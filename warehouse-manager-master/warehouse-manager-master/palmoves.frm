VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form15 
   Caption         =   "Pallet Movement"
   ClientHeight    =   8160
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13470
   LinkTopic       =   "Form15"
   ScaleHeight     =   8160
   ScaleWidth      =   13470
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check2 
      Caption         =   "Use Test Logs"
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   480
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View All Fields"
      Height          =   255
      Left            =   10440
      TabIndex        =   8
      Top             =   240
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   240
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   3495
      Left            =   0
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6165
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   6135
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10821
      _Version        =   327680
      BackColorSel    =   32768
      WordWrap        =   -1  'True
      FocusRect       =   0
      FillStyle       =   1
      AllowUserResizing=   3
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3600
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label sortshiptrig 
      Caption         =   "Label3"
      Height          =   255
      Left            =   9480
      TabIndex        =   10
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label ccol 
      Caption         =   "..."
      Height          =   255
      Left            =   12240
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label cntlit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8280
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Date:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Label hcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Pallet Moves:"
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
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
      Begin VB.Menu prtcur 
         Caption         =   "Current List"
      End
      Begin VB.Menu pstot 
         Caption         =   "Shipping Totals"
      End
      Begin VB.Menu ppsum 
         Caption         =   "Production Summary"
      End
      Begin VB.Menu pptot 
         Caption         =   "Production Details"
      End
      Begin VB.Menu pbartest 
         Caption         =   "BarCodes for Testing"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu sortmenu 
      Caption         =   "Sort"
      Begin VB.Menu sortbc 
         Caption         =   "BarCode"
      End
      Begin VB.Menu sortdt 
         Caption         =   "Date/Time"
      End
   End
   Begin VB.Menu findmenu 
      Caption         =   "Find"
      Begin VB.Menu findsku 
         Caption         =   "SKU"
      End
      Begin VB.Menu findcol 
         Caption         =   "Column"
      End
      Begin VB.Menu addbillbc 
         Caption         =   "Add Pallet to Bill"
      End
      Begin VB.Menu btran 
         Caption         =   "Branch Transfer"
      End
      Begin VB.Menu pcorr 
         Caption         =   "Pallet Correction"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu usermenu 
      Caption         =   "User"
      Visible         =   0   'False
      Begin VB.Menu emplook 
         Caption         =   "Lookup Employee Name"
      End
   End
   Begin VB.Menu histmenu 
      Caption         =   "History"
      Begin VB.Menu batonhand 
         Caption         =   "Batch - On Hand"
      End
      Begin VB.Menu shiphist 
         Caption         =   "Batch - Shipped"
      End
      Begin VB.Menu widrpt 
         Caption         =   "Batch - Withdrawl"
      End
      Begin VB.Menu billhist 
         Caption         =   "Bill of Lading"
      End
      Begin VB.Menu palhist 
         Caption         =   "Pallet"
      End
   End
   Begin VB.Menu addrec 
      Caption         =   "Add Record"
   End
   Begin VB.Menu postlogs 
      Caption         =   "Post Logs"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu posteorlogs 
         Caption         =   "EOR Pallets"
      End
      Begin VB.Menu ckpart 
         Caption         =   "Check Partials"
      End
      Begin VB.Menu postalllogs 
         Caption         =   "All Pallets"
      End
   End
   Begin VB.Menu psrlogs 
      Caption         =   "SR Logs"
      Enabled         =   0   'False
      Begin VB.Menu psrlogship 
         Caption         =   "Shipping"
      End
      Begin VB.Menu psrlogsques 
         Caption         =   "Queues"
      End
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function full_pallet(psku As String, qty As Integer) As Boolean
    If skurec(Val(psku)).sku = psku Then
        If qty >= skurec(Val(psku)).uom_per_pallet Then
            full_pallet = True
        Else
            full_pallet = False
        End If
    Else
        full_pallet = False
    End If
End Function

Function r12_lot(plot As String, ocode As String) As String
    Dim s As String, myear As Integer, mdays As Integer
    If Len(plot) >= 5 Then
        myear = Val(Left(plot, 2))
        mdays = Val(Mid(plot, 3, 3)) - 1
        s = "1-1-20" & Left(plot, 2)
        s = Format(DateAdd("d", mdays, s), "MMddyy")
        s = Left(s, 4) & Format(myear + 2, "00")
        If Len(plot) > 5 Then               'jv080315
            s = s & Right(plot, 3)          'jv080315
        Else                                'jv080315
            s = s & ocode                   'jv080315
        End If                              'jv080315
    Else
        s = " "
    End If
    r12_lot = s
End Function

Private Sub fetch_bill_of_lading()
    Dim plit As String
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String, wsku As String, wlot As String
    Dim wzone As String, wstat As String, wgma As Integer, wside As String
    Dim waisle As String, wrack As String, hrow As Boolean, r12flag As Boolean, ocode As String
    Dim cfile As String, s As String, bc As String, srflag As Boolean
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim dl As Long, wbc As String
    Dim syear As Integer, eyear As Integer, i As Integer                        'jv061215
    Dim logpath As String
    logpath = Form1.logdir

    Dim bt As String
    bt = InputBox("Batch Ticket # on Bill of Lading", "Batch Ticket", "")
    If Len(bt) = 0 Then Exit Sub
    cfile = Form1.logdir & "RO" & bt & ".txt"
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15
        Close #1
        sdate = f12
        'MsgBox sdate
        Text1 = sdate
    Else
        MsgBox "Batch ticket was not found on server.", vbOKOnly + vbInformation, "sorry, try again..."
        Exit Sub
    End If
    
    hcolor.Caption = "Bill of Lading"
    Screen.MousePointer = 11
    grid1.Clear: grid1.Cols = 19: grid1.Rows = 1
    grid1.Redraw = False
    
    cfile = Form1.logdir & "bill" & Format(sdate, "MMddyyyy") & ".txt"
    If Len(Dir(cfile)) > 0 Then
        'MsgBox cfile
        Open cfile For Input Shared As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
            If f16 = bt Then
                s = "B" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                grid1.AddItem s
            End If
                
        Loop
        Close #1
    End If
    
    s = "^Type|^RecId|<Area|<Description|<Source|<Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^Ticket"
    grid1.FormatString = s
    grid1.ColWidth(0) = 600
    grid1.ColWidth(1) = 1 '600
    grid1.ColWidth(2) = 1300
    grid1.ColWidth(3) = 1000
    grid1.ColWidth(4) = 1300
    grid1.ColWidth(5) = 1300
    grid1.ColWidth(6) = 3000
    grid1.ColWidth(7) = 1800
    grid1.ColWidth(8) = 600
    grid1.ColWidth(9) = 800
    grid1.ColWidth(10) = 900
    grid1.ColWidth(11) = 800
    grid1.ColWidth(12) = 900
    grid1.ColWidth(13) = 800
    grid1.ColWidth(14) = 1 '800
    grid1.ColWidth(15) = 1000
    grid1.ColWidth(16) = 1400
    grid1.ColWidth(17) = 1000
    grid1.ColWidth(18) = 1 '2100
    grid1.RowSel = grid1.Row
    grid1.Col = 16: grid1.ColSel = 16
    grid1.Sort = 5
    grid1.FillStyle = flexFillRepeat
    If grid1.Rows > 2 Then
        s = grid1.TextMatrix(1, 7)
        For i = 1 To grid1.Rows - 1
            If grid1.TextMatrix(i, 7) <> s Then
                hrow = Not hrow
                s = grid1.TextMatrix(i, 7)
            End If
            If hrow = True Then
                grid1.Row = i: grid1.RowSel = i
                grid1.Col = 1: grid1.ColSel = 17
                grid1.CellBackColor = cntlit.BackColor
            End If
        Next i
        grid1.Row = 1: grid1.Col = 2
    End If
    grid1.Redraw = True
            
    cntlit.Caption = grid1.Rows - 1 & " Records"
    addbillbc.Visible = True
    btran.Visible = False
    Screen.MousePointer = 0
End Sub

Private Sub pallet_history()
    Dim ds As ADODB.Recordset, ss As ADODB.Recordset, plit As String
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String, wsku As String, wlot As String
    Dim wzone As String, wstat As String, wgma As Integer, wside As String
    Dim waisle As String, wrack As String, hrow As Boolean, r12flag As Boolean, ocode As String
    Dim cfile As String, s As String, bc As String, srflag As Boolean
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim dl As Long, wbc As String, citem As String
    Dim syear As Integer, eyear As Integer, i As Integer
    Dim logpath As String
    Dim db5 As ADODB.Connection, ds5 As ADODB.Recordset, ds6 As ADODB.Recordset
    logpath = Form1.logdir
    'srpath = "C:\"                                          'jv060117
    srpath = logpath                                            'jv060117
    's = grid1.TextMatrix(grid1.Row, 7)
    'sdate = Format(Val(Mid(s, 9, 2)) - 2, "00")
    'sdate = "20" & sdate & Mid(s, 5, 4)
    'edate = Format(Now, "yyyymmdd")
    wbc = grid1.TextMatrix(grid1.Row, 7)
    wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
    wbc = InputBox("Enter a BarCode:", "BarCode Example....", wbc)
    If Len(wbc) = 0 Then Exit Sub
    wsku = Trim(Left(wbc, 4))
    wlot = barcode_to_lotnum(wbc)
    r12flag = True
    s = wbc                                                             'jv012116
    sdate = Format(Val(Mid(s, 9, 2)) - 2, "00")                         'jv012116
    sdate = "20" & sdate & Mid(s, 5, 4)                                 'jv012116
    edate = Format(Now, "yyyymmdd")                                     'jv012116
    
    hcolor.Caption = "Pallet History"
    Screen.MousePointer = 11
    grid1.Clear: grid1.Cols = 19: grid1.Rows = 1
    
    'Current location
    If skurec(Val(wsku)).sku = wsku Then
        plit = wsku & " " & skurec(Val(wsku)).prodname
    Else
        plit = wsku & " Undefined SKU"
    End If
    If Form1.plantno = "50" Then
        s = "select * from position where barcode = '" & wbc & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                wzone = "0": wstat = " ": wgma = 0: wside = " "
                s = "select zone_num, rack_side, lane_status, gmasize from lane where id = " & ds!laneno
                Set ss = Wdb.Execute(s)
                If ss.BOF = False Then
                    ss.MoveFirst
                    wzone = ss!zone_num
                    wstat = ss!lane_status
                    wgma = ss!gmasize
                    wside = ss!rack_side
                End If
                ss.Close
                s = "OH" & Chr(9)                               'Type
                s = s & ds!id & Chr(9)                          'Recid
                s = s & "Crane" & Chr(9)                        'Area
                If wstat = "H" Then s = s & "On Hold"
                If wstat = "B" Then s = s & "Blocked"
                s = s & " " & Chr(9)                            'Description
                s = s & "SR-" & ds!whse_num & Chr(9)            'Source
                If ds!whse_num < 4 Then                         'Target
                    s = s & ds!vert_loc & "-" & ds!horz_loc & "-" & ds!rack_side & " " & ds!posn_num & Chr(9)
                Else
                    s = s & wzone & " " & ds!vert_loc & "-" & ds!horz_loc & "-" & wside & Chr(9)
                End If
                s = s & plit & Chr(9)                           'Product
                s = s & bc000(ds!barcode) & Chr(9)                     'Pallet
                s = s & "1" & Chr(9)                            'Qty
                If wgma = 0 Then
                    s = s & "BBC" & Chr(9)                       'Uom
                Else
                    s = s & "GMA" & Chr(9)
                End If
                ocode = Mid(ds!barcode, 11, 3)                              'jv071715
                If r12flag = True Then
                    s = s & Mid(ds!barcode, 5, 9) & Chr(9)                  'jv071615
                Else
                    s = s & ds!lot_num & Chr(9)                 'Lot1
                End If
                s = s & ds!count_qty & Chr(9)                   'Units
                If r12flag = True Then
                    s = s & r12_lot(ds!lot2, ocode) & Chr(9)
                Else
                    s = s & ds!lot2 & Chr(9)                    'Lot2
                End If
                s = s & ds!qty2 & Chr(9)                        'Units
                s = s & "In-Stock" & Chr(9)                     'Status
                s = s & "WMS" & Chr(9)                          'User
                s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
                s = s & " "                                     'Reqid
                grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    
    If Form1.plantno = "52" Then
        citem = Trim(Left(wbc, 4)) & "-" & Mid(wbc, 11, 3)                              'jv100516
        Set db5 = CreateObject("ADODB.Connection")
        'db5.Open "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
        db5.Open "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=sywms;PWD=!Sylacauga_WMS1907"
        s = "select * from pallets where barcode = '" & wbc & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                's = "select * from vContainerLocation_1033 Where [pal id] = '" & ds!plateno & "'"   'jv081415
                's = "select * from vContainerLocation_1033 Where [pal id] in "      'jv092816
                s = "select * from vAllInventory_1033 Where LPN in "      'westfalia update
                s = s & "('" & ds!plateno & "', '" & ds!barcode & "')"              'jv092816
                s = s & " and item = '" & citem & "'"                                   'jv100516
                Set ds6 = db5.Execute(s)
                If ds6.BOF = False Then
                    ds6.MoveFirst
                    s = "OH" & Chr(9)
                    s = s & ds6(0) & Chr(9)         'container id
                    s = s & "Crane" & Chr(9)        'area
                    's = s & ds6(17) & Chr(9)        'Hold reason
                    If ds6(16) > "0" Then          'westfalia update
                        s = s & "Locked" & Chr(9)
                    Else
                        s = s & " " & Chr(9)
                    End If
                    s = s & "CS5" & Chr(9)          'Source
                    s = s & ds6!location & Chr(9)   'Target
                    s = s & plit & Chr(9)           'Product
                    s = s & bc000(ds!barcode) & Chr(9)     'Pallet
                    s = s & "1" & Chr(9)            'Qty
                    'If Trim(ds6!Type) = "BBCPallet" Then  'UOM
                    If Trim(ds6![Pallet Type]) = "BBCPallet" Then  'UOM        westalia update
                        s = s & "BBC" & Chr(9)
                    Else
                        s = s & "GMA" & Chr(9)
                    End If
                    ocode = Mid(ds!barcode, 11, 3)                              'jv071715
                    If r12flag = True Then
                        s = s & Mid(ds!barcode, 5, 9) & Chr(9)                  'jv071615
                    Else
                        s = s & ds!lot1 & Chr(9)                 'Lot1
                    End If
                    s = s & ds!qty1 & Chr(9)                   'Units
                    If r12flag = True Then
                        s = s & r12_lot(ds!lot2, ocode) & Chr(9)
                    Else
                        s = s & ds!lot2 & Chr(9)                    'Lot2
                    End If
                    s = s & ds!qty2 & Chr(9)                        'Units
                    s = s & "In-Stock" & Chr(9)                     'Status
                    s = s & "WMS" & Chr(9)                          'User
                    s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
                    s = s & ds!plateno                                     'Reqid
                    grid1.AddItem s
                End If
                ds6.Close
                ds.MoveNext
            Loop
        End If
        ds.Close
        db5.Close
    End If
    
    s = "select * from rackpos where barcode = '" & wbc & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            waisle = " ": wstat = " ": wrack = " "
            s = "select aisle, rack, hold from racks where id = " & ds!rackno
            Set ss = Wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                waisle = Trim(ss!aisle)
                wrack = Trim(ss!rack)
                If ss!hold = 0 Then
                    wstat = " "
                Else
                    wstat = "On Hold"
                End If
            End If
            ss.Close
            s = "OH" & Chr(9)                               'Type
            s = s & ds!id & Chr(9)                          'Recid
            s = s & "Racks" & Chr(9)                        'Area
            s = s & wstat & Chr(9)                          'Description
            s = s & "SR-4" & Chr(9)                         'Source
            s = s & waisle & "-" & wrack & Chr(9)           'Target
            s = s & plit & Chr(9)                           'Product
            s = s & bc000(ds!barcode) & Chr(9)                     'Pallet
            s = s & "1" & Chr(9)                            'Qty
            If ds!bbc = "Y" Then
                s = s & "BB" & Chr(9)                       'Uom
            Else
                s = s & "GMA" & Chr(9)
            End If
            ocode = Mid(ds!barcode, 11, 3)                                  'jv071715
            If r12flag = True Then
                s = s & Mid(ds!barcode, 5, 9) & Chr(9)                      'jv071715
            Else
                s = s & ds!lot_num & Chr(9)                 'Lot1
            End If
            s = s & ds!count_qty & Chr(9)                   'Units
            If r12flag = True Then
                s = s & r12_lot(ds!lot2, ocode) & Chr(9)
            Else
                s = s & ds!lot2 & Chr(9)                    'Lot2
            End If
            s = s & ds!qty2 & Chr(9)                        'Units
            s = s & "In-Stock" & Chr(9)                     'Status
            s = s & "WMS" & Chr(9)                          'User
            s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
            s = s & " "                                     'Reqid
            grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    
    'db.Close
    syear = Val(Left(sdate, 4))                                             'jv061215
    eyear = Val(Left(edate, 4))                                             'jv061215
    bc = " "
    spath = logpath & "recv*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = wbc Then       'jv080315
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    ocode = Mid(f6, 11, 3)                                  'jv071715
                    s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                    grid1.AddItem s
                End If
                
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\recv*.txt"                         'jv061215
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        ocode = Mid(f6, 11, 3)                      'jv071715
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                    End If
                
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i                                                                          'jv0612515
    
    spath = logpath & "tml*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = wbc Then       'jv080315
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "TM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    ocode = Mid(f6, 11, 3)                  'jv071715
                    s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    's = s & f14 & Chr(9) & f15 & Chr(9) & f16
                    s = s & "Traffic Master" & Chr(9) & f15 & Chr(9) & f16         'jv062117
                    grid1.AddItem s
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\tml*.txt"                          'jv061215
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "TM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        ocode = Mid(f6, 11, 3)                          'jv071715
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)   'jv071715
                        's = s & f14 & Chr(9) & f15 & Chr(9) & f16
                        s = s & "Traffic Master" & Chr(9) & f15 & Chr(9) & f16  'jv062117
                        grid1.AddItem s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    spath = logpath & "move*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)
        s = Mid(s, 5, 4) & Mid(s, 1, 4)
        fdate = s
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = wbc Then
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "M" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    ocode = Mid(f6, 11, 3)                              'jv071715
                    s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                    grid1.AddItem s
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\move*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "M" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        ocode = Mid(f6, 11, 3)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    spath = logpath & "sr4rem*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)
        s = Mid(s, 5, 4) & Mid(s, 1, 4)
        fdate = s
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = wbc Then
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "RR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    ocode = Mid(f6, 11, 3)
                    s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                    grid1.AddItem s
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\sr4rem*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "RR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        ocode = Mid(f6, 11, 3)                      'jv071715
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    spath = logpath & "ship*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)
        s = Mid(s, 5, 4) & Mid(s, 1, 4)
        fdate = s
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f9 = wlot And bc < wsku Then bc = f6
                If f6 = wbc Then
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    ocode = Mid(f6, 11, 3)
                    s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                    grid1.AddItem s
                End If
                
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\" & "\ship*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f9 = wlot And bc < wsku Then bc = f6
                    If f6 = wbc Then
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        ocode = Mid(f6, 11, 3)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    spath = logpath & "bill*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)
        s = Mid(s, 5, 4) & Mid(s, 1, 4)
        fdate = s
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f9 = wlot And bc < wsku Then bc = f6
                If f6 = wbc Then
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "B" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    ocode = Mid(f6, 11, 3)
                    s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                    grid1.AddItem s
                End If
                
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\" & "\bill*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f9 = wlot And bc < wsku Then bc = f6
                    If f6 = wbc Then
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "B" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        ocode = Mid(f6, 11, 3)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    
    spath = logpath & "wms*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)
        s = Mid(s, 5, 4) & Mid(s, 1, 4)
        fdate = s
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = wbc Then
                    If f1 = "DOCK" And f13 = "COMP" Then
                    Else
                        s = "WM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        ocode = Mid(f6, 11, 3)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                    End If
                    If f9 = wlot And bc < wsku Then bc = f6
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\wms*.txt"
        'MsgBox spath
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then
                        If f1 = "DOCK" And f13 = "COMP" Then
                        Else
                            s = "WM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                            s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                            ocode = Mid(f6, 11, 3)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                            s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                            grid1.AddItem s
                        End If
                        If f9 = wlot And bc < wsku Then bc = f6
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    spath = logpath & "pick*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)
        s = Mid(s, 5, 4) & Mid(s, 1, 4)
        fdate = s
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = wbc Then
                    s = "P" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    ocode = Mid(f6, 11, 3)
                    s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                    grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\pick*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                'Open logpath & sdir For Input Shared As #1
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then
                        s = "P" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        ocode = Mid(f6, 11, 3)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                        If f9 = wlot And bc < wsku Then bc = f6
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    srflag = False
    If Form1.plantno = 50 Then
        If MsgBox("Include SR Logs?", vbQuestion + vbYesNo, "SR Logs....") = vbYes Then srflag = True
    End If
    If srflag = True Then
    spath = srpath & "sr*.txt"                                     'jv060117
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)
        s = Mid(s, 5, 4) & Mid(s, 1, 4)
        fdate = s
        If fdate >= sdate And fdate <= edate Then
            Open srpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 = wbc Then
                    s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    ocode = Mid(f6, 11, 3)
                    s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                    s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                    grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                    'MsgBox srpath & sdir
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = srpath & Format(i, "0000") & "\sr*.txt"            'jv060117
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)
            s = Mid(s, 5, 4) & Mid(s, 1, 4)
            fdate = s
            If fdate >= sdate And fdate <= edate Then
                Open srpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then
                        s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        ocode = Mid(f6, 11, 3)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        s = s & wdempname(f14) & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                        If f9 = wlot And bc < wsku Then bc = f6
                        'MsgBox srpath & Format(i, "0000") & "\" & sdir
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    End If
    
    grid1.Redraw = False
    If grid1.Rows > 1 Then
        For i = 1 To grid1.Rows - 1
            If grid1.TextMatrix(i, 0) = "PR" Then
                grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7) & "0" & grid1.TextMatrix(i, 16)
            Else
                grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7) & grid1.TextMatrix(i, 16) & grid1.TextMatrix(i, 0)
            End If
        Next i
    End If
    
    s = "^Type|^RecId|<Area|<Description|<Source|<Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId"
    grid1.FormatString = s
    grid1.ColWidth(0) = 600
    grid1.ColWidth(1) = 1 '600
    grid1.ColWidth(2) = 1300
    grid1.ColWidth(3) = 1000
    grid1.ColWidth(4) = 1300
    grid1.ColWidth(5) = 1300
    grid1.ColWidth(6) = 3000
    grid1.ColWidth(7) = 1800
    grid1.ColWidth(8) = 600
    grid1.ColWidth(9) = 800
    grid1.ColWidth(10) = 900
    grid1.ColWidth(11) = 800
    grid1.ColWidth(12) = 900
    grid1.ColWidth(13) = 800
    grid1.ColWidth(14) = 1 '800
    grid1.ColWidth(15) = 1600
    grid1.ColWidth(16) = 1400
    grid1.ColWidth(17) = 1 '1000
    grid1.ColWidth(18) = 1 '2100
    grid1.RowSel = grid1.Row
    grid1.Col = 16: grid1.ColSel = 16
    grid1.Sort = 5
    grid1.FillStyle = flexFillRepeat
    If grid1.Rows > 2 Then
        s = grid1.TextMatrix(1, 16)
        For i = 1 To grid1.Rows - 1
            If grid1.TextMatrix(i, 16) <> s Then
                hrow = Not hrow
                s = grid1.TextMatrix(i, 16)
            End If
            If hrow = True Then
                grid1.Row = i: grid1.RowSel = i
                grid1.Col = 1: grid1.ColSel = 16
                grid1.CellBackColor = cntlit.BackColor
            End If
        Next i
        grid1.Row = 1
    End If
    grid1.Redraw = True
            
    cntlit.Caption = grid1.Rows - 1 & " Records"
    Screen.MousePointer = 0
End Sub

Private Sub ship_history()                  'jv012016
    Dim plit As String
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String, wsku As String, wlot As String
    Dim wzone As String, wstat As String, wgma As Integer, wside As String
    Dim waisle As String, wrack As String, hrow As Boolean, r12flag As Boolean, ocode As String
    Dim cfile As String, s As String, bc As String, srflag As Boolean
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim dl As Long, wbc As String
    Dim k10path As String, a10path As String, t10path As String, opcode As String
    Dim syear As Integer, eyear As Integer, i As Integer
    Dim logpath As String
    logpath = Form1.logdir
    k10path = "\\bbba-03-dc\f\user\waredist\data\pallogs\"
    a10path = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
    t10path = "\\bbc-01-prodtrk\wd\pallogs\"
    's = grid1.TextMatrix(grid1.Row, 7)
    'sdate = Format(Val(Mid(s, 9, 2)) - 2, "00")
    'sdate = "20" & sdate & Mid(s, 5, 4)
    'edate = Format(Now, "yyyymmdd")
    wbc = grid1.TextMatrix(grid1.Row, 7)
    wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
    wbc = InputBox("Enter a BarCode:", "BarCode Example....", wbc)
    If Len(wbc) = 0 Then Exit Sub
    wsku = Trim(Left(wbc, 4))
    wlot = barcode_to_lotnum(wbc)
    'r12flag = True
    r12flag = False
    opcode = Mid(wbc, 11, 3)
    'MsgBox opcode
    'MsgBox wbc & ", " & wsku & ", " & wlot & ", " & opcode
    s = wbc                                                             'jv012116
    sdate = Format(Val(Mid(s, 9, 2)) - 2, "00")                         'jv012116
    sdate = "20" & sdate & Mid(s, 5, 4)                                 'jv012116
    edate = Format(Now, "yyyymmdd")                                     'jv012116
    
    syear = Val(Left(sdate, 4))                                             'jv061215
    eyear = Val(Left(edate, 4))                                             'jv061215

    hcolor.Caption = "Ship History"
    Screen.MousePointer = 11
    grid1.Clear: grid1.Cols = 19: grid1.Rows = 1

    spath = logpath & "ship*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f9 = wlot And bc < wsku Then bc = f6
                If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                    If f9 = wlot And bc < wsku Then bc = f6
                    's = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = "S" & Chr(9) & f0 & Chr(9)                                              'jv012616
                    If f1 = "DOCK" Then                                                         'jv012616
                        If logpath = t10path Then f3 = "T10-" & Trim(f3)                        'jv012616
                        If logpath = k10path Then f3 = "K10-" & Trim(f3)                        'jv012616
                        If logpath = a10path Then f3 = "A10-" & Trim(f3)                        'jv012616
                        s = s & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)                   'jv012616
                    Else                                                                        'jv012616
                        s = s & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)                   'jv012616
                    End If                                                                      'jv012616
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    If r12flag = True Then
                        ocode = Mid(f6, 11, 3)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    Else
                        s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                    End If
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                    grid1.AddItem s
                End If
                
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear                                                      'jv061215
        'MsgBox syear & " - " & eyear
        spath = logpath & Format(i, "0000") & "\" & "\ship*.txt"                'jv061215
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f9 = wlot And bc < wsku Then bc = f6
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = "S" & Chr(9) & f0 & Chr(9)                                              'jv012616
                        If f1 = "DOCK" Then                                                         'jv012616
                            If logpath = t10path Then f3 = "T10-" & Trim(f3)                        'jv012616
                            If logpath = k10path Then f3 = "K10-" & Trim(f3)                        'jv012616
                            If logpath = a10path Then f3 = "A10-" & Trim(f3)                        'jv012616
                            s = s & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)                   'jv012616
                        Else                                                                        'jv012616
                            s = s & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)                   'jv012616
                        End If                                                                      'jv012616
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            ocode = Mid(f6, 11, 3)              'jv071715
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    ' ------------------ Sylacauga OP Code ------------------
    If Val(opcode) >= 200 And Val(opcode) <= 299 And Form1.plantno <> 52 Then
        spath = a10path & "ship*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            If fdate >= sdate And fdate <= edate Then
                Open a10path & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f9 = wlot And bc < wsku Then bc = f6
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "S" & Chr(9) & f0 & Chr(9)                                              'jv012616
                        If f1 = "DOCK" Then                                                         'jv012616
                            s = s & f1 & Chr(9) & f2 & Chr(9) & "A10-" & Trim(f3) & Chr(9)          'jv012616
                        Else                                                                        'jv012616
                            s = s & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)                   'jv012616
                        End If                                                                      'jv012616
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            ocode = Mid(f6, 11, 3)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                    End If
                
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
        For i = syear To eyear                                                      'jv061215
            spath = a10path & Format(i, "0000") & "\" & "\ship*.txt"                'jv061215
            sdir = Dir$(spath)
            Do While sdir <> ""
                s = Right(sdir, 12)                                                 'jv061215
                s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
                fdate = s                                                           'jv061215
                If fdate >= sdate And fdate <= edate Then
                    Open a10path & Format(i, "0000") & "\" & sdir For Input Shared As #1
                    Do Until EOF(1)
                        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                        If f9 = wlot And bc < wsku Then bc = f6
                        If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                            If f9 = wlot And bc < wsku Then bc = f6
                            's = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                            s = "S" & Chr(9) & f0 & Chr(9)                                              'jv012616
                            If f1 = "DOCK" Then                                                         'jv012616
                                s = s & f1 & Chr(9) & f2 & Chr(9) & "A10-" & Trim(f3) & Chr(9)          'jv012616
                            Else                                                                        'jv012616
                                s = s & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)                   'jv012616
                            End If                                                                      'jv012616
                            s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                            If r12flag = True Then
                                ocode = Mid(f6, 11, 3)              'jv071715
                                s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                            Else
                                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            End If
                            s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                            grid1.AddItem s
                        End If
                    Loop
                    Close #1
                End If
                sdir = Dir$
                DoEvents
            Loop
        Next i
    End If
    ' --------------------- End Sylacauga ----------------------
    
    ' ------------------ Broken Arrow OP Code ------------------
    If Val(opcode) >= 100 And Val(opcode) <= 199 And Form1.plantno <> 51 Then
        spath = k10path & "ship*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            If fdate >= sdate And fdate <= edate Then
                Open k10path & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f9 = wlot And bc < wsku Then bc = f6
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        's = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = "S" & Chr(9) & f0 & Chr(9)                                              'jv012616
                        If f1 = "DOCK" Then                                                         'jv012616
                            s = s & f1 & Chr(9) & f2 & Chr(9) & "K10-" & Trim(f3) & Chr(9)          'jv012616
                        Else                                                                        'jv012616
                            s = s & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)                   'jv012616
                        End If                                                                      'jv012616
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            ocode = Mid(f6, 11, 3)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                    End If
                
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
        For i = syear To eyear                                                      'jv061215
            spath = k10path & Format(i, "0000") & "\" & "\ship*.txt"                'jv061215
            sdir = Dir$(spath)
            Do While sdir <> ""
                s = Right(sdir, 12)                                                 'jv061215
                s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
                fdate = s                                                           'jv061215
                If fdate >= sdate And fdate <= edate Then
                    Open k10path & Format(i, "0000") & "\" & sdir For Input Shared As #1
                    Do Until EOF(1)
                        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                        If f9 = wlot And bc < wsku Then bc = f6
                        If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                            If f9 = wlot And bc < wsku Then bc = f6
                            's = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                            s = "S" & Chr(9) & f0 & Chr(9)                                              'jv012616
                            If f1 = "DOCK" Then                                                         'jv012616
                                s = s & f1 & Chr(9) & f2 & Chr(9) & "K10-" & Trim(f3) & Chr(9)          'jv012616
                            Else                                                                        'jv012616
                                s = s & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)                   'jv012616
                            End If                                                                      'jv012616
                            s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                            If r12flag = True Then
                                ocode = Mid(f6, 11, 3)              'jv071715
                                s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                            Else
                                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            End If
                            s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                            grid1.AddItem s
                        End If
                    Loop
                    Close #1
                End If
                sdir = Dir$
                DoEvents
            Loop
        Next i
    End If
    ' --------------------- End Broken Arrow ----------------------
    
    ' ------------------ Brenham OP Code ------------------
    If Val(opcode) >= 300 And Val(opcode) <= 599 And Form1.plantno <> 50 Then
        spath = t10path & "ship*.txt"
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            If fdate >= sdate And fdate <= edate Then
                Open t10path & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f9 = wlot And bc < wsku Then bc = f6
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        's = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = "S" & Chr(9) & f0 & Chr(9)                                              'jv012616
                        If f1 = "DOCK" Then                                                         'jv012616
                            s = s & f1 & Chr(9) & f2 & Chr(9) & "T10-" & Trim(f3) & Chr(9)          'jv012616
                        Else                                                                        'jv012616
                            s = s & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)                   'jv012616
                        End If                                                                      'jv012616
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            ocode = Mid(f6, 11, 3)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                        grid1.AddItem s
                    End If
                
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
        For i = syear To eyear                                                      'jv061215
            spath = t10path & Format(i, "0000") & "\" & "\ship*.txt"                'jv061215
            sdir = Dir$(spath)
            Do While sdir <> ""
                s = Right(sdir, 12)                                                 'jv061215
                s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
                fdate = s                                                           'jv061215
                If fdate >= sdate And fdate <= edate Then
                    Open t10path & Format(i, "0000") & "\" & sdir For Input Shared As #1
                    Do Until EOF(1)
                        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                        If f9 = wlot And bc < wsku Then bc = f6
                        If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                            If f9 = wlot And bc < wsku Then bc = f6
                            's = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                            s = "S" & Chr(9) & f0 & Chr(9)                                              'jv012616
                            If f1 = "DOCK" Then                                                         'jv012616
                                s = s & f1 & Chr(9) & f2 & Chr(9) & "T10-" & Trim(f3) & Chr(9)          'jv012616
                            Else                                                                        'jv012616
                                s = s & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)                   'jv012616
                            End If                                                                      'jv012616
                            s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                            If r12flag = True Then
                                ocode = Mid(f6, 11, 3)              'jv071715
                                s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                            Else
                                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            End If
                            s = s & f14 & Chr(9) & f15 & Chr(9) & f16
                            grid1.AddItem s
                        End If
                    Loop
                    Close #1
                End If
                sdir = Dir$
                DoEvents
            Loop
        Next i
    End If
    ' --------------------- End Brenham ----------------------
    

    grid1.Redraw = False
    If grid1.Rows > 1 Then
        For i = 1 To grid1.Rows - 1
            If grid1.TextMatrix(i, 0) = "PR" Then
                grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7) & "0" & grid1.TextMatrix(i, 16)
            Else
                grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7) & grid1.TextMatrix(i, 16) & grid1.TextMatrix(i, 0)
            End If
        Next i
        'If Form1.plantno = "50" Then btran.Visible = True       'jv012616
        btran.Visible = True       'jv033116
    End If
    addbillbc.Visible = False

    
    s = "^Type|^RecId|<Area|<Description|<Source|<Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId"
    grid1.FormatString = s
    grid1.ColWidth(0) = 600
    grid1.ColWidth(1) = 1 '600
    grid1.ColWidth(2) = 1300
    grid1.ColWidth(3) = 1000
    grid1.ColWidth(4) = 1300
    grid1.ColWidth(5) = 1300
    grid1.ColWidth(6) = 3000
    grid1.ColWidth(7) = 1800
    grid1.ColWidth(8) = 600
    grid1.ColWidth(9) = 800
    grid1.ColWidth(10) = 900
    grid1.ColWidth(11) = 800
    grid1.ColWidth(12) = 900
    grid1.ColWidth(13) = 800
    grid1.ColWidth(14) = 1 '800
    grid1.ColWidth(15) = 1600
    grid1.ColWidth(16) = 1400
    grid1.ColWidth(17) = 1 '1000
    grid1.ColWidth(18) = 1 '2100
    sort_ship_history
    'grid1.RowSel = grid1.Row
    'grid1.Col = 18: grid1.ColSel = 18
    'grid1.Sort = 5
    'grid1.FillStyle = flexFillRepeat
    'If grid1.Rows > 2 Then
    '    s = grid1.TextMatrix(1, 7)
    '    For i = 1 To grid1.Rows - 1
    '        If grid1.TextMatrix(i, 7) <> s Then
    '            hrow = Not hrow
    '            s = grid1.TextMatrix(i, 7)
    '        End If
    '        If hrow = True Then
    '            grid1.Row = i: grid1.RowSel = i
    '            grid1.Col = 1: grid1.ColSel = 16
    '            'grid1.CellBackColor = cntlit.BackColor
    '            grid1.CellBackColor = hcolor.BackColor
    '        End If
    '    Next i
    '    grid1.Row = 1
    'End If
    grid1.Redraw = True
            
    'cntlit.Caption = grid1.Rows - 1 & " Records"
    Screen.MousePointer = 0
End Sub

Private Sub sort_ship_history()
    grid1.RowSel = grid1.Row
    grid1.Col = 18: grid1.ColSel = 18
    grid1.Sort = 5
    grid1.FillStyle = flexFillRepeat
    If grid1.Rows > 2 Then
        s = grid1.TextMatrix(1, 7)
        For i = 1 To grid1.Rows - 1
            If grid1.TextMatrix(i, 7) <> s Then
                hrow = Not hrow
                s = grid1.TextMatrix(i, 7)
            End If
            If hrow = True Then
                grid1.Row = i: grid1.RowSel = i
                grid1.Col = 1: grid1.ColSel = 16
                'grid1.CellBackColor = cntlit.BackColor
                grid1.CellBackColor = hcolor.BackColor
            End If
        Next i
        grid1.Row = 1
    End If
    cntlit.Caption = grid1.Rows - 1 & " Records"
End Sub

Private Sub withdrawal()
    Dim ds As ADODB.Recordset, ss As ADODB.Recordset, plit As String
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String, wsku As String, wlot As String
    Dim wzone As String, wstat As String, wgma As Integer, wside As String
    Dim waisle As String, wrack As String, hrow As Boolean, r12flag As Boolean, ocode As String
    Dim cfile As String, s As String, bc As String, srflag As Boolean
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim dl As Long, wbc As String, citem As String
    Dim syear As Integer, eyear As Integer, i As Integer                        'jv061215
    Dim logpath As String
    Dim db5 As ADODB.Connection, ds5 As ADODB.Recordset, ds6 As ADODB.Recordset     'jv080315
    Dim sbc As String, ebc As String                                                'jv080315
    logpath = Form1.logdir
    'srpath = "C:\"
    srpath = logpath                                                            'jv060117
    's = grid1.TextMatrix(grid1.Row, 7)
    'sdate = Format(Val(Mid(s, 9, 2)) - 2, "00")
    'sdate = "20" & sdate & Mid(s, 5, 4)
    
    'sdate = InputBox("Start Date (YearMoDa):", "Start Date...", sdate)
    'If Len(sdate) = 0 Then Exit Sub
    'edate = InputBox("End Date (YearMoDa):", "End Date...", Format(Now, "yyyymmdd"))
    'If Len(edate) = 0 Then Exit Sub
    On Error GoTo vberror
    wbc = grid1.TextMatrix(grid1.Row, 7)
    wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
    wbc = InputBox("Enter a BarCode for the withdrawal:", "BarCode Example....", wbc)
    If Len(wbc) = 0 Then Exit Sub
    wsku = Trim(Left(wbc, 4))
    wlot = barcode_to_lotnum(wbc)
    If wlot = "01001" Then
        MsgBox "Invalid BarCode example.", vbExclamation + vbOKOnly, "problem with barcode..."
        Exit Sub
    End If
    If MsgBox("Display R12 Code Dates?", vbQuestion + vbYesNo, "R12 Lots...") = vbYes Then
        r12flag = True
    Else
        r12flag = False
    End If
    
    s = wbc                                                             'jv012116
    sdate = Format(Val(Mid(s, 9, 2)) - 2, "00")                         'jv012116
    sdate = "20" & sdate & Mid(s, 5, 4)                                 'jv012116
    edate = Format(Now, "yyyymmdd")                                     'jv012116
    
    hcolor.Caption = "Withdrawal"
    Screen.MousePointer = 11
    grid1.Clear: grid1.Cols = 19: grid1.Rows = 1
    
    'Current location
    If skurec(Val(wsku)).sku = wsku Then
        plit = wsku & " " & skurec(Val(wsku)).prodname
    Else
        plit = wsku & " Undefined SKU"
    End If
    If Form1.plantno = "50" Then
        sbc = Mid(wbc, 1, 13) & "001"                       'jv080315
        ebc = Mid(wbc, 1, 13) & "EOR"                       'jv080315
        s = "select * from position where (barcode >= '" & sbc & "' and barcode <= '" & ebc & "') or sku = '" & wsku & "' and lot2 = '" & wlot & Mid(wbc, 11, 3) & "'"         'jv080315
        's = "select * from position where sku = '" & wsku & "' and (lot_num = '" & wlot & "' or lot2 = '" & wlot & "')"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                wzone = "0": wstat = " ": wgma = 0: wside = " "
                s = "select zone_num, rack_side, lane_status, gmasize from lane where id = " & ds!laneno
                Set ss = Wdb.Execute(s)
                If ss.BOF = False Then
                    ss.MoveFirst
                    wzone = ss!zone_num
                    wstat = ss!lane_status
                    wgma = ss!gmasize
                    wside = ss!rack_side
                End If
                ss.Close
                s = "OH" & Chr(9)                               'Type
                s = s & ds!id & Chr(9)                          'Recid
                s = s & "Crane" & Chr(9)                        'Area
                If wstat = "H" Then s = s & "On Hold"
                If wstat = "B" Then s = s & "Blocked"
                s = s & " " & Chr(9)                            'Description
                s = s & "SR-" & ds!whse_num & Chr(9)            'Source
                If ds!whse_num < 4 Then                         'Target
                    s = s & ds!vert_loc & "-" & ds!horz_loc & "-" & ds!rack_side & " " & ds!posn_num & Chr(9)
                Else
                    s = s & wzone & " " & ds!vert_loc & "-" & ds!horz_loc & "-" & wside & Chr(9)
                End If
                s = s & plit & Chr(9)                           'Product
                s = s & bc000(ds!barcode) & Chr(9)                     'Pallet
                s = s & "1" & Chr(9)                            'Qty
                If wgma = 0 Then
                    s = s & "BBC" & Chr(9)                       'Uom
                Else
                    s = s & "GMA" & Chr(9)
                End If
                'ocode = Mid(ds!barcode, 12, 1)
                ocode = Mid(ds!barcode, 11, 3)                              'jv071715
                If r12flag = True Then
                    s = s & Mid(ds!barcode, 5, 9) & Chr(9)                  'jv071615
                Else
                    s = s & ds!lot_num & Chr(9)                 'Lot1
                End If
                s = s & ds!count_qty & Chr(9)                   'Units
                If r12flag = True Then
                    s = s & r12_lot(ds!lot2, ocode) & Chr(9)
                Else
                    s = s & ds!lot2 & Chr(9)                    'Lot2
                End If
                s = s & ds!qty2 & Chr(9)                        'Units
                s = s & "In-Stock" & Chr(9)                     'Status
                s = s & "WMS" & Chr(9)                          'User
                s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
                s = s & " " & Chr(9)                            'Reqid
                grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    
    If Form1.plantno = "52" Then
        Set db5 = CreateObject("ADODB.Connection")
        'db5.Open "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
        db5.Open "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=sywms;PWD=!Sylacauga_WMS1907"
        sbc = Mid(wbc, 1, 13) & "001"
        ebc = Mid(wbc, 1, 13) & "EOR"
        citem = Trim(Left(wbc, 4)) & "-" & Mid(wbc, 11, 3)                              'jv100516
        s = "select * from pallets where barcode >= '" & sbc & "'"
        s = s & " and barcode <= '" & ebc & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                's = "select iContainerDataSysID from tContainerData where sContainerID = '" & ds!plateno & "'"
                'Set ds5 = db5.Execute(s)
                'If ds5.BOF = False Then
                '    ds5.MoveFirst
                    's = "select * from vContainerLocation_1033 Where iContainerDataSysID = " & ds5(0)
                    's = "select * from vContainerLocation_1033 Where [pal id] = '" & ds!plateno & "'"   'jv081415
                    's = "select * from vContainerLocation_1033 Where [pal id] in "      'jv092816
                    s = "select * from vAllInventory_1033 Where LPN in "       'westfalia update
                    s = s & "('" & ds!plateno & "', '" & ds!barcode & "')"              'jv092816
                    s = s & " and item = '" & citem & "'"                               'jv100516
                    Set ds6 = db5.Execute(s)
                    If ds6.BOF = False Then
                        ds6.MoveFirst
                        s = "OH" & Chr(9)
                        s = s & ds6(0) & Chr(9)         'container id
                        s = s & "Crane" & Chr(9)        'area
                        's = s & ds6(17) & Chr(9)        'Hold reason
                        If ds6(16) > 0 Then            'westfalia update
                            s = s & "Locked" & Chr(9)
                        Else
                            s = s & " " & Chr(9)
                        End If
                        s = s & "CS5" & Chr(9)          'Source
                        s = s & ds6!location & Chr(9)   'Target
                        s = s & plit & Chr(9)           'Product
                        s = s & bc000(ds!barcode) & Chr(9)     'Pallet
                        s = s & "1" & Chr(9)            'Qty
                        'If Trim(ds6!Type) = "BBCPallet" Then  'UOM
                        If Trim(ds6![Pallet Type]) = "BBCPallet" Then   'westfalia Update
                            s = s & "BBC" & Chr(9)
                        Else
                            s = s & "GMA" & Chr(9)
                        End If
                        's = s & ds6!Type & Chr(9)
                        ocode = Mid(ds!barcode, 11, 3)                              'jv071715
                        If r12flag = True Then
                            s = s & Mid(ds!barcode, 5, 9) & Chr(9)                  'jv071615
                        Else
                            s = s & ds!lot1 & Chr(9)                 'Lot1
                        End If
                        s = s & ds!qty1 & Chr(9)                   'Units
                        If r12flag = True Then
                            s = s & r12_lot(ds!lot2, ocode) & Chr(9)
                        Else
                            s = s & ds!lot2 & Chr(9)                    'Lot2
                        End If
                        s = s & ds!qty2 & Chr(9)                        'Units
                        s = s & "In-Stock" & Chr(9)                     'Status
                        s = s & "WMS" & Chr(9)                          'User
                        s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
                        's = s & Format(ds6!creation, "yyMMdd hh:mm:ss") & Chr(9)
                        s = s & ds!plateno & Chr(9)                     'Reqid
                        grid1.AddItem s
                    End If
                    ds6.Close
                'End If
                'ds5.Close
                ds.MoveNext
            Loop
        End If
        ds.Close
        db5.Close
    End If
    
    sbc = Mid(wbc, 1, 13) & "001"                       'jv080315
    ebc = Mid(wbc, 1, 13) & "EOR"                       'jv080315
    's = "select * from rackpos where sku = '" & wsku & "' and (lot_num = '" & wlot & "' or lot2 = '" & wlot & "')"
    s = "select * from rackpos where (barcode >= '" & sbc & "' and barcode <= '" & ebc & "') or sku = '" & wsku & "' and lot2 = '" & wlot & Mid(wbc, 11, 3) & "'"         'jv080315
    'MsgBox s
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            waisle = " ": wstat = " ": wrack = " "
            s = "select aisle, rack, hold from racks where id = " & ds!rackno
            Set ss = Wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                waisle = Trim(ss!aisle)
                wrack = Trim(ss!rack)
                If ss!hold = 0 Then
                    wstat = " "
                Else
                    wstat = "On Hold"
                End If
            End If
            ss.Close
            s = "OH" & Chr(9)                               'Type
            s = s & ds!id & Chr(9)                          'Recid
            s = s & "Racks" & Chr(9)                        'Area
            s = s & wstat & Chr(9)                          'Description
            s = s & "SR-4" & Chr(9)                         'Source
            s = s & waisle & "-" & wrack & Chr(9)           'Target
            s = s & plit & Chr(9)                           'Product
            s = s & bc000(ds!barcode) & Chr(9)                     'Pallet
            s = s & "1" & Chr(9)                            'Qty
            If ds!bbc = "Y" Then
                s = s & "BB" & Chr(9)                       'Uom
            Else
                s = s & "GMA" & Chr(9)
            End If
            'ocode = Mid(ds!barcode, 12, 1)
            ocode = Mid(ds!barcode, 11, 3)                                  'jv071715
            If r12flag = True Then
                s = s & Mid(ds!barcode, 5, 9) & Chr(9)                      'jv071715
            Else
                s = s & ds!lot_num & Chr(9)                 'Lot1
            End If
            s = s & ds!count_qty & Chr(9)                   'Units
            If r12flag = True Then
                s = s & r12_lot(ds!lot2, ocode) & Chr(9)
            Else
                s = s & ds!lot2 & Chr(9)                    'Lot2
            End If
            s = s & ds!qty2 & Chr(9)                        'Units
            s = s & "In-Stock" & Chr(9)                     'Status
            s = s & "WMS" & Chr(9)                          'User
            s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
            s = s & " " & Chr(9)                            'Reqid
            grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
        
    'db.Close
    syear = Val(Left(sdate, 4))                                             'jv061215
    eyear = Val(Left(edate, 4))                                             'jv061215
    bc = " "
    spath = logpath & "recv*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    If r12flag = True Then
                        'ocode = Mid(f6, 12, 1)
                        ocode = Mid(f6, 11, 3)                                  'jv071715
                        's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    Else
                        s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                    End If
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                    grid1.AddItem s
                End If
                
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\recv*.txt"                         'jv061215
        'MsgBox spath
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            'ocode = Mid(f6, 12, 1)
                            ocode = Mid(f6, 11, 3)                      'jv071715
                            's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                        grid1.AddItem s
                    End If
                
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i                                                                          'jv0612515
    
    spath = logpath & "tml*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "TM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    If r12flag = True Then
                        'ocode = Mid(f6, 12, 1)
                        ocode = Mid(f6, 11, 3)                  'jv071715
                        's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    Else
                        s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                    End If
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                    grid1.AddItem s
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear
        spath = logpath & Format(i, "0000") & "\tml*.txt"                          'jv061215
        'MsgBox spath
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "TM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            'ocode = Mid(f6, 12, 1)
                            ocode = Mid(f6, 11, 3)                          'jv071715
                            's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)   'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                        grid1.AddItem s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    spath = logpath & "move*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "M" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    If r12flag = True Then
                        'ocode = Mid(f6, 12, 1)
                        ocode = Mid(f6, 11, 3)                              'jv071715
                        's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    Else
                        s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                    End If
                    's = s & f14 & Chr(9) & f15 & Chr(9) & f16
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                    grid1.AddItem s
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear                                                      'jv061215
        spath = logpath & Format(i, "0000") & "\move*.txt"                      'jv061215
        'MsgBox spath
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "M" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            'ocode = Mid(f6, 12, 1)
                            ocode = Mid(f6, 11, 3)                      'jv071715
                            's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                        grid1.AddItem s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    spath = logpath & "sr4rem*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "RR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    If r12flag = True Then
                        'ocode = Mid(f6, 12, 1)
                        ocode = Mid(f6, 11, 3)                  'jv071715
                        's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    Else
                        s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                    End If
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                    grid1.AddItem s
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear                                                      'jv061215
        spath = logpath & Format(i, "0000") & "\sr4rem*.txt"                    'jv061215
        'MsgBox spath
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "RR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            'ocode = Mid(f6, 12, 1)
                            ocode = Mid(f6, 11, 3)                      'jv071715
                            's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                        grid1.AddItem s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    spath = logpath & "ship*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f9 = wlot And bc < wsku Then bc = f6
                'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    If r12flag = True Then
                        'ocode = Mid(f6, 12, 1)
                        ocode = Mid(f6, 11, 3)
                        's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    Else
                        s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                    End If
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                    grid1.AddItem s
                End If
                
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear                                                      'jv061215
        spath = logpath & Format(i, "0000") & "\" & "\ship*.txt"                'jv061215
        'MsgBox spath
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                        If f9 = wlot And bc < wsku Then bc = f6
                    'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            'ocode = Mid(f6, 12, 1)
                            ocode = Mid(f6, 11, 3)              'jv071715
                            's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                        grid1.AddItem s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    spath = logpath & "bill*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f9 = wlot And bc < wsku Then bc = f6
                'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = "B" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    If r12flag = True Then
                        'ocode = Mid(f6, 12, 1)
                        ocode = Mid(f6, 11, 3)
                        's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    Else
                        s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                    End If
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                    grid1.AddItem s
                End If
                
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear                                                      'jv061215
        spath = logpath & Format(i, "0000") & "\" & "\bill*.txt"                'jv061215
        'MsgBox spath
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                        If f9 = wlot And bc < wsku Then bc = f6
                    'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = "B" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            'ocode = Mid(f6, 12, 1)
                            ocode = Mid(f6, 11, 3)              'jv071715
                            's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                        grid1.AddItem s
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    
    spath = logpath & "wms*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                    If f1 = "DOCK" And f13 = "COMP" Then            'jv013114
                    Else
                        s = "WM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            'ocode = Mid(f6, 12, 1)
                            ocode = Mid(f6, 11, 3)              'jv071715
                            's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                        grid1.AddItem s
                    End If
                    If f9 = wlot And bc < wsku Then bc = f6
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear                                                      'jv061215
        spath = logpath & Format(i, "0000") & "\wms*.txt"                       'jv061215
        'MsgBox spath
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f1 = "DOCK" And f13 = "COMP" Then            'jv013114
                        Else
                            s = "WM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                            s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                            If r12flag = True Then
                                'ocode = Mid(f6, 12, 1)
                                ocode = Mid(f6, 11, 3)              'jv071715
                                's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                                s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                            Else
                                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            End If
                            s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                            grid1.AddItem s
                        End If
                        If f9 = wlot And bc < wsku Then bc = f6
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    spath = logpath & "pick*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                    s = "P" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    If r12flag = True Then
                        'ocode = Mid(f6, 12, 1)
                        ocode = Mid(f6, 11, 3)                  'jv071715
                        's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    Else
                        s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                    End If
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                    grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear                                                      'jv061215
        spath = logpath & Format(i, "0000") & "\pick*.txt"                      'jv061215
        'MsgBox spath
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        s = "P" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            'ocode = Mid(f6, 12, 1)
                            ocode = Mid(f6, 11, 3)                  'jv071715
                            's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                        grid1.AddItem s
                        If f9 = wlot And bc < wsku Then bc = f6
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    
    
    srflag = False
    If Form1.plantno = 50 Then
        If MsgBox("Include SR Logs?", vbQuestion + vbYesNo, "SR Logs....") = vbYes Then srflag = True
    End If
    If srflag = True Then
    'spath = logpath & "sr*.txt"                                             'jv060117
    spath = srpath & "sr*.txt"                                             'jv060117
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
        If fdate >= sdate And fdate <= edate Then
            Open srpath & sdir For Input Shared As #1
            'Open "C:\Users\rlhalfmann\Desktop\debugthisfile.txt" For Input Shared As #1
            ' Put error handling here for Input Past End of File
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                    s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                    If r12flag = True Then
                        'ocode = Mid(f6, 12, 1)
                        ocode = Mid(f6, 11, 3)                  'jv071715
                        's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                    Else
                        s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                    End If
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                    grid1.AddItem s
                    If f9 = wlot And bc < wsku Then bc = f6
                End If
            Loop
            Close #1
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear                                                      'jv061215
        'spath = logpath & Format(i, "0000") & "\sr*.txt"                        'jv060117
        spath = srpath & Format(i, "0000") & "\sr*.txt"                        'jv060117
        'MsgBox spath
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open srpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    'If Left(f6, 10) = Left(wbc, 10) Or (Trim(Left(f6, 4)) = wsku And (f9 = wlot Or Mid(f11, 1, 5) = wlot Or wlot = "All")) Then
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                        s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        If r12flag = True Then
                            'ocode = Mid(f6, 12, 1)
                            ocode = Mid(f6, 11, 3)                  'jv071715
                            's = s & Mid(f6, 5, 8) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                            s = s & Mid(f6, 5, 9) & Chr(9) & f10 & Chr(9) & r12_lot(f11, ocode) & Chr(9) & f12 & Chr(9) & f13 & Chr(9)  'jv071715
                        Else
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                        End If
                        s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9)
                        grid1.AddItem s
                        If f9 = wlot And bc < wsku Then bc = f6
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
    End If
    
    grid1.Redraw = False
    If grid1.Rows > 1 Then
        For i = 1 To grid1.Rows - 1
            If grid1.TextMatrix(i, 0) = "PR" Then
                grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7) & "0" & grid1.TextMatrix(i, 16)
            Else
                grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7) & grid1.TextMatrix(i, 16) & grid1.TextMatrix(i, 0)
            End If
        Next i
    End If
    
    If Check1.Value = 1 Then
        s = "^Type|^RecId|<Area|<Description|<Source|<Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId|<FileSource"
        grid1.FormatString = s
        grid1.ColWidth(0) = 600
        grid1.ColWidth(1) = 1 '600
        grid1.ColWidth(2) = 1300
        grid1.ColWidth(3) = 1000
        grid1.ColWidth(4) = 1300
        grid1.ColWidth(5) = 1300
        grid1.ColWidth(6) = 3000
        grid1.ColWidth(7) = 1800
        grid1.ColWidth(8) = 600
        grid1.ColWidth(9) = 800
        grid1.ColWidth(10) = 900
        grid1.ColWidth(11) = 800
        grid1.ColWidth(12) = 900
        grid1.ColWidth(13) = 800
        grid1.ColWidth(14) = 1 '800
        grid1.ColWidth(15) = 1000
        grid1.ColWidth(16) = 1400
        grid1.ColWidth(17) = 1000 '1000
        grid1.ColWidth(18) = 1 '2100
    Else
        s = "^Type|^|^|<Description|<Source|<Target|<Product|^Pallet|^|^|^LotNum|^Units|^LotNum|^Units|^|^|<Time|^|<FileSource"
        grid1.FormatString = s
        grid1.ColWidth(0) = 600
        grid1.ColWidth(1) = 1 '600
        grid1.ColWidth(2) = 1 '300
        grid1.ColWidth(3) = 1000
        grid1.ColWidth(4) = 1300
        grid1.ColWidth(5) = 1300
        grid1.ColWidth(6) = 3000
        grid1.ColWidth(7) = 1800
        grid1.ColWidth(8) = 1 '600
        grid1.ColWidth(9) = 1 '800
        grid1.ColWidth(10) = 900
        grid1.ColWidth(11) = 800
        grid1.ColWidth(12) = 900
        grid1.ColWidth(13) = 800
        grid1.ColWidth(14) = 1 '800
        grid1.ColWidth(15) = 1 '000
        grid1.ColWidth(16) = 1400
        grid1.ColWidth(17) = 1 '1000
        grid1.ColWidth(18) = 1 '2100
    End If
    
    grid1.RowSel = grid1.Row
    grid1.Col = 18: grid1.ColSel = 18
    grid1.Sort = 5
    grid1.FillStyle = flexFillRepeat
    If grid1.Rows > 2 Then
        s = grid1.TextMatrix(1, 7)
        For i = 1 To grid1.Rows - 1
            If grid1.TextMatrix(i, 7) <> s Then
                hrow = Not hrow
                s = grid1.TextMatrix(i, 7)
            End If
            If hrow = True Then
                grid1.Row = i: grid1.RowSel = i
                grid1.Col = 1: grid1.ColSel = 7
                grid1.CellBackColor = cntlit.BackColor
            End If
        Next i
        grid1.Row = 1
    End If
    grid1.Redraw = True
            
    cntlit.Caption = grid1.Rows - 1 & " Records"
    Screen.MousePointer = 0
Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, "palmoves.frm", "withdrawal", WDUserId)
    If MsgBox(edesc & vbCrLf & srpath & sdir, vbRetryCancel + vbQuestion, "Malformed text file") = vbRetry Then
        Resume
    Else
        End
    End If
End Sub


Private Sub print_pgrid()
    Dim rt As String, rf As String, rh As String
    Dim i As Integer, k As Integer, j As Integer, s As String
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = grid1.Cols - 1
    For i = 1 To grid1.Rows - 1
        If Left(grid1.TextMatrix(i, 18), 3) <> "999" Then
            s = grid1.TextMatrix(i, 0)
            For k = 1 To grid1.Cols - 2
                s = s & Chr(9) & grid1.TextMatrix(i, k)
            Next k
            pgrid.AddItem s
        End If
    Next i
    pgrid.FormatString = grid1.FormatString
    For i = 0 To pgrid.Cols - 1
        pgrid.ColWidth(i) = grid1.ColWidth(i)
    Next i
    
    rt = Me.Caption
    rh = hcolor.Caption '& "  " & Text1
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, pgrid, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
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

Private Sub refresh_grid1()
    Dim cfile As String, s As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim logpath As String, srpath As String                                 'jv060117
    Screen.MousePointer = 11
    btran.Visible = False
    addbillbc.Visible = False
    logpath = Form1.logdir
    srpath = Form1.srserv                                                   'jv060117
    'srpath = "c:\"                                                          'jv060117
    srpath = logpath
    grid1.Clear: grid1.Rows = 1: grid1.Cols = 19
    If Combo1 = "Shipping" Then
        addrec.Enabled = True
    Else
        addrec.Enabled = False
    End If
    
    If Combo1 = "Production" Or Combo1 = "All" Then
        cfile = logpath & "recv" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do While Not EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile
                grid1.AddItem s
            Loop
            Close #1
        Else                                                                                            'jv061215
            cfile = logpath & Right(Text1, 4) & "\recv" & Format(Text1, "mmddyyyy") & ".txt"            'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If
    
    If Combo1 = "Shipping" Or Combo1 = "All" Then
        cfile = logpath & "ship" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "S" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile
                grid1.AddItem s
            Loop
            Close #1
        Else                                                                                            'jv061215
            cfile = logpath & Right(Text1, 4) & "\ship" & Format(Text1, "mmddyyyy") & ".txt"            'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If
    
    If Combo1 = "Bills" Or Combo1 = "All" Then
        cfile = logpath & "bill" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "B" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile
                grid1.AddItem s
            Loop
            Close #1
        Else                                                                                            'jv061215
            cfile = logpath & Right(Text1, 4) & "\bill" & Format(Text1, "mmddyyyy") & ".txt"            'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If
    
    
    If Combo1 = "Rack Moves" Or Combo1 = "All" Then
        cfile = logpath & "move" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "M" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile
                grid1.AddItem s
            Loop
            Close #1
        Else                                                                                            'jv061215
            cfile = logpath & Right(Text1, 4) & "\move" & Format(Text1, "mmddyyyy") & ".txt"            'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If
    
    If Combo1 = "Picks" Or Combo1 = "All" Then
        cfile = logpath & "pick" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "P" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile
                grid1.AddItem s
            Loop
            Close #1
        Else                                                                                            'jv061215
            cfile = logpath & Right(Text1, 4) & "\pick" & Format(Text1, "mmddyyyy") & ".txt"            'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If
    
    If Combo1 = "Traffic Master" Or Combo1 = "All" Then
        cfile = logpath & "tml" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "TM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile
                grid1.AddItem s
            Loop
            Close #1
        Else                                                                                            'jv061215
            cfile = logpath & Right(Text1, 4) & "\tml" & Format(Text1, "mmddyyyy") & ".txt"             'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If
    
    If Combo1 = "WMS" Or Combo1 = "All" Or Combo1 = "Rack Activity" Then
        cfile = logpath & "wms" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "WM" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile
                If Left(f6, 3) <> "ING" Then
                    grid1.AddItem s
                End If
            Loop
            Close #1
        Else                                                                                            'jv061215
            cfile = logpath & Right(Text1, 4) & "\wms" & Format(Text1, "mmddyyyy") & ".txt"             'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "PR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If
    
    If Combo1 = "Rack Activity" Or Combo1 = "All" Then
        cfile = logpath & "sr4rem" & Format(Text1, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                s = "RR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & Trim(f2) & Chr(9) & Trim(f3) & Chr(9)
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile
                grid1.AddItem s
            Loop
            Close #1
        Else                                                                                            'jv061215
            cfile = logpath & Right(Text1, 4) & "\sr4rem" & Format(Text1, "mmddyyyy") & ".txt"          'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "RR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If
        
    If Combo1 = "SR Logs" Or Combo1 = "All" Then                                                           'jv060117
        'cfile = srpath & "sr1\SR1" & Format(Text1, "mmddyyyy") & ".txt"                                 'jv060117
        cfile = srpath & "SR" & Format(Text1, "mmddyyyy") & ".txt"                                     'jv060117
        If Len(Dir(cfile)) > 0 Then                                                                     'jv060117
            Open cfile For Input Shared As #1                                                           'jv060117
            Do Until EOF(1)                                                                             'jv060117
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16     'jv060117
                s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & Trim(f2) & Chr(9) & Trim(f3) & Chr(9)  'jv060117
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9) 'jv060117
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)         'jv060117
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                              'jv060117
                grid1.AddItem s                                                                         'jv060117
            Loop                                                                                        'jv060117
            Close #1                                                                                    'jv060117
        Else                                                                                            'jv061215
            cfile = srpath & Right(Text1, 4) & "\sr" & Format(Text1, "mmddyyyy") & ".txt"               'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If                                                                                              'jv060117
        
    If Combo1 = "SR Logs" Or Combo1 = "All" Then                                                           'jv060117
        'cfile = srpath & "sr1\SR1" & Format(Text1, "mmddyyyy") & ".txt"                                 'jv060117
        cfile = srpath & "SR1" & Format(Text1, "mmddyyyy") & ".txt"                                     'jv060117
        If Len(Dir(cfile)) > 0 Then                                                                     'jv060117
            Open cfile For Input Shared As #1                                                           'jv060117
            Do Until EOF(1)                                                                             'jv060117
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16     'jv060117
                s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & Trim(f2) & Chr(9) & Trim(f3) & Chr(9)  'jv060117
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9) 'jv060117
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)         'jv060117
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                              'jv060117
                grid1.AddItem s                                                                         'jv060117
            Loop                                                                                        'jv060117
            Close #1                                                                                    'jv060117
        Else                                                                                            'jv061215
            cfile = srpath & Right(Text1, 4) & "\sr1" & Format(Text1, "mmddyyyy") & ".txt"               'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If                                                                                              'jv060117
        
    If Combo1 = "SR Logs" Or Combo1 = "All" Then                                                           'jv060117
        'cfile = srpath & "sr1\SR1" & Format(Text1, "mmddyyyy") & ".txt"                                 'jv060117
        cfile = srpath & "SR2" & Format(Text1, "mmddyyyy") & ".txt"                                     'jv060117
        If Len(Dir(cfile)) > 0 Then                                                                     'jv060117
            Open cfile For Input Shared As #1                                                           'jv060117
            Do Until EOF(1)                                                                             'jv060117
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16     'jv060117
                s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & Trim(f2) & Chr(9) & Trim(f3) & Chr(9)  'jv060117
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9) 'jv060117
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)         'jv060117
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                              'jv060117
                grid1.AddItem s                                                                         'jv060117
            Loop                                                                                        'jv060117
            Close #1                                                                                    'jv060117
        Else                                                                                            'jv061215
            cfile = srpath & Right(Text1, 4) & "\sr2" & Format(Text1, "mmddyyyy") & ".txt"               'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If                                                                                              'jv060117
        
    If Combo1 = "SR Logs" Or Combo1 = "All" Then                                                           'jv060117
        'cfile = srpath & "sr1\SR1" & Format(Text1, "mmddyyyy") & ".txt"                                 'jv060117
        cfile = srpath & "SR3" & Format(Text1, "mmddyyyy") & ".txt"                                     'jv060117
        If Len(Dir(cfile)) > 0 Then                                                                     'jv060117
            Open cfile For Input Shared As #1                                                           'jv060117
            Do Until EOF(1)                                                                             'jv060117
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16     'jv060117
                s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & Trim(f2) & Chr(9) & Trim(f3) & Chr(9)  'jv060117
                s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9) 'jv060117
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)         'jv060117
                s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                              'jv060117
                grid1.AddItem s                                                                         'jv060117
            Loop                                                                                        'jv060117
            Close #1                                                                                    'jv060117
        Else                                                                                            'jv061215
            cfile = srpath & Right(Text1, 4) & "\sr3" & Format(Text1, "mmddyyyy") & ".txt"               'jv061215
            If Len(Dir(cfile)) > 0 Then                                                                 'jv061215
                Open cfile For Input Shared As #1                                                       'jv061215
                Do Until EOF(1)                                                                         'jv061215
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16 'jv061215
                    s = "SR" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)     'jv061215
                    s = s & Trim(f4) & Chr(9) & f5 & Chr(9) & bc000(f6) & Chr(9) & f7 & Chr(9) & f8 & Chr(9)   'jv061215
                    s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)     'jv061215
                    s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & cfile                          'jv061215
                    grid1.AddItem s                                                                     'jv061215
                Loop                                                                                    'jv061215
                Close #1                                                                                'jv061215
            End If                                                                                      'jv061215
        End If                                                                                          'jv061215
        'MsgBox cfile                                                                                    'jv061215
    End If                                                                                              'jv060117
    
    If Combo1 = "SR Logs" Then Call sortdt_Click
        
    
    If Check1.Value = 1 Then
        s = "^Type|^RecId|<Area|<Description|<Source|<Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId|<FileSource"
        grid1.FormatString = s
        grid1.ColWidth(0) = 600
        grid1.ColWidth(1) = 600
        grid1.ColWidth(2) = 1300
        grid1.ColWidth(3) = 1000
        grid1.ColWidth(4) = 1300
        grid1.ColWidth(5) = 1300
        grid1.ColWidth(6) = 3000
        grid1.ColWidth(7) = 1800
        grid1.ColWidth(8) = 600
        grid1.ColWidth(9) = 800
        grid1.ColWidth(10) = 800
        grid1.ColWidth(11) = 800
        grid1.ColWidth(12) = 800
        grid1.ColWidth(13) = 800
        grid1.ColWidth(14) = 800
        grid1.ColWidth(15) = 1000
        grid1.ColWidth(16) = 1400
        grid1.ColWidth(17) = 1000
        grid1.ColWidth(18) = 1
    Else
        s = "^Type|^RecId|<Area|<Description|<Source|<Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId|<FileSource"
        grid1.FormatString = s
        grid1.ColWidth(0) = 600
        grid1.ColWidth(1) = 1 '600
        grid1.ColWidth(2) = 1 '1300
        grid1.ColWidth(3) = 1 '1000
        grid1.ColWidth(4) = 1300
        grid1.ColWidth(5) = 1300
        grid1.ColWidth(6) = 3000
        grid1.ColWidth(7) = 1800
        grid1.ColWidth(8) = 1 '600
        grid1.ColWidth(9) = 1 '800
        grid1.ColWidth(10) = 800
        grid1.ColWidth(11) = 800
        grid1.ColWidth(12) = 800
        grid1.ColWidth(13) = 800
        grid1.ColWidth(14) = 1 '800
        grid1.ColWidth(15) = 1 '1000
        grid1.ColWidth(16) = 1400
        grid1.ColWidth(17) = 1 '1000
        grid1.ColWidth(18) = 1
    End If
    hcolor.Caption = "All Records"
    cntlit.Caption = grid1.Rows - 1 & " Records"
    Screen.MousePointer = 0
End Sub

Private Sub addbillbc_Click()
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String, i As Integer
    Dim ds As ADODB.Recordset
    Dim cfile As String, s As String, bc As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    
    Dim t0 As String, t1 As String, t2 As String, t3 As String
    Dim t4 As String, t5 As String, t6 As String, t7 As String
    Dim t8 As String, t9 As String, t10 As String, t11 As String
    Dim t12 As String, t13 As String, t14 As String, t15 As String
    
    Dim dl As Long, wbc As String, tbc As String
    Dim logpath As String
    logpath = Form1.logdir
    If Val(grid1.TextMatrix(grid1.Row, 16)) < 1 Then Exit Sub
    wbc = grid1.TextMatrix(grid1.Row, 7)
    wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
    wbc = InputBox("Enter a BarCode to search for:", "BarCode Example....", wbc)
    If Len(wbc) = 0 Then Exit Sub
    wbc = UCase(wbc)
    For i = 1 To grid1.Rows - 1
        tbc = UCase(grid1.TextMatrix(i, 7))
        tbc = Mid(tbc, 1, 10) & Mid(tbc, 13, 3) & Mid(tbc, 18, 3)   'undo bc000
        If wbc = tbc And grid1.TextMatrix(i, 14) <> "CANC" Then
            MsgBox "BarCode is already on this bill.", vbOKOnly + vbInformation, "Duplicate barcode..."
            Exit Sub
        End If
    Next i
    
    Screen.MousePointer = 11
    t10 = "0"
    s = "Select * from pallets where barcode = '" & wbc & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        t5 = ds!sku
        t6 = ds!barcode
        t7 = "1"
        t8 = "Pallet"
        t9 = ds!lot1
        t10 = ds!qty1
        t11 = ds!lot2
        t12 = ds!qty2
        'MsgBox t6 & " found in pallet table.."
        'ds.Close
        's = "select description, uom_type from sku_config where sku = '" & t5 & "'"
        'Set ds = db.Execute(s)
        'If ds.BOF = False Then
        If skurec(Val(t5)).sku = t5 Then
            'ds.MoveFirst
            't5 = t5 & " " & ds!uom_type & " " & ds!description
            t5 = t5 & " " & skurec(Val(t5)).prodname
        End If
    End If
    ds.Close ': db.Close
    
    sdate = Mid(wbc, 5, 2) & "-" & Mid(wbc, 7, 2) & "-20" & Mid(wbc, 9, 2)
    sdate = DateAdd("yyyy", -2, sdate)
    sdate = Format(sdate, "yyyymmdd")
    edate = Format(Now, "yyyymmdd")
    If Val(t10) = 0 Then
        'Look for barcode in movement log
        spath = logpath & "move*.txt"
        sdir = Dir$(spath)
        Do While sdir <> "" And Val(t10) = 0
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            fdate = Mid(sdir, 5, 2) & "-" & Mid(sdir, 7, 2) & "-" & Mid(sdir, 9, 4)
            fdate = Format(fdate, "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then
                        If Val(Right(wbc, 3)) > 0 Or (Right(wbc, 3) = "EOR" And Val(f12) > 0) Then
                            t0 = f0: t1 = f1: t2 = f2: t3 = f3: t4 = f4
                            t5 = f5: t6 = f6: t7 = f7: t8 = f8: t9 = f9
                            t10 = f10: t11 = f11: t12 = f12: t13 = f13: t14 = f14
                            t15 = f15: t16 = f16
                            s = f2 & " " & f4 & " " & f5 & " .. " & sdir
                            'MsgBox s, vbOKOnly + vbInformation, f15 & " received...... " & f6
                        End If
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    
    If Val(t10) = 0 Then
        'Look for barcodes in shipping tasks
        spath = logpath & "ship*.txt"
        sdir = Dir$(spath)
        Do While sdir <> "" And Val(t10) = 0
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            fdate = Mid(sdir, 5, 2) & "-" & Mid(sdir, 7, 2) & "-" & Mid(sdir, 9, 4)
            fdate = Format(fdate, "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then
                        t0 = f0: t1 = f1: t2 = f2: t3 = f3: t4 = f4
                        t5 = f5: t6 = f6: t7 = f7: t8 = f8: t9 = f9
                        t10 = f10: t11 = f11: t12 = f12: t13 = f13: t14 = f14
                        t15 = f15: t16 = f16
                        s = f2 & " " & f4 & " " & f5 & " .. " & sdir
                        'MsgBox s, vbOKOnly + vbInformation, f15 & " shipped...... " & f6
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    
    If Val(t10) = 0 Then
        'Look for barcodes at wrappers
        sdate = Mid(wbc, 5, 2) & "-" & Mid(wbc, 7, 2) & "-20" & Format(Val(Mid(wbc, 9, 2)) - 2, "00")
        edate = Format(DateAdd("d", 5, sdate), "MM-dd-yyyy")
        sdate = Format(sdate, "yyyymmdd")
        edate = Format(edate, "yyyymmdd")
        spath = logpath & "recv*.txt"
        sdir = Dir$(spath)
        Do While sdir <> "" And Val(t10) = 0
            'fdate = Format(FileDateTime(logpath & sdir), "yyyymmdd")
            fdate = Mid(sdir, 5, 2) & "-" & Mid(sdir, 7, 2) & "-" & Mid(sdir, 9, 4)
            fdate = Format(fdate, "yyyymmdd")
            If fdate >= sdate And fdate <= edate Then
                Open logpath & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f6 = wbc Then
                        If Val(Right(wbc, 3)) > 0 Or (Right(wbc, 3) = "EOR" And Val(f12) > 0) Then
                            t0 = f0: t1 = f1: t2 = f2: t3 = f3: t4 = f4
                            t5 = f5: t6 = f6: t7 = f7: t8 = f8: t9 = f9
                            t10 = f10: t11 = f11: t12 = f12: t13 = f13: t14 = f14
                            t15 = f15: t16 = f16
                            s = f2 & " " & f4 & " " & f5 & " .. " & sdir
                            'MsgBox s, vbOKOnly + vbInformation, f15 & " received...... " & f6
                        End If
                    End If
                Loop
                Close #1
            End If
            sdir = Dir$
            DoEvents
        Loop
    End If
    
    Screen.MousePointer = 0
    If Val(t10) <> 0 Then
        i = grid1.Row
        s = "B" & Chr(9)
        s = s & grid1.TextMatrix(i, 1) & Chr(9)
        s = s & grid1.TextMatrix(i, 2) & Chr(9)
        s = s & grid1.TextMatrix(i, 3) & Chr(9)
        s = s & "WMS-Add" & Chr(9) 'billgrid.TextMatrix(i, 4) & Chr(9)
        s = s & grid1.TextMatrix(i, 5) & Chr(9)
        s = s & t5 & Chr(9)
        s = s & t6 & Chr(9)
        s = s & t7 & Chr(9)
        s = s & t8 & Chr(9)
        s = s & t9 & Chr(9)
        s = s & t10 & Chr(9)
        s = s & t11 & Chr(9)
        s = s & t12 & Chr(9)
        s = s & "PEND" & Chr(9) 'billgrid.TextMatrix(i, 14) & Chr(9)
        's = s & "wms" & Chr(9) 'billgrid.TextMatrix(i, 15) & Chr(9)
        s = s & Form1.userid & Chr(9) 'billgrid.TextMatrix(i, 15) & Chr(9)
        s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9) 'billgrid.TextMatrix(i, 16) & Chr(9)
        s = s & grid1.TextMatrix(i, 17) & Chr(9)
        grid1.AddItem s, i
        cntlit.Caption = grid1.Rows - 1 & " Records"
        srun = grid1.TextMatrix(i, 17)
        grid1.Row = i
        cfile = Form1.logdir & "wms" & Format(Text1.Text, "MMddyyyy") & ".txt"
        'cfile = "v:\testlogs\wms" & Format(Text1.Text, "MMddyyyy") & ".txt"
        'MsgBox cfile
        Open cfile For Append As #1
        For k = 1 To 16
            Write #1, grid1.TextMatrix(i, k);
        Next k
        Write #1, grid1.TextMatrix(i, 17)
        Close #1
        
        cfile = Form1.logdir & "bill" & Format(Text1.Text, "MMddyyyy") & ".txt"
        'cfile = "v:\testlogs\bill" & Format(Text1.Text, "MMddyyyy") & ".txt"
        'MsgBox cfile
        Open cfile For Append As #1
        For k = 1 To 16
            Write #1, grid1.TextMatrix(i, k);
        Next k
        Write #1, grid1.TextMatrix(i, 17)
        Close #1
        
        
        grid1.TextMatrix(i, 7) = bc000(t6)
    Else
        MsgBox "BarCode " & wbc & " was not found in the logs.", vbOKOnly + vbInformation, "sorry, cannot add..."
    End If

End Sub

Private Sub addrec_Click()
    Dim mgroup As String, msource As String, mtarget As String, msku As String
    Dim mlot As String, mqty As String, mlot2 As String, mqty2 As String, mbc As String
    Dim i As Integer, s As String, cfile As String
    Dim logpath As String
    logpath = Form1.logdir
    If grid1.Row = 0 Then Exit Sub
    i = grid1.Row
    mgroup = grid1.TextMatrix(i, 3)
    msource = grid1.TextMatrix(i, 4)
    mtarget = grid1.TextMatrix(i, 5)
    msku = grid1.TextMatrix(i, 6)
    mbc = Left(grid1.TextMatrix(i, 7), 12)
    mlot = grid1.TextMatrix(i, 10)
    mqty = grid1.TextMatrix(i, 11)
    mlot2 = grid1.TextMatrix(i, 12)
    mqty2 = grid1.TextMatrix(i, 13)
    mgroup = InputBox("Shipping Group:", "Shipping Group...", mgroup)
    If Len(mgroup) = 0 Then Exit Sub
    msource = InputBox("Source:", "Source...", msource)
    If Len(msource) = 0 Then Exit Sub
    mtarget = InputBox("Target:", "Target...", mtarget)
    If Len(mtarget) = 0 Then Exit Sub
    msku = InputBox("Product:", "Product...", msku)
    If Len(msku) = 0 Then Exit Sub
    mlot = InputBox("Lot Number 1:", "Lot Number 1...", mlot)
    If Len(mlot) = 0 Then Exit Sub
    mqty = InputBox("Units 1:", "Units 1...", mqty)
    If Len(mqty) = 0 Then Exit Sub
    mlot2 = InputBox("Lot Number 2:", "Lot Number 2...", mlot2)
    If Len(mlot2) = 0 Then Exit Sub
    mqty2 = InputBox("Units 2:", "Units 2...", mqty2)
    If Len(mqty2) = 0 Then Exit Sub
    mbc = UCase(InputBox("BarCode:", "BarCode...", mbc))
    If Len(mbc) = 0 Then Exit Sub
    If Len(mbc) < 16 Then
        MsgBox "Invalid BarCode length: " & mbc, vbOKOnly + vbInformation, "Try again..."
        Exit Sub
    End If
    If Left(mbc, 4) <> Left(msku, 4) Then
        MsgBox "BarCode: " & mbc & " and " & msku & " do not match.", vbOKOnly + vbInformation, "Try again..."
        Exit Sub
    End If
    s = "S" & Chr(9)
    s = s & "0" & Chr(9)
    s = s & "DOCK" & Chr(9)
    s = s & mgroup & Chr(9)
    s = s & msource & Chr(9)
    s = s & mtarget & Chr(9)
    s = s & msku & Chr(9)
    s = s & mbc & Chr(9)
    s = s & "1" & Chr(9)
    s = s & "Pallet" & Chr(9)
    s = s & mlot & Chr(9)
    s = s & mqty & Chr(9)
    s = s & mlot2 & Chr(9)
    s = s & mqty2 & Chr(9)
    s = s & "COMP" & Chr(9)
    s = s & "WMS" & Chr(9)
    s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
    grid1.AddItem s, i
    cfile = logpath & "ship" & Format(Text1, "mmddyyyy") & ".txt"
    Open cfile For Append Shared As #1
    Write #1, "0";
    Write #1, "DOCK";
    Write #1, mgroup;
    Write #1, msource;
    Write #1, mtarget;
    Write #1, msku;
    Write #1, mbc;
    Write #1, "1";
    Write #1, "Pallet";
    Write #1, mlot;
    Write #1, mqty;
    Write #1, mlot2;
    Write #1, mqty2;
    Write #1, "COMP";
    Write #1, "WMS";
    Write #1, Format(Now, "yyMMdd hh:mm:ss");
    Write #1, " "
    Close #1
End Sub

Private Sub batonhand_Click()
    Dim s As String
    s = Left(grid1.TextMatrix(grid1.Row, 7), 10)
    s = s & Mid(grid1.TextMatrix(grid1.Row, 7), 13, 3)
    tktonhand.bbarcode = s
    tktonhand.bproduct = Right(grid1.TextMatrix(grid1.Row, 6), Len(grid1.TextMatrix(grid1.Row, 6)) - 4)
    tktonhand.Show
End Sub

Private Sub billhist_Click()
    fetch_bill_of_lading
End Sub

Private Sub btran_Click()
    Dim i As Integer, wbc As String
    wbc = grid1.TextMatrix(grid1.Row, 7)
    wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
    If Len(wbc) = 0 Then Exit Sub
    For i = 1 To grid1.Cols - 2
        branchtrans.tagname(i - 1).Caption = grid1.TextMatrix(0, i)
    Next i
    branchtrans.tagname(15).Caption = "Date"
    i = grid1.Row
    branchtrans.cval(0) = grid1.TextMatrix(i, 1)                                        'Recid
    branchtrans.cval(1) = "BRANCH"                                                      'Area
    branchtrans.cval(2) = "TRANSFER"                                                    'Description
    branchtrans.cval(3) = Left(grid1.TextMatrix(i, 5), Len(grid1.TextMatrix(i, 5)) - 3) 'Source
    branchtrans.cval(4) = grid1.TextMatrix(i, 6)                                        'Product
    branchtrans.cval(5) = wbc 'grid1.TextMatrix(i, 7)                                        'BarCode
    branchtrans.cval(6) = grid1.TextMatrix(i, 8)                                        'Qty
    branchtrans.cval(7) = grid1.TextMatrix(i, 9)                                        'UOM
    branchtrans.cval(8) = grid1.TextMatrix(i, 10)                                       'Lotnum
    branchtrans.cval(9) = grid1.TextMatrix(i, 11)                                       'Units
    branchtrans.cval(10) = grid1.TextMatrix(i, 12)                                      'Lot2
    branchtrans.cval(11) = grid1.TextMatrix(i, 13)                                      'Units
    branchtrans.cval(12) = grid1.TextMatrix(i, 14)                                      'Status
    branchtrans.cval(13) = "WMS" 'grid1.TextMatrix(i, 15)                                      'User
    'branchtrans.cval(14) = grid1.TextMatrix(i, 16)                                      'DateTime
    branchtrans.cval(15) = grid1.TextMatrix(i, 17)                                      'Reqid
    branchtrans.Show
End Sub

Private Sub ccol_Change()
    findcol.Caption = ccol.Caption
End Sub

Private Sub Check1_Click()
    refresh_grid1
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Form1.logdir = "U:\"
    Else
        Form1.logdir = "V:\pallogs\"
    End If
End Sub

Private Sub ckpart_Click()
    Dim ds As ADODB.Recordset, s As String
    Screen.MousePointer = 11
    s = "select * from pallets where barcode <> '...' order by trandate"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If full_pallet(ds!sku, ds!qty1 + ds!qty2) = False Then
                MsgBox ds!barcode & " " & ds!qty1 + ds!qty2 & " " & ds!status & " " & ds!target
                'ds.Edit
                'ds!barcode = "..."
                'ds!status = "Order Pick"
                'ds.Update
                s = "Update pallets set barcode = '...', status = 'Order Pick'"
                s = s & " Where id = " & ds!id
                Wdb.Execute s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
End Sub

Private Sub Combo1_Click()
    refresh_grid1
End Sub

Private Sub emplook_Click()
    Dim ds As ADODB.Recordset, s As String
    If Len(grid1.Text) = 0 Then Exit Sub
    'SQL Database - bbsr
    s = "select * from valuelists where listname = 'wdempid'"
    s = s & " and listreturn = '" & grid1.Text & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = ds!listdisplay
    Else
        s = "Employee #: " & grid1.Text & " is not in WdEmp database."
    End If
    ds.Close
    MsgBox s, vbOKOnly + vbInformation, "WMS SQL Employee " & grid1.Text & " ...."
End Sub

Private Sub findcol_Click()
    Dim i As Integer, s As String, t As String, k As Integer, sc As Integer
    sc = grid1.Col
    k = 0
    s = grid1.Text
    s = InputBox(ccol & ": ", "Highlight " & ccol & "...", s)
    If Len(s) = 0 Then Exit Sub
    hcolor.Caption = ccol & ": " & s
    grid1.Redraw = False
    For i = 1 To grid1.Rows - 1
        grid1.TextMatrix(i, 18) = "99999999999"
        grid1.Row = i: grid1.RowSel = i
        grid1.Col = 1: grid1.ColSel = grid1.Cols - 1
        If UCase(grid1.TextMatrix(i, sc)) = UCase(s) Then
            grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7)
            grid1.CellBackColor = hcolor.BackColor
            k = k + 1
        Else
            grid1.CellBackColor = grid1.BackColor
        End If
    Next i
    grid1.Redraw = True
    grid1.TopRow = 1
    grid1.Row = 1: grid1.RowSel = 1
    grid1.Col = 18: grid1.ColSel = 18
    grid1.Sort = 5
    cntlit.Caption = k & " Records"
    grid1.Col = sc
End Sub

Private Sub findsku_Click()
    Dim i As Integer, s As String, t As String, k As Integer, n As Integer
    k = 0
    's = Left(Grid1.TextMatrix(Grid1.Row, 7), 3)
    s = Trim(Left(grid1.TextMatrix(grid1.Row, 7), 4))       'jv062916
    s = InputBox("SKU:", "Highlight SKU..", s)
    If Len(s) = 0 Then Exit Sub
    n = Len(s)                                              'jv062916
    hcolor.Caption = "SKU: " & s
    grid1.Redraw = False
    For i = 1 To grid1.Rows - 1
        grid1.TextMatrix(i, 18) = "99999999999"
        grid1.Row = i: grid1.RowSel = i
        grid1.Col = 1: grid1.ColSel = grid1.Cols - 1
        'If Left(Grid1.TextMatrix(i, 7), 3) = s Or Left(Grid1.TextMatrix(i, 6), 3) = s Then
        If Left(grid1.TextMatrix(i, 7), n) = s Or Left(grid1.TextMatrix(i, 6), n) = s Then          'jv062916
            grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7)
            grid1.CellBackColor = hcolor.BackColor
            k = k + 1
        Else
            grid1.CellBackColor = grid1.BackColor
        End If
        'If Left(Grid1.TextMatrix(i, 7), 3) <> Left(Grid1.TextMatrix(i, 6), 3) And Grid1.TextMatrix(i, 7) > "100" Then
        If Left(grid1.TextMatrix(i, 7), n) <> Left(grid1.TextMatrix(i, 6), n) And grid1.TextMatrix(i, 7) > "100" Then       'jv062916
            grid1.Row = i: grid1.RowSel = i
            grid1.Col = 6: grid1.ColSel = 7
            grid1.CellBackColor = cntlit.BackColor
            grid1.TextMatrix(i, 18) = grid1.TextMatrix(i, 7)
        End If
    Next i
    grid1.Redraw = True
    grid1.TopRow = 1
    grid1.Row = 1: grid1.RowSel = 1
    grid1.Col = 18: grid1.ColSel = 18
    grid1.Sort = 5
    cntlit.Caption = k & " Records"
End Sub

Private Sub Form_Load()
    Text1 = Format(Now, "mm-dd-yyyy")
    Combo1.Clear
    Combo1.AddItem "Production"
    Combo1.AddItem "Shipping"
    Combo1.AddItem "Rack Moves"
    Combo1.AddItem "Picks"
    Combo1.AddItem "Traffic Master"
    'If Form1.plantno = 50 Then                      'jv060117
    '    Combo1.AddItem "SR Logs"                    'jv060117
    'End If                                          'jv060117
    Combo1.AddItem "WMS"
    Combo1.AddItem "Bills"
    Combo1.AddItem "All"
    Combo1.ListIndex = 0
    'If Form1.plantno = 50 Then
        emplook.Enabled = True
    'Else
    '    emplook.Enabled = False
    'End If
    
    If Form1.plantno = 50 Then                      'jv071717
        pcorr.Visible = True                        'jv071717
    Else                                            'jv071717
        pcorr.Visible = False                       'jv071717
    End If                                          'jv071717
End Sub

Private Sub Form_Resize()
    grid1.Width = Me.Width - 80
    pgrid.Width = Me.Width - 80
    If Me.Height > 2000 Then grid1.Height = Me.Height - 1500
End Sub

Private Sub Grid1_Click()
    ccol = grid1.TextMatrix(0, grid1.Col)
End Sub

Private Sub Grid1_DblClick()
    Dim psku As String, q As Integer, s As String, i As Integer
    i = grid1.Row
    psku = Trim(Left(grid1.TextMatrix(i, 7), 4))
    q = Val(grid1.TextMatrix(i, 11)) + Val(grid1.TextMatrix(i, 13))
    If full_pallet(psku, q) = True Then
        s = " Full pallet"
    Else
        s = " Partial"
    End If
    MsgBox psku & " = " & q & s
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If grid1.TextMatrix(0, grid1.Col) = "User" And emplook.Enabled = True Then
            PopupMenu usermenu
        Else
            PopupMenu findmenu
        End If
    End If
End Sub

Private Sub palhist_Click()
    Call pallet_history
End Sub

Private Sub pbartest_Click()
    Dim ds As ADODB.Recordset, s As String
    Dim rs As ADODB.Recordset, dname As String
    Screen.MousePointer = 11
    s = "select id,area,source,target,product,palletid,description from paltasks"
    s = s & " where palletid > '0' and status = 'PEND'"
    s = s & " and area <> 'TRAFFIC MASTER'"
    s = s & " order by area, source, target"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Printer.FontName = "Arial"
            s = ds!id & " " & ds!area & "-" & ds!source & "-" & ds!target
            Printer.Print s
            Printer.Print ds!product
            Printer.FontName = "IDAutomationHC39M"
            s = "!" & Mid(ds!palletid, 1, 3) & "="
            s = s & Mid(ds!palletid, 5, 6) & "="
            If Mid(ds!palletid, 12, 1) = "_" Then
                s = s & "B="
            Else
                s = s & Mid(ds!palletid, 12, 1) & "="
            End If
            s = s & Mid(ds!palletid, 14, 3) & "!"
            Printer.Print s
            Printer.FontName = "Arial"
            dname = ds!target
            If ds!description > "  " Then
                s = "select target from paltasks where area = 'GROUP'"
                s = s & " and product = '" & ds!description & "   " & ds!target & "'"
                Set rs = Wdb.Execute(s)
                If rs.BOF = False Then
                    rs.MoveFirst
                    dname = rs!target
                Else
                    dname = "NODOOR"
                End If
                rs.Close
            End If
            Printer.Print dname
            Printer.FontName = "IDAutomationHC39M"
            If dname = "ANTE ROOM" Then dname = "ANTE=ROOM"
            If dname = "ORDER PICK" Then dname = "ORDER=PICK"
            If dname = "CRANE 3" Then dname = "CRANE=3"
            If Mid(dname, 2, 1) = "-" Then dname = Trim(dname)
            If Mid(dname, 2, 1) = " " Then dname = Left(dname, 1) & "=" & Trim(Right(dname, Len(dname) - 2))
            dname = Trim(dname)
            s = "!" & dname & "!"
            Printer.Print s
            Printer.FontName = "Arial"
            Printer.Print " "
            Printer.Print " "
            ds.MoveNext
        Loop
        Printer.EndDoc
    End If
    ds.Close
    Screen.MousePointer = 0
End Sub

Private Sub pcorr_Click()                                       'jv071717
    Dim wbc As String
    palcorr.pref = Combo1
    wbc = grid1.TextMatrix(grid1.Row, 7)
    wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
    palcorr.bckey = wbc
    palcorr.Show
End Sub

Private Sub postalllogs_Click()
    Dim db As Database, ds As Recordset, s As String, q As Integer
    Dim pid As Long, i As Long, pstat As String, psku As String
    Screen.MousePointer = 11
    If Combo1 = "Production" Then
        pstat = "Wrapper"
    Else
        If Combo1 = "Shipping" Then
            pstat = "Shipped"
        Else
            pstat = "Warehouse"
        End If
    End If
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, Form1.BBSR)
    s = "select sequence_id from sequences where seq = 'Pallets'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        pid = ds!sequence_id
    End If
    ds.Close
    For i = 1 To grid1.Rows - 1
        cntlit.Caption = Val(cntlit.Caption) - 1
        DoEvents
        psku = Trim(Left(grid1.TextMatrix(i, 7), 4))
        q = Val(grid1.TextMatrix(i, 11)) + Val(grid1.TextMatrix(i, 13))
        If full_pallet(psku, q) = True Then     'Check for full pallets
            s = "select * from pallets where barcode = '" & grid1.TextMatrix(i, 7) & "'"
            Set ds = db.OpenRecordset(s)
            If ds.BOF = False Then
                ds.MoveFirst
                Do Until ds.EOF
                    ds.Edit
                    If pstat = "Wrapper" And Val(grid1.TextMatrix(i, 11)) > 0 Then
                        If ds!status = "Wrapper" Then ds!plateno = grid1.TextMatrix(i, 17)
                        ds!barcode = grid1.TextMatrix(i, 7)
                        ds!qty1 = Val(grid1.TextMatrix(i, 11))
                        ds!lot1 = grid1.TextMatrix(i, 10)
                        ds!qty2 = Val(grid1.TextMatrix(i, 13))
                        ds!lot2 = grid1.TextMatrix(i, 12)
                        ds!sku = psku
                    End If
                    ds!source = grid1.TextMatrix(i, 4)
                    ds!target = grid1.TextMatrix(i, 5)
                    ds!bbc = "Y"
                    If grid1.TextMatrix(i, 5) = "ORDER PICK" Then
                        ds!status = "Order Pick"
                    Else
                        ds!status = pstat
                    End If
                    ds!trandate = grid1.TextMatrix(i, 16)
                    ds.Update
                    ds.MoveNext
                Loop
            Else
                If pstat = "Wrapper" And Val(grid1.TextMatrix(i, 11)) > 0 Then
                    ds.Close
                    s = "select * from pallets where status in ('Shipped','Order Pick')"
                    s = s & " order by trandate"
                    Set ds = db.OpenRecordset(s)
                    If ds.BOF = False Then
                        ds.MoveFirst
                        ds.Edit
                        'ds!plateno = Grid1.TextMatrix(i, 17)
                        ds!plateno = grid1.TextMatrix(i, 7)
                        ds!barcode = grid1.TextMatrix(i, 7)
                        ds!qty1 = Val(grid1.TextMatrix(i, 11))
                        ds!lot1 = grid1.TextMatrix(i, 10)
                        ds!qty2 = Val(grid1.TextMatrix(i, 13))
                        ds!lot2 = grid1.TextMatrix(i, 12)
                        ds!source = grid1.TextMatrix(i, 4)
                        ds!target = grid1.TextMatrix(i, 5)
                        ds!bbc = "Y"
                        If grid1.TextMatrix(i, 5) = "ORDER PICK" Then
                            ds!status = "Order Pick"
                        Else
                            ds!status = pstat
                        End If
                        ds!trandate = grid1.TextMatrix(i, 16)
                        ds!sku = psku
                        ds.Update
                    Else
                        pid = pid + 1
                        s = "Insert Into pallets Values (" & pid
                        's = s & ",'" & Grid1.TextMatrix(i, 17) & "'"
                        s = s & ",'" & grid1.TextMatrix(i, 7) & "'"
                        s = s & ",'" & grid1.TextMatrix(i, 7) & "'"
                        s = s & "," & Val(grid1.TextMatrix(i, 11))
                        s = s & ",'" & grid1.TextMatrix(i, 10) & "'"
                        s = s & "," & Val(grid1.TextMatrix(i, 13))
                        s = s & ",'" & grid1.TextMatrix(i, 12) & "'"
                        s = s & ",'" & grid1.TextMatrix(i, 4) & "'"
                        s = s & ",'" & grid1.TextMatrix(i, 5) & "'"
                        s = s & ",'Y'"
                        If grid1.TextMatrix(i, 5) = "ORDER PICK" Then
                            s = s & ",'Order Pick'"
                        Else
                            s = s & ",'" & pstat & "'"
                        End If
                        s = s & ",'" & grid1.TextMatrix(i, 16) & "'"
                        s = s & ",'" & psku & "')"
                        db.Execute s
                    End If
                    'ds.Close
                End If
            End If
            ds.Close
        End If
    Next i
    s = "Update sequences Set sequence_id = " & pid & " Where seq = 'Pallets'"
    db.Execute s
    db.Close
    Screen.MousePointer = 0

End Sub

Private Sub posteorlogs_Click()
    Dim db As Database, ds As Recordset, s As String, q As Integer
    Dim pid As Long, i As Long, pstat As String, psku As String, rcnt As Integer
    Screen.MousePointer = 11
    If Combo1 = "Production" Then
        pstat = "Wrapper"
    Else
        If Combo1 = "Shipping" Then
            pstat = "Shipped"
        Else
            pstat = "Warehouse"
        End If
    End If
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, Form1.BBSR)
    s = "select sequence_id from sequences where seq = 'Pallets'"
    Set ds = db.OpenRecordset(s)
    If ds.BOF = False Then
        ds.MoveFirst
        pid = ds!sequence_id
    End If
    ds.Close
    rcnt = 0
    For i = 1 To grid1.Rows - 1
        If Val(grid1.TextMatrix(i, 13)) > 0 Or Right(grid1.TextMatrix(i, 7), 3) = "EOR" Then     'Units2
            psku = Trim(Left(grid1.TextMatrix(i, 7), 4))
            q = Val(grid1.TextMatrix(i, 11)) + Val(grid1.TextMatrix(i, 13))
            If full_pallet(psku, q) = True Then     'Check for full pallets
                rcnt = rcnt + 1
                psku = Trim(Left(grid1.TextMatrix(i, 7), 4))
                s = "select * from pallets where barcode = '" & grid1.TextMatrix(i, 7) & "'"
                Set ds = db.OpenRecordset(s)
                If ds.BOF = False Then
                    ds.MoveFirst
                    Do Until ds.EOF
                        ds.Edit
                        If pstat = "Wrapper" And Val(grid1.TextMatrix(i, 11)) > 0 Then
                            If ds!status = "Wrapper" Then ds!plateno = grid1.TextMatrix(i, 17)
                            'ds!plateno = Grid1.TextMatrix(i, 7)
                            ds!barcode = grid1.TextMatrix(i, 7)
                            ds!qty1 = Val(grid1.TextMatrix(i, 11))
                            ds!lot1 = grid1.TextMatrix(i, 10)
                            ds!qty2 = Val(grid1.TextMatrix(i, 13))
                            ds!lot2 = grid1.TextMatrix(i, 12)
                            ds!sku = psku
                        End If
                        ds!source = grid1.TextMatrix(i, 4)
                        ds!target = grid1.TextMatrix(i, 5)
                        ds!bbc = "Y"
                        If grid1.TextMatrix(i, 5) = "ORDER PICK" Then
                            ds!status = "Order Pick"
                        Else
                            ds!status = pstat
                        End If
                        ds!trandate = grid1.TextMatrix(i, 16)
                        ds.Update
                        ds.MoveNext
                    Loop
                Else
                    If pstat = "Wrapper" And Val(grid1.TextMatrix(i, 11)) > 0 Then
                        ds.Close
                        s = "select * from pallets where status in ('Shipped','Order Pick')"
                        s = s & " order by trandate"
                        Set ds = db.OpenRecordset(s)
                        If ds.BOF = False Then
                            ds.MoveFirst
                            ds.Edit
                            ds!plateno = grid1.TextMatrix(i, 17)
                            'ds!plateno = Grid1.TextMatrix(i, 7)
                            ds!barcode = grid1.TextMatrix(i, 7)
                            ds!qty1 = Val(grid1.TextMatrix(i, 11))
                            ds!lot1 = grid1.TextMatrix(i, 10)
                            ds!qty2 = Val(grid1.TextMatrix(i, 13))
                            ds!lot2 = grid1.TextMatrix(i, 12)
                            ds!source = grid1.TextMatrix(i, 4)
                            ds!target = grid1.TextMatrix(i, 5)
                            ds!bbc = "Y"
                            If grid1.TextMatrix(i, 5) = "ORDER PICK" Then
                                ds!status = "Order Pick"
                            Else
                                ds!status = pstat
                            End If
                            ds!trandate = grid1.TextMatrix(i, 16)
                            ds!sku = psku
                            ds.Update
                        Else
                            pid = pid + 1
                            s = "Insert Into pallets Values (" & pid
                            s = s & ",'" & grid1.TextMatrix(i, 17) & "'"
                            's = s & ",'" & Grid1.TextMatrix(i, 7) & "'"
                            s = s & ",'" & grid1.TextMatrix(i, 7) & "'"
                            s = s & "," & Val(grid1.TextMatrix(i, 11))
                            s = s & ",'" & grid1.TextMatrix(i, 10) & "'"
                            s = s & "," & Val(grid1.TextMatrix(i, 13))
                            s = s & ",'" & grid1.TextMatrix(i, 12) & "'"
                            s = s & ",'" & grid1.TextMatrix(i, 4) & "'"
                            s = s & ",'" & grid1.TextMatrix(i, 5) & "'"
                            s = s & ",'Y'"
                            If grid1.TextMatrix(i, 5) = "ORDER PICK" Then
                                s = s & ",'Order Pick'"
                            Else
                                s = s & ",'" & pstat & "'"
                            End If
                            s = s & ",'" & grid1.TextMatrix(i, 16) & "'"
                            s = s & ",'" & psku & "')"
                            db.Execute s
                        End If
                        'ds.Close
                    End If
                End If
                ds.Close
            End If
        End If
    Next i
    s = "Update sequences Set sequence_id = " & pid & " Where seq = 'Pallets'"
    db.Execute s
    db.Close
    MsgBox "Posted " & rcnt & " records.", vbOKOnly + vbInformation, "Pallet records..."
    Screen.MousePointer = 0
End Sub

Private Sub ppsum_Click()
    Dim i As Integer, k As Integer, s As String, aflag As Boolean
    Dim rt As String, rf As String, rh As String
    Dim ds As ADODB.Recordset
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) = "Production" Then Combo1.ListIndex = i
    Next i
    DoEvents
    If grid1.Rows < 2 Then Exit Sub
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 6
    s = " " & Chr(9) & grid1.TextMatrix(1, 6) & Chr(9) & grid1.TextMatrix(1, 10)
    pgrid.AddItem s
    s = " " & Chr(9) & "tot" & grid1.TextMatrix(1, 6) & Chr(9) & grid1.TextMatrix(1, 10)
    pgrid.AddItem s
    
    
    For i = 1 To grid1.Rows - 1
        aflag = True
        For k = 1 To pgrid.Rows - 1
            If Trim(Left(pgrid.TextMatrix(k, 1), 4)) = Trim(Left(grid1.TextMatrix(i, 6), 4)) And pgrid.TextMatrix(k, 2) = grid1.TextMatrix(i, 10) Then
                pgrid.TextMatrix(k, 5) = Val(pgrid.TextMatrix(k, 5)) + Val(grid1.TextMatrix(i, 11))
                aflag = False
                Exit For
            End If
        Next k
        If aflag = True Then
            s = " " & Chr(9) & grid1.TextMatrix(i, 6) & Chr(9) & grid1.TextMatrix(i, 10) & Chr(9) & Chr(9) & Chr(9) & grid1.TextMatrix(i, 11)
            pgrid.AddItem s
        End If
        
        aflag = True
        For k = 1 To pgrid.Rows - 1
            If Trim(Left(pgrid.TextMatrix(k, 1), 7)) = "tot" & Trim(Left(grid1.TextMatrix(i, 6), 4)) Then
                pgrid.TextMatrix(k, 5) = Val(pgrid.TextMatrix(k, 5)) + Val(grid1.TextMatrix(i, 11))
                aflag = False
                Exit For
            End If
        Next k
        If aflag = True Then
            s = " " & Chr(9) & "tot" & grid1.TextMatrix(i, 6) & Chr(9) & " " & Chr(9) & Chr(9) & Chr(9) & grid1.TextMatrix(i, 11)
            pgrid.AddItem s
        End If
        
        
    Next i
    
    For i = 1 To grid1.Rows - 1
        If grid1.TextMatrix(i, 12) > "000" Then
            aflag = True
            For k = 1 To pgrid.Rows - 1
                If pgrid.TextMatrix(k, 1) = grid1.TextMatrix(i, 6) And pgrid.TextMatrix(k, 2) = grid1.TextMatrix(i, 12) Then
                    pgrid.TextMatrix(k, 5) = Val(pgrid.TextMatrix(k, 5)) + Val(grid1.TextMatrix(i, 13))
                    aflag = False
                    Exit For
                End If
            Next k
            If aflag = True Then
                s = " " & Chr(9) & grid1.TextMatrix(i, 6) & Chr(9) & grid1.TextMatrix(i, 12) & Chr(9) & Chr(9) & Chr(9) & grid1.TextMatrix(i, 13)
                pgrid.AddItem s
            End If
        End If
        
        If grid1.TextMatrix(i, 12) > "000" Then
            aflag = True
            For k = 1 To pgrid.Rows - 1
                If pgrid.TextMatrix(k, 1) = "tot" & grid1.TextMatrix(i, 6) Then
                    pgrid.TextMatrix(k, 5) = Val(pgrid.TextMatrix(k, 5)) + Val(grid1.TextMatrix(i, 13))
                    aflag = False
                    Exit For
                End If
            Next k
            If aflag = True Then
                s = " " & Chr(9) & grid1.TextMatrix(i, 6) & Chr(9) & " " & Chr(9) & Chr(9) & Chr(9) & grid1.TextMatrix(i, 13)
                pgrid.AddItem s
            End If
        End If
        
        
    Next i
    If pgrid.Rows > 1 Then
        For i = 1 To pgrid.Rows - 1
            If Left(pgrid.TextMatrix(i, 1), 3) = "tot" Then
                s = pgrid.TextMatrix(i, 1)
                s = Right(s, Len(s) - 3) & "tot"
                pgrid.TextMatrix(i, 1) = s
            End If
            s = Val(Trim(Left(pgrid.TextMatrix(i, 1), 4)))
            If skurec(Val(s)).sku = s Then
                If Val(pgrid.TextMatrix(i, 5)) < skurec(Val(s)).uom_per_pallet Then
                    pgrid.TextMatrix(i, 3) = "0"
                Else
                    pgrid.TextMatrix(i, 3) = Int(Val(pgrid.TextMatrix(i, 5)) / skurec(Val(s)).uom_per_pallet)
                End If
                k = Val(pgrid.TextMatrix(i, 5)) - (Val(pgrid.TextMatrix(i, 3)) * skurec(Val(s)).uom_per_pallet)
                pgrid.TextMatrix(i, 4) = k / (skurec(Val(s)).uom_per_pallet / skurec(Val(s)).qty_per_pallet)
            End If
        Next i
    End If
    pgrid.RowSel = pgrid.Row
    pgrid.Col = 1: pgrid.ColSel = 2
    pgrid.Sort = 5
    
    If pgrid.Rows > 1 Then
        For i = 1 To pgrid.Rows - 1
            If Right(pgrid.TextMatrix(i, 1), 3) = "tot" Then
                pgrid.TextMatrix(i, 1) = " "
            End If
        Next i
    End If
    
    s = "^|<Product|^Lot Number|^Pallets|^Wraps|^Units"
    pgrid.FormatString = s
    pgrid.ColWidth(0) = 1 '2800
    pgrid.ColWidth(1) = 4000
    pgrid.ColWidth(2) = 1200
    pgrid.ColWidth(3) = 1200
    pgrid.ColWidth(4) = 1200
    pgrid.ColWidth(5) = 1200
    rt = ppsum.Caption
    rh = Text1 & "  " & "Production"
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, pgrid, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
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

Private Sub pptot_Click()
    Dim i As Integer, k As Integer, s As String, aflag As Boolean
    Dim rt As String, rf As String, rh As String
    Dim ds As ADODB.Recordset
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) = "Production" Then Combo1.ListIndex = i
    Next i
    DoEvents
    If grid1.Rows < 2 Then Exit Sub
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 6
    s = grid1.TextMatrix(1, 4) & Chr(9) & grid1.TextMatrix(1, 6) & Chr(9) & grid1.TextMatrix(1, 10)
    pgrid.AddItem s
    For i = 1 To grid1.Rows - 1
        aflag = True
        For k = 1 To pgrid.Rows - 1
            If pgrid.TextMatrix(k, 0) = grid1.TextMatrix(i, 4) And Left(pgrid.TextMatrix(k, 1), 3) = Left(grid1.TextMatrix(i, 6), 3) And pgrid.TextMatrix(k, 2) = grid1.TextMatrix(i, 10) Then
                pgrid.TextMatrix(k, 5) = Val(pgrid.TextMatrix(k, 5)) + Val(grid1.TextMatrix(i, 11))
                aflag = False
                Exit For
            End If
        Next k
        If aflag = True Then
            s = grid1.TextMatrix(i, 4) & Chr(9) & grid1.TextMatrix(i, 6) & Chr(9) & grid1.TextMatrix(i, 10) & Chr(9) & Chr(9) & Chr(9) & grid1.TextMatrix(i, 11)
            pgrid.AddItem s
        End If
    Next i
    
    For i = 1 To grid1.Rows - 1
        If grid1.TextMatrix(i, 12) > "000" Then
            aflag = True
            For k = 1 To pgrid.Rows - 1
                If pgrid.TextMatrix(k, 0) = grid1.TextMatrix(i, 4) And pgrid.TextMatrix(k, 1) = grid1.TextMatrix(i, 6) And pgrid.TextMatrix(k, 2) = grid1.TextMatrix(i, 12) Then
                    pgrid.TextMatrix(k, 5) = Val(pgrid.TextMatrix(k, 5)) + Val(grid1.TextMatrix(i, 13))
                    aflag = False
                    Exit For
                End If
            Next k
            If aflag = True Then
                s = grid1.TextMatrix(i, 4) & Chr(9) & grid1.TextMatrix(i, 6) & Chr(9) & grid1.TextMatrix(i, 12) & Chr(9) & Chr(9) & Chr(9) & grid1.TextMatrix(i, 13)
                pgrid.AddItem s
            End If
        End If
    Next i
    If pgrid.Rows > 1 Then
        For i = 1 To pgrid.Rows - 1
            s = Trim(Left(pgrid.TextMatrix(i, 1), 4))
            If skurec(Val(s)).sku = s Then
                If Val(pgrid.TextMatrix(i, 5)) < skurec(Val(s)).uom_per_pallet Then
                    pgrid.TextMatrix(i, 3) = "0"
                Else
                    pgrid.TextMatrix(i, 3) = Int(Val(pgrid.TextMatrix(i, 5)) / skurec(Val(s)).uom_per_pallet)
                End If
                k = Val(pgrid.TextMatrix(i, 5)) - (Val(pgrid.TextMatrix(i, 3)) * skurec(Val(s)).uom_per_pallet)
                pgrid.TextMatrix(i, 4) = k / (skurec(Val(s)).uom_per_pallet / skurec(Val(s)).qty_per_pallet)
            End If
        Next i
    End If
    pgrid.RowSel = pgrid.Row
    pgrid.Col = 1: pgrid.ColSel = 2
    pgrid.Sort = 5
    s = "^Source|<Product|^Lot Number|^Pallets|^Wraps|^Units"
    pgrid.FormatString = s
    pgrid.ColWidth(0) = 2800
    pgrid.ColWidth(1) = 4000
    pgrid.ColWidth(2) = 1200
    pgrid.ColWidth(3) = 1200
    pgrid.ColWidth(4) = 1200
    pgrid.ColWidth(5) = 1200
    rt = pptot.Caption
    rh = Text1 & "  " & "Production"
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, pgrid, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
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

Private Sub prtcur_Click()
    Dim rt As String, rf As String, rh As String
    If hcolor.Caption <> "All Records" Then
        Call print_pgrid
        Exit Sub
    End If
    rt = Me.Caption
    rh = Combo1 & "  " & hcolor.Caption
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
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

Private Sub psrlogship_Click()
    Dim cfile As String, p As ptask, i As Integer, s As String, wbc As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim psku As String, plot As String, pplt As String, pgrp As String
    If Combo1 <> "Shipping" Then Exit Sub
    Screen.MousePointer = 11
    For i = 1 To grid1.Rows - 1
        If grid1.TextMatrix(i, 4) = "SR1" Then
            p.id = grid1.TextMatrix(i, 1)
            p.area = "SR-1"
            p.description = grid1.TextMatrix(i, 3)
            p.source = " "
            p.target = " "
            p.product = grid1.TextMatrix(i, 6)
            wbc = grid1.TextMatrix(i, 7)
            wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
            p.palletid = wbc
            p.qty = grid1.TextMatrix(i, 8)
            p.uom = grid1.TextMatrix(i, 9)
            p.lotnum = grid1.TextMatrix(i, 10)
            p.units = grid1.TextMatrix(i, 11)
            p.lotnum2 = grid1.TextMatrix(i, 12)
            p.units2 = grid1.TextMatrix(i, 13)
            p.status = " "
            p.userid = "SR-1"
            p.trandate = Left(grid1.TextMatrix(i, 16), 7)
            p.reqid = grid1.TextMatrix(i, 17)
            s = "V:\sr1\bin\SR1" & Mid(Text1, 1, 2) & Mid(Text1, 4, 2) & ".csv"
            'MsgBox s
            If Len(Dir(s)) > 0 Then
                psku = Trim(Left(p.palletid, 4))
                plot = p.lotnum
                pplt = Right(p.palletid, 3)
                pgrp = p.description
                t = p.palletid & " SKU=" & psku & " Lot=" & plot & " Pallet=" & pplt
                'MsgBox t
                Open s For Input As #5
                Do Until EOF(5)
                    Input #5, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    If f2 = psku And f3 = plot And f4 = pplt Then
                        p.description = f6
                        p.source = f7
                        p.target = f8
                        p.status = "COMP"
                        'p.trandate = p.trandate & Left(f9, 5) & ":00"
                        p.trandate = p.trandate & Format(f9, "hh:mm:ss")
                        cfile = Form1.logdir & "\2017\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'cfile = "c:\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'MsgBox cfile
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
                    End If
                Loop
                Close #5
            End If
        End If
    
        If grid1.TextMatrix(i, 4) = "SR2" Then
            p.id = grid1.TextMatrix(i, 1)
            p.area = "SR-2"
            p.description = grid1.TextMatrix(i, 3)
            p.source = " "
            p.target = " "
            p.product = grid1.TextMatrix(i, 6)
            wbc = grid1.TextMatrix(i, 7)
            wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
            p.palletid = wbc
            p.qty = grid1.TextMatrix(i, 8)
            p.uom = grid1.TextMatrix(i, 9)
            p.lotnum = grid1.TextMatrix(i, 10)
            p.units = grid1.TextMatrix(i, 11)
            p.lotnum2 = grid1.TextMatrix(i, 12)
            p.units2 = grid1.TextMatrix(i, 13)
            p.status = " "
            p.userid = "SR-2"
            p.trandate = Left(grid1.TextMatrix(i, 16), 7)
            p.reqid = grid1.TextMatrix(i, 17)
            s = "V:\sr2\bin\SR2" & Mid(Text1, 1, 2) & Mid(Text1, 4, 2) & ".csv"
            'MsgBox s
            If Len(Dir(s)) > 0 Then
                psku = Trim(Left(p.palletid, 4))
                plot = p.lotnum
                pplt = Right(p.palletid, 3)
                pgrp = p.description
                t = p.palletid & " SKU=" & psku & " Lot=" & plot & " Pallet=" & pplt
                'MsgBox t
                Open s For Input As #5
                Do Until EOF(5)
                    Input #5, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    If f2 = psku And f3 = plot And f4 = pplt Then
                        p.description = f6
                        p.source = f7
                        p.target = f8
                        p.status = "COMP"
                        'p.trandate = p.trandate & Left(f9, 5) & ":00"
                        p.trandate = p.trandate & Format(f9, "hh:mm:ss")
                        cfile = Form1.logdir & "\2017\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'cfile = "c:\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'MsgBox cfile
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
                    End If
                Loop
                Close #5
            End If
        End If
    
        If grid1.TextMatrix(i, 4) = "SR3" Then
            p.id = grid1.TextMatrix(i, 1)
            p.area = "SR-3"
            p.description = grid1.TextMatrix(i, 3)
            p.source = " "
            p.target = " "
            p.product = grid1.TextMatrix(i, 6)
            wbc = grid1.TextMatrix(i, 7)
            wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
            p.palletid = wbc
            p.qty = grid1.TextMatrix(i, 8)
            p.uom = grid1.TextMatrix(i, 9)
            p.lotnum = grid1.TextMatrix(i, 10)
            p.units = grid1.TextMatrix(i, 11)
            p.lotnum2 = grid1.TextMatrix(i, 12)
            p.units2 = grid1.TextMatrix(i, 13)
            p.status = " "
            p.userid = "SR-3"
            p.trandate = Left(grid1.TextMatrix(i, 16), 7)
            p.reqid = grid1.TextMatrix(i, 17)
            s = "V:\sr3\bin\SR3" & Mid(Text1, 1, 2) & Mid(Text1, 4, 2) & ".csv"
            'MsgBox s
            If Len(Dir(s)) > 0 Then
                psku = Trim(Left(p.palletid, 4))
                plot = p.lotnum
                pplt = Right(p.palletid, 3)
                pgrp = p.description
                t = p.palletid & " SKU=" & psku & " Lot=" & plot & " Pallet=" & pplt
                'MsgBox t
                Open s For Input As #5
                Do Until EOF(5)
                    Input #5, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    If f2 = psku And f3 = plot And f4 = pplt Then
                        p.description = f6
                        p.source = f7
                        p.target = f8
                        p.status = "COMP"
                        'p.trandate = p.trandate & Left(f9, 5) & ":00"
                        p.trandate = p.trandate & Format(f9, "hh:mm:ss")
                        cfile = Form1.logdir & "\2017\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'cfile = "c:\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'MsgBox cfile
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
                    End If
                Loop
                Close #5
            End If
        End If
    
        If grid1.TextMatrix(i, 4) = "SR5" Then
            p.id = grid1.TextMatrix(i, 1)
            p.area = "SR-5"
            p.description = grid1.TextMatrix(i, 3)
            p.source = " "
            p.target = " "
            p.product = grid1.TextMatrix(i, 6)
            wbc = grid1.TextMatrix(i, 7)
            wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
            p.palletid = wbc
            p.qty = grid1.TextMatrix(i, 8)
            p.uom = grid1.TextMatrix(i, 9)
            p.lotnum = grid1.TextMatrix(i, 10)
            p.units = grid1.TextMatrix(i, 11)
            p.lotnum2 = grid1.TextMatrix(i, 12)
            p.units2 = grid1.TextMatrix(i, 13)
            p.status = " "
            p.userid = "SR-5"
            p.trandate = Left(grid1.TextMatrix(i, 16), 7)
            p.reqid = grid1.TextMatrix(i, 17)
            s = "V:\sr5\bin\SR5" & Mid(Text1, 1, 2) & Mid(Text1, 4, 2) & ".csv"
            'MsgBox s
            If Len(Dir(s)) > 0 Then
                psku = Trim(Left(p.palletid, 4))
                plot = p.lotnum
                pplt = Right(p.palletid, 3)
                pgrp = p.description
                t = p.palletid & " SKU=" & psku & " Lot=" & plot & " Pallet=" & pplt
                'MsgBox t
                Open s For Input As #5
                Do Until EOF(5)
                    Input #5, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    'If f2 = psku And f3 = plot And f4 = pplt Then
                    If f5 = p.palletid Then
                        p.description = f6
                        p.source = "SR-5"
                        p.target = f8
                        p.status = "COMP"
                        'p.trandate = p.trandate & Left(f9, 5) & ":00"
                        p.trandate = p.trandate & Format(f9, "hh:mm:ss")
                        cfile = Form1.logdir & "\2017\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'cfile = "c:\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'MsgBox cfile
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
                    End If
                Loop
                Close #5
            End If
        End If
    
    
    Next i
    Screen.MousePointer = 0
End Sub

Private Sub psrlogsques_Click()
    Dim cfile As String, p As ptask, i As Integer, s As String, wbc As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim psku As String, plot As String, pplt As String, pgrp As String
    If Combo1 <> "Traffic Master" Then Exit Sub
    Screen.MousePointer = 11
    For i = 1 To grid1.Rows - 1
        If grid1.TextMatrix(i, 5) = "SR1" Then
            p.id = grid1.TextMatrix(i, 1)
            p.area = "SR-1"
            p.description = grid1.TextMatrix(i, 3)
            p.source = " "
            p.target = " "
            p.product = grid1.TextMatrix(i, 6)
            wbc = grid1.TextMatrix(i, 7)
            wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
            p.palletid = wbc
            p.qty = grid1.TextMatrix(i, 8)
            p.uom = grid1.TextMatrix(i, 9)
            p.lotnum = grid1.TextMatrix(i, 10)
            p.units = grid1.TextMatrix(i, 11)
            p.lotnum2 = grid1.TextMatrix(i, 12)
            p.units2 = grid1.TextMatrix(i, 13)
            p.status = " "
            p.userid = "SR-1"
            p.trandate = Left(grid1.TextMatrix(i, 16), 7)
            p.reqid = grid1.TextMatrix(i, 17)
            s = "V:\sr1\bin\SR1" & Mid(Text1, 1, 2) & Mid(Text1, 4, 2) & ".csv"
            'MsgBox s
            If Len(Dir(s)) > 0 Then
                psku = Trim(Left(p.palletid, 4))
                plot = p.lotnum
                pplt = Right(p.palletid, 3)
                pgrp = p.description
                t = p.palletid & " SKU=" & psku & " Lot=" & plot & " Pallet=" & pplt
                'MsgBox t
                Open s For Input As #5
                Do Until EOF(5)
                    Input #5, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    If f2 = psku And f3 = plot And f4 = pplt Then
                        p.description = f6
                        p.source = f7
                        p.target = f8
                        p.status = "COMP"
                        'p.trandate = p.trandate & Left(f9, 5) & ":00"
                        p.trandate = p.trandate & Format(f9, "hh:mm:ss")
                        cfile = Form1.logdir & "\2017\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'cfile = "c:\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'MsgBox cfile
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
                    End If
                Loop
                Close #5
            End If
        End If
    
        If grid1.TextMatrix(i, 5) = "SR2" Then
            p.id = grid1.TextMatrix(i, 1)
            p.area = "SR-2"
            p.description = grid1.TextMatrix(i, 3)
            p.source = " "
            p.target = " "
            p.product = grid1.TextMatrix(i, 6)
            wbc = grid1.TextMatrix(i, 7)
            wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
            p.palletid = wbc
            p.qty = grid1.TextMatrix(i, 8)
            p.uom = grid1.TextMatrix(i, 9)
            p.lotnum = grid1.TextMatrix(i, 10)
            p.units = grid1.TextMatrix(i, 11)
            p.lotnum2 = grid1.TextMatrix(i, 12)
            p.units2 = grid1.TextMatrix(i, 13)
            p.status = " "
            p.userid = "SR-2"
            p.trandate = Left(grid1.TextMatrix(i, 16), 7)
            p.reqid = grid1.TextMatrix(i, 17)
            s = "V:\sr2\bin\SR2" & Mid(Text1, 1, 2) & Mid(Text1, 4, 2) & ".csv"
            'MsgBox s
            If Len(Dir(s)) > 0 Then
                psku = Trim(Left(p.palletid, 4))
                plot = p.lotnum
                pplt = Right(p.palletid, 3)
                pgrp = p.description
                t = p.palletid & " SKU=" & psku & " Lot=" & plot & " Pallet=" & pplt
                'MsgBox t
                Open s For Input As #5
                Do Until EOF(5)
                    Input #5, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    If f2 = psku And f3 = plot And f4 = pplt Then
                        p.description = f6
                        p.source = f7
                        p.target = f8
                        p.status = "COMP"
                        'p.trandate = p.trandate & Left(f9, 5) & ":00"
                        p.trandate = p.trandate & Format(f9, "hh:mm:ss")
                        cfile = Form1.logdir & "\2017\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'cfile = "c:\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'MsgBox cfile
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
                    End If
                Loop
                Close #5
            End If
        End If
    
        If grid1.TextMatrix(i, 5) = "SR3" Then
            p.id = grid1.TextMatrix(i, 1)
            p.area = "SR-3"
            p.description = grid1.TextMatrix(i, 3)
            p.source = " "
            p.target = " "
            p.product = grid1.TextMatrix(i, 6)
            wbc = grid1.TextMatrix(i, 7)
            wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
            p.palletid = wbc
            p.qty = grid1.TextMatrix(i, 8)
            p.uom = grid1.TextMatrix(i, 9)
            p.lotnum = grid1.TextMatrix(i, 10)
            p.units = grid1.TextMatrix(i, 11)
            p.lotnum2 = grid1.TextMatrix(i, 12)
            p.units2 = grid1.TextMatrix(i, 13)
            p.status = " "
            p.userid = "SR-3"
            p.trandate = Left(grid1.TextMatrix(i, 16), 7)
            p.reqid = grid1.TextMatrix(i, 17)
            s = "V:\sr3\bin\SR3" & Mid(Text1, 1, 2) & Mid(Text1, 4, 2) & ".csv"
            'MsgBox s
            If Len(Dir(s)) > 0 Then
                psku = Trim(Left(p.palletid, 4))
                plot = p.lotnum
                pplt = Right(p.palletid, 3)
                pgrp = p.description
                t = p.palletid & " SKU=" & psku & " Lot=" & plot & " Pallet=" & pplt
                'MsgBox t
                Open s For Input As #5
                Do Until EOF(5)
                    Input #5, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    If f2 = psku And f3 = plot And f4 = pplt Then
                        p.description = f6
                        p.source = f7
                        p.target = f8
                        p.status = "COMP"
                        'p.trandate = p.trandate & Left(f9, 5) & ":00"
                        p.trandate = p.trandate & Format(f9, "hh:mm:ss")
                        cfile = Form1.logdir & "\2017\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'cfile = "c:\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'MsgBox cfile
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
                    End If
                Loop
                Close #5
            End If
        End If
    
        If grid1.TextMatrix(i, 5) = "SR5" Then
            p.id = grid1.TextMatrix(i, 1)
            p.area = "SR-5"
            p.description = grid1.TextMatrix(i, 3)
            p.source = " "
            p.target = " "
            p.product = grid1.TextMatrix(i, 6)
            wbc = grid1.TextMatrix(i, 7)
            wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
            p.palletid = wbc
            p.qty = grid1.TextMatrix(i, 8)
            p.uom = grid1.TextMatrix(i, 9)
            p.lotnum = grid1.TextMatrix(i, 10)
            p.units = grid1.TextMatrix(i, 11)
            p.lotnum2 = grid1.TextMatrix(i, 12)
            p.units2 = grid1.TextMatrix(i, 13)
            p.status = " "
            p.userid = "SR-5"
            p.trandate = Left(grid1.TextMatrix(i, 16), 7)
            p.reqid = grid1.TextMatrix(i, 17)
            s = "V:\sr5\bin\SR5" & Mid(Text1, 1, 2) & Mid(Text1, 4, 2) & ".csv"
            'MsgBox s
            If Len(Dir(s)) > 0 Then
                psku = Trim(Left(p.palletid, 4))
                plot = p.lotnum
                pplt = Right(p.palletid, 3)
                pgrp = p.description
                t = p.palletid & " SKU=" & psku & " Lot=" & plot & " Pallet=" & pplt
                'MsgBox t
                Open s For Input As #5
                Do Until EOF(5)
                    Input #5, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
                    'If f2 = psku And f3 = plot And f4 = pplt Then
                    If f5 = p.palletid Then
                        p.description = f6
                        p.source = f7
                        p.target = f8
                        p.status = "COMP"
                        'p.trandate = p.trandate & Left(f9, 5) & ":00"
                        p.trandate = p.trandate & Format(f9, "hh:mm:ss")
                        cfile = Form1.logdir & "\2017\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'cfile = "c:\SR" & Format(Text1, "mmddyyyy") & ".txt"
                        'MsgBox cfile
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
                    End If
                Loop
                Close #5
            End If
        End If
    
    
    Next i
    Screen.MousePointer = 0

End Sub

Private Sub pstot_Click()
    Dim i As Integer, k As Integer, s As String, aflag As Boolean
    Dim rt As String, rf As String, rh As String
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) = "Shipping" Then Combo1.ListIndex = i
    Next i
    DoEvents
    If grid1.Rows < 2 Then Exit Sub
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 3
    s = grid1.TextMatrix(1, 5) & Chr(9) & grid1.TextMatrix(1, 6)
    pgrid.AddItem s
    For i = 1 To grid1.Rows - 1
        aflag = True
        For k = 1 To pgrid.Rows - 1
            If pgrid.TextMatrix(k, 0) = grid1.TextMatrix(i, 5) And pgrid.TextMatrix(k, 1) = grid1.TextMatrix(i, 6) Then
                pgrid.TextMatrix(k, 2) = Val(pgrid.TextMatrix(k, 2)) + Val(grid1.TextMatrix(i, 11)) + Val(grid1.TextMatrix(i, 13))
                aflag = False
                Exit For
            End If
        Next k
        If aflag = True Then
            s = grid1.TextMatrix(i, 5) & Chr(9) & grid1.TextMatrix(i, 6) & Chr(9) & Val(grid1.TextMatrix(i, 11)) + Val(grid1.TextMatrix(i, 13))
            pgrid.AddItem s
        End If
    Next i
    
    pgrid.RowSel = pgrid.Row
    pgrid.Col = 0: pgrid.ColSel = 1
    pgrid.Sort = 5
    
    s = "<Trailer|<Product|^Units"
    pgrid.FormatString = s
    pgrid.ColWidth(0) = 2800
    pgrid.ColWidth(1) = 4000
    pgrid.ColWidth(2) = 1200
    rt = pstot.Caption
    rh = Text1 & "  " & "Shipping"
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, pgrid, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
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

Private Sub shiphist_Click()
    Call ship_history
End Sub

Private Sub sortbc_Click()
    grid1.Row = 0: grid1.RowSel = 0
    grid1.Col = 7: grid1.ColSel = 7
    grid1.Sort = 5
End Sub

Private Sub sortdt_Click()
    grid1.Row = 0: grid1.RowSel = 0
    grid1.Col = 16: grid1.ColSel = 16
    grid1.Sort = 5
End Sub

Private Sub sortshiptrig_Change()
    Dim i As Integer
    i = grid1.Row
    grid1.Redraw = False
    sort_ship_history
    grid1.Row = i
    grid1.Redraw = True
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If MsgBox("Add Rack Logs to list.", vbYesNo + vbQuestion, "Howdy!!!!!!!!!!!") = vbYes Then
            Combo1.AddItem "Rack Activity"
        End If
    End If
End Sub

Private Sub widrpt_Click()
    Call withdrawal
End Sub
