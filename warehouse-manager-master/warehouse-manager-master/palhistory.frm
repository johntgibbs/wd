VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form palhistory 
   Caption         =   "Pallet History"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14295
   LinkTopic       =   "Form14"
   ScaleHeight     =   4020
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Pallet Correction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
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
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid opgrid 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6165
      _Version        =   327680
      ForeColor       =   128
      BackColorFixed  =   65535
      AllowUserResizing=   3
   End
   Begin VB.Label Label1 
      Caption         =   "Barcode:"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label barkey 
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
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "palhistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub pallet_history(pbarcode As String)
    Dim ds As ADODB.Recordset, ss As ADODB.Recordset
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String, wsku As String, wlot As String
    Dim wzone As String, wstat As String, wgma As Integer, wside As String
    Dim waisle As String, wrack As String, hrow As Boolean, ocode As String
    Dim cfile As String, s As String, bc As String, srflag As Boolean
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim dl As Long, wbc As String, citem As String
    Dim syear As Integer, eyear As Integer, i As Integer, j As Integer
    Dim logpath As String
    logpath = Form1.logdir
    srpath = logpath                                            'jv060117
    wbc = pbarcode
    If Len(wbc) = 0 Then Exit Sub
    wsku = Trim(Left(wbc, 4))
    wlot = barcode_to_lotnum(wbc)
    s = wbc                                                             'jv012116
    sdate = Format(Val(Mid(s, 9, 2)) - 2, "00")                         'jv012116
    sdate = "20" & sdate & Mid(s, 5, 4)                                 'jv012116
    edate = Format(Now, "yyyymmdd")                                     'jv012116
    
    Screen.MousePointer = 11
    opgrid.Visible = True
    opgrid.Clear: opgrid.Cols = 19: opgrid.Rows = 1
    opgrid.Left = 0: opgrid.Width = Me.Width - 200
    
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
                'If r12flag = True Then
                    s = s & Mid(ds!barcode, 5, 9) & Chr(9)                  'jv071615
                'Else
                '    s = s & ds!lot_num & Chr(9)                 'Lot1
                'End If
                s = s & ds!count_qty & Chr(9)                   'Units
                'If r12flag = True Then
                    s = s & r12_lot(ds!lot2, ocode) & Chr(9)
                'Else
                '    s = s & ds!lot2 & Chr(9)                    'Lot2
                'End If
                s = s & ds!qty2 & Chr(9)                        'Units
                s = s & "In-Stock" & Chr(9)                     'Status
                s = s & "WMS" & Chr(9)                          'User
                s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
                s = s & " "                                     'Reqid
                opgrid.AddItem s
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
                ''s = "select * from vContainerLocation_1033 Where [pal id] = '" & ds!plateno & "'"   'jv081415
                's = "select * from vContainerLocation_1033 Where [pal id] in "      'jv092816
                's = s & "('" & ds!plateno & "', '" & ds!barcode & "')"              'jv092816
                's = s & " and item = '" & citem & "'"                                   'jv100516
                s = "select * from vAllInventory_1033 where lpn = '" & ds!barcode & "'"
                's = "select * from vAllInventory_1033 where lpn = '777 053020226038'"
                'MsgBox s
                Set ds6 = db5.Execute(s)
                If ds6.BOF = False Then
                    ds6.MoveFirst
                    s = "OH" & Chr(9)
                    s = s & ds6(0) & Chr(9)         'container id
                    s = s & "Crane" & Chr(9)        'area
                    's = s & ds6(17) & Chr(9)        'Hold reason
                    If ds6!lock > 0 Then s = s & "Locked"
                    s = s & Chr(9)
                    s = s & "CS5" & Chr(9)          'Source
                    s = s & ds6!location & Chr(9)   'Target
                    s = s & plit & Chr(9)           'Product
                    s = s & bc000(ds!barcode) & Chr(9)     'Pallet
                    s = s & "1" & Chr(9)            'Qty
                    'If Trim(ds6!Type) = "BBCPallet" Then  'UOM
                    If Trim(ds6(4)) = "BBCPallet" Then  'UOM
                        s = s & "BBC" & Chr(9)
                    Else
                        s = s & "GMA" & Chr(9)
                    End If
                    ocode = Mid(ds!barcode, 11, 3)                              'jv071715
                    'If r12flag = True Then
                        s = s & Mid(ds!barcode, 5, 9) & Chr(9)                  'jv071615
                    'Else
                    '    s = s & ds!lot1 & Chr(9)                 'Lot1
                    'End If
                    s = s & ds!qty1 & Chr(9)                   'Units
                    'If r12flag = True Then
                        s = s & r12_lot(ds!lot2, ocode) & Chr(9)
                    'Else
                    '    s = s & ds!lot2 & Chr(9)                    'Lot2
                    'End If
                    s = s & ds!qty2 & Chr(9)                        'Units
                    s = s & "In-Stock" & Chr(9)                     'Status
                    s = s & "WMS" & Chr(9)                          'User
                    s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
                    s = s & ds!plateno                                     'Reqid
                    opgrid.AddItem s
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
            'If r12flag = True Then
                s = s & Mid(ds!barcode, 5, 9) & Chr(9)                      'jv071715
            'Else
            '    s = s & ds!lot_num & Chr(9)                 'Lot1
            'End If
            s = s & ds!count_qty & Chr(9)                   'Units
            'If r12flag = True Then
                s = s & r12_lot(ds!lot2, ocode) & Chr(9)
            'Else
            '    s = s & ds!lot2 & Chr(9)                    'Lot2
            'End If
            s = s & ds!qty2 & Chr(9)                        'Units
            s = s & "In-Stock" & Chr(9)                     'Status
            s = s & "WMS" & Chr(9)                          'User
            s = s & Format(Now, "yyMMdd hh:mm:ss") & Chr(9)
            s = s & " "                                     'Reqid
            opgrid.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    
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
                    opgrid.AddItem s
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
                        opgrid.AddItem s
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
                    opgrid.AddItem s
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
                        opgrid.AddItem s
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
                    opgrid.AddItem s
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
                        opgrid.AddItem s
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
                    opgrid.AddItem s
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
                        opgrid.AddItem s
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
                    opgrid.AddItem s
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
                        opgrid.AddItem s
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
                    opgrid.AddItem s
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
                        opgrid.AddItem s
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
                        opgrid.AddItem s
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
                            opgrid.AddItem s
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
    
    
    'srflag = False
    'If Form1.plantno = 50 Then
    '    If MsgBox("Include SR Logs?", vbQuestion + vbYesNo, "SR Logs....") = vbYes Then srflag = True
    'End If
    'If srflag = True Then
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
                    opgrid.AddItem s
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
                        opgrid.AddItem s
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
    'End If
    
    
    opgrid.Redraw = False
    j = 0
    If opgrid.Rows > 1 Then
        For i = 1 To opgrid.Rows - 1
            If opgrid.TextMatrix(i, 0) = "PR" Then
                opgrid.TextMatrix(i, 18) = opgrid.TextMatrix(i, 7) & "0" & opgrid.TextMatrix(i, 16)
            Else
                opgrid.TextMatrix(i, 18) = opgrid.TextMatrix(i, 7) & opgrid.TextMatrix(i, 16) & opgrid.TextMatrix(i, 0)
            End If
            If opgrid.TextMatrix(i, 0) = "OH" Then j = j + 1
        Next i
    End If
    
    's = "^Type|^RecId|<Area|<Description|<Source|<Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId"
    s = "^Type||<Area|<Description|<Source|<Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units||^User|<Time|"
    opgrid.FormatString = s
    opgrid.ColWidth(0) = 600
    opgrid.ColWidth(1) = 1 '600
    opgrid.ColWidth(2) = 1500
    opgrid.ColWidth(3) = 1000
    opgrid.ColWidth(4) = 1300
    opgrid.ColWidth(5) = 1500
    opgrid.ColWidth(6) = 3000
    opgrid.ColWidth(7) = 1800
    opgrid.ColWidth(8) = 600
    opgrid.ColWidth(9) = 800
    opgrid.ColWidth(10) = 900
    opgrid.ColWidth(11) = 800
    opgrid.ColWidth(12) = 900
    opgrid.ColWidth(13) = 800
    opgrid.ColWidth(14) = 1 '800
    opgrid.ColWidth(15) = 1600
    opgrid.ColWidth(16) = 1400
    opgrid.ColWidth(17) = 1 '1000
    opgrid.ColWidth(18) = 1 '2100
    opgrid.RowSel = opgrid.Row
    opgrid.Col = 16: opgrid.ColSel = 16
    opgrid.Sort = 5
    opgrid.FillStyle = flexFillRepeat
    opgrid.Redraw = True
    Screen.MousePointer = 0
    If j > 1 Then
        s = "This pallet label has been found at " & j & " locations."
        MsgBox s, vbOKOnly + vbInformation, "multiple locations found...."
    End If
End Sub


Private Sub barkey_Change()
    If Len(barkey) = 16 Then
        If Val(Mid(barkey, 11, 3)) > 0 Then
            Call pallet_history(barkey.Caption)
        End If
    End If
End Sub

Private Sub Command1_Click()                'Print List
    Dim rt As String, rh As String, rf As String
    rt = "Pallet History - " & barkey.Caption
    rh = barkey.Caption
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    opgrid.Redraw = False
    'If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
    '    Call printflexgrid(Printer, Grid1, rt, rh, rf)
    'Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", opgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    'End If
    opgrid.Redraw = True
End Sub

Private Sub Command2_Click()
    Dim wbc As String
    palcorr.pref = opgrid.TextMatrix(opgrid.Row, 2)
    wbc = opgrid.TextMatrix(opgrid.Row, 7)
    wbc = Mid(wbc, 1, 10) & Mid(wbc, 13, 3) & Mid(wbc, 18, 3)   'undo bc000
    palcorr.bckey = wbc
    palcorr.Show
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Width = Screen.Width - 200
End Sub

Private Sub Form_Resize()
    opgrid.Width = Me.Width - 220
End Sub
