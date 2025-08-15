VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form16 
   Caption         =   "Forklift Shipping and Pallet Builds"
   ClientHeight    =   9705
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13560
   LinkTopic       =   "Form16"
   ScaleHeight     =   9705
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   1815
      Left            =   0
      TabIndex        =   11
      Top             =   7800
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3201
      _Version        =   327680
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   6480
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   960
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10560
      TabIndex        =   5
      Top             =   0
      Width           =   2535
      Begin VB.OptionButton Option2 
         Caption         =   "Pallets"
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
         Left            =   1440
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Trailers"
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
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   960
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   10335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8895
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   15690
      _Version        =   327680
      ForeColor       =   4194368
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
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
      Left            =   1440
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   8895
   End
   Begin VB.Label rc 
      Caption         =   "0"
      Height          =   255
      Left            =   13200
      TabIndex        =   9
      Top             =   240
      Width           =   495
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Active / Complete Tasks"
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
      Left            =   7800
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Trailers:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Width           =   1095
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu edsrc 
         Caption         =   "Change Source"
      End
      Begin VB.Menu edbc 
         Caption         =   "Change BarCode"
      End
      Begin VB.Menu edpalsize 
         Caption         =   "Change Pallet Size"
      End
      Begin VB.Menu edloc 
         Caption         =   "Change Picked Pallet Location"
      End
      Begin VB.Menu mtc 
         Caption         =   "Mark Task - Complete"
      End
      Begin VB.Menu mtp 
         Caption         =   "Mark Task - Pending"
      End
      Begin VB.Menu cu 
         Caption         =   "Clear User"
      End
      Begin VB.Menu mac 
         Caption         =   "Mark All - Complete"
      End
      Begin VB.Menu map 
         Caption         =   "Mark All - Pending"
      End
      Begin VB.Menu addalt 
         Caption         =   "Add Alternate Product"
      End
      Begin VB.Menu pco 
         Caption         =   "Print Check off"
      End
      Begin VB.Menu prtlab 
         Caption         =   "Print Pallet Barcode Label"
      End
      Begin VB.Menu batonhand 
         Caption         =   "View Batch Inventory"
      End
      Begin VB.Menu palhist 
         Caption         =   "View Pallet History"
      End
   End
   Begin VB.Menu impmenu 
      Caption         =   "Import"
      Begin VB.Menu imptrl 
         Caption         =   "Branch Trailers from Shipping"
      End
      Begin VB.Menu impjob 
         Caption         =   "Jobbing Groups from Cranes"
      End
   End
   Begin VB.Menu usermenu 
      Caption         =   "User"
      Visible         =   0   'False
      Begin VB.Menu emplook 
         Caption         =   "Lookup Employee Name"
      End
   End
   Begin VB.Menu resmenu 
      Caption         =   "Re-Stock"
      Begin VB.Menu restock 
         Caption         =   "Generate Re-stock tasks for Dock "
      End
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub post_wms(rk As Integer)
    Dim cfile As String, p As ptask
    p.id = Grid1.TextMatrix(rk, 0)
    p.area = Grid1.TextMatrix(rk, 1)
    p.description = " "
    p.source = Grid1.TextMatrix(rk, 2)
    p.target = Grid1.TextMatrix(rk, 3)
    p.product = Grid1.TextMatrix(rk, 4)
    p.palletid = Grid1.TextMatrix(rk, 5)
    p.qty = Grid1.TextMatrix(rk, 6)
    p.uom = Grid1.TextMatrix(rk, 7)
    p.lotnum = Grid1.TextMatrix(rk, 9)
    p.units = Grid1.TextMatrix(rk, 8)
    p.lotnum2 = Grid1.TextMatrix(rk, 11)
    p.units2 = Grid1.TextMatrix(rk, 10)
    p.status = Grid1.TextMatrix(rk, 12)
    p.userid = Grid1.TextMatrix(rk, 13)
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")
    p.reqid = ".."
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"
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
End Sub

Private Sub post_jobbing(gno As String)
    Dim ds As ADODB.Recordset, s As String
    Dim ss As ADODB.Recordset, pwraps
    Dim punit As String, pdesc As String, ppal As Integer
    Dim p As ptask, i As Integer, k As Integer
    s = "select * from paltasks where area = 'GROUP'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If Left(ds!product, Len(gno)) = gno Then
                s = "Update paltasks set area = 'NONE', status = 'COMP' Where id = " & ds!id
                Wdb.Execute s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select status,area,product from paltasks where area in ('FORKLIFT', 'DOCK')"
    s = s & " and description >= '" & gno & "'"
    s = s & " and description < '" & gno & "ZZZZ'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update paltasks set area = 'NONE', status = 'COMP' Where id = " & ds!id
            Wdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select * from ship_infc where order_num = '" & gno & "'"
    s = s & " and ship_status in ('NEW','ACTV')"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        rc.Caption = Val(rc.Caption) + 1
        DoEvents
        'Build Header
        p.area = "GROUP"
        p.description = " "
        p.source = gno & " ZO"
        p.target = "..."
        p.product = gno & Space(8 - Len(gno)) & " " & gno & " ZO"
        p.palletid = "..."
        p.qty = 0
        p.uom = " "
        p.lotnum = " "
        p.units = 0
        p.lotnum2 = " "
        p.units2 = 0
        p.status = "PEND"
        p.userid = " "
        p.trandate = Format(Now, "yymmdd hh:mm:ss")
        p.reqid = " "
        Call insert_trans(p)
        Do Until ds.EOF
            rc.Caption = Val(rc.Caption) + 1
            DoEvents
            k = ds!order_qty - ds!ship_plt_qty
            If k > 0 Then
                For i = 1 To k
                    If skurec(Val(ds!sku)).sku = ds!sku Then
                        punit = skurec(Val(ds!sku)).uom_type
                        pdesc = skurec(Val(ds!sku)).desc
                        ppal = skurec(Val(ds!sku)).uom_per_pallet
                        pwraps = skurec(Val(ds!sku)).uom_per_pallet / skurec(Val(ds!sku)).qty_per_pallet
                    Else
                        punit = "Invalid"
                        pdesc = "SKU"
                        ppal = 1
                        pwraps = 1
                    End If
                    If ds!to_whse_num <> 4 Then
                        p.area = "DOCK"
                        p.description = ds!order_num
                        p.source = "SR" & ds!to_whse_num
                        p.target = "STAGING " & ds!order_num        'jv062211
                        p.product = ds!sku & " " & punit & " " & pdesc
                        p.palletid = "..."
                        p.qty = 1
                        p.uom = "Pallet"
                        p.lotnum = " "
                        p.units = ppal
                        p.lotnum2 = " "
                        p.units2 = 0
                        p.status = "PEND"
                        p.userid = " "
                        p.trandate = Format(Now, "yymmdd hh:mm:ss")
                        p.reqid = " "
                    Else
                        p.area = "FORKLIFT"
                        p.description = gno & Space(8 - Len(gno)) & " " & gno & " ZO"
                        If ds!gmasize > 0 Then
                            p.source = "4WAY"
                        Else
                            p.source = "RACKS"
                        End If
                        p.target = "ORDER PICK"
                        p.product = ds!sku & " " & punit & " " & pdesc
                        p.palletid = "..."
                        p.qty = 1
                        p.uom = "Pallet"
                        p.lotnum = " "
                        If ds!gmasize > 0 Then
                            p.units = ds!gmasize * pwraps
                        Else
                            p.units = ppal
                        End If
                        p.lotnum2 = " "
                        p.units2 = 0
                        p.status = "PEND"
                        p.userid = " "
                        p.trandate = Format(Now, "yymmdd hh:mm:ss")
                        p.reqid = " "
                    End If
                    Call insert_trans(p)
                Next i
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub


Private Sub postgroup(gno As String)
    Dim ds As ADODB.Recordset, s As String
    s = "select * from paltasks where area = 'GROUP'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If Left(ds!product, Len(gno)) = gno Then
                s = "Update paltasks set area = 'NONE', status = 'COMP' Where id = " & ds!id
                Wdb.Execute s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select id,area,status from paltasks where area in ('FORKLIFT', 'DOCK')"
    s = s & " and description >= '" & gno & "'"
    s = s & " and description < '" & gno & "ZZZZ'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update paltasks set area = 'NONE', status = 'COMP' Where id = " & ds!id
            Wdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select distinct runid from trailers where groupcode = '" & gno & "'"
    s = s & " and pb_flag = 'N'"          'jv090911
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Call build_run_file(ds(0))
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Private Sub build_run_file(rno As Long)
    Dim ds As ADODB.Recordset, s As String
    Dim lname As String, tno As String, msku As String
    Dim mlot As String, mpal As String, mpdesc As String, mwhs As String, mline As String
    Dim gc As String, lbr As Integer, dbr As Integer, tdate As String
    Dim p As ptask
    Dim bagc As String
    rc.Caption = "0"
    DoEvents
    lname = "???": tno = "??"
    
    s = "select plant,trailers.branch,branchname,trlno,shipdate from trailers,branches"
    s = s & " where runid = " & rno
    s = s & " and branches.branch = trailers.branch"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        lname = ds!branchname
        tno = ds!trlno
        lbr = ds!plant
        dbr = ds(1)
        tdate = Format(ds(4), "mm-dd-yyyy")
        bagc = "T" & Mid(tdate, 4, 2) & Format(Val(dbr), "00") & Mid(tno, 2, 1)        'jv090911
    Else
        ds.Close
        Exit Sub
    End If
    ds.Close
    s = "select groupcode,trailers.sku,pallets,trailers.whs_num,fgunit,fgdesc,pallet"
    s = s & " from trailers,skumast where trailers.runid = " & rno
    s = s & " and skumast.sku = trailers.sku"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        rc.Caption = Val(rc.Caption) + 1
        DoEvents
        If Form1.plantno = 50 Then                                          'jv090911
            gc = ds!groupcode
        Else                                                                'jv090911
            gc = bagc                                                       'jv090911
        End If                                                              'jv090911
        'Build Header
        s = "GROUP" & Chr(9) & Chr(9)
        p.area = "GROUP"
        p.description = " "
        s = s & lname & " " & tno & Chr(9)      'Source = Trailer Name
        p.source = lname & " " & tno
        s = s & "..." & Chr(9)                  'Target Door
        p.target = "..."
        s = s & gc & Space(8 - Len(gc)) & lname & " " & tno & Chr(9)    'Group Code and Trailer Name
        p.product = gc & Space(8 - Len(gc)) & lname & " " & tno
        s = s & "..." & Chr(9)
        p.palletid = "..."
        s = s & " " & Chr(9)
        p.qty = 0
        s = s & " " & Chr(9)
        p.uom = " "
        s = s & " " & Chr(9)
        p.lotnum = " "
        s = s & " " & Chr(9)
        p.units = 0
        s = s & " " & Chr(9)
        p.lotnum2 = " "
        s = s & " " & Chr(9)
        p.units2 = 0
        s = s & "PEND" & Chr(9)
        p.status = "PEND"
        s = s & " " & Chr(9)
        p.userid = " "
        s = s & Format(Now, "yymmdd hh:mm:ss") & Chr(9)
        p.trandate = Format(Now, "yymmdd hh:mm:ss")
        s = s & " "
        p.reqid = rno                                                   'jv021815
        Call insert_trans(p)
        Do Until ds.EOF
            rc.Caption = Val(rc.Caption) + 1
            DoEvents
            If ds!pallets > 0 Then
                'If ds(3) <= 3 Then                      'Cranes
                If ds(3) <= 3 Or ds(3) = 5 Then          'Cranes   dai0625
                    For i = 1 To ds!pallets
                        s = "DOCK" & Chr(9)
                        p.area = "DOCK"
                        s = s & ds!groupcode & Chr(9)
                        'p.description = ds!groupcode
                        p.description = gc                           'jv090911
                        s = s & "SR" & ds(3) & Chr(9)
                        p.source = "SR" & ds(3)
                        If tno = "ZO" Or tno = "OP" Or tno = "QC" Then                   'jv121015
                            s = s & "STAGING " & ds!groupcode & Chr(9)      'jv062211
                            p.target = "STAGING " & ds!groupcode            'jv062211
                        Else
                            s = s & lname & " " & tno & Chr(9)
                            p.target = lname & " " & tno
                        End If
                        s = s & ds(1) & " " & ds!fgunit & " " & ds!fgdesc & Chr(9)
                        p.product = ds(1) & " " & ds!fgunit & " " & ds!fgdesc
                        If Form1.plantno = "52" Or ds(3) = 5 Then   'dai0625
                            s = s & ds(1) & " ...... . ..." & Chr(9)
                            p.palletid = ds(1) & " ...... . ..."
                        Else
                            s = s & "..." & Chr(9)
                            p.palletid = "..."
                        End If
                        s = s & "1" & Chr(9)            'Move Qty
                        p.qty = 1
                        s = s & "Pallet" & Chr(9)       'Uom
                        p.uom = "Pallet"
                        s = s & "..." & Chr(9)          'Lot
                        p.lotnum = "..."
                        s = s & ds!pallet & Chr(9)      'Units
                        p.units = ds!pallet
                        s = s & "..." & Chr(9)          'Lot2
                        p.lotnum2 = "..."
                        s = s & " " & Chr(9)            'Units
                        p.units2 = 0
                        s = s & "PEND" & Chr(9)
                        p.status = "PEND"
                        s = s & " " & Chr(9)
                        p.userid = " "
                        s = s & Format(Now, "yymmdd hh:mm:ss") & Chr(9)
                        p.trandate = Format(Now, "yymmdd hh:mm:ss")
                        s = s & " "
                        p.reqid = " "
                        Call insert_trans(p)
                    Next i
                End If
                'Generic rack assignment
                'If ds(3) = 4 Or ds(3) = 5 Or ds(3) = 14 Or ds(3) = 15 Then              'Regular, Regular A, Broken Arrow, Sylacauga   jv090911
                If ds(3) = 4 Or ds(3) = 7 Or ds(3) = 14 Or ds(3) = 15 Then    'dai0625          'Regular, Regular A, Broken
                    For i = 1 To ds!pallets
                        s = "FORKLIFT" & Chr(9)
                        p.area = "FORKLIFT"
                        s = s & ds!groupcode & Chr(9)
                        'p.description = ds!groupcode
                        p.description = gc & Space(8 - Len(gc)) & lname & " " & tno
                        s = s & "RACKS" & Chr(9)
                        p.source = "RACKS"
                        If tno = "ZO" Or tno = "OP" Then                                'jv091815
                            s = s & "ORDER PICK" & Chr(9)
                            p.target = "ORDER PICK"
                        Else
                            s = s & "STAGING" & Chr(9)
                            p.target = "STAGING"
                        End If
                        s = s & ds(1) & " " & ds!fgunit & " " & ds!fgdesc & Chr(9)
                        p.product = ds(1) & " " & ds!fgunit & " " & ds!fgdesc
                        s = s & "..." & Chr(9)
                        p.palletid = "..."
                        s = s & "1" & Chr(9)            'Move Qty
                        p.qty = 1
                        s = s & "Pallet" & Chr(9)       'Uom
                        p.uom = "Pallet"
                        s = s & "..." & Chr(9)          'Lot
                        p.lotnum = "..."
                        s = s & ds!pallet & Chr(9)      'Units
                        p.units = ds!pallet
                        s = s & "..." & Chr(9)          'Lot2
                        p.lotnum2 = "..."
                        s = s & " " & Chr(9)            'Units
                        p.units2 = 0
                        s = s & "PEND" & Chr(9)
                        p.status = "PEND"
                        s = s & " " & Chr(9)
                        p.userid = " "
                        s = s & Format(Now, "yymmdd hh:mm:ss") & Chr(9)
                        p.trandate = Format(Now, "yymmdd hh:mm:ss")
                        s = s & " "
                        p.reqid = " "
                        Call insert_trans(p)
                    
                        s = "DOCK" & Chr(9)
                        p.area = "DOCK"
                        s = s & ds!groupcode & Chr(9)
                        'p.description = ds!groupcode
                        p.description = gc                          'jv090911
                        s = s & "STAGING" & Chr(9)
                        p.source = "STAGING"
                        s = s & lname & " " & tno & Chr(9)
                        p.target = lname & " " & tno
                        s = s & ds(1) & " " & ds!fgunit & " " & ds!fgdesc & Chr(9)
                        p.product = ds(1) & " " & ds!fgunit & " " & ds!fgdesc
                        s = s & "..." & Chr(9)
                        p.palletid = "..."
                        s = s & "1" & Chr(9)            'Move Qty
                        p.qty = 1
                        s = s & "Pallet" & Chr(9)       'Uom
                        p.uom = "Pallet"
                        s = s & "..." & Chr(9)          'Lot
                        p.lotnum = "..."
                        s = s & ds!pallet & Chr(9)      'Units
                        p.units = ds!pallet
                        s = s & "..." & Chr(9)          'Lot2
                        p.lotnum2 = "..."
                        s = s & " " & Chr(9)            'Units
                        p.units2 = 0
                        s = s & "PEND" & Chr(9)
                        p.status = "PEND"
                        s = s & " " & Chr(9)
                        p.userid = " "
                        s = s & Format(Now, "yymmdd hh:mm:ss") & Chr(9)
                        p.trandate = Format(Now, "yymmdd hh:mm:ss")
                        s = s & " "
                        p.reqid = " "
                        If tno <> "ZO" And tno <> "OP" Then                         'jv091815
                            Call insert_trans(p)
                        End If
                    Next i
                End If
                
                If ds(3) = 13 Then                          '4Way Pallets
                    For i = 1 To ds!pallets
                        s = "FORKLIFT" & Chr(9)
                        p.area = "FORKLIFT"
                        s = s & ds!groupcode & Chr(9)
                        'p.description = ds!groupcode
                        p.description = gc & Space(8 - Len(gc)) & lname & " " & tno
                        s = s & "4WAY" & Chr(9)
                        p.source = "4WAY"
                        If tno = "ZO" Or tno = "OP" Then                            'jv091815
                            s = s & "ORDER PICK" & Chr(9)
                            p.target = "ORDER PICK"
                        Else
                            s = s & "STAGING" & Chr(9)
                            p.target = "STAGING"
                        End If
                        s = s & ds(1) & " " & ds!fgunit & " " & ds!fgdesc & Chr(9)
                        p.product = ds(1) & " " & ds!fgunit & " " & ds!fgdesc
                        s = s & "..." & Chr(9)
                        p.palletid = "..."
                        s = s & "1" & Chr(9)            'Move Qty
                        p.qty = 1
                        s = s & "Pallet" & Chr(9)       'Uom
                        p.uom = "Pallet"
                        s = s & "..." & Chr(9)          'Lot
                        p.lotnum = "..."
                        s = s & ds!pallet & Chr(9)      'Units
                        p.units = ds!pallet
                        s = s & "..." & Chr(9)          'Lot2
                        p.lotnum2 = "..."
                        s = s & " " & Chr(9)            'Units
                        p.units2 = 0
                        s = s & "PEND" & Chr(9)
                        p.status = "PEND"
                        s = s & " " & Chr(9)
                        p.userid = " "
                        s = s & Format(Now, "yymmdd hh:mm:ss") & Chr(9)
                        p.trandate = Format(Now, "yymmdd hh:mm:ss")
                        s = s & " "
                        p.reqid = " "
                        Call insert_trans(p)
                    
                        s = "DOCK" & Chr(9)
                        p.area = "DOCK"
                        s = s & ds!groupcode & Chr(9)
                        'p.description = ds!groupcode
                        p.description = gc                      'jv090911
                        s = s & "STAGING" & Chr(9)
                        p.source = "STAGING"
                        s = s & lname & " " & tno & Chr(9)
                        p.target = lname & " " & tno
                        s = s & ds(1) & " " & ds!fgunit & " " & ds!fgdesc & Chr(9)
                        p.product = ds(1) & " " & ds!fgunit & " " & ds!fgdesc
                        s = s & "..." & Chr(9)
                        p.palletid = "..."
                        s = s & "1" & Chr(9)            'Move Qty
                        p.qty = 1
                        s = s & "Pallet" & Chr(9)       'Uom
                        p.uom = "Pallet"
                        s = s & "..." & Chr(9)          'Lot
                        p.lotnum = "..."
                        s = s & ds!pallet & Chr(9)      'Units
                        p.units = ds!pallet
                        s = s & "..." & Chr(9)          'Lot2
                        p.lotnum2 = "..."
                        s = s & " " & Chr(9)            'Units
                        p.units2 = 0
                        s = s & "PEND" & Chr(9)
                        p.status = "PEND"
                        s = s & " " & Chr(9)
                        p.userid = " "
                        s = s & Format(Now, "yymmdd hh:mm:ss") & Chr(9)
                        p.trandate = Format(Now, "yymmdd hh:mm:ss")
                        s = s & " "
                        p.reqid = " "
                        If tno <> "ZO" And tno <> "OP" Then                             'jv091815
                            Call insert_trans(p)
                        End If
                    Next i
                End If
                
                
                If ds(3) = 6 Or ds(3) = 8 Then          'Snack Plant, Snack Plant Drop
                    For i = 1 To ds!pallets
                        s = "DOCK" & Chr(9)
                        p.area = "DOCK"
                        s = s & ds!groupcode & Chr(9)
                        'p.description = ds!groupcode
                        p.description = gc                      'jv090911
                        s = s & "SNACK PLANT" & Chr(9)
                        p.source = "SNACK PLANT"
                        If tno = "ZO" Or tno = "OP" Then                                'jv091815
                            s = s & "STAGING " & ds!groupcode & Chr(9)      'jv062211
                            p.target = "STAGING " & ds!groupcode            'jv062211
                        Else
                            s = s & lname & " " & tno & Chr(9)
                            p.target = lname & " " & tno
                        End If
                        s = s & ds(1) & " " & ds!fgunit & " " & ds!fgdesc & Chr(9)
                        p.product = ds(1) & " " & ds!fgunit & " " & ds!fgdesc
                        s = s & "..." & Chr(9)
                        p.palletid = "..."
                        s = s & "1" & Chr(9)            'Move Qty
                        p.qty = 1
                        s = s & "Pallet" & Chr(9)       'Uom
                        p.uom = "Pallet"
                        s = s & "..." & Chr(9)          'Lot
                        p.lotnum = "..."
                        s = s & ds!pallet & Chr(9)      'Units
                        p.units = ds!pallet
                        s = s & "..." & Chr(9)          'Lot2
                        p.lotnum2 = "..."
                        s = s & " " & Chr(9)            'Units
                        p.units2 = 0
                        s = s & "PEND" & Chr(9)
                        p.status = "PEND"
                        s = s & " " & Chr(9)
                        p.userid = " "
                        s = s & Format(Now, "yymmdd hh:mm:ss") & Chr(9)
                        p.trandate = Format(Now, "yymmdd hh:mm:ss")
                        s = s & " "
                        p.reqid = " "
                        Call insert_trans(p)
                    Next i
                End If
                If ds(3) = 11 Then                          'Ante Room
                    For i = 1 To ds!pallets
                        s = "FORKLIFT" & Chr(9)
                        p.area = "FORKLIFT"
                        s = s & ds!groupcode & Chr(9)
                        'p.description = ds!groupcode
                        p.description = gc & Space(8 - Len(gc)) & lname & " " & tno
                        s = s & "ANTE ROOM" & Chr(9)
                        p.source = "ANTE ROOM"
                        If tno = "ZO" Or tno = "OP" Then                                'jv091815
                            s = s & "ORDER PICK" & Chr(9)
                            p.target = "ORDER PICK"
                        Else
                            s = s & "STAGING" & Chr(9)
                            p.target = "STAGING"
                        End If
                        s = s & ds(1) & " " & ds!fgunit & " " & ds!fgdesc & Chr(9)
                        p.product = ds(1) & " " & ds!fgunit & " " & ds!fgdesc
                        s = s & "..." & Chr(9)
                        p.palletid = "..."
                        s = s & "1" & Chr(9)            'Move Qty
                        p.qty = 1
                        s = s & "Pallet" & Chr(9)       'Uom
                        p.uom = "Pallet"
                        s = s & "..." & Chr(9)          'Lot
                        p.lotnum = "..."
                        s = s & ds!pallet & Chr(9)      'Units
                        p.units = ds!pallet
                        s = s & "..." & Chr(9)          'Lot2
                        p.lotnum2 = "..."
                        s = s & " " & Chr(9)            'Units
                        p.units2 = 0
                        s = s & "PEND" & Chr(9)
                        p.status = "PEND"
                        s = s & " " & Chr(9)
                        p.userid = " "
                        s = s & Format(Now, "yymmdd hh:mm:ss") & Chr(9)
                        p.trandate = Format(Now, "yymmdd hh:mm:ss")
                        s = s & " "
                        p.reqid = " "
                        Call insert_trans(p)
                    
                        s = "DOCK" & Chr(9)
                        p.area = "DOCK"
                        s = s & ds!groupcode & Chr(9)
                        'p.description = ds!groupcode
                        p.description = gc                      'jv090911
                        s = s & "ANTE ROOM" & Chr(9)
                        p.source = "ANTE ROOM"
                        s = s & lname & " " & tno & Chr(9)
                        p.target = lname & " " & tno
                        s = s & ds(1) & " " & ds!fgunit & " " & ds!fgdesc & Chr(9)
                        p.product = ds(1) & " " & ds!fgunit & " " & ds!fgdesc
                        s = s & "..." & Chr(9)
                        p.palletid = "..."
                        s = s & "1" & Chr(9)            'Move Qty
                        p.qty = 1
                        s = s & "Pallet" & Chr(9)       'Uom
                        p.uom = "Pallet"
                        s = s & "..." & Chr(9)          'Lot
                        p.lotnum = "..."
                        s = s & ds!pallet & Chr(9)      'Units
                        p.units = ds!pallet
                        s = s & "..." & Chr(9)          'Lot2
                        p.lotnum2 = "..."
                        s = s & " " & Chr(9)            'Units
                        p.units2 = 0
                        s = s & "PEND" & Chr(9)
                        p.status = "PEND"
                        s = s & " " & Chr(9)
                        p.userid = " "
                        s = s & Format(Now, "yymmdd hh:mm:ss") & Chr(9)
                        p.trandate = Format(Now, "yymmdd hh:mm:ss")
                        s = s & " "
                        p.reqid = " "
                        If tno <> "ZO" And tno <> "OP" Then                             'jv091815
                            Call insert_trans(p)
                        End If
                    Next i
                End If
                
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select sku,fgunit,fgdesc,pallet from skumast where sku in ("
    s = s & "select sku from brorders where plant = " & lbr
    s = s & " and branch = " & dbr
    s = s & " and orddate = '" & tdate & "'"
    s = s & " and altflag = 'Y')"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "DOCK" & Chr(9)
            p.area = "DOCK"
            s = s & gc & Chr(9)
            p.description = gc
            s = s & "ALT" & Chr(9)
            p.source = "ALT"
            s = s & lname & " " & tno & Chr(9)
            p.target = lname & " " & tno
            s = s & ds!sku & " " & ds!fgunit & " " & ds!fgdesc & Chr(9)
            p.product = ds!sku & " " & ds!fgunit & " " & ds!fgdesc
            If Form1.plantno = "52" Then        'Assign barcode for sylacauga cranes
                If Len(ds!sku) = 4 Then                                 'jv082415
                    s = s & ds!sku & "000000 X 001"                     'jv082415
                    p.palletid = ds!sku & "000000 X 001"                'jv082415
                Else
                    s = s & ds!sku & " 000000 X 001"
                    p.palletid = ds!sku & " 000000 X 001"
                End If
            Else
                s = s & "..." & Chr(9)
                p.palletid = "..."
            End If
            s = s & "1" & Chr(9)            'Move Qty
            p.qty = 1
            s = s & "Pallet" & Chr(9)       'Uom
            p.uom = "Pallet"
            s = s & "..." & Chr(9)          'Lot
            p.lotnum = "..."
            s = s & ds!pallet & Chr(9)      'Units
            p.units = ds!pallet
            s = s & "..." & Chr(9)          'Lot2
            p.lotnum2 = "..."
            s = s & " " & Chr(9)            'Units
            p.units2 = 0
            s = s & "PEND" & Chr(9)
            p.status = "PEND"
            s = s & " " & Chr(9)
            p.userid = " "
            s = s & Format(Now, "yymmdd hh:mm:ss") & Chr(9)
            p.trandate = Format(Now, "yymmdd hh:mm:ss")
            s = s & " "
            p.reqid = " "
            Call insert_trans(p)
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Private Sub refresh_picks()
    Dim ds As ADODB.Recordset, s As String, i As Integer
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 15
    s = "select * from picktasks " & List1
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!branch & Chr(9)
            s = s & ds!brname & Chr(9)
            s = s & ds!shipdate & Chr(9)
            s = s & ds!palnum & Chr(9)
            s = s & ds!opseq & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & ds!qty & Chr(9)
            s = s & ds!uom & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!palletid & Chr(9)
            s = s & ds!status & Chr(9)
            s = s & ds!userid & Chr(9)
            s = s & ds!location & Chr(9)
            s = s & ds!reqid
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    ycolor.Visible = False
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 11) <> "PEND" Or Grid1.TextMatrix(i, 12) > "." Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ycolor.BackColor
                ycolor.Visible = True
            End If
        Next i
        Grid1.Row = 1: Grid1.Col = 1
    End If
    's = "^ID|^Area|^Source|^Target|<Product|^BarCode|^Qty|^UOM|^Units|^Lot|^Units2|^Lot2|^Status|^User"
    s = "^ID|^Branch|<Name|^Date|^Tag #|^OPSeq|^SKU|^Qty|^UOM|^units|^PalletID|^Status|^User|^Location|^ReqId"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 800
    Grid1.ColWidth(10) = 800
    Grid1.ColWidth(11) = 800
    Grid1.ColWidth(12) = 800
    Grid1.ColWidth(13) = 1200
    Grid1.ColWidth(14) = 800
    Screen.MousePointer = 0
End Sub

Private Sub unmatched_tasks()
    Dim ds As ADODB.Recordset, s As String, i As Integer
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 15
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 3
    s = "select * from paltasks where area = 'DOCK' and description > ' ' and status = 'PEND'"
    s = s & " and userid > ' ' and description not in (select rtrim(left(product, 6))"
    s = s & " from paltasks where area = 'GROUP' and status in ('PEND','ACTV'))"
    s = s & " order by id"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!area & Chr(9)
            s = s & Trim(ds!source) & Chr(9)
            s = s & Trim(ds!target) & Chr(9)
            s = s & ds!product & Chr(9)
            s = s & ds!palletid & Chr(9)
            s = s & ds!qty & Chr(9)
            s = s & ds!uom & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!lotnum & Chr(9)
            s = s & ds!units2 & Chr(9)
            s = s & ds!lotnum2 & Chr(9)
            s = s & ds!status & Chr(9)
            s = s & ds!userid & Chr(9)
            s = s & ds!description
            Grid1.AddItem s
            If ds!area = "DOCK" Then
                If ds!source = "STAGING" Then
                    s = "RACKS" & Chr(9)
                Else
                    s = ds!source & Chr(9)
                End If
                s = s & ds!product & Chr(9)
                pgrid.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    ycolor.Visible = False
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 12) <> "PEND" Or Grid1.TextMatrix(i, 13) > "." Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ycolor.BackColor
                ycolor.Visible = True
            End If
        Next i
        Grid1.Row = 1: Grid1.Col = 1
    End If
    's = "^ID|^Area|^Source|^Target|<Product|^BarCode|^Qty|^UOM|^Units|^Lot|^Units2|^Lot2|^Status|^User"
    s = "^ID|^Area|^Source|<Target|<Product|^BarCode|^Qty|^UOM|^Size||||^Status|^User|^Group"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 3200
    Grid1.ColWidth(4) = 3800
    Grid1.ColWidth(5) = 1800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 1 '800
    Grid1.ColWidth(10) = 1 '800
    Grid1.ColWidth(11) = 1 '800
    Grid1.ColWidth(12) = 800
    Grid1.ColWidth(13) = 1200
    Grid1.ColWidth(14) = 800
    pgrid.FormatString = "^Source|<Product|^_"
    pgrid.ColWidth(0) = 2000
    pgrid.ColWidth(1) = 4000
    pgrid.ColWidth(2) = 2500
    Screen.MousePointer = 0
End Sub
Private Sub refresh_grid1()
    Dim ds As ADODB.Recordset, s As String, i As Integer
    If Combo1 = "Misc Shipping Tasks" Then
        Call unmatched_tasks
        Exit Sub
    End If
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 15
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 3
    If Option1 = True Then
        s = "select * from paltasks where (area = 'DOCK'"
        s = s & " and description = '" & Trim(Mid(Combo1, 6, 6)) & "'"
        's = s & " and target in ('" & Right(Combo1, Len(Combo1) - 13) & "', 'STAGING OP'))"
        s = s & " and target in ('" & Right(Combo1, Len(Combo1) - 13) & "', 'STAGING OP', 'STAGING " & Trim(Mid(Combo1, 6, 6)) & "'))"
        s = s & " or (area = 'FORKLIFT' and description = '" & Right(Combo1, Len(Combo1) - 5) & "')"
        s = s & " order by status desc, userid desc, id"
    End If
    If Option2 = True Then
        s = "select * from paltasks where area = 'PICK'"
        s = s & " and description = '" & Mid(Combo1, 6, Len(Combo1) - 7) & "'"
        s = s & " and target = '" & Right(Combo1, 1) & "'"
        s = s & " order by source,product"
    End If
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!area & Chr(9)
            s = s & Trim(ds!source) & Chr(9)
            s = s & Trim(ds!target) & Chr(9)
            s = s & ds!product & Chr(9)
            s = s & ds!palletid & Chr(9)
            s = s & ds!qty & Chr(9)
            s = s & ds!uom & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!lotnum & Chr(9)
            s = s & ds!units2 & Chr(9)
            s = s & ds!lotnum2 & Chr(9)
            s = s & ds!status & Chr(9)
            s = s & ds!userid & Chr(9)
            s = s & ds!reqid
            Grid1.AddItem s
            If ds!area = "DOCK" Then
                If ds!source = "STAGING" Then
                    s = "RACKS" & Chr(9)
                Else
                    s = ds!source & Chr(9)
                End If
                s = s & ds!product & Chr(9)
                pgrid.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    ycolor.Visible = False
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 12) <> "PEND" Or Grid1.TextMatrix(i, 13) > "." Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ycolor.BackColor
                ycolor.Visible = True
            End If
        Next i
        Grid1.Row = 1: Grid1.Col = 1
    End If
    's = "^ID|^Area|^Source|^Target|<Product|^BarCode|^Qty|^UOM|^Units|^Lot|^Units2|^Lot2|^Status|^User"
    s = "^ID|^Area|^Source|<Target|<Product|^BarCode|^Qty|^UOM|^Size||||^Status|^User|^ReqId"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 3800
    Grid1.ColWidth(5) = 1800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 1 '800
    Grid1.ColWidth(10) = 1 '800
    Grid1.ColWidth(11) = 1 '800
    Grid1.ColWidth(12) = 800
    Grid1.ColWidth(13) = 1200
    Grid1.ColWidth(14) = 800
    pgrid.FormatString = "^Source|<Product|^_"
    pgrid.ColWidth(0) = 2000
    pgrid.ColWidth(1) = 4000
    pgrid.ColWidth(2) = 2500
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_jobbing()
    Dim ds As ADODB.Recordset, s As String
    Dim ss As ADODB.Recordset
    List3.Clear
    s = "select order_num,count(*) from ship_infc"
    s = s & " where ship_status = 'NEW'"
    s = s & " group by order_num order by order_num"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "select groupcode from trailers where groupcode = '" & ds(0) & "'"
            s = s & " and plant = " & Form1.plantno
            If Right(Form1.shipdb, 4) = ".mdb" Then
                s = s & " and pb_flag = false and ra_flag = false"
            Else
                s = s & " and pb_flag = 'N' and ra_flag = 'N'"
            End If
            Set ss = Sdb.Execute(s)
            If ss.BOF = True Then
                List3.AddItem ds(0)
            End If
            ss.Close
            ds.MoveNext
        Loop
    Else
        List3.AddItem "no crane groups...."
    End If
    ds.Close
End Sub

Private Sub refresh_shipping()
    Dim ds As ADODB.Recordset, s As String
    List2.Clear
    s = "select groupcode,count(*) from trailers"
    s = s & " where plant = " & Form1.plantno
    If Form1.plantno = "50" Then
        s = s & " and branch not in (15, 16)"
    End If
    s = s & " and pb_flag = 'N'"
    s = s & " and ra_flag = 'N'"
    s = s & " group by groupcode order by groupcode"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            List2.AddItem ds(0)
            ds.MoveNext
        Loop
    Else
        List2.AddItem "no trailers...."
    End If
    ds.Close
End Sub

Private Sub refresh_trailers()
    Dim ds As ADODB.Recordset, s As String, ts As ADODB.Recordset
    Screen.MousePointer = 11
    Combo1.Clear: List1.Clear
    
    If Option1 = True Then
        s = "select id,area,source,target,product,palletid,status from paltasks"
        s = s & " where area = 'GROUP'"
        s = s & " order by product"
        imptrl.Enabled = True
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = "select count(*) from paltasks where area = 'DOCK'"
                s = s & " and description = '" & Left(ds!product, 6) & "'"
                s = s & " and status = 'PEND'"
                Set ts = Wdb.Execute(s)
                If ts.BOF = False Then
                    ts.MoveFirst
                    If ts(0) > 0 Then
                        s = ds!status & " " & ds!product
                        Combo1.AddItem s
                        List1.AddItem ds!id
                    End If
                End If
                ts.Close
                ds.MoveNext
            Loop
        Else
            Combo1.AddItem "...."
            List1.AddItem 0
        End If
        Combo1.AddItem "Misc Shipping Tasks": List1.AddItem "99999"
    End If
    
    If Option2 = True Then
        s = "select brname,palnum,shipdate,count(*) from picktasks"
        s = s & " where status in ('PEND', 'PICKED')"
        s = s & " group by brname,palnum,shipdate"
        s = s & " order by shipdate,brname,palnum"
        imptrl.Enabled = False
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
    End If
    
    ds.Close
    Screen.MousePointer = 0
    If Combo1.ListCount > 1 Then Combo1.ListIndex = 0
End Sub


Private Sub addalt_Click()
    Dim i As Long, ns As String
    Dim p As ptask, mprod As String, mqty As Integer
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    ns = InputBox("SKU:", "Add Alternate SKU...", ns)
    If Len(ns) = 0 Then Exit Sub
    ns = UCase(ns)
    mprod = "0"
    If skurec(Val(ns)).sku = ns Then
        mprod = ns & " " & skurec(Val(ns)).prodname
        mqty = skurec(Val(ns)).uom_per_pallet
    Else
        MsgBox ns & " is not a valid SKU.."
    End If
    If mprod > "0" Then
        p.area = "DOCK"
        p.description = Trim(Mid(Combo1, 6, 6))
        p.source = "ALT"
        p.target = Right(Combo1, Len(Combo1) - 13)
        p.product = mprod
        p.palletid = "..."
        p.qty = "1"
        p.uom = "Pallet"
        p.lotnum = "..."
        p.units = mqty
        p.lotnum2 = "..."
        p.units2 = 0
        p.status = "PEND"
        p.userid = " "
        p.trandate = Format(Now, "yymmdd hh:mm:ss")
        p.reqid = " "
        Call insert_trans(p)
        DoEvents
        refresh_grid1
    End If
End Sub

Private Sub batonhand_Click()
    Dim s As String
    s = Left(Grid1.TextMatrix(Grid1.Row, 5), 13)
    tktonhand.bbarcode = s
    tktonhand.bproduct = Grid1.TextMatrix(Grid1.Row, 4)
    tktonhand.Show
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    If Option1.Value = True Then
        If List1 > "0" Then refresh_grid1
    End If
    If Option2.Value = True Then
        If Left(List1, 5) = "where" Then refresh_picks
    End If
End Sub

Private Sub cu_Click()
    Dim i As Long, ds As ADODB.Recordset, s As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    If Option1 = True Then
        s = "select userid from paltasks where id = " & i
        s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
    End If
    If Option2 = True Then
        s = "select userid from picktasks where id = " & i
        s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 10) & "'"
    End If
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If Option1 = True Then
            Grid1.TextMatrix(Grid1.Row, 13) = " "
            s = "Update paltasks set userid = ' ' where id = " & i
            Wdb.Execute s
        Else
            s = "Update picktasks set userid = ' ' where id = " & i
            Wdb.Execute s
            Grid1.TextMatrix(Grid1.Row, 12) = " "
        End If
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
        If Option1 = True Then
            If Grid1.TextMatrix(Grid1.Row, 12) = "PEND" Then
                Grid1.CellBackColor = Grid1.BackColor
            Else
                Grid1.CellBackColor = ycolor.BackColor
            End If
        Else
            If Grid1.TextMatrix(Grid1.Row, 11) = "PEND" Then
                Grid1.CellBackColor = Grid1.BackColor
            Else
                Grid1.CellBackColor = ycolor.BackColor
            End If
        End If
        Grid1.Col = 1
    End If
    ds.Close
End Sub

Private Sub edbc_Click()
    Dim i As Long, ds As ADODB.Recordset, s As String, ns As String, uqty As Integer
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 12) <> "PEND" Then
        MsgBox "Task is not Pending.", vbOKOnly + vbExclamation, "Sorry, not this task..."
        Exit Sub
    End If
    If Grid1.TextMatrix(Grid1.Row, 2) = "ALT" Then
        MsgBox "This is an ALTERNATE.", vbOKOnly + vbExclamation, "Sorry, not this task..."
        Exit Sub
    End If
    ns = Grid1.TextMatrix(Grid1.Row, 5)
    'If ns = "..." Then ns = Left(Grid1.TextMatrix(Grid1.Row, 4), 3) & " 000000 X 001"
    If ns = "..." Then ns = Left(Grid1.TextMatrix(Grid1.Row, 4), 4) & "000000 X 001"
    ns = InputBox("New BarCode:", "Change BarCode...", ns)
    If Len(ns) = 0 Then Exit Sub
    If ns = Grid1.TextMatrix(Grid1.Row, 5) Then Exit Sub
    ns = UCase(ns)
    uqty = 0
    s = Trim(Left(ns, 4))
    If skurec(Val(s)).sku = s Then
        uqty = skurec(Val(s)).uom_per_pallet
    Else
        MsgBox ns & " unrecognized SKU..."
    End If
    If uqty > 0 Then
        s = "select palletid from paltasks where id = " & i
        s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            s = "Update paltasks set palletid = '" & ns & "' Where id = " & i
            Wdb.Execute s
            Grid1.TextMatrix(Grid1.Row, 5) = ns
        End If
        ds.Close
    End If
End Sub

Private Sub edloc_Click()
    Dim i As Long, nl As String
    Dim s As String
    If Grid1.Rows < 2 Then Exit Sub
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    nl = InputBox("Location for picked pallet:", "Pallet Storegage...", "A AISLE")
    If Len(nl) = 0 Then Exit Sub
    nl = UCase(nl)
    Screen.MousePointer = 11
    i = Grid1.Row
    Grid1.TextMatrix(i, 13) = nl
    s = "Update picktasks set location = '" & nl & "' Where id = " & Grid1.TextMatrix(i, 0)
    Wdb.Execute s
    Screen.MousePointer = 0
End Sub

Private Sub edpalsize_Click()
    Dim i As Long, ds As ADODB.Recordset, s As String, ns As String
    If Grid1.TextMatrix(Grid1.Row, 12) <> "PEND" Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 13) > "..." Then Exit Sub
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    ns = Grid1.TextMatrix(Grid1.Row, 8)
    ns = InputBox("Pallet Size (units):", "Edit Pallet Size...", ns)
    If Len(ns) = 0 Then Exit Sub
    If Val(ns) <= 0 Then Exit Sub
    ns = UCase(ns)
    s = "select id, units from paltasks where id = " & i
    s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set units = " & Val(ns) & " Where id = " & ds!id
        Wdb.Execute s
        Grid1.TextMatrix(Grid1.Row, 8) = Val(ns)
        Grid1.Col = 1
    End If
    ds.Close
End Sub

Private Sub edsrc_Click()
    Dim i As Long, ds As ADODB.Recordset, s As String, ns As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    ns = Grid1.TextMatrix(Grid1.Row, 2)
    ns = InputBox("Source:", "Edit source...", ns)
    If Len(ns) = 0 Then Exit Sub
    ns = UCase(ns)
    s = "select id, source from paltasks where id = " & i
    s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update paltasks set source = '" & ns & "' Where id = " & ds!id
        Wdb.Execute s
        Grid1.TextMatrix(Grid1.Row, 2) = ns
        Grid1.Col = 1
    End If
    ds.Close
End Sub

Private Sub emplook_Click()
    Dim ds As ADODB.Recordset, s As String
    If Len(Grid1.Text) = 0 Then Exit Sub
    'SQL Database - bbsr
    s = "select * from valuelists where listname = 'wdempid'"
    s = s & " and listreturn = '" & Grid1.Text & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = ds!listdisplay
    Else
        s = "Employee #: " & Grid1.Text & " is not in WdEmp database."
    End If
    ds.Close
    MsgBox s, vbOKOnly + vbInformation, "WMS SQL Employee " & Grid1.Text & " ...."
End Sub

Private Sub Form_Load()
    If Form1.plantno = "50" Then
        impjob.Enabled = True
    Else
        impjob.Enabled = False
    End If
    Option1_Click
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    pgrid.Width = Me.Width - 80
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1500
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Grid1.TextMatrix(0, Grid1.Col) = "User" Then
            PopupMenu usermenu
        Else
            PopupMenu edmenu
        End If
    End If
End Sub

Private Sub Grid1_RowColChange()
    prtlab.Enabled = False
    If Grid1.Row = 0 Then edmenu.Enabled = False
    If Grid1.Rows = 1 Then edmenu.Enabled = False
    If Grid1.TextMatrix(Grid1.Row, 12) = "COMP" Then
        mtc.Enabled = False
        mtp.Enabled = True
    End If
    If Grid1.TextMatrix(Grid1.Row, 12) = "PEND" Then
        mtp.Enabled = False
        mtc.Enabled = True
    End If
    If Grid1.TextMatrix(Grid1.Row, 5) >= "100" Then prtlab.Enabled = True
End Sub

Private Sub impjob_Click()
    refresh_jobbing
    List3.Visible = True
    List3.SetFocus
End Sub

Private Sub imptrl_Click()
    refresh_shipping
    List2.Visible = True
    List2.SetFocus
End Sub

Private Sub List2_Click()
    Screen.MousePointer = 11
    Call postgroup(List2)
    refresh_trailers
    List2.Visible = False
    Screen.MousePointer = 0
End Sub

Private Sub List2_LostFocus()
    List2.Visible = False
End Sub

Private Sub List3_Click()
    Screen.MousePointer = 11
    Call post_jobbing(List3)
    refresh_trailers
    List3.Visible = False
    Screen.MousePointer = 0
End Sub

Private Sub List3_LostFocus()
    List3.Visible = False
End Sub

Private Sub mac_Click()
    Dim i As Long, ds As ADODB.Recordset, s As String, k As Integer
    If Grid1.Rows > 1 Then
        For k = 1 To Grid1.Rows - 1
            i = Val(Grid1.TextMatrix(k, 0))
            If i <> 0 Then
                If Option1 = True Then
                    If Len(Grid1.TextMatrix(k, 5)) = 16 Then                            'jv062915
                        s = "Update pallets Set status = 'Shipped' Where barcode = '"   'jv062915
                        s = s & Grid1.TextMatrix(k, 5) & "'"                            'jv062915
                        Wdb.Execute s                                                    'jv062915
                    End If                                                              'jv062915
                    s = "select id,userid,status from paltasks where id = " & i
                    s = s & " and palletid = '" & Grid1.TextMatrix(k, 5) & "'"
                End If
                If Option2 = True Then
                    s = "select id,userid,status from picktasks where id = " & i
                    s = s & " and palletid = '" & Grid1.TextMatrix(k, 10) & "'"
                End If
                Set ds = Wdb.Execute(s)
                If ds.BOF = False Then
                    ds.MoveFirst
                    If Option1 = True Then
                        s = "Update paltasks set status = 'COMP', userid = ' ' Where id = " & ds!id
                        Wdb.Execute s
                    Else
                        s = "Update picktasks set status = 'COMP', userid = ' ' Where id = " & ds!id
                        Wdb.Execute s
                    End If
                    If Option1 = True Then
                        If Grid1.TextMatrix(k, 12) = "PEND" And Grid1.TextMatrix(k, 2) <> "ALT" Then
                            Grid1.TextMatrix(k, 12) = "COMP"
                            Grid1.TextMatrix(k, 13) = Form1.userid
                            Call post_wms(k)       'jv0113
                        End If
                        Grid1.TextMatrix(k, 12) = "COMP"
                        Grid1.TextMatrix(k, 13) = Form1.userid
                    Else
                        Grid1.TextMatrix(k, 11) = "COMP"
                        Grid1.TextMatrix(k, 12) = " "
                    End If
                    Grid1.Row = k
                    Grid1.RowSel = Grid1.Row
                    Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                    Grid1.CellBackColor = ycolor.BackColor
                    Grid1.Col = 1
                End If
                ds.Close
            End If
        Next k
        If Option1 = True And Val(List1) <> 0 Then
            s = "select area,status,target from paltasks where id = " & List1
            s = s & " and area = 'GROUP'"
            s = s & " and product = '" & Right(Combo1, Len(Combo1) - 5) & "'"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                s = "Update paltasks set area = 'GROUP-COMP', status = 'COMP', target = '...'"
                s = s & " Where id = " & List1
                Wdb.Execute s
            End If
            ds.Close
        End If
    End If
End Sub

Private Sub map_Click()
    Dim i As Long, ds As ADODB.Recordset, s As String, k As Integer
    If Grid1.Rows > 1 Then
        For k = 1 To Grid1.Rows - 1
            i = Val(Grid1.TextMatrix(k, 0))
            If i <> 0 Then
                If Option1 = True Then
                    s = "select id,userid,status from paltasks where id = " & i
                    s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
                End If
                If Option2 = True Then
                    s = "select id,userid,status from picktasks where id = " & i
                    s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 10) & "'"
                End If
                Set ds = Wdb.Execute(s)
                If ds.BOF = False Then
                    ds.MoveFirst
                    If Option1 = True Then
                        s = "Update paltasks set status = 'PEND', userid = ' ' Where id = " & ds!id
                        Wdb.Execute s
                        Grid1.TextMatrix(k, 12) = "PEND"
                        Grid1.TextMatrix(k, 13) = " "
                    Else
                        s = "Update picktasks set status = 'PEND', userid = ' ' Where id = " & ds!id
                        Wdb.Execute s
                        Grid1.TextMatrix(k, 11) = "PEND"
                        Grid1.TextMatrix(k, 12) = " "
                    End If
                    Grid1.Row = k
                    Grid1.RowSel = Grid1.Row
                    Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                    Grid1.CellBackColor = Grid1.BackColor
                    Grid1.Col = 1
                End If
                ds.Close
            End If
        Next k
        If Option1 = True And Val(List1) <> 0 Then
            s = "select status from paltasks where id = " & List1
            s = s & " and area = 'GROUP'"
            s = s & " and product = '" & Right(Combo1, Len(Combo1) - 5) & "'"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                s = "Update paltasks set status = 'PEND' where id = " & List1
                Wdb.Execute s
            End If
            ds.Close
        End If
    End If
End Sub

Private Sub mtc_Click()
    Dim i As Long, ds As ADODB.Recordset, s As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    If Option1 = True Then
        If Len(Grid1.TextMatrix(Grid1.Row, 5)) = 16 Then                            'jv062915
            s = "Update pallets Set status = 'Shipped' Where barcode = '"           'jv062915
            s = s & Grid1.TextMatrix(Grid1.Row, 5) & "'"                            'jv062915
            Wdb.Execute s                                                            'jv062915
        End If                                                                      'jv062915
        s = "select id,userid,status from paltasks where id = " & i
        s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
    End If
    If Option2 = True Then
        s = "select id,userid,status from picktasks where id = " & i
        s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 10) & "'"
    End If
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If Option1 = True Then
            s = "Update paltasks set status = 'COMP', userid = ' ' Where id = " & ds!id
            Wdb.Execute s
            Grid1.TextMatrix(Grid1.Row, 12) = "COMP"
            Grid1.TextMatrix(Grid1.Row, 13) = Form1.userid
            Call post_wms(Grid1.Row)
        Else
            s = "Update picktasks set status = 'COMP', userid = ' ' Where id = " & ds!id
            Wdb.Execute s
            Grid1.TextMatrix(Grid1.Row, 11) = "COMP"
            Grid1.TextMatrix(Grid1.Row, 12) = " "
        End If
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
        Grid1.CellBackColor = ycolor.BackColor
        Grid1.Col = 1
    End If
    ds.Close
End Sub

Private Sub mtp_Click()
    Dim i As Long, ds As ADODB.Recordset, s As String
    i = Val(Grid1.TextMatrix(Grid1.Row, 0))
    If i = 0 Then Exit Sub
    If Option1 = True Then
        s = "select id,userid,status from paltasks where id = " & i
        s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 5) & "'"
    End If
    If Option2 = True Then
        s = "select id,userid,status from picktasks where id = " & i
        s = s & " and palletid = '" & Grid1.TextMatrix(Grid1.Row, 10) & "'"
    End If
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If Option1 = True Then
            s = "Update paltasks set status = 'PEND', userid = ' ' Where id = " & ds!id
            Wdb.Execute s
            Grid1.TextMatrix(Grid1.Row, 12) = "PEND"
            Grid1.TextMatrix(Grid1.Row, 13) = " "
        Else
            s = "Update picktasks set status = 'PEND', userid = ' ' Where id = " & ds!id
            Wdb.Execute s
            Grid1.TextMatrix(Grid1.Row, 11) = "PEND"
            Grid1.TextMatrix(Grid1.Row, 12) = " "
        End If
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
        Grid1.CellBackColor = Grid1.BackColor
        Grid1.Col = 1
    End If
    ds.Close
End Sub

Private Sub Option1_Click()
    edsrc.Enabled = True
    edloc.Enabled = False
    edpalsize.Enabled = True
    refresh_trailers
End Sub

Private Sub Option2_Click()
    edsrc.Enabled = False
    edloc.Enabled = True
    edpalsize.Enabled = False
    refresh_trailers
End Sub

Private Sub palhist_Click()
    palhistory.Show
    palhistory.barkey = Grid1.TextMatrix(Grid1.Row, 5)
End Sub

Private Sub pco_Click()
    Dim rt As String, rf As String, rh As String, tno As String, rfp As String
    Dim bname As String, bno As String, gc As String, sdate As String
    Dim ds As ADODB.Recordset, s As String
    Dim tsb As ADODB.Connection, srun As String
    gc = Trim(Mid(Combo1, 6, 6))
    bname = Right(Combo1, Len(Combo1) - 12)
    bname = Trim(Left(bname, Len(bname) - 3))
    tno = Right(Combo1, 1)
    s = "select branch,shipdate,runid from trailers where groupcode = '" & gc & "'"
    s = s & " and branch in (select branch from branches where branchname = '" & bname & "')"
    s = s & " order by shipdate desc"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        bno = Format(ds(0), "000")
        sdate = Format(ds(1), "MM-dd-yyyy")
        srun = ds(2)
    End If
    ds.Close
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    rfp = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If IsDate(sdate) Then
        If bno = "001" Then bno = "T10"
        If bno = "047" Then bno = "K10"
        If bno = "052" Then bno = "A10"
        Set tsb = CreateObject("ADODB.Connection")                          'jv060216
        tsb.Open Form1.schdb                                                  'jv060216
        s = "select wodate,startime,origin,destination,description,contents,driver"
        's = s & " from truckwo,drivers where truckwo.destination = '" & bno & "'"
        s = s & " from truckwo,drivers where truckwo.r12ticket = '" & srun & "'"
        's = s & " and truckwo.wodate = '" & sdate & "'"
        's = s & " and truckwo.trlno = '" & tno & "'"
        's = s & " and truckwo.wtype in ('Start','SameDay')"
        s = s & " and truckwo.wostatus not in ('CANC','COMP')"
        s = s & " and drivers.id = truckwo.drvid"
        'MsgBox s
        Set ds = tsb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            rf = Format(ds!wodate, "dddd") & " " & Format(ds!wodate, "MM-dd-yy")
            rf = rf & " Start Time: " & Format(ds!startime, "h:mm AM/PM") & "<BR>"
            rf = rf & ds!origin & "-" & ds!destination & " " & ds!description & "<BR>"
            rf = rf & ds!contents & "<BR>"
            rf = rf & ds!driver
            rfp = Format(ds!wodate, "dddd") & " " & Format(ds!wodate, "MM-dd-yy")
            rfp = rfp & " Start Time: " & Format(ds!startime, "h:mm AM/PM") & vbCrLf
            rfp = rfp & ds!origin & "-" & ds!destination & " " & ds!description & vbCrLf
            rfp = rfp & ds!contents & vbCrLf
            rfp = rfp & ds!driver
        End If
        ds.Close: tsb.Close
    End If
    pgrid.RowSel = pgrid.Row
    pgrid.Col = 0: pgrid.ColSel = 1
    pgrid.Sort = 5
    rt = "Pallet Check Off"
    rh = Right(Combo1, Len(Combo1) - 5)
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, pgrid, rt, rh, rfp)
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

Private Sub prtlab_Click()
    Dim u As String, d As String, s As String
    Dim s1 As Integer, s2 As Integer, i As Integer, k As Integer
    k = Grid1.Row
    If Grid1.TextMatrix(k, 5) <= " " Then
        s = "Warning this line does not have a designated OP code.  "
        s = s & "Do you wish to continue printing this line?"
        If MsgBox(s, vbYesNo + vbQuestion, "are you sure...") = vbNo Then Exit Sub
    End If
    
    s = Right(Grid1.TextMatrix(k, 4), Len(Grid1.TextMatrix(k, 4)) - 4)
    u = " "
    d = s
    If UCase(Left(s, 4)) = "BULK" Then
        u = "BULK"
        d = Right(s, Len(s) - 5)
    End If
    If UCase(Left(s, 3)) = "CUP" Then
        u = "CUPS"
        d = Right(s, Len(s) - 4)
    End If
    If UCase(Left(s, 4)) = "3GAL" Then
        u = "3 GALLON"
        d = Right(s, Len(s) - 5)
    End If
    If UCase(Left(s, 3)) = "1/2" Then
        u = "1/2 GAL"
        d = Right(s, Len(s) - 4)
    End If
    If UCase(Left(s, 2)) = "PT" Then
        u = "PINTS"
        d = Right(s, Len(s) - 3)
    End If
    If UCase(Left(s, 2)) = "QT" Then
        u = "QUARTS"
        d = Right(s, Len(s) - 3)
    End If
    If UCase(Left(s, 4)) = "12PK" Then
        u = "12 PACK"
        d = Right(s, Len(s) - 5)
    End If
    If UCase(Left(s, 4)) = "24PK" Then
        u = "24 PACK"
        d = Right(s, Len(s) - 5)
    End If
    If UCase(Left(s, 3)) = "6PK" Then
        u = "6 PACK"
        d = Right(s, Len(s) - 4)
    End If
    If UCase(Left(s, 3)) = "8PK" Then
        u = "8 PACK"
        d = Right(s, Len(s) - 4)
    End If
    If UCase(Left(s, 3)) = "4PK" Then
        u = "4 PACK"
        d = Right(s, Len(s) - 4)
    End If
    If UCase(Left(s, 3)) = "3PK" Then
        u = "3 PACK"
        d = Right(s, Len(s) - 4)
    End If
    If UCase(Left(s, 7)) = "HALF PT" Then
        u = "HALF PINT"
        d = Right(s, Len(s) - 8)
    End If
    
    pallabprt.Show
    If MsgBox("Send to printer?", vbYesNo + vbQuestion, "Ready to print?") = vbYes Then
        pallabprt.prtdevice = "Printer"
        'If Form1.paplegal.Checked = True Then
            Printer.PaperSize = 5
        'Else
        '    Printer.PaperSize = 1
        'End If
    Else
        pallabprt.prtdevice = "Screen"
    End If
    'For i = s1 To s2
        pallabprt.skulab = Trim(Left(Grid1.TextMatrix(k, 4), 4))
        pallabprt.desc1lab = d
        'MsgBox d
        pallabprt.pkglab = u
        'MsgBox u
        'pallabprt.lotlab = Grid1.TextMatrix(k, 4) & " " & Grid1.TextMatrix(k, 5)
        pallabprt.lotlab = Mid(Grid1.TextMatrix(k, 5), 5, 7)
        If Mid(Grid1.TextMatrix(k, 5), 12, 1) = "_" Then
            pallabprt.lotlab = pallabprt.lotlab & "R"
        Else
            pallabprt.lotlab = pallabprt.lotlab & Mid(Grid1.TextMatrix(k, 5), 12, 1)
        End If
        'MsgBox pallabprt.lotlab & " lotlab"
        pallabprt.seqlab = Right(Grid1.TextMatrix(k, 5), 3)
        'MsgBox pallabprt.seqlab & " sequenxe"
        's = Grid1.TextMatrix(k, 1)
        's = s & Grid1.TextMatrix(k, 4)
        's = s & Grid1.TextMatrix(k, 5)
        's = s & i
        pallabprt.ptrig = Grid1.TextMatrix(k, 5)
        'If pallabprt.prtdevice = "Printer" And i <> s2 Then Printer.NewPage
    'Next i
    If pallabprt.prtdevice = "Printer" Then Printer.EndDoc

End Sub

Private Sub restock_Click()
    Dim i As Integer, s As String, bc As String
    Screen.MousePointer = 11
    For i = 1 To Grid1.Rows - 1
        bc = Trim(Grid1.TextMatrix(i, 5))
        If Grid1.TextMatrix(i, 1) = "DOCK" And Len(bc) = 16 Then
            If Right(bc, 1) > "." Then
                s = "Update paltasks set description = 'RE-STOCK'"
                s = s & ", source = 'OC500'"
                s = s & ", target = 'STAGING'"
                s = s & ", status = 'PEND'"
                s = s & ", userid = ''"
                s = s & " Where id = " & Val(Grid1.TextMatrix(i, 0))
                'MsgBox s, vbOKOnly + vbInformation, bc
                Wdb.Execute s
            End If
        End If
    Next i
    Screen.MousePointer = 0
    refresh_grid1
End Sub
