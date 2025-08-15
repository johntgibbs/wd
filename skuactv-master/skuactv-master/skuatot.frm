VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Pallet Totals"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form2"
   ScaleHeight     =   3015
   ScaleWidth      =   5040
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4471
      _Version        =   327680
      Rows            =   8
      Cols            =   5
      BackColorFixed  =   12648447
      FocusRect       =   0
      Appearance      =   0
      FormatString    =   " ^               |^OnHand  |^Orders     |^Available |^Incoming   "
   End
   Begin VB.Label tprod 
      Caption         =   "Label2"
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
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label tsku 
      Caption         =   "Label1"
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
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function r12lot(plot As String) As String
    Dim s As String, t As String
    If Val(plot) > 0 Then
        t = "1-1-20" & Left(plot, 2)
        'MsgBox t
        s = Format(DateAdd("d", Val(Right(plot, 3)) - 1, t), "MM-dd-yyyy")
        'MsgBox s
        s = Format(DateAdd("yyyy", 2, s), "MM-dd-yyyy")
        'MsgBox s
        s = Format(s, "MMddyy")
    End If
    r12lot = s
End Function

Private Sub Form_Load()
    Dim i As Integer
    Form2.Caption = Form2.Caption & " " & Form1.plantdesc
    'Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Cols = 5
    If Form1.plantno = "50" Then Grid1.Rows = 9
    If Form1.plantno = "51" Then Grid1.Rows = 5
    If Form1.plantno = "52" Then Grid1.Rows = 6
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form2" Then
            Form2.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form2.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form2.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form2.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Grid1.FormatString = "^|^OnHand|^Orders|^Available|^Incoming"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    If Form1.plantno = "50" Then
        Grid1.TextMatrix(1, 0) = "SR-1"
        Grid1.TextMatrix(2, 0) = "SR-2"
        Grid1.TextMatrix(3, 0) = "SR-3"
        Grid1.TextMatrix(4, 0) = "Racks"
        Grid1.TextMatrix(5, 0) = "SR-5"
        Grid1.TextMatrix(7, 0) = "BB Total"
        Grid1.TextMatrix(8, 0) = "4Ways"
    End If
    If Form1.plantno = "51" Then
        Grid1.TextMatrix(1, 0) = "Racks"
        Grid1.TextMatrix(3, 0) = "BB Total"
        Grid1.TextMatrix(4, 0) = "4Ways"
    End If
    If Form1.plantno = "52" Then
        Grid1.TextMatrix(1, 0) = "CS5"
        Grid1.TextMatrix(2, 0) = "Racks"
        Grid1.TextMatrix(4, 0) = "BB Total"
        Grid1.TextMatrix(5, 0) = "4Ways"
    End If
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 120
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (tsku.Height * 2)
End Sub

Private Sub Form_Terminate()
    Dim i As Integer
    If Form2.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form2" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form2.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form2.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form2.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form2.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Terminate
End Sub

Private Sub Grid1_Click()
    If Form1.plantno = "50" Then
        If Grid1.Row < 5 Then
            Form3.lwhs = Grid1.Row
            Form4.bwhs = Grid1.Row
        Else
            Form3.lwhs = "0"
            Form4.bwhs = "0"
        End If
    End If
    If Form1.plantno = "51" Then
        Form3.lwhs = "4"
        Form4.bwhs = "4"
    End If
    If Form1.plantno = "52" Then
        If Grid1.Row = 1 Then
            Form3.lwhs = "5"
            Form4.bwhs = "5"
        Else
            If Grid1.Row = 2 Then
                Form3.lwhs = "4"
                Form4.bwhs = "4"
            Else
                Form3.lwhs = "0"
                Form4.bwhs = "0"
            End If
        End If
    End If
    Form4.blot = "0"
End Sub

Private Sub tsku_Change()
    Dim ds As ADODB.Recordset, s As String, ps As ADODB.Recordset
    Dim db5 As Database, ds5 As ADODB.Recordset, bbpallet As Integer
    Dim i As Integer, k As Integer, r4 As Integer, rt As Integer, rr As Integer
    
    ' Clear grid 1 cells
    For i = 1 To Grid1.Rows - 1
        For k = 1 To Grid1.Cols - 1
            Grid1.TextMatrix(i, k) = ""
        Next k
    Next i
    
    If Form1.plantno = "50" Then
        s = "select whse_num,gmasize,sum(qty) from lane where sku = '" & tsku & "'"
        s = s & " group by whse_num,gmasize"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ds!gmasize > 0 Then                  'jv082813
                    Grid1.TextMatrix(8, 1) = Format(Val(Grid1.TextMatrix(8, 1)) + ds(2), "#####")
                Else
                    Grid1.TextMatrix(ds!whse_num, 1) = Format(ds(2), "#####")
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
        s = "select to_whse_num,sum(order_qty-ship_plt_qty) from ship_infc"
        s = s & " where sku = '" & tsku & "'"
        s = s & " and ship_status not in ('CANC','DONE')"
        s = s & " group by to_whse_num"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                Grid1.TextMatrix(ds!to_whse_num, 2) = Format(ds(1), "#####")
                ds.MoveNext
            Loop
        End If
        ds.Close
        
        s = "select source,count(*) from paltasks"
        s = s & " where area = 'DOCK' and status = 'PEND'"
        s = s & " and source in ('SR5', 'SR6')"             'jv082813
        s = s & " and description > ' '"
        s = s & " and product >= '" & tsku & "'"
        s = s & " and product < '" & tsku & "ZZZ'"
        s = s & " and lotnum < '0'"
        s = s & " group by source"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ds!Source = "SR5" Then                               'jv082813
                    Grid1.TextMatrix(5, 2) = Format(ds(1), "#####")
                Else
                    Grid1.TextMatrix(8, 2) = Format(ds(1), "#####")
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
        
        s = "select * from prodrcv where sku = '" & tsku & "'"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                Grid1.TextMatrix(1, 4) = Format(Val(Grid1.TextMatrix(1, 4)) + ds!sr1, "###")
                Grid1.TextMatrix(2, 4) = Format(Val(Grid1.TextMatrix(2, 4)) + ds!sr2, "###")
                Grid1.TextMatrix(3, 4) = Format(Val(Grid1.TextMatrix(3, 4)) + ds!sr3, "###")
                Grid1.TextMatrix(4, 4) = Format(Val(Grid1.TextMatrix(4, 4)) + ds!sr4, "###")
                Grid1.TextMatrix(5, 4) = Format(Val(Grid1.TextMatrix(5, 4)) + ds!sr5, "###")
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    If Form1.plantno = "52" Then
        ' Get pallet size for selected SKU
        s = "select uom_per_pallet from sku_config where sku = '" & tsku & "'"          'jv070915
        Set ds = Form1.wdb.Execute(s)                                                    'jv070915
        If ds.BOF = False Then                                                          'jv070915
            ds.MoveFirst                                                                'jv070915
            bbpallet = ds(0)                                                            'jv070915
        Else                                                                            'jv070915
            bbpallet = 1                                                                'jv070915
        End If                                                                          'jv070915
        ds.Close                                                                        'jv070915
        
        ' Now fetch all records from Westfalia of this SKU
        s = "SELECT [Default Quantity], ISNULL([Pallet Count], 0) FROM vAllItems_1033 WHERE Item >= '" & tsku & "' and Item < '" & tsku & "ZZZZZ'"
        Set ds5 = Form1.db5.Execute(s)
        If ds5.BOF = False Then
            ds5.MoveFirst
            Do Until ds5.EOF
                If ds5(0) = bbpallet Then                                               'jv070915
                    ' If same pallet size as in SKU Config, add to OnHand column
                    Grid1.TextMatrix(1, 1) = Format(Val(Grid1.TextMatrix(1, 1)) + ds5(1), "#####")
                Else
                    ' Else add to Incoming column
                    Grid1.TextMatrix(5, 1) = Format(ds5(1) + Val(Grid1.TextMatrix(5, 1)), "###")
                End If
                ds5.MoveNext
            Loop
        End If
        ds5.Close
        
        s = "select area,count(*) from paltasks where area = 'DOCK'"
        s = s & " and status = 'PEND' and source not in ('STAGING','ALT')"
        s = s & " and product >= '" & tsku & "'"
        s = s & " and product < '" & tsku & "ZZZ'"
        s = s & " and lotnum < '0'"
        s = s & " group by area"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            'Grid1.TextMatrix(4, 2) = Format(ds(1), "#####")
            Grid1.TextMatrix(1, 2) = Format(ds(1), "#####")
        End If
        ds.Close
        'MsgBox "dock tasks"
    End If
    If Form1.Check2.Value = 1 Then
        If Form1.plantno = "50" Then
            rr = 4: rt = 7: r4 = 8
        End If
        If Form1.plantno = "51" Then
            rr = 1: rt = 3: r4 = 4
        End If
        If Form1.plantno = "52" Then
            rr = 2: rt = 4: r4 = 5
        End If
        s = "select bbc,count(*) from rackpos where sku = '" & tsku & "'"
        s = s & " group by bbc"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ds(0) = "Y" Then
                    Grid1.TextMatrix(rr, 1) = Format(ds(1), "#####")
                Else
                    Grid1.TextMatrix(r4, 1) = Format(ds(1) + Val(Grid1.TextMatrix(r4, 1)), "#####")
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
        s = "select area,count(*) from paltasks where area = 'FORKLIFT'"
        s = s & " and status = 'PEND' and target = 'STAGING'"
        s = s & " and product >= '" & tsku & "'"
        s = s & " and product < '" & tsku & "ZZZ'"
        s = s & " group by area"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Grid1.TextMatrix(rr, 2) = Format(ds(1), "#####")
        End If
        ds.Close
        'MsgBox "efl tasks"
        Grid1.TextMatrix(rt - 1, 0) = "Hold"
        s = "select * from holdlist where sku = '" & tsku & "' order by lot_num, opcode"
        Set ds = Form1.wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If Len(ds!opcode) = 1 Then
                    If Len(tsku) = 3 Then                                                       'jv082415
                        sbc = tsku & " " & r12lot(ds!lot_num) & " " & ds!opcode & " " & ds!spallet
                        ebc = tsku & " " & r12lot(ds!lot_num) & " " & ds!opcode & " " & ds!epallet
                    Else
                        sbc = tsku & r12lot(ds!lot_num) & " " & ds!opcode & " " & ds!spallet    'jv082415
                        ebc = tsku & r12lot(ds!lot_num) & " " & ds!opcode & " " & ds!epallet    'jv082415
                    End If
                Else
                    If Len(tsku) = 3 Then                                                       'jv082415
                        sbc = tsku & " " & r12lot(ds!lot_num) & ds!opcode & ds!spallet
                        ebc = tsku & " " & r12lot(ds!lot_num) & ds!opcode & ds!epallet
                    Else
                        sbc = tsku & r12lot(ds!lot_num) & ds!opcode & ds!spallet                'jv082415
                        ebc = tsku & r12lot(ds!lot_num) & ds!opcode & ds!epallet                'jv082415
                    End If
                End If
                'MsgBox sbc & " ... " & ebc
                s = "select sku, count(*) from pallets where sku = '" & tsku & "'"
                s = s & " and barcode >= '" & sbc & "'"
                s = s & " and barcode <= '" & ebc & "'"
                s = s & " and status in ('Wrapper', 'Warehouse')"
                s = s & " group by sku"
                Set ps = Form1.wdb.Execute(s)
                If ps.BOF = False Then
                    ps.MoveFirst
                    Do Until ps.EOF
                        Grid1.TextMatrix(rt - 1, 1) = Val(Grid1.TextMatrix(rt - 1, 1)) - ps(1)
                        'MsgBox sbc & " ... " & ebc & "  = " & ps(1)
                        ps.MoveNext
                    Loop
                End If
                ps.Close
                ds.MoveNext
            Loop
        End If
    End If
    'db.Close
    For i = 1 To rt - 1 'rr
        Grid1.TextMatrix(i, 3) = Format(Val(Grid1.TextMatrix(i, 1)) - Val(Grid1.TextMatrix(i, 2)), "######")
        For k = 1 To 4
            Grid1.TextMatrix(rt, k) = Format(Val(Grid1.TextMatrix(rt, k)) + Val(Grid1.TextMatrix(i, k)), "######")
        Next k
    Next i
    Grid1.TextMatrix(r4, 3) = Format(Val(Grid1.TextMatrix(r4, 1)) - Val(Grid1.TextMatrix(r4, 2)), "######")
End Sub

