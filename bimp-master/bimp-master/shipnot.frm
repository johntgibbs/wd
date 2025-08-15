VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form shipnot 
   Caption         =   "Shipments Not Received"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   14565
   StartUpPosition =   3  'Windows Default
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
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4815
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8493
      _Version        =   327680
      ForeColor       =   12582912
      BackColorFixed  =   12648447
      ForeColorFixed  =   12582912
      BackColorSel    =   16384
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5530
      _Version        =   327680
      BackColorFixed  =   12648384
      BackColorSel    =   192
      FocusRect       =   0
   End
   Begin VB.Label tktkey 
      Caption         =   "tktkey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "shipnot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()
    Dim q As String, i As Integer
    Dim ds As ADODB.Recordset, s As String
    
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
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 2
    
    
    
    s = "Select t.shipment_number, sum(t.transaction_quantity)"
    s = s & " From mtl_material_transactions t"
    s = s & " Where t.transaction_date > sysdate - 5"
    s = s & " and t.shipment_number > ' '"
    s = s & " and t.source_code in ('RCV', 'TRAILER TRANSFER')"
    s = s & " group by t.shipment_number"
    s = s & " Having Sum(t.transaction_quantity) < 0"
    s = s & " order by t.shipment_number"
    
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(0) & Chr(9) & ds(1)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    
    ds.Close
    Grid1.FormatString = "^Ticket|^Unit Qty"
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 1200
    Grid1.Redraw = True
    Screen.MousePointer = 0

End Sub

Private Sub refresh_grid2()
    Dim ds As ADODB.Recordset, s As String, i As Integer, r12 As Boolean
    Dim pplant As String, pb As ADODB.Connection
    

    
    Screen.MousePointer = 11
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 8
    
    pplant = "T10"
    s = "select t.subinventory_code from mtl_material_transactions t"
    s = s & " where t.shipment_number = '" & tktkey.Caption & "'"
    s = s & " and t.source_code = 'TRAILER TRANSFER'"
    Set ds = r12db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        pplant = ds(0)
    End If
    ds.Close
    
    'If plant_server_status(pplant) = True Then                      'jv010417
    '    If pplant <> "T10" And pplant <> "001" Then
    '        Set pb = CreateObject("ADODB.Connection")
    '        If pplant = "K10" Or pplant = "047" Then pb.Open k10ship
    '        If pplant = "A10" Or pplant = "052" Then pb.Open a10ship
    '    End If
    'End If
    
    r12 = False
    s = "select groupcode, plant, branch, shipdate, trlno, sku, sum(units)"
    s = s & " from trailers where runid = " & Left(tktkey, Len(tktkey) - 1)
    s = s & " group by groupcode, plant, branch, shipdate, trlno, sku"
    
    If plant_server_status(pplant) = True Then                      'jv010417
        If pplant = "T10" Or pplant = "001" Then
            Set ds = wdb.Execute(s)
        Else
            Set pb = CreateObject("ADODB.Connection")
            If pplant = "K10" Or pplant = "047" Then pb.Open k10ship
            If pplant = "A10" Or pplant = "052" Then pb.Open a10ship
            Set ds = pb.Execute(s)
        End If
    
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds!groupcode & Chr(9)
                s = s & plantrec(ds!plant).orawhs & Chr(9)
                s = s & branchrec(ds!branch).oraloc & Chr(9)
                s = s & branchrec(ds!branch).branchname & " " & ds!trlno & Chr(9)
                s = s & Format(ds!shipdate, "M-dd-yyyy") & Chr(9)
                s = s & ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                s = s & ds(6)
                Grid2.AddItem s
                ds.MoveNext
            Loop
        Else
            r12 = True
        End If
        ds.Close
    
        If pplant <> "T10" And pplant <> "001" Then pb.Close
    Else                                                                'jv010417
        s = "WD Server for " & pplant & " is not on line.  Non-receipts cannot be processed."
        MsgBox s, vbOKOnly + vbExclamation, "Server cannot be accessed."
    End If                                                              'jv010417
    
    
    If r12 = True Then
        Call r12_batch(Left(tktkey, Len(tktkey) - 1))
        Exit Sub
    End If
    
    Grid2.FormatString = "^Group|^Plant|^Branch|<Trailer|^Date|^SKU|<Product|^Units"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 2000
    Grid2.ColWidth(4) = 1200
    Grid2.ColWidth(5) = 800
    Grid2.ColWidth(6) = 3200
    Grid2.ColWidth(7) = 1000
    Grid2.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub r12_batch(rid As String)
    Dim q As String, i As Integer, k As Integer
    Dim ds As ADODB.Recordset
    Screen.MousePointer = 11
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 8
    
    
    q = "select t.subinventory_code, i.segment1, i.description, t.transaction_quantity," & _
        " t.transaction_uom, t.source_code, t.transaction_reference, t.transaction_date" & _
        " from mtl_material_transactions t, mtl_system_items_b i" & _
        " where t.shipment_number in ('" & rid & "P', '" & rid & "W')" & _
        " and i.inventory_item_id = t.inventory_item_id" & _
        " and i.organization_id = t.organization_id" & _
        " order by t.source_code, i.segment1, t.subinventory_code"
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(0) & Chr(9)
            s = s & ds(1) & Chr(9)
            s = s & ds(2) & Chr(9)
            s = s & ds(3) & Chr(9)
            s = s & ds(4) & Chr(9)
            s = s & ds(5) & Chr(9)
            s = s & ds(6) & Chr(9)
            s = s & ds(7)
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
        
    Grid2.FormatString = "^SubInv|^SKU|<Product|^Qty|^UOM|^SourceCode|<Reference|^Date/Time"
    Grid2.ColWidth(0) = 800
    Grid2.ColWidth(1) = 800
    Grid2.ColWidth(2) = 2000
    Grid2.ColWidth(3) = 800
    Grid2.ColWidth(4) = 600
    Grid2.ColWidth(5) = 1800
    Grid2.ColWidth(6) = 2600
    Grid2.ColWidth(7) = 1800
    Grid2.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
    refresh_grid1
    DoEvents
    If Grid1.Rows > 1 Then tktkey.Caption = Grid1.TextMatrix(1, 0)
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    'Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    refresh_grid1
    DoEvents
    If Grid1.Rows > 1 Then tktkey.Caption = Grid1.TextMatrix(1, 0)
End Sub

Private Sub Form_Resize()
    If Me.Height > 2000 Then
        Grid1.Height = Me.Height - 1000
        Grid2.Height = Grid1.Height
    End If
    Grid2.Width = Me.Width - Grid1.Width
End Sub

Private Sub Grid1_RowColChange()
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) > 0 Then tktkey.Caption = Grid1.TextMatrix(Grid1.Row, 0)
End Sub

Private Sub Grid2_DblClick()
    Dim blit As String, i As Integer, k As Integer
    i = Grid2.Row
    If Grid2.TextMatrix(0, 6) = "Reference" Then
        blit = Left(Grid2.TextMatrix(i, 6), Len(Grid2.TextMatrix(i, 6)) - 7)
        For k = 1 To 99
            If UCase(branchrec(k).branchname) = UCase(blit) Then
                'MsgBox blit & "=" & i
                s = "Update bimp set onorder = onorder + " & Val(Grid2.TextMatrix(i, 3) * -1)
                s = s & " where plantwhs = '" & Grid2.TextMatrix(i, 0) & "'"
                s = s & " and branchwhs = '" & Format(k, "000") & "'"
                s = s & " and sku = '" & Grid2.TextMatrix(i, 1) & "'"
                MsgBox s
                Exit For
            End If
        Next k
    End If
End Sub

Private Sub tktkey_Change()
    refresh_grid2
End Sub

