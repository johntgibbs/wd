VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form shipinfc 
   Caption         =   "Shipping Groups"
   ClientHeight    =   6300
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   ScaleHeight     =   6300
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option4 
      Caption         =   "Racks"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   0
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Caption         =   "SR-3"
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
      Left            =   2760
      TabIndex        =   9
      Top             =   0
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "SR-2"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   0
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "SR-1"
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
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel Group"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cancel SR"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add SKU"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Group"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2295
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4048
      _Version        =   327680
      BackColorFixed  =   12648447
      FocusRect       =   0
      HighLight       =   2
      Appearance      =   0
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   960
      TabIndex        =   1
      Top             =   3480
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   2535
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu cangrp 
         Caption         =   "Cancel Group"
      End
      Begin VB.Menu addgrp 
         Caption         =   "Add Group"
      End
      Begin VB.Menu addsku 
         Caption         =   "Add SKU"
      End
      Begin VB.Menu cansr 
         Caption         =   "Cancel SR"
      End
      Begin VB.Menu ed4 
         Caption         =   "Edit 4Way Size"
      End
   End
End
Attribute VB_Name = "shipinfc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_groups()
    Dim ds As adodb.Recordset, sqlx As String
    Combo1.Clear
    sqlx = "select distinct order_num from ship_infc"
    sqlx = sqlx & " where ship_status = 'NEW' or ship_status = 'ACTV'"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds!order_num
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Private Sub refresh_ship()
    Dim ds As adodb.Recordset, sqlx As String, pwhse As Integer
    If Option1.Value = True Then pwhse = 1
    If Option2.Value = True Then pwhse = 2
    If Option3.Value = True Then pwhse = 3
    If Option4.Value = True Then pwhse = 4
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 9: Grid2.FixedCols = 4
    sqlx = "select id,to_whse_num,sku,order_qty,ship_plt_qty,ship_status,gmasize"
    sqlx = sqlx & " from ship_infc where order_num = '" & Combo1 & "'"
    sqlx = sqlx & " and to_whse_num = " & pwhse
    sqlx = sqlx & " and ship_status not in ('CANC','DONE')"
    sqlx = sqlx & " order by sku"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!id & Chr$(9)
            sqlx = sqlx & ds!to_whse_num & Chr$(9)
            sqlx = sqlx & ds!sku & Chr(9)
            sqlx = sqlx & skurec(Val(ds!sku)).prodname & Chr(9)
            sqlx = sqlx & ds!order_qty & Chr$(9)
            sqlx = sqlx & ds!ship_plt_qty & Chr$(9)
            sqlx = sqlx & ds!order_qty - ds!ship_plt_qty & Chr$(9)
            sqlx = sqlx & ds!ship_status
            sqlx = sqlx & Chr(9) & ds!gmasize
            Grid2.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid2.FormatString = "ID|^SR|^SKU|<Product|^Ordered|^Shipped|^Net|^Status|^4Way Size"
    Grid2.ColWidth(0) = 1: Grid2.ColWidth(1) = 600
    Grid2.ColWidth(2) = 800: Grid2.ColWidth(3) = 3600
    Grid2.ColWidth(4) = 1000: Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 1000: Grid2.ColWidth(7) = 1000
    Grid2.ColWidth(8) = 1100
    Grid2.Redraw = True
End Sub

Private Sub addgrp_Click()
    Command1_Click
End Sub

Private Sub addsku_Click()
    Command2_Click
End Sub

Private Sub cangrp_Click()
    Command4_Click
End Sub

Private Sub cansr_Click()
    Command6_Click
End Sub

Private Sub Combo1_Click()
    Call refresh_ship
End Sub

Private Sub Command1_Click()            'New Group
    Dim pgrp As String
    pgrp = InputBox("Group:", "New Group")
    If Len(pgrp) = 0 Then Exit Sub
    Combo1.AddItem pgrp
    Combo1.ListIndex = Combo1.ListCount - 1
End Sub

Private Sub Command2_Click()            'Add SKU
    Dim psku As String, pwhs As Integer
    Dim ds As adodb.Recordset, sqlx As String, s As String, rid As Long
    Dim pvert As Integer, phorz As Integer, pside As String
    Dim pdesc As String
    psku = InputBox("SKU:", "Add SKU", "777")
    If Len(psku) = 0 Then Exit Sub
    If Option1 = True Then pwhs = 1
    If Option2 = True Then pwhs = 2
    If Option3 = True Then pwhs = 3
    If Option4 = True Then pwhs = 4
    If skurec(Val(psku)).sku = psku Then
        pdesc = skurec(Val(psku)).prodname
    Else
        MsgBox "SKU not on file...", vbOKOnly, "Sorry cannot add"
        Exit Sub
    End If
    sqlx = "select * from sr_config where whs_num = " & pwhs
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        pvert = ds!ship1_lane_vert
        phorz = ds!ship1_lane_horz
        pside = ds!ship1_lane_side
    Else
        pvert = "1"
        phorz = "1"
        pside = "R"
    End If
    ds.Close
    sqlx = "select * from ship_infc where ship_status = 'CANC' or ship_status = 'DONE'"
    sqlx = sqlx & " order by id"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update ship_infc set order_num = '" & Combo1 & "'"
        s = s & ", sku = '" & psku & "', lot_num = ' ', ship_date = '" & Format(Now, "M-d-yyyy") & "'"
        s = s & ", order_qty = 0, ship_uom_qty = 0, ship_plt_qty = 0, ship_status = 'NEW'"
        s = s & ", to_whse_num = " & Val(pwhs)
        s = s & ", to_vert_loc = " & pvert
        s = s & ", to_horz_loc = " & phorz
        s = s & ", to_rack_side = '" & pside & "', resv_strategy = 'A'"
        s = s & " Where id = " & ds!id
        Wdb.Execute s
        sqlx = ds!id & Chr(9)
        sqlx = sqlx & pwhs & Chr(9)
        sqlx = sqlx & psku & Chr(9)
        sqlx = sqlx & " " & pdesc & Chr(9)
        sqlx = sqlx & "0" & Chr(9)
        sqlx = sqlx & "0" & Chr(9)
        sqlx = sqlx & "0" & Chr(9)
        sqlx = sqlx & "NEW"
    Else
        rid = wd_seq("Ship_Infc")
        s = "INSERT INTO Ship_Infc (ID, Order_Num, SKU, Lot_Num, Ship_Date, Order_Qty,"
        s = s & " Ship_Uom_Qty, Ship_Plt_Qty, Ship_Status, To_Whse_Num, To_Vert_Loc,"
        s = s & " To_Rack_Side, Resv_Strategy, GMASize)"
        s = s & " VALUES (" & rid & ","
        s = s & "'" & Combo1 & "',"
        s = s & "'" & psku & "',"
        s = s & "'.',"
        s = s & "'" & Format(Now, "mm-dd-yyyy") & "',"
        s = s & "0,0,0,'NEW',"
        s = s & Val(pwhs) & ","
        s = s & pvert & ","
        s = s & phorz & ","
        s = s & "'" & pside & "',"
        s = s & "'A',0)"
        Wdb.Execute s
        sqlx = rid & Chr(9)
        sqlx = sqlx & pwhs & Chr(9)
        sqlx = sqlx & psku & Chr(9)
        sqlx = sqlx & " " & pdesc & Chr(9)
        sqlx = sqlx & "0" & Chr(9)
        sqlx = sqlx & "0" & Chr(9)
        sqlx = sqlx & "0" & Chr(9)
        sqlx = sqlx & "NEW"
    End If
    ds.Close
    Grid2.AddItem sqlx
    Grid2.Row = Grid2.Rows - 1
End Sub

Private Sub Command4_Click()        'Cancel Group
    Dim sqlx As String
    If Combo1.ListIndex < 0 Then Exit Sub
    If MsgBox("Cancel Group: " & Combo1, vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    sqlx = "update ship_infc set ship_status = 'CANC' where order_num = '" & Combo1 & "'"
    MsgBox sqlx
    Wdb.Execute sqlx
    Call refresh_ship
End Sub

Private Sub Command6_Click()        'Cancel / Actv SKU
    Dim sqlx As String
    Dim pwhs As String
    Dim pstat As String, plit As String
    If Grid2.Row < 1 Then Exit Sub
    If Grid2.TextMatrix(Grid2.Row, 7) = "DONE" Then Exit Sub
    If Grid2.TextMatrix(Grid2.Row, 7) = "CANC" Then
        If Val(Grid2.TextMatrix(Grid2.Row, 5)) > 0 Then
            pstat = "ACTV": plit = "Activate "
        Else
            pstat = "NEW": plit = "Insert "
        End If
    Else
        pstat = "CANC": plit = "Cancel "
    End If
    pwhs = Grid2.TextMatrix(Grid2.Row, 1)
    If MsgBox(plit & Grid2.TextMatrix(Grid2.Row, 3) & " from SR " & pwhs, vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    sqlx = "update ship_infc set ship_status = '" & pstat & "' where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    Wdb.Execute sqlx
    Grid2.TextMatrix(Grid2.Row, 7) = pstat
End Sub

Private Sub ed4_Click()
    Dim db As Database, ds As Recordset, s As String, psz As String
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    psz = Grid2.TextMatrix(Grid2.Row, Grid2.Cols - 1)
    psz = InputBox("4Way Pallet Size (Wraps):", "4Way Wraps Per Pallet..", psz)
    If Len(psz) = 0 Then Exit Sub
    Grid2.TextMatrix(Grid2.Row, Grid2.Cols - 1) = Val(psz)
    s = "Update ship_infc set gmasize = " & Val(psz) & " Where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    Wdb.Execute s
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If shipinfc.WindowState = 0 Then
        For i = 1 To Form1.Frmgrid.Rows - 1
            If Form1.Frmgrid.TextMatrix(i, 0) = "shipinfc" Then
                Form1.Frmgrid.TextMatrix(i, 1) = shipinfc.Top
                Form1.Frmgrid.TextMatrix(i, 2) = shipinfc.Left
                Form1.Frmgrid.TextMatrix(i, 3) = shipinfc.Height
                Form1.Frmgrid.TextMatrix(i, 4) = shipinfc.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If shipinfc.ActiveControl.Name = "Combo1" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command1_Click
        If KeyCode = 46 Or KeyCode = 120 Then Call Command4_Click
    End If
    If shipinfc.ActiveControl.Name = "Grid2" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command2_Click
        If KeyCode = 46 Or KeyCode = 120 Then Call Command6_Click
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.Frmgrid.Rows - 1
        If Form1.Frmgrid.TextMatrix(i, 0) = "shipinfc" Then
            shipinfc.Top = Val(Form1.Frmgrid.TextMatrix(i, 1))
            shipinfc.Left = Val(Form1.Frmgrid.TextMatrix(i, 2))
            shipinfc.Height = Val(Form1.Frmgrid.TextMatrix(i, 3))
            shipinfc.Width = Val(Form1.Frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    ed4.Enabled = False
    Call refresh_groups
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid2.Height = shipinfc.Height - (Combo1.Height * 3.2) ' 800
    Grid2.Width = Me.Width - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid2_KeyPress(KeyAscii As Integer)
    Dim mflag As Integer, pqty As Integer, pqty2 As Integer
    Dim sqlx As String, mstat As String
    If Grid2.Row = 0 Then Exit Sub
    Grid2.Col = 4: mflag = 0
    If KeyAscii = 8 Then
        If Len(Grid2.Text) > 1 Then
            Grid2.Text = Left$(Grid2.Text, Len(Grid2.Text) - 1)
        Else
            Grid2.Text = "0"
        End If
        mflag = 1
    End If
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Grid2.Text = Val(Grid2.Text + Chr$(KeyAscii))
        If Val(Grid2.Text) > 65355 Then
            Grid2.Text = Left$(Grid2.Text, 4)
            Beep
        End If
        mflag = 1
    End If
    If mflag = 1 Then
        pqty = Val(Grid2.Text)
        pqty2 = Val(Grid2.TextMatrix(Grid2.Row, 5))
        Grid2.TextMatrix(Grid2.Row, 6) = pqty - pqty2
        If pqty2 > 0 Then
            mstat = "ACTV"
        Else
            mstat = "NEW"
        End If
        If pqty2 >= pqty Then mstat = "DONE"
        Grid2.TextMatrix(Grid2.Row, 7) = mstat
        sqlx = "update ship_infc set order_qty = " & pqty
        sqlx = sqlx & ", ship_status = '" & mstat & "' where id = " & Grid2.TextMatrix(Grid2.Row, 0)
        Wdb.Execute sqlx
    End If
End Sub

Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Option1_Click()
    Call refresh_ship
    ed4.Enabled = False
End Sub

Private Sub Option2_Click()
    Call refresh_ship
    ed4.Enabled = False
End Sub

Private Sub Option3_Click()
    Call refresh_ship
    ed4.Enabled = False
End Sub

Private Sub Option4_Click()
    Call refresh_ship
    ed4.Enabled = True
End Sub
