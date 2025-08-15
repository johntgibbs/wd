VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Move Products in Rack"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form6"
   ScaleHeight     =   3225
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "F2:Edit Rack"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "F9:Clear Rack"
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "F3:Split Rack"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Targets 
      Height          =   1455
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2566
      _Version        =   327680
      BackColor       =   16777152
      FocusRect       =   2
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Source "
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton Command1 
         Caption         =   "Move Product"
         Height          =   255
         Left            =   6240
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label RKey 
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label pqty4 
         Caption         =   "pqty4"
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
         Left            =   4800
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
      Begin VB.Label pqty 
         Caption         =   "pqty"
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
         Left            =   2880
         TabIndex        =   17
         Top             =   480
         Width           =   735
      End
      Begin VB.Label plot 
         Caption         =   "plot"
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
         Left            =   840
         TabIndex        =   16
         Top             =   480
         Width           =   615
      End
      Begin VB.Label pdesc 
         Caption         =   "pdesc"
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
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label psku 
         Caption         =   "psku"
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
         Left            =   2880
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.Label paisle 
         Caption         =   "paisle"
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
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "4Way:"
         Height          =   255
         Left            =   4200
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "BB:"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Lot:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "SKU:"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Rack:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Target Aisle:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_targets()
    Dim ds As ADODB.Recordset, sqlx As String
    Dim i As Integer, psku As String
    Screen.MousePointer = 11
    Targets.Visible = False: Targets.Clear
    Targets.Rows = 1: Targets.Cols = 13
    sqlx = "select * from racks where aisle = '" & Combo1 & "'"
    sqlx = sqlx & " order by slot"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!id & Chr$(9)
            sqlx = sqlx & " " & ds!rack & Chr$(9)
            sqlx = sqlx & ds!capacity & Chr$(9)
            sqlx = sqlx & ds!sku & Chr$(9)
            sqlx = sqlx & ds!lot_num & Chr$(9)
            sqlx = sqlx & ds!qty & Chr$(9)
            sqlx = sqlx & ds!qty4 & Chr$(9)
            sqlx = sqlx & ds!resv_sku & Chr$(9)
            sqlx = sqlx & ds!resv_lot & Chr$(9)
            sqlx = sqlx & ds!fo & Chr$(9)
            sqlx = sqlx & ds!hold & Chr$(9)
            sqlx = sqlx & " " & Chr$(9)
            sqlx = sqlx & ds!slot
            Targets.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    For i = 1 To Targets.Rows - 1
        Targets.Row = i: Targets.Col = 3: psku = " "
        If Val(Targets.Text) > 0 Then
            psku = Targets.Text
        Else
            Targets.Col = 7
            If Val(Targets.Text) > 0 Then psku = Targets.Text
        End If
        If Val(psku) > 0 Then
            If skurec(Val(psku)).sku = psku Then
                Targets.Text = skurec(Val(psku)).prodname
            Else
                Targets.Text = " Invalid SKU"
            End If
        End If
    Next i
    Targets.FormatString = "#|^Rack|^Cap|^SKU|^Lot|^Qty|^4W|^Resv|^Lot|^FO|^Hold|^Description|^Slot"
    Targets.ColWidth(0) = 1
    Targets.ColWidth(1) = 1000: Targets.ColWidth(2) = 400
    Targets.ColWidth(3) = 500: Targets.ColWidth(4) = 600
    Targets.ColWidth(5) = 400: Targets.ColWidth(6) = 400
    Targets.ColWidth(7) = 500: Targets.ColWidth(8) = 600
    Targets.ColWidth(9) = 300: Targets.ColWidth(10) = 450
    Targets.ColWidth(11) = 2700: Targets.ColWidth(12) = 400
    Screen.MousePointer = 0
    Targets.Visible = True
    If Targets.Rows > 1 Then
        Targets.Row = 1
        Call Targets_Click
    End If
End Sub

Private Sub Combo1_Click()
    Call refresh_targets
End Sub

Private Sub Command1_Click()
    Dim i As Integer, sqlx As String
    Targets.Col = 5
    If Val(Targets.Text) > 0 Then
        MsgBox "Target Rack contains BB Pallets", vbOKOnly, "Invalid Move"
        Exit Sub
    End If
    Targets.Col = 6
    If Val(Targets.Text) > 0 Then
        MsgBox "Target Rack contains 4Way Pallets", vbOKOnly, "Invalid Move"
        Exit Sub
    End If
    Targets.Col = 2
    If Val(Targets.Text) < Val(pqty) + Val(pqty4) Then
        MsgBox "Target Rack shows insuffecient capacity", vbOKOnly, "Sorry"
        Exit Sub
    End If
    sqlx = "Update racks set sku = ' ', lot_num = ' ', qty = 0, qty4 = 0 Where id = " & RKey
    Wdb.Execute sqlx
    Targets.Col = 0
    sqlx = "Update racks set sku = '" & psku & "', lot_num = '" & plot & "'"
    sqlx = sqlx & ", qty = " & pqty & ", qty4 = " & pqty4 & " Where id = " & Targets.Text
    Wdb.Execute sqlx
    i = Form4.RGrid.Row
    Form4.RGrid.TextMatrix(i, 3) = " "
    Form4.RGrid.TextMatrix(i, 4) = " "
    Form4.RGrid.TextMatrix(i, 5) = " "
    Form4.RGrid.TextMatrix(i, 6) = " "
    Form4.RGrid.TextMatrix(i, 11) = " "
    If Trim(Form4.PAisle) = Trim(Combo1) Then
        For i = 1 To Form4.RGrid.Rows - 1
            If Val(Form4.RGrid.TextMatrix(i, 0)) = Val(Targets.TextMatrix(Targets.Row, 0)) Then
                Form4.RGrid.TextMatrix(i, 3) = " " & psku
                Form4.RGrid.TextMatrix(i, 4) = " " & plot
                Form4.RGrid.TextMatrix(i, 5) = pqty
                Form4.RGrid.TextMatrix(i, 6) = pqty4
                Form4.RGrid.TextMatrix(i, 11) = " " & pdesc
            End If
        Next i
    End If
    Unload Form6
End Sub

Private Sub Command2_Click()
    Dim ptop As Integer, pbot As Integer, pkey As Long, nkey As Long
    Dim proom As String, PAisle As String, prack As String
    Dim sqlx As String
    Dim pans As String, y As Integer, i As Integer, j As Integer
    y = Targets.Row
    prack = Targets.TextMatrix(y, 1)
    pkey = Val(Targets.TextMatrix(y, 0))
    If pkey = 0 Then Exit Sub
    pans = InputBox$("Top Capacity:", "Split Rack " & prack, 0)
    If Val(pans) = 0 Then Exit Sub
    ptop = Val(pans)
    pans = InputBox$("Bottom Capacity:", "Split Rack " & prack, 0)
    If Val(pans) = 0 Then Exit Sub
    pbot = Val(pans)
    Targets.TextMatrix(y, 2) = ptop
    nkey = wd_seq("Racks")
    s = "INSERT INTO Racks (ID, Room, Aisle, Rack, Slot, Capacity, SKU, Lot_Num,"
    s = s & " Qty, Qty4, Resv_SKU, Resv_Lot, FO, Hold)"
    s = s & " VALUES (" & nkey & ","
    s = s & "'" & proom & "',"
    s = s & "'" & PAisle & "',"
    s = s & "'" & prack & "',"
    s = s & Val(Targets.TextMatrix(y, 12)) & ","
    s = s & pbot & ","
    s = s & "' ',"
    s = s & "'.',"
    s = s & "0,"
    s = s & "0,"
    s = s & "' ',"
    s = s & "' ',"
    s = s & "0,"
    s = s & "0)"
    Wdb.Execute s
    'Call refresh_targets
    Targets.AddItem " "
    For i = Targets.Rows - 2 To y Step -1
        For j = 0 To Targets.Cols - 1
            Targets.TextMatrix(i + 1, j) = Targets.TextMatrix(i, j)
        Next j
    Next i
    Targets.TextMatrix(y, 0) = nkey
    Targets.TextMatrix(y, 2) = pbot: Targets.TextMatrix(y, 3) = " "
    Targets.TextMatrix(y, 4) = " ": Targets.TextMatrix(y, 5) = "0"
    Targets.TextMatrix(y, 6) = "0": Targets.TextMatrix(y, 7) = " "
    Targets.TextMatrix(y, 8) = " ": Targets.TextMatrix(y, 9) = "0"
    Targets.TextMatrix(y, 10) = "0": Targets.TextMatrix(y, 11) = " "
    If Trim(Form4.PAisle) = Trim(Combo1) Then
        y = 0
        For i = 1 To Form4.RGrid.Rows - 1
            If Val(Form4.RGrid.TextMatrix(i, 0)) = pkey Then
                y = i
                Exit For
            End If
        Next i
        If y > 0 Then
            Form4.RGrid.TextMatrix(y, 2) = ptop
            Form4.RGrid.AddItem " "
            For i = Form4.RGrid.Rows - 2 To y Step -1
                For j = 0 To Form4.RGrid.Cols - 1
                    Form4.RGrid.TextMatrix(i + 1, j) = Form4.RGrid.TextMatrix(i, j)
                Next j
            Next i
            Form4.RGrid.TextMatrix(y, 0) = nkey
            Form4.RGrid.TextMatrix(y, 2) = pbot
            Form4.RGrid.TextMatrix(y, 3) = " "
            Form4.RGrid.TextMatrix(y, 4) = " "
            Form4.RGrid.TextMatrix(y, 5) = "0"
            Form4.RGrid.TextMatrix(y, 6) = "0"
            Form4.RGrid.TextMatrix(y, 7) = " "
            Form4.RGrid.TextMatrix(y, 8) = " "
            Form4.RGrid.TextMatrix(y, 9) = "0"
            Form4.RGrid.TextMatrix(y, 10) = "0"
            Form4.RGrid.TextMatrix(y, 11) = " "
        End If
    End If
End Sub

Private Sub Command3_Click()
    Dim sqlx As String, pkey As Long, y As Integer
    y = Targets.Row
    pkey = Val(Targets.TextMatrix(y, 0))
    If pkey = 0 Then Exit Sub
    If MsgBox("Are you sure?", vbOKCancel, "Clear Rack") = vbCancel Then Exit Sub
    Targets.TextMatrix(y, 3) = " ": Targets.TextMatrix(y, 4) = " "
    Targets.TextMatrix(y, 5) = "0": Targets.TextMatrix(y, 6) = "0"
    Targets.TextMatrix(y, 7) = " ": Targets.TextMatrix(y, 8) = " "
    Targets.TextMatrix(y, 11) = " "
    sqlx = "update racks set sku=' ',lot_num=' ',qty=0,qty4=0,resv_sku=' ',resv_lot=' '"
    sqlx = sqlx & " where id = " & pkey
    Wdb.Execute sqlx
    If Trim(Form4.PAisle) = Trim(Combo1) Then
        For y = 1 To Form4.RGrid.Rows - 1
            If Val(Form4.RGrid.TextMatrix(y, 0)) = pkey Then
                Form4.RGrid.TextMatrix(y, 3) = " "
                Form4.RGrid.TextMatrix(y, 4) = " "
                Form4.RGrid.TextMatrix(y, 5) = "0"
                Form4.RGrid.TextMatrix(y, 6) = "0"
                Form4.RGrid.TextMatrix(y, 7) = " "
                Form4.RGrid.TextMatrix(y, 8) = " "
                Form4.RGrid.TextMatrix(y, 11) = " "
                Exit For
            End If
        Next y
    End If
    Call Targets_Click
End Sub

Private Sub Command4_Click()
    y = Targets.Row
    Form5.RKey = Val(Targets.TextMatrix(y, 0))
    Form5.Caption = "Rack " & Trim$(Targets.TextMatrix(y, 1))
    Form5.Text1 = Trim$(Targets.TextMatrix(y, 3))
    Form5.Text2 = Trim$(Targets.TextMatrix(y, 4))
    Form5.Text3 = Trim$(Targets.TextMatrix(y, 5))
    Form5.Text4 = Trim$(Targets.TextMatrix(y, 6))
    Form5.Text5 = Trim$(Targets.TextMatrix(y, 7))
    Form5.Text6 = Trim$(Targets.TextMatrix(y, 8))
    Form5.Text7 = Trim$(Targets.TextMatrix(y, 2))
    Form5.Text8 = Trim$(Targets.TextMatrix(y, 12))
    Form5.Check1.Value = Val(Targets.TextMatrix(y, 9))
    Form5.Check2.Value = Val(Targets.TextMatrix(y, 10))
    Form5.Show
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Form6.WindowState = 0 Then
        For i = 1 To Form1.Frmgrid.Rows - 1
            If Form1.Frmgrid.TextMatrix(i, 0) = "form6" Then
                Form1.Frmgrid.TextMatrix(i, 1) = Form6.Top
                Form1.Frmgrid.TextMatrix(i, 2) = Form6.Left
                Form1.Frmgrid.TextMatrix(i, 3) = Form6.Height
                Form1.Frmgrid.TextMatrix(i, 4) = Form6.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim ds As ADODB.Recordset, sqlx As String
    Dim i As Integer
    For i = 1 To Form1.Frmgrid.Rows - 1
        If Form1.Frmgrid.TextMatrix(i, 0) = "form6" Then
            Form6.Top = Val(Form1.Frmgrid.TextMatrix(i, 1))
            Form6.Left = Val(Form1.Frmgrid.TextMatrix(i, 2))
            Form6.Height = Val(Form1.Frmgrid.TextMatrix(i, 3))
            Form6.Width = Val(Form1.Frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Combo1.Clear
    sqlx = "select distinct aisle from racks where aisle > ' ' order by aisle"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds(0)
            ds.MoveNext
        Loop
    End If
    ds.Close
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    If Form6.Height > 3630 Then
        Command2.Top = Form6.Height - 750
        Command3.Top = Command2.Top
        Command4.Top = Command2.Top
        Targets.Height = Command2.Top - Targets.Top - 150
    End If
    If Form6.Width > 8845 Then
        Targets.Width = 8845
    Else
        Targets.Width = Form6.Width - 100
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Targets_Click()
    Targets.Col = 0: Targets.ColSel = Targets.Cols - 1
    Targets.RowSel = Targets.Row
End Sub

Private Sub Targets_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Targets_Click
End Sub

Private Sub Targets_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then Call Command4_Click 'F2 - Edit Rack
    If KeyCode = 114 Then Call Command2_Click 'F3 - Split Rack
    If KeyCode = 120 Then Call Command3_Click 'F9 - Clear Rack
End Sub
