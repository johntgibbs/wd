VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Warehouses 
   Caption         =   "Warehouse Listing"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form2"
   ScaleHeight     =   7650
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   11456
      _Version        =   327680
      ForeColor       =   8421376
      BackColorFixed  =   12648384
      BackColorBkg    =   -2147483633
      FocusRect       =   0
   End
End
Attribute VB_Name = "Warehouses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    s = "select * from warehouses order by whs_num"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!whs_num & Chr(9)
            s = s & ds!plant & Chr(9)
            s = s & ds!whs & Chr(9)
            s = s & ds!whsname & Chr(9)
            s = s & ds!vert_loc & Chr(9)
            s = s & ds!horz_loc & Chr(9)
            s = s & ds!rack_side
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^Whs_Num|^Plant|^Whs|<Name|^Vert_loc|^Horz_loc|^Rack_side"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1600
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Warehouses.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
            If Form1.FrmGrid.Text = "warehouses" Then
                Form1.FrmGrid.Col = 1: Form1.FrmGrid.Text = Warehouses.Top
                Form1.FrmGrid.Col = 2: Form1.FrmGrid.Text = Warehouses.Left
                Form1.FrmGrid.Col = 3: Form1.FrmGrid.Text = Warehouses.Height
                Form1.FrmGrid.Col = 4: Form1.FrmGrid.Text = Warehouses.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.FrmGrid.Rows - 1
        Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
        If Form1.FrmGrid.Text = "warehouses" Then
            Form1.FrmGrid.Col = 1: Warehouses.Top = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 2: Warehouses.Left = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 3: Warehouses.Height = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 4: Warehouses.Width = Val(Form1.FrmGrid.Text)
            Exit For
        End If
    Next i
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 480
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub
