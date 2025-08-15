VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form allocations 
   Caption         =   "Product Allocations"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form2"
   ScaleHeight     =   5955
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6800
      _Version        =   327680
   End
   Begin VB.Label Label1 
      Caption         =   "trigkey"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "allocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function sku_alloc(psku As String, plot As String, whs As Integer) As Integer
    Dim i As Integer
    If Grid1.Rows < 2 Then
        sku_alloc = 0
    Else
        sku_alloc = 0
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 1) = psku And Grid1.TextMatrix(i, 5) = plot Then
                If whs = 1 Then sku_alloc = Val(Grid1.TextMatrix(i, 6))
                If whs = 2 Then sku_alloc = Val(Grid1.TextMatrix(i, 7))
                If whs = 3 Then sku_alloc = Val(Grid1.TextMatrix(i, 8))
                If whs = 4 Then sku_alloc = Val(Grid1.TextMatrix(i, 9))
                If whs = 5 Then sku_alloc = Val(Grid1.TextMatrix(i, 9))
                Exit For
            End If
        Next i
    End If
End Function

Sub refresh_sched()
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 10
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    s = "select * from prodrcv"
    s = s & " order by sku, lot_num"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & Format(ds!proddate, "MM-dd-yyyy") & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!sp_flag & Chr(9)
            s = s & ds!lot_num & Chr(9)
            s = s & ds!sr1 & Chr(9)
            s = s & ds!sr2 & Chr(9)
            s = s & ds!sr3 & Chr(9)
            s = s & ds!sr4
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    s = "^Id|^SKU|^ProdDate|^Units|^SP|^Lot|^Sr1|^Sr2|^Sr3|^Sr4"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1000
End Sub

Private Sub Form_Load()
    refresh_sched
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
End Sub

Private Sub Label1_Change()
    refresh_sched
End Sub

