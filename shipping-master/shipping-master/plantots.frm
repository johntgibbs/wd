VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form plantots 
   Caption         =   "Plant Pallet Totals"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form2"
   ScaleHeight     =   3765
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5530
      _Version        =   327680
      Cols            =   11
      FixedCols       =   2
      BackColorFixed  =   16777152
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "plantots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    Dim i As Integer, k As Integer
    On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1
    sqlx = "select * from skumast order by sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!sku & Chr(9) & ds!fgunit & " " & ds!fgdesc
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    k = 1
    sqlx = "select * from whstotals where whs_num in (11,1,2,3,13,10,4,5,6,12)"
    sqlx = sqlx & " order by sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = k To Grid1.Rows - 1
                If ds!sku = Grid1.TextMatrix(i, 0) Then
                    Grid1.TextMatrix(i, 2) = Val(Grid1.TextMatrix(i, 2)) + ds!count_qty
                    Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + ds!grp_qty
                    Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + ds!avail
                    k = i
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    k = 1
    sqlx = "select * from whstotals where whs_num = 14"
    sqlx = sqlx & " order by sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = k To Grid1.Rows - 1
                If ds!sku = Grid1.TextMatrix(i, 0) Then
                    Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds!count_qty
                    Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds!grp_qty
                    Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 7)) + ds!avail
                    k = i
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    k = 1
    sqlx = "select * from whstotals where whs_num = 15"
    sqlx = sqlx & " order by sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = k To Grid1.Rows - 1
                If ds!sku = Grid1.TextMatrix(i, 0) Then
                    Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 8)) + ds!count_qty
                    Grid1.TextMatrix(i, 9) = Val(Grid1.TextMatrix(i, 9)) + ds!grp_qty
                    Grid1.TextMatrix(i, 10) = Val(Grid1.TextMatrix(i, 10)) + ds!avail
                    k = i
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    For i = Grid1.Rows - 1 To 1 Step -1
        k = Val(Grid1.TextMatrix(i, 2))
        k = k + Val(Grid1.TextMatrix(i, 3))
        k = k + Val(Grid1.TextMatrix(i, 4))
        k = k + Val(Grid1.TextMatrix(i, 5))
        k = k + Val(Grid1.TextMatrix(i, 6))
        k = k + Val(Grid1.TextMatrix(i, 7))
        k = k + Val(Grid1.TextMatrix(i, 8))
        k = k + Val(Grid1.TextMatrix(i, 9))
        k = k + Val(Grid1.TextMatrix(i, 10))
        If k = 0 Then Grid1.RemoveItem i
    Next i
    Grid1.FormatString = "^SKU|<Product|^TX Onhand|^TX Orders|^TX Avail|^OK Onhand|^OK Orders|^OK Avail|^AL Onhand|^AL Orders|^AL Avail"
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 3500
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 1200
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 1200
    Grid1.ColWidth(7) = 1200
    Grid1.ColWidth(8) = 1200
    Grid1.ColWidth(9) = 1200
    Grid1.ColWidth(10) = 1200
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

Private Sub Command1_Click()
    Dim ps As String
    Printer.FontName = "Courier New"
    Printer.FontSize = 8
    Printer.Orientation = 2
    Printer.Duplex = 3
    Printer.Print "Plant Pallet Totals"
    Printer.Print Format(Now, "m-dd-yyyy  h:mm Am/Pm")
    Printer.Print " "
    For i = 0 To Grid1.Rows - 1
        ps = Grid1.TextMatrix(i, 0) & " "
        ps = ps & Grid1.TextMatrix(i, 1)
        ps = ps & Space(50 - Len(ps))
        ps = ps & Grid1.TextMatrix(i, 2)
        ps = ps & Space(60 - Len(ps))
        ps = ps & Grid1.TextMatrix(i, 3)
        ps = ps & Space(70 - Len(ps))
        ps = ps & Grid1.TextMatrix(i, 4)
        ps = ps & Space(80 - Len(ps))
        ps = ps & Grid1.TextMatrix(i, 5)
        ps = ps & Space(90 - Len(ps))
        ps = ps & Grid1.TextMatrix(i, 6)
        ps = ps & Space(100 - Len(ps))
        ps = ps & Grid1.TextMatrix(i, 7)
        ps = ps & Space(110 - Len(ps))
        ps = ps & Grid1.TextMatrix(i, 8)
        ps = ps & Space(120 - Len(ps))
        ps = ps & Grid1.TextMatrix(i, 9)
        ps = ps & Space(130 - Len(ps))
        ps = ps & Grid1.TextMatrix(i, 10)
        Printer.Print ps
    Next i
    Printer.EndDoc
    Printer.Orientation = 1
    Printer.Duplex = 1
End Sub

Private Sub Form_Load()
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Call refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = plantots.Width - 120
    Grid1.Height = plantots.Height - 580
End Sub
