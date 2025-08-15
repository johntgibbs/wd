VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form oplanes 
   Caption         =   "Order Pick Lanes"
   ClientHeight    =   7425
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form11"
   ScaleHeight     =   7425
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7646
      _Version        =   327680
      BackColorFixed  =   16777152
      FocusRect       =   0
   End
   Begin VB.Menu prtmenu 
      Caption         =   "&Print"
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu insrec 
         Caption         =   "Insert Bay - F10"
      End
      Begin VB.Menu delrec 
         Caption         =   "Remove Bay - F9"
      End
      Begin VB.Menu edsku 
         Caption         =   "Change SKU - F2"
      End
   End
End
Attribute VB_Name = "oplanes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, sqlx As String
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Cols = 5: Grid1.Rows = 1
    Grid1.FixedCols = 3
    sqlx = "select * from opbays order by whse_num,vert_loc, horz_loc"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!id & Chr(9)
            sqlx = sqlx & ds!whse_num & Chr(9)
            sqlx = sqlx & ds!vert_loc & " "
            sqlx = sqlx & ds!horz_loc & " "
            sqlx = sqlx & ds!rack_side & Chr(9)
            sqlx = sqlx & ds!sku & Chr(9)
            sqlx = sqlx & ds!oplabel
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^ID|^Whs|^Lane|^Sku|<Description"
    Grid1.ColWidth(0) = 700
    Grid1.ColWidth(1) = 700
    Grid1.ColWidth(2) = 900
    Grid1.ColWidth(3) = 700
    Grid1.ColWidth(4) = 4500
    Grid1.Redraw = True
End Sub

Private Sub delrec_Click()
    Dim sqlx As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    sqlx = "Ok to delete lane " & Grid1.TextMatrix(Grid1.Row, 2) & "?"
    If MsgBox(sqlx, vbYesNo + vbQuestion, "Are you sure...") = vbNo Then Exit Sub
    sqlx = "DELETE FROM OPBays WHERE id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Wdb.Execute sqlx
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        refresh_grid
    End If
End Sub

Private Sub edsku_Click()
    Dim ds As ADODB.Recordset, sqlx As String
    Dim s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    s = Grid1.TextMatrix(Grid1.Row, 3)
    s = InputBox("SKU:", "Sku number...", s)
    If Len(s) = 0 Then Exit Sub
    If skurec(Val(s)).sku = s Then
        pdesc = skurec(Val(s)).uom_type
        If Left(pdesc, 3) = "1/2" Then pdesc = "1/2"
        pdesc = pdesc & " " & skurec(Val(s)).desc
        pdesc = StrConv(pdesc, vbProperCase)
    Else
        s = "..."
        pdesc = "..."
    End If
    sqlx = "select * from opbays where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        sqlx = "Update opbays set sku = '" & s & "', oplabel = '" & pdesc & "' Where id = " & ds!id
        Wdb.Execute sqlx
    End If
    ds.Close
    Grid1.TextMatrix(Grid1.Row, 3) = s
    Grid1.TextMatrix(Grid1.Row, 4) = pdesc
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 120 Then
        KeyCode = 0
        delrec_Click
    End If
    If KeyCode = 121 Then
        KeyCode = 0
        insrec_Click
    End If
    If KeyCode = 113 Then
        KeyCode = 0
        edsku_Click
    End If
End Sub

Private Sub Form_Load()
    refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 780
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub insrec_Click()
    Dim sqlx As String
    Dim w As String, v As String, h As String, s As String, oid As Long
    w = InputBox("Warehouse:", "Warehouse...", "1")
    If Len(w) = 0 Then Exit Sub
    If Val(w) < 1 Or Val(w) > 3 Then
        MsgBox "Invalid Crane " & w & "..", vbOKOnly + vbInformation, "sorry, try again..."
        Exit Sub
    End If
    v = InputBox("Vertical:", "Vertical...", "2")
    If Len(v) = 0 Then Exit Sub
    If Val(v) < 1 Or Val(v) > 7 Then
        MsgBox "Invalid Vertical " & v & "..", vbOKOnly + vbInformation, "sorry, try again..."
        Exit Sub
    End If
    h = InputBox("Horizontal:", "Horizontal...", "1")
    If Len(h) = 0 Then Exit Sub
    If Val(h) < 1 Or Val(h) > 50 Then
        MsgBox "Invalid Horizontal " & h & "..", vbOKOnly + vbInformation, "sorry, try again..."
        Exit Sub
    End If
    s = InputBox("Side:", "Side...", "R")
    If Len(s) = 0 Then Exit Sub
    If s <> "R" And s <> "L" Then
        MsgBox "Invalid Side " & s & "..", vbOKOnly + vbInformation, "sorry, try again..."
        Exit Sub
    End If
    oid = wd_seq("OPBays")
    sqlx = "INSERT INTO OPBays (ID, Whse_Num, Vert_Loc, Horz_Loc, Rack_Side, SKU, OPLabel, OPSeq)"
    sqlx = sqlx & " VALUES (" & oid & ","
    sqlx = sqlx & Val(w) & ","
    sqlx = sqlx & Val(v) & ","
    sqlx = sqlx & Val(h) & ","
    sqlx = sqlx & "'" & s & "',"
    sqlx = sqlx & "'...',"
    sqlx = sqlx & "'...',"
    sqlx = sqlx & "0)"
    Wdb.Execute sqlx
    sqlx = oid & Chr(9)
    sqlx = sqlx & w & Chr(9)
    sqlx = sqlx & v & " "
    sqlx = sqlx & h & " "
    sqlx = sqlx & s & Chr(9)
    sqlx = sqlx & "..." & Chr(9)
    sqlx = sqlx & "..."
    Grid1.AddItem sqlx, Grid1.Row
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    rt = Me.Caption
    rh = Format(Now, "mmmm d, yyyy")
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Call printflexgrid(Printer, Grid1, rt, rh, rf)
End Sub
