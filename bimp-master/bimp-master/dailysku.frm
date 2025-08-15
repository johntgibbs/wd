VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form dailysku 
   Caption         =   "Daily SKU Pallet Shipments"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12570
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   12570
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6015
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10610
      _Version        =   327680
      ForeColor       =   16512
      BackColorFixed  =   8454143
      ForeColorFixed  =   16512
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Warehouse:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Ship Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "dailysku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1_T10()
    Dim ds As ADODB.Recordset, s As String, i As Integer, nr As Boolean, t0 As Integer, c As Long
    Dim t1 As Integer, t2 As Integer, t3 As Integer, t4 As Integer, t5 As Integer, tt As Integer
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 9: Grid1.FixedCols = 2
    s = "select sku, sum(ordqty) from brorders where orddate = '" & Text1 & "'"
    s = s & " and plant = 50"
    s = s & " group by sku having sum(ordqty) > 0 order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds(1) & Chr(9)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    s = "select sku, whs_num, sum(pallets) from trailers where plant = 50"
    s = s & " and shipdate = '" & Text1 & "'"
    s = s & " group by sku, whs_num"
    s = s & " having sum(pallets) > 0"
    s = s & " order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            nr = True
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = ds!sku Then
                    If ds!whs_num = 1 Then Grid1.TextMatrix(i, 3) = ds(2)
                    If ds!whs_num = 2 Then Grid1.TextMatrix(i, 4) = ds(2)
                    If ds!whs_num = 3 Then Grid1.TextMatrix(i, 5) = ds(2)
                    If ds!whs_num <> 1 And ds!whs_num <> 2 And ds!whs_num <> 3 And ds!whs_num <> 5 Then
                        Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds(2)
                    End If
                    If ds!whs_num = 5 Then Grid1.TextMatrix(i, 7) = ds(2)
                    Grid1.TextMatrix(i, 8) = Val(Grid1.TextMatrix(i, 8)) + ds(2)
                    nr = False
                    Exit For
                End If
            Next i
            If nr = True Then
                s = ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                If ds!whs_num = 1 Then
                    s = s & Chr(9) & ds(2) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ds(2)
                End If
                If ds!whs_num = 2 Then
                    s = s & Chr(9) & Chr(9) & ds(2) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ds(2)
                End If
                If ds!whs_num = 3 Then
                    s = s & Chr(9) & Chr(9) & Chr(9) & ds(2) & Chr(9) & Chr(9) & Chr(9) & ds(2)
                End If
                If ds!whs_num <> 1 And ds!whs_num <> 2 And ds!whs_num <> 3 And ds!whs_num <> 5 Then
                    s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ds(2) & Chr(9) & Chr(9) & ds(2)
                End If
                If ds!whs_num = 5 Then
                    s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ds(2) & Chr(9) & ds(2)
                End If
                Grid1.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    t0 = 0: t1 = 0: t2 = 0: t3 = 0: t4 = 0: t5 = 0: tt = 0
    Grid1.FillStyle = flexFillRepeat
    c = Grid1.BackColor
    pbr = " "
    
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            t0 = t0 + Val(Grid1.TextMatrix(i, 2))
            t1 = t1 + Val(Grid1.TextMatrix(i, 3))
            t2 = t2 + Val(Grid1.TextMatrix(i, 4))
            t3 = t3 + Val(Grid1.TextMatrix(i, 5))
            t4 = t4 + Val(Grid1.TextMatrix(i, 6))
            t5 = t5 + Val(Grid1.TextMatrix(i, 7))
            tt = tt + Val(Grid1.TextMatrix(i, 8))
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            If c = Grid1.BackColorFixed Then
                c = Grid1.BackColor
            Else
                c = Grid1.BackColorFixed
            End If
            Grid1.CellBackColor = c
            If Val(Grid1.TextMatrix(i, 2)) > Val(Grid1.TextMatrix(i, Grid1.Cols - 1)) Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 2: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = Text1.ForeColor
                Grid1.CellForeColor = Text1.BackColor
            End If
        Next i
        Grid1.Row = 1
    End If
    
    s = "All" & Chr(9) & "Summary" & Chr(9) & t0 & Chr(9) & t1 & Chr(9) & t2 & Chr(9) & t3 & Chr(9) & t4 & Chr(9)
    s = s & t5 & Chr(9) & tt
    Grid1.AddItem s
    s = "^SKU|<Product|^Orders|^SR1|^SR2|^SR3|^Racks|^SR5|^Total"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 3500
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_grid1_A10()
    Dim pb As ADODB.Connection
    Dim ds As ADODB.Recordset, s As String, i As Integer, nr As Boolean
    Dim t1 As Integer, t2 As Integer, tt As Integer, t0 As Integer, c As Long
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6: Grid1.FixedCols = 2
    s = "select sku, sum(ordqty) from brorders where orddate = '" & Text1 & "'"
    s = s & " and plant = 52"
    s = s & " group by sku having sum(ordqty) > 0 order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds(1) & Chr(9)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    Set pb = CreateObject("ADODB.Connection")
    pb.Open a10ship
    s = "select sku, whs_num, sum(pallets) from trailers where plant = 52"
    s = s & " and shipdate = '" & Text1 & "'"
    s = s & " group by sku, whs_num"
    s = s & " having sum(pallets) > 0"
    s = s & " order by sku"
    Set ds = pb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            nr = True
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = ds!sku Then
                    If ds!whs_num = 1 Then
                        Grid1.TextMatrix(i, 3) = ds(2)
                    Else
                        Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + ds(2)
                    End If
                    Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(2)
                    nr = False
                    Exit For
                End If
            Next i
            If nr = True Then
                s = ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                If ds!whs_num = 1 Then
                    s = s & Chr(9) & ds(2) & Chr(9) & Chr(9) & ds(2)
                Else
                    s = s & Chr(9) & Chr(9) & ds(2) & Chr(9) & ds(2)
                End If
                Grid1.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close: pb.Close
    t0 = 0: t1 = 0: t2 = 0:  tt = 0
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            t0 = t0 + Val(Grid1.TextMatrix(i, 2))
            t1 = t1 + Val(Grid1.TextMatrix(i, 3))
            t2 = t2 + Val(Grid1.TextMatrix(i, 4))
            tt = tt + Val(Grid1.TextMatrix(i, 5))
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            If c = Grid1.BackColorFixed Then
                c = Grid1.BackColor
            Else
                c = Grid1.BackColorFixed
            End If
            Grid1.CellBackColor = c
            If Val(Grid1.TextMatrix(i, 2)) > Val(Grid1.TextMatrix(i, Grid1.Cols - 1)) Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 2: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = Text1.ForeColor
                Grid1.CellForeColor = Text1.BackColor
            End If
        Next i
        Grid1.Row = 1
    End If
    s = "All" & Chr(9) & "Summary" & Chr(9) & t0 & Chr(9) & t1 & Chr(9) & t2 & Chr(9) & tt
    Grid1.AddItem s
    s = "^SKU|<Product|^Orders|^CS5|^Racks|^Total"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 3500
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_grid1_K10()
    Dim pb As ADODB.Connection
    Dim ds As ADODB.Recordset, s As String, i As Integer, nr As Boolean
    Dim tt As Integer, t0 As Integer, c As Long
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4: Grid1.FixedCols = 2
    s = "select sku, sum(ordqty) from brorders where orddate = '" & Text1 & "'"
    s = s & " and plant = 51"
    s = s & " group by sku having sum(ordqty) > 0 order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds(1) & Chr(9)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    Set pb = CreateObject("ADODB.Connection")
    pb.Open k10ship
    s = "select sku, sum(pallets) from trailers where plant = 51"
    s = s & " and shipdate = '" & Text1 & "'"
    s = s & " group by sku"
    s = s & " having sum(pallets) > 0"
    s = s & " order by sku"
    Set ds = pb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            nr = True
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = ds!sku Then
                    Grid1.TextMatrix(i, 3) = ds(1)
                    nr = False
                    Exit For
                End If
            Next i
            If nr = True Then
                s = ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                s = s & Chr(9) & ds(1)
                Grid1.AddItem s
            End If
            ds.MoveNext
        Loop
        
    End If
    ds.Close: pb.Close
    tt = 0: t0 = 0
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            t0 = t0 + Val(Grid1.TextMatrix(i, 2))
            tt = tt + Val(Grid1.TextMatrix(i, 3))
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            If c = Grid1.BackColorFixed Then
                c = Grid1.BackColor
            Else
                c = Grid1.BackColorFixed
            End If
            Grid1.CellBackColor = c
            If Val(Grid1.TextMatrix(i, 2)) > Val(Grid1.TextMatrix(i, Grid1.Cols - 1)) Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 2: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = Text1.ForeColor
                Grid1.CellForeColor = Text1.BackColor
            End If
        Next i
        Grid1.Row = 1
    End If
    s = "All" & Chr(9) & "Summary" & Chr(9) & t0 & Chr(9) & tt
    Grid1.AddItem s
    s = "^SKU|<Product|^Orders|^Pallets"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 3500
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
    If Combo1 = "T10" Then refresh_grid1_T10
    If Combo1 = "A10" Then refresh_grid1_A10
    If Combo1 = "K10" Then refresh_grid1_K10
End Sub

Private Sub Command2_Click()
    Dim rt As String, rf As String, rh As String
    rt = Me.Caption
    rh = "Ship Date: " & Text1 & "  Plant: " & Combo1
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "white": gndc(0) = Me.Grid1.BackColorFixed
    htdc(1) = "yellow": gndc(1) = Me.Grid1.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid1.BackColor
    Grid1.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Text1 = Format(Now, "M-d-yyyy")
    Combo1.Clear
    Combo1.AddItem "T10"
    Combo1.AddItem "K10"
    Combo1.AddItem "A10"
    Combo1.ListIndex = 0
    Me.Height = whssales.Height
    Me.Top = whssales.Top
    Me.Left = whssales.Width - Me.Width
    Command1_Click
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 200
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Command1.Height * 3)
End Sub
