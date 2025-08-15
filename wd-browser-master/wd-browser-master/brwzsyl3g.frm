VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brwzsyl3g 
   Caption         =   "Form13"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13710
   LinkTopic       =   "Form13"
   ScaleHeight     =   10095
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8281
      _Version        =   327680
      BackColorFixed  =   12640511
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   16536
      _Version        =   327680
      Cols            =   3
      FixedCols       =   2
      BackColorFixed  =   12648447
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "label3"
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
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   9600
      Visible         =   0   'False
      Width           =   11175
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "> 30 Day Supply"
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
      Left            =   5880
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "30 Day Supply"
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
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2 Week Supply"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "< 2 Week Supply"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Brenham Manufactured Products Issued from Sylacauga"
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
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sylacauga 3 Gallon Products"
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
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   5775
   End
End
Attribute VB_Name = "brwzsyl3g"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_bimp_3g()
    Dim ds As ADODB.Recordset, s As String, i As Integer
    'Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    
    s = s & " Sylacauga 3 Gallon Inventory"
    s = s & "  Last update: " & bimp_status_time
    's = s & Format(FileDateTime(Form1.webdir & "\stock\gsales." & rpcode), "m-d-yyyy h:mm am/pm")
    Me.Caption = s
    
    
    s = "select sku, count(*) from bimp where plantwhs = 'A10' group by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = Val(ds!sku)
            If skurec(i).unit = "3GAL" Then
                s = i & Chr(9) & skurec(i).desc
                Grid1.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select * from bimp where plantwhs = 'A10' and branchwhs in "
    s = s & "(select listreturn from valuelists where listname = 'branchplants' and listdisplay = 'A10')"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = ds!sku Then
                    Grid1.TextMatrix(i, 2) = ds!plantpool + ds!onorder
                    Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + ds!onhand
                    Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds!onorder
                    Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds!sales
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select sku, onhand from bimp where plantwhs = 'A10' and branchwhs = '052'"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = ds!sku Then
                    Grid1.TextMatrix(i, 2) = Val(Grid1.TextMatrix(i, 2)) + ds!onhand
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    
    s = "select sku, plantpool from bimp where plantwhs = 'T10' and branchwhs = '004'"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = ds!sku Then
                    Grid1.TextMatrix(i, 11) = ds!plantpool '+ ds!onorder
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 2)) + Val(Grid1.TextMatrix(i, 3))
            Grid1.TextMatrix(i, 7) = Val(Grid1.TextMatrix(i, 4)) - Val(Grid1.TextMatrix(i, 6))
            If Val(Grid1.TextMatrix(i, 7)) <> 0 Then
                Grid1.TextMatrix(i, 8) = CInt(Val(Grid1.TextMatrix(i, 7)) / 60)
            End If
            If Val(Grid1.TextMatrix(i, 4)) <> 0 And Val(Grid1.TextMatrix(i, 6)) <> 0 Then
                Grid1.TextMatrix(i, 10) = CLng((Val(Grid1.TextMatrix(i, 4)) / Val(Grid1.TextMatrix(i, 6))) * 30)
            End If
            If Val(Grid1.TextMatrix(i, 8)) > 0 Then
                Grid1.TextMatrix(i, 9) = "G"
            Else
                If Val(Grid1.TextMatrix(i, 8)) = 0 Then
                    Grid1.TextMatrix(i, 9) = "B"
                Else
                    If Val(Grid1.TextMatrix(i, 10)) < 14 Then
                        Grid1.TextMatrix(i, 9) = "W"
                    Else
                        Grid1.TextMatrix(i, 9) = "Y"
                    End If
                End If
            End If
        Next i
    End If
    
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 6: Grid1.ColSel = 6
    Grid1.Sort = 4
    Grid1.FormatString = "^SKU|<Flavor|^Sylacauga Stock|^Branch Units|^Total Units|^Branch Orders|^Sales Last 30|^Units Diff|^Pallet Diff|||^Brenham Stock"
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 3200
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1
    Grid1.ColWidth(10) = 1
    Grid1.ColWidth(11) = 1000
    Grid1.FillStyle = flexFillRepeat
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i: Grid1.Col = 2: Grid1.ColSel = 9
        If Grid1.TextMatrix(i, 9) = "W" Then Grid1.CellBackColor = wcolor.BackColor
        If Grid1.TextMatrix(i, 9) = "B" Then Grid1.CellBackColor = bcolor.BackColor
        If Grid1.TextMatrix(i, 9) = "G" Then Grid1.CellBackColor = gcolor.BackColor
        If Grid1.TextMatrix(i, 9) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
    Next i
    If Grid1.Rows > 1 Then
        Grid1.RowHeight(0) = Grid1.RowHeight(1) * 2
        Grid1.Row = 1: Grid1.Col = 4
    End If
End Sub

Private Sub refresh_grid1()
    Dim psku As String, pdesc As String, ppoh As String
    Dim puoh As String, poo As String, psales As String
    Dim udiff As String, pdiff As String, plnt As String
    Dim plit As String, pcc As String
    Dim s As String, pro As String, rpcode As String
    Dim ts As String
    Grid1.Clear: Grid1.Rows = 1
    Grid1.Cols = 12
    rpcode = "502"
    If Len(Dir(Form1.webdir & "\stock\gsales." & rpcode)) > 0 Then
        s = s & " Sylacauga 3 Gallon Inventory"
        s = s & "  Last update: "
        s = s & Format(FileDateTime(Form1.webdir & "\stock\gsales." & rpcode), "m-d-yyyy h:mm am/pm")
        Me.Caption = s
    
        Open Form1.webdir & "\stock\gsales." & rpcode For Input As #1
        Do Until EOF(1)
            Input #1, psku, pdesc, ppoh, puoh, poo, psales, udiff, pdiff, plnt, pro, plit, pcc
            If plnt = "50" Then plnt = "TX"
            If plnt = "51" Then plnt = "OK"
            If plnt = "52" Then plnt = "AL"
            s = psku & Chr(9) & pdesc & Chr(9)
            s = s & Format(plit, "#") & Chr(9)
            s = s & Format(Val(puoh) - Val(plit), "#") & Chr(9)
            s = s & Format(puoh, "#") & Chr(9)
            s = s & Format(poo, "#") & Chr(9)
            s = s & Format(Val(psales), "0") & Chr(9)
            s = s & Format(udiff, "#") & Chr(9)
            s = s & Format(pdiff, "#") & Chr(9)
            s = s & pcc & Chr(9)
            s = s & plit
            If UCase(Left(pdesc, 2)) = "3G" Then Grid1.AddItem s
        Loop
    End If
    Close #1
    
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 6: Grid1.ColSel = 6
    Grid1.Sort = 4
    Grid1.FormatString = "^SKU|<Description|^Sylacauga Stock|^Branch Units|^Total Units|^Branch Orders|^Sales Last 30|^Units Diff|^Pallet Diff|||^Brenham Stock"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 3200
    Grid1.ColWidth(2) = 900
    Grid1.ColWidth(3) = 900
    Grid1.ColWidth(4) = 900
    Grid1.ColWidth(5) = 900
    Grid1.ColWidth(6) = 900
    Grid1.ColWidth(7) = 900
    Grid1.ColWidth(8) = 900
    Grid1.ColWidth(9) = 1
    Grid1.ColWidth(10) = 1
    Grid1.ColWidth(11) = 1000
    Grid1.FillStyle = flexFillRepeat
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i: Grid1.Col = 0: Grid1.ColSel = 9
        If Grid1.TextMatrix(i, 9) = "W" Then Grid1.CellBackColor = wcolor.BackColor
        If Grid1.TextMatrix(i, 9) = "B" Then Grid1.CellBackColor = bcolor.BackColor
        If Grid1.TextMatrix(i, 9) = "G" Then Grid1.CellBackColor = gcolor.BackColor
        If Grid1.TextMatrix(i, 9) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
    Next i
    If Grid1.Rows > 1 Then
        Grid1.RowHeight(0) = Grid1.RowHeight(1) * 2
        Grid1.Row = 1: Grid1.Col = 4
    End If
End Sub

Private Sub refresh_grid2()
    Dim psku As String, pdesc As String, ppoh As String
    Dim puoh As String, poo As String, psales As String
    Dim udiff As String, pdiff As String, plnt As String
    Dim plit As String, pcc As String
    Dim s As String, pro As String, rpcode As String
    Dim ts As String
    Grid2.Clear: Grid2.Rows = 1
    Grid2.Cols = 12
    rpcode = "507"
    If Len(Dir(Form1.webdir & "\stock\gsales." & rpcode)) > 0 Then
        Open Form1.webdir & "\stock\gsales." & rpcode For Input As #1
        Do Until EOF(1)
            Input #1, psku, pdesc, ppoh, puoh, poo, psales, udiff, pdiff, plnt, pro, plit, pcc
            If plnt = "50" Then plnt = "TX"
            If plnt = "51" Then plnt = "OK"
            If plnt = "52" Then plnt = "AL"
            s = psku & Chr(9) & pdesc & Chr(9)
            s = s & Format(plit, "#") & Chr(9)
            s = s & Format(Val(puoh) - Val(plit), "#") & Chr(9)
            s = s & Format(puoh, "#") & Chr(9)
            s = s & Format(poo, "#") & Chr(9)
            s = s & Format(Val(psales), "0") & Chr(9)
            s = s & Format(udiff, "#") & Chr(9)
            s = s & Format(pdiff, "#") & Chr(9)
            s = s & pcc & Chr(9)
            s = s & plit
            If UCase(Left(pdesc, 2)) = "3G" Then Grid2.AddItem s
        Loop
    End If
    Close #1
    Grid2.RowSel = Grid2.Row
    Grid2.Col = 6: Grid2.ColSel = 6
    'Grid2.Col = 1: Grid2.ColSel = 1
    Grid2.Sort = 4
    
    Grid2.FormatString = "^SKU|<Description|^Sylacauga Stock|^* Branch Units|^Total Units|^Branch Orders|^Sales Last 30|^Units Diff|^Pallet Diff|||^Brenham Stock"
    Grid2.ColWidth(0) = 500
    Grid2.ColWidth(1) = 3200
    Grid2.ColWidth(2) = 900
    Grid2.ColWidth(3) = 900
    Grid2.ColWidth(4) = 900
    Grid2.ColWidth(5) = 900
    Grid2.ColWidth(6) = 900
    Grid2.ColWidth(7) = 900
    Grid2.ColWidth(8) = 900
    Grid2.ColWidth(9) = 1
    Grid2.ColWidth(10) = 1
    Grid2.ColWidth(11) = 1000
    Grid2.FillStyle = flexFillRepeat
    For i = 1 To Grid2.Rows - 1
        Grid2.Row = i: Grid2.RowSel = i: Grid2.Col = 0: Grid2.ColSel = 9
        If Grid2.TextMatrix(i, 9) = "W" Then Grid2.CellBackColor = wcolor.BackColor
        If Grid2.TextMatrix(i, 9) = "B" Then Grid2.CellBackColor = bcolor.BackColor
        If Grid2.TextMatrix(i, 9) = "G" Then Grid2.CellBackColor = gcolor.BackColor
        If Grid2.TextMatrix(i, 9) = "Y" Then Grid2.CellBackColor = ycolor.BackColor
    Next i
    If Grid2.Rows > 1 Then
        Grid2.RowHeight(0) = Grid2.RowHeight(1) * 2
        Grid2.Row = 1: Grid2.Col = 4
    End If
End Sub

Private Sub brenham_oh()
    Dim psku As String, pdesc As String, ppoh As String
    Dim puoh As String, poo As String, psales As String
    Dim udiff As String, pdiff As String, plnt As String
    Dim plit As String, pcc As String
    Dim s As String, pro As String, rpcode As String
    Dim ts As String
    
    'Brenham Inventory
    rpcode = "500"
    If Len(Dir(Form1.webdir & "\stock\gsales." & rpcode)) > 0 Then
        Open Form1.webdir & "\stock\gsales." & rpcode For Input As #1
        Do Until EOF(1)
            Input #1, psku, pdesc, ppoh, puoh, poo, psales, udiff, pdiff, plnt, pro, plit, pcc
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = psku Then
                    Grid1.TextMatrix(i, 11) = plit
                    Grid1.Row = i: Grid1.RowSel = i
                    Grid1.Col = 11: Grid1.ColSel = 11
                    If pcc = "W" Then Grid1.CellBackColor = wcolor.BackColor
                    If pcc = "B" Then Grid1.CellBackColor = bcolor.BackColor
                    If pcc = "G" Then Grid1.CellBackColor = gcolor.BackColor
                    If pcc = "Y" Then Grid1.CellBackColor = ycolor.BackColor
                    Exit For
                End If
            Next i
            For i = 1 To Grid2.Rows - 1
                If Grid2.TextMatrix(i, 0) = psku Then
                    Grid2.TextMatrix(i, 11) = plit
                    Grid2.Row = i: Grid2.RowSel = i
                    Grid2.Col = 11: Grid2.ColSel = 11
                    If pcc = "W" Then Grid2.CellBackColor = wcolor.BackColor
                    If pcc = "B" Then Grid2.CellBackColor = bcolor.BackColor
                    If pcc = "G" Then Grid2.CellBackColor = gcolor.BackColor
                    If pcc = "Y" Then Grid2.CellBackColor = ycolor.BackColor
                    Exit For
                End If
            Next i
        Loop
    End If
    Close #1
End Sub

Private Sub Form_Load()
    'refresh_grid1
    'refresh_grid2
    'brenham_oh
    'Me.Width = Form1.Width
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
    refresh_bimp_3g
    Label3.Caption = " "
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    Grid2.Width = Me.Width - 80
    Label1.Width = Me.Width - 80
    Label2.Width = Me.Width - 80
    Label3.Width = Me.Width - 80
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (ycolor.Height * 4.5)
End Sub
