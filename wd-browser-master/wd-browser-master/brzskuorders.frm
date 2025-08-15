VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brzskuorders 
   Caption         =   "SKU Orders"
   ClientHeight    =   8640
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Pallets"
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
      Left            =   6240
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Units"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Left            =   1080
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   600
      Width           =   4695
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
      Left            =   1080
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Width           =   4695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   11668
      _Version        =   327680
      ForeColor       =   128
      BackColorFixed  =   16777152
   End
   Begin VB.Label Label2 
      Caption         =   "Product:"
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
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Plant:"
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
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label msku 
      Caption         =   "msku"
      Height          =   255
      Left            =   10200
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label mplant 
      Caption         =   "mplant"
      Height          =   255
      Left            =   10200
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "brzskuorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, s As String, i As Integer, c As Long
    Dim t1 As Long, t2 As Long, t3 As Long, t4 As Long, oqty As Long
    Dim ss As ADODB.Recordset
    t1 = 0: t2 = 0: t3 = 0: t4 = 0
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5: Grid1.FixedCols = 1
    s = "select * from bimp where sku = '" & msku & "'"
    s = s & " and plantwhs = '" & mplant.Caption & "'"
    s = s & " order by branchwhs"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            oqty = ds!onorder
            s = "select sku, sum(netqty) from brorders where sku = '" & msku & "'"
            s = s & " and branch = " & Val(ds!branchwhs)
            If mplant = "T10" Then s = s & " and plant = 50"
            If mplant = "K10" Then s = s & " and plant = 51"
            If mplant = "A10" Then s = s & " and plant = 52"
            s = s & " group by sku having sum(netqty) <> 0"
            Set ss = wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                If IsNull(ss(1)) = False Then oqty = oqty + (ss(1) * ds!roqty)
                'MsgBox s & " = " & oqty
            End If
            ss.Close
            oqty = oqty + (groupitems_qty(msku, mplant, ds!branchwhs) * ds!roqty)           'jv081516
            
            s = ds!plantwhs & Chr(9)
            s = ds!branchwhs & "-" & branchrec(Val(ds!branchwhs)).branchname & Chr(9)
            's = s & Format(ds!onorder, "#") & Chr(9)
            s = s & Format(oqty, "#") & Chr(9)
            s = s & Format((ds!thiswknewpals * ds!roqty), "#") & Chr(9)
            s = s & Format((ds!nextwknewpals * ds!roqty), "#") & Chr(9)
            's = s & (ds!onorder + (ds!thiswknewpals * ds!roqty) + (ds!nextwknewpals * ds!roqty))
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Option2 = True Then
                mpal = 0
                If skurec(Val(msku)).sku = msku Then
                    mpal = skurec(Val(msku)).pallet
                End If
                If mpal <> 0 Then
                    Grid1.TextMatrix(i, 1) = Format(CInt(Val(Grid1.TextMatrix(i, 1)) / mpal), "#")
                    Grid1.TextMatrix(i, 2) = Format(CInt(Val(Grid1.TextMatrix(i, 2)) / mpal), "#")
                    Grid1.TextMatrix(i, 3) = Format(CInt(Val(Grid1.TextMatrix(i, 3)) / mpal), "#")
                End If
            End If
            Grid1.TextMatrix(i, 4) = Grid1.TextMatrix(i, 1)
            Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + Val(Grid1.TextMatrix(i, 2))
            Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + Val(Grid1.TextMatrix(i, 3))
            Grid1.TextMatrix(i, 4) = Format(Val(Grid1.TextMatrix(i, 4)), "#")
            t1 = t1 + Val(Grid1.TextMatrix(i, 1))
            t2 = t2 + Val(Grid1.TextMatrix(i, 2))
            t3 = t3 + Val(Grid1.TextMatrix(i, 3))
            t4 = t4 + Val(Grid1.TextMatrix(i, 4))
        Next i
        Grid1.Row = 1
    End If
    s = "Totals" & Chr(9) & t1 & Chr(9)
    s = s & t2 & Chr(9) & t3 & Chr(9) & t4
    Grid1.AddItem s
    
    c = Grid1.BackColor
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
        Grid1.CellBackColor = c
        If c = Grid1.BackColorFixed Then
            c = Grid1.BackColor
        Else
            c = Grid1.BackColorFixed ' wcolor.BackColor
        End If
    Next i
    Grid1.Row = 1
        
    'Grid1.Cols = Grid1.Cols - 1
    s = "<Branch|^Active|^This Week|^Next Week|^Total"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 2200
    Grid1.ColWidth(1) = 1300
    Grid1.ColWidth(2) = 1300
    Grid1.ColWidth(3) = 1300
    Grid1.ColWidth(4) = 1300
    Grid1.Redraw = True
End Sub

Private Sub Form_Load()
    Me.Height = brzplantorders.Height
    Me.Top = brzplantorders.Top
    Me.Left = Form1.Width - Me.Width
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Text1.Height * 6)
End Sub

Private Sub mplant_Change()
    If Val(msku.Caption) > 0 Then refresh_grid
End Sub

Private Sub msku_Change()
    If mplant.Caption <> "mplant" Then refresh_grid
End Sub

Private Sub Option1_Click()
    If Val(msku.Caption) > 0 Then refresh_grid
End Sub

Private Sub Option2_Click()
    If Val(msku.Caption) > 0 Then refresh_grid
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    rt = Text1 & " Branch Orders"
    rh = Text2
    If Option1 = True Then
        rh = rh & " - Units"
    Else
        rh = rh & " - Pallets"
    End If
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "lightcyan": gndc(0) = Grid1.BackColorFixed
    'htdc(1) = "yellow": gndc(1) = Me.ycolor.BackColor
    'htdc(2) = "white": gndc(2) = Me.wcolor.BackColor
    Grid1.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, localAppDataPath & "\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "lightyellow", "white")
        i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmlgrid.htm", vbNormalFocus)
        Grid1.Redraw = True: Grid1.Row = 1
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, localAppDataPath & "\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "lightyellow", "white")
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmlgrid.htm", vbNormalFocus)
        Grid1.Redraw = True: Grid1.Row = 1
        Exit Sub
    End If
End Sub
