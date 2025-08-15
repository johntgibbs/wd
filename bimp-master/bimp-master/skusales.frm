VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form skusales 
   Caption         =   "Form1"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows Default
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
      Height          =   375
      Left            =   10200
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   4560
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   14843
      _Version        =   327680
      BackColorFixed  =   14737632
      FocusRect       =   0
      GridLines       =   2
   End
   Begin VB.Label psku 
      Caption         =   "Label1"
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Surplus"
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
      Left            =   6960
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Month Supply"
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
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "< Month"
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
      Left            =   5280
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "< 2 Weeks"
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
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "skusales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, s As String, i As Integer
    Dim t4 As Long, t5 As Long, t6 As Long, t7 As Long, t8 As Long, t12 As Long
    t4 = 0: t5 = 0: t6 = 0: t7 = 0: t8 = 0: t12 = 0
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 15: Grid1.FixedCols = 1 '3
    s = "select * from bimp where sku = '" & psku & "'"
    If List1 = "T10" Then s = s & " and plantwhs = 'T10'"
    If List1 = "K10" Then s = s & " and plantwhs = 'K10'"
    If List1 = "A10" Then s = s & " and plantwhs = 'A10'"
    s = s & " order by branchwhs"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!plantwhs & Chr(9)
            s = s & ds!branchwhs & Chr(9)
            If ds!discflag = "B" Then                                                   'jv091918
                s = s & branchrec(Val(ds!branchwhs)).branchname & " Blocked" & Chr(9)   'jv091918
            Else
                s = s & branchrec(Val(ds!branchwhs)).branchname & Chr(9)
            End If
            s = s & Format(DateDiff("d", ds!lastrecpt, Now), "0") & Chr(9)
            s = s & ds!onhand & Chr(9)
            s = s & ds!onorder & Chr(9)
            's = s & (ds!onorder + (ds!thiswknewpals * ds!roqty) + (ds!nextwknewpals * ds!roqty)) & Chr(9) 'jv072216
            s = s & ds!sales & Chr(9)
            s = s & ds!undiff & Chr(9)
            s = s & ds!paldiff & Chr(9)
            s = s & Format(ds!ohpct, ".000") & Chr(9)
            s = s & ds!roqty & Chr(9)
            s = s & Format(ds!pctgain, ".000") & Chr(9)
            s = s & ds!needqty & Chr(9)
            s = s & calc_bimp_status(30, ds!ohpct, ds!paldiff)          'jv053018
            'If ds!ohpct > 0 And ds!ohpct < 0.5 Then
            '    s = s & "W"
            'Else
            '    If ds!paldiff = 0 Then
            '        s = s & "B"
            '    Else
            '        If ds!paldiff > 0 Then
            '            s = s & "G"
            '        Else
            '            s = s & "Y"
            '        End If
            '    End If
            'End If
            If List1 = "All" Then                                                   'jv031616
                If ds!plantwhs = branchrec(Val(ds!branchwhs)).supplier Then         'jv031616
                    If ds!onhand <> 0 Or ds!onorder <> 0 Or ds!sales <> 0 Then Grid1.AddItem s  'jv101316
                    'Grid1.AddItem s
                ElseIf Val(ds!branchwhs) = 34 And ds!plantwhs <> branchrec(Val(ds!branchwhs)).supplier Then
                    If ds!onhand <> 0 Or ds!onorder <> 0 Or ds!sales <> 0 Then Grid1.AddItem s
                End If                                                              'jv031616
            Else
                If ds!onhand <> 0 Or ds!onorder <> 0 Or ds!sales <> 0 Then Grid1.AddItem s      'jv101316
                'Grid1.AddItem s
            End If                                                                  'jv031616
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 9)) > 0 Then                                         'jv122115
                Grid1.TextMatrix(i, 14) = Format(Val(Grid1.TextMatrix(i, 9)) * 30, "0")     'jv122115
            End If                                                                          'jv122115
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            If Grid1.TextMatrix(i, 13) = "W" Then Grid1.CellBackColor = wcolor.BackColor
            If Grid1.TextMatrix(i, 13) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
            If Grid1.TextMatrix(i, 13) = "B" Then Grid1.CellBackColor = bcolor.BackColor
            If Grid1.TextMatrix(i, 13) = "G" Then Grid1.CellBackColor = gcolor.BackColor
            Grid1.TextMatrix(i, 13) = ""
            t4 = t4 + Val(Grid1.TextMatrix(i, 4))
            t5 = t5 + Val(Grid1.TextMatrix(i, 5))
            t6 = t6 + Val(Grid1.TextMatrix(i, 6))
            t7 = t7 + Val(Grid1.TextMatrix(i, 7))
            t8 = t8 + Val(Grid1.TextMatrix(i, 8))
            t12 = t12 + Val(Grid1.TextMatrix(i, 12))
        Next i
        Grid1.Row = 1
    End If
    s = Chr(9) & "All" & Chr(9) & "Totals" & Chr(9) & Chr(9) & t4 & Chr(9)
    s = s & t5 & Chr(9) & t6 & Chr(9) & t7 & Chr(9) & t8 & Chr(9)
    s = s & Chr(9) & Chr(9) & Chr(9) & t12
    Grid1.AddItem s
    'Grid1.Cols = Grid1.Cols - 1
    s = "^Plant|^Branch|<Location|^Days|^OnHand|^OnOrder|^Sales|^UnDiff|^PalDiff|^OH%|^PalQty|^%GPP|^Need||^Days Supply"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 900
    Grid1.ColWidth(1) = 900
    Grid1.ColWidth(2) = 1900
    Grid1.ColWidth(3) = 900
    Grid1.ColWidth(4) = 900
    Grid1.ColWidth(5) = 900
    Grid1.ColWidth(6) = 900
    Grid1.ColWidth(7) = 900
    Grid1.ColWidth(8) = 900
    Grid1.ColWidth(9) = 900
    Grid1.ColWidth(10) = 900
    Grid1.ColWidth(11) = 900
    Grid1.ColWidth(12) = 900
    Grid1.ColWidth(13) = 0
    Grid1.ColWidth(14) = 1100
    Grid1.Redraw = True
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
    Dim rt As String, rf As String, rh As String
    rt = "Sales Analysis " & Combo1
    rh = "Product:  " & Me.Caption
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "seagreen": gndc(0) = Me.Grid1.BackColorFixed
    'htdc(1) = "cyan": gndc(1) = Me.rcolor.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid2.BackColor
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
    Combo1.Clear: List1.Clear
    Combo1.AddItem "All Plants": List1.AddItem "All"
    Combo1.AddItem "T10 - Brenham": List1.AddItem "T10"
    Combo1.AddItem "K10 - Broken Arrow": List1.AddItem "K10"
    Combo1.AddItem "A10 - Sylacauga": List1.AddItem "A10"
    Me.Height = whssales.Height
    Me.Top = whssales.Top
    Me.Left = whssales.Width - Me.Width
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 4)
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)  'jv121416
    If Button = 2 And Grid1.Rows > 3 Then
        If MsgBox("Sort by " & Grid1.TextMatrix(0, Grid1.Col) & "?", vbYesNo + vbQuestion, Grid1.TextMatrix(0, Grid1.Col)) = vbYes Then
            Grid1.Row = 1
            Grid1.RowSel = Grid1.Rows - 2
            Grid1.ColSel = Grid1.Col
            If Grid1.Col < 3 Then
                Grid1.Sort = 5
            Else
                If Grid1.Col = 3 Then
                    Grid1.Sort = 3
                Else
                    Grid1.Sort = 4
                End If
            End If
            Grid1.Row = 1: Grid1.RowSel = 1
            Grid1.TopRow = 1
        End If
    End If
End Sub

Private Sub Grid1_RowColChange()
    Dim i As Integer, pals As Currency
    i = Grid1.Row
    Grid1.ToolTipText = ""
    If Val(Grid1.TextMatrix(i, 10)) = 0 Then Exit Sub
    If Grid1.Col = 4 Then
        If Val(Grid1.TextMatrix(i, 4)) > 0 Then
            pals = Format(Val(Grid1.TextMatrix(i, 4)) / Val(Grid1.TextMatrix(i, 10)), "0.00")
            Grid1.ToolTipText = Grid1.TextMatrix(i, 2) & " OnHand Pallets: " & pals
        End If
    End If
    If Grid1.Col = 5 Then
        If Val(Grid1.TextMatrix(i, 5)) > 0 Then
            pals = Format(Val(Grid1.TextMatrix(i, 5)) / Val(Grid1.TextMatrix(i, 10)), "0.00")
            Grid1.ToolTipText = Grid1.TextMatrix(i, 2) & " OnOrder Pallets: " & pals
        End If
    End If
    If Grid1.Col = 6 Then
        If Val(Grid1.TextMatrix(i, 6)) > 0 Then
            pals = Format(Val(Grid1.TextMatrix(i, 6)) / Val(Grid1.TextMatrix(i, 10)), "0.00")
            Grid1.ToolTipText = Grid1.TextMatrix(i, 2) & " Pallet Sales: " & pals
        End If
    End If
End Sub

Private Sub List1_Click()
    refresh_grid
End Sub

Private Sub psku_Change()
    Me.Caption = psku & " " & skurec(Val(psku)).unit & " " & skurec(Val(psku)).desc
    refresh_grid
End Sub

