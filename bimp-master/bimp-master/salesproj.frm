VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form salesproj 
   Caption         =   "Sales Projections"
   ClientHeight    =   11220
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   11220
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Include New Release Sales"
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
      Left            =   7440
      TabIndex        =   21
      Top             =   600
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   3840
      TabIndex        =   19
      Text            =   "Combo3"
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   720
      TabIndex        =   17
      Text            =   "Combo2"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-->"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Target"
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
      Left            =   13920
      TabIndex        =   14
      Top             =   120
      Width           =   1455
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
      Left            =   12240
      TabIndex        =   9
      Top             =   120
      Width           =   1455
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
      Left            =   720
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   120
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7455
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   13150
      _Version        =   327680
      ForeColor       =   128
      BackColorFixed  =   12648447
      BackColorSel    =   16711680
      FocusRect       =   0
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
      Left            =   10560
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   8040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Don't Include:"
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
      Left            =   2280
      TabIndex        =   20
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Unit:"
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
      TabIndex        =   18
      Top             =   600
      Width           =   615
   End
   Begin VB.Label gcolor 
      BackColor       =   &H00FFC0C0&
      Caption         =   "gcolor"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   9360
      Width           =   1095
   End
   Begin VB.Label bcolor 
      BackColor       =   &H00FFFF80&
      Caption         =   "bcolor"
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   9000
      Width           =   1095
   End
   Begin VB.Label ycolor 
      BackColor       =   &H0080FFFF&
      Caption         =   "ycolor"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Label wcolor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "wcolor"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   8280
      Width           =   1095
   End
   Begin VB.Label Label3 
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
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Days:"
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
      Left            =   7440
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label bname 
      Caption         =   "Label2"
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
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label bcode 
      Alignment       =   2  'Center
      Caption         =   "001"
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
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Branch:"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu postmenu 
      Caption         =   "Post"
      Visible         =   0   'False
      Begin VB.Menu postnwk 
         Caption         =   "Post to Next Week"
      End
      Begin VB.Menu posttwk 
         Caption         =   "Post to This Week"
         Enabled         =   0   'False
      End
      Begin VB.Menu postbrords 
         Caption         =   "Post to Branch Orders"
         Enabled         =   0   'False
      End
      Begin VB.Menu edqty 
         Caption         =   "Change Order Qty"
      End
   End
End
Attribute VB_Name = "salesproj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_plants()
    Dim ds As ADODB.Recordset, s As String, j As Integer, i As Integer
    Dim itot As Long, stot As Long, noq As Integer
    Dim a10tot As Integer, k10tot As Integer, t10tot As Integer
    Dim a10twk As Integer, k10twk As Integer, t10twk As Integer
    Dim a10nwk As Integer, k10nwk As Integer, t10nwk As Integer
    Dim a10pal As Integer, k10pal As Integer, t10pal As Integer
    Dim a10inv As Integer, k10inv As Integer, t10inv As Integer
    Dim a10ord As Integer, k10ord As Integer, t10ord As Integer
    Dim a10sal As Integer, k10sal As Integer, t10sal As Integer
    Dim limpal As Integer, c As Long
    's = Combo1 & ": " & bcode & ": " & Text1
    'MsgBox s, vbOKOnly + vbInformation, "refresh_Plants....."
    'Exit Sub
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 14: Grid1.Rows = 1
    Grid1.FixedCols = 3
    Grid1.Clear
    s = "select plantwhs,branchwhs,sku,onhand,onorder,thiswknewpals,nextwknewpals,sales,bimpstatus"
    s = s & ",plantpool,poolsched,quotapct,roqty,promoqty,discflag from bimp"              'jv082818
    's = s & " where branchwhs = '" & Format(Val(bcode), "000") & "'"
    If Combo1 = "ALL" Then
        s = s & " where plantwhs in ('A10', 'K10', 'T10')"
    Else
        s = s & " where plantwhs = '" & Combo1 & "'"
    End If
    If Combo1 = "T10" Then s = s & " and branchwhs not in ('047', '052')"
    If Combo1 = "K10" Then s = s & " and branchwhs not in ('001', '052')"
    If Combo1 = "A10" Then s = s & " and branchwhs not in ('047', '001')"
    
    If Combo2 <> "ALL" Then                                                                 'jv072216
        s = s & " and sku in (select sku from skumast where fgunit = '" & Combo2 & "')"     'jv072216
    End If                                                                                  'jv072216
    If Combo3 <> "DONE" Then                                                                'jv072216
        s = s & " and sku not in (select sku from skumast where fgunit = '" & Combo3 & "')" 'jv072216
    End If                                                                                  'jv072216
    s = s & " order by sku, plantwhs, branchwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If Command3.FontStrikethru = True Then                                              'jv081516
                noq = 0                                                                         'jv081516
            Else                                                                                'jv081516
                noq = net_order_qty(ds!plantwhs, ds!branchwhs, ds!sku)         'jv081216
                noq = noq * skurec(Val(ds!sku)).pallet                  'jv081216
                noq = noq + (groupitems_qty(ds!sku, ds!plantwhs, ds!branchwhs) * ds!roqty)             'jv081516
            End If                                                                              'jv081516
            itot = ds!onhand + ds!onorder + noq                     'jv081216
            If ds!thiswknewpals > 0 Then itot = itot + (ds!thiswknewpals * skurec(Val(ds!sku)).pallet)
            If ds!nextwknewpals > 0 Then itot = itot + (ds!nextwknewpals * skurec(Val(ds!sku)).pallet)
            otot = CLng(ds!sales * (Val(Text1) / 30))
            If Check1 = 1 And ds!promoqty > 0 Then                  'jv083116
                otot = otot + (ds!promoqty * ds!roqty)              'jv083116
            End If                                                  'jv083116
            If otot > itot Then
                s = ds!plantwhs & Chr(9)
                s = s & ds!branchwhs & "-" & branchrec(Val(ds!branchwhs)).branchname & Chr(9)
                s = s & ds!sku & Chr(9)
                If ds!discflag = "B" Then s = s & "Blocked "        'jv082818
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                s = s & ds!onhand & Chr(9)
                's = s & ds!onorder & Chr(9)
                noq = noq + ds!onorder                              'jv081216
                s = s & noq & Chr(9)                                'jv081216
                s = s & Format(ds!thiswknewpals * skurec(Val(ds!sku)).pallet, "#") & Chr(9)
                s = s & Format(ds!nextwknewpals * skurec(Val(ds!sku)).pallet, "#") & Chr(9)
                s = s & otot & Chr(9)
                s = s & Format(itot - otot, "#") & Chr(9)
                's = s & Format((itot - otot) / skurec(Val(ds!sku)).pallet, "#.00") '& Chr(9)
                s = s & Format(CInt((itot - otot) / skurec(Val(ds!sku)).pallet) * -1, "0") & Chr(9)
                limpal = (ds!plantpool * (ds!quotapct / 100)) / ds!roqty
                If ds!poolsched > 0 Then limpal = limpal + (ds!poolsched * (ds!quotapct / 100))
                j = plant_transfers(ds!plantwhs, ds!sku)                        'jv090216
                If j > 0 Then limpal = limpal + (j * (ds!quotapct / 100))       'jv090216
                If Check1 = 1 And ds!promoqty > 0 Then              'jv083116
                    limpal = limpal + ds!promoqty                   'jv083116
                End If                                              'jv083116
                s = s & limpal & Chr(9) & Chr(9)
                s = s & ds!bimpstatus
                'If limpal > 0 And CInt((itot - otot) / skurec(Val(ds!sku)).pallet) * -1 > 0 Then Grid1.AddItem s
                Grid1.AddItem s     'jv031717
            End If
            ds.MoveNext
        Loop
    End If
    a10tot = 0: k10tot = 0: t10tot = 0
    a10inv = 0: k10inv = 0: t10inv = 0
    a10ord = 0: k10ord = 0: t10ord = 0
    a10twk = 0: k10twk = 0: t10twk = 0
    a10nwk = 0: k10nwk = 0: t10nwk = 0
    a10sal = 0: k10sal = 0: t10sal = 0
    a10pal = 0: k10pal = 0: t10pal = 0
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 11)) <= 0 Then
                Grid1.TextMatrix(i, 12) = " "
            Else
                If Val(Grid1.TextMatrix(i, 10)) <= Val(Grid1.TextMatrix(i, 11)) Then
                    Grid1.TextMatrix(i, 12) = Format(Val(Grid1.TextMatrix(i, 10)), "#")
                Else
                    Grid1.TextMatrix(i, 12) = Format(Val(Grid1.TextMatrix(i, 11)), "#")
                End If
            End If
            Grid1.Row = i: Grid1.RowSel = i
            'Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            Grid1.Col = 1: Grid1.ColSel = 2
            If Grid1.TextMatrix(i, Grid1.Cols - 1) = "B" Then
                Grid1.CellBackColor = bcolor.BackColor
            Else
                If Grid1.TextMatrix(i, Grid1.Cols - 1) = "Y" Then
                    Grid1.CellBackColor = ycolor.BackColor
                Else
                    If Grid1.TextMatrix(i, Grid1.Cols - 1) = "G" Then
                        Grid1.CellBackColor = gcolor.BackColor
                    Else
                        Grid1.CellBackColor = wcolor.BackColor
                    End If
                End If
            End If
            'Grid1.CellBackColor = c
            'If c = Grid1.BackColorFixed Then
            '    c = Grid1.BackColor
            'Else
            '    c = Grid1.BackColorFixed
            'End If
            j = Val(Grid1.TextMatrix(i, 2))
            If Grid1.TextMatrix(i, 0) = "A10" Then
                a10tot = a10tot + Val(Grid1.TextMatrix(i, 12))
                a10inv = a10inv + Val(Grid1.TextMatrix(i, 4)) / skurec(j).pallet
                a10ord = a10ord + Val(Grid1.TextMatrix(i, 5)) / skurec(j).pallet
                a10twk = a10twk + Val(Grid1.TextMatrix(i, 6)) / skurec(j).pallet
                a10nwk = a10nwk + Val(Grid1.TextMatrix(i, 7)) / skurec(j).pallet
                a10sal = a10sal + Val(Grid1.TextMatrix(i, 8)) / skurec(j).pallet
                a10pal = a10pal + Val(Grid1.TextMatrix(i, 10))
            End If
            If Grid1.TextMatrix(i, 0) = "K10" Then
                k10tot = k10tot + Val(Grid1.TextMatrix(i, 12))
                k10inv = k10inv + Val(Grid1.TextMatrix(i, 4)) / skurec(j).pallet
                k10ord = k10ord + Val(Grid1.TextMatrix(i, 5)) / skurec(j).pallet
                k10twk = k10twk + Val(Grid1.TextMatrix(i, 6)) / skurec(j).pallet
                k10nwk = k10nwk + Val(Grid1.TextMatrix(i, 7)) / skurec(j).pallet
                k10sal = k10sal + Val(Grid1.TextMatrix(i, 8)) / skurec(j).pallet
                k10pal = k10pal + Val(Grid1.TextMatrix(i, 10))
            End If
            If Grid1.TextMatrix(i, 0) = "T10" Then
                t10tot = t10tot + Val(Grid1.TextMatrix(i, 12))
                t10inv = t10inv + Val(Grid1.TextMatrix(i, 4)) / skurec(j).pallet
                t10ord = t10ord + Val(Grid1.TextMatrix(i, 5)) / skurec(j).pallet
                t10twk = t10twk + Val(Grid1.TextMatrix(i, 6)) / skurec(j).pallet
                t10nwk = t10nwk + Val(Grid1.TextMatrix(i, 7)) / skurec(j).pallet
                t10sal = t10sal + Val(Grid1.TextMatrix(i, 8)) / skurec(j).pallet
                t10pal = t10pal + Val(Grid1.TextMatrix(i, 10))
            End If
        Next i
        Grid1.Row = 1: Grid1.Col = 3
    End If
    If a10tot > 0 Then
        i = CInt(a10tot / 34)
        s = "A10" & Chr(9) & Chr(9) & i & " Loads" & Chr(9) & "Total Pallets" & Chr(9) & a10inv & Chr(9) & a10ord & Chr(9)
        s = s & a10twk & Chr(9) & a10nwk & Chr(9) & a10sal & Chr(9) & Chr(9) & a10pal & Chr(9) & Chr(9) & a10tot
        Grid1.AddItem s
    End If
    If k10tot > 0 Then
        i = CInt(k10tot / 34)
        s = "K10" & Chr(9) & Chr(9) & i & " Loads" & Chr(9) & "Total Pallets" & Chr(9) & k10inv & Chr(9) & k10ord & Chr(9)
        s = s & k10twk & Chr(9) & k10nwk & Chr(9) & k10sal & Chr(9) & Chr(9) & k10pal & Chr(9) & Chr(9) & k10tot
        Grid1.AddItem s
    End If
    If t10tot > 0 Then
        i = CInt(t10tot / 34)
        s = "T10" & Chr(9) & Chr(9) & i & " Loads" & Chr(9) & "Total Pallets" & Chr(9) & t10inv & Chr(9) & t10ord & Chr(9)
        s = s & t10twk & Chr(9) & t10nwk & Chr(9) & t10sal & Chr(9) & Chr(9) & t10pal & Chr(9) & Chr(9) & t10tot
        Grid1.AddItem s
    End If
    
    'c = Grid1.BackColor ' wcolor.BackColor
    c = wcolor.BackColor
    If Grid1.Rows > 1 Then
        s = Grid1.TextMatrix(1, 2)
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 2) <> s Then
                'If c = Grid1.BackColorFixed Then ' wcolor.BackColor Then
                '    c = Grid1.BackColor ' bcolor.BackColor
                'Else
                '    c = Grid1.BackColorFixed ' wcolor.BackColor
                'End If
                If c = wcolor.BackColor Then
                    c = bcolor.BackColor
                Else
                    c = wcolor.BackColor
                End If
                s = Grid1.TextMatrix(i, 2)
            End If
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 3: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellBackColor = c
        Next i
        Grid1.Row = 1
    End If
    
    s = "^Plant|<Branch|^SKU|<Product|^OnHand|^OnOrder|^ThisWeek|^NextWeek|^Sales|^UnDiff|^Pallets|^Limit|^Order"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 2200
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 3000
    Grid1.ColWidth(4) = 1100
    Grid1.ColWidth(5) = 1100
    Grid1.ColWidth(6) = 1100
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1100
    Grid1.ColWidth(9) = 1100
    Grid1.ColWidth(10) = 1100
    Grid1.ColWidth(11) = 1100
    Grid1.ColWidth(12) = 1100
    Grid1.ColWidth(13) = 0
    Grid1.Redraw = True
    Screen.MousePointer = 0

End Sub


Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, s As String, j As Integer, i As Integer
    Dim itot As Long, stot As Long, noq As Integer
    Dim a10tot As Integer, k10tot As Integer, t10tot As Integer
    Dim a10twk As Integer, k10twk As Integer, t10twk As Integer
    Dim a10nwk As Integer, k10nwk As Integer, t10nwk As Integer
    Dim a10pal As Integer, k10pal As Integer, t10pal As Integer
    Dim a10inv As Integer, k10inv As Integer, t10inv As Integer
    Dim a10ord As Integer, k10ord As Integer, t10ord As Integer
    Dim a10sal As Integer, k10sal As Integer, t10sal As Integer
    Dim limpal As Integer, c As Long
    If bcode = "ALL" Then
        refresh_plants
        Exit Sub
    End If
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 13: Grid1.Rows = 1
    Grid1.FixedCols = 3
    Grid1.Clear
    s = "select plantwhs,sku,onhand,onorder,thiswknewpals,nextwknewpals,sales,bimpstatus"
    s = s & ",plantpool,poolsched,quotapct,roqty,promoqty,discflag from bimp"              'jv082818
    s = s & " where branchwhs = '" & Format(Val(bcode), "000") & "'"
    If Combo1 = "ALL" Then
        s = s & " and plantwhs in ('A10', 'K10', 'T10')"
    Else
        s = s & " and plantwhs = '" & Combo1 & "'"
    End If
    If Combo2 <> "ALL" Then                                                                 'jv072216
        If Combo2 = "12PK CUP" Then
            s = s & " and sku in (SELECT sku FROM skumast WHERE fgunit = '12PK' AND (ISNULL(proddesc, '') LIKE '%cup%' OR ISNULL(fgdesc, '') LIKE '%cup%'))"
        ElseIf Combo2 = "12PK SAND" Then
            s = s & " and sku in (SELECT sku FROM skumast WHERE fgunit = '12PK' AND (ISNULL(proddesc, '') LIKE '%SAND%' OR ISNULL(fgdesc, '') LIKE '%SAND%'))"
        ElseIf Combo2 = "12PK BAR" Then
            s = s & " and sku in (SELECT sku FROM skumast WHERE fgunit = '12PK' AND NOT (ISNULL(proddesc, '') LIKE '%cup%' OR ISNULL(fgdesc, '') LIKE '%cup%'"
            s = s & " OR ISNULL(proddesc, '') LIKE '%SAND%' OR ISNULL(fgdesc, '') LIKE '%SAND%'))"
        ElseIf Combo2 = "BULK BAR" Then
            s = s & " and sku in (SELECT sku FROM skumast WHERE fgunit = 'BULK' AND (ISNULL(proddesc, '') LIKE '%BAR' OR ISNULL(fgdesc, '') LIKE '%BAR'))"
        ElseIf Combo2 = "BULK SAND" Then
            s = s & " and sku in (SELECT sku FROM skumast WHERE fgunit = 'BULK' AND (ISNULL(proddesc, '') LIKE '%SAND%' OR ISNULL(fgdesc, '') LIKE '%SAND%'))"
        Else
            s = s & " and sku in (select sku from skumast where fgunit = '" & Combo2 & "')"
        End If
    End If                                                                                  'jv072216
    If Combo3 <> "DONE" Then                                                                'jv072216
        If Combo3 = "12PK CUP" Then
            s = s & " and sku not in (SELECT sku FROM skumast WHERE fgunit = '12PK' AND (ISNULL(proddesc, '') LIKE '%cup%' OR ISNULL(fgdesc, '') LIKE '%cup%'))"
        ElseIf Combo3 = "12PK SAND" Then
            s = s & " and sku not in (SELECT sku FROM skumast WHERE fgunit = '12PK' AND (ISNULL(proddesc, '') LIKE '%SAND%' OR ISNULL(fgdesc, '') LIKE '%SAND%'))"
        ElseIf Combo3 = "12PK BAR" Then
            s = s & " and sku not in (SELECT sku FROM skumast WHERE fgunit = '12PK' AND NOT (ISNULL(proddesc, '') LIKE '%cup%' OR ISNULL(fgdesc, '') LIKE '%cup%'"
            s = s & " OR ISNULL(proddesc, '') LIKE '%SAND%' OR ISNULL(fgdesc, '') LIKE '%SAND%'))"
        ElseIf Combo3 = "BULK BAR" Then
            s = s & " and sku not in (SELECT sku FROM skumast WHERE fgunit = 'BULK' AND (ISNULL(proddesc, '') LIKE '%BAR' OR ISNULL(fgdesc, '') LIKE '%BAR'))"
        ElseIf Combo3 = "BULK SAND" Then
            s = s & " and sku not in (SELECT sku FROM skumast WHERE fgunit = 'BULK' AND (ISNULL(proddesc, '') LIKE '%SAND%' OR ISNULL(fgdesc, '') LIKE '%SAND%'))"
        Else
            s = s & " and sku not in (select sku from skumast where fgunit = '" & Combo3 & "')"
        End If
    End If                                                                                  'jv072216
    s = s & " order by sku, plantwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If Command3.FontStrikethru = True Then                                              'jv081516
                noq = 0                                                                         'jv081516
            Else                                                                                'jv081516
                noq = net_order_qty(ds!plantwhs, bcode, ds!sku)         'jv081216
                noq = noq * skurec(Val(ds!sku)).pallet                  'jv081216
                noq = noq + (groupitems_qty(ds!sku, ds!plantwhs, bcode) * ds!roqty)             'jv081516
            End If                                                                              'jv081516
            itot = ds!onhand + ds!onorder + noq                     'jv081216
            If ds!thiswknewpals > 0 Then itot = itot + (ds!thiswknewpals * skurec(Val(ds!sku)).pallet)
            If ds!nextwknewpals > 0 Then itot = itot + (ds!nextwknewpals * skurec(Val(ds!sku)).pallet)
            otot = CLng(ds!sales * (Val(Text1) / 30))
            If Check1 = 1 And ds!promoqty > 0 Then                  'jv083116
                otot = otot + (ds!promoqty * ds!roqty)              'jv083116
            End If                                                  'jv083116
            If otot > itot Then
                s = ds!plantwhs & Chr(9)
                s = s & ds!sku & Chr(9)
                If ds!discflag = "B" Then s = s & "Blocked "        'jv082818
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                s = s & ds!onhand & Chr(9)
                's = s & ds!onorder & Chr(9)
                noq = noq + ds!onorder                              'jv081216
                s = s & noq & Chr(9)                                'jv081216
                s = s & Format(ds!thiswknewpals * skurec(Val(ds!sku)).pallet, "#") & Chr(9)
                s = s & Format(ds!nextwknewpals * skurec(Val(ds!sku)).pallet, "#") & Chr(9)
                s = s & otot & Chr(9)
                s = s & Format(itot - otot, "#") & Chr(9)
                's = s & Format((itot - otot) / skurec(Val(ds!sku)).pallet, "#.00") '& Chr(9)
                s = s & Format(CInt((itot - otot) / skurec(Val(ds!sku)).pallet) * -1, "0") & Chr(9)
                limpal = (ds!plantpool * (ds!quotapct / 100)) / ds!roqty
                If ds!poolsched > 0 Then limpal = limpal + (ds!poolsched * (ds!quotapct / 100))
                j = plant_transfers(ds!plantwhs, ds!sku)                        'jv090216
                If j > 0 Then limpal = limpal + (j * (ds!quotapct / 100))       'jv090216
                If Check1 = 1 And ds!promoqty > 0 Then              'jv083116
                    limpal = limpal + ds!promoqty                   'jv083116
                End If                                              'jv083116
                s = s & limpal & Chr(9) & Chr(9)
                s = s & ds!bimpstatus
                Grid1.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    a10tot = 0: k10tot = 0: t10tot = 0
    a10inv = 0: k10inv = 0: t10inv = 0
    a10ord = 0: k10ord = 0: t10ord = 0
    a10twk = 0: k10twk = 0: t10twk = 0
    a10nwk = 0: k10nwk = 0: t10nwk = 0
    a10sal = 0: k10sal = 0: t10sal = 0
    a10pal = 0: k10pal = 0: t10pal = 0
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 10)) <= 0 Then
                Grid1.TextMatrix(i, 11) = " "
            Else
                If Val(Grid1.TextMatrix(i, 9)) <= Val(Grid1.TextMatrix(i, 10)) Then
                    Grid1.TextMatrix(i, 11) = Format(Val(Grid1.TextMatrix(i, 9)), "#")
                Else
                    Grid1.TextMatrix(i, 11) = Format(Val(Grid1.TextMatrix(i, 10)), "#")
                End If
            End If
            Grid1.Row = i: Grid1.RowSel = i
            'Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            Grid1.Col = 1: Grid1.ColSel = 2
            If Grid1.TextMatrix(i, Grid1.Cols - 1) = "B" Then
                Grid1.CellBackColor = bcolor.BackColor
            Else
                If Grid1.TextMatrix(i, Grid1.Cols - 1) = "Y" Then
                    Grid1.CellBackColor = ycolor.BackColor
                Else
                    If Grid1.TextMatrix(i, Grid1.Cols - 1) = "G" Then
                        Grid1.CellBackColor = gcolor.BackColor
                    Else
                        Grid1.CellBackColor = wcolor.BackColor
                    End If
                End If
            End If
            'Grid1.CellBackColor = c
            'If c = Grid1.BackColorFixed Then
            '    c = Grid1.BackColor
            'Else
            '    c = Grid1.BackColorFixed
            'End If
            j = Val(Grid1.TextMatrix(i, 1))
            If Grid1.TextMatrix(i, 0) = "A10" Then
                a10tot = a10tot + Val(Grid1.TextMatrix(i, 11))
                a10inv = a10inv + Val(Grid1.TextMatrix(i, 3)) / skurec(j).pallet
                a10ord = a10ord + Val(Grid1.TextMatrix(i, 4)) / skurec(j).pallet
                a10twk = a10twk + Val(Grid1.TextMatrix(i, 5)) / skurec(j).pallet
                a10nwk = a10nwk + Val(Grid1.TextMatrix(i, 6)) / skurec(j).pallet
                a10sal = a10sal + Val(Grid1.TextMatrix(i, 7)) / skurec(j).pallet
                a10pal = a10pal + Val(Grid1.TextMatrix(i, 9))
            End If
            If Grid1.TextMatrix(i, 0) = "K10" Then
                k10tot = k10tot + Val(Grid1.TextMatrix(i, 11))
                k10inv = k10inv + Val(Grid1.TextMatrix(i, 3)) / skurec(j).pallet
                k10ord = k10ord + Val(Grid1.TextMatrix(i, 4)) / skurec(j).pallet
                k10twk = k10twk + Val(Grid1.TextMatrix(i, 5)) / skurec(j).pallet
                k10nwk = k10nwk + Val(Grid1.TextMatrix(i, 6)) / skurec(j).pallet
                k10sal = k10sal + Val(Grid1.TextMatrix(i, 7)) / skurec(j).pallet
                k10pal = k10pal + Val(Grid1.TextMatrix(i, 9))
            End If
            If Grid1.TextMatrix(i, 0) = "T10" Then
                t10tot = t10tot + Val(Grid1.TextMatrix(i, 11))
                t10inv = t10inv + Val(Grid1.TextMatrix(i, 3)) / skurec(j).pallet
                t10ord = t10ord + Val(Grid1.TextMatrix(i, 4)) / skurec(j).pallet
                t10twk = t10twk + Val(Grid1.TextMatrix(i, 5)) / skurec(j).pallet
                t10nwk = t10nwk + Val(Grid1.TextMatrix(i, 6)) / skurec(j).pallet
                t10sal = t10sal + Val(Grid1.TextMatrix(i, 7)) / skurec(j).pallet
                t10pal = t10pal + Val(Grid1.TextMatrix(i, 9))
            End If
        Next i
        Grid1.Row = 1: Grid1.Col = 3
    End If
    If a10tot > 0 Then
        i = CInt(a10tot / 34)
        s = "A10" & Chr(9) & i & " Loads" & Chr(9) & "Total Pallets" & Chr(9) & a10inv & Chr(9) & a10ord & Chr(9)
        s = s & a10twk & Chr(9) & a10nwk & Chr(9) & a10sal & Chr(9) & Chr(9) & a10pal & Chr(9) & Chr(9) & a10tot
        Grid1.AddItem s
    End If
    If k10tot > 0 Then
        i = CInt(k10tot / 34)
        s = "K10" & Chr(9) & i & " Loads" & Chr(9) & "Total Pallets" & Chr(9) & k10inv & Chr(9) & k10ord & Chr(9)
        s = s & k10twk & Chr(9) & k10nwk & Chr(9) & k10sal & Chr(9) & Chr(9) & k10pal & Chr(9) & Chr(9) & k10tot
        Grid1.AddItem s
    End If
    If t10tot > 0 Then
        i = CInt(t10tot / 34)
        s = "T10" & Chr(9) & i & " Loads" & Chr(9) & "Total Pallets" & Chr(9) & t10inv & Chr(9) & t10ord & Chr(9)
        s = s & t10twk & Chr(9) & t10nwk & Chr(9) & t10sal & Chr(9) & Chr(9) & t10pal & Chr(9) & Chr(9) & t10tot
        Grid1.AddItem s
    End If
    
    c = Grid1.BackColor ' wcolor.BackColor
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 3: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellBackColor = c
            If c = Grid1.BackColorFixed Then ' wcolor.BackColor Then
                c = Grid1.BackColor ' bcolor.BackColor
            Else
                c = Grid1.BackColorFixed ' wcolor.BackColor
            End If
        Next i
        Grid1.Row = 1
    End If
    
    s = "^Plant|^SKU|<Product|^OnHand|^OnOrder|^ThisWeek|^NextWeek|^Sales|^UnDiff|^Pallets|^Limit|^Order"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 1100
    Grid1.ColWidth(4) = 1100
    Grid1.ColWidth(5) = 1100
    Grid1.ColWidth(6) = 1100
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1100
    Grid1.ColWidth(9) = 1100
    Grid1.ColWidth(10) = 1100
    Grid1.ColWidth(11) = 1100
    Grid1.ColWidth(12) = 0
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub bcode_Change()
    Dim ds As ADODB.Recordset, s As String
    If bcode = "ALL" Then
        bname.Caption = "ALL Branches"
        'Combo1.ListIndex = 0
    Else
        bname.Caption = branchrec(Val(bcode)).branchname
        s = "select listdisplay from valuelists where listname = 'branchplants'"
        s = s & " and listreturn = '" & bcode & "'"
        'MsgBox s
        Set ds = wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            For i = 0 To Combo1.ListCount - 1
                If Combo1.List(i) = ds!listdisplay Then
                    Combo1.ListIndex = i
                    Exit For
                End If
            Next i
        'Else
        '   refresh_grid
        End If
        ds.Close
    End If
    refresh_grid
End Sub

Private Sub Check1_Click()
    refresh_grid
End Sub

Private Sub Combo1_Click()
    If Val(Text1) > 0 And Val(bcode) > 0 Then refresh_grid
End Sub

Private Sub Combo2_Click()
    If Val(Text1) > 0 And Val(bcode) > 0 Then refresh_grid
End Sub

Private Sub Combo3_Click()
    If Val(Text1) > 0 And Val(bcode) > 0 Then refresh_grid
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub Command2_Click()
    Dim rt As String, rh As String, rf As String
    rt = "Sales Projections - " & Text1 & " Days"
    rh = Combo1 & " " & bcode & " " & bname
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "lightcyan": gndc(0) = Me.bcolor.BackColor
    htdc(1) = "yellow": gndc(1) = Me.ycolor.BackColor
    ''htdc(2) = "lightgrey": gndc(2) = Me.wcolor.BackColor
    htdc(2) = "white": gndc(2) = Me.wcolor.BackColor
    htdc(3) = "lightgrey": gndc(3) = Me.gcolor.BackColor
    Grid1.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "lightyellow", "white")
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Grid1.Redraw = True: Grid1.Row = 1
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "lightyellow", "white")
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Grid1.Redraw = True: Grid1.Row = 1
        Exit Sub
    End If
End Sub

Private Sub Command3_Click()
    Dim i As Integer, t As Integer, r As Integer
    t = InputBox("Pallet Order Qty:", "Pallet Qty Target....", 34)
    If Len(t) = 0 Then Exit Sub
    r = Val(Grid1.TextMatrix(Grid1.Rows - 1, 11))
    
    Command3.FontStrikethru = True
    DoEvents
    
    If r > t Then
        For i = Val(Text1) To 0 Step -1
            Text1 = i
            refresh_grid
            r = Val(Grid1.TextMatrix(Grid1.Rows - 1, 11))
            If r < t Then
                Text1 = i + 1
                refresh_grid
                Exit For
            End If
        Next i
    Else
        For i = Val(Text1) To 90
            Text1 = i
            refresh_grid
            r = Val(Grid1.TextMatrix(Grid1.Rows - 1, 11))
            If r > t Then
                Text1 = i - 1
                refresh_grid
                Exit For
            End If
        Next i
    End If
    
    Command3.FontStrikethru = False
    refresh_grid
End Sub

Private Sub Command4_Click()
    If Val(Text1) > 1 Then
        Text1 = Val(Text1) - 1
        refresh_grid
    End If
End Sub

Private Sub Command5_Click()
    Text1 = Val(Text1) + 1
    refresh_grid
End Sub

Private Sub edqty_Click()
    Dim pqty As String, t As Integer
    If bcode = "ALL" Then
        pqty = Grid1.TextMatrix(Grid1.Row, 12)
        pqty = InputBox("Qty", "Qty", pqty)
        If Len(pqty) = 0 Then Exit Sub
        Grid1.TextMatrix(Grid1.Row, 12) = pqty
        t = 0
        For i = 1 To Grid1.Rows - 2
            t = t + Val(Grid1.TextMatrix(i, 12))
        Next i
        Grid1.TextMatrix(Grid1.Rows - 1, 12) = t
    Else
        pqty = Grid1.TextMatrix(Grid1.Row, 11)
        pqty = InputBox("Qty", "Qty", pqty)
        If Len(pqty) = 0 Then Exit Sub
        Grid1.TextMatrix(Grid1.Row, 11) = pqty
        t = 0
        For i = 1 To Grid1.Rows - 2
            t = t + Val(Grid1.TextMatrix(i, 11))
        Next i
        Grid1.TextMatrix(Grid1.Rows - 1, 11) = t
    End If
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    If bimpbanner.Command1.Visible = False Then
        postnwk.Enabled = False
        posttwk.Enabled = False
        postbrords.Enabled = False
        edqty.Enabled = False
    End If
    Combo1.Clear
    Combo1.AddItem "T10"
    Combo1.AddItem "K10"
    Combo1.AddItem "A10"
    If bcode <> "ALL" Then Combo1.AddItem "ALL"
    'MsgBox branchturnover.Combo1
    If branchturnover.Combo1 = "K10" Then
        Combo1.ListIndex = 1
    Else
        If branchturnover.Combo1 = "A10" Then
            Combo1.ListIndex = 2
        Else
            Combo1.ListIndex = 0
        End If
    End If
    Combo2.AddItem "ALL"
    Combo2.AddItem "1/2"
    Combo2.AddItem "PT"
    Combo2.AddItem "3GAL"
    Combo2.AddItem "CUP"
    Combo2.AddItem "12PK CUP"
    Combo2.AddItem "12PK BAR"
    Combo2.AddItem "12PK SAND"
    Combo2.AddItem "BULK BAR"
    Combo2.AddItem "BULK SAND"
    Combo2.AddItem "QT"                     'jv062317
    Combo2.ListIndex = 0
    Combo3.AddItem "NONE"
    Combo3.AddItem "1/2"
    Combo3.AddItem "PT"
    Combo3.AddItem "3GAL"
    Combo3.AddItem "CUP"
    Combo3.AddItem "12PK CUP"
    Combo3.AddItem "12PK BAR"
    Combo3.AddItem "12PK SAND"
    Combo3.AddItem "BULK BAR"
    Combo3.AddItem "BULK SAND"
    Combo3.AddItem "QT"                     'jv062317
    Combo3.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Text1.Height * 5)
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu postmenu
End Sub

Private Sub postbrords_Click()
    Dim psku As String, pqty As String, i As Integer, pplant As Integer, pbranch As Integer, pdate As String
    Dim z As Long
    If Combo1 = "A10" Then pplant = 52
    If Combo1 = "K10" Then pplant = 51
    If Combo1 = "T10" Then pplant = 50
    pbranch = Val(bcode.Caption)
    pdate = Format(DateAdd("d", 1, Now), "M-dd-yyyy")
    pdate = InputBox("Order Date:", "Order Date....", pdate)
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then Exit Sub
    For i = 1 To Grid1.Rows - 2
        If bcode = "ALL" Then
            pqty = Val(Grid1.TextMatrix(i, 12))
        Else
            pqty = Val(Grid1.TextMatrix(i, 11))
        End If
        If pqty > 0 Then
            If bcode = "ALL" Then
                pbranch = Val(Left(Grid1.TextMatrix(i, 1), 3))
                psku = Grid1.TextMatrix(i, 2)
            Else
                pbranch = Val(bcode.Caption)
                psku = Grid1.TextMatrix(i, 1)
            End If
            z = wd_seq("brorders")
            s = "Insert into brorders (id, plant, branch, account, sku, orddate, ordqty, gprqty"
            s = s & ", netqty, altflag, partqty) Values (" & z
            s = s & ", " & pplant
            s = s & ", " & pbranch
            s = s & ", '......'"
            s = s & ", '" & psku & "'"
            s = s & ", '" & pdate & "'"
            s = s & ", " & pqty
            s = s & ", 0, 0, 'N', 0)"
            'MsgBox s
            wdb.Execute s
        End If
    Next i
End Sub

Private Sub postnwk_Click()
    Dim psku As String, pqty As String, i As Integer
    For i = 1 To Grid1.Rows - 2
        If bcode = "ALL" Then
            pqty = Val(Grid1.TextMatrix(i, 12))
        Else
            pqty = Val(Grid1.TextMatrix(i, 11))
        End If
        If pqty > 0 Then
            If bcode = "ALL" Then
                psku = Grid1.TextMatrix(i, 2)
            Else
                psku = Grid1.TextMatrix(i, 1)
            End If
            s = "Update bimp set nextwknewpals = " & pqty
            s = s & " Where plantwhs = '" & Combo1 & "'"
            If bcode = "ALL" Then
                s = s & " And branchwhs = '" & Left(Grid1.TextMatrix(i, 1), 3) & "'"
            Else
                s = s & " And branchwhs = '" & bcode & "'"
            End If
            s = s & " and sku = '" & psku & "'"
            'MsgBox s
            wdb.Execute s
        End If
    Next i
End Sub

Private Sub posttwk_Click()
    Dim psku As String, pqty As String, i As Integer
    For i = 1 To Grid1.Rows - 2
        If bcode = "ALL" Then
            pqty = Val(Grid1.TextMatrix(i, 12))
        Else
            pqty = Val(Grid1.TextMatrix(i, 11))
        End If
        If pqty > 0 Then
            If bcode = "ALL" Then
                psku = Grid1.TextMatrix(i, 2)
            Else
                psku = Grid1.TextMatrix(i, 1)
            End If
            s = "Update bimp set thiswknewpals = " & pqty
            s = s & " Where plantwhs = '" & Combo1 & "'"
            If bcode = "ALL" Then
                s = s & " And branchwhs = '" & Left(Grid1.TextMatrix(i, 1), 3) & "'"
            Else
                s = s & " And branchwhs = '" & bcode & "'"
            End If
            s = s & " and sku = '" & psku & "'"
            'MsgBox s
            wdb.Execute s
        End If
    Next i
End Sub

