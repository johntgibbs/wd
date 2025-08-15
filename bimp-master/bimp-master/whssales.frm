VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form whssales 
   Caption         =   "Sales Analysis"
   ClientHeight    =   12105
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15045
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12105
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   375
      Left            =   14160
      TabIndex        =   22
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000C0&
      Height          =   3150
      Left            =   0
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   7200
      TabIndex        =   19
      Top             =   11520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3480
      TabIndex        =   17
      Text            =   "Combo2"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "WMS Inventory"
      Height          =   375
      Left            =   14160
      TabIndex        =   14
      Top             =   120
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   11520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   11400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   360
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   13361
      _Version        =   327680
      BackColorFixed  =   14737632
      FocusRect       =   0
      GridLines       =   2
   End
   Begin VB.Label menulabel 
      Caption         =   "Reports"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "..."
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Supplier:"
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      Height          =   255
      Left            =   5760
      TabIndex        =   15
      Top             =   840
      Width           =   8175
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Surplus"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11880
      TabIndex        =   12
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label gpct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".000"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11880
      TabIndex        =   11
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label bpct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".000"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9840
      TabIndex        =   10
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Month Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9840
      TabIndex        =   9
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label ypct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".000"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   8
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "< Month"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7800
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label wpct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".000"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "< 2 Weeks"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   5
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "..."
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Warehouse:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "whssales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub plantot(plantcode As String)
    Dim query As String, i As Integer, ss As ADODB.Recordset
    Dim db As ADODB.Connection, ds As ADODB.Recordset, sqlx As String
    Dim tc As Integer, mpal As Integer, tord As Long
    Dim psource As Integer
    Dim pstat As String                             'jv053018
    'On Error GoTo vberror
    tc = Check1.Value
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 12: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    'sqlx = "select sku,plantpool,sum(onhand),sum(onorder),sum(sales) from bimp"
    sqlx = "select sku,plantpool,sum(onhand),sum(onorder),sum(sales),sum(thiswknewpals)"    'jv072216
    sqlx = sqlx & ",sum(nextwknewpals) from bimp"                                           'jv072216
    sqlx = sqlx & " where plantwhs = '" & plantcode & "'"
    sqlx = sqlx & " group by sku, plantpool"
    sqlx = sqlx & " having plantpool > 0 or sum(onhand) > 0 or sum(sales) > 0 order by sku" 'jv012517
    'MsgBox sqlx
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = Val(ds(0))
            sqlx = ds(0) & Chr(9)                                                   'SKU
            If skurec(i).sku <> ds(0) Then
                sqlx = sqlx & "...." & Chr(9)
                mpal = 0: psource = 0
            Else
                sqlx = sqlx & skurec(i).unit & " " & skurec(i).desc & Chr(9)        'Product
                mpal = skurec(i).pallet: psource = skurec(i).psrc
            End If
                
            
            sqlx = sqlx & Format(ds(1), "#") & Chr(9)                               'Plant Units
            sqlx = sqlx & Format(ds(2), "#") & Chr(9)                               'Branch Units
            'sqlx = sqlx & Format(ds(3), "#") & Chr(9)                               'Branch Orders
            tord = ds(3)                                'jv072216
            If IsNull(ds(5)) = False Then tord = tord + (ds(5) * mpal)                'jv072216
            If IsNull(ds(6)) = False Then tord = tord + (ds(6) * mpal)                 'jv072216
            sqlx = sqlx & Format(tord, "#") & Chr(9)                               'Branch Orders
            sqlx = sqlx & Format(ds(4), "#") & Chr(9)                               'Sales
            sqlx = sqlx & Format((ds(1) + ds(2)) - ds(4), "#") & Chr(9)
            sqlx = sqlx & Format(((ds(1) + ds(2)) - ds(4)) / mpal, "#") & Chr(9)    '(Plant Units + Branch Units) - Sales
            If ds(4) > 0 Then
                sqlx = sqlx & Format((ds(1) + ds(2)) / ds(4), ".000") & Chr(9)
            Else
                sqlx = sqlx & Chr(9)
            End If
            If mpal <> 0 Then 'And (ds(2) >= mpal Or ds(4) >= mpal) Then
                'If plantcode = 55 Then
                '    If psource = 4 Then Grid1.AddItem sqlx
                'Else
                    Grid1.AddItem sqlx
                'End If
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
    'Grid1.FormatString = "^SKU|<Product|^Days|^OnHand|^OnOrd|^Sales|^UDiff|^PDiff|^OH%|^ROQty|^PG|^Need"
    'Grid1.FormatString = "^SKU|<Product||^OnHand||^Sales|^UDiff|^PDiff|^OH%|||"
    Grid1.FormatString = "^SKU|<Product|^Plant Units|^Branch Units|^Branch Orders|^Sales|^UnitDiff|^PalDiff|^OH%|^Days Supply||"
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1400 '6 '00
    Grid1.ColWidth(3) = 1400
    Grid1.ColWidth(4) = 1400 '6 '00
    Grid1.ColWidth(5) = 1400
    Grid1.ColWidth(6) = 1400
    Grid1.ColWidth(7) = 1400
    Grid1.ColWidth(8) = 1400
    Grid1.ColWidth(9) = 1200 '6 '00
    Grid1.ColWidth(10) = 0 '6 '00
    Grid1.ColWidth(11) = 0 '6 '00
    If Check1.Value = 1 Then Check1.Value = 0
    For i = 1 To Grid1.Rows - 1
        pstat = calc_bimp_status(30, Val(Grid1.TextMatrix(i, 8)), Val(Grid1.TextMatrix(i, 7)))  'jv053018
        Grid1.TextMatrix(i, 9) = Format(Val(Grid1.TextMatrix(i, 8)) * 30, "0")
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = 11 '10
        If pstat = "W" Then                                                 'jv053018
        'If Val(Grid1.TextMatrix(i, 8)) < 0.5 And Val(Grid1.TextMatrix(i, 8)) > 0 Then
            Grid1.CellBackColor = wcolor.BackColor
            wpct = Val(wpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
        Else
            If pstat = "B" Then                                             'jv053018
            'If Val(Grid1.TextMatrix(i, 7)) = 0 Then
                Grid1.CellBackColor = bcolor.BackColor
                bpct = Val(bpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
            Else
            
                If pstat = "G" Then                                         'jv053018
                'If Val(Grid1.TextMatrix(i, 7)) > 0 Then
                    Grid1.CellBackColor = gcolor.BackColor
                    gpct = Val(gpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
                Else
                    Grid1.CellBackColor = ycolor.BackColor
                    ypct = Val(ypct) + Abs(Val(Grid1.TextMatrix(i, 6)))
                End If
            End If
        End If
        Grid1.FillStyle = flexFillRepeat
    Next i
    If Grid1.Rows > 1 Then
        pstat = calc_bimp_status(30, Val(Grid1.TextMatrix(1, 8)), Val(Grid1.TextMatrix(1, 7)))  'jv053018
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 0: Grid1.ColSel = 11 '10
        If pstat = "W" Then                                                 'jv053018
        'If Val(Grid1.TextMatrix(1, 8)) < 0.5 And Val(Grid1.TextMatrix(1, 8)) > 0 Then
        ''If Val(Grid1.TextMatrix(1, 11)) > 0 Then
            Grid1.CellBackColor = wcolor.BackColor
        Else
            If pstat = "B" Then                                             'jv053018
            'If Val(Grid1.TextMatrix(1, 7)) = 0 Then
                Grid1.CellBackColor = bcolor.BackColor
            Else
                If pstat = "G" Then                                         'jv053018
                'If Val(Grid1.TextMatrix(1, 7)) > 0 Then
                    Grid1.CellBackColor = gcolor.BackColor
                Else
                    Grid1.CellBackColor = ycolor.BackColor
                End If
            End If
        End If
        Grid1.FillStyle = flexFillRepeat
        Check1.Value = tc
        Grid1.Row = 1: Grid1.Col = 3
    End If
    Grid1.Redraw = True
    stot = Val(wpct) + Val(ypct) + Val(bpct) + Val(gpct)
    If stot > 0 And Val(wpct) > 0 Then
        wpct = Format(Val(wpct) / stot, ".000")
    Else
        wpct = "..."
    End If
    If stot > 0 And Val(ypct) > 0 Then
        ypct = Format(Val(ypct) / stot, ".000")
    Else
        ypct = "..."
    End If
    If stot > 0 And Val(bpct) > 0 Then
        bpct = Format(Val(bpct) / stot, ".000")
    Else
        bpct = "..."
    End If
    If stot > 0 And Val(gpct) > 0 Then
        gpct = Format(Val(gpct) / stot, ".000")
    Else
        gpct = "..."
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.Description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "plantot", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " plantot - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_branch()
    Dim query As String, i As Integer, ss As ADODB.Recordset, b As bimprec, p As bimprec
    Dim db As ADODB.Connection, ds As ADODB.Recordset, sqlx As String
    Dim tc As Integer, stot As Long
    Dim pstat As String                                             'jv053018
    'On Error GoTo vberror
    Label3.Caption = bimp_status_time                                                        'jv022316
    If Label3.Caption > " " Then Label3.Caption = "Last R12 import @ " & Label3.Caption      'jv022316
    ypct = "": wpct = "": gpct = "": bpct = ""
    Command1.Visible = False
    If Combo1 = "T10" Or Combo1 = "A10" Or Combo1 = "K10" Then
        Command1.Visible = True
        Call plantot(Combo1)
        Exit Sub
    End If
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 14: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    sqlx = "select sku,lastrecpt,"
    sqlx = sqlx & "onhand,onorder,sales,"
    sqlx = sqlx & "undiff,paldiff,ohpct,roqty,"
    sqlx = sqlx & "pctgain,needqty,plantwhs"
    sqlx = sqlx & ",thiswknewpals,nextwknewpals"                             'jv072216
    sqlx = sqlx & " from bimp"
    sqlx = sqlx & " where branchwhs = '" & Combo1 & "'"
    If Combo2 <> "ALL" Then sqlx = sqlx & " and plantwhs = '" & Combo2 & "'"
    sqlx = sqlx & " and plantwhs <> 'DRY'"                          'jv020516
    sqlx = sqlx & " order by sku"
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            b.sku = ds!sku
            b.lastrcpt = ds!lastrecpt
            b.onhand = ds!onhand
            b.onorder = ds!onorder
            b.sales = ds!sales
            b.undiff = ds!undiff
            b.paldiff = ds!paldiff
            b.ohpct = ds!ohpct
            b.roqty = ds!roqty
            b.pctgain = ds!pctgain
            b.needqty = ds!needqty
            b.plantwhs = ds!plantwhs
            If IsNull(ds!thiswknewpals) Then
                b.thiswknewpals = 0
            Else
                b.thiswknewpals = ds!thiswknewpals
            End If
            If IsNull(ds!nextwknewpals) Then
                b.nextwknewpals = 0
            Else
                b.nextwknewpals = ds!nextwknewpals
            End If
            p = calc_bimp_line(b)
            
            sqlx = p.sku & Chr(9)
            i = Val(p.sku)
            If skurec(i).sku <> p.sku Then
                sqlx = sqlx & "..." & Chr(9)
            Else
                sqlx = sqlx & skurec(i).unit & " " & skurec(i).desc & Chr(9)
            End If
            sqlx = sqlx & DateDiff("d", p.lastrcpt, Now) & Chr(9)
            sqlx = sqlx & Format(p.onhand, "#") & Chr(9)
            sqlx = sqlx & Format(p.onorder, "#") & Chr(9)
            sqlx = sqlx & Format(p.sales, "#") & Chr(9)
            sqlx = sqlx & Format(p.undiff, "#") & Chr(9)
            sqlx = sqlx & Format(p.paldiff, "#") & Chr(9)
            sqlx = sqlx & Format(p.ohpct, ".000") & Chr(9)
            sqlx = sqlx & Format(p.roqty, "#") & Chr(9)
            sqlx = sqlx & Format(p.pctgain, ".000") & Chr(9)
            sqlx = sqlx & Format(p.needqty, "#") & Chr(9)
            sqlx = sqlx & ds(11)
            If Val(age) > 0 Then
                If (DateDiff("d", p.lastrcpt, Now) >= Val(age) Or IsDate(p.lastrcpt) = False) And p.onhand > 0 Then Grid1.AddItem sqlx
            Else
                Grid1.AddItem sqlx
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
    Grid1.FormatString = "^SKU|<Product|^Days|^OnHand|^OnOrder|^Sales|^UnitDiff|^PalletDiff|^OH%|^PalSize|^%Gain|^Need|^Source|^Days Supply"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 700
    Grid1.ColWidth(3) = 1100
    Grid1.ColWidth(4) = 1100
    Grid1.ColWidth(5) = 1100
    Grid1.ColWidth(6) = 1100
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1100
    Grid1.ColWidth(9) = 1100
    Grid1.ColWidth(10) = 1100
    Grid1.ColWidth(11) = 1100
    Grid1.ColWidth(12) = 1100
    If Check1.Value = 1 Then Check1.Value = 0
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 8)) > 0 Then                                         'jv041116
            Grid1.TextMatrix(i, 13) = Format(Val(Grid1.TextMatrix(i, 8)) * 30, "0")     'jv041116
        End If                                                                          'jv041116
        'If Val(Grid1.TextMatrix(Grid1.Row, 8)) > 0 Then                                                 'jv122115
        '    Grid1.TextMatrix(Grid1.Row, 13) = Format(Val(Grid1.TextMatrix(Grid1.Row, 8)) * 30, "0")     'jv122115
        'End If                                                                                          'jv122115
        pstat = calc_bimp_status(30, Val(Grid1.TextMatrix(i, 8)), Val(Grid1.TextMatrix(i, 7)))  'jv053018
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = 13 '10
        If pstat = "W" Then                                                     'jv053018
        'If Val(Grid1.TextMatrix(i, 11)) > 0 Then                                'need
            Grid1.CellBackColor = wcolor.BackColor
            wpct = Val(wpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
        Else
            If pstat = "B" Then                                                 'jv053018
            'If Val(Grid1.TextMatrix(i, 7)) = 0 Then                             'pallet diff
                Grid1.CellBackColor = bcolor.BackColor
                bpct = Val(bpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
            Else
                If pstat = "G" Then                                             'jv053018
                'If Val(Grid1.TextMatrix(i, 7)) > 0 Then                         'pallet diff
                    Grid1.CellBackColor = gcolor.BackColor
                    gpct = Val(gpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
                Else
                    Grid1.CellBackColor = ycolor.BackColor
                    ypct = Val(ypct) + Abs(Val(Grid1.TextMatrix(i, 6)))
                End If
            End If
        End If
        Grid1.FillStyle = flexFillRepeat
    Next i
    If Grid1.Rows > 1 Then
        pstat = calc_bimp_status(30, Val(Grid1.TextMatrix(1, 8)), Val(Grid1.TextMatrix(1, 7)))  'jv053018
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 0: Grid1.ColSel = 11 '10
        If pstat = "W" Then                                                     'jv053018
        'If Val(Grid1.TextMatrix(1, 11)) > 0 Then
            Grid1.CellBackColor = wcolor.BackColor
        Else
            If pstat = "B" Then                                                 'jv053018
            'If Val(Grid1.TextMatrix(1, 7)) = 0 Then
                Grid1.CellBackColor = bcolor.BackColor
            Else
                If pstat = "G" Then                                             'jv053018
                'If Val(Grid1.TextMatrix(1, 7)) > 0 Then
                    Grid1.CellBackColor = gcolor.BackColor
                Else
                    Grid1.CellBackColor = ycolor.BackColor
                End If
            End If
        End If
        Grid1.FillStyle = flexFillRepeat
        Check1.Value = tc
        Grid1.Row = 1: Grid1.Col = 3
    End If
    Grid1.Redraw = True
    stot = Val(wpct) + Val(ypct) + Val(bpct) + Val(gpct)
    If stot > 0 And Val(wpct) > 0 Then
        wpct = Format(Val(wpct) / stot, ".000")
    Else
        wpct = "..."
    End If
    If stot > 0 And Val(ypct) > 0 Then
        ypct = Format(Val(ypct) / stot, ".000")
    Else
        ypct = "..."
    End If
    If stot > 0 And Val(bpct) > 0 Then
        bpct = Format(Val(bpct) / stot, ".000")
    Else
        bpct = "..."
    End If
    If stot > 0 And Val(gpct) > 0 Then
        gpct = Format(Val(gpct) / stot, ".000")
    Else
        gpct = "..."
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.Description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_branch", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_branch - Error Number: " & eno
        End
    End If

End Sub

Private Sub refresh_grid()
    Dim query As String, i As Integer, ss As ADODB.Recordset
    Dim db As ADODB.Connection, ds As ADODB.Recordset, sqlx As String
    Dim tc As Integer, stot As Long
    Dim pstat As String                                                     'jv053018
    'On Error GoTo vberror
    Label3.Caption = bimp_status_time                                                        'jv022316
    If Label3.Caption > " " Then Label3.Caption = "Last R12 import @ " & Label3.Caption      'jv022316
    ypct = "": wpct = "": gpct = "": bpct = ""
    Command1.Visible = False
    If Combo1 = "T10" Or Combo1 = "A10" Or Combo1 = "K10" Then
        Command1.Visible = True
        Call plantot(Combo1)
        Exit Sub
    End If
    'If Left(List1, 1) = "P" Then
    '    Call plantot(Right(List1, 2))
    '    explants.Enabled = True
    '    Exit Sub
    'End If
    'explants.Enabled = False
    'tc = Check1.Value
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 14: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    sqlx = "select sku,lastrecpt,"
    sqlx = sqlx & "onhand,onorder,sales,"
    sqlx = sqlx & "undiff,paldiff,ohpct,roqty,"
    sqlx = sqlx & "pctgain,needqty,plantwhs"
    sqlx = sqlx & ",thiswknewpals,nextwknewpals"                             'jv072216
    sqlx = sqlx & " from bimp"
    sqlx = sqlx & " where branchwhs = '" & Combo1 & "'"
    If Combo2 <> "ALL" Then                                                 'jv090216
        'sqlx = sqlx & " and plantwhs in ('VENDOR', '" & Combo2 & "')"       'jv090216
        sqlx = sqlx & " and plantwhs ='" & Combo2 & "'"       'jv090216
    End If                                                                  'jv090216
    sqlx = sqlx & " and plantwhs <> 'DRY'"                          'jv020516
    sqlx = sqlx & " order by sku"
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!sku & Chr(9)
            i = Val(ds!sku)
            If skurec(i).sku <> ds(0) Then
                sqlx = sqlx & "..." & Chr(9)
            Else
                sqlx = sqlx & skurec(i).unit & " " & skurec(i).desc & Chr(9)
            End If
            sqlx = sqlx & DateDiff("d", ds(1), Now) & Chr(9)
            sqlx = sqlx & Format(ds(2), "#") & Chr(9)
            sqlx = sqlx & Format(ds(3), "#") & Chr(9)
            'sqlx = sqlx & Format(ds(3) + (ds(12) * ds(8)) + (ds(13) * ds(8)), "#") & Chr(9)         'jv072216
            sqlx = sqlx & Format(ds(4), "#") & Chr(9)
            sqlx = sqlx & Format(ds(5), "#") & Chr(9)
            sqlx = sqlx & Format(ds(6), "#") & Chr(9)
            sqlx = sqlx & Format(ds(7), ".000") & Chr(9)
            sqlx = sqlx & Format(ds(8), "#") & Chr(9)
            sqlx = sqlx & Format(ds(9), ".000") & Chr(9)
            sqlx = sqlx & Format(ds(10), "#") & Chr(9)
            sqlx = sqlx & ds(11)
            If Val(age) > 0 Then
                If (DateDiff("d", ds(1), Now) >= Val(age) Or IsDate(ds(1)) = False) And ds(2) > 0 Then Grid1.AddItem sqlx
            Else
                If ds!onhand <> 0 Or ds!onorder <> 0 Or ds!sales <> 0 Then Grid1.AddItem sqlx   'jv101316
                'Grid1.AddItem sqlx
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
    Grid1.FormatString = "^SKU|<Product|^Days|^OnHand|^OnOrder|^Sales|^UnitDiff|^PalletDiff|^OH%|^PalSize|^%Gain|^Need|^Source|^Days Supply"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 700
    Grid1.ColWidth(3) = 1100
    Grid1.ColWidth(4) = 1100
    Grid1.ColWidth(5) = 1100
    Grid1.ColWidth(6) = 1100
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1100
    Grid1.ColWidth(9) = 1100
    Grid1.ColWidth(10) = 1100
    Grid1.ColWidth(11) = 1100
    Grid1.ColWidth(12) = 1100
    If Check1.Value = 1 Then Check1.Value = 0
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 8)) > 0 Then                                         'jv041116
            Grid1.TextMatrix(i, 13) = Format(Val(Grid1.TextMatrix(i, 8)) * 30, "0")     'jv041116
        End If                                                                          'jv041116
        'If Val(Grid1.TextMatrix(Grid1.Row, 8)) > 0 Then                                                 'jv122115
        '    Grid1.TextMatrix(Grid1.Row, 13) = Format(Val(Grid1.TextMatrix(Grid1.Row, 8)) * 30, "0")     'jv122115
        'End If                                                                                          'jv122115
        pstat = calc_bimp_status(30, Val(Grid1.TextMatrix(i, 8)), Val(Grid1.TextMatrix(i, 7)))  'jv053018
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = 13 '10
        If pstat = "W" Then                                                     'jv053018
        'If Val(Grid1.TextMatrix(i, 11)) > 0 Then                                'need
            Grid1.CellBackColor = wcolor.BackColor
            wpct = Val(wpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
        Else
            If pstat = "B" Then                                                 'jv053018
            'If Val(Grid1.TextMatrix(i, 7)) = 0 Then                             'pallet diff
                Grid1.CellBackColor = bcolor.BackColor
                bpct = Val(bpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
            Else
                If pstat = "G" Then                                             'jv053018
                'If Val(Grid1.TextMatrix(i, 7)) > 0 Then                         'pallet diff
                    Grid1.CellBackColor = gcolor.BackColor
                    gpct = Val(gpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
                Else
                    Grid1.CellBackColor = ycolor.BackColor
                    ypct = Val(ypct) + Abs(Val(Grid1.TextMatrix(i, 6)))
                End If
            End If
        End If
        Grid1.FillStyle = flexFillRepeat
    Next i
    If Grid1.Rows > 1 Then
        pstat = calc_bimp_status(30, Val(Grid1.TextMatrix(1, 8)), Val(Grid1.TextMatrix(1, 7)))  'jv053018
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 0: Grid1.ColSel = 11 '10
        If pstat = "W" Then                                                     'jv053018
        'If Val(Grid1.TextMatrix(1, 11)) > 0 Then
            Grid1.CellBackColor = wcolor.BackColor
        Else
            If pstat = "B" Then                                                 'jv053018
            'If Val(Grid1.TextMatrix(1, 7)) = 0 Then
                Grid1.CellBackColor = bcolor.BackColor
            Else
                If pstat = "G" Then                                             'jv053018
                'If Val(Grid1.TextMatrix(1, 7)) > 0 Then
                    Grid1.CellBackColor = gcolor.BackColor
                Else
                    Grid1.CellBackColor = ycolor.BackColor
                End If
            End If
        End If
        Grid1.FillStyle = flexFillRepeat
        Check1.Value = tc
        Grid1.Row = 1: Grid1.Col = 3
    End If
    Grid1.Redraw = True
    stot = Val(wpct) + Val(ypct) + Val(bpct) + Val(gpct)
    If stot > 0 And Val(wpct) > 0 Then
        wpct = Format(Val(wpct) / stot, ".000")
    Else
        wpct = "..."
    End If
    If stot > 0 And Val(ypct) > 0 Then
        ypct = Format(Val(ypct) / stot, ".000")
    Else
        ypct = "..."
    End If
    If stot > 0 And Val(bpct) > 0 Then
        bpct = Format(Val(bpct) / stot, ".000")
    Else
        bpct = "..."
    End If
    If stot > 0 And Val(gpct) > 0 Then
        gpct = Format(Val(gpct) / stot, ".000")
    Else
        gpct = "..."
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.Description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "download_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " download_click - Error Number: " & eno
        End
    End If
End Sub

Sub refresh_vlists()
    Combo1.Clear: List1.Clear
    For i = 1 To 99
        If branchrec(i).oraloc > " " Then
            'Combo1.AddItem Format(branchrec(i).branchno, "000") & "-" & branchrec(i).branchname
            Combo1.AddItem Format(branchrec(i).branchno, "000")
            If i = 1 Then                                       'jv090216
                List1.AddItem "Brenham Sales"                   'jv090216
            Else                                                'jv090216
                If i = 47 Then                                  'jv090216
                    List1.AddItem "Tulsa Sales"                 'jv090216
                Else                                            'jv090216
                    If i = 52 Then                              'jv090216
                        List1.AddItem "Sylacauga Sales"         'jv090216
                    Else                                        'jv090216
                        List1.AddItem branchrec(i).branchname
                    End If                                      'jv090216
                End If                                          'jv090216
            End If                                              'jv090216
        End If                                                  'jv090216
    Next i                                                      'jv090216
    For i = 50 To 52
        Combo1.AddItem plantrec(i).orawhs
        List1.AddItem plantrec(i).plantname & " Plant"          'jv090216
        Combo2.AddItem plantrec(i).orawhs                       'jv090216
        List2.AddItem plantrec(i).plantname                     'jv090216
    Next i
    Combo2.AddItem "ALL"                                        'jv090216
    List2.AddItem "All Plants"                                  'jv090216
End Sub

Private Sub batscheds_Click()
    plantmgrsched.Show
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    Label2.Caption = List1
    DoEvents
    refresh_grid
    'refresh_branch                     'jv072216
End Sub

Private Sub Combo2_Click()                                  'jv090216
    List2.ListIndex = Combo2.ListIndex
    Label5.Caption = List2
    refresh_grid
End Sub

Private Sub Command1_Click()
    plantmgrsched.unitsoh = Val(Grid1.TextMatrix(Grid1.Row, 2)) + Val(Grid1.TextMatrix(Grid1.Row, 3))
    plantmgrsched.sales30 = Val(Grid1.TextMatrix(Grid1.Row, 5))
    plantmgrsched.daysupply = Val(Grid1.TextMatrix(Grid1.Row, 9))
    plantmgrsched.plantcode = Combo1
    plantmgrsched.prodcode = Grid1.TextMatrix(Grid1.Row, 0)
    plantmgrsched.prodname = Grid1.TextMatrix(Grid1.Row, 1)
    plantmgrsched.Show
End Sub

Private Sub Command2_Click()
    Dim rt As String, rf As String, rh As String
    rt = "Sales Analysis"
    rh = "Warehouse:  " & Combo1 & " - " & Label2.Caption
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
    refresh_vlists
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    Combo1.ListIndex = 0
    List3.Clear
    List3.AddItem "Plant SKU Orders"
    List3.AddItem "Trailer Status"
    List3.AddItem "Branch Pallet Turnover"
    List3.AddItem "Branch Capacities"
    List3.AddItem "Branch Pallet Order Totals"
    List3.AddItem "E-O-P Units"
    List3.AddItem "Low Stock Report"
    List3.AddItem "Out of Stock Report"
    List3.AddItem "Shipments Not Received"
    List3.AddItem "Daily SKU Pallet Shipments"
    List3.AddItem "OverStocked Branch SKUs"                 'jv032317
    List3.AddItem "Daily Pallet Route Loads"                'jv042417
    List3.AddItem "Hub Inventories"                         'jv112717
    List3.AddItem "Transfer Stations"                       'jv061818
    'List3.AddItem "Branch Out of Stock Report"
    List3.AddItem "Home Page"
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 5)
End Sub

'Private Sub Grid1_Click()
'    Dim pdaze As Integer, i As Integer
'    pdaze = 35 'bimp_sales_days
'    MsgBox avg_sales(Combo1, Grid1.TextMatrix(Grid1.Row, 0), pdaze)
'End Sub

Private Sub Grid1_DblClick()
    skusales.psku = Grid1.TextMatrix(Grid1.Row, 0)
    skusales.Show
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)  'jv121416
    List3.Visible = False
    If Button = 2 And Grid1.Rows > 3 Then
        If MsgBox("Sort by " & Grid1.TextMatrix(0, Grid1.Col) & "?", vbYesNo + vbQuestion, Grid1.TextMatrix(0, Grid1.Col)) = vbYes Then
            Grid1.Row = 1
            Grid1.RowSel = Grid1.Rows - 1
            Grid1.ColSel = Grid1.Col
            If Grid1.Col < 2 Then
                Grid1.Sort = 5
            Else
                'If Grid1.Col = 2 Then
                If Grid1.TextMatrix(0, Grid1.Col) = "Days" Then
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
    Dim i As Integer, pals As Currency, psku As Integer
    i = Grid1.Row
    Grid1.ToolTipText = ""
    If Grid1.TextMatrix(0, 2) = "Plant Units" Then
        psku = Val(Grid1.TextMatrix(i, 0))
        If psku = 0 Then Exit Sub
        If skurec(psku).pallet = 0 Then Exit Sub
        If Grid1.Col = 2 Then
            If Val(Grid1.TextMatrix(i, 2)) > 0 Then
                pals = Format(Val(Grid1.TextMatrix(i, 2)) / skurec(psku).pallet, "0.00")
                Grid1.ToolTipText = Grid1.TextMatrix(i, 1) & " Plant Pallets: " & pals
            End If
        End If
        If Grid1.Col = 3 Then
            If Val(Grid1.TextMatrix(i, 3)) > 0 Then
                pals = Format(Val(Grid1.TextMatrix(i, 3)) / skurec(psku).pallet, "0.00")
                Grid1.ToolTipText = Grid1.TextMatrix(i, 1) & " Branch Pallets: " & pals
            End If
        End If
        If Grid1.Col = 4 Then
            If Val(Grid1.TextMatrix(i, 4)) > 0 Then
                pals = Format(Val(Grid1.TextMatrix(i, 4)) / skurec(psku).pallet, "0.00")
                Grid1.ToolTipText = Grid1.TextMatrix(i, 1) & " Branch Order Pallets: " & pals
            End If
        End If
        If Grid1.Col = 5 Then
            If Val(Grid1.TextMatrix(i, 5)) > 0 Then
                pals = Format(Val(Grid1.TextMatrix(i, 5)) / skurec(psku).pallet, "0.00")
                Grid1.ToolTipText = Grid1.TextMatrix(i, 1) & " Pallet Sales: " & pals
            End If
        End If
                
                
    Else
        If Val(Grid1.TextMatrix(i, 9)) = 0 Then Exit Sub
        If Grid1.Col = 3 Then
            If Val(Grid1.TextMatrix(i, 3)) > 0 Then
                pals = Format(Val(Grid1.TextMatrix(i, 3)) / Val(Grid1.TextMatrix(i, 9)), "0.00")
                Grid1.ToolTipText = Grid1.TextMatrix(i, 1) & " OnHand Pallets: " & pals
            End If
        End If
        If Grid1.Col = 4 Then
            If Val(Grid1.TextMatrix(i, 4)) > 0 Then
                pals = Format(Val(Grid1.TextMatrix(i, 4)) / Val(Grid1.TextMatrix(i, 9)), "0.00")
                Grid1.ToolTipText = Grid1.TextMatrix(i, 1) & " OnOrder Pallets: " & pals
            End If
        End If
        If Grid1.Col = 5 Then
            If Val(Grid1.TextMatrix(i, 5)) > 0 Then
                pals = Format(Val(Grid1.TextMatrix(i, 5)) / Val(Grid1.TextMatrix(i, 9)), "0.00")
                Grid1.ToolTipText = Grid1.TextMatrix(i, 1) & " Pallet Sales: " & pals
            End If
        End If
    End If
End Sub

Private Sub Label2_Change()                                         'jv090216
    Dim pplant As String, i As Integer
    Combo2.Visible = False: Label5.Visible = False: Label4.Visible = False
    If Combo1 = "A10" Or Combo1 = "K10" Or Combo1 = "T10" Then
        Combo2.ListIndex = Combo2.ListCount - 1
        Exit Sub
    End If
    Combo2.Visible = True: Label5.Visible = True: Label4.Visible = True
    If Combo1 = "001" Then
        Combo2.ListIndex = 0
        Exit Sub
    End If
    If Combo1 = "047" Then
        Combo2.ListIndex = 1
        Exit Sub
    End If
    If Combo1 = "052" Then
        Combo2.ListIndex = 2
        Exit Sub
    End If
    
    For i = 1 To 99
        If branchrec(i).oraloc = Combo1 Then
            pplant = branchrec(i).supplier
            Exit For
        End If
    Next i
    For i = 0 To Combo2.ListCount - 1
        If Combo2.List(i) = pplant Then
            Combo2.ListIndex = i
            Exit For
        End If
    Next i
    
End Sub

Private Sub List3_Click()
    If List3.Visible = False Then Exit Sub
    List3.Visible = False
    DoEvents
    If List3 = "Plant SKU Orders" Then plantorders.Show
    If List3 = "Trailer Status" Then trlstatus.Show
    If List3 = "Branch Pallet Turnover" Then branchturnover.Show
    If List3 = "Branch Capacities" Then branchcaps.Show
    If List3 = "Branch Pallet Order Totals" Then branchpalship.Show
    If List3 = "E-O-P Units" Then bimpeop.Show
    'If List3 = "Low Stock Report" Then bimplowstock.Show
    If List3 = "Low Stock Report" Then branchouts.Show                      'jv032118
    If List3 = "Out of Stock Report" Then bimpoutstock.Show
    If List3 = "Shipments Not Received" Then shipnot.Show
    If List3 = "Daily SKU Pallet Shipments" Then dailysku.Show
    If List3 = "OverStocked Branch SKUs" Then whseover.Show                 'jv032317
    If List3 = "Home Page" Then browserpage.Show
    If List3 = "Daily Pallet Route Loads" Then dailypaltots.Show
    If List3 = "Hub Inventories" Then bimphubrpt.Show                       'jv112717
    If List3 = "Transfer Stations" Then tstations.Show                      'jv061818
    'If List3 = "Branch Out of Stock Report" Then branchouts.Show
End Sub

Private Sub menulabel_Click()
    List3.Visible = True
End Sub

Private Sub menulabel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'List3.Visible = True
End Sub

