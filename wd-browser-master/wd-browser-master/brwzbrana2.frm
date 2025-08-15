VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brwzbrana2 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List57 
      Height          =   1815
      Left            =   7440
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid wgrid 
      Height          =   1575
      Left            =   2640
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   2778
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5741
      _Version        =   327680
      GridColor       =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label calledby 
      Alignment       =   2  'Center
      Caption         =   "Region"
      Height          =   255
      Left            =   7920
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label wsku 
      Caption         =   "..."
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "brwzbrana2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_wgrid()
    Dim cfile As String, f0 As String, f1 As String
    wgrid.Clear: wgrid.Rows = 1: wgrid.Cols = 2
    'cfile = "s:\wd\html\brana\whslist.csv"
    cfile = Form1.webdir & "\brana\whslist.csv"
    Open cfile For Input As #7
    Do Until EOF(7)
        Input #7, f0, f1
        wgrid.AddItem f0 & Chr(9) & f1
    Loop
    Close #7
    wgrid.FormatString = "^Whs|<Description"
    wgrid.ColWidth(0) = 1000
    wgrid.ColWidth(1) = 2000
End Sub
Private Sub branches_57()
    Dim cfile As String
    List57.Clear
    'cfile = "s:\wd\html\brana\branches.csv"
    cfile = Form1.webdir & "\brana\branches.csv"
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13
        If f1 = "777" And f13 = "52" Then List57.AddItem f0
    Loop
    Close #1
End Sub

Private Sub refresh_57()
    Dim i As Integer, k As Integer
    Dim t3 As Long, t4 As Long, t5 As Long, t6 As Long
    Dim t7 As Long, t11 As Long, mpal As Integer
    Dim tsp As Single, tw As String, f13 As String, tw2 As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String, f12 As String
    Dim cfile As String, w As String
    tw = Left(Combo1, 2)
    If calledby.Caption = "Brana" Then
        If brwzbrana.Grid1.Row < 1 Then Exit Sub
        If Val(brwzbrana.Grid1.TextMatrix(brwzbrana.Grid1.Row, 0)) = 0 Then Exit Sub
        Me.Caption = brwzbrana.Grid1.TextMatrix(brwzbrana.Grid1.Row, 0) & " " & brwzbrana.Grid1.TextMatrix(brwzbrana.Grid1.Row, 1)
    End If
    If calledby.Caption = "Plana" Then
        If brwzplana.Grid1.Row < 1 Then Exit Sub
        If Val(brwzplana.Grid1.TextMatrix(brwzplana.Grid1.Row, 0)) = 0 Then Exit Sub
        Me.Caption = brwzplana.Grid1.TextMatrix(brwzplana.Grid1.Row, 0) & " " & brwzplana.Grid1.TextMatrix(brwzplana.Grid1.Row, 1)
    End If
    Grid1.Visible = False: Grid1.Cols = 13: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t11 = 0
    'cfile = "s:\wd\html\brana\branches.csv"
    cfile = Form1.webdir & "\brana\branches.csv"
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13
        If (tw = "Al" Or f13 = "50" Or f13 = "52") Then
            'If f1 = wsku And f0 >= "000" And f0 <= "999" Then
            If f1 = wsku And Left(f0, 1) <> "P" And Left(f0, 1) <> "R" Then     'R12
                w = "...."
                For k = 0 To wgrid.Rows - 1
                    If wgrid.TextMatrix(k, 0) = f0 Then
                        w = wgrid.TextMatrix(k, 1)
                        'w = Right(w, Len(w) - 4)
                        If Left(f0, 1) = "0" Then w = Right(w, Len(w) - 4)      'R12
                        Exit For
                    End If
                Next k
                For k = 0 To List57.ListCount - 1
                    If f0 = List57.List(k) Then 'Or f0 = "052" Then
                        If Val(f7) < 0 Or f0 = "052" Or f0 = "A10" Then         'R12
                            'MsgBox f0 & " " & f7
                            s = f0 & Chr(9) & w & Chr(9) & f3 & Chr(9) & f4 & Chr(9)
                            s = s & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12
                            Grid1.AddItem s
                            t3 = t3 + Val(f4)
                            t4 = t4 + Val(f5)
                            t5 = t5 + Val(f6)
                            t6 = t6 + Val(f7)
                            t7 = t7 + Val(f8)
                            t11 = t11 + Val(f12)
                        End If
                        Exit For
                    End If
                Next k
            End If
        End If
    Loop
    Close #1
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 0
    Grid1.Sort = 5
    
    t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t11 = 0
    If tw = "52" Then
        tw2 = "P57"
    Else
        tw2 = "PPP"
    End If
    'cfile = "s:\wd\html\brana\plants.csv"
    cfile = Form1.webdir & "\brana\plants.csv"
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12
        If (tw = "Al" Or f0 = "P" & tw Or f0 = tw2) Then
            If f1 = wsku Then
                w = "...."
                For k = 0 To wgrid.Rows - 1
                    If wgrid.TextMatrix(k, 0) = f0 Then
                        w = wgrid.TextMatrix(k, 1)
                        w = Right(w, Len(w) - 4)
                        Exit For
                    End If
                Next k
                's = f0 & Chr(9) & w & Chr(9) & f3 & Chr(9) & f4 & Chr(9)
                's = s & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                's = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12
                'Grid1.AddItem s
                t3 = t3 + Val(f4)
                t4 = t4 + Val(f5)
                t5 = t5 + Val(f6)
                t6 = t6 + Val(f7)
                t7 = t7 + Val(f8)
                t11 = t11 + Val(f12)
            End If
        End If
    Loop
    Close #1
    
    t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t11 = 0
    For i = 1 To Grid1.Rows - 1
        t3 = t3 + Val(Grid1.TextMatrix(i, 3))
        t4 = t4 + Val(Grid1.TextMatrix(i, 4))
        t5 = t5 + Val(Grid1.TextMatrix(i, 5))
        t6 = t6 + Val(Grid1.TextMatrix(i, 6))
        t7 = t7 + Val(Grid1.TextMatrix(i, 7))
        t11 = t11 + Val(Grid1.TextMatrix(i, 11))
    Next i
        
    
    s = "All" & Chr(9) & "Totals" & Chr(9) & Chr(9)
    s = s & Format(t3, "#") & Chr(9)
    s = s & Format(t4, "#") & Chr(9)
    s = s & Format(t5, "#") & Chr(9)
    's = s & Format(t3 - t5, "#") & Chr(9)
    'mpal = 2
    's = s & Format((t3 - t5) / mpal, "#") & Chr(9)
    s = s & Format(t6, "#") & Chr(9)
    s = s & Format(t7, "#") & Chr(9)
    If t5 > 0 Then
        s = s & Format(t3 / t5, ".000") & Chr(9)
    Else
        s = s & Chr(9)
    End If
    s = s & Chr(9) & Chr(9)
    s = s & Format(t11, "#")
    Grid1.AddItem s
    Screen.MousePointer = 0
    Grid1.FormatString = "^Whs|<Branch|^Days|^OnHand|^OnOrd|^Sales|^UDiff|^PDiff|^OH%|^ROQty|^PG|^Need|^% Tot Sales"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 1400
    Grid1.ColWidth(2) = 500
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 600
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 600
    Grid1.ColWidth(8) = 600
    Grid1.ColWidth(9) = 600
    Grid1.ColWidth(10) = 600
    Grid1.ColWidth(11) = 600
    Grid1.ColWidth(12) = 1000
    For i = 1 To Me.Grid1.Rows - 1
        tsp = Val(Me.Grid1.TextMatrix(i, 5))
        If tsp <> 0 And t5 <> 0 Then Grid1.TextMatrix(i, 12) = Format(tsp / t5, ".000")
        Me.Grid1.Row = i: Me.Grid1.RowSel = i
        Me.Grid1.Col = 0: Me.Grid1.ColSel = 11 '10
        If Val(Me.Grid1.TextMatrix(i, 11)) > 0 Then
            Me.Grid1.CellBackColor = brwzbrana.wcolor.BackColor
        Else
            If Val(Me.Grid1.TextMatrix(i, 7)) = 0 Then
                Me.Grid1.CellBackColor = brwzbrana.bcolor.BackColor
            Else
                If Val(Me.Grid1.TextMatrix(i, 7)) > 0 Then
                    Me.Grid1.CellBackColor = brwzbrana.gcolor.BackColor
                Else
                    Me.Grid1.CellBackColor = brwzbrana.ycolor.BackColor
                End If
            End If
        End If
        Me.Grid1.FillStyle = flexFillRepeat
    Next i
    Me.Grid1.Row = 1: Me.Grid1.RowSel = 1
    Me.Grid1.Col = 0: Me.Grid1.ColSel = 11 '10
    If Val(Me.Grid1.TextMatrix(1, 11)) > 0 Then
        Me.Grid1.CellBackColor = brwzbrana.wcolor.BackColor
    Else
        If Val(Me.Grid1.TextMatrix(1, 7)) = 0 Then
            Me.Grid1.CellBackColor = brwzbrana.bcolor.BackColor
        Else
            If Val(Me.Grid1.TextMatrix(1, 7)) > 0 Then
                Me.Grid1.CellBackColor = brwzbrana.gcolor.BackColor
            Else
                Me.Grid1.CellBackColor = brwzbrana.ycolor.BackColor
            End If
        End If
    End If
    Me.Grid1.FillStyle = flexFillRepeat
    DoEvents
    i = Me.Grid1.Rows - 1
    Me.Grid1.Row = i: Me.Grid1.RowSel = i
    Me.Grid1.Col = 0: Me.Grid1.ColSel = 11 '10
    If Val(Me.Grid1.TextMatrix(i, 7)) = 0 Then
        Me.Grid1.CellBackColor = brwzbrana.bcolor.BackColor
    Else
        If Val(Me.Grid1.TextMatrix(i, 7)) > 0 Then
            Me.Grid1.CellBackColor = brwzbrana.gcolor.BackColor
        Else
            If Val(Me.Grid1.TextMatrix(i, 8)) > 0 And Val(Me.Grid1.TextMatrix(i, 8)) < 0.5 Then
                Me.Grid1.CellBackColor = brwzbrana.wcolor.BackColor
            Else
                Me.Grid1.CellBackColor = brwzbrana.ycolor.BackColor
            End If
        End If
    End If
    Me.Grid1.FillStyle = flexFillRepeat
    
    Me.Grid1.Row = 1: Me.Grid1.Col = 2
    Me.Grid1.Visible = True
End Sub

Private Sub refresh_region()
    Dim i As Integer, k As Integer
    Dim t3 As Long, t4 As Long, t5 As Long, t6 As Long
    Dim t7 As Long, t11 As Long, mpal As Integer
    Dim tsp As Single, tw As String, f13 As String, tw2 As String
    Dim cfile As String, w As String
    tw = Left(Combo1, 2)
    If Form8.Grid1.Row < 1 Then Exit Sub
    If Val(Form8.Grid1.TextMatrix(Form8.Grid1.Row, 0)) = 0 Then Exit Sub
    Me.Caption = Form8.Grid1.TextMatrix(Form8.Grid1.Row, 0) & " " & Form8.Grid1.TextMatrix(Form8.Grid1.Row, 1)
    Grid1.Visible = False: Grid1.Cols = 13: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t11 = 0
    'cfile = "s:\wd\html\brana\branches.csv"
    cfile = Form1.webdir & "\brana\branches.csv"
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13
        'If (tw = "Al" Or f13 = tw) Then
            'If f1 = wsku And f0 >= "000" And f0 <= "999" Then
            If f1 = wsku And Left(f0, 1) <> "P" And Left(f0, 1) <> "R" Then     'R12
                w = "...."
                For k = 0 To wgrid.Rows - 1
                    If wgrid.TextMatrix(k, 0) = f0 Then
                        w = wgrid.TextMatrix(k, 1)
                        'w = Right(w, Len(w) - 4)
                        If Left(f0, 1) = "0" Then w = Right(w, Len(w) - 4)     'R12
                        Exit For
                    End If
                Next k
                For k = 0 To Form1.bclist.ListCount - 1
                    If Val(f0) = Val(Form1.bclist.List(k)) Then
                    
                    'If Val(f0) = Val(Form1.bclist.List(k)) Or f0 = "T10" Or f0 = "K10" Or f0 = "A10" Then
                        'MsgBox wsku & " " & f0
                        s = f0 & Chr(9) & w & Chr(9) & f3 & Chr(9) & f4 & Chr(9)
                        s = s & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                        s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12
                        Grid1.AddItem s
                        t3 = t3 + Val(f4)
                        t4 = t4 + Val(f5)
                        t5 = t5 + Val(f6)
                        t6 = t6 + Val(f7)
                        t7 = t7 + Val(f8)
                        t11 = t11 + Val(f12)
                        Exit For
                    End If
                Next k
            End If
        'End If
    Loop
    Close #1
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 0
    Grid1.Sort = 5
    
    t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t11 = 0
    For i = 1 To Grid1.Rows - 1
        t3 = t3 + Val(Grid1.TextMatrix(i, 3))
        t4 = t4 + Val(Grid1.TextMatrix(i, 4))
        t5 = t5 + Val(Grid1.TextMatrix(i, 5))
        t6 = t6 + Val(Grid1.TextMatrix(i, 6))
        t7 = t7 + Val(Grid1.TextMatrix(i, 7))
        t11 = t11 + Val(Grid1.TextMatrix(i, 11))
    Next i

    
    s = "All" & Chr(9) & "Totals" & Chr(9) & Chr(9)
    s = s & Format(t3, "#") & Chr(9)
    s = s & Format(t4, "#") & Chr(9)
    s = s & Format(t5, "#") & Chr(9)
    's = s & Format(t3 - t5, "#") & Chr(9)
    'mpal = 2
    's = s & Format((t3 - t5) / mpal, "#") & Chr(9)
    s = s & Format(t6, "#") & Chr(9)
    s = s & Format(t7, "#") & Chr(9)
    If t5 > 0 Then
        s = s & Format(t3 / t5, ".000") & Chr(9)
    Else
        s = s & Chr(9)
    End If
    s = s & Chr(9) & Chr(9)
    s = s & Format(t11, "#")
    Grid1.AddItem s
    Screen.MousePointer = 0
    Grid1.FormatString = "^Whs|<Branch|^Days|^OnHand|^OnOrd|^Sales|^UDiff|^PDiff|^OH%|^ROQty|^PG|^Need|^% Tot Sales"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 1400
    Grid1.ColWidth(2) = 500
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 600
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 600
    Grid1.ColWidth(8) = 600
    Grid1.ColWidth(9) = 600
    Grid1.ColWidth(10) = 600
    Grid1.ColWidth(11) = 600
    Grid1.ColWidth(12) = 1000
    For i = 1 To Me.Grid1.Rows - 1
        tsp = Val(Me.Grid1.TextMatrix(i, 5))
        If tsp <> 0 And t5 <> 0 Then Grid1.TextMatrix(i, 12) = Format(tsp / t5, ".000")
        Me.Grid1.Row = i: Me.Grid1.RowSel = i
        Me.Grid1.Col = 0: Me.Grid1.ColSel = 11 '10
        If Val(Me.Grid1.TextMatrix(i, 11)) > 0 Then
            Me.Grid1.CellBackColor = brwzbrana.wcolor.BackColor
        Else
            If Val(Me.Grid1.TextMatrix(i, 7)) = 0 Then
                Me.Grid1.CellBackColor = brwzbrana.bcolor.BackColor
            Else
                If Val(Me.Grid1.TextMatrix(i, 7)) > 0 Then
                    Me.Grid1.CellBackColor = brwzbrana.gcolor.BackColor
                Else
                    Me.Grid1.CellBackColor = brwzbrana.ycolor.BackColor
                End If
            End If
        End If
        Me.Grid1.FillStyle = flexFillRepeat
    Next i
    Me.Grid1.Row = 1: Me.Grid1.RowSel = 1
    Me.Grid1.Col = 0: Me.Grid1.ColSel = 11 '10
    If Val(Me.Grid1.TextMatrix(1, 11)) > 0 Then
        Me.Grid1.CellBackColor = brwzbrana.wcolor.BackColor
    Else
        If Val(Me.Grid1.TextMatrix(1, 7)) = 0 Then
            Me.Grid1.CellBackColor = brwzbrana.bcolor.BackColor
        Else
            If Val(Me.Grid1.TextMatrix(1, 7)) > 0 Then
                Me.Grid1.CellBackColor = brwzbrana.gcolor.BackColor
            Else
                Me.Grid1.CellBackColor = brwzbrana.ycolor.BackColor
            End If
        End If
    End If
    Me.Grid1.FillStyle = flexFillRepeat
    DoEvents
    i = Me.Grid1.Rows - 1
    Me.Grid1.Row = i: Me.Grid1.RowSel = i
    Me.Grid1.Col = 0: Me.Grid1.ColSel = 11 '10
    If Val(Me.Grid1.TextMatrix(i, 7)) = 0 Then
        Me.Grid1.CellBackColor = brwzbrana.bcolor.BackColor
    Else
        If Val(Me.Grid1.TextMatrix(i, 7)) > 0 Then
            Me.Grid1.CellBackColor = brwzbrana.gcolor.BackColor
        Else
            If Val(Me.Grid1.TextMatrix(i, 8)) > 0 And Val(Me.Grid1.TextMatrix(i, 8)) < 0.5 Then
                Me.Grid1.CellBackColor = brwzbrana.wcolor.BackColor
            Else
                Me.Grid1.CellBackColor = brwzbrana.ycolor.BackColor
            End If
        End If
    End If
    Me.Grid1.FillStyle = flexFillRepeat
    
    Me.Grid1.Row = 1: Me.Grid1.Col = 2
    Me.Grid1.Visible = True
End Sub

Private Sub refresh_sku_locations()
    Dim i As Integer, k As Integer
    Dim t3 As Long, t4 As Long, t5 As Long, t6 As Long
    Dim t7 As Long, t11 As Long, mpal As Integer
    Dim tsp As Single, tw As String, f13 As String, tw2 As String
    Dim cfile As String, w As String
    tw = Left(Combo1, 2)
    If calledby.Caption = "Brana" Then
        If brwzbrana.Grid1.Row < 1 Then Exit Sub
        If Val(brwzbrana.Grid1.TextMatrix(brwzbrana.Grid1.Row, 0)) = 0 Then Exit Sub
        Me.Caption = brwzbrana.Grid1.TextMatrix(brwzbrana.Grid1.Row, 0) & " " & brwzbrana.Grid1.TextMatrix(brwzbrana.Grid1.Row, 1)
    End If
    If calledby.Caption = "Plana" Then
        If brwzplana.Grid1.Row < 1 Then Exit Sub
        If Val(brwzplana.Grid1.TextMatrix(brwzplana.Grid1.Row, 0)) = 0 Then Exit Sub
        Me.Caption = brwzplana.Grid1.TextMatrix(brwzplana.Grid1.Row, 0) & " " & brwzplana.Grid1.TextMatrix(brwzplana.Grid1.Row, 1)
    End If
    Grid1.Visible = False: Grid1.Cols = 13: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t11 = 0
    'cfile = "s:\wd\html\brana\branches.csv"
    cfile = Form1.webdir & "\brana\branches.csv"
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13
        If (tw = "Al" Or f13 = tw) Then
            'If f1 = wsku And f0 >= "000" And f0 <= "999" Then
            If f1 = wsku And Left(f0, 1) <> "R" And Left(f0, 1) <> "P" Then 'R12
                w = "...."
                For k = 0 To wgrid.Rows - 1
                    If wgrid.TextMatrix(k, 0) = f0 Then
                        w = wgrid.TextMatrix(k, 1)
                        'w = Right(w, Len(w) - 4)
                        If Left(f0, 1) = "0" Then w = Right(w, Len(w) - 4)  'R12
                        Exit For
                    End If
                Next k
                s = f0 & Chr(9) & w & Chr(9) & f3 & Chr(9) & f4 & Chr(9)
                s = s & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12
                Grid1.AddItem s
                t3 = t3 + Val(f4)
                t4 = t4 + Val(f5)
                t5 = t5 + Val(f6)
                t6 = t6 + Val(f7)
                t7 = t7 + Val(f8)
                t11 = t11 + Val(f12)
            End If
        End If
    Loop
    Close #1
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 0
    Grid1.Sort = 5
    
    t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t11 = 0
    If tw = "52" Then
        tw2 = "P57"
    Else
        tw2 = "PPP"
    End If
    'cfile = "s:\wd\html\brana\plants.csv"
    cfile = Form1.webdir & "\brana\plants.csv"
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12
        If (tw = "Al" Or f0 = "P" & tw Or f0 = tw2) Then
            If f1 = wsku Then
                w = "...."
                For k = 0 To wgrid.Rows - 1
                    If wgrid.TextMatrix(k, 0) = f0 Then
                        w = wgrid.TextMatrix(k, 1)
                        w = Right(w, Len(w) - 4)
                        Exit For
                    End If
                Next k
                's = f0 & Chr(9) & w & Chr(9) & f3 & Chr(9) & f4 & Chr(9)
                's = s & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                's = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12
                'Grid1.AddItem s
                t3 = t3 + Val(f4)
                t4 = t4 + Val(f5)
                t5 = t5 + Val(f6)
                t6 = t6 + Val(f7)
                t7 = t7 + Val(f8)
                t11 = t11 + Val(f12)
            End If
        End If
    Loop
    Close #1
    
    t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t11 = 0
    For i = 1 To Grid1.Rows - 1
        t3 = t3 + Val(Grid1.TextMatrix(i, 3))
        t4 = t4 + Val(Grid1.TextMatrix(i, 4))
        t5 = t5 + Val(Grid1.TextMatrix(i, 5))
        t6 = t6 + Val(Grid1.TextMatrix(i, 6))
        t7 = t7 + Val(Grid1.TextMatrix(i, 7))
        t11 = t11 + Val(Grid1.TextMatrix(i, 11))
    Next i
    
    
    s = "All" & Chr(9) & "Totals" & Chr(9) & Chr(9)
    s = s & Format(t3, "#") & Chr(9)
    s = s & Format(t4, "#") & Chr(9)
    s = s & Format(t5, "#") & Chr(9)
    's = s & Format(t3 - t5, "#") & Chr(9)
    'mpal = 2
    's = s & Format((t3 - t5) / mpal, "#") & Chr(9)
    s = s & Format(t6, "#") & Chr(9)
    s = s & Format(t7, "#") & Chr(9)
    If t5 > 0 Then
        s = s & Format(t3 / t5, ".000") & Chr(9)
    Else
        s = s & Chr(9)
    End If
    s = s & Chr(9) & Chr(9)
    s = s & Format(t11, "#")
    Grid1.AddItem s
    Screen.MousePointer = 0
    Grid1.FormatString = "^Whs|<Branch|^Days|^OnHand|^OnOrd|^Sales|^UDiff|^PDiff|^OH%|^ROQty|^PG|^Need|^% Tot Sales"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 1400
    Grid1.ColWidth(2) = 500
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 600
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 600
    Grid1.ColWidth(8) = 600
    Grid1.ColWidth(9) = 600
    Grid1.ColWidth(10) = 600
    Grid1.ColWidth(11) = 600
    Grid1.ColWidth(12) = 1000
    For i = 1 To Me.Grid1.Rows - 1
        tsp = Val(Me.Grid1.TextMatrix(i, 5))
        If tsp <> 0 And t5 <> 0 Then Grid1.TextMatrix(i, 12) = Format(tsp / t5, ".000")
        Me.Grid1.Row = i: Me.Grid1.RowSel = i
        Me.Grid1.Col = 0: Me.Grid1.ColSel = 11 '10
        If Val(Me.Grid1.TextMatrix(i, 11)) > 0 Then
            Me.Grid1.CellBackColor = brwzbrana.wcolor.BackColor
        Else
            If Val(Me.Grid1.TextMatrix(i, 7)) = 0 Then
                Me.Grid1.CellBackColor = brwzbrana.bcolor.BackColor
            Else
                If Val(Me.Grid1.TextMatrix(i, 7)) > 0 Then
                    Me.Grid1.CellBackColor = brwzbrana.gcolor.BackColor
                Else
                    Me.Grid1.CellBackColor = brwzbrana.ycolor.BackColor
                End If
            End If
        End If
        Me.Grid1.FillStyle = flexFillRepeat
    Next i
    Me.Grid1.Row = 1: Me.Grid1.RowSel = 1
    Me.Grid1.Col = 0: Me.Grid1.ColSel = 11 '10
    If Val(Me.Grid1.TextMatrix(1, 11)) > 0 Then
        Me.Grid1.CellBackColor = brwzbrana.wcolor.BackColor
    Else
        If Val(Me.Grid1.TextMatrix(1, 7)) = 0 Then
            Me.Grid1.CellBackColor = brwzbrana.bcolor.BackColor
        Else
            If Val(Me.Grid1.TextMatrix(1, 7)) > 0 Then
                Me.Grid1.CellBackColor = brwzbrana.gcolor.BackColor
            Else
                Me.Grid1.CellBackColor = brwzbrana.ycolor.BackColor
            End If
        End If
    End If
    Me.Grid1.FillStyle = flexFillRepeat
    DoEvents
    i = Me.Grid1.Rows - 1
    Me.Grid1.Row = i: Me.Grid1.RowSel = i
    Me.Grid1.Col = 0: Me.Grid1.ColSel = 11 '10
    If Val(Me.Grid1.TextMatrix(i, 7)) = 0 Then
        Me.Grid1.CellBackColor = brwzbrana.bcolor.BackColor
    Else
        If Val(Me.Grid1.TextMatrix(i, 7)) > 0 Then
            Me.Grid1.CellBackColor = brwzbrana.gcolor.BackColor
        Else
            If Val(Me.Grid1.TextMatrix(i, 8)) > 0 And Val(Me.Grid1.TextMatrix(i, 8)) < 0.5 Then
                Me.Grid1.CellBackColor = brwzbrana.wcolor.BackColor
            Else
                Me.Grid1.CellBackColor = brwzbrana.ycolor.BackColor
            End If
        End If
    End If
    Me.Grid1.FillStyle = flexFillRepeat
    
    Me.Grid1.Row = 1: Me.Grid1.Col = 2
    Me.Grid1.Visible = True
End Sub

Private Sub calledby_Change()
    Combo1.Clear
    If Left(calledby.Caption, 6) = "Region" Then
        Combo1.Clear
        Combo1.AddItem calledby.Caption
        Combo1.Visible = False
    Else
        If Form1.pi01.Visible Then
            Combo1.AddItem "All Branches"
            Combo1.AddItem "50 - Brenham Distribution"
            'Combo1.AddItem "51 - Broken Arrow Distribution"
            'Combo1.AddItem "52 - Sylacauga Distribution"
        End If
        If Form1.pi47.Visible Then Combo1.AddItem "51 - Broken Arrow Distribution"
        If Form1.pi52.Visible Then Combo1.AddItem "52 - Sylacauga Distribution"
    End If
    Combo1.ListIndex = 0
End Sub

Private Sub Combo1_Click()
    If calledby.BackColor <> Me.BackColor And Left(Combo1, 2) = "52" Then
        refresh_57
    Else
        Call refresh_sku_locations
    End If
End Sub

Private Sub Command1_Click()
    Dim rt As String, rh As String, rf As String
    rt = Me.Caption
    rh = Combo1
    rf = "Printed:  " & Format(Now, "m-d-yyyy  h:mm am/pm")
    Call printflexgrid(Printer, Grid1, rt, rh, rf)
End Sub

Private Sub Form_Load()
    'Combo1.Clear
    'If Left(calledby.Caption, 6) = "Region" Then
    '    Combo1.AddItem calledby.Caption
    'Else
    '    If Form1.pi01.Visible Then
    '        Combo1.AddItem "All Branches"
    '        Combo1.AddItem "50 - Brenham Distribution"
    '        'Combo1.AddItem "51 - Broken Arrow Distribution"
    '        'Combo1.AddItem "52 - Sylacauga Distribution"
    '    End If
    '    If Form1.pi47.Visible Then Combo1.AddItem "51 - Broken Arrow Distribution"
    '    If Form1.pi52.Visible Then Combo1.AddItem "52 - Sylacauga Distribution"
    'End If
    'Combo1.ListIndex = 0
    refresh_wgrid
    branches_57
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    If Me.Height > 2000 Then
        Grid1.Height = Me.Height - 855
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    brwzbrana.Check1 = 0
    brwzplana.Check1 = 0
    Form8.Check1 = 0
End Sub

Private Sub wsku_Change()
    If Left(calledby.Caption, 6) = "Region" Then
        Combo1.Visible = False
        refresh_region
        Exit Sub
    Else
        Combo1.Visible = True
    End If
    calledby.BackColor = Me.BackColor
    If calledby = "Brana" And brwzbrana.List1 = "P57" Then calledby.BackColor = brwzbrana.ycolor.BackColor
    If calledby = "Plana" And brwzplana.brcode = "507" Then calledby.BackColor = brwzbrana.ycolor.BackColor
    If calledby.BackColor <> Me.BackColor And Left(Combo1, 2) = "52" Then
        refresh_57
    Else
        Call refresh_sku_locations
    End If
End Sub

