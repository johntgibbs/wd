VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form branchcaps 
   Caption         =   "Branch Storage Capacities"
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11430
   LinkTopic       =   "Form15"
   ScaleHeight     =   6570
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   4471
      _Version        =   327680
      ForeColor       =   128
      BackColorFixed  =   12648447
      BackColorSel    =   32768
      FocusRect       =   0
   End
   Begin VB.Label gcolor 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   5160
      Width           =   4095
   End
   Begin VB.Menu prtgrid 
      Caption         =   "&Print"
   End
   Begin VB.Menu refgrid 
      Caption         =   "&Refresh"
   End
End
Attribute VB_Name = "branchcaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fetch_branch_pallets()
    Dim i As Integer, s As String, ds As ADODB.Recordset
    Dim ppal As String, ptot As Long, ptype As String
    Dim t3gal As Long, ttray As Long, itot As Long, isku As Integer
    Dim psku As String
    For i = 1 To Grid1.Rows - 2
        ptot = 0: t3gal = 0: ttray = 0: psku = " "
        s = "select sku, onhand from bimp where branchwhs = '" & Grid1.TextMatrix(i, 0) & "'"
        s = s & " and plantwhs <> 'DRY' order by sku"                          'jv020516
        Set ds = wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ds!sku <> psku Then
                    isku = Val(ds!sku)
                    If skurec(isku).sku = ds!sku Then
                        ppal = skurec(isku).pallet
                        ptype = Left(skurec(isku).unit & ".", 1)
                    Else
                        ppal = 0: ptype = " "
                    End If
                    If ds!onhand > 0 And ppal > 0 Then
                        If ptype = "3" Then
                            t3gal = t3gal + ds!onhand
                        Else
                            If ptype = "T" Then
                                ttray = ttray + ds!onhand
                            Else
                                'ptot = ptot + Int((ds!onhand / ppal) + 0.999)
                                ptot = ptot + pallet_space(ds!sku, ds!onhand)       'jv022616
                            End If
                        End If
                    End If
                    psku = ds!sku
                End If
                ds.MoveNext
            Loop
            ds.Close
        End If
        'ptot = ptot + Int((t3gal / 60) + 0.999)
        ptot = ptot + pallet_space("507", t3gal)
        ptot = ptot + Int((ttray / 132) + 0.999)
        Grid1.TextMatrix(i, 5) = ptot
        DoEvents
    Next i
End Sub

Private Sub Form_Load()
    Dim ds As ADODB.Recordset, sqlx As String
    Dim tc As Long, tu As Long, ti As Long
    Dim bc As String, bu As String, i As Integer
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12: Grid1.FixedCols = 3
    tc = 0: tu = 0: ti = 0
    sqlx = "select branch,gemmsid,branchname,modem,fax from branches where gemmsid > '0'"
    sqlx = sqlx & " and modem > '0' and fax > '0'"                          'jv020316
    sqlx = sqlx & " and branch not in (15,16)"
    sqlx = sqlx & " order by branch"
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!gemmsid & Chr(9)              'Whs
            sqlx = sqlx & ds!branch & Chr(9)        'Branch
            sqlx = sqlx & ds!branchname & Chr(9)    'Location
            sqlx = sqlx & ds!modem & Chr(9)         'Tot Cap
            sqlx = sqlx & ds!fax & Chr(9)           'Usable
            sqlx = sqlx & Chr(9)                    'Product
            sqlx = sqlx & Chr(9)                    '%Usable
            bc = ds!modem & " "
            bu = ds!fax & " "
            sqlx = sqlx & Format(Val(bc) - Val(bu), "#") & Chr(9) 'Ing Cap
            sqlx = sqlx & Chr(9)                    'Ings
            sqlx = sqlx & Chr(9)                    'FGoods
            sqlx = sqlx & Chr(9)                    'Diff
            
            tc = tc + Val(bc)
            tu = tu + Val(bu)
            ti = tc - tu
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.AddItem Chr(9) & Chr(9) & "Branch Totals" & Chr(9) & tc & Chr(9) & tu & Chr(9) & Chr(9) & Chr(9) & ti
    Grid1.FormatString = "^Whs|^Branch|<Location|^Tot Cap|^Usable|^Product|^%Usable|^Ing Cap|^Oracle Ings|^WD Storage|^Diff|^%Total"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 900
    Grid1.ColWidth(2) = 2200
    Grid1.ColWidth(3) = 1100
    Grid1.ColWidth(4) = 1100
    Grid1.ColWidth(5) = 1100
    Grid1.ColWidth(6) = 1100
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1400
    Grid1.ColWidth(9) = 1400
    Grid1.ColWidth(10) = 1000
    Grid1.ColWidth(11) = 1000
    Grid1.FillStyle = flexFillRepeat
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 7: Grid1.ColSel = 10
        Grid1.CellBackColor = gcolor.BackColor
    Next i
    Grid1.Row = 1: Grid1.Col = 3
    DoEvents
    refgrid_Click
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 720
End Sub

Private Sub prtgrid_Click()
    Dim rt As String, rh As String, rf As String
    rt = Me.Caption
    rh = "Pallet Quantities"
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Call printflexgrid(Printer, Grid1, rt, rh, rf)
End Sub

Private Sub refgrid_Click()
    Dim i As Integer, cfile As String, ti As Long
    Dim t5 As Long, t8 As Long, t9 As Long, c10 As Long, t10 As Long
    Dim c6 As Single, c11 As Single, d11 As Single
    Dim rt As String, rh As String, rf As String
    Screen.MousePointer = 11
    Grid1.Redraw = False                                        'jv021516
    For i = 1 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 5) = ""
        Grid1.TextMatrix(i, 6) = ""
        Grid1.TextMatrix(i, 8) = ""
        Grid1.TextMatrix(i, 9) = ""
        Grid1.TextMatrix(i, 10) = ""
        Grid1.TextMatrix(i, 11) = ""
    Next i
    Call fetch_branch_pallets
    For i = 1 To Grid1.Rows - 2
        
        c10 = Val(Grid1.TextMatrix(i, 7))
        c10 = c10 - Val(Grid1.TextMatrix(i, 8))
        c10 = c10 - Val(Grid1.TextMatrix(i, 9))
        Grid1.TextMatrix(i, 10) = Format(c10, "#")
        t10 = t10 + c10
        c6 = Val(Grid1.TextMatrix(i, 4))
        If c6 > 0 Then
            t5 = t5 + Val(Grid1.TextMatrix(i, 5))
            t8 = t8 + Val(Grid1.TextMatrix(i, 8))
            t9 = t9 + Val(Grid1.TextMatrix(i, 9))
            c6 = Val(Grid1.TextMatrix(i, 5)) / c6
            Grid1.TextMatrix(i, 6) = Format(c6, ".000")
        End If
        c11 = Val(Grid1.TextMatrix(i, 3))
        If c11 > 0 Then
            d11 = Val(Grid1.TextMatrix(i, 5))
            d11 = d11 + Val(Grid1.TextMatrix(i, 8))
            d11 = d11 + Val(Grid1.TextMatrix(i, 9))
            c11 = d11 / c11
            Grid1.TextMatrix(i, 11) = Format(c11, ".000")
        End If
    Next i
    i = Grid1.Rows - 1
    Grid1.TextMatrix(i, 5) = t5
    If t5 <> 0 Then Grid1.TextMatrix(i, 6) = Format(t5 / Val(Grid1.TextMatrix(i, 4)), ".000")
    Grid1.TextMatrix(i, 8) = t8
    Grid1.TextMatrix(i, 9) = t9
    Grid1.TextMatrix(i, 10) = t10
    c11 = t5 + t8 + t9
    If c11 <> 0 Then c11 = c11 / Val(Grid1.TextMatrix(i, 3))
    Grid1.TextMatrix(i, 11) = Format(c11, ".000")
    
    ti = Val(Grid1.TextMatrix(i, 7)) + Val(Grid1.TextMatrix(i, 8)) + Val(Grid1.TextMatrix(i, 9))    'jv021516
    If ti = 0 Then                                                                                  'jv021516
        Grid1.ColWidth(7) = 0: Grid1.TextMatrix(0, 7) = "": Grid1.TextMatrix(i, 7) = ""             'jv021516
        Grid1.ColWidth(8) = 0: Grid1.TextMatrix(0, 8) = "": Grid1.TextMatrix(i, 8) = ""             'jv021516
        Grid1.ColWidth(9) = 0: Grid1.TextMatrix(0, 9) = "": Grid1.TextMatrix(i, 9) = ""             'jv021516
        Grid1.ColWidth(10) = 0: Grid1.TextMatrix(0, 10) = "": Grid1.TextMatrix(i, 10) = ""          'jv021516
    Else                                                                                            'jv021516
        Grid1.ColWidth(7) = 1000: Grid1.TextMatrix(0, 7) = "Ing Cap"                                'jv021516
        Grid1.ColWidth(8) = 1400: Grid1.TextMatrix(0, 8) = "Oracle Ings"                            'jv021516
        Grid1.ColWidth(9) = 1400: Grid1.TextMatrix(0, 9) = "WD Storage"                             'jv021516
        Grid1.ColWidth(10) = 1000: Grid1.TextMatrix(0, 10) = "Diff"                                 'jv021516
    End If                                                                                          'jv021516
    Grid1.Redraw = True                                                                             'jv021516
    rt = Me.Caption
    rh = Me.Caption
    rf = "Posted: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "lightgreen": gndc(0) = gcolor.BackColor
    'htdc(1) = "yellow": gndc(1) = Form1.ycolor.BackColor
    'htdc(2) = "tomato": gndc(2) = rcolor.BackColor
    'cfile = Form1.webdir & "\stock\brancaps.htm"
    Grid1.Redraw = False                                                                            'jv021516
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\stock\brancaps.htm"
    Call htmlcolorgrid(Me, cfile, Grid1, rt, rh, rf, "linen", "khaki", "white")
    Grid1.Redraw = True                                                                             'jv021516
    Screen.MousePointer = 0
End Sub
