VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brzplantorders 
   Caption         =   "Plant SKU Orders"
   ClientHeight    =   7125
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   12270
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
      Left            =   4680
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
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
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   11033
      _Version        =   327680
      ForeColor       =   8421376
      BackColorFixed  =   12648384
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   11520
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label dlit 
      Caption         =   "..."
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
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label rcolor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "rcolor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   8520
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "brzplantorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid(plantcode As String)
    Dim i As Integer, ss As ADODB.Recordset
    Dim ds As ADODB.Recordset, sqlx As String, oqty As Long
    Dim tc As Integer, mpal As Integer, np As Long, c As Long
    Dim psource As Integer
    'On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 6: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    sqlx = "select sku,plantpool,sum(onorder),sum(thiswknewpals),sum(nextwknewpals)"
    sqlx = sqlx & " from bimp"
    sqlx = sqlx & " where plantwhs = '" & plantcode & "'"
    sqlx = sqlx & " group by sku, plantpool"
    sqlx = sqlx & " order by sku"
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
            np = 0
            s = "select poolsched from bimp where sku = '" & ds(0) & "'"
            s = s & " and plantwhs = '" & plantcode & "'"
            s = s & " and poolsched > 0"
            Set ss = wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                np = ss(0)
            End If
            ss.Close
            
            s = "select sku, sum(thiswknewpals), sum(nextwknewpals), sum(onorder / roqty) from bimp"
            s = s & " where sku = '" & ds(0) & "'"
            If plantcode = "T10" Then
                s = s & " and plantwhs in ('A10', 'K10') and branchwhs = '001'"
            End If
            If plantcode = "K10" Then
                s = s & " and plantwhs in ('A10', 'T10') and branchwhs = '047'"
            End If
            If plantcode = "A10" Then
                s = s & " and plantwhs in ('K10', 'T10') and branchwhs = '052'"
            End If
            s = s & " group by sku"
            Set ss = wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                'MsgBox s
                If IsNull(ss(1)) = False Then
                    np = np + ss(1)
                    'If ss(1) > 0 Then MsgBox s & " = " & ss(1), vbOKOnly + vbInformation, "This week new.."
                End If
                If IsNull(ss(2)) = False Then
                    np = np + ss(2)
                    'If ss(2) > 0 Then MsgBox s & " = " & ss(2), vbOKOnly + vbInformation, "Next week new.."
                End If
                If IsNull(ss(3)) = False Then
                    np = np + ss(3)
                    'If ss(3) > 0 Then MsgBox s & " = " & ss(3), vbOKOnly + vbInformation, "On order.."
                End If
            End If
            ss.Close
            
            s = "select sku, sum(netqty) from brorders where sku = '" & ds(0) & "'"
            If plantcode = "T10" Then
                s = s & " and plant in (52, 51) and branch = 1"
            End If
            If plantcode = "K10" Then
                s = s & " and plant in (52, 50) and branch = 47"
            End If
            If plantcode = "A10" Then
                s = s & " and plant in (51, 50) and branch = 52"
            End If
            s = s & " group by sku having sum(netqty) <> 0"
            Set ss = wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                'MsgBox s
                If IsNull(ss(1)) = False Then
                    np = np + ss(1)
                    'If ss(1) > 0 Then MsgBox s & " = " & ss(1), vbOKOnly + vbInformation, "Branch Orders.."
                End If
            End If
            ss.Close
            
            oqty = ds(2)
            If IsNull(ds(3)) = False Then oqty = oqty + (ds(3) * mpal)
            If IsNull(ds(4)) = False Then oqty = oqty + (ds(4) * mpal)
            s = "select sku, sum(netqty) from brorders where sku = '" & ds(0) & "'"
            If plantcode = "T10" Then s = s & " and plant = 50"
            If plantcode = "K10" Then s = s & " and plant = 51"
            If plantcode = "A10" Then s = s & " and plant = 52"
            s = s & " group by sku having sum(netqty) <> 0"
            Set ss = wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                'MsgBox s
                If IsNull(ss(1)) = False Then
                    oqty = oqty + (ss(1) * mpal)
                    'If ss(1) > 0 Then MsgBox s & " = " & ss(1), vbOKOnly + vbInformation, "Branch Orders.."
                End If
            End If
            ss.Close
            oqty = oqty + (groupitems_qty(ds(0), plantcode, "ALL") * mpal)          'jv081516
            
            sqlx = sqlx & Format(ds(1) + ds(2), "#") & Chr(9)                                       'Plant Units
            sqlx = sqlx & Format((np * mpal), "#") & Chr(9)                              'New Pool Units
            'sqlx = sqlx & Format(ds(2) + (ds(3) * mpal) + (ds(4) * mpal), "#") & Chr(9)     'Branch Orders
            sqlx = sqlx & Format(oqty, "#") & Chr(9)     'Branch Orders
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            np = Val(Grid1.TextMatrix(i, 2)) + Val(Grid1.TextMatrix(i, 3))
            np = np - Val(Grid1.TextMatrix(i, 4))
            Grid1.TextMatrix(i, 5) = Format(np, "#")
            If Option2 = True Then
                mpal = 0
                If skurec(Val(Grid1.TextMatrix(i, 0))).sku = Grid1.TextMatrix(i, 0) Then
                    mpal = skurec(Val(Grid1.TextMatrix(i, 0))).pallet
                End If
                If mpal <> 0 Then
                    Grid1.TextMatrix(i, 2) = CInt(Val(Grid1.TextMatrix(i, 2)) / mpal)
                    Grid1.TextMatrix(i, 3) = CInt(Val(Grid1.TextMatrix(i, 3)) / mpal)
                    Grid1.TextMatrix(i, 4) = CInt(Val(Grid1.TextMatrix(i, 4)) / mpal)
                    np = Val(Grid1.TextMatrix(i, 2)) + Val(Grid1.TextMatrix(i, 3))
                    np = np - Val(Grid1.TextMatrix(i, 4))
                    Grid1.TextMatrix(i, 5) = Format(np, "#")
                End If
            End If
        Next i
    End If
    
    For i = Grid1.Rows - 1 To 1 Step -1                                                             'jv101016
        k = Val(Grid1.TextMatrix(i, 2)) + Val(Grid1.TextMatrix(i, 3)) + Val(Grid1.TextMatrix(i, 4)) 'jv101016
        If k = 0 Then                                                                               'jv101016
            If Grid1.Rows > 2 Then                                                                  'jv101016
                Grid1.RemoveItem i                                                                  'jv101016
            Else                                                                                    'jv101016
                Grid1.Rows = 1                                                                      'jv101016
            End If                                                                                  'jv101016
        End If                                                                                      'jv101016
    Next i                                                                                          'jv101016
    
    Screen.MousePointer = 0
    If Option1 = True Then
        Grid1.FormatString = "^SKU|<Product|^Plant Units|^New Pool|^Branch Orders|^Net"
    Else
        Grid1.FormatString = "^SKU|<Product|^Plant Pallets|^New Pool|^Branch Orders|^Net"
    End If
    
    Grid1.FillStyle = flexFillRepeat
    c = Grid1.BackColor
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
        Grid1.CellBackColor = c
        If c = Grid1.BackColorFixed Then
            c = Grid1.BackColor
        Else
            c = Grid1.BackColorFixed
        End If
        If Val(Grid1.TextMatrix(i, 5)) < 0 Then Grid1.CellForeColor = rcolor.ForeColor
    Next i
    Grid1.Row = 1
    
    
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1400 '6 '00
    Grid1.ColWidth(3) = 1400
    Grid1.ColWidth(4) = 1400 '6 '00
    Grid1.ColWidth(5) = 1400
    Grid1.Redraw = True
    
    s = "select listdisplay from valuelists where listname = 'bimpproddates'"   'jv091616
    s = s & " and listreturn = '" & List1 & "'"                                 'jv091616
    Set ds = wdb.Execute(s)                                                     'jv091616
    If ds.BOF = False Then                                                      'jv091616
        dlit = "Production dates thru " & ds(0) & "."                           'jv091616
    Else                                                                        'jv091616
        dlit = "..."                                                            'jv091616
    End If                                                                      'jv091616
    ds.Close                                                                    'jv091616
    
'    Exit Sub
'vberror:
'    eno = Err.Number: edesc = Err.Description: Err.Clear
'    Call vb_elog(eno, edesc, Me.Name, "plantot", Form1.UserId)
'    If eno = -2147467259 Then
'        Resume
'    Else
'        MsgBox edesc, vbOKOnly, Me.Name & " plantot - Error Number: " & eno
'        End
'    End If
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    Call refresh_grid(List1)
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Combo1.Clear: List1.Clear
    'Combo1.AddItem "All Plants": List1.AddItem "All"
    Combo1.AddItem "T10 - Brenham": List1.AddItem "T10"
    Combo1.AddItem "K10 - Broken Arrow": List1.AddItem "K10"
    Combo1.AddItem "A10 - Sylacauga": List1.AddItem "A10"
    Me.Height = Form1.WebBrowser1.Height 'whssales.Height
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Left = Form1.Width - Me.Width
    'sqlx = " update bimp set poolsched = 0 where poolsched is null"
    'wdb.Execute sqlx
    If Form1.wdbranch = "47" Then
        Combo1.ListIndex = 1
    Else
        If Form1.wdbranch = "52" Then
            Combo1.ListIndex = 2
        Else
            Combo1.ListIndex = 0
        End If
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 4)
End Sub

Private Sub Grid1_DblClick()
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) > 0 Then
        brzskuorders.Option1 = Me.Option1
        brzskuorders.Option2 = Me.Option2
        brzskuorders.Text1 = Combo1
        brzskuorders.Text2 = Grid1.TextMatrix(Grid1.Row, 0) & " " & Grid1.TextMatrix(Grid1.Row, 1)
        brzskuorders.msku = Grid1.TextMatrix(Grid1.Row, 0)
        brzskuorders.mplant = List1
        brzskuorders.Show
    End If
End Sub

Private Sub Option1_Click()
    Call refresh_grid(List1)
End Sub

Private Sub Option2_Click()
    Call refresh_grid(List1)
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    rt = Combo1 & " SKU Orders"
    rh = "SKU Orders"
    If Option1 = True Then
        rh = rh & " - Units"
    Else
        rh = rh & " - Pallets"
    End If
    rh = rh & "  " & dlit                       'jv091616
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    'htdc(0) = "lightcyan": gndc(0) = Me.bcolor.BackColor
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
