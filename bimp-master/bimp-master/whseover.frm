VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form whseover 
   Caption         =   "Over Stocked Branch Items"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   13980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "No Sales List"
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
      Left            =   12120
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
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
      Left            =   10560
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
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
      Left            =   9000
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      TabIndex        =   6
      Text            =   "180"
      Top             =   120
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   6240
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   3615
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
      Left            =   1320
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   12515
      _Version        =   327680
      BackColorFixed  =   14737632
   End
   Begin VB.Label hcolor 
      BackColor       =   &H0000FFFF&
      Caption         =   "hcolor"
      Height          =   255
      Left            =   10560
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Days."
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
      Left            =   7800
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Unit Supply >="
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
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Warehouse:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "whseover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refresh_vlists()
    Combo1.Clear: List1.Clear
    Combo1.AddItem "ALL": List1.AddItem "ALL Sources"
    Combo1.AddItem "A10": List1.AddItem "Sylacauga Plant"
    Combo1.AddItem "K10": List1.AddItem "Broken Arrow Plant"
    Combo1.AddItem "T10": List1.AddItem "Brenham Plant"
    Combo1.AddItem "VENDOR": List1.AddItem "Vendor Items"
End Sub

Private Sub zero_sales()
    Dim ds As ADODB.Recordset, sqlx As String, i As Integer
    Dim psku As String, hflag As Boolean
    'On Error GoTo vberror
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
    sqlx = sqlx & ",thiswknewpals,nextwknewpals,branchwhs"                             'jv072216
    sqlx = sqlx & " from bimp"
    If Combo1 = "ALL" Then
        sqlx = sqlx & " Where plantwhs not in ('DRY', 'VENDOR')"
    Else
        sqlx = sqlx & " Where plantwhs = '" & Combo1 & "'"
    End If
    sqlx = sqlx & " and branchwhs not in ('001', '012', '047', '052')"
    sqlx = sqlx & " and onhand > 0 and sales < 1"
    sqlx = sqlx & " order by sku, branchwhs"
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
            sqlx = sqlx & ds!branchwhs & "-" & branchrec(Val(ds!branchwhs)).branchname & Chr(9)
            
            sqlx = sqlx & Format(ds(2), "#") & Chr(9)
            sqlx = sqlx & Format(ds(3), "#") & Chr(9)
            sqlx = sqlx & Format(ds(4), "#") & Chr(9)
            sqlx = sqlx & Format(ds(5), "#") & Chr(9)
            sqlx = sqlx & Format(ds(6), "#") & Chr(9)
            sqlx = sqlx & Format(ds(7), ".000") & Chr(9)
            sqlx = sqlx & DateDiff("d", ds(1), Now) & Chr(9) & Chr(9) & Chr(9)
            sqlx = sqlx & ds(11)
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
    Grid1.FormatString = "^SKU|<Product|<Branch|^OnHand|^OnOrder|^Sales|^UnitDiff|^PalletDiff|^OH%|^PalSize|^%Gain|^Need|^Source|^Days Supply"
    Grid1.FormatString = "^SKU|<Product|<Branch|^OnHand|^OnOrder|^Sales|^UnitDiff|^PalletDiff|^OH%|^Days InStock|^|^|^Source|^Days Supply"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 3200 '4000
    Grid1.ColWidth(2) = 2200 '2700
    Grid1.ColWidth(3) = 1100
    Grid1.ColWidth(4) = 1100
    Grid1.ColWidth(5) = 1100
    Grid1.ColWidth(6) = 1100
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1100
    Grid1.ColWidth(9) = 1300
    Grid1.ColWidth(10) = 0 '1100
    Grid1.ColWidth(11) = 0 '1100
    Grid1.ColWidth(12) = 1000
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 8)) > 0 Then                                         'jv041116
                Grid1.TextMatrix(i, 13) = Format(Val(Grid1.TextMatrix(i, 8)) * 30, "0")     'jv041116
            End If                                                                          'jv041116
        Next i
    
        Grid1.FillStyle = flexFillRepeat
        hflag = True: psku = Grid1.TextMatrix(1, 0)
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 0) <> psku Then
                hflag = Not hflag
                psku = Grid1.TextMatrix(i, 0)
            End If
            If hflag = True Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = hcolor.BackColor
            End If
        Next i
    End If
    If Grid1.Rows > 1 Then Grid1.Row = 1
    
    Grid1.Redraw = True
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

Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, sqlx As String, i As Integer
    Dim psku As String, hflag As Boolean
    'On Error GoTo vberror
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
    sqlx = sqlx & ",thiswknewpals,nextwknewpals,branchwhs"                             'jv072216
    sqlx = sqlx & " from bimp"
    If Combo1 = "ALL" Then
        sqlx = sqlx & " Where plantwhs not in ('DRY', 'VENDOR')"
    Else
        sqlx = sqlx & " Where plantwhs = '" & Combo1 & "'"
    End If
    sqlx = sqlx & " and branchwhs not in ('001', '047', '052')"
    sqlx = sqlx & " order by sku, branchwhs"
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
            sqlx = sqlx & ds!branchwhs & "-" & branchrec(Val(ds!branchwhs)).branchname & Chr(9)
            
            sqlx = sqlx & Format(ds(2), "#") & Chr(9)
            sqlx = sqlx & Format(ds(3), "#") & Chr(9)
            sqlx = sqlx & Format(ds(4), "#") & Chr(9)
            sqlx = sqlx & Format(ds(5), "#") & Chr(9)
            sqlx = sqlx & Format(ds(6), "#") & Chr(9)
            sqlx = sqlx & Format(ds(7), ".000") & Chr(9)
            sqlx = sqlx & DateDiff("d", ds(1), Now) & Chr(9) & Chr(9) & Chr(9)
            sqlx = sqlx & ds(11)
            If ds!ohpct * 30 > Val(Text1) And DateDiff("d", ds(1), Now) > Val(Text1) Then Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
    Grid1.FormatString = "^SKU|<Product|<Branch|^OnHand|^OnOrder|^Sales|^UnitDiff|^PalletDiff|^OH%|^PalSize|^%Gain|^Need|^Source|^Days Supply"
    Grid1.FormatString = "^SKU|<Product|<Branch|^OnHand|^OnOrder|^Sales|^UnitDiff|^PalletDiff|^OH%|^Days InStock|^|^|^Source|^Days Supply"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 3200 '4000
    Grid1.ColWidth(2) = 2200 '2700
    Grid1.ColWidth(3) = 1100
    Grid1.ColWidth(4) = 1100
    Grid1.ColWidth(5) = 1100
    Grid1.ColWidth(6) = 1100
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1100
    Grid1.ColWidth(9) = 1300
    Grid1.ColWidth(10) = 0 '1100
    Grid1.ColWidth(11) = 0 '1100
    Grid1.ColWidth(12) = 1000
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 8)) > 0 Then                                         'jv041116
                Grid1.TextMatrix(i, 13) = Format(Val(Grid1.TextMatrix(i, 8)) * 30, "0")     'jv041116
            End If                                                                          'jv041116
        Next i
    
        Grid1.FillStyle = flexFillRepeat
        hflag = True: psku = Grid1.TextMatrix(1, 0)
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 0) <> psku Then
                hflag = Not hflag
                psku = Grid1.TextMatrix(i, 0)
            End If
            If hflag = True Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = hcolor.BackColor
            End If
        Next i
    End If
    If Grid1.Rows > 1 Then Grid1.Row = 1
    
    Grid1.Redraw = True
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

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    Label2.Caption = List1
    DoEvents
    refresh_grid
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub Command2_Click()
    Dim rt As String, rf As String, rh As String
    rt = Me.Caption
    rh = "Days of Supply >= " & Text1 & "    Source: " & Label2.Caption
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

Private Sub Command3_Click()
    zero_sales
End Sub

Private Sub Form_Load()
    refresh_vlists
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 3.5)
End Sub

Private Sub Grid1_DblClick()
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) > 0 Then
        branchostk.Label2.Caption = Grid1.TextMatrix(Grid1.Row, 2)
        branchostk.Show
    End If
End Sub
