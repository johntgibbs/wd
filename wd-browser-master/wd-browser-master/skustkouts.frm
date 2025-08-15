VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form skustkouts 
   Caption         =   "Product Stock History"
   ClientHeight    =   8160
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15705
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form15"
   ScaleHeight     =   8160
   ScaleWidth      =   15705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   12360
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7455
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   13150
      _Version        =   327680
      BackColorFixed  =   16777152
      BackColorSel    =   128
      FocusRect       =   0
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
      Height          =   1035
      Left            =   7920
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label usertype 
      Caption         =   "Label3"
      Height          =   255
      Left            =   14160
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label rcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10080
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label proddesc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label skukey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Product:"
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Supplier:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu sortmenu 
      Caption         =   "Sort"
      Visible         =   0   'False
      Begin VB.Menu sortbranch 
         Caption         =   "Branch"
         Checked         =   -1  'True
      End
      Begin VB.Menu sortdays 
         Caption         =   "Days Out of Stock"
      End
      Begin VB.Menu sortloads 
         Caption         =   "Loads"
      End
      Begin VB.Menu sortadl 
         Caption         =   "Avg Daily Loads"
      End
      Begin VB.Menu sortlost 
         Caption         =   "Lost Sales"
      End
   End
End
Attribute VB_Name = "skustkouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()
    Dim ds As ADODB.Recordset, s As String, i As Long, hflag As Boolean
    Dim t6 As Long, t7 As Long, t8 As Long, t9 As Long, t10 As Long
    Grid1.Redraw = False
    Grid1.FontName = "Callibri"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 11: Grid1.FixedCols = 0
    s = "select * from stockhistory where sku = '" & skukey & "'"
    s = s & " and startdate <> 'N/A'"
    s = s & " and branchwhs <> '052'"
    s = s & " order by branchwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!branchwhs & "-" & branchrec(Val(ds!branchwhs)).branchname & Chr(9)
            s = s & branchrec(Val(ds!branchwhs)).supplier & Chr(9)
            s = s & branchrec(Val(ds!branchwhs)).region & Chr(9)
            s = s & Format(ds!startdate, "M-d-yyyy") & Chr(9)
            s = s & Format(ds!enddate, "M-d-yyyy") & Chr(9)
            s = s & Format(ds!totaldays, "#") & Chr(9)
            s = s & Format(ds!daysin, "#") & Chr(9)
            s = s & Format(ds!daysout, "#") & Chr(9)
            s = s & Format(ds!loads, "#") & Chr(9)
            If ds!loads > 0 And ds!daysin > 0 Then
                i = (ds!loads / ds!daysin)
                s = s & i & Chr(9)
                'i = (ds!loads / ds!daysin) * ds!daysout
                i = i * ds!daysout
            Else
                s = s & Chr(9)
                i = 0
            End If
            s = s & Format(i, "#")
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    If Me.usertype = "SU" Or Me.usertype = "01" Then
        If List1 <> "ALL" And Grid1.Rows > 1 Then
            For i = Grid1.Rows - 1 To 1 Step -1
                If Grid1.TextMatrix(i, 1) <> List1 Then
                    If Grid1.Rows > 2 Then
                        Grid1.RemoveItem i
                    Else
                        Grid1.Rows = 1
                    End If
                End If
            Next i
        End If
    Else
        For i = Grid1.Rows - 1 To 1 Step -1
            If Grid1.TextMatrix(i, 2) <> Me.usertype Then
                If Grid1.Rows > 2 Then
                    Grid1.RemoveItem i
                Else
                    Grid1.Rows = 1
                End If
            End If
        Next i
    End If
    If sortdays.Checked = True Then
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 7: Grid1.ColSel = 7
        Grid1.Sort = 4
    End If
    If sortloads.Checked = True Then
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 8: Grid1.ColSel = 8
        Grid1.Sort = 4
    End If
    If sortadl.Checked = True Then
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 9: Grid1.ColSel = 9
        Grid1.Sort = 4
    End If
    If sortlost.Checked = True Then
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 10: Grid1.ColSel = 10
        Grid1.Sort = 4
    End If
    If Grid1.Rows > 1 Then
        t6 = 0: t7 = 0: t8 = 0: t9 = 0: t10 = 0
        For i = 1 To Grid1.Rows - 1
            t6 = t6 + Val(Grid1.TextMatrix(i, 6))
            t7 = t7 + Val(Grid1.TextMatrix(i, 7))
            t8 = t8 + Val(Grid1.TextMatrix(i, 8))
            t9 = t9 + Val(Grid1.TextMatrix(i, 9))
            t10 = t10 + Val(Grid1.TextMatrix(i, 10))
        Next i
        s = ".." & Chr(9)
        s = s & ".." & Chr(9)
        s = s & ".." & Chr(9)
        s = s & ".." & Chr(9)
        s = s & "Totals" & Chr(9)
        s = s & ".." & Chr(9)
        s = s & Format(t6, "#") & Chr(9)
        s = s & Format(t7, "#") & Chr(9)
        s = s & Format(t8, "#") & Chr(9)
        s = s & Format(t9, "#") & Chr(9)
        s = s & Format(t10, "#")
        Grid1.AddItem s
    End If
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            hflag = Not hflag
            If hflag = True Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 0: Grid1.ColSel = 10
                Grid1.CellBackColor = Grid1.BackColorFixed
            End If
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 10: Grid1.ColSel = 10
            Grid1.CellForeColor = rcolor.BackColor
        Next i
        Grid1.Row = 1: Grid1.Col = 2
    End If
    Grid1.FormatString = "<Branch|^Plant|^Region|^Start Date|^End Date|^Total Days|^Days In-Stock|^Days Out-of-Stock|^Loads|^Avg Daily Loads|^Lost Unit Sales"
    Grid1.ColWidth(0) = 2200
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 1400
    Grid1.ColWidth(4) = 1400
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 1600
    Grid1.ColWidth(7) = 1600
    Grid1.ColWidth(8) = 1400
    Grid1.ColWidth(9) = 1400
    Grid1.ColWidth(10) = 1400
    Grid1.Redraw = True
End Sub


Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    refresh_grid1
End Sub

Private Sub Command1_Click()
    Dim rt As String, rh As String, rf As String, hfile As String
    rt = Me.Caption & "  " & Combo1 '& "<br>" & sdate & " thru " & edate
    rh = skukey.Caption & " " & proddesc.Caption
    rf = "Printed:  " & Format(Now, "M-d-yyyy h:mm:ss am/pm")
    hfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\htmtemp.htm"
    Grid1.Redraw = False
    Call htmlcolorgrid(Me, hfile, Grid1, rt, rh, rf, "lemonchiffon", "linen", "white")
    Grid1.Redraw = True
    Grid1.Row = 1: Grid1.Col = 2
    Form1.WebBrowser1.Navigate hfile
    'Unload Me
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe " & hfile, vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & hfile, vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Me.Top = brzstkouts.Top
    Me.Height = brzstkouts.Height
    Me.Left = brzstkouts.Width - Me.Width
    Combo1.Clear: List1.Clear
    Combo1.AddItem "All Plants": List1.AddItem "ALL"
    Combo1.AddItem "T10-Brenham": List1.AddItem "T10"
    Combo1.AddItem "K10-Broken Arrow": List1.AddItem "K10"
    Combo1.AddItem "A10-Sylacauga": List1.AddItem "A10"
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 200
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 3.5)
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu sortmenu
End Sub

Private Sub Grid1_RowColChange()
    Dim i As Integer, pals As Currency, pconv As Integer
    i = Grid1.Row
    pconv = skurec(Val(skukey)).pallet
    Grid1.ToolTipText = ""
    'If Val(Grid1.TextMatrix(i, 10)) = 0 Then Exit Sub
    If Grid1.Col = 8 Then
        If Val(Grid1.TextMatrix(i, 8)) > 0 Then
            pals = Format(Val(Grid1.TextMatrix(i, 8)) / pconv, "0.00")
            Grid1.ToolTipText = "Loaded Pallets: " & pals
        End If
    End If
    If Grid1.Col = 9 Then
        If Val(Grid1.TextMatrix(i, 9)) > 0 Then
            pals = Format(Val(Grid1.TextMatrix(i, 9)) / pconv, "0.00")
            Grid1.ToolTipText = "Daily Loaded Pallets: " & pals
        End If
    End If
    If Grid1.Col = 10 Then
        If Val(Grid1.TextMatrix(i, 10)) > 0 Then
            pals = Format(Val(Grid1.TextMatrix(i, 10)) / pconv, "0.00")
            Grid1.ToolTipText = "Lost Pallet Sales: " & pals
        End If
    End If
End Sub

Private Sub skukey_Change()
    refresh_grid1
End Sub

Private Sub sortadl_Click()
    sortadl.Checked = True
    sortbranch.Checked = False
    sortdays.Checked = False
    sortloads.Checked = False
    sortlost.Checked = False
    refresh_grid1
End Sub

Private Sub sortbranch_Click()
    sortadl.Checked = False
    sortbranch.Checked = True
    sortdays.Checked = False
    sortloads.Checked = False
    sortlost.Checked = False
    refresh_grid1
End Sub

Private Sub sortdays_Click()
    sortadl.Checked = False
    sortbranch.Checked = False
    sortdays.Checked = True
    sortloads.Checked = False
    sortlost.Checked = False
    refresh_grid1
End Sub

Private Sub sortloads_Click()
    sortadl.Checked = False
    sortbranch.Checked = False
    sortdays.Checked = False
    sortloads.Checked = True
    sortlost.Checked = False
    refresh_grid1
End Sub

Private Sub sortlost_Click()
    sortadl.Checked = False
    sortbranch.Checked = False
    sortdays.Checked = False
    sortloads.Checked = False
    sortlost.Checked = True
    refresh_grid1
End Sub
