VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brzstkouts 
   Caption         =   "Stock History"
   ClientHeight    =   9945
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   15825
   LinkTopic       =   "Form15"
   ScaleHeight     =   9945
   ScaleWidth      =   15825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
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
      Left            =   8880
      TabIndex        =   5
      Top             =   240
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   11040
      TabIndex        =   4
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh Date"
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
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4575
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8070
      _Version        =   327680
      ForeColor       =   0
      BackColorFixed  =   12648447
      ForeColorFixed  =   0
      BackColorSel    =   32768
      FocusRect       =   0
      Appearance      =   0
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
      Left            =   1560
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label rcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8640
      TabIndex        =   6
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Branch Whs:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Menu sortmenu 
      Caption         =   "Sort Options"
      Begin VB.Menu sortuf 
         Caption         =   "Unit Flavor"
      End
      Begin VB.Menu sortsku 
         Caption         =   "SKU"
      End
      Begin VB.Menu sortostk 
         Caption         =   "Days Out-of-Stock"
      End
      Begin VB.Menu sortloads 
         Caption         =   "Loads"
      End
      Begin VB.Menu sortadl 
         Caption         =   "Avg Daily Loads"
      End
      Begin VB.Menu sortls 
         Caption         =   "Lost Unit Sales"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu helpmenu 
      Caption         =   "Help"
      Begin VB.Menu helpfile 
         Caption         =   "About Stock History"
      End
   End
End
Attribute VB_Name = "brzstkouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub refresh_grid1()
    Dim ds As ADODB.Recordset, s As String, i As Long
    Grid1.Redraw = False
    Grid1.FontName = "Callibri"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 11: Grid1.FixedCols = 3
    s = "select * from stockhistory where branchwhs = '" & List1 & "'"
    s = s & " and startdate <> 'N/A'"
    s = s & " order by loads desc"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            i = Val(ds!sku)
            s = s & skurec(i).unit & Chr(9)
            s = s & skurec(i).desc & Chr(9)
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
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 10: Grid1.ColSel = 10
            Grid1.CellForeColor = rcolor.BackColor
        Next i
    End If
    If sortls.Checked = True Then sortls_Click
    If sortuf.Checked = True Then sortuf_Click
    If sortsku.Checked = True Then sortsku_Click
    If sortostk.Checked = True Then sortostk_Click
    If sortloads.Checked = True Then sortloads_Click
    If sortadl.Checked = True Then sortadl_Click
    'Grid1.Row = 1: Grid1.RowSel = 1
    'Grid1.Col = 1: Grid1.ColSel = 2
    'Grid1.Sort = 5
    'Grid1.Col = 2
    Grid1.FormatString = "^SKU|^Unit|<Flavor|^Start Date|^End Date|^Total Days|^Days In-Stock|^Days Out-of-Stock|^Loads|^Avg Daily Loads|^Lost Unit Sales"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 2800
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


Private Sub refresh_lists()
    Dim i As Integer, k As Integer
    If Val(Form1.wdbranch) > 0 Then
        i = Val(Form1.wdbranch)
        Combo1.AddItem Format(branchrec(i).branchno, "000") & "-" & branchrec(i).branchname
        List1.AddItem Format(branchrec(i).branchno, "000")
    Else
        For i = 1 To 99
            If branchrec(i).oraloc > " " Then
                If Form1.wdbranch = "SU" Then
                    Combo1.AddItem Format(branchrec(i).branchno, "000") & "-" & branchrec(i).branchname
                    List1.AddItem Format(branchrec(i).branchno, "000")
                Else
                    's = "D" & Mid(Form1.wdbranch, 2, 1)
                    If branchrec(i).region = Form1.wdbranch Then
                        Combo1.AddItem Format(branchrec(i).branchno, "000") & "-" & branchrec(i).branchname
                        List1.AddItem Format(branchrec(i).branchno, "000")
                    End If
                End If
            End If
        Next i
        sortls.Checked = False
        sortuf.Checked = True
    End If
    k = Val(Mid(Form1.Combo1, 1, 2))
    For i = 0 To List1.ListCount - 1
        If Val(List1.List(i)) = k Then
            Combo1.ListIndex = i
            Exit For
        End If
    Next i
    'Combo1.ListIndex = 0
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    refresh_grid1
End Sub

Private Sub Command4_Click()
    Dim rt As String, rh As String, rf As String, hfile As String
    rt = Me.Caption & "  " & Combo1 '& "<br>" & sdate & " thru " & edate
    'rt = rt & "<br>" & Combo2 & " " & List2
    'rh = "Total Days:  " & tdaze
    'rh = rh & "<br>Days In-Stock:  " & idaze
    'rh = rh & "   Days Out-of-Stock:  " & odaze
    'rh = rh & "<br>Total Loads:  " & psales
    rf = "Printed:  " & Format(Now, "M-d-yyyy h:mm:ss am/pm")
    'EXCEL
    Grid1.Redraw = False
    hfile = "s:\wd\html\stock\" & List1 & "\skustk" & List1 & ".xls"
    Call htmlcolorgrid(Me, hfile, Grid1, rt, rh, rf, "lemonchiffon", "linen", "white")
    'HTML
    hfile = "s:\wd\html\stock\" & List1 & "\skustk" & List1 & ".htm"
    Call htmlcolorgrid(Me, hfile, Grid1, rt, rh, rf, "lemonchiffon", "linen", "white")
    Grid1.Redraw = True
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
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
    Me.Width = Form1.Width
    refresh_lists
End Sub

Private Sub Form_Resize()
    If Me.Height > 2000 Then
        Grid1.Height = Me.Height - (Combo1.Height * 4.5)
    End If
    Grid1.Width = Me.Width - 200
End Sub

Private Sub Grid1_DblClick()
    Dim pcheck As Boolean
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    pcheck = False
    If Form1.wdbranch = "SU" Then pcheck = True
    If Form1.wdbranch = "01" Then pcheck = True
    If Form1.wdbranch = "D1" Then pcheck = True
    If Form1.wdbranch = "D2" Then pcheck = True
    If Form1.wdbranch = "D3" Then pcheck = True
    If Form1.wdbranch = "D4" Then pcheck = True
    If Form1.wdbranch = "D5" Then pcheck = True
    If Form1.wdbranch = "D6" Then pcheck = True
    If Form1.wdbranch = "D7" Then pcheck = True
    If pcheck = True Then
        skustkouts.usertype = Form1.wdbranch
        skustkouts.skukey = Grid1.TextMatrix(Grid1.Row, 0)
        skustkouts.proddesc = Grid1.TextMatrix(Grid1.Row, 1) & " " & Grid1.TextMatrix(Grid1.Row, 2)
        skustkouts.Show
    End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu sortmenu
End Sub

Private Sub helpfile_Click()
    Form2.wdfile = Form1.webdir & "\stock\stkhelp.txt"
    Form2.Caption = "Stock History..."
    Form2.Left = Me.Left
    Form2.Width = Me.Width
    Form2.Top = Me.Top + (Command1.Height * 6)
    Form2.Height = Me.Height - (Command1.Height * 6)
    Form2.Show
End Sub

Private Sub sortadl_Click()
    sortadl.Checked = True
    sortls.Checked = False
    sortuf.Checked = False
    sortsku.Checked = False
    sortostk.Checked = False
    sortloads.Checked = False
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 9: Grid1.ColSel = 9
    Grid1.Sort = 4
    Grid1.Col = 8
End Sub

Private Sub sortloads_Click()
    sortadl.Checked = False
    sortls.Checked = False
    sortuf.Checked = False
    sortsku.Checked = False
    sortostk.Checked = False
    sortloads.Checked = True
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 8: Grid1.ColSel = 8
    Grid1.Sort = 4
    Grid1.Col = 8
End Sub

Private Sub sortls_Click()
    sortadl.Checked = False
    sortls.Checked = True
    sortuf.Checked = False
    sortsku.Checked = False
    sortostk.Checked = False
    sortloads.Checked = False
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 10: Grid1.ColSel = 10
    Grid1.Sort = 4
    Grid1.Col = 10
End Sub

Private Sub sortostk_Click()
    sortadl.Checked = False
    sortls.Checked = False
    sortuf.Checked = False
    sortsku.Checked = False
    sortostk.Checked = True
    sortloads.Checked = False
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 7: Grid1.ColSel = 7
    Grid1.Sort = 4
    Grid1.Col = 7
End Sub

Private Sub sortsku_Click()
    sortadl.Checked = False
    sortls.Checked = False
    sortuf.Checked = False
    sortsku.Checked = True
    sortostk.Checked = False
    sortloads.Checked = False
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 0
    Grid1.Sort = 5
    Grid1.Col = 3
End Sub

Private Sub sortuf_Click()
    sortadl.Checked = False
    sortls.Checked = False
    sortuf.Checked = True
    sortsku.Checked = False
    sortostk.Checked = False
    sortloads.Checked = False
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 1: Grid1.ColSel = 2
    Grid1.Sort = 5
    Grid1.Col = 3
End Sub
