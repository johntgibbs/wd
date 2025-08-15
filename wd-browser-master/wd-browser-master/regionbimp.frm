VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form regionbimp 
   Caption         =   "Form15"
   ClientHeight    =   11100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13935
   LinkTopic       =   "Form15"
   ScaleHeight     =   11100
   ScaleWidth      =   13935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Plant Inventory"
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
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pallet Turnover"
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
      Left            =   10680
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4215
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7435
      _Version        =   327680
      ForeColor       =   12582912
      BackColorFixed  =   16777152
      FocusRect       =   0
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
      Left            =   12360
      TabIndex        =   2
      Top             =   120
      Width           =   1575
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
      Left            =   7200
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
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
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label hcolor 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label2"
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2 Week Supply"
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
      Left            =   3840
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label regkey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Region:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "regionbimp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, s As String, i As Integer, k As Integer, psku As String, hflag As Boolean
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontSize = 8
    Grid1.FontBold = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12: Grid1.FixedCols = 0
    s = "select * from bimp where plantwhs <> 'DRY' order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "D" & regkey
            i = Val(ds!branchwhs)
            If branchrec(i).region = s And (branchrec(i).supplier = ds!plantwhs Or ds!plantwhs = "VENDOR") Then
                k = Val(ds!sku)
                s = ds!sku & Chr(9)
                s = s & skurec(k).unit & Chr(9)
                s = s & skurec(k).desc & Chr(9)
                s = s & Format(branchrec(i).branchno, "000") & "-" & branchrec(i).branchname & Chr(9)
                s = s & branchrec(i).supplier
                If branchrec(i).supplier = "T10" Then s = s & "-Brenham"
                If branchrec(i).supplier = "K10" Then s = s & "-Broken Arrow"
                If branchrec(i).supplier = "A10" Then s = s & "-Sylacauga"
                s = s & Chr(9)
                s = s & Format(ds!lastrecpt, "M-dd-yyyy") & Chr(9)
                s = s & Format(DateDiff("d", ds!lastrecpt, Now), "0") & Chr(9)
                s = s & Format(ds!onhand + ds!onorder, "#") & Chr(9)
                s = s & ds!sales & Chr(9)
                s = s & Format(ds!ohpct * 30, "#") & Chr(9)
                s = s & ds!bimpstatus
                Grid1.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    For i = 1 To Grid1.Rows - 1
        s = Grid1.TextMatrix(i, 1) & Grid1.TextMatrix(i, 2)
        s = s & Format(999999 - Val(Grid1.TextMatrix(i, 8)), "000000")
        Grid1.TextMatrix(i, 11) = s
    Next i
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 11: Grid1.ColSel = 11
    Grid1.Sort = 5
    psku = " ": hflag = False
    Grid1.FillStyle = flexFillRepeat
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) <> psku Then
            hflag = Not hflag
            psku = Grid1.TextMatrix(i, 0)
        Else
            Grid1.TextMatrix(i, 0) = " "
            Grid1.TextMatrix(i, 1) = " "
            Grid1.TextMatrix(i, 2) = " "
        End If
        If hflag = True Then
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
            Grid1.CellBackColor = hcolor.BackColor
        End If
    Next i
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 7: Grid1.ColSel = 8
        If Grid1.TextMatrix(i, 10) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
        If Grid1.TextMatrix(i, 10) = "W" Then Grid1.CellBackColor = wcolor.BackColor
        If Grid1.TextMatrix(i, 10) = "B" Then Grid1.CellBackColor = bcolor.BackColor
        If Grid1.TextMatrix(i, 10) = "G" Then Grid1.CellBackColor = gcolor.BackColor
    Next i
    Grid1.Row = 1: Grid1.Col = 0
    s = "^SKU|^Unit|<Flavor|<Branch|<Supplier|^Last Order|^Days|^On Hand|^30 Day Sales|^Days Supply"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 700
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 2100
    Grid1.ColWidth(4) = 1900
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 600
    Grid1.ColWidth(7) = 1200
    Grid1.ColWidth(8) = 1200
    Grid1.ColWidth(9) = 1200
    Grid1.ColWidth(10) = 0 '200
    Grid1.ColWidth(11) = 0
    Grid1.Redraw = True
    Me.Caption = "Region " & regkey & " Sales and Inventory"
    Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub Command2_Click()
    brzturnover.regkey = Val(regkey)
    DoEvents
    brzturnover.Show
End Sub

Private Sub Command3_Click()                                        'jv050616
    Dim s As String
    s = Left(Grid1.TextMatrix(Grid1.Row, 4), 3)
    If s = "T10" Then
        whssalesbrz.Caption = "Brenham Plant Distribution"
        whssalesbrz.qstr = "plana50"
        whssalesbrz.Show
    End If
    If s = "K10" Then
        whssalesbrz.Caption = "Broken Arrow Plant Distribution"
        whssalesbrz.qstr = "plana51"
        whssalesbrz.Show
    End If
    If s = "A10" Then
        whssalesbrz.Caption = "Sylacauga Plant Distribution"
        whssalesbrz.qstr = "plana52"
        whssalesbrz.Show
    End If
End Sub

Private Sub Form_Load()
    'Me.Width = Form1.Width - Form1.Combo1.Width
    'Me.Left = 0 'Form1.Left
    'refresh_grid
    Me.Width = Form1.Width
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 120
    If Me.Height > 2000 Then
        Grid1.Height = Me.Height - (regkey.Height * 4)
    End If
End Sub

Private Sub regkey_Change()
    refresh_grid
End Sub

