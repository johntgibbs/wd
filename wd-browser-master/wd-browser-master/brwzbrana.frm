VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brwzbrana 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Branch Inventory Management Program"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox age 
      Height          =   285
      Left            =   9480
      TabIndex        =   12
      Text            =   "0"
      Top             =   120
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4335
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7646
      _Version        =   327680
      BackColorFixed  =   65535
      GridColor       =   0
      ScrollTrack     =   -1  'True
      GridLinesFixed  =   1
      Appearance      =   0
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   3840
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label agelabel 
      Caption         =   "Label3"
      Height          =   255
      Left            =   8520
      TabIndex        =   15
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Days >="
      Height          =   255
      Left            =   8760
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Branch:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label gpct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "> 30 Day Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label bpct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "30 Day Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6000
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label ypct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2 Week Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label wpct 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "< 2 Week Supply"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "brwzbrana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()
    Dim tc As Integer
    tc = Check1.Value
    Check1.Value = 0
    If Left(List1, 1) = "P" Then
        refresh_plant
    Else
        refresh_branch
    End If
    Check1.Value = tc
End Sub

Private Sub refresh_branch()
    Dim cfile As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, i As Integer, s As String, tc As Integer
    Dim f13 As String
    Dim mpal As Integer
    Screen.MousePointer = 11
    Grid1.Visible = False: Grid1.Cols = 12: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    'cfile = "s:\wd\html\brana\branches.csv"
    cfile = Form1.webdir & "\brana\branches.csv"
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13
        If f0 = List1 Then
            If Val(f3) >= Val(age.Text) Then
                s = f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & f4 & Chr(9)
                s = s & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
                s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12
                Grid1.AddItem s
            End If
        End If
    Loop
    Close #1
    Screen.MousePointer = 0
    Grid1.FormatString = "^SKU|<Product|^Days|^OnHand|^OnOrd|^Sales|^UDiff|^PDiff|^OH%|^ROQty|^PG|^Need"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 3200
    Grid1.ColWidth(2) = 600
    Grid1.ColWidth(3) = 700
    Grid1.ColWidth(4) = 600
    Grid1.ColWidth(5) = 600
    Grid1.ColWidth(6) = 600
    Grid1.ColWidth(7) = 600
    Grid1.ColWidth(8) = 600
    Grid1.ColWidth(9) = 600
    Grid1.ColWidth(10) = 600
    Grid1.ColWidth(11) = 600
    tc = Check1.Value
    If Check1.Value = 1 Then Check1.Value = 0
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = 11 '10
        If Val(Grid1.TextMatrix(i, 11)) > 0 Then
            Grid1.CellBackColor = wcolor.BackColor
            wpct = Val(wpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
        Else
            If Val(Grid1.TextMatrix(i, 7)) = 0 Then
                Grid1.CellBackColor = bcolor.BackColor
                bpct = Val(bpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
            Else
                If Val(Grid1.TextMatrix(i, 7)) > 0 Then
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
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 0: Grid1.ColSel = 11 '10
        If Val(Grid1.TextMatrix(1, 11)) > 0 Then
            Grid1.CellBackColor = wcolor.BackColor
        Else
            If Val(Grid1.TextMatrix(1, 7)) = 0 Then
                Grid1.CellBackColor = bcolor.BackColor
            Else
                If Val(Grid1.TextMatrix(1, 7)) > 0 Then
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
    Grid1.Visible = True
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

End Sub

Private Sub refresh_plant()
    Dim cfile As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, i As Integer, s As String
    Dim mpal As Integer
    Screen.MousePointer = 11
    Grid1.Visible = False: Grid1.Cols = 12: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    'cfile = "s:\wd\html\brana\plants.csv"
    cfile = Form1.webdir & "\brana\plants.csv"
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12
        If f0 = List1 Then
            s = f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & f4 & Chr(9)
            s = s & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12
            Grid1.AddItem s
        End If
    Loop
    Close #1
    Screen.MousePointer = 0
    'Grid1.FormatString = "^SKU|<Product|^Days|^OnHand|^OnOrd|^Sales|^UDiff|^PDiff|^OH%|^ROQty|^PG|^Need"
    Grid1.FormatString = "^SKU|<Product||^OnHand||^Sales|^UDiff|^PDiff|^OH%|||"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 3200
    Grid1.ColWidth(2) = 1 '00
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1 '00
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1 '00
    Grid1.ColWidth(10) = 1 '00
    Grid1.ColWidth(11) = 1 '00
    'If Check1.Value = 1 Then Check1.Value = 0
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = 11 '10
        If Val(Grid1.TextMatrix(i, 8)) < 0.5 And Val(Grid1.TextMatrix(i, 8)) > 0 Then
            Grid1.CellBackColor = wcolor.BackColor
            wpct = Val(wpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
        Else
            If Val(Grid1.TextMatrix(i, 7)) = 0 Then
                Grid1.CellBackColor = bcolor.BackColor
                bpct = Val(bpct) + Abs(Val(Grid1.TextMatrix(i, 6)))
            Else
                If Val(Grid1.TextMatrix(i, 7)) > 0 Then
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
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 0: Grid1.ColSel = 11 '10
        If Val(Grid1.TextMatrix(1, 8)) < 0.5 And Val(Grid1.TextMatrix(1, 8)) > 0 Then
        'If Val(Grid1.TextMatrix(1, 11)) > 0 Then
            Grid1.CellBackColor = wcolor.BackColor
        Else
            If Val(Grid1.TextMatrix(1, 7)) = 0 Then
                Grid1.CellBackColor = bcolor.BackColor
            Else
                If Val(Grid1.TextMatrix(1, 7)) > 0 Then
                    Grid1.CellBackColor = gcolor.BackColor
                Else
                    Grid1.CellBackColor = ycolor.BackColor
                End If
            End If
        End If
        Grid1.FillStyle = flexFillRepeat
        'Check1.Value = tc
        Grid1.Row = 1: Grid1.Col = 3
    End If
    Grid1.Visible = True
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

End Sub

Private Sub refresh_whslist()
    Dim f0 As String, f1 As String, cfile As String
    Combo1.Clear: List1.Clear
    'cfile = "s:\wd\html\brana\whslist.csv"
    cfile = Form1.webdir & "\brana\whslist.csv"
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1
        If Form1.wdbranch = "500" Or Form1.wdbranch = "01" Or Form1.wdbranch = "SU" Then
            Combo1.AddItem f1
            List1.AddItem f0
        Else
            If Val(Form1.wdbranch) = Val(f0) Then
                Combo1.AddItem f1
                List1.AddItem f0
            End If
        End If
    Loop
    Close #1
    Combo1.ListIndex = 0
End Sub

Private Sub age_LostFocus()
    agelabel.Caption = age.Text
End Sub

Private Sub agelabel_Change()
    refresh_grid1
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Form_Load()
    refresh_whslist
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Me.Height > 2000 Then
        Grid1.Height = Me.Height - 980 '655
    End If
    
End Sub

Private Sub Grid1_DblClick()
    If Combo1.ListCount <= 1 Then Exit Sub
    Screen.MousePointer = 11
    brwzbrana2.calledby = "Brana"
    brwzbrana2.wsku = Grid1.TextMatrix(Grid1.Row, 0)
    Screen.MousePointer = 0
    brwzbrana2.Show
    Check1 = 1
End Sub

Private Sub Grid1_RowColChange()
    Dim i As Integer
    i = Grid1.Row
    If Check1 = 1 And Left(brwzbrana2.Caption, Len(Grid1.TextMatrix(i, 1))) <> Grid1.TextMatrix(i, 1) Then
        brwzbrana2.calledby = "Brana"
        brwzbrana2.wsku = Grid1.TextMatrix(Grid1.Row, 0)
    End If
End Sub

Private Sub List1_Click()
    refresh_grid1
End Sub
