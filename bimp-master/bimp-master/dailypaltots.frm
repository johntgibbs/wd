VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form dailypaltots 
   Caption         =   "Daily Pallet Loads"
   ClientHeight    =   11430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   ScaleHeight     =   11430
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
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
      Left            =   7080
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   600
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   10080
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   3255
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
      Left            =   960
      TabIndex        =   14
      Text            =   "Combo2"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
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
      Left            =   9000
      TabIndex        =   12
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Refresh All"
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
      Left            =   9840
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   3975
      Left            =   0
      TabIndex        =   9
      Top             =   7320
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7011
      _Version        =   327680
      BackColorFixed  =   16776960
      BackColorSel    =   255
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   6015
      Left            =   6480
      TabIndex        =   8
      Top             =   1080
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   10610
      _Version        =   327680
      BackColorFixed  =   12648447
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6015
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   10610
      _Version        =   327680
      Cols            =   4
      ForeColor       =   16711680
      BackColorFixed  =   49152
      ForeColorFixed  =   16777215
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
      Left            =   8400
      TabIndex        =   6
      Top             =   120
      Width           =   1335
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
      Left            =   7080
      TabIndex        =   5
      Text            =   "30"
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   10080
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   960
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label hcolor 
      BackColor       =   &H0080FF80&
      Caption         =   "hcolor"
      Height          =   255
      Left            =   7440
      TabIndex        =   19
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Process Date:"
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
      Left            =   5760
      TabIndex        =   17
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "SKU:"
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
      TabIndex        =   13
      Top             =   600
      Width           =   735
   End
   Begin VB.Label datelit 
      Alignment       =   2  'Center
      Caption         =   "Label4"
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
      Left            =   0
      TabIndex        =   11
      Top             =   7080
      Width           =   6495
   End
   Begin VB.Label Label3 
      Caption         =   "Load days:"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   2535
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "dailypaltots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refresh_vlists()
    Combo1.Clear: List1.Clear
    For i = 1 To 99
        If branchrec(i).oraloc > " " Then
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
    Combo2.Clear: List2.Clear
    Combo2.AddItem "ALL"
    List2.AddItem "All Products"
    For i = 0 To 9999
        If skurec(i).sku > "0" Then
            Combo2.AddItem skurec(i).sku
            List2.AddItem skurec(i).unit & " " & skurec(i).desc
        End If
    Next i
    Combo2.ListIndex = 0
    Combo1.ListIndex = 0
End Sub

Sub refresh_grid()
    Dim ds As ADODB.Recordset, q As String, s As String
    Dim i As Integer, t As Currency, j As Long, newrec As Boolean
    Dim d1 As Integer, d2 As Integer
    d2 = DateDiff("d", Now, Text2) * -1
    d1 = d2 + Val(Text1) + 1
    'MsgBox "d1=" & d1 & " : d2=" & d2
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
        
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 5: Grid1.Rows = 1
    Grid1.FixedCols = 1
    Grid1.Clear
    'For i = 1 To Val(Text1) - 1
    'j = 0
    j = Val(Text1) + 1
    'For i = Val(Text1) - 1 To 0 Step -1
    For i = Val(Text1) To 1 Step -1
        'j = j + 1
        j = j - 1
        s = j & Chr(9) & Format(DateAdd("d", i * -1, Text2), "M/d/yyyy")
        Grid1.AddItem s
    Next i
    
    
    q = "select tran_date,product_no,sum(tran_qty)"
    q = q & " from bolinf.inv_adj_input_detail"
    q = q & " where tran_type = '1'"
    If Combo2 <> "ALL" Then q = q & " and product_no = '" & Combo2 & "'"
    'q = q & " and trunc(tran_date) > trunc(SYSDATE - " & Val(Text1) + 1 & ")"
    q = q & " and trunc(tran_date) > trunc(SYSDATE - " & d1 & ")"
    'q = q & " and trunc(tran_date) < trunc(SYSDATE - " & d2 & ")"
    q = q & " and branch_no = '" & Combo1 & "'"
    q = q & " group by tran_date, product_no"
    q = q & " order by tran_date, product_no"
    'MsgBox q
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        j = 0
        Do Until ds.EOF
            If skurec(Val(ds!product_no)).pallet > 0 Then
                'newrec = True
                t = ds(2) / skurec(Val(ds!product_no)).pallet
                For i = 0 To Grid1.Rows - 1
                    If Grid1.TextMatrix(i, 1) = ds!tran_date Then
                        't = ds(2) / skurec(Val(ds!product_no)).pallet
                        Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + ds(2)
                        Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + t
                        'newrec = False
                    End If
                Next i
                'If newrec = True Then
                '    j = j + 1
                '    s = j & Chr(9) & ds!tran_date & Chr(9) & " " & Chr(9) & ds(2) & Chr(9) & t
                '    If j <= Val(Text1) Then Grid1.AddItem s
                'End If
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            Grid1.TextMatrix(i, 2) = Format(Grid1.TextMatrix(i, 1), "dddd")
        Next i
    End If
    'If (Combo2 <> "ALL" Or Combo3 <> "ALL") And Grid1.Rows > 1 Then
    '    t = 0: j = 0
    '    For i = 1 To Grid1.Rows - 1
    '        t = t + Val(Grid1.TextMatrix(i, 3))
    '        j = j + Val(Grid1.TextMatrix(i, 4))
    '    Next i
    '    s = " " & Chr(9) & " " & Chr(9) & "Totals" & Chr(9) & t & Chr(9) & j
    '    Grid1.AddItem s
    'End If
    Grid1.FormatString = "^#|^Date|^Day of Week|^Units|^Pallets"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 1600
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.Redraw = True
    Screen.MousePointer = 1
    refresh_grid2
End Sub

Private Sub refresh_grid2()
    Dim i As Integer, k As Integer, s As String, t1 As Long, t2 As Currency     'jv090617
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Cols = 4: Grid2.Rows = 1
    Grid2.FixedCols = 1
    Grid2.Clear
    Grid2.AddItem "Sunday"
    Grid2.AddItem "Monday"
    Grid2.AddItem "Tuesday"
    Grid2.AddItem "Wednesday"
    Grid2.AddItem "Thursday"
    Grid2.AddItem "Friday"
    Grid2.AddItem "Saturday"
    For i = 1 To Grid1.Rows - 1
        For k = 1 To Grid2.Rows - 1
            If Grid1.TextMatrix(i, 2) = Grid2.TextMatrix(k, 0) Then
                Grid2.TextMatrix(k, 1) = Val(Grid2.TextMatrix(k, 1)) + Val(Grid1.TextMatrix(i, 3))
                Grid2.TextMatrix(k, 2) = Val(Grid2.TextMatrix(k, 2)) + Val(Grid1.TextMatrix(i, 4))
                Grid2.TextMatrix(k, 3) = Val(Grid2.TextMatrix(k, 3)) + 1
                Exit For
            End If
        Next k
    Next i
    'Factor qtys to 30 days
    'For i = 1 To Grid2.Rows - 1
    '    Grid2.TextMatrix(i, 1) = Format(Val(Grid2.TextMatrix(i, 1)) * (30 / Val(Text1)), "0")
    '    Grid2.TextMatrix(i, 2) = Format(Val(Grid2.TextMatrix(i, 2)) * (30 / Val(Text1)), "0.0000")
    'Next i
    t1 = 0: t2 = 0
    For i = 1 To Grid2.Rows - 1
        t1 = t1 + Val(Grid2.TextMatrix(i, 1))
        t2 = t2 + Val(Grid2.TextMatrix(i, 2))
    Next i
    Grid2.AddItem "Total" & Chr(9) & t1 & Chr(9) & t2
    Grid2.FormatString = "^Day of Week|^Units|^Pallets|^# Days"
    Grid2.ColWidth(0) = 1600
    Grid2.ColWidth(1) = 1200
    Grid2.ColWidth(2) = 1200
    Grid2.ColWidth(3) = 1200
    Grid2.Redraw = True
    post_grid3
End Sub

Private Sub refresh_grid3()
    Dim i As Integer, s As String
    Grid3.Redraw = False
    Grid3.FontName = "Arial"
    Grid3.FontBold = True
    Grid3.FontSize = 8
    Grid3.Cols = 10: Grid3.Rows = 1
    Grid3.FixedCols = 2
    Grid3.Clear
    
    For i = 0 To Combo1.ListCount - 1
        s = Combo1.List(i) & Chr(9) & List1.List(i)
        Grid3.AddItem s
    Next i
    s = "All" & Chr(9) & "Totals"
    Grid3.AddItem s
    Grid3.FormatString = "^Whs|<Location|^Sun|^Mon|^Tue|^Wed|^Thu|^Fri|^Sat|^Total"
    Grid3.ColWidth(0) = 800
    Grid3.ColWidth(1) = 2200
    Grid3.ColWidth(2) = 900
    Grid3.ColWidth(3) = 900
    Grid3.ColWidth(4) = 900
    Grid3.ColWidth(5) = 900
    Grid3.ColWidth(6) = 900
    Grid3.ColWidth(7) = 900
    Grid3.ColWidth(8) = 900
    Grid3.ColWidth(9) = 900
    Grid3.Redraw = True
End Sub

Private Sub post_grid3()
    Dim i As Integer
    Dim t1 As Long, t2 As Long, t3 As Long, t4 As Long, t5 As Long, t6 As Long, t7 As Long, t8 As Long
    t1 = 0: t2 = 0: t3 = 0: t4 = 0: t5 = 0: t6 = 0: t7 = 0: t8 = 0
    If Grid3.Cols < 5 Then refresh_grid3
    datelit = Text1 & " Days:  " & Grid1.TextMatrix(1, 1) & " thru " & Grid1.TextMatrix(Grid1.Rows - 1, 1)
    For i = 1 To Grid3.Rows - 1
        If Grid3.TextMatrix(i, 0) = Combo1 Then
            Grid3.TextMatrix(i, 2) = Format(CInt(Grid2.TextMatrix(1, 2)), "#")
            Grid3.TextMatrix(i, 3) = Format(CInt(Grid2.TextMatrix(2, 2)), "#")
            Grid3.TextMatrix(i, 4) = Format(CInt(Grid2.TextMatrix(3, 2)), "#")
            Grid3.TextMatrix(i, 5) = Format(CInt(Grid2.TextMatrix(4, 2)), "#")
            Grid3.TextMatrix(i, 6) = Format(CInt(Grid2.TextMatrix(5, 2)), "#")
            Grid3.TextMatrix(i, 7) = Format(CInt(Grid2.TextMatrix(6, 2)), "#")
            Grid3.TextMatrix(i, 8) = Format(CInt(Grid2.TextMatrix(7, 2)), "#")
            Grid3.TextMatrix(i, 9) = Format(CInt(Grid2.TextMatrix(8, 2)), "#")
            Exit For
        End If
    Next i
    For i = 1 To Grid3.Rows - 2
        t1 = t1 + Val(Grid3.TextMatrix(i, 2))
        t2 = t2 + Val(Grid3.TextMatrix(i, 3))
        t3 = t3 + Val(Grid3.TextMatrix(i, 4))
        t4 = t4 + Val(Grid3.TextMatrix(i, 5))
        t5 = t5 + Val(Grid3.TextMatrix(i, 6))
        t6 = t6 + Val(Grid3.TextMatrix(i, 7))
        t7 = t7 + Val(Grid3.TextMatrix(i, 8))
        t8 = t8 + Val(Grid3.TextMatrix(i, 9))
    Next i
    i = Grid3.Rows - 1
    Grid3.TextMatrix(i, 2) = t1
    Grid3.TextMatrix(i, 3) = t2
    Grid3.TextMatrix(i, 4) = t3
    Grid3.TextMatrix(i, 5) = t4
    Grid3.TextMatrix(i, 6) = t5
    Grid3.TextMatrix(i, 7) = t6
    Grid3.TextMatrix(i, 8) = t7
    Grid3.TextMatrix(i, 9) = t8
    
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    Label2.Caption = List1
End Sub

Private Sub Combo2_Click()
    List2.ListIndex = Combo2.ListIndex
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub Command2_Click()
    Dim i As Integer, c As Integer, v As Long
    refresh_grid3
    If Combo1.ListIndex = 0 Then refresh_grid
    For i = 0 To Combo1.ListCount - 1
        Combo1.ListIndex = i
        DoEvents
    Next i
    For i = Grid3.Rows - 1 To 1 Step -1
        If Val(Grid3.TextMatrix(i, 9)) = 0 Then
            If Grid3.Rows > 2 Then
                Grid3.RemoveItem i
            Else
                Grid3.Rows = 1
            End If
        End If
    Next i
    Grid3.FillStyle = flexFillRepeat
    For i = 1 To Grid3.Rows - 1
        c = 2
        v = 0
        For k = 2 To 8
            If Val(Grid3.TextMatrix(i, k)) > v Then
                v = Val(Grid3.TextMatrix(i, k))
                c = k
            End If
        Next k
        Grid3.Row = i: Grid3.RowSel = i
        Grid3.Col = c: Grid3.ColSel = c
        Grid3.CellBackColor = hcolor.BackColor
        Grid3.CellForeColor = hcolor.ForeColor
    Next i
    Grid3.Row = 1: Grid3.Col = 2
End Sub

Private Sub Command3_Click()
    Dim rt As String, rf As String, rh As String
    rt = Me.Caption & " - " & Label5.Caption
    rh = datelit.Caption
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "white": gndc(0) = Me.Grid1.BackColorFixed
    htdc(1) = "yellow": gndc(1) = Me.Grid1.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid1.BackColor
    Grid3.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid3, rt, rh, rf, "linen", "khaki", "white")
        Grid3.Redraw = True
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid3, rt, rh, rf, "linen", "khaki", "white")
        Grid3.Redraw = True
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub datelit_Change()
    refresh_grid3
End Sub

Private Sub Form_Load()
    Text1 = bimp_sales_days                                 'jv050117
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    'Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    Text2 = Format(Now, "M-dd-yyyy")
    refresh_vlists
    'refresh_grid3
End Sub

Private Sub Form_Resize()
    'Grid1.Width = Me.Width - 100
    'If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 5)
    'Grid2.Height = Grid1.Height
End Sub

Private Sub Label2_Change()
    refresh_grid
End Sub

Private Sub List2_Click()
    Label5.Caption = List2
End Sub
