VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form planskuship 
   Caption         =   "Planned Shipments"
   ClientHeight    =   10365
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   10365
   ScaleWidth      =   15225
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6495
      Left            =   -120
      TabIndex        =   8
      Top             =   600
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11456
      _Version        =   327680
      Cols            =   8
      FixedCols       =   3
      BackColorFixed  =   12648447
      WordWrap        =   -1  'True
      FocusRect       =   0
      GridLines       =   2
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   6240
      TabIndex        =   5
      Top             =   7440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   3720
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   5880
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   120
      Width           =   1095
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
      Left            =   840
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Next Week:"
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
      Left            =   11280
      TabIndex        =   18
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "This Week:"
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
      Left            =   11280
      TabIndex        =   17
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label nend 
      Caption         =   "nend"
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
      Left            =   13800
      TabIndex        =   16
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label nstart 
      Caption         =   "nstart"
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
      Left            =   12480
      TabIndex        =   15
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label cend 
      Caption         =   "cend"
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
      Left            =   13800
      TabIndex        =   14
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label cstart 
      Caption         =   "cstart"
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
      Left            =   12480
      TabIndex        =   13
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "gcolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Caption         =   "bcolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ycolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "wcolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   9600
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Left            =   7080
      TabIndex        =   7
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
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
      Left            =   2040
      TabIndex        =   6
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Branch:"
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
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Plant:"
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
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu procmenu 
      Caption         =   "Process"
      Begin VB.Menu procweeks 
         Caption         =   "Move Next Week Pallets to Current"
      End
   End
End
Attribute VB_Name = "planskuship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refresh_vlists()
    Combo1.Clear: List1.Clear
    Combo2.Clear: List2.Clear
    For i = 50 To 52
        Combo1.AddItem plantrec(i).orawhs
        List1.AddItem plantrec(i).plantname
    Next i
    For i = 1 To 99
        If branchrec(i).oraloc > " " Then
            Combo2.AddItem Format(Val(branchrec(i).branchno), "000")
            List2.AddItem branchrec(i).branchname
        End If
    Next i
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
End Sub

Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, s As String, i As Integer, ctot As Integer, ntot As Integer
    Dim atot1 As Integer, atot2 As Integer
    If Combo1 < " " Then Exit Sub
    If Combo2 < " " Then Exit Sub
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
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 13
    s = "select id, sku, onhand, onorder, sales, undiff, paldiff, bimpstatus, thiswknewpals, nextwknewpals"
    s = s & " from bimp where plantwhs = '" & Combo1 & "' and branchwhs = '" & Combo2 & "'"
    s = s & " order by sku"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!onhand & Chr(9)
            s = s & ds!onorder & Chr(9)
            s = s & ds!sales & Chr(9)
            s = s & ds!undiff & Chr(9)
            s = s & ds!paldiff & Chr(9)
            s = s & ds!bimpstatus & Chr(9)
            s = s & Chr(9)
            s = s & Format(ds!thiswknewpals, "#") & Chr(9)
            s = s & Chr(9)
            s = s & Format(ds!nextwknewpals, "#")
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        If Combo1 = "T10" Then pplant = 50
        If Combo1 = "K10" Then pplant = 51
        If Combo1 = "A10" Then pplant = 52
        s = "select runid, sku, pallets, shipdate from trailers where plant = " & pplant
        s = s & " and branch = " & Val(Combo2)
        s = s & " and shipdate >= '" & Format(Now, "MM-dd-yyyy") & "'"
        'MsgBox s
        Set ds = wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ticket_post(ds!runid) = False Then
                    For i = 1 To Grid1.Rows - 1
                        If Grid1.TextMatrix(i, 1) = ds!sku Then
                            If Format(ds!shipdate, "yyyyMMdd") <= Format(cend, "yyyyMMdd") Then
                                Grid1.TextMatrix(i, 9) = Val(Grid1.TextMatrix(i, 9)) + ds!pallets
                            Else
                                Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) + ds!pallets
                            End If
                        End If
                    Next i
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        ntot = 0: ctot = 0: atot1 = 0: atot2 = 0
        For i = 1 To Grid1.Rows - 1
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 3: Grid1.ColSel = 8
            If Grid1.TextMatrix(i, 8) = "W" Then Grid1.CellBackColor = wcolor.BackColor
            If Grid1.TextMatrix(i, 8) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
            If Grid1.TextMatrix(i, 8) = "B" Then Grid1.CellBackColor = bcolor.BackColor
            If Grid1.TextMatrix(i, 8) = "G" Then Grid1.CellBackColor = gcolor.BackColor
            atot1 = atot1 + Val(Grid1.TextMatrix(i, 9))
            ctot = ctot + Val(Grid1.TextMatrix(i, 10))
            atot2 = atot2 + Val(Grid1.TextMatrix(i, 11))
            ntot = ntot + Val(Grid1.TextMatrix(i, 12))
        Next i
        s = Chr(9) & Chr(9) & "Total Pallets" & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9)
        s = s & atot1 & Chr(9) & ctot & Chr(9) & atot2 & Chr(9) & ntot
        's = s & "ctot" & Chr(9) & "ntot"
        Grid1.AddItem s
        Grid1.Row = 1: Grid1.Col = 3
        Grid1.RowHeight(0) = Grid1.RowHeight(1) * 2
    End If
    'Grid1.FormatString = "^ID|^SKU|<Product|^On Hand|^On Order|^Sales|^Unit Diff|^Pallet Diff|^Status|^This Week (Pallets)|^Next Week (Pallets)"
    Grid1.FormatString = "^ID|^SKU|<Product|^On Hand|^On Order|^Sales|^Unit Diff|^Pallet Diff||^Active Trailers|^This Week (Pallets)|^Active Trailers|^Next Week (Pallets)"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3500
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 0 '1000
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 1200
    Grid1.ColWidth(11) = 1000
    Grid1.ColWidth(12) = 1200
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    Label3.Caption = List1
    refresh_grid
End Sub

Private Sub Combo2_Click()
    List2.ListIndex = Combo2.ListIndex
    Label4.Caption = List2
    refresh_grid
End Sub

Private Sub Form_Load()
    Dim pday As Integer
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    refresh_vlists
    cstart = Format(Now, "ddd")
    If cstart = "Sun" Then pday = 0
    If cstart = "Mon" Then pday = -1
    If cstart = "Tue" Then pday = -2
    If cstart = "Wed" Then pday = -3
    If cstart = "Thu" Then pday = -4
    If cstart = "Fri" Then pday = -5
    If cstart = "Sat" Then pday = -6
    cstart = Format(DateAdd("d", pday, Now), "M-dd-yyyy")
    cend = Format(DateAdd("d", 6, cstart), "M-dd-yyyy")
    nstart = Format(DateAdd("d", 1, cend), "M-dd-yyyy")
    nend = Format(DateAdd("d", 6, nstart), "M-dd-yyyy")
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 140
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 4)
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim i As Integer, cp As Long, np As Long
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    If Grid1.Col = 10 Or Grid1.Col = 12 Then
        If edcol = True Then
            Grid1.Text = ""
            edcol = False
        End If
        If KeyAscii = 8 Then
            If Len(Grid1.Text) > 1 Then
                Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
            Else
                Grid1.Text = ""
            End If
        End If
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            Grid1.Text = Grid1.Text & Chr(KeyAscii)
        End If
        If Grid1.Col = 10 Then
            s = "Update bimp set thiswknewpals = " & Val(Grid1.Text) & " where id = " & Grid1.TextMatrix(Grid1.Row, 0)
            'MsgBox s
            wdb.Execute s
        End If
        If Grid1.Col = 12 Then
            s = "Update bimp set nextwknewpals = " & Val(Grid1.Text) & " where id = " & Grid1.TextMatrix(Grid1.Row, 0)
            'MsgBox s
            wdb.Execute s
        End If
        cp = 0: np = 0
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 0)) > 0 Then
                cp = cp + Val(Grid1.TextMatrix(i, 10))
                np = np + Val(Grid1.TextMatrix(i, 12))
            End If
        Next i
        Grid1.TextMatrix(Grid1.Rows - 1, 10) = Format(cp, "#")
        Grid1.TextMatrix(Grid1.Rows - 1, 12) = Format(np, "#")
    End If
End Sub

Private Sub procweeks_Click()
    Dim s As String, i As Integer
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 0)) > 0 Then
            s = "Update bimp set thiswknewpals = " & Val(Grid1.TextMatrix(i, 12))
            s = s & ", nextwknewpals = 0 where id = " & Grid1.TextMatrix(i, 0)
            'MsgBox s
            wdb.Execute s
        End If
    Next i
    refresh_grid
End Sub
