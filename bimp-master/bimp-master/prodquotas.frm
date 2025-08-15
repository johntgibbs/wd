VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form prodquotas 
   Caption         =   "Product Quotas"
   ClientHeight    =   8100
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13530
   LinkTopic       =   "Form2"
   ScaleHeight     =   8100
   ScaleWidth      =   13530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Post Pcts."
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
      Left            =   11040
      TabIndex        =   15
      Top             =   720
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   6375
      Left            =   8160
      TabIndex        =   13
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   11245
      _Version        =   327680
      BackColorFixed  =   16777152
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Branch"
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
      Left            =   6000
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6375
      Left            =   0
      TabIndex        =   9
      Top             =   1080
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11245
      _Version        =   327680
      FocusRect       =   0
   End
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   9840
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   11280
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
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
      Left            =   840
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   600
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   11280
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
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
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Sales History"
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
      Left            =   8280
      TabIndex        =   14
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label pcterr 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Allocated"
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
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label pctok 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "100% Allocated"
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
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label4 
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
      Left            =   2400
      TabIndex        =   7
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label3 
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
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "SKU:"
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
      TabIndex        =   1
      Top             =   600
      Width           =   735
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
      Width           =   1575
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu delrec 
         Caption         =   "Delete Record"
      End
      Begin VB.Menu updrec 
         Caption         =   "Update Pct."
      End
   End
   Begin VB.Menu procmenu 
      Caption         =   "Process"
      Begin VB.Menu procskusales 
         Caption         =   "Process Plant SKU Sales Pcts."
      End
   End
End
Attribute VB_Name = "prodquotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub process_sales_pcts()
    Dim ds As adodb.Recordset, s As String, i As Integer
    s = "select sku, count(*) from bimp where plantwhs = '" & Combo1 & "' group by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 0 To Combo2.ListCount - 1
                If Combo2.List(i) = ds!sku Then
                    Combo2.ListIndex = i
                    DoEvents
                    'MsgBox "Check"
                    Command2_Click
                    DoEvents
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Sub refresh_grid()
    Dim s As String, ds As adodb.Recordset, i As Integer, p As Currency, c As Integer
    c = Grid1.Row
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 3
    pcterr.Visible = False: pctok.Visible = False
    s = "select * from bimp where plantwhs = '" & Combo1 & "'"
    s = s & " and sku = '" & Combo2 & "'"
    s = s & " order by branchwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            For i = 0 To List3.ListCount
                If Left(List3.List(i), 3) = ds!branchwhs Then
                    s = s & List3.List(i)
                    Exit For
                End If
            Next i
            s = s & Chr(9)
            s = s & Format(ds!quotapct, "0.000")
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    If Grid1.Rows > 1 Then
        p = 0
        For i = 1 To Grid1.Rows - 1
            p = p + Val(Grid1.TextMatrix(i, 2))
        Next i
        pcterr = p & "% Allocated"
        If Int(p) <> 100 Then
            pcterr.Visible = True
        Else
            pctok.Visible = True
        End If
    End If
    Grid1.FormatString = "^ID|<Branch|^Quota Pct."
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 2000
    If c < Grid1.Rows Then Grid1.Row = c
    Grid1.Redraw = True
End Sub

Sub refresh_grid2()
    Dim ds As adodb.Recordset, s As String, ts As Long, tp As Currency
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 3
    s = "select * from bimp where sku = '" & Combo2 & "'"
    s = s & " and plantwhs = '" & Combo1 & "'"
    s = s & " and sales > 0"
    s = s & " order by branchwhs"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!branchwhs & Chr(9)
            s = s & ds!sales
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid2.Rows > 1 Then
        ts = 0: tp = 0
        For i = 1 To Grid2.Rows - 1
            ts = ts + Val(Grid2.TextMatrix(i, 1))
        Next i
        For i = 1 To Grid2.Rows - 1
            Grid2.TextMatrix(i, 2) = Format((Val(Grid2.TextMatrix(i, 1)) / ts) * 100, ".000")
            tp = tp + Val(Grid2.TextMatrix(i, 2))
        Next i
        Grid2.AddItem "Total" & Chr(9) & ts & Chr(9) & Format(tp, ".000")
    End If
    Grid2.FormatString = "^Whs|^Sales|^Pct."
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1400
    Grid2.ColWidth(2) = 1400
    Grid2.Redraw = True
End Sub

Sub refresh_vlists()
    Combo1.Clear: List1.Clear
    Combo2.Clear: List2.Clear
    List3.Clear
    For i = 50 To 52
        Combo1.AddItem plantrec(i).orawhs
        List1.AddItem plantrec(i).plantname
    Next i
    Combo1.AddItem "VENDOR": List1.AddItem "Vendor Items"
    Combo1.AddItem "DRY": List1.AddItem "Dry Storage Items"
    For i = 0 To 9999
        If skurec(i).pallet > 0 Then
            Combo2.AddItem skurec(i).sku
            List2.AddItem skurec(i).unit & " " & skurec(i).desc
        End If
    Next i
    For i = 1 To 99
        If branchrec(i).oraloc > " " Then
            List3.AddItem Format(Val(branchrec(i).branchno), "000") & "-" & branchrec(i).branchname
        End If
    Next i
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    Label3.Caption = List1
    refresh_grid
    refresh_grid2
End Sub

Private Sub Combo2_Click()
    List2.ListIndex = Combo2.ListIndex
    Label4.Caption = List2
    refresh_grid
    refresh_grid2
End Sub

Private Sub Command1_Click()
    Dim s As String, i As Integer, nrec As Boolean, b As String
    b = InputBox("Branch Code:", "add branch...", "001")
    If Len(b) = 0 Then Exit Sub
    b = Format(Val(b), "000")
    If Val(b) = 0 Then Exit Sub
    nrec = False
    For i = 0 To List3.ListCount - 1
        If b = Left(List3.List(i), 3) Then
            nrec = True
            Exit For
        End If
    Next i
    If nrec = False Then
        MsgBox b & " branch not on file...", vbOKOnly + vbInformation, "sorry, try again...."
    End If
    s = "Insert into bimp (id, plantwhs, branchwhs, sku, onhand, onorder, sales, undiff, paldiff"
    s = s & ", ohpct, roqty, pctgain, needqty, bimpstatus, promoqty, lowqty, outqty, quotapct, plantpool"
    s = s & ", poolsched, discflag, promoflag, lowflag, outflag, lastrecpt, skunotes) Values ("
    s = s & wd_seq("Bimp")                              'ID
    s = s & ", '" & Combo1 & "'"                        'plantwhs
    s = s & ", '" & b & "'"                             'branchwhs
    s = s & ", '" & Combo2 & "'"                        'sku
    s = s & ", 0"                                       'onhand
    s = s & ", 0"                                       'onorder
    s = s & ", 0"                                       'sales
    s = s & ", 0"                                       'undiff
    s = s & ", 0"                                       'paldiff
    s = s & ", 0"                                       'ohpct
    s = s & ", " & skurec(Val(Combo2)).pallet           'roqty
    s = s & ", 0"                                       'pctgain
    s = s & ", 0"                                       'needqty
    s = s & ", 'B'"                                     'bimpstatus
    s = s & ", 0"                                       'promoqty
    's = s & ", 0"                                       'lowqty
    s = s & ", " & plant_lowstock(Combo1, Combo2)       'lowqty                 'jv071117
    's = s & ", 0"                                       'outqty
    s = s & ", " & plant_outstock(Combo1, Combo2)       'outqty                 'jv071117
    s = s & ", 0"                                       'quotapct
    s = s & ", 0"                                       'plantpool
    s = s & ", 0"                                       'poolsched
    s = s & ", 'N'"                                     'discflag
    s = s & ", 'N'"                                     'promoflag
    s = s & ", 'N'"                                     'lowflag
    s = s & ", 'N'"                                     'outflag
    s = s & ", '" & Format(Now, "m-d-yyyy") & "'"       'lastrcpt
    s = s & ", ' ')"                                    'skunotes
    'MsgBox s
    wdb.Execute s
    refresh_grid
End Sub

Private Sub Command2_Click()
    Dim i As Integer, s As String
    For i = 1 To Grid2.Rows - 1
        If Val(Grid2.TextMatrix(i, 0)) > 0 Then
            s = "Update bimp set quotapct = " & Grid2.TextMatrix(i, 2)
            s = s & " Where branchwhs = '" & Grid2.TextMatrix(i, 0) & "'"
            s = s & " and plantwhs = '" & Combo1 & "'"
            s = s & " and sku = '" & Combo2 & "'"
            'MsgBox s
            wdb.Execute s
        End If
    Next i
    refresh_grid
End Sub

Private Sub delrec_Click()
    Dim s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    s = "Delete " & Combo2 & " from " & Grid1.TextMatrix(Grid1.Row, 1)
    Call delete_bimp_log(bimpuserid, Me.Caption, s)                         'jv030817
    s = "Delete from bimp where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    'MsgBox s
    wdb.Execute s
    refresh_grid
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
    If Me.Height > 2000 Then
        Grid1.Height = Me.Height - (Command1.Height * 4.5)
        Grid2.Height = Grid1.Height
    End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub procskusales_Click()
    process_sales_pcts
End Sub

Private Sub updrec_Click()
    Dim s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    s = InputBox("Pct. Qty:", "Quota pct....", Grid1.TextMatrix(Grid1.Row, 2))
    If Len(s) = 0 Then Exit Sub
    s = "Update bimp set quotapct = " & Format(Val(s), "0.000")
    s = s & " where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    'MsgBox s
    wdb.Execute s
    refresh_grid
End Sub
