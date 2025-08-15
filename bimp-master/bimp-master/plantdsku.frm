VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form plantdsku 
   Caption         =   "Plant SKUs"
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13530
   LinkTopic       =   "Form3"
   ScaleHeight     =   6450
   ScaleWidth      =   13530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Synch"
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
      Left            =   13440
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add SKU"
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
      Left            =   5280
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5415
      Left            =   0
      TabIndex        =   8
      Top             =   960
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9551
      _Version        =   327680
      ForeColor       =   8388608
      BackColorFixed  =   8454143
      FocusRect       =   0
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   7920
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
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
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   480
      Width           =   1335
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
      TabIndex        =   5
      Top             =   600
      Width           =   4935
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
      TabIndex        =   4
      Top             =   240
      Width           =   1935
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
      Width           =   1575
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
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu delsku 
         Caption         =   "Drop SKU"
      End
      Begin VB.Menu edoutq 
         Caption         =   "Out Quantity"
      End
      Begin VB.Menu edlowq 
         Caption         =   "Low Quantity"
      End
      Begin VB.Menu edrelease 
         Caption         =   "New Release"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "plantdsku"
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
    Combo1.AddItem "VENDOR": List1.AddItem "Vendor Items"
    Combo1.AddItem "DRY": List1.AddItem "Dry Storage Items"
    For i = 0 To 9999
        If skurec(i).pallet > 0 Or skurec(i).unit = "MISC" Then         'jv021616
            Combo2.AddItem skurec(i).sku
            List2.AddItem skurec(i).unit & " " & skurec(i).desc
        End If
    Next i
End Sub

Sub reset_zeros()
    Dim i As Integer, s As String
    Screen.MousePointer = 11
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 2)) > 0 Then
            s = "Update bimp set outqty = " & Val(Grid1.TextMatrix(i, 2))
            s = s & ", lowqty = " & Val(Grid1.TextMatrix(i, 3))
            s = s & " Where plantwhs = '" & Combo1 & "'"
            s = s & " and sku = '" & Grid1.TextMatrix(i, 0) & "'"
            'MsgBox s
            wdb.Execute s
        End If
    Next i
    refresh_grid
    Screen.MousePointer = 0
End Sub

Sub refresh_grid()
    Dim ds As adodb.Recordset, s As String
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    'Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    's = "Select sku, outqty, lowqty, promoqty, count(*) from bimp where plantwhs = '" & Combo1 & "'"
    's = s & " group by sku, outqty, lowqty, promoqty"
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    s = "Select sku, outqty, lowqty, count(*) from bimp where plantwhs = '" & Combo1 & "'"
    s = s & " group by sku, outqty, lowqty"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!outqty & Chr(9)
            s = s & ds!lowqty '& Chr(9)
            's = s & ds!promoqty
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    'Grid1.FormatString = "^SKU|<Product|^OutQty(Pallets)|^LowQty(Pallets)|^New Release (Pallets)"
    Grid1.FormatString = "^SKU|<Product|^OutQty(Pallets)|^LowQty(Pallets)"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 4500
    Grid1.ColWidth(2) = 2000
    Grid1.ColWidth(3) = 2000
    'Grid1.ColWidth(4) = 2000
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    Label3.Caption = List1
    refresh_grid
End Sub

Private Sub Combo2_Click()
    Dim i As Integer
    List2.ListIndex = Combo2.ListIndex
    Label4.Caption = List2
    For i = 0 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = Combo2 Then
            Grid1.Row = i
            Exit For
        End If
    Next i
End Sub

Private Sub Command1_Click()
    Dim s As String, ds As adodb.Recordset, i As Integer
    For i = 0 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = Combo2 Then
            MsgBox Combo2 & " is already on the list.", vbOKOnly + vbInformation, "add failed.."
            Exit Sub
        End If
    Next i
    s = "Add " & List2 & " to " & List1 & " Product List?"
    If MsgBox(s, vbYesNo + vbQuestion, "are you sure....") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    s = "Select branchwhs, count(*) from bimp where plantwhs = '" & Combo1 & "'"
    s = s & " group by branchwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Insert into bimp (id, plantwhs, branchwhs, sku, onhand, onorder, sales, undiff, paldiff"
            s = s & ", ohpct, roqty, pctgain, needqty, bimpstatus, promoqty, lowqty, outqty, quotapct, plantpool"
            s = s & ", poolsched, discflag, promoflag, lowflag, outflag, lastrecpt, skunotes) Values ("
            s = s & wd_seq("Bimp")                              'ID
            s = s & ", '" & Combo1 & "'"                        'plantwhs
            s = s & ", '" & ds!branchwhs & "'"                  'branchwhs
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
            s = s & ", " & plant_lowstock(Combo1, Combo2)       'lowqty             'jv071117
            's = s & ", 0"                                       'outqty
            s = s & ", " & plant_outstock(Combo1, Combo2)       'outqty             'jv071117
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
            ds.MoveNext
        Loop
    End If
    ds.Close
    refresh_grid
    Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
    reset_zeros
End Sub

Private Sub delsku_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 0)) = 0 Then Exit Sub
    s = "Drop " & Grid1.TextMatrix(i, 1) & " from " & List1 & " Product List?"
    If MsgBox(s, vbYesNo + vbQuestion, "are you sure....") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    s = "Delete from bimp where plantwhs = '" & Combo1 & "'"
    s = s & " and sku = '" & Grid1.TextMatrix(i, 0) & "'"
    Call delete_bimp_log(bimpuserid, Me.Caption, s)             'jv030817
    'MsgBox s
    wdb.Execute s
    refresh_grid
    If i < Grid1.Rows Then Grid1.Row = i
    Screen.MousePointer = 0
End Sub

Private Sub edlowq_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 0)) = 0 Then Exit Sub
    s = Grid1.TextMatrix(i, 3)
    s = InputBox("Pallet Qty:", "Low Stock Pallet Qty...", s)
    If Len(s) = 0 Then Exit Sub
    Screen.MousePointer = 11
    s = "Update bimp set lowqty = " & Val(s)
    s = s & " Where plantwhs = '" & Combo1 & "'"
    s = s & " and sku = '" & Grid1.TextMatrix(i, 0) & "'"
    'MsgBox s
    wdb.Execute s
    refresh_grid
    If i < Grid1.Rows Then Grid1.Row = i
    Screen.MousePointer = 0
End Sub

Private Sub edoutq_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 0)) = 0 Then Exit Sub
    s = Grid1.TextMatrix(i, 2)
    s = InputBox("Pallet Qty:", "Out of Stock Pallet Qty...", s)
    If Len(s) = 0 Then Exit Sub
    Screen.MousePointer = 11
    s = "Update bimp set outqty = " & Val(s)
    s = s & " Where plantwhs = '" & Combo1 & "'"
    s = s & " and sku = '" & Grid1.TextMatrix(i, 0) & "'"
    'MsgBox s
    wdb.Execute s
    refresh_grid
    If i < Grid1.Rows Then Grid1.Row = i
    Screen.MousePointer = 0
End Sub

Private Sub edrelease_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 0)) = 0 Then Exit Sub
    s = InputBox("Pallet Qty:", "New Release Pallet Qty...", 2)
    If Len(s) = 0 Then Exit Sub
    Screen.MousePointer = 11
    s = "Update bimp set promoqty = " & Val(s)
    s = s & " Where plantwhs = '" & Combo1 & "'"
    s = s & " and sku = '" & Grid1.TextMatrix(i, 0) & "'"
    MsgBox s
    wdb.Execute s
    refresh_grid
    If i < Grid1.Rows Then Grid1.Row = i
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    refresh_vlists
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Command1.Height * 4.5)
End Sub

Private Sub Grid1_DblClick()
    Call edoutq_Click
    DoEvents
    Call edlowq_Click
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub
