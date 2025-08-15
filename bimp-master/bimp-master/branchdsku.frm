VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form branchdsku 
   Caption         =   "Branch SKUs"
   ClientHeight    =   12660
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13575
   LinkTopic       =   "Form4"
   ScaleHeight     =   12660
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Height          =   450
      Left            =   8880
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   6960
      TabIndex        =   11
      Text            =   "Combo3"
      Top             =   120
      Width           =   1335
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
      Left            =   10800
      TabIndex        =   9
      Top             =   480
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   8760
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   8760
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
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
      Left            =   960
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   480
      Width           =   1575
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
      Left            =   960
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   14631
      _Version        =   327680
      ForeColor       =   8421376
      BackColorFixed  =   12632319
      FocusRect       =   0
   End
   Begin VB.Label Label6 
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
      Left            =   8520
      TabIndex        =   13
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label5 
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
      Left            =   5640
      TabIndex        =   10
      Top             =   120
      Width           =   1215
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
      Left            =   2760
      TabIndex        =   6
      Top             =   600
      Width           =   2535
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
      Left            =   2760
      TabIndex        =   5
      Top             =   240
      Width           =   2535
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
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu delrec 
         Caption         =   "Drop SKU"
      End
      Begin VB.Menu edpq 
         Caption         =   "Update Release Qty"
      End
      Begin VB.Menu ednotes 
         Caption         =   "Update SKU Notes"
      End
   End
End
Attribute VB_Name = "branchdsku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refresh_vlists()
    Combo1.Clear: List1.Clear
    Combo2.Clear: List2.Clear
    Combo3.Clear: List3.Clear
    For i = 1 To 99
        If branchrec(i).oraloc > " " Then
            Combo1.AddItem Format(Val(branchrec(i).branchno), "000")
            List1.AddItem branchrec(i).branchname
        End If
    Next i
    For i = 0 To 9999
        'If skurec(i).pallet > 0 Then
        If skurec(i).wrapunits > 0 Then
            Combo2.AddItem skurec(i).sku
            List2.AddItem skurec(i).desc
        End If
    Next i
    For i = 50 To 52
        Combo3.AddItem plantrec(i).orawhs
        List3.AddItem plantrec(i).plantname
    Next i
    Combo3.AddItem "VENDOR": List3.AddItem "Vendor Items"
    Combo3.AddItem "DRY": List3.AddItem "Dry Storage Items"
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Combo3.ListIndex = 0
End Sub

Sub refresh_grid()
    Dim ds As ADODB.Recordset, s As String
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    s = "Select id, plantwhs, sku, promoqty, discflag, skunotes from bimp where branchwhs = '" & Combo1 & "'"
    s = s & " order by sku, plantwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!plantwhs & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!promoqty & Chr(9)
            s = s & ds!discflag & Chr(9)
            s = s & ds!skunotes
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    Grid1.FormatString = "^ID|^Source|^SKU|<Product|^New Release (Pallets)|^Discflag|<SKU Notes"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 3500
    Grid1.ColWidth(4) = 2200
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 3500
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
        If Grid1.TextMatrix(i, 2) = Combo2 Then
            Grid1.Row = i
            Exit For
        End If
    Next i
End Sub

Private Sub Combo3_Click()
    List3.ListIndex = Combo3.ListIndex
    Label6.Caption = List3
End Sub

Private Sub Command1_Click()
    Dim s As String, ds As ADODB.Recordset, i As Integer
    For i = 0 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 1) = Combo3 And Grid1.TextMatrix(i, 2) = Combo2 Then
            MsgBox Combo2 & " is already on the list for " & List3 & ".", vbOKOnly + vbInformation, "add failed.."
            Exit Sub
        End If
    Next i
    s = "Add " & List2 & " to " & List1 & " Product List?"
    If MsgBox(s, vbYesNo + vbQuestion, "are you sure....") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    s = "Insert into bimp (id, plantwhs, branchwhs, sku, onhand, onorder, sales, undiff, paldiff"
    s = s & ", ohpct, roqty, pctgain, needqty, bimpstatus, promoqty, lowqty, outqty, quotapct, plantpool"
    s = s & ", poolsched, discflag, promoflag, lowflag, outflag, lastrecpt, skunotes) Values ("
    s = s & wd_seq("Bimp")                              'ID
    s = s & ", '" & Combo3 & "'"                        'plantwhs
    s = s & ", '" & Combo1 & "'"                        'branchwhs
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
    s = s & ", " & plant_lowstock(Combo3, Combo2)       'lowqty                     'jv071117
    's = s & ", 0"                                       'outqty
    s = s & ", " & plant_outstock(Combo3, Combo2)       'outqty                     'jv071117
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
    i = Grid1.Row
    refresh_grid
    If i < Grid1.Rows Then Grid1.Row = i
    Screen.MousePointer = 0
End Sub

Private Sub delrec_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 0)) = 0 Then Exit Sub
    s = "Drop " & Grid1.TextMatrix(i, 1) & " " & Grid1.TextMatrix(i, 3) & " from " & List1 & " Product List?"
    If MsgBox(s, vbYesNo + vbQuestion, "are you sure....") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    Call delete_bimp_log(bimpuserid, Me.Caption, s)                     'jv030817
    s = "Delete from bimp where id = " & Grid1.TextMatrix(i, 0)
    'MsgBox s
    wdb.Execute s
    refresh_grid
    If i < Grid1.Rows Then Grid1.Row = i
    Screen.MousePointer = 0
End Sub

Private Sub ednotes_Click()
    Dim s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    If Len(Grid1.TextMatrix(Grid1.Row, 6)) > 1 Then
        s = Grid1.TextMatrix(Grid1.Row, 6)
    Else
        s = "."
    End If
    s = InputBox("SKU Notes:", "SKU Notes..", s)
    If Len(s) = 0 Then Exit Sub
    Grid1.TextMatrix(Grid1.Row, 6) = s
    If s > "0" Then
        s = "Update bimp set skunotes = '" & s & "', discflag = 'B' where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Grid1.TextMatrix(Grid1.Row, 5) = "B"
    Else
        s = "Update bimp set skunotes = ' ', discflag = 'N' where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Grid1.TextMatrix(Grid1.Row, 5) = "N"
    End If
    'MsgBox s
    wdb.Execute s
End Sub

Private Sub edpq_Click()
    Dim s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    s = InputBox("New Release Qty:", "New Release Pallety Quantity..", 2)
    If Len(s) = 0 Then Exit Sub
    Grid1.TextMatrix(Grid1.Row, 4) = s
    If Val(s) > 0 Then
        s = "Update bimp set promoqty = " & Val(s) & ", promoflag = 'Y' where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Else
        s = "Update bimp set promoqty = 0, promoflag = 'N' where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    End If
    'MsgBox s
    wdb.Execute s
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    refresh_vlists
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Command1.Height * 4.5)
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

