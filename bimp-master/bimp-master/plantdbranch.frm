VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form plantdbranch 
   Caption         =   "Plant Branches"
   ClientHeight    =   8565
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
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
      Left            =   8760
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   10200
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   10200
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   1200
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   600
      Width           =   3015
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
      Left            =   1200
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12726
      _Version        =   327680
      BackColorFixed  =   12648384
      FocusRect       =   0
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
      Left            =   4560
      TabIndex        =   6
      Top             =   600
      Width           =   3375
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
      Left            =   4560
      TabIndex        =   5
      Top             =   240
      Width           =   3255
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
      Left            =   120
      TabIndex        =   2
      Top             =   600
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
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu delbranch 
         Caption         =   "Drop Branch"
      End
   End
End
Attribute VB_Name = "plantdbranch"
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
    For i = 1 To 99
        If branchrec(i).oraloc > " " Then
            Combo2.AddItem Format(branchrec(i).branchno, "000") & "-" & branchrec(i).branchname
            List2.AddItem branchrec(i).branchname
        End If
    Next i
End Sub

Sub refresh_grid()
    Dim ds As adodb.Recordset, s As String
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 3
    s = s & "select plantwhs, branchwhs, count(*) from bimp where plantwhs = '" & Combo1 & "'"
    s = s & " group by plantwhs, branchwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!plantwhs & Chr(9)
            s = s & ds!branchwhs & Chr(9)
            s = s & branchrec(Val(ds!branchwhs)).branchname
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    Grid1.FormatString = "^Plant|^Branch|<Location"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 4500
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
End Sub

Private Sub Command1_Click()
    Dim s As String, ds As adodb.Recordset, i As Integer
    For i = 0 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 1) = Left(Combo2, 3) Then
            MsgBox Combo2 & " is already on the list.", vbOKOnly + vbInformation, "add failed.."
            Exit Sub
        End If
    Next i
    s = "Add " & List2 & " to " & List1 & " Branch List?"
    If MsgBox(s, vbYesNo + vbQuestion, "are you sure....") = vbNo Then Exit Sub
    
    
    Screen.MousePointer = 11
    s = "Select sku, count(*) from bimp where plantwhs = '" & Combo1 & "'"
    s = s & " group by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Insert into bimp (id, plantwhs, branchwhs, sku, onhand, onorder, sales, undiff, paldiff"
            s = s & ", ohpct, roqty, pctgain, needqty, bimpstatus, promoqty, lowqty, outqty, quotapct, plantpool"
            s = s & ", poolsched, discflag, promoflag, lowflag, outflag, lastrecpt, skunotes) Values ("
            s = s & wd_seq("Bimp")                              'ID
            s = s & ", '" & Combo1 & "'"                        'plantwhs
            s = s & ", '" & Left(Combo2, 3) & "'"               'branchwhs
            s = s & ", '" & ds!sku & "'"                        'sku
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
            s = s & ", " & plant_lowstock(Combo1, ds!sku)       'lowqty                 'jv071117
            's = s & ", 0"                                       'outqty
            s = s & ", " & plant_outstock(Combo1, ds!sku)       'outqty                 'jv071117
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

Private Sub delbranch_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 1)) = 0 Then Exit Sub
    s = "Drop " & Grid1.TextMatrix(i, 2) & " from " & List1 & " Product List?"
    If MsgBox(s, vbYesNo + vbQuestion, "are you sure....") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    s = "Delete from bimp where plantwhs = '" & Combo1 & "'"
    s = s & " and branchwhs = '" & Grid1.TextMatrix(i, 1) & "'"
    Call delete_bimp_log(bimpuserid, Me.Caption, s)                 'jv030817
    'MsgBox s
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

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub
