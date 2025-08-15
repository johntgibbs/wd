VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brbarcodes 
   Caption         =   "Ticket BarCodes"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
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
      Left            =   9480
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6615
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   11668
      _Version        =   327680
      ForeColor       =   16384
      BackColorFixed  =   12648447
      FocusRect       =   0
   End
   Begin VB.Label tktkey 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "Branch:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "brbarcodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim s As String, f1 As String, f2 As String, f3 As String, f4 As String, f0 As String
    Dim t As String, k As String, i As Integer
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\brbarcodes.txt"
    Open cfile For Input Shared As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4
        k = Format(DateDiff("d", f1, Now), "0") & f0 & f4
        'If Left(f2, Len(tktkey.Caption)) = UCase(tktkey.Caption) Then
        If Left(f2, Len(f2) - 3) = UCase(tktkey.Caption) Then
            If Format(DateDiff("d", f1, Now), "0") & f0 <> t Then
                If Grid1.Rows > 1 Then Grid1.AddItem " " & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & t & "99999"
                t = Format(DateDiff("d", f1, Now), "0") & f0
            End If
            bc = Mid(f4, 5, 6) & "  " & Mid(f4, 11, 3) & "  " & Mid(f4, 14, 3)
            's = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & f4 & Chr(9) & k
            s = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & bc & Chr(9) & k
            Grid1.AddItem s
        End If
    Loop
    If Grid1.Rows > 1 Then Grid1.AddItem " " & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & t & "99999"
    Close #1
    If Grid1.Rows > 1 Then
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 5: Grid1.ColSel = 5
        Grid1.Sort = 5
        For i = Grid1.Rows - 2 To 1 Step -1
            If Grid1.TextMatrix(i, 5) <> t Then
                t = Grid1.TextMatrix(i, 5)
            Else
                Grid1.RemoveItem i
            End If
        Next i
    End If
    'Grid1.FormatString = "^Ticket|^Ship Date|<Trailer|<Product|^BarCode"
    Grid1.FormatString = "^Group|^Loaded|<Trailer|<Product|^BarCode"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 2400
    Grid1.ColWidth(3) = 3500
    Grid1.ColWidth(4) = 2000
    Grid1.ColWidth(5) = 0 '3000
    Grid1.Redraw = True
End Sub

Private Sub Command1_Click()
    export_branchbarcodes_ships
    'export_branchbarcodes_bills
    DoEvents
    refresh_grid
End Sub

Private Sub Form_Load()
    Me.Height = trlstatus.Height
    Me.Top = trlstatus.Top
    'Me.Left = trlstatus.Width - Me.Width
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (tktkey.Height * 4)
End Sub

Private Sub tktkey_Change()
    refresh_grid
End Sub

