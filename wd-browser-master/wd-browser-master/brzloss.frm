VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brzloss 
   Caption         =   "Branches Out of Stock"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12735
   LinkTopic       =   "Form15"
   ScaleHeight     =   9300
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6495
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   11456
      _Version        =   327680
      BackColorFixed  =   16777152
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
      Left            =   7080
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Supplier Warehouse:"
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
      Width           =   2055
   End
   Begin VB.Label plantkey 
      Caption         =   "Label1"
      Height          =   375
      Left            =   9720
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "brzloss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim s As String, daysin As Integer, daysout As Integer, lostsales As Long, salesperday As Long
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Grid1.Redraw = False: Grid1.Font = "Arial": Grid1.FontBold = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    'Open brzloss For Input As #2
    Open "S:\wd\html\boutstock.csv" For Input As #2
    Do Until EOF(2)
        Input #2, f0, f1, f2, f3, f4, f5, f6, f7
        If f3 = plantkey.Caption Then
            s = f0 & Chr(9)
            s = s & f1 & Chr(9)
            s = s & skurec(Val(f1)).unit & " " & skurec(Val(f1)).desc & Chr(9)
            s = s & f2 & Chr(9)
            s = s & f4 & "-" & branchrec(Val(f4)).branchname & Chr(9)
            s = s & f5 & Chr(9)
            s = s & f6 & Chr(9)
            s = s & f7 & Chr(9)
            daysin = DateDiff("d", f5, f6) + 1
            s = s & daysin & Chr(9)
            salesperday = f7 / daysin
            s = s & salesperday & Chr(9)
            daysout = DateDiff("d", f6, Now)
            s = s & daysout & Chr(9)
            'lostsales = f7 * (daysout / daysin)
            lostsales = salesperday * daysout
            s = s & lostsales
            If lostsales > 0 Then Grid1.AddItem s
        End If
    Loop
    Close #2
    s = "^ID|^SKU|<Product|^PalSize|<Branch|^Last Receipt|^Last Issue|^Sales|^Days In Stock|^Sales Per Day|^Days Out|^Lost Sales"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 0 '1000
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 2400
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 1200
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1400
    Grid1.ColWidth(9) = 1400
    Grid1.ColWidth(10) = 1200
    Grid1.ColWidth(11) = 1200
    Grid1.Redraw = True
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub Form_Load()
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
    Me.Width = Form1.Width
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 200
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Command1.Height * 2.5)
End Sub

Private Sub plantkey_Change()
    If plantkey.Caption = "T10" Then Label2.Caption = "T10-Brenham"
    If plantkey.Caption = "K10" Then Label2.Caption = "K10-Broken Arrow"
    If plantkey.Caption = "A10" Then Label2.Caption = "A10-Sylacauga"
    refresh_grid
End Sub

