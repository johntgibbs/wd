VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form branchostk 
   Caption         =   "Over Stocked Items"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   14820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
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
      Left            =   7680
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   12726
      _Version        =   327680
      ForeColor       =   32768
      BackColorFixed  =   12648384
      FocusRect       =   0
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
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   6255
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
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "branchostk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim s As String, i As Integer
    'On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 11: Grid1.Rows = 1
    Grid1.FixedCols = 2
    Grid1.Clear
    
    For i = 1 To whseover.Grid1.Rows - 1
        If whseover.Grid1.TextMatrix(i, 2) = Label2.Caption Then
            s = whseover.Grid1.TextMatrix(i, 0) & Chr(9)
            s = s & whseover.Grid1.TextMatrix(i, 1) & Chr(9)
            s = s & whseover.Grid1.TextMatrix(i, 3) & Chr(9)
            s = s & whseover.Grid1.TextMatrix(i, 4) & Chr(9)
            s = s & whseover.Grid1.TextMatrix(i, 5) & Chr(9)
            s = s & whseover.Grid1.TextMatrix(i, 6) & Chr(9)
            s = s & whseover.Grid1.TextMatrix(i, 7) & Chr(9)
            s = s & whseover.Grid1.TextMatrix(i, 8) & Chr(9)
            s = s & whseover.Grid1.TextMatrix(i, 9) & Chr(9)
            s = s & whseover.Grid1.TextMatrix(i, 12) & Chr(9)
            s = s & whseover.Grid1.TextMatrix(i, 13)
            Grid1.AddItem s
        End If
    Next i
    Screen.MousePointer = 0
    Grid1.FormatString = "^SKU|<Product|^OnHand|^OnOrder|^Sales|^UnitDiff|^PalletDiff|^OH%|^Days InStock|^Source|^Days Supply"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 3200 '4000
    Grid1.ColWidth(2) = 1100
    Grid1.ColWidth(3) = 1100
    Grid1.ColWidth(4) = 1100
    Grid1.ColWidth(5) = 1100
    Grid1.ColWidth(6) = 1100
    Grid1.ColWidth(7) = 1100
    Grid1.ColWidth(8) = 1300
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 1300
    If Grid1.Rows > 1 Then Grid1.Row = 1
    Grid1.Redraw = True
End Sub


Private Sub Command1_Click()
    Dim rt As String, rf As String, rh As String
    rt = Me.Caption
    rh = "Days of Supply >= " & whseover.Text1 & "    Branch: " & Label2.Caption
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "white": gndc(0) = Me.Grid1.BackColorFixed
    htdc(1) = "yellow": gndc(1) = Me.Grid1.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid1.BackColor
    Grid1.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Me.Top = bimpbanner.Label2.Top
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Label1.Height * 3.5)
End Sub

Private Sub Label2_Change()
    refresh_grid
End Sub

