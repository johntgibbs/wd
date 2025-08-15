VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10125
   LinkTopic       =   "Form11"
   ScaleHeight     =   7920
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9975
      _Version        =   327680
      FocusRect       =   0
   End
   Begin VB.Label rcolor 
      BackColor       =   &H00000080&
      Caption         =   "rcolor"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label lotyear 
      Caption         =   "Year:"
      Height          =   255
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu prtmenu 
      Caption         =   "&Print"
   End
   Begin VB.Menu edyear 
      Caption         =   "Change Year"
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim i As Integer, k As Integer, leapyear As Boolean
    'MsgBox lotyear Mod 4
    If lotyear Mod 4 = 0 Then leapyear = True
    Grid1.Clear: Grid1.Cols = 13: Grid1.Rows = 32
    k = 1
    'Jan
    For i = 1 To 31
        Grid1.TextMatrix(i, 0) = i
        Grid1.TextMatrix(i, 1) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    'Feb
    For i = 1 To 28
        Grid1.TextMatrix(i, 2) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    If leapyear = True Then
        Grid1.TextMatrix(29, 2) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    End If
    'Mar
    For i = 1 To 31
        Grid1.TextMatrix(i, 3) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    'Apr
    For i = 1 To 30
        Grid1.TextMatrix(i, 4) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    'May
    For i = 1 To 31
        Grid1.TextMatrix(i, 5) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    'Jun
    For i = 1 To 30
        Grid1.TextMatrix(i, 6) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    'Jul
    For i = 1 To 31
        Grid1.TextMatrix(i, 7) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    'Aug
    For i = 1 To 31
        Grid1.TextMatrix(i, 8) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    'Sep
    For i = 1 To 30
        Grid1.TextMatrix(i, 9) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    'Oct
    For i = 1 To 31
        Grid1.TextMatrix(i, 10) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    'Nov
    For i = 1 To 30
        Grid1.TextMatrix(i, 11) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    'Dec
    For i = 1 To 31
        Grid1.TextMatrix(i, 12) = Right(lotyear, 2) & Format(k, "000")
        k = k + 1
    Next i
    Grid1.FormatString = "^Day|^Jan|^Feb|^Mar|^Apr|^May|^Jun|^Jul|^Aug|^Sep|^Oct|^Nov|^Dec"
    For i = 0 To 12
        Grid1.ColWidth(i) = 700
    Next i
End Sub

Private Sub edyear_Click()
    Dim s As String
    s = InputBox("Year:", "Enter New Year....", lotyear)
    If Len(s) = 0 Then Exit Sub
    lotyear = Format(Val(s), "0000")
End Sub

Private Sub Form_Load()
    lotyear = Format(Now, "yyyy")
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 780
End Sub

Private Sub Grid1_Click()
    Dim r As Integer, c As Integer
    rcolor.Caption = "clicked"
    Grid1.Redraw = False
    Grid1.FillStyle = flexFillRepeat
    r = Grid1.Row: c = Grid1.Col
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = 0
        Grid1.CellFontBold = False
        Grid1.CellBackColor = Grid1.BackColorFixed
        Grid1.CellForeColor = Grid1.ForeColorFixed
    Next i
    For i = 1 To Grid1.Cols - 1
        Grid1.Row = 0: Grid1.RowSel = 0
        Grid1.Col = i: Grid1.ColSel = i
        Grid1.CellFontBold = False
        Grid1.CellBackColor = Grid1.BackColorFixed
        Grid1.CellForeColor = Grid1.ForeColorFixed
    Next i
    Grid1.Row = r: Grid1.RowSel = r
    Grid1.Col = 0: Grid1.ColSel = 0
    Grid1.CellFontBold = True
    Grid1.CellBackColor = rcolor.BackColor
    Grid1.CellForeColor = rcolor.ForeColor
    Grid1.Row = 0: Grid1.RowSel = 0
    Grid1.Col = c: Grid1.ColSel = c
    Grid1.CellFontBold = True
    Grid1.CellBackColor = rcolor.BackColor
    Grid1.CellForeColor = rcolor.ForeColor
    Grid1.Row = r: Grid1.Col = c
    Grid1.Redraw = True
    rcolor.Caption = "rcolor"
End Sub

Private Sub Grid1_RowColChange()
    If rcolor.Caption = "rcolor" Then Call Grid1_Click
End Sub

Private Sub lotyear_Change()
    Me.Caption = "Lot Code Table - " & lotyear
    refresh_grid
    Grid1_Click
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    rt = Me.Caption
    rh = " "
    rf = "Printed: " & Format(Now, "mmmm d, yyyy h:mm am/pm")
    Call printflexgrid(Printer, Grid1, rt, rh, rf)
End Sub
