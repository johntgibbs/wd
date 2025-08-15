VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form holdhist 
   Caption         =   "Hold Status History"
   ClientHeight    =   8085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13680
   LinkTopic       =   "Form23"
   ScaleHeight     =   8085
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6015
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   10610
      _Version        =   327680
      ForeColor       =   16384
      BackColorFixed  =   8454143
      BackColorSel    =   8388736
      FocusRect       =   0
   End
   Begin VB.Label hprod 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "hprod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label hsku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "hsku"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "holdhist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Label1.Height * 4)
End Sub

Private Sub hsku_Change()
    Dim cfile As String, s As String, t As String, tcolor As Boolean
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim f10 As String, f11 As String, f12 As String, f13 As String, f14 As String
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 16
    For i = 1 To holdlist.Grid5.Rows - 1
        If holdlist.Grid5.TextMatrix(i, 1) = hsku Then
            s = holdlist.Grid5.TextMatrix(i, 0) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 1) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 2) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 3) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 4) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 5) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 6) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 7) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 8) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 9) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 10) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 11) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 12) & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 13) & Chr(9)
            s = s & "Current" & Chr(9)
            s = s & holdlist.Grid5.TextMatrix(i, 1)
            s = s & holdlist.Grid5.TextMatrix(i, 3)
            s = s & holdlist.Grid5.TextMatrix(i, 4)
            s = s & holdlist.Grid5.TextMatrix(i, 0)
            's = s & holdlist.Grid5.TextMatrix(i, 13)
            Grid1.AddItem s
        End If
    Next i
    cfile = wdlogdir & "holdlogs\hold" & hsku & ".txt"
    If Len(Dir(cfile)) > 0 Then
        i = 0
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14
            s = f0 & Chr(9)
            s = s & f1 & Chr(9)
            s = s & f2 & Chr(9)
            s = s & f3 & Chr(9)
            s = s & f4 & Chr(9)
            s = s & f5 & Chr(9)
            s = s & f6 & Chr(9)
            s = s & f7 & Chr(9)
            s = s & f8 & Chr(9)
            s = s & f9 & Chr(9)
            s = s & f10 & Chr(9)
            s = s & f11 & Chr(9)
            s = s & f12 & Chr(9)
            s = s & f13 & Chr(9)
            s = s & f14 & Chr(9)
            s = s & f1 & f3 & f4 & f0 & i '& f13
            Grid1.AddItem s
            i = i + 1
        Loop
        Close #1
    End If
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 15: Grid1.ColSel = 15
    Grid1.Sort = 5
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        't = Grid1.TextMatrix(1, 0): tcolor = False
        t = Grid1.TextMatrix(1, 1) & Grid1.TextMatrix(1, 3) & Grid1.TextMatrix(1, 4): tcolor = False
        For i = 1 To Grid1.Rows - 1
            'If Grid1.TextMatrix(i, 0) <> t Then
            If Grid1.TextMatrix(i, 1) & Grid1.TextMatrix(i, 3) & Grid1.TextMatrix(i, 4) <> t Then
                tcolor = Not tcolor
                't = Grid1.TextMatrix(i, 0)
                t = Grid1.TextMatrix(i, 1) & Grid1.TextMatrix(i, 3) & Grid1.TextMatrix(i, 4)
            End If
            If tcolor = True Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = Grid1.BackColorFixed
            End If
        Next i
        Grid1.Row = 1
    End If
    s = "^ID|^SKU|<Product|^Lot|^OpCode|^Start|^End|^Source|^R12Lot||||<UserId|<DateTime|^TranType"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 1800
    Grid1.ColWidth(8) = 1200
    Grid1.ColWidth(9) = 0 '1000
    Grid1.ColWidth(10) = 0 '1000
    Grid1.ColWidth(11) = 0 '1000
    Grid1.ColWidth(12) = 1400
    Grid1.ColWidth(13) = 1400
    Grid1.ColWidth(14) = 1400
    Grid1.ColWidth(15) = 0
    
End Sub

