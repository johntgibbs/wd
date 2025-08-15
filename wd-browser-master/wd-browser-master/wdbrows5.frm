VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Transport Request Schedule"
   ClientHeight    =   4260
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8850
   LinkTopic       =   "Form5"
   ScaleHeight     =   4260
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4471
      _Version        =   327680
      ForeColor       =   255
      BackColorBkg    =   4210752
      FocusRect       =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label brcode 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcol As Boolean
Private Sub brcode_Change()
    Dim filler As String, br As String, dat As String
    Dim t1 As String, t2 As String, t3 As String
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Rows = 1: Grid1.Cols = 5: Grid1.FixedCols = 2
    If Len(Dir(Form1.webdir & "\orders\trsched." & brcode)) > 1 Then
        Open Form1.webdir & "\orders\trsched." & brcode For Input As #1
        Line Input #1, filler
        Label1 = filler
        Do Until EOF(1)
            Input #1, br, dat, t1, t2, t3
            filler = br & Chr$(9)
            filler = filler & dat & Chr$(9)
            filler = filler & t1 & Chr$(9)
            filler = filler & t2 & Chr$(9)
            filler = filler & t3
            Grid1.AddItem filler
        Loop
        Close #1
    End If
    Grid1.FormatString = "^Branch|^Date|^Brenham|^Bkn Arrow|^Sylacauga"
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 1400: Grid1.ColWidth(2) = 1400
    Grid1.ColWidth(3) = 1400: Grid1.ColWidth(4) = 1400
    Grid1.Redraw = True
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer, k As Integer, sqlx As String
    If Val(brcode) = 0 Then Exit Sub
    Open Form1.webdir & "\orders\trsched." & brcode For Output As #1
    Print #1, Label1
    For i = 1 To Grid1.Rows - 1
        sqlx = ""
        For k = 0 To Grid1.Cols - 1
            sqlx = sqlx & Grid1.TextMatrix(i, k) & ","
        Next k
        sqlx = Left(sqlx, Len(sqlx) - 1)
        Print #1, sqlx
    Next i
    Close #1
    If Form5.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form5" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form5.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form5.Left
                'Form1.frmgrid.TextMatrix(i, 3) = Form5.Height
                'Form1.frmgrid.TextMatrix(i, 4) = Form5.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form5" Then
            Form5.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form5.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            'Form5.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            'Form5.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    Label1.Width = Grid1.Width
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Label1.Height * 3.5)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    If Grid1.Col < 2 Then Exit Sub
    If Grid1.Row < 1 Then Exit Sub
    If edcol = True Then
        Grid1.Text = ""
        edcol = False
    End If
    If KeyAscii = 8 Then
        If Len(Grid1.Text) > 1 Then
            Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
        Else
            Grid1.Text = "0"
        End If
    End If
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Grid1.Text = Grid1.Text & Chr$(KeyAscii)
    End If
    Grid1.Text = Val(Grid1.Text)
End Sub

Private Sub Grid1_RowColChange()
    edcol = True
End Sub

Private Sub prtmenu_Click()
    Dim i As Integer, sqlx As String, k As Integer
    If Val(brcode) = 0 Then Exit Sub
    Open Form1.webdir & "\orders\trsched." & brcode For Output As #1
    Print #1, Label1
    For i = 1 To Grid1.Rows - 1
        sqlx = ""
        For k = 0 To Grid1.Cols - 1
            sqlx = sqlx & Grid1.TextMatrix(i, k) & ","
        Next k
        sqlx = Left(sqlx, Len(sqlx) - 1)
        Print #1, sqlx
    Next i
    Close #1
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print Label1
    For i = 0 To Grid1.Rows - 1
        sqlx = Grid1.TextMatrix(i, 0)
        sqlx = sqlx & Space(8 - Len(sqlx))
        sqlx = sqlx & Grid1.TextMatrix(i, 1)
        sqlx = sqlx & Space(18 - Len(sqlx))
        sqlx = sqlx & Grid1.TextMatrix(i, 2)
        sqlx = sqlx & Space(27 - Len(sqlx))
        sqlx = sqlx & Grid1.TextMatrix(i, 3)
        sqlx = sqlx & Space(38 - Len(sqlx))
        sqlx = sqlx & Grid1.TextMatrix(i, 4)
        Printer.Print sqlx
    Next i
    Printer.EndDoc
End Sub
