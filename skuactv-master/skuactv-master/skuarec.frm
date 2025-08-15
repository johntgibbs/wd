VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Last Receipts"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form6"
   ScaleHeight     =   3135
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4683
      _Version        =   327680
      Cols            =   7
      BackColorFixed  =   12648447
      FocusRect       =   0
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label reclit 
      Caption         =   ".."
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
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label rprod 
      Caption         =   "rprod"
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
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label rsku 
      Caption         =   "rsku"
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
      Width           =   375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function calc_date(lotcode As String) As String
    Dim seed As String
    'seed = "12-31-19" & Val(Left(lotcode, 2)) - 1
    If Val(Left(lotcode, 2)) > 90 Then
        seed = "19" & Left(lotcode, 2)
    Else
        seed = "20" & Left(lotcode, 2)
    End If
    seed = "12-31-" & Val(seed) - 1
    calc_date = Format(DateAdd("d", Val(Right(lotcode, 3)), seed), "m-d-yyyy")
End Function

Private Sub refresh_logs()
    Dim cfile As String, logpath As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim f10 As String, f11 As String, f12 As String, f13 As String, f14 As String
    Dim f16 As String
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    logpath = Form1.pallogs
    cfile = logpath & "move" & Format(Now, "mmddyyyy") & ".txt"
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input Shared As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
            If f6 >= rsku And f6 < rsku & "ZZZZZZZZZZ" Then
                s = StrConv(f3, vbProperCase) & Chr(9) & f4 & Chr(9) & f6 & Chr(9) & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12
                If UCase(f3) = "ROBOT ZERO" Then Grid1.AddItem s
                If UCase(f3) = "TRI LEVEL" Then Grid1.AddItem s
                If UCase(f3) = "ROLLER BED" Then Grid1.AddItem s
                If UCase(f3) = "WRAPPER" Then Grid1.AddItem s
                If UCase(f3) = "BACKHAUL" Then Grid1.AddItem s
                If UCase(f3) = "STAGING" Then Grid1.AddItem s
                If UCase(f3) = "SNACK PLANT" Then Grid1.AddItem s
                If UCase(f3) = "1731" Then Grid1.AddItem s
                If UCase(f3) = "1405" Then Grid1.AddItem s
                If UCase(f3) = "1406" Then Grid1.AddItem s
            End If
        Loop
        Close #1
    End If
    If Form1.plantno = "50" Or Form1.plantno = "52" Then
        cfile = logpath & "tml" & Format(Now, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 >= rsku And f6 < rsku & "ZZZZZZZZZZ" Then
                    s = StrConv(f3, vbProperCase) & Chr(9) & f4 & Chr(9) & f6 & Chr(9) & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12
                    'If UCase(f3) = "ROBOT ZERO" Then Grid1.AddItem s
                    'If UCase(f3) = "TRI LEVEL" Then Grid1.AddItem s
                    'If UCase(f3) = "ROLLER BED" Then Grid1.AddItem s
                    'If UCase(f3) = "WRAPPER" Then Grid1.AddItem s
                    'If UCase(f3) = "BACKHAUL" Then Grid1.AddItem s
                    'If UCase(f3) = "STAGING" Then Grid1.AddItem s
                    'If UCase(f3) = "SNACK PLANT" Then Grid1.AddItem s
                    'If UCase(f3) = "1731" Then Grid1.AddItem s
                    'If UCase(f3) = "1405" Then Grid1.AddItem s
                    'If UCase(f3) = "1406" Then Grid1.AddItem s
                    If UCase(f4) <> "SR4" Then Grid1.AddItem s
                End If
            Loop
            Close #1
        End If
    End If
    If Form1.plantno = "50" Then
        cfile = logpath & "recv" & Format(Now, "mmddyyyy") & ".txt"
        If Len(Dir(cfile)) > 0 Then
            Open cfile For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f6 >= rsku And f6 < rsku & "ZZZZZZZZZZ" Then
                    s = StrConv(f3, vbProperCase) & Chr(9) & f4 & Chr(9) & f6 & Chr(9) & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12
                    'If UCase(f3) = "ROBOT ZERO" Then Grid1.AddItem s
                    'If UCase(f3) = "TRI LEVEL" Then Grid1.AddItem s
                    If UCase(f3) = "ROLLER BED" Then Grid1.AddItem s
                    'If UCase(f3) = "WRAPPER" Then Grid1.AddItem s
                    'If UCase(f3) = "BACKHAUL" Then Grid1.AddItem s
                    'If UCase(f3) = "STAGING" Then Grid1.AddItem s
                    'If UCase(f3) = "SNACK PLANT" Then Grid1.AddItem s
                    'If UCase(f3) = "1731" Then Grid1.AddItem s
                    'If UCase(f3) = "1405" Then Grid1.AddItem s
                    'If UCase(f3) = "1406" Then Grid1.AddItem s
                    'If UCase(f4) <> "SR4" Then Grid1.AddItem s
                    'MsgBox s & " " & UCase(f3)
                End If
            Loop
            Close #1
        End If
    End If
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 2: Grid1.ColSel = 2
    Grid1.Sort = 5
    Grid1.FormatString = "^Source|^Target|^BarCode|^Lot|^Units|^Lot2|^Units"
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 1600
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    Grid1.Redraw = True
    reclit.Caption = Format(Now, "m-d-yyyy") & "  " & Grid1.Rows - 1 & " Records"
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Form6.Caption = Form6.Caption & " " & Form1.plantdesc
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form6" Then
            Form6.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form6.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form6.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form6.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Resize()
    If Form6.Height > 3540 Then Grid1.Height = Form6.Height - 885
    Grid1.Width = Me.Width - 100
End Sub

Private Sub Form_Terminate()
    Dim i As Integer
    If Form6.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form6" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form6.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form6.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form6.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form6.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Terminate
End Sub

Private Sub rsku_Change()
    Call refresh_logs
End Sub

