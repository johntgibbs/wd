VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brzbo 
   Caption         =   "Recent Back Orders"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12285
   LinkTopic       =   "Form14"
   ScaleHeight     =   9015
   ScaleWidth      =   12285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   9615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9615
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11033
      _Version        =   327680
      ForeColor       =   16711680
      BackColorFixed  =   12648447
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label bobrorder 
      Caption         =   ".............."
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      Top             =   8640
      Width           =   2175
   End
   Begin VB.Label bobranch 
      Caption         =   "..."
      Height          =   255
      Left            =   9600
      TabIndex        =   1
      Top             =   8640
      Width           =   1335
   End
End
Attribute VB_Name = "brzbo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()
    Dim cfile As String, s As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim f10 As String, f11 As String
    'cfile = "c:\jvwork\backorders." & bobranch.Caption
    cfile = Form1.webdir & "\stock\backorders." & bobranch.Caption
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11
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
            s = s & f11
            Grid1.AddItem s
        Loop
        Grid1.RemoveItem 1
        Close #1
    End If
    If bobranch.Caption = "01" Then
        s = "^Plant|<Branch|^Date|^SKU|<Product|^Pallets|^Wraps|^Units|^Shipped|^Short|^BO Pallets|^BO Wraps"
    Else
        s = "^Plant||^Date|^SKU|<Product|^Pallets|^Wraps|^Units|^Shipped|^Short|^BO Pallets|^BO Wraps"
    End If
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1200
    If bobranch.Caption = "01" Then
        Grid1.ColWidth(1) = 1800
    Else
        Grid1.ColWidth(1) = 0
    End If
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 3000
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 800
    Grid1.ColWidth(10) = 1000
    Grid1.ColWidth(11) = 1000
    Grid1.Col = 10
    Grid1_RowColChange
End Sub

Private Sub bobranch_Change()
    refresh_grid1
End Sub

Private Sub Command1_Click()
    Dim psku As String, pplant As String, i As Integer, pqty As String, pav As Boolean
    psku = Grid1.TextMatrix(Grid1.Row, 3)
    pplant = Grid1.TextMatrix(Grid1.Row, 0)
    pqty = Grid1.TextMatrix(Grid1.Row, 10)
    pav = False
    For i = 1 To Form3.Grid1.Rows - 1
        If Form3.Grid1.TextMatrix(i, 0) = psku Then
            Form3.Grid1.TextMatrix(i, 3) = pqty
            pav = True
            Exit For
        End If
    Next i
    If pav = False Then
        s = "SKU: " & psku & " is not available for " & pplant & " orders."
        MsgBox s, vbOKOnly + vbInformation, "sorry, "
    End If
End Sub

Private Sub Command2_Click()
    Dim psku As String, pplant As String, i As Integer, wqty As String, wav As Boolean
    psku = Grid1.TextMatrix(Grid1.Row, 3)
    pplant = Grid1.TextMatrix(Grid1.Row, 0)
    wqty = Grid1.TextMatrix(Grid1.Row, 10)
    wav = False
    For i = 1 To Form3.Grid1.Rows - 1
        If Form3.Grid1.TextMatrix(i, 0) = psku Then
            Form3.Grid1.TextMatrix(i, 4) = wqty
            wav = True
            Exit For
        End If
    Next i
    If wav = False Then
        s = "SKU: " & psku & " is not available for " & pplant & " orders."
        MsgBox s, vbOKOnly + vbInformation, "sorry, " & Form1.wduser
    End If
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    If Form1.Width > 2000 Then
        Me.Width = Form1.Width
        Me.Left = 10
    End If
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Me.Height > 3000 Then Grid1.Height = Me.Height - 1700
End Sub

Private Sub Grid1_RowColChange()
    Dim s As String
    Command1.Caption = "Re-order pallet(s) from " & Grid1.TextMatrix(Grid1.Row, 0)
    Command2.Caption = "Re-order wraps from " & Grid1.TextMatrix(Grid1.Row, 0)
    Command1.Enabled = False
    Command2.Enabled = False
    If Grid1.TextMatrix(Grid1.Row, 0) = "Brenham" And Left(bobrorder.Caption, 5) <> "Ord50" Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 0) = "Sylacauga" And Left(bobrorder.Caption, 5) <> "Ord52" Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 0) = "Broken Arrow" And Left(bobrorder.Caption, 5) <> "Ord51" Then Exit Sub
    If Val(Grid1.TextMatrix(Grid1.Row, 10)) > 0 Then
        If Val(Grid1.TextMatrix(Grid1.Row, 10)) = 1 Then
            s = "Re-order 1 pallet - " & Grid1.TextMatrix(Grid1.Row, 4)
        Else
            s = "Re-order " & Grid1.TextMatrix(Grid1.Row, 10) & " pallets - " & Grid1.TextMatrix(Grid1.Row, 4)
        End If
        s = s & " from " & Grid1.TextMatrix(Grid1.Row, 0) & "."
        Command1.Caption = s
        Command1.Enabled = True
    End If
    If Val(Grid1.TextMatrix(Grid1.Row, 11)) > 0 Then
        s = "Re-Order " & Grid1.TextMatrix(Grid1.Row, 11) & " wraps - " & Grid1.TextMatrix(Grid1.Row, 4)
        s = s & " from " & Grid1.TextMatrix(Grid1.Row, 0)
        Command2.Caption = s
        Command2.Enabled = True
    End If
End Sub
