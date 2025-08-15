VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brwzplana 
   BackColor       =   &H0080FF80&
   Caption         =   "PLANA"
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8460
   LinkTopic       =   "Form13"
   ScaleHeight     =   7920
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox gemmies 
      Height          =   285
      Left            =   360
      TabIndex        =   11
      Text            =   "s:\wd\bin\gemmies.txt"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox oradsn 
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Text            =   "pbbcri"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox orapwd 
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Text            =   "gmd0207"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox orauser 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Text            =   "bbcgmd"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4215
      Left            =   0
      TabIndex        =   4
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7435
      _Version        =   327680
      BackColorFixed  =   8454016
      GridColor       =   0
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      GridLinesFixed  =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label schlit 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label wduser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ".."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label brcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ".."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">30 Day Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "30 Day Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2 Week Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<2 Week Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Menu prtgrid 
      Caption         =   "Print"
   End
   Begin VB.Menu orabat 
      Caption         =   "Oracle Batches"
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu edpals 
         Caption         =   "Pallet Adjustment"
      End
   End
End
Attribute VB_Name = "brwzplana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim psku As String, pdesc As String, ppoh As String
    Dim puoh As String, poo As String, psales As String
    Dim udiff As String, pdiff As String, plnt As String
    Dim plit As String, pcc As String
    Dim s As String, pro As String, rpcode As String
    Dim ts As String
    Grid1.Clear: Grid1.Rows = 1
    Grid1.Cols = 11
    If brcode = "503" Then
        rpcode = "505"
    Else
        rpcode = brcode
    End If
    If Len(Dir(Form1.webdir & "\stock\gsales." & rpcode)) > 0 Then
        If brcode = "500" Then s = "Brenham Plants " & brcode
        If brcode = "501" Then s = "Bkn Arrow Plant " & brcode
        If brcode = "502" Then s = "Sylacauga Plant " & brcode
        If brcode = "503" Then s = "Snack Plant 503"
        If brcode = "507" Then
            s = "Sylacauga Staged Products " & brcode
            orabat.Visible = False
        End If
        s = s & " Sales vs. Inventory"
        s = s & "  Last update: "
        s = s & Format(FileDateTime(Form1.webdir & "\stock\gsales." & rpcode), "m-d-yyyy h:mm am/pm")
        If rpcode <> brcode Then s = s & " All Routes"
        Me.Caption = s
        'MsgBox Form1.webdir & "\stock\gsales." & rpcode
        Open Form1.webdir & "\stock\gsales." & rpcode For Input As #1
        Do Until EOF(1)
            Input #1, psku, pdesc, ppoh, puoh, poo, psales, udiff, pdiff, plnt, pro, plit, pcc
            If plnt = "50" Then plnt = "TX"
            If plnt = "51" Then plnt = "OK"
            If plnt = "52" Then plnt = "AL"
            s = psku & Chr(9) & pdesc & Chr(9)
            s = s & Format(plit, "#") & Chr(9)
            s = s & Format(Val(puoh) - Val(plit), "#") & Chr(9)
            s = s & Format(puoh, "#") & Chr(9)
            s = s & Format(poo, "#") & Chr(9)
            s = s & Format(psales, "#") & Chr(9)
            s = s & Format(udiff, "#") & Chr(9)
            s = s & Format(pdiff, "#") & Chr(9)
            s = s & pcc & Chr(9)
            s = s & plit
            Grid1.AddItem s
        Loop
    End If
    Close #1
    Grid1.FormatString = "^SKU|<Description|^Plant Units|^Branch Units|^Total Units|^Branch Orders|^Sales Last 30|^Units Diff|^Pallet Diff"
    Grid1.ColWidth(0) = 400
    Grid1.ColWidth(1) = 2800
    Grid1.ColWidth(2) = 700
    Grid1.ColWidth(3) = 700
    Grid1.ColWidth(4) = 700
    Grid1.ColWidth(5) = 700
    Grid1.ColWidth(6) = 700
    Grid1.ColWidth(7) = 700
    Grid1.ColWidth(8) = 700
    Grid1.ColWidth(9) = 1
    Grid1.ColWidth(10) = 1
    Grid1.FillStyle = flexFillRepeat
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i: Grid1.Col = 0: Grid1.ColSel = 9
        If Grid1.TextMatrix(i, 9) = "W" Then Grid1.CellBackColor = wcolor.BackColor
        If Grid1.TextMatrix(i, 9) = "B" Then Grid1.CellBackColor = bcolor.BackColor
        If Grid1.TextMatrix(i, 9) = "G" Then Grid1.CellBackColor = gcolor.BackColor
        If Grid1.TextMatrix(i, 9) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
    Next i
    If Grid1.Rows > 1 Then
        Grid1.RowHeight(0) = Grid1.RowHeight(1) * 2
        Grid1.Row = 1: Grid1.Col = 4
    End If
End Sub

Private Sub edpals_Click()
    Dim s As String, i As Integer
    If Grid1.Cols < 16 Then Exit Sub
    s = Grid1.TextMatrix(Grid1.Row, 15)
    s = InputBox("Pallet Adjustment", "Pallet Adjustment...", s)
    If Len(s) = 0 Then Exit Sub
    Grid1.TextMatrix(Grid1.Row, 15) = Format(Val(s), "#")
    Grid1.TextMatrix(Grid1.Row, 16) = Val(s) + Val(Grid1.TextMatrix(Grid1.Row, 14))
    i = Grid1.Row
    Grid1.RowSel = i: Grid1.Col = 16: Grid1.ColSel = 16
    If Val(Grid1.TextMatrix(i, 16)) < 0 Then
        If Grid1.TextMatrix(i, 9) = "W" Then
            Grid1.CellBackColor = wcolor.BackColor
        Else
            Grid1.CellBackColor = ycolor.BackColor
        End If
    End If
    If Val(Grid1.TextMatrix(i, 16)) = 0 Then Grid1.CellBackColor = bcolor.BackColor
    If Val(Grid1.TextMatrix(i, 16)) > 0 Then Grid1.CellBackColor = gcolor.BackColor
End Sub

Private Sub brcode_Change()
    Dim tc As Integer
    tc = Check1.Value
    Check1.Value = 0
    refresh_grid
    Check1.Value = tc
End Sub

Private Sub Form_Load()
    Dim lpbuff As String * 25
    Dim ret As Long, UserId As String, filler As String
    Dim t1 As String, t2 As String, t3 As String, cfile As String
    'cfile = "s:\wd\bin\plana.ini"
    'If Len(Dir(cfile)) > 0 Then
    '    Open cfile For Input As #1
    '    Do Until EOF(1)
    '        Line Input #1, filler
    '        If LCase(Left(filler, 7)) = "orapwd=" Then Me.orapwd = Right(filler, Len(filler) - 7)
    '        If LCase(Left(filler, 7)) = "oradsn=" Then Me.oradsn = Right(filler, Len(filler) - 7)
    '        If LCase(Left(filler, 8)) = "orauser=" Then Me.orauser = Right(filler, Len(filler) - 8)
    '        If LCase(Left(filler, 8)) = "gemmies=" Then Me.gemmies = Right(filler, Len(filler) - 8)
    '        'If LCase(Left(filler, 7)) = "webdir=" Then Me.webdir = Right(filler, Len(filler) - 7)
    '    Loop
    '    Close #1
    'End If
    Me.orapwd = "gmd0207"
    Me.oradsn = "pbbcri"
    Me.orauser = "bbcgmd"
    Me.gemmies = Form1.webdir & "\counts\gemmies.txt"
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 930
End Sub

Private Sub Grid1_DblClick()
    'MsgBox Form1.wdbranch & " " & Me.brcode
    If Form1.wdbranch = "52" And Me.brcode = "502" Then Exit Sub
    'If Form1.wdbranch = "52" Then Exit Sub
    If Form1.wdbranch = "47" Then Exit Sub
    Screen.MousePointer = 11
    brwzbrana2.calledby = "Plana"
    brwzbrana2.wsku = Grid1.TextMatrix(Grid1.Row, 0)
    Screen.MousePointer = 0
    brwzbrana2.Show
    Check1 = 1
End Sub

Private Sub Grid1_RowColChange()
    Dim i As Integer
    i = Grid1.Row
    If Check1 = 1 And Left(brwzbrana2.Caption, Len(Grid1.TextMatrix(i, 1))) <> Grid1.TextMatrix(i, 1) Then
        brwzbrana2.calledby = "Plana"
        brwzbrana2.wsku = Grid1.TextMatrix(Grid1.Row, 0)
    End If
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Grid1.Cols > 11 Then PopupMenu edmenu
    End If
End Sub

Private Sub orabat_Click()
    brwzorabat.Show
End Sub

Private Sub Prtgrid_Click()
    Dim rt As String, rf As String, rh As String
    Dim pl As String, i As Integer, lc As Integer
    rt = Me.Caption
    rh = " "
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If Grid1.Cols > 11 Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
        Exit Sub
    End If
    Screen.MousePointer = 11
    lc = 4
    Printer.Font = "Courier New"
    Printer.FontSize = 8
    Printer.Print Me.Caption
    Printer.Print " "
    'Printer.Print "---------1---------2---------3---------4---------5---------6---------7---------8---------9---------A---------B"
    'Printer.Print "                                     Pallets     Units     Units    Sales      Unit     Pallet              Reorder"
    'Printer.Print " SKU                                  OnHand    OnHand    OnOrder  Last 30     Diff      Diff      Plant   Pallet Qty"
    Printer.Print "                                      Plant    Branch     Total     Units     Sales      Unit     Pallet"
    Printer.Print " SKU                                  Units     Units     Units    OnOrder   Last 30     Diff      Diff "
    
    For i = 1 To Grid1.Rows - 1
        If lc > 76 Then
            Printer.NewPage
            lc = 4
            Printer.Print Me.Caption
            Printer.Print " "
            'Printer.Print "                                       Plant    Branch     Total   Sales   Units       Unit     Pallet"
            'Printer.Print " SKU                                   Units     Units     Units   Last 30 OnOrder     Diff      Diff "
            Printer.Print "                                      Plant    Branch     Total     Units     Sales      Unit     Pallet"
            Printer.Print " SKU                                  Units     Units     Units    OnOrder   Last 30     Diff      Diff "
            
        End If
        Grid1.TextMatrix(i, 1) = Left(Grid1.TextMatrix(i, 1), 30)
        pl = Space(4 - Len(Grid1.TextMatrix(i, 0))) & Grid1.TextMatrix(i, 0) & " "
        pl = pl & Grid1.TextMatrix(i, 1) & Space(30 - Len(Grid1.TextMatrix(i, 1))) & " "
        pl = pl & Space(7 - Len(Grid1.TextMatrix(i, 2))) & Grid1.TextMatrix(i, 2) & " "
        pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 3))) & Grid1.TextMatrix(i, 3) & " "
        pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 4))) & Grid1.TextMatrix(i, 4) & " "
        pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 5))) & Grid1.TextMatrix(i, 5) & " "
        pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 6))) & Grid1.TextMatrix(i, 6) & " "
        pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 7))) & Grid1.TextMatrix(i, 7) & " "
        pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 8))) & Grid1.TextMatrix(i, 8) & " "
        'pl = pl & Space(9 - Len(Grid1.TextMatrix(i, 9))) & Grid1.TextMatrix(i, 9)
        Printer.Print pl
        lc = lc + 1
    Next i
    Printer.EndDoc
    Screen.MousePointer = 0
End Sub
