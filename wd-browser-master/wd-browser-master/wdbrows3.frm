VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Branch Order"
   ClientHeight    =   9165
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9300
   LinkTopic       =   "Form3"
   ScaleHeight     =   9165
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   6960
      Width           =   8055
   End
   Begin VB.ListBox List1 
      ForeColor       =   &H00C00000&
      Height          =   1035
      Left            =   0
      TabIndex        =   22
      Top             =   240
      Width           =   8175
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   975
      Left            =   0
      TabIndex        =   21
      Top             =   8160
      Visible         =   0   'False
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   1720
      _Version        =   327680
      BackColorFixed  =   16776960
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      Height          =   1230
      Left            =   6840
      TabIndex        =   10
      Top             =   8520
      Visible         =   0   'False
      Width           =   4575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4335
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7646
      _Version        =   327680
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8454143
      BackColorSel    =   49344
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   2
      GridLinesFixed  =   1
      Appearance      =   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label5 
      Caption         =   "Notes to Brenham"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Messages From Brenham"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label rcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New Release"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label ts 
      Caption         =   "0"
      Height          =   255
      Left            =   6120
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label singlit 
      Caption         =   "..."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6360
      TabIndex        =   18
      Top             =   6360
      Width           =   3615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "On Hand Pallet Color Legend"
      Height          =   495
      Left            =   6480
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Over Month Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      TabIndex        =   16
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Full Month Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Below Month Supply"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Below 2 Week Level"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label odate 
      Caption         =   "..."
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label brcode 
      Caption         =   "0"
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label ordfile 
      Caption         =   "..."
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Plant 
      Caption         =   "Label4"
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label ta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label tw 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label tp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alternates"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wraps"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pallets"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   6360
      Width           =   975
   End
   Begin VB.Menu prtord 
      Caption         =   "Print"
      Enabled         =   0   'False
   End
   Begin VB.Menu postord 
      Caption         =   "Post Order"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu userec 
      Caption         =   "Use Recommended Order"
   End
   Begin VB.Menu restoreord 
      Caption         =   "Restore Order File"
   End
   Begin VB.Menu showbo 
      Caption         =   "Show Back Orders"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcol As Boolean, pflag As Boolean
Dim bname As String, singrow As Integer
Dim edcell As String

Private Sub refresh_notes_messages()
    Dim cfile As String, s As String
    Dim flen As Long, t1 As String
    cfile = Form1.webdir & "\stock\message." & Me.brcode
    List1.Clear
    Open cfile For Input As #1
    Do Until EOF(1)
        Line Input #1, s
        List1.AddItem s
    Loop
    Close #1
    Text1 = ""
    If Len(Dir(Form1.webdir & "\orders\notes." & Me.brcode)) > 0 Then
        flen = FileLen(Form1.webdir & "\orders\notes." & Me.brcode)
        Open Form1.webdir & "\orders\notes." & Me.brcode For Input As #1
        t1 = Input(flen, #1)
        Close #1
        Text1 = Trim(t1)
    End If
End Sub
Private Sub refresh_grid2()
    Dim tfile As String, i As Integer
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String, f5 As String
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 6
    Grid2.FixedCols = 1
    tfile = Form1.webdir & "\recycles\recycle." & Format(Val(brcode), "000")
    'MsgBox tfile
    If Len(Dir(tfile)) > 0 Then
        Open tfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5
            f0 = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9)
            f0 = f0 & f3 & Chr(9) & f4 & Chr(9) & f5
            Grid2.AddItem f0
        Loop
        Close #1
    Else
        tfile = Form1.webdir & "\recycles\recycle.csv"
        If Len(Dir(tfile)) > 0 Then
            Open tfile For Input As #1
            Do Until EOF(1)
                Input #1, f1, f3
                f0 = Format(Val(brcode), "000")
                f2 = " "
                f4 = "Brenham"
                f5 = Format(DateAdd("d", -1, Now), "m-d-yyyy")
                f0 = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9)
                f0 = f0 & f3 & Chr(9) & f4 & Chr(9) & f5
                Grid2.AddItem f0
            Loop
            Close #1
        End If
    End If
    Grid2.FormatString = "^Branch|^Recycle Item|^Qty|^UOM|^Send To|^Posted"
    Grid2.ColWidth(0) = 700
    Grid2.ColWidth(1) = 2000
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 1200
    Grid2.ColWidth(5) = 1200
    If Grid2.Rows > 1 Then
        Grid2.FillStyle = flexFillRepeat
        For i = 1 To Grid2.Rows - 1
            If Grid2.TextMatrix(i, 5) > "  " Then
                If DateDiff("d", Now, Grid2.TextMatrix(i, 5)) <> 0 Then
                    Grid2.Row = i: Grid2.RowSel = i
                    Grid2.Col = 5: Grid2.ColSel = 5
                    Grid2.CellBackColor = ycolor.BackColor
                End If
            Else
                Grid2.Row = i: Grid2.RowSel = i
                Grid2.Col = 5: Grid2.ColSel = 5
                Grid2.CellBackColor = ycolor.BackColor
            End If
        Next i
        Grid2.Row = 1: Grid2.Col = 3
    End If
End Sub
Private Sub save_fuel()
    Dim tfile As String
    If Grid2.Rows < 2 Then Exit Sub
    tfile = Form1.webdir & "\recycles\recycle." & Format(Val(brcode), "000")
    Open tfile For Output As #7
    For i = 1 To Grid2.Rows - 1
        For k = 0 To Grid2.Cols - 2
            Write #7, Grid2.TextMatrix(i, k);
        Next k
        Write #7, Grid2.TextMatrix(i, Grid2.Cols - 1)
    Next i
    Close #7
End Sub

Private Sub refresh_grid()
    Dim rtype As String, bno As String
    Dim sdate As String, mplant As Integer, mpals As Integer
    Dim gs As String, msku As String, mdesc As String, msrc As String
    Dim scnt As Integer, oc As String, af As String, wc As String
    Dim i As Integer, psku As String, mro As String
    Dim unam As String, sqlx As String
    rcolor.Visible = False
    unam = Left(Form1.wduser, InStr(1, Form1.wduser, " "))
    unam = "Hey " & unam & "....."
    If pflag = True Then
        sqlx = "Changes made to order: " & ordfile & " have not been posted."
        sqlx = sqlx & "  Do you wish to post the changes now?"
        If MsgBox(sqlx, vbQuestion + vbYesNo, unam) = vbYes Then
            Call prtord_Click
        End If
    End If
    pflag = False
    odate = Left(Combo1, 10)
    If Right(Combo1, 7) = "Brenham" Then Plant = "50"
    If Right(Combo1, 5) = "Arrow" Then Plant = "51"
    If Right(Combo1, 9) = "Sylacauga" Then Plant = "52"
    If Mid(Combo1, 12, 7) = "BobTail" Then Plant = Right(Combo1, 2)
    If Val(Plant) = 50 Then Grid1.BackColorFixed = &HFFFF80
    If Val(Plant) = 51 Then Grid1.BackColorFixed = &H80FF80
    If Val(Plant) = 52 Then Grid1.BackColorFixed = &H80FFFF
    mdesc = "Ord" & Plant & brcode & "." & DateDiff("d", "01-01-1999", Left(Combo1, 10))
    If ordfile.Caption = mdesc Then
        'List1.Visible = False: Grid1.Visible = True
        Call tp_Change
        Grid1.SetFocus
        Exit Sub
    Else
        ordfile = mdesc
    End If
    Screen.MousePointer = 11
    tp = "0": tw = "0": ta = "0"
    List2.Clear: List2.AddItem "000 0000": scnt = 0
    Grid1.Visible = False: Grid1.Clear
    Grid1.Rows = 1: Grid1.Cols = 12: Grid1.FixedCols = 3
    If Mid(Combo1, 12, 7) = "BobTail" Then
        Open Form1.webdir & "\stock\avdate.bob" For Input As #1
    Else
        'Open "C:\avord." & brcode For Input As #1
        'Open Form1.webdir & "\stock\avordtest.52" For Input As #1
        Open Form1.webdir & "\stock\avord." & brcode For Input As #1
        
    End If
    Input #1, rtype
    Do Until EOF(1) Or rtype = "End"
        If rtype = "B" Then Input #1, bno, bname
        If rtype = "S" Then Input #1, sdate, mplant, mpals
        If rtype = "P" Then
            Input #1, msrc, msku, mdesc, moh, moo, mud, mpd, mord, mro
            If msrc = Plant Or Plant > "B" Then
                If mord > 0 Then
                    mc = "1W"
                Else
                    If mpd < 0 Then mc = "2Y"
                    If mpd = 0 Then mc = "3B"
                    If mpd > 0 Then mc = "4G"
                End If
                scnt = scnt + 1
                sqlx = msku & Chr(9) & mdesc & Chr(9)
                sqlx = sqlx & moh & Chr(9)
                sqlx = sqlx & Chr(9) & Chr(9) & Chr(9)
                sqlx = sqlx & Format(mud, "#") & Chr(9)
                sqlx = sqlx & Format(mpd, "#") & Chr(9)
                sqlx = sqlx & Format(mord, "#") & Chr(9)
                sqlx = sqlx & mro & Chr(9)
                sqlx = sqlx & mc & Format(Val(msku), "00000") & Chr(9)
                sqlx = sqlx & Format(moo, "#")
                Grid1.AddItem sqlx
                List2.AddItem msku & Format(scnt, "0000")
            End If
        End If
        If rtype = "R" Then
            Input #1, msrc, msku, mdesc, moh, moo, mud, mpd, mord, mro
            If msrc = Plant Or Plant > "B" Then
                mc = "5R"
                rcolor.Visible = True
                scnt = scnt + 1
                sqlx = msku & Chr(9) & mdesc & Chr(9)
                sqlx = sqlx & moh & Chr(9)
                sqlx = sqlx & Chr(9) & Chr(9) & Chr(9)
                sqlx = sqlx & Format(mud, "#") & Chr(9)
                sqlx = sqlx & Format(mpd, "#") & Chr(9)
                sqlx = sqlx & Format(mord, "#") & Chr(9)
                sqlx = sqlx & mro & Chr(9)
                sqlx = sqlx & mc & Format(Val(msku), "00000") & Chr(9)
                sqlx = sqlx & Format(moo, "#")
                Grid1.AddItem sqlx
                List2.AddItem msku & Format(scnt, "0000")
            End If
        End If
        
        Input #1, rtype
    Loop
    Close #1
    Grid1.FormatString = "^SKU|<Product|^Stock|^Pallets|^Wraps|^Alt?|^UDiff|^PDiff|^Need|^ROQty|>Code|>OnOrd"
    Grid1.ColWidth(0) = 400
    Grid1.ColWidth(1) = 2900
    Grid1.ColWidth(2) = 700
    Grid1.ColWidth(3) = 600
    Grid1.ColWidth(4) = 600
    Grid1.ColWidth(5) = 400
    Grid1.ColWidth(6) = 600
    Grid1.ColWidth(7) = 500
    Grid1.ColWidth(8) = 500
    Grid1.ColWidth(9) = 600
    Grid1.ColWidth(10) = 0
    Grid1.ColWidth(11) = 0
    If Len(Dir(Form1.webdir & "\orders\" & ordfile)) > 0 Then
        Open Form1.webdir & "\orders\" & ordfile For Input As #1
        Do Until EOF(1)
            Input #1, bno, msku, oc, mpals, oc, af, wc
            If bno = brcode Then
                For i = 1 To scnt
                    'psku = Left(List2.List(i), 3)
                    psku = Trim(Left(List2.List(i), 4))                     'jv082415
                    If psku = msku Then
                        If mpals > 0 Then
                            tp = Val(tp) + mpals
                            Grid1.TextMatrix(i, 3) = Format(mpals, "#")
                        End If
                        If Val(wc) > 0 Then
                            tw = Val(tw) + Val(wc)
                            Grid1.TextMatrix(i, 4) = Format(Val(wc), "#")
                        End If
                        If af = "Y" Then
                            ta = Val(ta) + 1
                            Grid1.TextMatrix(i, 5) = "Y"
                        End If
                    End If
                Next i
            End If
        Loop
        Close #1
    End If
    'If Len(Dir(Form1.webdir & "\stock\goh." & brcode)) > 0 Then
    '    Open Form1.webdir & "\stock\goh." & brcode For Input As #1
    '    Do Until EOF(1)
    '        Line Input #1, mdesc
    '        If Len(mdesc) > 10 Then
    '            msku = Left(mdesc, 3)
    '            For i = 0 To Grid1.Rows - 1
    '                If msku = Grid1.TextMatrix(i, 0) Then
    '                    Grid1.TextMatrix(i, 2) = Format(Val(Right(mdesc, 10)), "#")
    '                    Exit For
    '                End If
    '            Next i
    '        End If
    '    Loop
    '    Close #1
    'End If
    ts.Caption = "0"
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 10: Grid1.ColSel = 10
    'Grid1.Sort = 5
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 3) = "1" And Grid1.TextMatrix(i, 9) = "2" Then ts = Val(ts.Caption) + 1
        Grid1.Row = i: Grid1.RowSel = i: Grid1.Col = 0: Grid1.ColSel = 2 '10
        If Left(Grid1.TextMatrix(i, 10), 2) = "1W" Then Grid1.CellBackColor = wcolor.BackColor
        If Left(Grid1.TextMatrix(i, 10), 2) = "3B" Then Grid1.CellBackColor = bcolor.BackColor
        If Left(Grid1.TextMatrix(i, 10), 2) = "4G" Then Grid1.CellBackColor = gcolor.BackColor
        If Left(Grid1.TextMatrix(i, 10), 2) = "2Y" Then Grid1.CellBackColor = ycolor.BackColor
        If Left(Grid1.TextMatrix(i, 10), 2) = "5R" Then Grid1.CellBackColor = rcolor.BackColor
        Grid1.FillStyle = flexFillRepeat
        Grid1.Col = 3: Grid1.ColSel = 5
        Grid1.CellFontBold = True
        Grid1.FillStyle = flexFillRepeat
        Grid1.Col = 6: Grid1.ColSel = 11
        If Left(Grid1.TextMatrix(i, 10), 2) = "1W" Then Grid1.CellBackColor = wcolor.BackColor
        If Left(Grid1.TextMatrix(i, 10), 2) = "3B" Then Grid1.CellBackColor = bcolor.BackColor
        If Left(Grid1.TextMatrix(i, 10), 2) = "4G" Then Grid1.CellBackColor = gcolor.BackColor
        If Left(Grid1.TextMatrix(i, 10), 2) = "2Y" Then Grid1.CellBackColor = ycolor.BackColor
        If Left(Grid1.TextMatrix(i, 10), 2) = "5R" Then Grid1.CellBackColor = rcolor.BackColor
        Grid1.FillStyle = flexFillRepeat
    Next i
    Grid1.Row = 1: Grid1.RowSel = 1: Grid1.Col = 0: Grid1.ColSel = 2 '10
    If Left(Grid1.TextMatrix(1, 10), 2) = "1W" Then Grid1.CellBackColor = wcolor.BackColor
    If Left(Grid1.TextMatrix(1, 10), 2) = "3B" Then Grid1.CellBackColor = bcolor.BackColor
    If Left(Grid1.TextMatrix(1, 10), 2) = "4G" Then Grid1.CellBackColor = gcolor.BackColor
    If Left(Grid1.TextMatrix(1, 10), 2) = "2Y" Then Grid1.CellBackColor = ycolor.BackColor
    If Left(Grid1.TextMatrix(1, 10), 2) = "5R" Then Grid1.CellBackColor = rcolor.BackColor
    Grid1.FillStyle = flexFillRepeat
    Grid1.Row = 1: Grid1.Col = 3
    Screen.MousePointer = 0
    Grid1.Visible = True
End Sub

Private Sub brcode_Change()
    Dim filler As String, t1 As String, t2 As String, t3 As String
    Dim rtype As String, bno As Integer, bname As String
    Dim sdate As String, mplant As Integer, mpals As Integer
    Dim mpl As String, i As Integer, k As Integer, pflag As Boolean
    Dim flen As Long
    Combo1.Clear
    Combo1.AddItem "Messages"
    Combo1.AddItem "Oracle Inventory"
    Combo1.AddItem "Trailer Status Report"
    Combo1.AddItem "Notes to Brenham"
    If Len(Dir(Form1.webdir & "\stock\avord." & brcode)) > 1 Then
    'If Len(Dir("c:\avord." & brcode)) > 1 Then
        Open Form1.webdir & "\stock\avord." & brcode For Input As #1
        'Open Form1.webdir & "\stock\avordtest.52" For Input As #1
        Input #1, rtype
        Do Until EOF(1) Or rtype = "P" Or rtype = "R"
            If rtype = "B" Then
                Input #1, bno, bname
                If bno = Val(brcode) Then
                    brlabel = "Trailer Orders " & bname & "  " & Format(Val(brcode), "00")
                End If
            End If
            If rtype = "S" Then
                Input #1, sdate, t2, t3
                If sdate = "No orders" Then
                    Combo1.AddItem sdate
                    userec.Enabled = False
                Else
                    userec.Enabled = True
                    If t2 = "50" Then Combo1.AddItem Format(sdate, "mm-dd-yyyy") & " " & t3 & " pallets from Brenham"
                    If t2 = "51" Then Combo1.AddItem Format(sdate, "mm-dd-yyyy") & " " & t3 & " pallets from Bkn Arrow"
                    If t2 = "52" Then Combo1.AddItem Format(sdate, "mm-dd-yyyy") & " " & t3 & " pallets from Sylacauga"
                    If UCase(Left(t2, 1)) = "B" Then Combo1.AddItem Format(sdate, "mm-dd-yyyy") & " BobTail " & UCase(t2)
                End If
            End If
            Input #1, rtype
        Loop
        Close #1
    End If
    Combo1.ListIndex = Combo1.ListCount - 1
    refresh_notes_messages
    'Recycle Items
    'refresh_grid2
End Sub

Private Sub Combo1_Click()
    If Left(Combo1, 9) = "No orders" Then Exit Sub
    If Combo1 = "Messages" Then
        Form2.wdfile = Form1.webdir & "\stock\message." & Format(Val(Form3.brcode), "00")
        Form2.Caption = "Messages From Brenham..."
        Form2.Show
    Else
        If Combo1 = "Notes to Brenham" Then
            Form4.brcode = Format(Val(Form3.brcode), "00")
            Form4.Show
        Else
            If Combo1 = "Oracle Inventory" Then
                Form2.wdfile = Form1.webdir & "\stock\goh." & Format(Val(Form3.brcode), "00")
                Form2.Caption = "Oracle Branch Inventory"
                Form2.Show
            Else
                If Combo1 = "Trailer Status Report" Then
                    'Form2.wdfile = Form1.webdir & "\stock\trlstat." & Format(Val(Form3.brcode), "00")
                    'Form2.Caption = "Trailer Status Report"
                    'Form2.Show
                    Form1.WebBrowser1.Navigate Form1.webdir & "\stock\trlstat" & Format(Val(brcode), "00") & ".htm"
                Else
                    Call refresh_grid
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Form3.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form3" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form3.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form3.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form3.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form3.Width
                Exit For
            End If
        Next i
    End If
End Sub
Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form3" Then
            Form3.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form3.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form3.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form3.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            'If Form3.Height < 7560 Then Form3.Height = 7560
            Exit For
        End If
    Next i
    pflag = False
    restoreord.Visible = False
End Sub

Private Sub Form_Resize()
    Grid1.Width = Form3.Width - 80
    'If Form3.Height > 2000 Then
    '    Grid1.Height = Form3.Height - 1000
    'End If
    List1.Width = Me.Width - 80
    Text1.Width = Me.Width - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim unam As String, sqlx As String
    unam = Left(Form1.wduser, InStr(1, Form1.wduser, " "))
    unam = "Hey " & unam & "....."
    If pflag = True Then
        sqlx = "Changes made to the order: " & ordfile & " have not been posted."
        sqlx = sqlx & "  Do you wish to post the changes now?"
        If MsgBox(sqlx, vbQuestion + vbYesNo, unam) = vbYes Then
            Call prtord_Click
        End If
    End If
    If Len(edcell) > 0 Then save_fuel
    Call Form_Deactivate
End Sub

Private Sub Grid1_GotFocus()
    Grid1.BackColorSel = &H800000
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim i As Integer, msg As String, omt As Integer
    If Grid1.Col = 3 Or Grid1.Col = 4 Then
        pflag = True
        If edcol = True Then
            Grid1.Text = ""
            edcol = False
        End If
        If KeyAscii = 8 Then
            If Len(Grid1.Text) > 1 Then
                Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
            Else
                Grid1.Text = ""
            End If
        End If
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            'MsgBox Grid1.TextMatrix(Grid1.Row, 10)
            omt = Val(Grid1.Text & Chr(KeyAscii))
            If Left(Grid1.TextMatrix(Grid1.Row, 10), 2) = "4G" Then
                msg = Grid1.TextMatrix(Grid1.Row, 1) & "?  "
                msg = msg & "Sale history indicates that you currently have"
                msg = msg & " a surplus amount of "
                If Grid1.Col = 3 Then
                    If omt > Val(Grid1.TextMatrix(Grid1.Row, 8)) Then
                        msg = msg & Grid1.TextMatrix(Grid1.Row, 7) & " pallets already."
                        msg = msg & "  Do you still wish to order this product?"
                        If MsgBox(msg, vbYesNo + vbQuestion, "Are you sure....") = vbNo Then Exit Sub
                    End If
                Else
                    msg = msg & Grid1.TextMatrix(Grid1.Row, 6) & " units on hand."
                    msg = msg & "  Do you still wish to order this product?"
                    If MsgBox(msg, vbYesNo + vbQuestion, "Are you sure....") = vbNo Then Exit Sub
                End If
            End If
            If Left(Grid1.TextMatrix(Grid1.Row, 10), 2) = "3B" Then
                msg = Grid1.TextMatrix(Grid1.Row, 1) & "?  "
                If Grid1.Col = 3 Then
                    If omt > Val(Grid1.TextMatrix(Grid1.Row, 8)) Then
                        msg = msg & "Sale history indicates that you currently have"
                        msg = msg & " a full month supply of this product."
                        msg = msg & "  Do you still wish to order this product?"
                        If MsgBox(msg, vbYesNo + vbQuestion, "Are you sure....") = vbNo Then Exit Sub
                    End If
                Else
                    If omt + Val(Grid1.TextMatrix(Grid1.Row, 6)) > 0 Then
                        msg = msg & "An order of " & omt
                        msg = msg & " wraps will create a monthly surplus amount"
                        msg = msg & " for this product.  Do you still wish to order these wraps?"
                        If MsgBox(msg, vbYesNo + vbQuestion, "Are you sure....") = vbNo Then Exit Sub
                    End If
                End If
            End If
            If Left(Grid1.TextMatrix(Grid1.Row, 10), 2) = "2Y" Then
                If Grid1.Col = 3 Then
                    If omt > Val(Grid1.TextMatrix(Grid1.Row, 8)) Then
                        If omt + Val(Grid1.TextMatrix(Grid1.Row, 7)) > 0 Then
                            msg = Grid1.TextMatrix(Grid1.Row, 1) & "?  "
                            msg = msg & "An order of " & omt
                            msg = msg & " pallets will create a monthly surplus amount of "
                            msg = msg & omt + Val(Grid1.TextMatrix(Grid1.Row, 7))
                            msg = msg & " pallets.  Do you still wish to order this amount?"
                            If MsgBox(msg, vbYesNo + vbQuestion, "Are you sure....") = vbNo Then Exit Sub
                        End If
                    End If
                Else
                    If omt + Val(Grid1.TextMatrix(Grid1.Row, 6)) > 0 Then
                        msg = msg & "An order of " & omt
                        msg = msg & " wraps will create a montly surplus amount"
                        msg = msg & " for this product.  Do you still wish to order these wraps?"
                        If MsgBox(msg, vbYesNo + vbQuestion, "Are you sure....") = vbNo Then Exit Sub
                    End If
                End If
            End If
            If Left(Grid1.TextMatrix(Grid1.Row, 10), 2) = "1W" Then
                If Grid1.Col = 3 Then
                    If omt > Val(Grid1.TextMatrix(Grid1.Row, 8)) Then
                        If omt + Val(Grid1.TextMatrix(Grid1.Row, 7)) > 0 Then
                            msg = Grid1.TextMatrix(Grid1.Row, 1) & "?  "
                            msg = msg & "An order of " & omt
                            msg = msg & " pallets will create a monthly surplus amount of "
                            msg = msg & omt + Val(Grid1.TextMatrix(Grid1.Row, 7))
                            msg = msg & " pallets.  Do you still wish to order this amount?"
                            If MsgBox(msg, vbYesNo + vbQuestion, "Are you sure....") = vbNo Then Exit Sub
                        End If
                    End If
                Else
                    If omt + Val(Grid1.TextMatrix(Grid1.Row, 6)) > 0 Then
                        msg = msg & "An order of " & omt
                        msg = msg & " wraps will create a monthly surplus amount"
                        msg = msg & " for this product.  Do you still wish to order these wraps?"
                        If MsgBox(msg, vbYesNo + vbQuestion, "Are you sure....") = vbNo Then Exit Sub
                    End If
                End If
            End If
            If Left(Grid1.TextMatrix(Grid1.Row, 10), 2) = "5R" Then
                If Grid1.Col = 3 Then
                    If omt > Val(Grid1.TextMatrix(Grid1.Row, 8)) Then
                        msg = "Order is limited to " & Grid1.TextMatrix(Grid1.Row, 8) & " pallets."
                        MsgBox msg, vbOKOnly + vbInformation, "New Product Release..."
                        Exit Sub
                    End If
                Else
                    Exit Sub
                End If
            End If
            Grid1.Text = Grid1.Text & Chr(KeyAscii)
        End If
    End If
    'singrow = 0
    If Grid1.Col = 3 And Grid1.TextMatrix(Grid1.Row, 9) = "2" Then
        If Val(Grid1.Text) = 1 Then
            msg = "This product will be shipped from the crane warehouses.  "
            msg = msg & "Please order two pallets or more if possible."
            MsgBox msg, vbOKOnly, "Crane item......"
            'singrow = Grid1.Row
        End If
    End If
    
    If Grid1.Col = 5 And Left(Grid1.TextMatrix(Grid1.Row, 10), 2) <> "5R" Then
        pflag = True
        If Grid1.Text >= "Y" Then
            Grid1.Text = ""
        Else
            Grid1.Text = "Y"
        End If
    End If
    tp = 0: tw = 0: ta = 0: ts = 0
    For i = 0 To Grid1.Rows - 1
        tp = tp + Val(Grid1.TextMatrix(i, 3))
        tw = tw + Val(Grid1.TextMatrix(i, 4))
        If Grid1.TextMatrix(i, 5) = "Y" Then ta = ta + 1
        If Grid1.TextMatrix(i, 3) = "1" And Grid1.TextMatrix(i, 9) = "2" Then ts = ts + 1
    Next i
    
End Sub

Private Sub Grid1_LostFocus()
    Grid1.BackColorSel = &HC0C0&
End Sub

Private Sub Grid1_RowColChange()
    edcol = True
    'If singrow > 0 Then
    '    msg = "This single pallet will be shipped from the crane warehouses.  "
    '    If Grid1.TextMatrix(singrow, 10) = "W" Or Grid1.TextMatrix(singrow, 10) = "Y" Then
    '        msg = msg & "Sales history indicates that a 2 pallet order will not"
    '        msg = msg & " create a surplus amount.  "
    '        msg = msg & "Do you wish to order 2 pallets?"
    '        If MsgBox(msg, vbYesNo + vbQuestion, "Crane item.. " & Grid1.TextMatrix(singrow, 1)) = vbYes Then
    '            Grid1.TextMatrix(singrow, 3) = "2"
    '            tp = Val(tp) + 1
    '            ts = Val(ts) - 1
    '        Else
    '            If Val(Grid1.TextMatrix(singrow, 11)) > 0 Then
    '                Grid1.TextMatrix(singrow, 3) = ""
    '                tp = Val(tp) - 1
    '                ts = Val(ts) - 1
    '                msg = "Sorry, there is already " & Grid1.TextMatrix(singrow, 11)
    '                msg = msg & " units on order for this product: "
    '                msg = msg & Grid1.TextMatrix(singrow, 1) & "."
    '                MsgBox msg, vbOKOnly + vbExclamation, "Single pallet order denied.."
    '            End If
    '        End If
    '    Else
    '        Grid1.TextMatrix(singrow, 3) = ""
    '        tp = Val(tp) - 1
    '        ts = Val(ts) - 1
    '        If Grid1.TextMatrix(singrow, 10) = "G" Then
    '            msg = msg & "Sales history indicates a surplus amount of "
    '            msg = msg & Grid1.TextMatrix(singrow, 7)
    '            msg = msg & " pallets already on hand."
    '        End If
    '        If Grid1.TextMatrix(singrow, 10) = "B" Then
    '            msg = msg & "Sales history indicates that a full 30 day supply"
    '            msg = msg & " is already on hand."
    '        End If
    '        msg = msg & "  Wait until sales increase enough to warrant this request."
    '        MsgBox msg, vbOKOnly + vbExclamation, "Single pallet order denied.."
    '    End If
    'End If
    'singrow = 0
End Sub
Private Sub Grid2_KeyPress(KeyAscii As Integer)
    If Grid2.Col <> 2 Then Exit Sub
    If edcell = "" Then Grid2.Text = ""
    If KeyAscii = 8 Then
        If Len(Grid2.Text) > 1 Then
            Grid2.Text = Left(Grid2.Text, Len(Grid2.Text) - 1)
        Else
            Grid2.Text = ""
        End If
        edcell = Grid2.TextMatrix(0, Grid2.Col)
    Else
        'If KeyAscii >= 31 And KeyAscii <= 127 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            Grid2.Text = Grid2.Text & Chr(KeyAscii)
            edcell = Grid2.TextMatrix(0, Grid2.Col)
            Grid2.TextMatrix(Grid2.Row, 5) = Format(Now, "m-d-yyyy")
        End If
    End If
End Sub

Private Sub grid2_LeaveCell()
    If Len(edcell) > 0 Then
        save_fuel
        edcell = ""
    End If
End Sub


Private Sub ordfile_Change()
    Dim cfile As String
    restoreord.Visible = False
    If Form1.wdbranch <> "SU" Then Exit Sub
    If ordfile < "Ord" Then Exit Sub
    cfile = Form1.webdir & "\orders\x" & ordfile
    If Len(Dir(cfile)) > 0 Then
        restoreord.Visible = True
        'MsgBox cfile
    End If
End Sub

Private Sub postord_Click()
    Dim i As Integer
    Open Form1.webdir & "\orders\" & ordfile For Output As #1
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 3)) > 0 Or Val(Grid1.TextMatrix(i, 4)) > 0 Or Grid1.TextMatrix(i, 5) = "Y" Then
            Write #1, brcode, Grid1.TextMatrix(i, 0), odate, Val(Grid1.TextMatrix(i, 3)), Plant, Grid1.TextMatrix(i, 5), Val(Grid1.TextMatrix(i, 4))
        End If
    Next i
    Close #1
    pflag = False
End Sub

Private Sub prtord_Click()
    Dim i As Integer, sqlx As String, ofile As String
    'Recycle Grid
    'If Grid2.Rows > 1 Then
    '    For i = 1 To Grid2.Rows - 1
    '        If DateDiff("d", Now, Grid2.TextMatrix(i, 5)) <> 0 Then
    '            sqlx = "Please update the quantities for recycle items."
    '            MsgBox sqlx, vbOKOnly + vbInformation, "recycle items..."
    '            Exit Sub
    '        End If
    '    Next i
    'End If
        
    If Len(Dir(Form1.webdir & "\orderoff.txt")) > 0 Then
        Form1.WebBrowser1.Navigate Form1.webdir & "\orderoff.htm"
        pflag = False
        Unload Form3
        Exit Sub
    End If
    If Len(Dir(Form1.webdir & "\orders\aord" & Right(ordfile, 4))) > 0 Then
        i = 0
    Else
        sqlx = Left(Form1.wduser, InStr(1, Form1.wduser, " "))
        sqlx = "Sorry " & sqlx & "!  "
        sqlx = sqlx & "The home office is no longer accepting orders for this date: "
        sqlx = sqlx & odate & "."
        MsgBox sqlx, vbOKOnly + vbExclamation, "Date expired..."
        pflag = False
        Unload Form3
        Exit Sub
    End If
    Call postord_Click
    DoEvents    '4-27-2000
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.Print ordfile & "   " & Combo1 & "    Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Printer.Print " "
    Printer.Print "SKU                                            Pallets   Wraps  Alternate"
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 3)) > 0 Or Val(Grid1.TextMatrix(i, 4)) > 0 Or Grid1.TextMatrix(i, 5) = "Y" Then
            sqlx = Grid1.TextMatrix(i, 0) & "  " & Grid1.TextMatrix(i, 1)
            sqlx = sqlx & Space(55 - Len(sqlx))
            sqlx = sqlx & Space(8 - Len(Grid1.TextMatrix(i, 3))) & Grid1.TextMatrix(i, 3)
            sqlx = sqlx & Space(8 - Len(Grid1.TextMatrix(i, 4))) & Grid1.TextMatrix(i, 4)
            sqlx = sqlx & Space(8 - Len(Grid1.TextMatrix(i, 5))) & Grid1.TextMatrix(i, 5)
            Printer.Print sqlx
        End If
    Next i
    Printer.Print " "
    Printer.Print "Total Pallets:  " & tp
    Printer.Print "        Wraps:  " & tw
    Printer.Print "   Alternates:  " & ta
    Printer.EndDoc
    ofile = Form1.webdir & "\orders\old" & brcode & ".txt"
    If Len(Dir(ofile)) > 0 Then
        Open ofile For Input As #1
        Do Until EOF(1)
            Line Input #1, sqlx
            Printer.Print sqlx
        Loop
        Close #1
        Kill ofile
    End If
End Sub

Private Sub restoreord_Click()
    Dim ofile As String, xfile As String
    ofile = Form1.webdir & "\orders\" & ordfile.Caption
    xfile = Form1.webdir & "\orders\x" & ordfile
    Name xfile As ofile
    DoEvents
    Call ordfile_Change
    DoEvents
    ordfile.Caption = "..."
    refresh_grid
End Sub

Private Sub showbo_Click()
    brzbo.bobrorder = Me.ordfile
    brzbo.bobranch = Me.brcode
    brzbo.Show
End Sub

Private Sub tp_Change()
    Dim stot As Integer
    stot = Val(tp) + Val(tw)
    'If stot > 0 And Val(ts) < 5 Then
    If stot > 0 Then
        prtord.Enabled = True
        postord.Enabled = True
    Else
        prtord.Enabled = False
        postord.Enabled = False
    End If
End Sub

Private Sub ts_Change()
    If Val(ts) > 2 Then
        singlit.Caption = ts & " Crane Singles" & String(Val(ts) - 1, ".")
    Else
        singlit.Caption = " "
    End If

End Sub

Private Sub tw_Change()
    Call tp_Change
End Sub

Private Sub userec_Click()
    Dim i As Integer, ptot As Integer
    If Grid1.Cols < 10 And Grid1.Rows > 1 Then Exit Sub
    ptot = 0
    For i = 1 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 3) = Grid1.TextMatrix(i, 8)
        ptot = ptot + Val(Grid1.TextMatrix(i, 3))
    Next i
    tp = ptot: pflag = True
End Sub
