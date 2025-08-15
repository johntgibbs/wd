VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form wdbphone 
   BackColor       =   &H00C0FFFF&
   Caption         =   "BB Pallets"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13515
   LinkTopic       =   "Form12"
   ScaleHeight     =   8895
   ScaleWidth      =   13515
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6615
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
         Left            =   4680
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         Max             =   1
         Min             =   1
         TabIndex        =   7
         Top             =   120
         Value           =   1
         Width           =   855
      End
      Begin VB.Label pagelit 
         Caption         =   "Label1"
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3855
      Left            =   7320
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4560
      Width           =   7335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   7275
      TabIndex        =   1
      Top             =   480
      Width           =   7335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3201
      _Version        =   327680
      AllowUserResizing=   3
   End
   Begin VB.Label pcode 
      Caption         =   "pcode"
      Height          =   255
      Left            =   9600
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label rtype 
      Caption         =   "rtype"
      Height          =   255
      Left            =   7800
      TabIndex        =   10
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label qsort 
      Caption         =   "..."
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   6495
   End
   Begin VB.Label qstr 
      Caption         =   "qstr"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   6375
   End
End
Attribute VB_Name = "wdbphone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function parse_detail(gr As Integer) As sdetail
    Dim s As sdetail, j As Integer
    s.wonum = wdbtrkwo.Grid1.TextMatrix(gr, 0)
    s.date = Format(wdbtrkwo.Grid1.TextMatrix(gr, 1), "M-dd-yyyy")
    s.trip = Trim(wdbtrkwo.Grid1.TextMatrix(gr, 2))
    s.comments = Trim(wdbtrkwo.Grid1.TextMatrix(gr, 3))
    s.trlno = wdbtrkwo.Grid1.TextMatrix(gr, 4)
    s.driver = wdbtrkwo.Grid1.TextMatrix(gr, 5)
    s.trlsize = wdbtrkwo.Grid1.TextMatrix(gr, 6)
    s.startime = wdbtrkwo.Grid1.TextMatrix(gr, 7)
    s.hours = wdbtrkwo.Grid1.TextMatrix(gr, 8)
    s.worktype = wdbtrkwo.Grid1.TextMatrix(gr, 9)
    s.contents = wdbtrkwo.Grid1.TextMatrix(gr, 10)
    s.meals = wdbtrkwo.Grid1.TextMatrix(gr, 11)
    s.wostatus = wdbtrkwo.Grid1.TextMatrix(gr, 12)
    s.parentwo = wdbtrkwo.Grid1.TextMatrix(gr, 13)
    s.sortstart = wdbtrkwo.Grid1.TextMatrix(gr, 14)
    s.sorttrip = wdbtrkwo.Grid1.TextMatrix(gr, 15)
    s.sortdriver = wdbtrkwo.Grid1.TextMatrix(gr, 16)
    s.origin = wdbtrkwo.Grid1.TextMatrix(gr, 17)
    s.destination = wdbtrkwo.Grid1.TextMatrix(gr, 18)
    
    If Left(s.origin, 1) = "K" And Val(Right(s.origin, Len(s.origin) - 1)) > 0 Then s.origin = "K10"
    If Left(s.destination, 1) = "K" And Val(Right(s.destination, Len(s.destination) - 1)) > 0 Then s.destination = "K10"
    If s.origin = "047" Then s.origin = "K10"
    If s.destination = "047" Then s.destination = "K10"
    If Left(s.origin, 1) = "A" And Val(Right(s.origin, Len(s.origin) - 1)) > 0 Then s.origin = "A10"
    If Left(s.destination, 1) = "A" And Val(Right(s.destination, Len(s.destination) - 1)) > 0 Then s.destination = "A10"
    If s.origin = "052" Then s.origin = "A10"
    If s.destination = "052" Then s.destination = "A10"
    If Left(s.origin, 1) = "T" And Val(Right(s.origin, Len(s.origin) - 1)) > 0 Then s.origin = "T10"
    If Left(s.destination, 1) = "T" And Val(Right(s.destination, Len(s.destination) - 1)) > 0 Then s.destination = "T10"
    If s.origin = "001" Then s.origin = "T10"
    If s.destination = "001" Then s.destination = "T10"
    
    
    s.Plant = " "
    If s.origin = "A10" Then s.Plant = "SY>"
    If s.origin = "K10" Then s.Plant = "BA>"
    
    j = InStr(1, wdbtrkwo.Grid1.TextMatrix(gr, 2), ">")
    If j > 0 Then
        s.oname = Trim(Left(wdbtrkwo.Grid1.TextMatrix(gr, 2), j - 1))
        s.dname = Trim(Right(wdbtrkwo.Grid1.TextMatrix(gr, 2), Len(wdbtrkwo.Grid1.TextMatrix(gr, 2)) - j))
    End If
    s.endtime = Format(DateAdd("n", Val(s.hours) * 60, s.startime), "h:mm am/pm")
    If s.worktype = "Return" Then
        s.trip = s.oname & " Return to " & s.dname
        j = InStr(1, s.comments, "Return")
        If j > 0 Then
            If Len(s.comments) >= j + 5 Then
                s.comments = " "
            Else
                s.comments = Trim(Right(s.comments, Len(s.comments) - (j + 5)))
            End If
        End If
    Else
        If s.origin < "0" Then
            s.trip = s.comments
            s.comments = " "
        Else
            'MsgBox s.comments & "|" & s.dname
            If Left(s.comments, Len(s.dname)) = s.dname Then
                s.comments = Trim(Right(s.comments, Len(s.comments) - Len(s.dname)))
            End If
            If s.worktype = "Start" Or s.worktype = "SameDay" Or s.worktype = "Job" And UCase(Left(s.contents, 2)) = "IC" Then
                s.trip = s.dname & " #" & s.trlno
            End If
            If s.worktype = "Swap" Then
                s.trip = s.dname & " Swap - " & swap_driver(s.parentwo, s.driver)
            End If
        End If
    End If
    parse_detail = s
End Function

Private Function parse_brief(gr As Integer) As sbrief
    Dim s As sbrief, j As Integer
    s.date = Format(wdbtrkwo.Grid2.TextMatrix(gr, 0), "M-dd-yyyy")
    s.trip = wdbtrkwo.Grid2.TextMatrix(gr, 1)
    s.comments = wdbtrkwo.Grid2.TextMatrix(gr, 2)
    s.trlno = wdbtrkwo.Grid2.TextMatrix(gr, 3)
    s.driver = wdbtrkwo.Grid2.TextMatrix(gr, 4)
    s.trlsize = wdbtrkwo.Grid2.TextMatrix(gr, 5)
    s.startime = wdbtrkwo.Grid2.TextMatrix(gr, 6)
    s.hours = wdbtrkwo.Grid2.TextMatrix(gr, 7)
    s.worktype = wdbtrkwo.Grid2.TextMatrix(gr, 8)
    s.contents = wdbtrkwo.Grid2.TextMatrix(gr, 9)
    s.meals = wdbtrkwo.Grid2.TextMatrix(gr, 10)
    s.wostatus = wdbtrkwo.Grid2.TextMatrix(gr, 11)
    s.endtime = wdbtrkwo.Grid2.TextMatrix(gr, 12)
    s.recno = Val(wdbtrkwo.Grid2.TextMatrix(gr, 13))

    j = InStr(1, wdbtrkwo.Grid2.TextMatrix(gr, 1), ">")
    If j > 0 Then
        s.oname = Trim(Left(wdbtrkwo.Grid2.TextMatrix(gr, 1), j - 1))
        s.dname = Trim(Right(wdbtrkwo.Grid2.TextMatrix(gr, 1), Len(wdbtrkwo.Grid2.TextMatrix(gr, 1)) - j))
    End If
    s.Plant = " "
    If s.oname = "Sylacauga" Then s.Plant = "SY>"
    If s.oname = "Broken Arrow" Then s.Plant = "BA>"
    
    
    s.endtime = Format(DateAdd("n", Val(s.hours) * 60, s.startime), "h:mm am/pm")
    If s.worktype = "Return" Then
        s.trip = s.oname & " Return to " & s.dname
        j = InStr(1, s.comments, "Return")
        If j > 0 Then
            If Len(s.comments) >= j + 5 Then
                s.comments = " "
            Else
                s.comments = Right(s.comments, Len(s.comments) - (j + 5))
            End If
        End If
    Else
        If Left(s.comments, Len(s.dname)) = s.dname Then
            s.comments = Right(s.comments, Len(s.comments) - Len(s.dname))
        End If
    
        If s.worktype = "Start" Or s.worktype = "SameDay" Or s.worktype = "Job" And UCase(Left(s.contents, 2)) = "IC" Then
            s.trip = s.dname & " #" & s.trlno
        End If
    End If
    parse_brief = s
End Function

Private Function swap_driver(wno As String, dname As String) As String
    Dim i As Integer, s As String
    s = ""
    For i = 1 To wdbtrkwo.Grid1.Rows - 1
        If wdbtrkwo.Grid1.TextMatrix(i, 13) = wno And wdbtrkwo.Grid1.TextMatrix(i, 9) = "Swap" And wdbtrkwo.Grid1.TextMatrix(i, 5) <> dname Then
            s = wdbtrkwo.Grid1.TextMatrix(i, 5)
            Exit For
        End If
    Next i
    swap_driver = s
End Function

Private Sub phonebook_portrait(pd As Control)
    Dim pxs As Long, pxe As Long, pys As Long, pye As Long
    Dim gxs As Long, gxe As Long, gys As Long, gye As Long
    Dim maxc As Integer, curc As Integer, gwdth As Long
    Dim ftx As Long, fty As Long, rstr As String
    Dim cx As Long, p As Integer
    Dim cw(0 To 128) As Long, i As Integer
    'Set Picture to Portrait
    Picture1.Width = 8.5 * 1440
    Picture1.Height = 11 * 1440
    If Len(Dir(localAppDataPath & "\blnk8x11.bmp")) = 0 Then
        SavePicture Picture1.Image, localAppDataPath & "\blnk8x11.bmp"
    End If
    DoEvents
    Picture1.Picture = LoadPicture(localAppDataPath & "\blnk8x11.bmp")
    DoEvents
    
    
    pxs = 0: pxe = 7.6 * 1440
    pys = 0: pye = 10 * 1440
    gxs = 0: gxe = 0: gys = 0: gye = 0
    maxc = 1: curc = 1
    ftx = 1440: fty = 10.25 * 1440
    If TypeOf pd Is Printer Then
        pd.DrawWidth = 4
        'Printer.Orientation = 2
        Printer.Orientation = 1
        Printer.FontName = Grid1.FontName
        'Printer.FontSize = 8
        Printer.FontSize = Grid1.FontSize
    Else
        pd.FontName = Grid1.FontName
        pd.FontSize = Grid1.FontSize
        rstr = localAppDataPath & "\blnk8x11.bmp"
        pd.Picture = LoadPicture(rstr)
        rstr = Dir(localAppDataPath & "\cic*.bmp")
        Do While Len(rstr) > 0
            Kill localAppDataPath & "\" & rstr
            rstr = Dir
        Loop
        DoEvents
    End If
    For i = 0 To Grid1.Cols - 1
        cw(i) = Grid1.ColWidth(i)
        If Grid1.ColWidth(i) > 100 Then gwdth = gwdth + Grid1.ColWidth(i)
    Next i
    For i = 1 To 6
        If gwdth * i < 14400 Then
            pxe = gwdth * i
            maxc = i
        End If
    Next i
    pd.CurrentY = pys
    pd.CurrentX = pxs
    p = 1
    For i = 1 To Grid1.Rows - 1
        gxe = (curc * gwdth) + (curc * 400)
        gxs = gxe - gwdth
        If Grid1.TextMatrix(i, 0) = "ng" Then       'header
            If gys <> 0 Then
                jj = pd.CurrentY
                pd.Line (gxs, gys)-(gxs, pd.CurrentY)
                cx = gxs
                For k = 0 To Grid1.Cols - 1
                    cx = cx + cw(k)
                    pd.Line (cx, gys)-(cx, pd.CurrentY)
                Next k
            End If
            'pd.CurrentX = gxs: pd.Print ""
            pd.CurrentY = pd.CurrentY + 50
            If TypeOf pd Is Printer Then
                Printer.FontBold = True
            Else
                pd.FontBold = True
            End If
            pd.CurrentX = gxs: pd.Print Grid1.TextMatrix(i, 1) & " " & Grid1.TextMatrix(i, 2);
            pd.CurrentX = gxe - (pd.TextWidth(Grid1.TextMatrix(i, 3)) + 50): pd.Print Grid1.TextMatrix(i, 3)
            If TypeOf pd Is Printer Then
                Printer.FontBold = False
            Else
                pd.FontBold = False
            End If
            gys = pd.CurrentY
        Else
            pd.Line (gxs, pd.CurrentY)-(gxe, pd.CurrentY)
            cx = gxs + 100
            pd.CurrentX = cx
            pd.CurrentY = pd.CurrentY + 30
            For k = 0 To Grid1.Cols - 1
                If cw(k) > 100 Then
                    pd.Print Grid1.TextMatrix(i, k);
                    cx = cx + cw(k)
                    pd.CurrentX = cx
                End If
            Next k
            pd.Print ""
            pd.Line (gxs, pd.CurrentY)-(gxe, pd.CurrentY)
        End If
        If pd.CurrentY >= pye Then
            If Grid1.TextMatrix(i, 0) <> "ng" Then
                pd.Line (gxs, gys)-(gxs, pd.CurrentY)
                cx = gxs
                For k = 0 To Grid1.Cols - 1
                    cx = cx + cw(k)
                    pd.Line (cx, gys)-(cx, pd.CurrentY)
                Next k
            End If
            If curc = maxc Then
                'pd.Print ""
                pd.CurrentX = ftx: pd.CurrentY = fty
                pd.Print Format(Now, "mmmm d, yyyy  h:mm am/pm")
                If TypeOf pd Is Printer Then
                    pd.NewPage
                Else
                    rstr = localAppDataPath & "\cic" & Format(p, "00000") & ".bmp"
                    SavePicture pd.Image, rstr
                    p = p + 1
                    pd.Cls
                End If
                curc = 1
            Else
                curc = curc + 1
            End If
            gys = 1
            pd.CurrentY = pys
        End If
    Next i
    
    pd.Line (gxs, gys)-(gxs, pd.CurrentY)
    cx = gxs
    For k = 0 To Grid1.Cols - 1
        cx = cx + cw(k)
        pd.Line (cx, gys)-(cx, pd.CurrentY)
    Next k
    pd.CurrentX = ftx: pd.CurrentY = fty
    pd.Print Format(Now, "mmmm d, yyyy  h:mm am/pm")
    If TypeOf pd Is Printer Then pd.EndDoc
    If TypeOf pd Is PictureBox Then
        If p > 1 Then
            rstr = localAppDataPath & "\cic" & Format(p, "00000") & ".bmp"
            SavePicture pd.Image, rstr
            pd.Picture = LoadPicture(rstr)
            HScroll1.Visible = True
            HScroll1.Value = 1
        Else
            HScroll1.Visible = False
        End If
        pagelit.Caption = "Page 1 of " & p
        HScroll1.Max = p
    End If

End Sub

Private Sub phonebook_landscape(pd As Control)
    Dim pxs As Long, pxe As Long, pys As Long, pye As Long
    Dim gxs As Long, gxe As Long, gys As Long, gye As Long
    Dim maxc As Integer, curc As Integer, gwdth As Long
    Dim ftx As Long, fty As Long, rstr As String
    Dim cx As Long, p As Integer
    Dim cw(0 To 128) As Long, i As Integer
    pxs = 0: pxe = 10 * 1440
    pys = 0: pye = 7.6 * 1440
    gxs = 0: gxe = 0: gys = 0: gye = 0
    maxc = 1: curc = 1
    ftx = 1440: fty = 8 * 1440
    If TypeOf pd Is Printer Then
        pd.DrawWidth = 4
        Printer.Orientation = 2
        Printer.FontName = Grid1.FontName
        'Printer.FontSize = 8
        Printer.FontSize = Grid1.FontSize
    Else
        pd.FontName = Grid1.FontName
        pd.FontSize = Grid1.FontSize
        rstr = localAppDataPath & "\blnk11x8.bmp"
        pd.Picture = LoadPicture(rstr)
        rstr = Dir(localAppDataPath & "\cic*.bmp")
        Do While Len(rstr) > 0
            Kill localAppDataPath & "\" & rstr
            rstr = Dir
        Loop
        DoEvents
    End If
    For i = 0 To Grid1.Cols - 1
        cw(i) = Grid1.ColWidth(i)
        If Grid1.ColWidth(i) > 100 Then gwdth = gwdth + Grid1.ColWidth(i)
    Next i
    For i = 1 To 6
        If gwdth * i < 14400 Then
            pxe = gwdth * i
            maxc = i
        End If
    Next i
    pd.CurrentY = pys
    pd.CurrentX = pxs
    p = 1
    For i = 1 To Grid1.Rows - 1
        gxe = (curc * gwdth) + (curc * 400)
        gxs = gxe - gwdth
        If Grid1.TextMatrix(i, 0) = "ng" Then       'header
            If gys <> 0 Then
                jj = pd.CurrentY
                pd.Line (gxs, gys)-(gxs, pd.CurrentY)
                cx = gxs
                For k = 0 To Grid1.Cols - 1
                    cx = cx + cw(k)
                    pd.Line (cx, gys)-(cx, pd.CurrentY)
                Next k
            End If
            'pd.CurrentX = gxs: pd.Print ""
            pd.CurrentY = pd.CurrentY + 50
            If TypeOf pd Is Printer Then
                Printer.FontBold = True
            Else
                pd.FontBold = True
            End If
            pd.CurrentX = gxs: pd.Print Grid1.TextMatrix(i, 1) & " " & Grid1.TextMatrix(i, 2);
            pd.CurrentX = gxe - (pd.TextWidth(Grid1.TextMatrix(i, 3)) + 50): pd.Print Grid1.TextMatrix(i, 3)
            If TypeOf pd Is Printer Then
                Printer.FontBold = False
            Else
                pd.FontBold = False
            End If
            gys = pd.CurrentY
        Else
            pd.Line (gxs, pd.CurrentY)-(gxe, pd.CurrentY)
            cx = gxs + 100
            pd.CurrentX = cx
            pd.CurrentY = pd.CurrentY + 30
            For k = 0 To Grid1.Cols - 1
                If cw(k) > 100 Then
                    pd.Print Grid1.TextMatrix(i, k);
                    cx = cx + cw(k)
                    pd.CurrentX = cx
                End If
            Next k
            pd.Print ""
            pd.Line (gxs, pd.CurrentY)-(gxe, pd.CurrentY)
        End If
        If pd.CurrentY >= pye Then
            If Grid1.TextMatrix(i, 0) <> "ng" Then
                pd.Line (gxs, gys)-(gxs, pd.CurrentY)
                cx = gxs
                For k = 0 To Grid1.Cols - 1
                    cx = cx + cw(k)
                    pd.Line (cx, gys)-(cx, pd.CurrentY)
                Next k
            End If
            If curc = maxc Then
                'pd.Print ""
                pd.CurrentX = ftx: pd.CurrentY = fty
                pd.Print Format(Now, "mmmm d, yyyy  h:mm am/pm")
                If TypeOf pd Is Printer Then
                    pd.NewPage
                Else
                    rstr = localAppDataPath & "\cic" & Format(p, "00000") & ".bmp"
                    SavePicture pd.Image, rstr
                    p = p + 1
                    pd.Cls
                End If
                curc = 1
            Else
                curc = curc + 1
            End If
            gys = 1
            pd.CurrentY = pys
        End If
    Next i
    
    pd.Line (gxs, gys)-(gxs, pd.CurrentY)
    cx = gxs
    For k = 0 To Grid1.Cols - 1
        cx = cx + cw(k)
        pd.Line (cx, gys)-(cx, pd.CurrentY)
    Next k
    pd.CurrentX = ftx: pd.CurrentY = fty
    pd.Print Format(Now, "mmmm d, yyyy  h:mm am/pm")
    If TypeOf pd Is Printer Then pd.EndDoc
    If TypeOf pd Is PictureBox Then
        If p > 1 Then
            rstr = localAppDataPath & "\cic" & Format(p, "00000") & ".bmp"
            SavePicture pd.Image, rstr
            pd.Picture = LoadPicture(rstr)
            HScroll1.Visible = True
            HScroll1.Value = 1
        Else
            HScroll1.Visible = False
        End If
        pagelit.Caption = "Page 1 of " & p
        HScroll1.Max = p
    End If
End Sub

Private Sub refresh_planttrk_ba()
    Dim i As Integer, k As Integer, th As Currency, sd As String, s As String
    Dim j As Integer, d As sdetail
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    For i = 0 To wdbtrkwo.sdriver.ListCount - 1
        If Right(wdbtrkwo.sdriver.List(i), 2) = "BA" Then
            s = "ng" & Chr(9) & wdbtrkwo.sdriver.List(i)
            th = 0
            For k = 1 To wdbtrkwo.Grid1.Rows - 1
                d = parse_detail(k)
                If d.driver = wdbtrkwo.sdriver.List(i) Then
                    th = th + Val(d.hours)
                End If
            Next k
            If th > 0 Then
                s = s & Chr(9) & Chr(9) & Format(th, "#.0")
                Grid1.AddItem s
                For k = 1 To wdbtrkwo.Grid1.Rows - 1
                    d = parse_detail(k)
                    If d.driver = wdbtrkwo.sdriver.List(i) Then
                        s = Format(d.date, "ddd m-d-yy") & Chr(9)
                        s = s & Format(d.startime, "h:mm am/pm") & Chr(9)
                        s = s & d.trip & Chr(9)
                        s = s & d.endtime
                        Grid1.AddItem s
                        s = d.Plant & Chr(9)
                        s = s & d.hours & Chr(9)
                        s = s & d.comments
                        Grid1.AddItem s
                    End If
                Next k
            End If
        End If
    Next i
    sd = "..."
    For i = 1 To wdbtrkwo.Grid1.Rows - 1
        d = parse_detail(i)
        If d.origin = "K10" Or d.destination = "K10" Then
            If Right(d.driver, 2) <> "BA" Then
                If d.driver <> sd Then
                    s = "ng" & Chr(9) & d.driver
                    Grid1.AddItem s
                    sd = d.driver
                End If
                s = Format(d.date, "ddd m-d-yy") & Chr(9)
                s = s & Format(d.startime, "h:mm am/pm") & Chr(9)
                s = s & d.trip & Chr(9)
                s = s & d.endtime
                Grid1.AddItem s
                s = d.Plant & Chr(9)
                s = s & d.hours & Chr(9)
                s = s & d.comments
                Grid1.AddItem s
            End If
        End If
    Next i
    Grid1.ColWidth(0) = 1300
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3600
    Grid1.ColWidth(3) = 800
    Screen.MousePointer = 0
End Sub

Private Sub refresh_planttrk_tx()
    Dim i As Integer, k As Integer, th As Integer, sd As String
    Dim j As Integer, d As sdetail
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    For i = 0 To wdbtrkwo.sdriver.ListCount - 1
        If Right(wdbtrkwo.sdriver.List(i), 2) <> "SY" And Right(wdbtrkwo.sdriver.List(i), 2) <> "BA" Then
            s = "ng" & Chr(9) & wdbtrkwo.sdriver.List(i)
            th = 0
            For k = 1 To wdbtrkwo.Grid1.Rows - 1
                d = parse_detail(k)
                If d.driver = wdbtrkwo.sdriver.List(i) Then
                    th = th + Val(d.hours)
                End If
            Next k
            If th > 0 Then
                s = s & Chr(9) & Chr(9) & Format(th, "#.0")
                Grid1.AddItem s
                For k = 1 To wdbtrkwo.Grid1.Rows - 1
                    d = parse_detail(k)
                    If d.driver = wdbtrkwo.sdriver.List(i) Then
                        s = Format(d.date, "ddd m-d-yy") & Chr(9)
                        s = s & Format(d.startime, "h:mm am/pm") & Chr(9)
                        s = s & d.trip & Chr(9)
                        s = s & d.endtime
                        Grid1.AddItem s
                        s = d.Plant & Chr(9)
                        s = s & d.hours & Chr(9)
                        s = s & d.comments
                        Grid1.AddItem s
                    End If
                Next k
            End If
        End If
    Next i
    
    sd = "..."
    For i = 1 To wdbtrkwo.Grid1.Rows - 1
        d = parse_detail(i)
        If d.origin = "T10" Or d.destination = "T10" Then
            If Right(d.driver, 2) = "SY" Or Right(d.driver, 2) = "BA" Then
                If d.driver <> sd Then
                    s = "ng" & Chr(9) & d.driver
                    Grid1.AddItem s
                    sd = d.driver
                End If
                s = Format(d.date, "ddd m-d-yy") & Chr(9)
                s = s & Format(d.startime, "h:mm am/pm") & Chr(9)
                s = s & d.trip & Chr(9)
                s = s & d.endtime
                Grid1.AddItem s
                s = d.Plant & Chr(9)
                s = s & d.hours & Chr(9)
                s = s & d.comments
                Grid1.AddItem s
            End If
        End If
    Next i
    Grid1.ColWidth(0) = 1300
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3600
    Grid1.ColWidth(3) = 800
    Screen.MousePointer = 0
End Sub

Private Sub refresh_planttrk_syl()
    Dim i As Integer, k As Integer, th As Currency, sd As String, s As String
    Dim j As Integer, d As sdetail
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    For i = 0 To wdbtrkwo.sdriver.ListCount - 1
        If Right(wdbtrkwo.sdriver.List(i), 2) = "SY" Then
            s = "ng" & Chr(9) & wdbtrkwo.sdriver.List(i)
            th = 0
            For k = 1 To wdbtrkwo.Grid1.Rows - 1
                d = parse_detail(k)
                If d.driver = wdbtrkwo.sdriver.List(i) Then
                    th = th + Val(d.hours)
                End If
            Next k
            If th > 0 Then
                s = s & Chr(9) & Chr(9) & Format(th, "#.0")
                Grid1.AddItem s
                For k = 1 To wdbtrkwo.Grid1.Rows - 1
                    d = parse_detail(k)
                    If d.driver = wdbtrkwo.sdriver.List(i) Then
                        s = Format(d.date, "ddd m-d-yy") & Chr(9)
                        s = s & Format(d.startime, "h:mm am/pm") & Chr(9)
                        s = s & d.trip & Chr(9)
                        s = s & d.endtime
                        Grid1.AddItem s
                        s = d.Plant & Chr(9)
                        s = s & d.hours & Chr(9)
                        s = s & d.comments
                        Grid1.AddItem s
                    End If
                Next k
            End If
        End If
    Next i
    sd = "..."
    For i = 1 To wdbtrkwo.Grid1.Rows - 1
        d = parse_detail(i)
        If d.origin = "A10" Or d.destination = "A10" Then
            If Right(d.driver, 2) <> "SY" Then
                If d.driver <> sd Then
                    s = "ng" & Chr(9) & d.driver
                    Grid1.AddItem s
                    sd = d.driver
                End If
                s = Format(d.date, "ddd m-d-yy") & Chr(9)
                s = s & Format(d.startime, "h:mm am/pm") & Chr(9)
                s = s & d.trip & Chr(9)
                s = s & d.endtime
                Grid1.AddItem s
                s = d.Plant & Chr(9)
                s = s & d.hours & Chr(9)
                s = s & d.comments
                Grid1.AddItem s
            End If
        End If
    Next i
    Grid1.ColWidth(0) = 1300
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3600
    Grid1.ColWidth(3) = 800
    Screen.MousePointer = 0
End Sub

Private Sub refresh_inbound(orgid As String)
    Dim i As Integer, s As String, th As Integer, sd As String
    Dim k As Integer, j As Integer, morg As String, mdest As String, d As sdetail
    sd = "..."
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    For i = 1 To wdbtrkwo.Grid1.Rows - 1
        d = parse_detail(i)
        If d.date <> sd Then
            sd = d.date
            'Overnight Returns
            th = 0
            For k = 1 To wdbtrkwo.Grid1.Rows - 1
                d = parse_detail(k)
                If d.destination = orgid And d.worktype = "Return" And d.date = sd Then
                    th = th + 1
                End If
            Next k
            If th > 0 Then
                d = parse_detail(i)
                s = "ng" & Chr(9) & Format(d.date, "ddd MM-dd") & " BB Overnight Return Trailers"
                s = s & Chr(9) & Chr(9) & th
                Grid1.AddItem s
                For k = 1 To wdbtrkwo.Grid1.Rows - 1
                    d = parse_detail(k)
                    If d.destination = orgid And d.worktype = "Return" And d.date = sd Then
                        s = orgid & Chr(9)
                        s = s & d.oname & Chr(9)
                        s = s & d.trlsize & Chr(9)
                        s = s & Format(d.startime, "h:mm a/p") & Chr(9)
                        s = s & d.driver & Chr(9)
                        s = s & d.contents & Chr(9)
                        s = s & "ETA: " & d.endtime
                        Grid1.AddItem s
                    End If
                Next k
            End If
            
            'SameDay Returns
            th = 0
            For k = 1 To wdbtrkwo.Grid1.Rows - 1
                d = parse_detail(k)
                If d.origin = orgid And d.worktype = "SameDay" And d.date = sd And Val(d.meals) = 0 Then
                    th = th + 1
                End If
            Next k
            If th > 0 Then
                d = parse_detail(i)
                s = "ng" & Chr(9) & Format(d.date, "ddd MM-dd") & " BB SameDay Return Trailers"
                s = s & Chr(9) & Chr(9) & th
                Grid1.AddItem s
                For k = 1 To wdbtrkwo.Grid1.Rows - 1
                    d = parse_detail(k)
                    If d.origin = orgid And d.worktype = "SameDay" And d.date = sd And Val(d.meals) = 0 Then
                        s = orgid & Chr(9)
                        s = s & d.dname & " #" & d.trlno & Chr(9)
                        s = s & d.trlsize & Chr(9)
                        s = s & Format(d.startime, "h:mm a/p") & Chr(9)
                        s = s & d.driver & Chr(9)
                        If d.contents <> "IceCream" Then
                            s = s & d.contents & Chr(9)
                        Else
                            s = s & Chr(9)
                        End If
                        s = s & "ETA: " & d.endtime
                        Grid1.AddItem s
                    End If
                Next k
            End If
            
            'Backhauls
            th = 0
            For k = 1 To wdbtrkwo.Grid1.Rows - 1
                d = parse_detail(k)
                If d.destination = orgid And d.contents <> "Empty" And d.date = sd And Val(d.meals) > 0 Then
                    th = th + 1
                End If
            Next k
            If th > 0 Then
                d = parse_detail(i)
                s = "ng" & Chr(9) & Format(d.date, "ddd MM-dd") & " BB Backhaul Trailers"
                s = s & Chr(9) & Chr(9) & th
                Grid1.AddItem s
                For k = 1 To wdbtrkwo.Grid1.Rows - 1
                    d = parse_detail(k)
                    If d.destination = orgid And d.contents <> "Empty" And d.date = sd And Val(d.meals) > 0 Then
                        s = d.origin & Chr(9)
                        s = s & d.dname & " #" & d.trlno & Chr(9)
                        s = s & d.trlsize & Chr(9)
                        s = s & Format(d.startime, "h:mm a/p") & Chr(9)
                        s = s & d.driver & Chr(9)
                        s = s & d.contents & Chr(9)
                        s = s & "ETA: " & d.endtime
                        Grid1.AddItem s
                    End If
                Next k
            End If
        End If
      Next i
      Grid1.ColWidth(0) = 500
      Grid1.ColWidth(1) = 3000
      Grid1.ColWidth(2) = 500
      Grid1.ColWidth(3) = 900
      Grid1.ColWidth(4) = 2000
      Grid1.ColWidth(5) = 1500
      Grid1.ColWidth(6) = 1500
    Screen.MousePointer = 0
End Sub

Private Sub refresh_outbound(orgid)
    Dim i As Integer, s As String, th As Integer, sd As String
    Dim k As Integer, j As Integer, d As sdetail
    sd = "..."
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    For i = 1 To wdbtrkwo.Grid1.Rows - 1
        d = parse_detail(i)
        If d.date <> sd Then
            sd = d.date
            'Ice Cream Loads
            th = 0
            For k = 1 To wdbtrkwo.Grid1.Rows - 1
                d = parse_detail(k)
                If d.origin = orgid And (d.worktype = "SameDay" Or d.worktype = "Start") And d.date = sd And UCase(Left(d.contents, 2)) = "IC" Then
                    th = th + 1
                End If
            Next k
            If th > 0 Then
                d = parse_detail(i)
                s = "ng" & Chr(9) & Format(d.date, "ddd MM-dd") & " Blue Bell Trailers"
                s = s & Chr(9) & Chr(9) & th
                Grid1.AddItem s
                For k = 1 To wdbtrkwo.Grid1.Rows - 1
                    d = parse_detail(k)
                    If d.origin = orgid And (d.worktype = "SameDay" Or d.worktype = "Start") And d.date = sd And UCase(Left(d.contents, 2)) = "IC" Then
                        s = orgid & Chr(9)
                        s = s & d.dname & " #" & d.trlno & Chr(9)
                        s = s & d.trlsize & Chr(9)
                        s = s & Format(d.startime, "h:mm a/p") & Chr(9)
                        If d.contents <> "IceCream" Then
                            s = s & d.contents & Chr(9)
                        Else
                            s = s & Chr(9)
                        End If
                        s = s & d.comments & Chr(9)
                        s = s & d.driver
                        Grid1.AddItem s
                    End If
                Next k
            End If
            
            'Jobbing Loads
            th = 0
            For k = 1 To wdbtrkwo.Grid1.Rows - 1
                d = parse_detail(k)
                If d.origin = orgid And (d.worktype = "SameDay" Or d.worktype = "Start" Or d.worktype = "Job") And d.date = sd And (d.contents = "Jobbing" Or d.contents = "IC+Jobbing") Then
                    th = th + 1
                End If
            Next k
            If th > 0 Then
                d = parse_detail(i)
                s = "ng" & Chr(9) & Format(d.date, "ddd MM-dd") & " Jobbing Trailers"
                s = s & Chr(9) & Chr(9) & th
                Grid1.AddItem s
                For k = 1 To wdbtrkwo.Grid1.Rows - 1
                    d = parse_detail(k)
                    If d.origin = orgid And (d.worktype = "SameDay" Or d.worktype = "Start" Or d.worktype = "Job") And d.date = sd And (d.contents = "Jobbing" Or d.contents = "IC+Jobbing") Then
                        s = orgid & Chr(9)
                        s = s & d.dname & " #" & d.trlno & Chr(9)
                        s = s & d.trlsize & Chr(9)
                        s = s & Format(d.startime, "h:mm a/p") & Chr(9)
                        If d.contents <> "IceCream" Then
                            s = s & d.contents & Chr(9)
                        Else
                            s = s & Chr(9)
                        End If
                        s = s & d.comments & Chr(9)
                        s = s & d.driver
                        Grid1.AddItem s
                    End If
                Next k
            End If
        End If
      Next i
      Grid1.ColWidth(0) = 500
      Grid1.ColWidth(1) = 2000
      Grid1.ColWidth(2) = 500
      Grid1.ColWidth(3) = 900
      Grid1.ColWidth(4) = 1200
      Grid1.ColWidth(5) = 4200
      Grid1.ColWidth(6) = 1800

End Sub

Private Sub refresh_clist_driver()
    Dim i As Integer, s As String, th As Integer, sd As String
    Dim b As sbrief, d As sdetail
    sd = "..."
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    For i = 1 To wdbtrkwo.Grid2.Rows - 1
        b = parse_brief(i)
        If b.driver <> sd Then
            s = "ng" & Chr(9) & b.driver
            Grid1.AddItem s
            sd = b.driver
        End If
        d = parse_detail(b.recno)
        s = d.Plant & Chr(9)
        s = s & Format(d.date, "ddd m-d-yy") & Chr(9)
        s = s & Format(d.startime, "h:mm am/pm") & Chr(9)
        s = s & d.trip & Chr(9)
        s = s & d.comments & Chr(9)
        s = s & d.contents & Chr(9)
        s = s & d.hours
        Grid1.AddItem s
    Next i
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 1600
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 4000
    Grid1.ColWidth(4) = 5000
    Grid1.ColWidth(5) = 1800
    Grid1.ColWidth(6) = 800
End Sub

Private Sub refresh_clist_trip()
    Dim i As Integer, s As String, th As Integer, sd As String
    Dim b As sbrief, d As sdetail
    sd = "..."
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    For i = 1 To wdbtrkwo.Grid2.Rows - 1
        b = parse_brief(i)
        If b.dname <> sd Then
            s = "ng" & Chr(9) & b.dname
            Grid1.AddItem s
            sd = b.dname
        End If
        d = parse_detail(b.recno)
        s = d.Plant & Chr(9)
        If d.worktype = "Start" Then s = s & "#" & d.trlno
        If d.worktype = "SameDay" Then s = s & "#" & d.trlno
        If d.worktype = "Delivery" Then s = s & "#" & d.trlno
        If d.worktype = "2ndDay" Then s = s & "#" & d.trlno
        s = s & Chr(9)
        s = s & Format(d.date, "ddd m-d-yy") & Chr(9)
        s = s & Format(d.startime, "h:mm am/pm") & Chr(9)
        s = s & d.driver & Chr(9)
        If d.comments > " " Then
            s = s & d.comments & Chr(9)
        Else
            s = s & d.trip & Chr(9)
        End If
        s = s & d.contents & Chr(9)
        s = s & d.hours
        Grid1.AddItem s
    Next i
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 600
    Grid1.ColWidth(2) = 1600
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 3000
    Grid1.ColWidth(5) = 4000
    Grid1.ColWidth(6) = 1800
    Grid1.ColWidth(7) = 800
End Sub

Private Sub refresh_clist_date()
    Dim i As Integer, s As String, sd As String
    Dim b As sbrief, d As sdetail
    sd = "..."
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    For i = 1 To wdbtrkwo.Grid2.Rows - 1
        b = parse_brief(i)
        If sd <> b.date Then
            s = "ng" & Chr(9) & Format(b.date, "dddd M-d-yyyy")
            Grid1.AddItem s
            sd = b.date
        End If
        d = parse_detail(b.recno)
        s = d.Plant & Chr(9)
        s = s & Format(d.startime, "h:mm am/pm") & Chr(9)
        s = s & d.trip & Chr(9)
        s = s & d.driver & Chr(9)
        s = s & d.comments & Chr(9)
        s = s & d.contents & Chr(9)
        s = s & d.hours
        Grid1.AddItem s
    Next i
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 4000
    Grid1.ColWidth(3) = 2200
    Grid1.ColWidth(4) = 4000
    Grid1.ColWidth(5) = 1800
    Grid1.ColWidth(6) = 800
End Sub

Private Sub refresh_branchorder()
    Dim i As Integer, s As String, k As Integer
    Dim tp As Integer, tw As Integer, ta As Integer
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    Form13.Grid1.Col = 0
    Form13.Grid1.Col = 1
    For i = 1 To Form13.Grid1.Rows - 1
        Form13.Grid1.Row = i
        DoEvents
        s = "ng" & Chr(9) & Format(Form13.Grid1.TextMatrix(i, 1), "ddd m-dd") & " "
        s = s & Form13.Grid1.TextMatrix(i, 3) & Chr(9) & Chr(9)
        s = s & Form13.Grid1.TextMatrix(i, 4) & " Pallets"
        Grid1.AddItem s
        s = "SKU" & Chr(9) & "Product" & Chr(9) & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Alternate"
        Grid1.AddItem s
        tp = 0: tw = 0: ta = 0
        For k = 1 To Form13.Grid2.Rows - 1
            If Val(Form13.Grid2.TextMatrix(k, 9)) > 0 Then
                tp = tp + Val(Form13.Grid2.TextMatrix(k, 3))
                tw = tw + Val(Form13.Grid2.TextMatrix(k, 4))
                If Form13.Grid2.TextMatrix(k, 5) = "Y" Then ta = ta + 1
                s = Form13.Grid2.TextMatrix(k, 0) & Chr(9)
                s = s & Form13.Grid2.TextMatrix(k, 1) & Chr(9)
                s = s & Form13.Grid2.TextMatrix(k, 3) & Chr(9)
                s = s & Form13.Grid2.TextMatrix(k, 4) & Chr(9)
                s = s & Form13.Grid2.TextMatrix(k, 5)
                Grid1.AddItem s
            End If
        Next k
        s = Chr(9) & "  Totals" & Chr(9) & tp & Chr(9) & tw & Chr(9) & ta
        Grid1.AddItem s
    Next i
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 1200
End Sub

Private Sub Command1_Click()
    If rtype = "planttrk_syl" Then
        Call phonebook_landscape(Printer)
    End If
    If rtype = "planttrk_ba" Then
        Call phonebook_landscape(Printer)
    End If
    If rtype = "planttrk_tx" Then
        Call phonebook_landscape(Printer)
    End If
    If rtype = "inboundt10" Then
        Call phonebook_portrait(Printer)
    End If
    If rtype = "inboundk10" Then
        Call phonebook_portrait(Printer)
    End If
    If rtype = "inbounda10" Then
        Call phonebook_portrait(Printer)
    End If
    If rtype = "outboundt10" Then
        Call phonebook_portrait(Printer)
    End If
    If rtype = "outboundk10" Then
        Call phonebook_portrait(Printer)
    End If
    If rtype = "outbounda10" Then
        Call phonebook_portrait(Printer)
    End If
    If rtype = "clist_driver" Then
        Call phonebook_landscape(Printer)
    End If
    If rtype = "clist_trip" Then
        Call phonebook_landscape(Printer)
    End If
    If rtype = "clist_date" Then
        Call phonebook_landscape(Printer)
    End If
    If rtype = "branchorder" Then
        Call phonebook_portrait(Printer)
    End If
End Sub

Private Sub Form_Load()
    Picture1.Width = 11 * 1440
    Picture1.Height = 8.5 * 1440
    If Len(Dir(localAppDataPath & "\blnk11x8.bmp")) = 0 Then
        SavePicture Picture1.Image, localAppDataPath & "\blnk11x8.bmp"
    End If
    DoEvents

    Picture1.Picture = LoadPicture(localAppDataPath & "\blnk11x8.bmp")
    
    DoEvents
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    If Me.Height > 2000 Then
        VScroll1.Height = Me.Height - 940 '1050 '920
        VScroll1.Max = 15840 - VScroll1.Height
        VScroll1.SmallChange = Int(VScroll1.Max / 8)
        VScroll1.LargeChange = Int(VScroll1.Max / 3)
        HScroll2.Top = Me.Height - 720 '1050 '920 '880
    End If
    If Me.Width > 2000 Then
        VScroll1.Left = Me.Width - 380
        HScroll2.Width = Me.Width - 380
        HScroll2.Max = 12240 - HScroll2.Width
        If HScroll2.Max > 0 Then
            HScroll2.SmallChange = Int(HScroll2.Max / 8)
            HScroll2.LargeChange = Int(HScroll2.Max / 3)
        End If
        Frame1.Width = Me.Width
    End If
    If Me.Width > 12000 Then
        HScroll2.Visible = False
    Else
        HScroll2.Visible = True
    End If
    
End Sub

Private Sub HScroll1_Change()
    'Picture1.Cls
    rstr = localAppDataPath & "\cic" & Format(HScroll1.Value, "00000") & ".bmp"
    If Len(Dir(rstr)) > 0 Then
        Picture1.Picture = LoadPicture(rstr)
        pagelit.Caption = "Page " & HScroll1.Value & " of " & HScroll1.Max
    End If
End Sub

Private Sub HScroll2_Change()
    Picture1.Move 0 - HScroll2.Value
End Sub


Private Sub qstr_Change()
    Picture1.Cls
    If rtype = "planttrk_syl" Then
        Me.Caption = "Sylacauga Transports"
        Call refresh_planttrk_syl
        Call phonebook_landscape(Picture1)
    End If
    If rtype = "planttrk_ba" Then
        Me.Caption = "Broken Arrow Transports"
        Call refresh_planttrk_ba
        Call phonebook_landscape(Picture1)
    End If
    If rtype = "planttrk_tx" Then
        Me.Caption = "Brenham Transports"
        Call refresh_planttrk_tx
        Call phonebook_landscape(Picture1)
    End If
    If rtype = "inboundt10" Then
        Me.Caption = "Inbound Transports"
        Call refresh_inbound("T10")
        Call phonebook_portrait(Picture1)
    End If
    If rtype = "inboundk10" Then
        Me.Caption = "Inbound Transports"
        Call refresh_inbound("K10")
        Call phonebook_portrait(Picture1)
    End If
    If rtype = "inbounda10" Then
        Me.Caption = "Inbound Transports"
        Call refresh_inbound("A10")
        Call phonebook_portrait(Picture1)
    End If
    If rtype = "outboundt10" Then
        Me.Caption = "Out Bound Transports"
        Call refresh_outbound("T10")
        Call phonebook_portrait(Picture1)
    End If
    If rtype = "outboundk10" Then
        Me.Caption = "Out Bound Transports"
        Call refresh_outbound("K10")
        Call phonebook_portrait(Picture1)
    End If
    If rtype = "outbounda10" Then
        Me.Caption = "Out Bound Transports"
        Call refresh_outbound("A10")
        Call phonebook_portrait(Picture1)
    End If
    If rtype = "clist_driver" Then
        Me.Caption = "Driver Schedule"
        Call refresh_clist_driver
        Call phonebook_landscape(Picture1)
    End If
    If rtype = "clist_trip" Then
        Me.Caption = "Driver Schedule"
        Call refresh_clist_trip
        Call phonebook_landscape(Picture1)
    End If
    If rtype = "clist_date" Then
        Me.Caption = "Transport Schedule"
        Call refresh_clist_date
        Call phonebook_landscape(Picture1)
    End If
    If rtype = "branchorder" Then
        Me.Caption = "Branch Order"
        Call refresh_branchorder
        Call phonebook_portrait(Picture1)
    End If
End Sub

Private Sub VScroll1_Change()
    Picture1.Move Picture1.Left, Frame1.Height - VScroll1.Value
End Sub


