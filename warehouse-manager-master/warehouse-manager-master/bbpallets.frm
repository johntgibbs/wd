VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form12 
   Caption         =   "BB Pallets"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   LinkTopic       =   "Form12"
   ScaleHeight     =   8895
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         Height          =   255
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
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3201
      _Version        =   327680
      AllowUserResizing=   3
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
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub phonebook(pd As Control)
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
        Printer.FontSize = 8
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
                ' (400, 240) - (400, 465)
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
            ' (400, 240) - (3700, 240)
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
            ' (400, 465) - (3700, 465)
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
            pd.Picture = LoadPicture(localAppDataPath & "\cic00001.bmp")
            HScroll1.Visible = True
            HScroll1.Value = 1
        Else
            HScroll1.Visible = False
        End If
        pagelit.Caption = "Page 1 of " & p
        HScroll1.Max = p
    End If
End Sub

Private Sub refresh_query()
    Dim ds As ADODB.Recordset, s As String
    Dim rs As ADODB.Recordset
    Dim pdesc As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    s = "select p.sku,count(*) from racks r, rackpos p "
    s = s & qstr & " and p.rackno = r.id"
    s = s & " group by p.sku"
    s = s & " order by p.sku"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            pdesc = " "
            i = Val(ds(0))
            If skurec(i).sku = ds(0) Then pdesc = skurec(i).prodname
            s = "ng" & Chr(9) & ds(0) & Chr(9) & pdesc & Chr(9) & ds(1)
            Grid1.AddItem s
            s = "select r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
            s = s & " from racks r, rackpos p "
            s = s & qstr & " and p.sku = '" & ds(0) & "' and p.rackno = r.id"
            s = s & " group by r.aisle,r.rack,p.sku,p.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
            If Len(qsort) > 5 Then
                s = s & " " & qsort
            Else
                s = s & " order by r.aisle,r.slot"
            End If
            Set rs = Wdb.Execute(s)
            If rs.BOF = False Then
                rs.MoveFirst
                Do Until rs.EOF
                    s = rs(3) & Chr(9)
                    s = s & rs(0) & "-" & rs(1) & Chr(9)
                    s = s & rs(10) & Chr(9)
                    If rs(8) = "N" Then
                        s = s & rs(10) & " 4 Way"
                    Else
                        If rs(5) > "0" Then
                            s = s & "<" & rs(5)
                            's = s & "<" & Form1.bb_codedate(rs!resv_lot)
                        Else
                            If rs(7) <> 0 Then
                                s = s & "On Hold"
                            Else
                                If rs(6) <> 0 Then
                                    s = s & "1st Out"
                                End If
                            End If
                        End If
                    End If
                    Grid1.AddItem s
                    rs.MoveNext
                Loop
            End If
            rs.Close
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.ColWidth(0) = 700
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 600
    Grid1.ColWidth(3) = 800
End Sub

Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, s As String
    Dim rs As ADODB.Recordset, cs5flag As Boolean
    Dim pdesc As String
    cs5flag = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    s = "select p.sku,count(*) from racks r, rackpos p "
    s = s & " where p.sku > '0' and p.rackno = r.id"
    s = s & " and r.aisle <> 'M' and p.bbc = 'Y'"
    If LCase(Right(Form1.Caption, 9)) = "sylacauga" Then
        If MsgBox("Include CS5 Pallets?", vbYesNo + vbQuestion, "sylacauga...") = vbNo Then
            s = s & " and r.room <> '5'"
            cs5flag = False
        End If
    End If
    s = s & " group by p.sku" ',p.bbc"
    s = s & " order by p.sku"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            pdesc = " "
            i = Val(ds(0))
            If skurec(i).sku = ds(0) Then pdesc = skurec(i).prodname
            s = "ng" & Chr(9) & ds!sku & Chr(9) & pdesc & Chr(9) & ds(1)
            Grid1.AddItem s
            If ds!sku > "0" Then
                s = "select r.aisle,r.rack,p.sku,r.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
                s = s & " from racks r, rackpos p "
                s = s & "where (p.sku = '" & ds(0) & "' or r.resv_sku = '" & ds(0) & "') and p.rackno = r.id"
                s = s & " and r.aisle <> 'M'"
                s = s & " group by r.aisle,r.rack,p.sku,r.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
            Else
                s = "select r.aisle,r.rack,p.sku,r.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot,count(*)"
                s = s & " from racks r, rackpos p "
                s = s & "where p.sku = '" & ds(0) & "' and r.resv_lot > '0' and p.rackno = r.id"
                s = s & " and r.aisle <> 'M'"
                s = s & " group by r.aisle,r.rack,p.sku,r.lot_num,r.resv_sku,r.resv_lot,r.fo,r.hold,p.bbc,r.slot"
            End If
            s = s & " order by r.lot_num,r.aisle,r.slot"
            Set rs = Wdb.Execute(s)
            If rs.BOF = False Then
                rs.MoveFirst
                Do Until rs.EOF
                    s = rs(3) & Chr(9)
                    s = s & rs(0) & "-" & rs(1) & Chr(9)
                    If rs(2) = ds(0) Then               'check for empty spaces
                        s = s & rs(10) & Chr(9)
                    Else
                        s = s & " " & Chr(9)
                    End If
                    If rs(8) = "N" Then
                        's = s & rs(10) & " 4 Way"
                        s = s & "4 Way"
                    Else
                        If rs(6) <> 0 Then
                            s = s & "1st Out"
                        Else
                            If rs(7) <> 0 Then
                                s = s & "On Hold"
                            Else
                                If rs(5) > "0" Then
                                    s = s & "<" & rs(5)
                                End If
                            End If
                        End If
                    End If
                    Grid1.AddItem s
                    rs.MoveNext
                Loop
            End If
            rs.Close
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.ColWidth(0) = 700
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 600
    Grid1.ColWidth(3) = 800
End Sub

Private Sub refresh_cranes()
    Dim ds As ADODB.Recordset, s As String
    Dim rs As ADODB.Recordset, cs5flag As Boolean
    Dim pdesc As String
    cs5flag = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    s = "select sku,sum(qty) from lane where qty > 0"
    If qstr = "Crane1" Then s = s & " and whse_num = 1"
    If qstr = "Crane2" Then s = s & " and whse_num = 2"
    If qstr = "Crane3" Then s = s & " and whse_num = 3"
    If qstr = "Crane5" Then s = s & " and whse_num = 5"
    s = s & " group by sku order by sku"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            pdesc = " "
            i = Val(ds(0))
            If skurec(i).sku = ds(0) Then pdesc = skurec(i).prodname
            s = "ng" & Chr(9) & ds!sku & Chr(9) & pdesc & Chr(9) & ds(1)
            Grid1.AddItem s
            If qstr = "CraneAll" Then
                s = "select whse_num,gmasize,lane_status,sum(qty) from lane where sku = '" & ds!sku & "'"
                s = s & " group by whse_num,gmasize,lane_status"
            Else
                s = "select whse_num,lot_num,gmasize,lane_status,sum(qty) from lane where sku = '" & ds!sku & "'"
                If qstr = "Crane1" Then s = s & " and whse_num = 1"
                If qstr = "Crane2" Then s = s & " and whse_num = 2"
                If qstr = "Crane3" Then s = s & " and whse_num = 3"
                If qstr = "Crane5" Then s = s & " and whse_num = 5"
                s = s & " group by whse_num, lot_num, gmasize, lane_status order by lot_num, whse_num"
            End If
            Set rs = Wdb.Execute(s)
            If rs.BOF = False Then
                rs.MoveFirst
                Do Until rs.EOF
                    If qstr = "CraneAll" Then
                        s = " " & Chr(9)
                    Else
                        s = rs!lot_num & Chr(9)
                    End If
                    s = s & "SR-" & rs!whse_num & Chr(9)
                    If qstr = "CraneAll" Then
                        s = s & rs(3) & Chr(9)
                    Else
                        s = s & rs(4) & Chr(9)
                    End If
                    If rs!lane_status = "H" Then
                        s = s & "On Hold"
                    Else
                        If rs!lane_status = "B" Then
                            s = s & "Blocked"
                        Else
                            If rs!gmasize > 0 Then s = s & "4 Way"
                        End If
                    End If
                    Grid1.AddItem s
                    rs.MoveNext
                Loop
            End If
            rs.Close
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.ColWidth(0) = 700
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 600
    Grid1.ColWidth(3) = 800
End Sub

Private Sub outgoing_list()
    Dim i As Integer, s As String, psku As String
    If Form4.RGrid.Rows <= 1 Then Exit Sub
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    psku = ".."
    For i = 1 To Form4.RGrid.Rows - 1
        If Form4.RGrid.TextMatrix(i, 3) <> psku Then
            s = "ng" & Chr(9) & Form4.RGrid.TextMatrix(i, 3) & Chr(9)
            s = s & Form4.RGrid.TextMatrix(i, 11) & Chr(9)
            s = s & Form4.RGrid.TextMatrix(i, 2)
            Grid1.AddItem s
            psku = Form4.RGrid.TextMatrix(i, 3)
        End If
        s = Form4.RGrid.TextMatrix(i, 4) & Chr(9)
        s = s & Form4.RGrid.TextMatrix(i, 1) & Chr(9)
        s = s & Form4.RGrid.TextMatrix(i, 5) & Chr(9)
        If Form4.RGrid.TextMatrix(i, 9) = "1" Then
            s = s & "1st out"
        End If
        Grid1.AddItem s
    Next i
End Sub

Private Sub Command1_Click()
    Call phonebook(Printer)
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
    If qstr = "BB Pallets" Then
        refresh_grid
    Else
        If qstr = "outgoing" Then
            outgoing_list
        Else
            If qstr = "Crane1" Or qstr = "Crane2" Or qstr = "Crane3" Or qstr = "Crane5" Or qstr = "CraneAll" Then
                refresh_cranes
            Else
                refresh_query
            End If
        End If
    End If
    Call phonebook(Picture1)
End Sub

Private Sub VScroll1_Change()
    Picture1.Move Picture1.Left, Frame1.Height - VScroll1.Value
End Sub

