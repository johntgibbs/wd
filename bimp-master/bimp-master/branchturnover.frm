VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form branchturnover 
   Caption         =   "Branch Pallet Turnover"
   ClientHeight    =   10875
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   0
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid hgrid 
      Height          =   10815
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   19076
      _Version        =   327680
      BackColorSel    =   192
      WordWrap        =   -1  'True
      FocusRect       =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Plant:"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label ycolor 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ycolor"
      Height          =   255
      Left            =   8160
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label bcolor 
      BackColor       =   &H00FFFFC0&
      Caption         =   "bcolor"
      Height          =   255
      Left            =   8040
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label wcolor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "wcolor"
      Height          =   255
      Left            =   8040
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Menu prtmenu 
      Caption         =   "&Print"
   End
   Begin VB.Menu sortmenu 
      Caption         =   "&Sort"
      Begin VB.Menu sortsales 
         Caption         =   "Sales"
         Checked         =   -1  'True
      End
      Begin VB.Menu sortturn 
         Caption         =   "Turnover"
      End
      Begin VB.Menu sortdoh 
         Caption         =   "Days On Hand"
      End
      Begin VB.Menu sortwhs 
         Caption         =   "Warehouse"
      End
   End
   Begin VB.Menu projmenu 
      Caption         =   "Projections"
   End
End
Attribute VB_Name = "branchturnover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fetch_branch_pallets()
    Dim i As Integer, s As String, ds As ADODB.Recordset
    Dim ppal As String, ptot As Long, ptype As String
    Dim t3gal As Long, ttray As Long, itot As Long, isku As Integer
    Dim psku As String
    For i = 1 To Grid1.Rows - 2
        ptot = 0: t3gal = 0: ttray = 0: psku = " "
        s = "select sku, onhand from bimp where branchwhs = '" & Grid1.TextMatrix(i, 0) & "'"
        s = s & " and plantwhs <> 'DRY' order by sku"                          'jv020516
        Set ds = wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                If ds!sku <> psku Then
                    isku = Val(ds!sku)
                    If skurec(isku).sku = ds!sku Then
                        ppal = skurec(isku).pallet
                        ptype = Left(skurec(isku).unit & ".", 1)
                    Else
                        ppal = 0: ptype = " "
                    End If
                    If ds!onhand > 0 And ppal > 0 Then
                        If ptype = "3" Then
                            t3gal = t3gal + ds!onhand
                        Else
                            If ptype = "T" Then
                                ttray = ttray + ds!onhand
                            Else
                                ptot = ptot + Int((ds!onhand / ppal) + 0.999)
                            End If
                        End If
                    End If
                    psku = ds!sku
                End If
                ds.MoveNext
            Loop
            ds.Close
        End If
        ptot = ptot + Int((t3gal / 60) + 0.999)
        ptot = ptot + Int((ttray / 132) + 0.999)
        Grid1.TextMatrix(i, 5) = ptot
        DoEvents
    Next i
End Sub



Private Sub refresh_grid()
    Dim i As Integer, s As String, k As Integer, w As Integer, c As Long
    Dim rt As String, rh As String, rf As String, ps As String, po As String, j As Integer
    Dim ds As ADODB.Recordset, ss As ADODB.Recordset
    Dim pcnt As Currency, icnt As Long, rcnt As Long
    Dim pcap As Long, puse As Long, psku As String              'jv082416
    Dim sales3 As Long, onhand3 As Long
    Screen.MousePointer = 11
    hgrid.Redraw = False
    hgrid.FontName = "Arial"
    hgrid.FontBold = True
    hgrid.FontSize = 8
    
    hgrid.Clear: hgrid.Rows = 1: hgrid.Cols = 15                    'jv020118
    hgrid.FillStyle = flexFillRepeat
    s = "select gemmsid, branchname, modem, fax from branches where modem > '0' and fax > '0'"
    If Combo1 = "ALL" Then
        s = s & " and gemmsid > '0' order by gemmsid"
    Else
        's = s & " and gemmsid in (select branchwhs from bimp where plantwhs = '" & Combo1 & "') order by gemmsid"
        s = s & " and gemmsid in (select listreturn from valuelists where listname = 'branchplants'"    'jv030316
        s = s & " and listdisplay = '" & Combo1 & "')"                                                  'jv030316
        s = s & " order by gemmsid"                                                                     'jv030316
    End If
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            pcnt = 0: icnt = 0: rcnt = 0
            sales3 = 0: onhand3 = 0
            psku = " "
            s = "select sku, onhand, onorder, sales from bimp where branchwhs = '" & ds!gemmsid & "'"
            'If Combo1 <> "ALL" Then s = s & " and plantwhs in ('" & Combo1 & "', 'VENDOR')"
            s = s & " and plantwhs <> 'DRY' order by sku"
            Set ss = wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                Do Until ss.EOF
                    If ss!sku <> psku Then
                        j = Val(ss!sku)
                        If skurec(j).sku = ss!sku Then
                            If ss!sales > 0 Then
                                ps = Format(ss!sales / skurec(j).pallet, ".00")
                            Else
                                ps = "0"
                            End If
                            If ss!onhand > 0 Or ss!onorder > 0 Then
                                po = Format((ss!onhand + ss!onorder) / skurec(j).pallet, ".00")
                            Else
                                po = " "
                            End If
                        Else
                            ps = " "
                            po = " "
                        End If
                        If skurec(j).unit = "3GAL" Then
                            sales3 = sales3 + ss!sales
                            onhand3 = onhand3 + ss!onhand + ss!onorder
                        Else
                            pcnt = pcnt + CInt(Val(po) + 0.499) ' 0.499 is forcing a round up
                            rcnt = rcnt + CInt(Val(ps))
                        End If
                        psku = ss!sku
                    End If
                    ss.MoveNext
                Loop
            End If
            ss.Close
            If onhand3 > 0 Then pcnt = pcnt + CInt((onhand3 / 60) + 0.499) ' 60 is the pallet size for 3 gallons. 0.499 is forcing a round up.
            If sales3 > 0 Then rcnt = rcnt + CInt((sales3 / 60) + 0.499)
            'MsgBox ds!gemmsid & " " & onhand3 & " " & sales3
            
            If pcnt > 0 Then
                pcap = Val(ds!modem)
                puse = Val(ds!fax)
                s = ds!gemmsid & Chr(9)
                s = s & ds!branchname & Chr(9)
                s = s & ds!modem & Chr(9)                                               'Capacity
                s = s & ds!fax & Chr(9)                                                 'Usable
                s = s & Format(pcnt, "0") & Chr(9)                                      'Onhand + Onorder
                If pcap <> 0 Then s = s & Format((pcnt / pcap) * 100, "0") & " %"       '%Cap in use
                s = s & Chr(9)
                If puse <> 0 Then s = s & Format((pcnt / puse) * 100, "0") & " %"       '%Usable in use
                s = s & Chr(9)
                s = s & Format(rcnt, "0") & Chr(9)                                      'Sales
                If pcap <> 0 Then s = s & Format((rcnt / pcap) * 100, "0") & " %"       '%Cap Turnover
                s = s & Chr(9)
                If puse <> 0 Then s = s & Format((rcnt / puse) * 100, "0") & " %"       '%Use Turnover
                s = s & Chr(9)
                s = s & Format(pcnt - rcnt, "0") & Chr(9)                               'Onhand - Sold
                If rcnt <> 0 And pcnt <> 0 Then s = s & Format((pcnt / rcnt) * 100, "0") & " %" '%Sales on Hand
                s = s & Chr(9)
                If rcnt <> 0 Then s = s & Format(rcnt / 30, ".00")                      'Pallet Per Day
                s = s & Chr(9)
                If rcnt <> 0 Then s = s & CInt(puse / (rcnt / 30))                      'Turnover Days
                s = s & Chr(9)              'jv020118
                If rcnt <> 0 And pcnt <> 0 Then s = s & Format((pcnt / rcnt) * 30, "0") 'Days Supply
                hgrid.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    ' ----------- Plants ---------------------
    If Combo1 = "ALL" Then s = "select gemmsid, branchname, modem, fax, branch from branches where branch in (1, 47, 52)"
    If Combo1 = "A10" Then s = "select gemmsid, branchname, modem, fax, branch from branches where branch = 52"
    If Combo1 = "K10" Then s = "select gemmsid, branchname, modem, fax, branch from branches where branch = 47"
    If Combo1 = "T10" Then s = "select gemmsid, branchname, modem, fax, branch from branches where branch = 1"
    
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            pcnt = 0: icnt = 0: rcnt = 0
            sales3 = 0: onhand3 = 0
            psku = " "
            s = "select sku, onhand, onorder, sales from bimp where branchwhs = '" & Format(ds!branch, "000") & "'"
            s = s & " and plantwhs <> 'DRY' order by sku"
            Set ss = wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                Do Until ss.EOF
                    If ss!sku <> psku Then
                        j = Val(ss!sku)
                        If skurec(j).sku = ss!sku Then
                            If ss!sales > 0 Then
                                ps = Format(ss!sales / skurec(j).pallet, ".00")
                            Else
                                ps = "0"
                            End If
                            If ss!onhand > 0 Or ss!onorder > 0 Then
                                po = Format((ss!onhand + ss!onorder) / skurec(j).pallet, ".00")
                            Else
                                po = " "
                            End If
                        Else
                            ps = " "
                            po = " "
                        End If
                        If skurec(j).unit = "3GAL" Then
                            sales3 = sales3 + ss!sales
                            onhand3 = onhand3 + ss!onhand + ss!onorder
                        Else
                            pcnt = pcnt + CInt(Val(po) + 0.499)
                            rcnt = rcnt + CInt(Val(ps))
                        End If
                        psku = ss!sku
                    End If
                    ss.MoveNext
                Loop
            End If
            ss.Close
            If onhand3 > 0 Then pcnt = pcnt + CInt((onhand3 / 60) + 0.499)
            If sales3 > 0 Then rcnt = rcnt + CInt((sales3 / 60) + 0.499)
            'MsgBox ds!gemmsid & " " & onhand3 & " " & sales3
            
            pcap = pcnt 'Val(ds!modem)
            puse = pcnt 'Val(ds!fax)
            s = Format(ds!branch, "000") & Chr(9)
            s = s & ds!branchname & Chr(9)
            s = s & pcnt & Chr(9) 'ds!modem & Chr(9)
            s = s & pcnt & Chr(9) 'ds!fax & Chr(9)
            s = s & Format(pcnt, "0") & Chr(9)
            If pcap <> 0 Then s = s & Format((pcnt / pcap) * 100, "0") & " %"
            s = s & Chr(9)
            If puse <> 0 Then s = s & Format((pcnt / puse) * 100, "0") & " %"
            s = s & Chr(9)
            s = s & Format(rcnt, "0") & Chr(9)
            If pcap <> 0 Then s = s & Format((rcnt / pcap) * 100, "0") & " %"
            s = s & Chr(9)
            If puse <> 0 Then s = s & Format((rcnt / puse) * 100, "0") & " %"
            s = s & Chr(9)
            s = s & Format(pcnt - rcnt, "0") & Chr(9)
            If rcnt <> 0 And pcnt <> 0 Then s = s & Format((pcnt / rcnt) * 100, "0") & " %"
            s = s & Chr(9)
            If rcnt <> 0 Then s = s & Format(rcnt / 30, ".00")
            s = s & Chr(9)
            If rcnt <> 0 Then
                s = s & CInt(puse / (rcnt / 30))
            Else
                s = s & Format(pcnt, "0")
            End If
            s = s & Chr(9)              'jv020118
            If rcnt <> 0 And pcnt <> 0 Then s = s & Format((pcnt / rcnt) * 30, "0") 'Days Supply
            hgrid.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    
    'Open "c:\bto.txt" For Output As #1
    'For i = 1 To hgrid.Rows - 1
    '    Write #1, hgrid.TextMatrix(i, 0), hgrid.TextMatrix(i, 13)
    'Next i
    'Close #1
    
    hgrid.Row = 1: hgrid.RowSel = 1
    'hgrid.Col = 13: hgrid.ColSel = 13
    If sortwhs.Checked = True Then hgrid.Col = 0: hgrid.ColSel = 0
    If sortsales.Checked = True Then hgrid.Col = 12: hgrid.ColSel = 12
    If sortturn.Checked = True Then hgrid.Col = 13: hgrid.ColSel = 13
    If sortdoh.Checked = True Then hgrid.Col = 14: hgrid.ColSel = 14            'jv020118
    If sortsales.Checked = True Then
        hgrid.Sort = 4
    Else
        If sortturn.Checked = True Or sortdoh.Checked = True Then               'jv020118
            hgrid.Sort = 3
        Else
            hgrid.Sort = 5
        End If
    End If
    
    pcap = 0: puse = 0: pcnt = 0: rcnt = 0
    For i = 1 To hgrid.Rows - 1
        pcap = pcap + Val(hgrid.TextMatrix(i, 2))
        puse = puse + Val(hgrid.TextMatrix(i, 3))
        pcnt = pcnt + Val(hgrid.TextMatrix(i, 4))
        rcnt = rcnt + Val(hgrid.TextMatrix(i, 7))
    Next i
    's = "---" & Chr(9)
    s = "ALL" & Chr(9)
    s = s & "Summary" & Chr(9)
    s = s & pcap & Chr(9)
    s = s & puse & Chr(9)
    s = s & Format(pcnt, "0") & Chr(9)
    If pcap <> 0 Then s = s & Format((pcnt / pcap) * 100, "0") & " %"
    s = s & Chr(9)
    If puse <> 0 Then s = s & Format((pcnt / puse) * 100, "0") & " %"
    s = s & Chr(9)
    s = s & Format(rcnt, "0") & Chr(9)
    If pcap <> 0 Then s = s & Format((rcnt / pcap) * 100, "0") & " %"
    s = s & Chr(9)
    If puse <> 0 Then s = s & Format((rcnt / puse) * 100, "0") & " %"
    s = s & Chr(9)
    s = s & Format(pcnt - rcnt, "0") & Chr(9)
    If rcnt <> 0 And pcnt <> 0 Then s = s & Format((pcnt / rcnt) * 100, "0") & " %"
    s = s & Chr(9)
    If rcnt <> 0 Then s = s & Format(rcnt / 30, ".00")
    s = s & Chr(9)
    If rcnt <> 0 Then s = s & CInt(puse / (rcnt / 30))
    s = s & Chr(9)              'jv020118
    If rcnt <> 0 And pcnt <> 0 Then s = s & Format((pcnt / rcnt) * 30, "0") 'Days Supply
    hgrid.AddItem s
    
    c = wcolor.BackColor
    For i = 1 To hgrid.Rows - 1
        hgrid.Row = i: hgrid.RowSel = i
        hgrid.Col = 1: hgrid.ColSel = hgrid.Cols - 1
        hgrid.CellBackColor = c
        If c = wcolor.BackColor Then
            c = bcolor.BackColor
        Else
            c = wcolor.BackColor
        End If
    Next i
    's = "^Whs|<Location|^Capacity|^Usable"
    's = s & "|^OnHand|^%Cap OH|^%Use OH|^Loads|^%Cap RO|^%Use RO|^PalDiff|^%SalesOH|^Pals/Day"
    s = "^Whs|<Location|^Storage Capacity|^Usable Storage"
    's = s & "|^OnHand|^%Cap In Use|^%Use In Use|^Loads|^%Cap Turnover|^%Use Turnover|^OnHand -Loads|^%Loads OnHand|^Pallets per Day|^TurnDays"
    s = s & "|^OnHand + OnOrder|^%Cap In Use|^%Usable In Use|^Sold 30 Days|^%Cap Turnover|^%Use Turnover|^OnHand - Sold|^%Sales OnHand|^Pallets per Day|^Turnover Days|^Days OnHand"
    
    hgrid.FormatString = s
    hgrid.ColWidth(0) = 600
    hgrid.ColWidth(1) = 1700
    hgrid.ColWidth(2) = 900
    hgrid.ColWidth(3) = 900
    hgrid.ColWidth(4) = 900
    hgrid.ColWidth(5) = 1000
    hgrid.ColWidth(6) = 1000
    hgrid.ColWidth(7) = 900
    hgrid.ColWidth(8) = 1000
    hgrid.ColWidth(9) = 1000
    hgrid.ColWidth(10) = 800
    hgrid.ColWidth(11) = 1000
    hgrid.ColWidth(12) = 1080
    hgrid.ColWidth(13) = 900
    hgrid.ColWidth(14) = 900
    Screen.MousePointer = 0
    
    hgrid.Redraw = True
    hgrid.RowHeight(0) = hgrid.RowHeight(1) * 2
    hgrid.Row = 1
End Sub

Private Sub Combo1_Click()
    refresh_grid
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    Combo1.Clear
    Combo1.AddItem "ALL"
    Combo1.AddItem "A10"
    Combo1.AddItem "K10"
    Combo1.AddItem "T10"
    Combo1.ListIndex = 0
    'refresh_grid
End Sub

Private Sub Form_Resize()
    hgrid.Width = Me.Width - 120
    If Me.Height > 2000 Then hgrid.Height = Me.Height - (720 + Combo1.Height) '680
End Sub

Private Sub projmenu_Click()
    Dim s As String
    If hgrid.TextMatrix(hgrid.Row, 0) = "ALL" Then
        s = InputBox("Sale Days", "Days of Sales...", "30")
    Else
        s = InputBox("Sale Days", "Days of Sales...", Val(hgrid.TextMatrix(hgrid.Row, hgrid.Cols - 1)) + 14)
    End If
    If Len(s) = 0 Then Exit Sub
    salesproj.Text1 = Val(s)
    salesproj.Show
    If hgrid.TextMatrix(hgrid.Row, 0) = "ALL" Then
        'salesproj.Text1 = Val(s)
        salesproj.bcode = "ALL"
    Else
        'salesproj.Text1 = Val(hgrid.TextMatrix(hgrid.Row, hgrid.Cols - 1)) + 14
        salesproj.bcode = hgrid.TextMatrix(hgrid.Row, 0)
    End If
    'salesproj.Show
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    rt = Combo1 & " Branch Pallet Turnover - 30 Days"
    rh = Format(DateAdd("d", -30, Now), "m-d-yyyy")
    rh = rh & " Thru " & Format(Now, "m-d-yyyy")
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    htdc(0) = "lightcyan": gndc(0) = Me.bcolor.BackColor
    htdc(1) = "yellow": gndc(1) = Me.ycolor.BackColor
    'htdc(2) = "lightgrey": gndc(2) = Me.wcolor.BackColor
    htdc(2) = "white": gndc(2) = Me.wcolor.BackColor
    hgrid.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", hgrid, rt, rh, rf, "linen", "lightyellow", "white")
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        hgrid.Redraw = True: hgrid.Row = 1
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", hgrid, rt, rh, rf, "linen", "lightyellow", "white")
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        hgrid.Redraw = True: hgrid.Row = 1
        Exit Sub
    End If

End Sub

Private Sub sortdoh_Click()                     'jv020118
    sortsales.Checked = False
    sortturn.Checked = False
    sortwhs.Checked = False
    sortdoh.Checked = True
    refresh_grid
End Sub

Private Sub sortsales_Click()
    sortsales.Checked = True
    sortturn.Checked = False
    sortwhs.Checked = False
    sortdoh.Checked = False
    refresh_grid
End Sub

Private Sub sortturn_Click()
    sortsales.Checked = False
    sortturn.Checked = True
    sortwhs.Checked = False
    sortdoh.Checked = False
    refresh_grid
End Sub

Private Sub sortwhs_Click()
    sortsales.Checked = False
    sortturn.Checked = False
    sortwhs.Checked = True
    sortdoh.Checked = False
    refresh_grid
End Sub
