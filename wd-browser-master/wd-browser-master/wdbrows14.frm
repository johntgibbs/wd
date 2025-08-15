VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form14 
   Caption         =   "Form14"
   ClientHeight    =   10695
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11580
   LinkTopic       =   "Form14"
   ScaleHeight     =   10695
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid hgrid 
      Height          =   3495
      Left            =   8160
      TabIndex        =   12
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   6165
      _Version        =   327680
      ForeColor       =   128
      BackColorFixed  =   12648384
      Appearance      =   0
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send to Home Page for Printing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   975
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "wdbrows14.frx":0000
      Top             =   0
      Width           =   6855
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   9340
      _Version        =   327680
      Cols            =   4
      ForeColor       =   4210688
      BackColorFixed  =   12648447
      WordWrap        =   -1  'True
      FocusRect       =   0
      GridLines       =   2
   End
   Begin VB.Label ncolor 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ncolor"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label pcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "pcolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Surplus - Over 30"
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
      Left            =   6120
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "30 Day Supply"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2 Week Supply"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Below 2 Week Level"
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
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label whsno 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label qstr 
      Caption         =   "Label1"
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   6360
      Width           =   6255
   End
   Begin VB.Menu listmenu 
      Caption         =   "Lists"
      Begin VB.Menu allskus 
         Caption         =   "All Products"
         Checked         =   -1  'True
      End
      Begin VB.Menu wskus 
         Caption         =   "Below 2 Week Level"
      End
      Begin VB.Menu yskus 
         Caption         =   "2 Week Supply"
      End
      Begin VB.Menu bskus 
         Caption         =   "30 Day Supply"
      End
      Begin VB.Menu gskus 
         Caption         =   "Overstocked"
      End
      Begin VB.Menu promos 
         Caption         =   "Promotions"
      End
      Begin VB.Menu discprod 
         Caption         =   "Discontinued Products"
      End
      Begin VB.Menu nosales 
         Caption         =   "No Recent Sales"
      End
      Begin VB.Menu ordskus 
         Caption         =   "On Order"
      End
      Begin VB.Menu unsales 
         Caption         =   "Unit Sales"
         Begin VB.Menu ushg 
            Caption         =   "1/2 Gallons"
         End
         Begin VB.Menu us48 
            Caption         =   "48 oz"
         End
         Begin VB.Menu uspt 
            Caption         =   "Pints"
         End
         Begin VB.Menu usqt 
            Caption         =   "Quarts"
         End
         Begin VB.Menu us3g 
            Caption         =   "3 Gallons"
         End
         Begin VB.Menu us6p 
            Caption         =   "6 Pack"
         End
         Begin VB.Menu us12p 
            Caption         =   "12 Pack"
         End
         Begin VB.Menu us24p 
            Caption         =   "24 Pack"
         End
         Begin VB.Menu uscup 
            Caption         =   "Cups"
         End
         Begin VB.Menu usthome 
            Caption         =   "Take Home"
         End
         Begin VB.Menu usbulk 
            Caption         =   "Bulk Snacks"
         End
      End
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clear_list_checks()
    allskus.Checked = False
    wskus.Checked = False
    yskus.Checked = False
    bskus.Checked = False
    gskus.Checked = False
    promos.Checked = False
    discprod.Checked = False
    nosales.Checked = False
    ordskus.Checked = False
    ushg.Checked = False
    us48.Checked = False
    uspt.Checked = False
    usqt.Checked = False
    us3g.Checked = False
    us6p.Checked = False
    us12p.Checked = False
    us24p.Checked = False
    uscup.Checked = False
    usthome.Checked = False
    usbulk.Checked = False
End Sub

Private Sub branch_out_of_stock()
    Dim s As String, daysin As Integer, daysout As Integer, lostsales As Long, salesperday As Long
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim cfile As String
    cfile = "s:\wd\html\boutstock.csv"
    Grid1.Redraw = False: Grid1.Font = "Arial": Grid1.FontBold = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    Open cfile For Input As #2
    Do Until EOF(2)
        Input #2, f0, f1, f2, f3, f4, f5, f6, f7
        If f4 = Me.whsno Then
            s = f0 & Chr(9)
            s = s & f1 & Chr(9)
            s = s & skurec(Val(f1)).unit & " " & skurec(Val(f1)).desc & Chr(9)
            s = s & f2 & Chr(9)
            s = s & f3 & "-"
            If f3 = "A10" Then s = s & "Sylacauga"
            If f3 = "K10" Then s = s & "Broken Arrow"
            If f3 = "T10" Then s = s & "Brenham"
            s = s & Chr(9)
            s = s & f5 & Chr(9)
            s = s & f6 & Chr(9)
            s = s & f7 & Chr(9)
            daysin = DateDiff("d", f5, f6) + 1
            s = s & daysin & Chr(9)
            salesperday = f7 / daysin
            s = s & salesperday & Chr(9)
            daysout = DateDiff("d", f6, Now)
            s = s & daysout & Chr(9)
            'lostsales = f7 * (daysout / daysin)
            lostsales = salesperday * daysout
            s = s & lostsales
            Grid1.AddItem s
        End If
    Loop
    Close #2
    s = "^ID|^SKU|<Product|^PalSize|^Supplier|^Last Transport Receipt|^Last Load Date|^Sales|^Days In Stock|^Sales Per Day|^Days Out|^Lost Sales"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 0 '1000
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 0 '1000
    Grid1.ColWidth(4) = 2400
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 1200
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1400
    Grid1.ColWidth(9) = 1400
    Grid1.ColWidth(10) = 1200
    Grid1.ColWidth(11) = 1200
    Grid1.Redraw = True
    Text1 = "Branch Out of Stock"
End Sub

Private Sub branch_turnover(pbr As String)
    Dim i As Integer, s As String, k As Integer, w As Integer, c As Long
    Dim rt As String, rh As String, rf As String, ps As String, po As String, j As Integer
    Dim ds As ADODB.Recordset, ss As ADODB.Recordset
    Dim pcnt As Currency, icnt As Long, rcnt As Long
    Dim pcap As Long, puse As Long, psku As String                  'jv082416
    Dim sales3 As Integer, onhand3 As Integer
    'Screen.MousePointer = 11
    hgrid.Redraw = False
    hgrid.FontName = "Arial"
    hgrid.FontBold = True
    hgrid.FontSize = 8
    
    hgrid.Clear: hgrid.Rows = 1: hgrid.Cols = 2 '14
    hgrid.FillStyle = flexFillRepeat
    s = "select gemmsid, branchname, modem, fax from branches where modem > '0' and fax > '0'"
    'If Combo1 = "ALL" Then
    '    s = s & " and gemmsid > '0' order by gemmsid"
    'Else
    '    's = s & " and gemmsid in (select branchwhs from bimp where plantwhs = '" & Combo1 & "') order by gemmsid"
    '    s = s & " and gemmsid in (select listreturn from valuelists where listname = 'branchplants'"    'jv030316
    '    s = s & " and listdisplay = '" & Combo1 & "')"                                                  'jv030316
    '    s = s & " order by gemmsid"                                                                     'jv030316
    'End If
    s = s & " and gemmsid = '" & pbr & "'"
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
            
            If pcnt > 0 Then
                pcap = Val(ds!modem)
                puse = Val(ds!fax)
                's = ds!gemmsid & Chr(9)                                                     '0
                's = s & ds!branchname & Chr(9)                                              '1
                's = s & ds!modem & Chr(9)                                                   '2
                's = s & ds!fax & Chr(9)                                                     '3
                's = s & Format(pcnt, "0") & Chr(9)                                          '4
                'If pcap <> 0 Then s = s & Format((pcnt / pcap) * 100, "0") & " %"
                's = s & Chr(9)                                                              '5
                'If puse <> 0 Then s = s & Format((pcnt / puse) * 100, "0") & " %"
                's = s & Chr(9)                                                              '6
                's = s & Format(rcnt, "0") & Chr(9)                                          '7
                'If pcap <> 0 Then s = s & Format((rcnt / pcap) * 100, "0") & " %"
                's = s & Chr(9)                                                              '8
                'If puse <> 0 Then s = s & Format((rcnt / puse) * 100, "0") & " %"
                's = s & Chr(9)                                                              '9
                's = s & Format(pcnt - rcnt, "0") & Chr(9)                                   '10
                'If rcnt <> 0 And pcnt <> 0 Then s = s & Format((pcnt / rcnt) * 100, "0") & " %"
                's = s & Chr(9)                                                              '11
                'If rcnt <> 0 Then s = s & Format(rcnt / 30, ".00")
                's = s & Chr(9)                                                              '12
                'If rcnt <> 0 Then s = s & CInt(puse / (rcnt / 30))
                'hgrid.AddItem s                                                             '13
                hgrid.AddItem ds!branchname
                s = "Storage Capacity" & Chr(9) & ds!modem
                hgrid.AddItem s
                s = "OnHand + OnOrder" & Chr(9) & pcnt
                hgrid.AddItem s
                If pcap <> 0 Then
                    's = "% Capacity in Use" & Chr(9)
                    s = "Capacity in Use" & Chr(9)
                    s = s & Format((pcnt / pcap) * 100, "0") & " %"
                    hgrid.AddItem s
                End If
                s = "Sold Last 30 Days" & Chr(9) & Format(rcnt, "0")
                hgrid.AddItem s
                If pcap <> 0 Then
                    's = "% Capacity Turnover" & Chr(9)
                    s = "Capacity Turnover" & Chr(9)
                    s = s & Format((rcnt / pcap) * 100, "0") & " %"
                    hgrid.AddItem s
                End If
                s = "OnHand - Sold" & Chr(9) & Format(pcnt - rcnt, "0")
                hgrid.AddItem s
                If rcnt <> 0 And pcnt <> 0 Then
                    's = "% Sales On Hand" & Chr(9)
                    s = "Sales On Hand" & Chr(9)
                    s = s & Format((pcnt / rcnt) * 100, "0") & " %"
                    hgrid.AddItem s
                End If
                If rcnt <> 0 Then
                    s = "Pallets per Day" & Chr(9)
                    s = s & Format(rcnt / 30, ".00")
                    hgrid.AddItem s
                End If
                If rcnt <> 0 Then
                    s = "Turnover Days" & Chr(9)
                    s = s & CInt(puse / (rcnt / 30))
                    hgrid.AddItem s
                End If
                If rcnt <> 0 And pcnt <> 0 Then                                 'jv020118
                    s = "Days OnHand" & Chr(9)                                  'jv020118
                    s = s & Format((pcnt / rcnt) * 30, "0") 'Days Supply        'jv020118
                    hgrid.AddItem s                                             'jv020118
                End If                                                          'jv020118
                
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    hgrid.FormatString = "<Pallet Turnover|^"
    hgrid.ColWidth(0) = 1800 '2000
    hgrid.ColWidth(1) = 800 '1000
    hgrid.Redraw = True
    hgrid.Visible = True
End Sub

Private Sub ticket_barcodes()
    Dim s As String, f1 As String, f2 As String, f3 As String, f4 As String, f0 As String
    Dim t As String, k As String, bname As String, i As Integer
    'branchrec(Val(bcode)).branchname
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    cfile = "s:\wd\html\brbarcodes.txt"
    Open cfile For Input Shared As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4
        k = Format(DateDiff("d", f1, Now), "0") & f0 & f4
        'If Left(f2, Len(branchrec(Val(whsno)).branchname)) = UCase(branchrec(Val(whsno)).branchname) Then
        If Left(f2, Len(f2) - 3) = UCase(branchrec(Val(whsno)).branchname) Then
            If Format(DateDiff("d", f1, Now), "0") & f0 <> t Then
                If Grid1.Rows > 1 Then Grid1.AddItem " " & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & t & "99999"
                t = Format(DateDiff("d", f1, Now), "0") & f0
            End If
            bc = Mid(f4, 5, 6) & "  " & Mid(f4, 11, 3) & "  " & Mid(f4, 14, 3)
            's = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & f4 & Chr(9) & k
            s = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & bc & Chr(9) & k
            Grid1.AddItem s
        End If
    Loop
    If Grid1.Rows > 1 Then Grid1.AddItem " " & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & t & "99999"
    Close #1
    If Grid1.Rows > 1 Then
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 5: Grid1.ColSel = 5
        Grid1.Sort = 5
        t = Grid1.TextMatrix(Grid1.Rows - 1, 5)
        For i = Grid1.Rows - 2 To 1 Step -1
            If Grid1.TextMatrix(i, 5) <> t Then
                t = Grid1.TextMatrix(i, 5)
            Else
                Grid1.RemoveItem i
            End If
        Next i
    End If
    'Grid1.FormatString = "^Ticket|^Ship Date|<Trailer|<Product|^BarCode"
    Grid1.FormatString = "^Group|^Loaded|<Trailer|<Product|^BarCode"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 2400
    Grid1.ColWidth(3) = 3500
    Grid1.ColWidth(4) = 2000
    Grid1.ColWidth(5) = 0 '3000
    Grid1.Redraw = True
    Text1 = "Ticket BarCodes"
End Sub

Private Sub tstation_rpt()
    Dim cfile As String, s As String, f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim pdaze As Integer, i As Integer
    Text1.Visible = False
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 10
    cfile = Form1.webdir & "\counts\tstation" & Me.whsno & ".csv"
    'MsgBox cfile
    If Len(Dir(cfile)) > 0 Then
        Text2 = "Updated: " & Format(FileDateTime(cfile), "M-d-yyyy h:mm am/pm")  'jv032318
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
            s = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & f4 & Chr(9)
            s = s & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9) & f9
            Grid1.AddItem s
        Loop
        Close #1
        Grid1.FormatString = "^SKU|<Product|^Count Date|^Beg Inventory|^Trans In|^Trans Out|^Net|^30 Day Sales|^Days Supply"
        Grid1.ColWidth(0) = 800
        Grid1.ColWidth(1) = 4000
        Grid1.ColWidth(2) = 1500 '2000
        Grid1.ColWidth(3) = 1500 '2000
        Grid1.ColWidth(4) = 1500 '2000
        Grid1.ColWidth(5) = 1500 '2000
        Grid1.ColWidth(6) = 1500 '2000
        Grid1.ColWidth(7) = 1500 '2000
        Grid1.ColWidth(8) = 1500 '2000
        Grid1.ColWidth(9) = 0
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 8) > " " Then
                pdaze = Val(Grid1.TextMatrix(i, 8))
            Else
                pdaze = 30
            End If
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
            If pdaze < 14 Then
                Grid1.CellBackColor = wcolor.BackColor
            Else
                If pdaze < 30 Then
                    Grid1.CellBackColor = ycolor.BackColor
                Else
                    If pdaze > 35 Then
                        Grid1.CellBackColor = gcolor.BackColor
                    Else
                        Grid1.CellBackColor = bcolor.BackColor
                    End If
                End If
            End If
        Next i
        
    End If
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub trailer_status_rpt()
    Dim ds As ADODB.Recordset, ts As ADODB.Recordset, s As String, tstat As String, q As String
    Dim pno As String, pt10 As Integer, pk10 As Integer, pa10 As Integer            'jv030216
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    s = "select * from runs where trldate >= '" & Format(Now, "M-d-yyyy") & "'"
    s = s & " and destination = '" & Val(whsno) & "'"               'jv121115
    s = s & " and trlno not in ('ZO', 'OP', 'QC')"                  'jv121115
    s = s & " order by trldate, trlno, startime"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = Format(ds!trldate, "dddd M-d-yyyy") & " "
            s = s & branchrec(Val(ds!destination)).branchname & " " & ds!trlno & vbCrLf
            s = s & " From: " & plantrec(Val(ds!loaded)).plantname & " "
            tstat = "In Process"
            If ds!runstat > " " Then tstat = ds!runstat
            'If ticket_post(ds!id) = True Then tstat = "In Transit"
            'If ticket_receipt(ds!id) = True Then tstat = "Received"
            s = s & tstat
            'Grid1.AddItem tstat & Chr(9) & s & Chr(9) & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
            
            q = "select sku, sum(pallets), sum(wraps), sum(units) from trailers"
            q = q & " where runid = " & ds!id & " group by sku"
            Set ts = wdb.Execute(q)
            If ts.BOF = False Then
                ts.MoveFirst
                Grid1.AddItem tstat & Chr(9) & s & Chr(9) & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
                Do Until ts.EOF
                    s = ts!sku & Chr(9)
                    s = s & skurec(Val(ts!sku)).unit & " " & skurec(Val(ts!sku)).desc & Chr(9)
                    s = s & ts(1) & Chr(9)
                    s = s & ts(2) & Chr(9)
                    s = s & ts(3)
                    Grid1.AddItem s
                    ts.MoveNext
                Loop
            End If
            ts.Close
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    ' Check Branch Orders                                                       'jv081516
    pno = "0": odate = " "
    s = "select * from brorders where branch = " & Val(whsno) & " and netqty > 0"
    s = s & " order by orddate, plant, sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!Plant <> pno Or Format(ds!orddate, "M-d-yyyy") <> odate Then
                s = "In Process" & Chr(9) & Format(ds!orddate, "dddd M-d-yyyy") & " "
                s = s & branchrec(ds!branch).branchname & " Orders" & vbCrLf
                s = s & " From: " & plantrec(Val(ds!Plant)).plantname & " In Process"
                s = s & Chr(9) & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
                Grid1.AddItem s
                pno = ds!Plant: odate = Format(ds!orddate, "M-d-yyyy")
            End If
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!netqty & Chr(9)
            s = s & ds!partqty & Chr(9)
            s = s & Format(ds!netqty * skurec(Val(ds!sku)).pallet, "0")
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    'Find pallet qtys in groupitems that have not been posted to trailers.      'jv081516
    s = "select * from runs where destination = '" & whsno & "'"                'jv081916
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "select * from trgroups where run1 = " & ds!id
            s = s & " or run2 = " & ds!id
            s = s & " or run3 = " & ds!id
            s = s & " or run4 = " & ds!id
            Set ts = wdb.Execute(s)
            If ts.BOF = False Then
                ts.MoveFirst
                Do Until ts.EOF
                    s = "select * from groupitems where groupcode = '" & ts!groupcode & "'"
                    s = s & " and groupcode not in (select groupcode from trailers)"
                    Set gs = wdb.Execute(s)
                    If gs.BOF = False Then
                        gs.MoveFirst
                        s = Format(ds!trldate, "dddd M-d-yyyy") & " "
                        s = s & branchrec(Val(whsno)).branchname & " " & ds!trlno & vbCrLf
                        s = s & " From: " & plantrec(Val(ds!loaded)).plantname & " "
                        tstat = "In Process"
                        If ds!runstat > " " Then tstat = ds!runstat
                        s = s & tstat
                        Grid1.AddItem tstat & Chr(9) & s & Chr(9) & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
                        
                        Do Until gs.EOF
                            s = gs!sku & Chr(9)
                            s = s & skurec(Val(gs!sku)).unit & " " & skurec(Val(gs!sku)).desc & Chr(9)
                            If ts!run1 = ds!id Then
                                If gs!qty1 > 0 Then                                             'jv081916
                                    s = s & Format(gs!qty1, "0") & Chr(9) & "0" & Chr(9)
                                    s = s & Format(gs!qty1 * skurec(Val(gs!sku)).pallet, "0")
                                    Grid1.AddItem s                                             'jv081916
                                End If                                                          'jv081916
                            End If
                            If ts!run2 = ds!id Then
                                If gs!qty2 > 0 Then                                             'jv081916
                                    s = s & Format(gs!qty2, "0") & Chr(9) & "0" & Chr(9)
                                    s = s & Format(gs!qty2 * skurec(Val(gs!sku)).pallet, "0")
                                    Grid1.AddItem s                                             'jv081916
                                End If                                                          'jv081916
                            End If
                            If ts!run3 = ds!id Then
                                If gs!qty3 > 0 Then                                             'jv081916
                                    s = s & Format(gs!qty3, "0") & Chr(9) & "0" & Chr(9)
                                    s = s & Format(gs!qty3 * skurec(Val(gs!sku)).pallet, "0")
                                    Grid1.AddItem s                                             'jv081916
                                End If                                                          'jv081916
                            End If
                            If ts!run4 = ds!id Then
                                If gs!qty4 > 0 Then                                             'jv081916
                                    s = s & Format(gs!qty4, "0") & Chr(9) & "0" & Chr(9)
                                    s = s & Format(gs!qty4 * skurec(Val(gs!sku)).pallet, "0")
                                    Grid1.AddItem s                                             'jv081916
                                End If                                                          'jv081916
                            End If
                            'Grid1.AddItem s
                            gs.MoveNext
                        Loop
                    End If
                    gs.Close
                    ts.MoveNext
                Loop
            End If
            ts.Close
            ds.MoveNext
        Loop
    End If
    ds.Close
    '-------------------------------------------------------------------------------
    
    
    pno = "0": pt10 = 0: pk10 = 0: pa10 = 0                                                     'jv030216
    s = "select plantwhs, sum(thiswknewpals) from bimp where branchwhs = '" & Format(Val(whsno), "000") & "'"      'jv030216
    s = s & " and thiswknewpals > 0"                                                            'jv030216
    s = s & " group by plantwhs"                                                                'jv030216
    Set ds = wdb.Execute(s)                                                                     'jv030216
    If ds.BOF = False Then                                                                      'jv030216
        ds.MoveFirst                                                                            'jv030216
        Do Until ds.EOF                                                                         'jv030216
            If ds!plantwhs = "T10" Then pt10 = pt10 + ds(1)                                     'jv030216
            If ds!plantwhs = "K10" Then pk10 = pk10 + ds(1)                                     'jv030216
            If ds!plantwhs = "A10" Then pa10 = pa10 + ds(1)                                     'jv030216
            ds.MoveNext                                                                         'jv030216
        Loop                                                                                    'jv030216
    End If                                                                                      'jv030216
    
    s = "select plantwhs, sku, thiswknewpals from bimp where branchwhs = '" & Format(Val(whsno), "000") & "'"
    s = s & " and thiswknewpals > 0"
    s = s & " order by plantwhs, sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!plantwhs <> pno Then
                s = "New" & ds!plantwhs & Chr(9)
                If ds!plantwhs = "T10" Then s = s & pt10 & " Additional Pallets Scheduled" & vbCrLf & "From: " & "Brenham"          'jv030216
                If ds!plantwhs = "K10" Then s = s & pk10 & " Additional Pallets Scheduled" & vbCrLf & "From: " & "Broken Arrow"     'jv030216
                If ds!plantwhs = "A10" Then s = s & pa10 & " Additional Pallets Scheduled" & vbCrLf & "From: " & "Sylacauga"        'jv030216
                's = s & "Additional Pallets Scheduled" & vbCrLf & "From: "
                'If ds!plantwhs = "T10" Then s = s & "Brenham"
                'If ds!plantwhs = "K10" Then s = s & "Broken Arrow"
                'If ds!plantwhs = "A10" Then s = s & "Sylacauga"
                s = s & " Later This Week." & Chr(9)
                s = s & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
                Grid1.AddItem s
                pno = ds!plantwhs
            End If
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!thiswknewpals & Chr(9)
            s = s & "0" & Chr(9)
            s = s & Format(ds!thiswknewpals * skurec(Val(ds!sku)).pallet, "0")
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    pno = "0": pt10 = 0: pk10 = 0: pa10 = 0                                                     'jv030216
    s = "select plantwhs, sum(nextwknewpals) from bimp where branchwhs = '" & Format(Val(whsno), "000") & "'"      'jv030216
    s = s & " and nextwknewpals > 0"                                                            'jv030216
    s = s & " group by plantwhs"                                                                'jv030216
    Set ds = wdb.Execute(s)                                                                     'jv030216
    If ds.BOF = False Then                                                                      'jv030216
        ds.MoveFirst                                                                            'jv030216
        Do Until ds.EOF                                                                         'jv030216
            If ds!plantwhs = "T10" Then pt10 = pt10 + ds(1)                                     'jv030216
            If ds!plantwhs = "K10" Then pk10 = pk10 + ds(1)                                     'jv030216
            If ds!plantwhs = "A10" Then pa10 = pa10 + ds(1)                                     'jv030216
            ds.MoveNext                                                                         'jv030216
        Loop                                                                                    'jv030216
    End If                                                                                      'jv030216
    
    s = "select plantwhs, sku, nextwknewpals from bimp where branchwhs = '" & Format(Val(whsno), "000") & "'"
    s = s & " and nextwknewpals > 0"
    s = s & " order by plantwhs, sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!plantwhs <> pno Then
                s = "New" & ds!plantwhs & Chr(9)
                If ds!plantwhs = "T10" Then s = s & pt10 & " Pallets Scheduled" & vbCrLf & "From: " & "Brenham"          'jv030216
                If ds!plantwhs = "K10" Then s = s & pk10 & " Pallets Scheduled" & vbCrLf & "From: " & "Broken Arrow"     'jv030216
                If ds!plantwhs = "A10" Then s = s & pa10 & " Pallets Scheduled" & vbCrLf & "From: " & "Sylacauga"        'jv030216
                's = s & "Pallets Scheduled" & vbCrLf & "From: "
                'If ds!plantwhs = "T10" Then s = s & "Brenham"
                'If ds!plantwhs = "K10" Then s = s & "Broken Arrow"
                'If ds!plantwhs = "A10" Then s = s & "Sylacauga"
                s = s & " Next Week." & Chr(9)
                s = s & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
                Grid1.AddItem s
                pno = ds!plantwhs
            End If
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!nextwknewpals & Chr(9)
            s = s & "0" & Chr(9)
            s = s & Format(ds!nextwknewpals * skurec(Val(ds!sku)).pallet, "0")
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    
    
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 0)) = 0 Then Grid1.RowHeight(i) = Grid1.RowHeight(i) * 4
            If Grid1.TextMatrix(i, 0) = "In Process" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("s:\wd\html\images\inproc.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = pcolor.BackColor
            End If
            If Grid1.TextMatrix(i, 0) = "In Transit" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("s:\wd\html\images\loaded.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ycolor.BackColor
            End If
            If Grid1.TextMatrix(i, 0) = "Received" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("s:\wd\html\stock\images\bbtruck.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = bcolor.BackColor
            End If
            If Grid1.TextMatrix(i, 0) = "NewT10" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("s:\wd\html\images\plant500.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ncolor.BackColor
            End If
            If Grid1.TextMatrix(i, 0) = "NewK10" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("s:\wd\html\images\plant501.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ncolor.BackColor
            End If
            If Grid1.TextMatrix(i, 0) = "NewA10" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("s:\wd\html\images\plant502.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ncolor.BackColor
            End If
            'Grid1.CellPicture = LoadPicture(Picture1.Picture)
        Next i
        Grid1.Row = 2: Grid1.Col = 1
    End If
    Text1 = "Trailer Status"
    Grid1.FormatString = "^|<|^|^|^"
    Grid1.ColWidth(0) = 1400
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1600
    Grid1.ColWidth(3) = 1600
    Grid1.ColWidth(4) = 1600
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub reco_rpt(pwhs As String)
    Dim ds As ADODB.Recordset, s As String, i As Integer, ppal As Integer
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7: Grid1.FixedCols = 1
    s = "select * from bimp where plantwhs = '" & pwhs & "' and branchwhs = '" & whsno & "'"
    s = s & " and paldiff < 0 order by paldiff"
    'Text1 = s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            ppal = skurec(Val(ds!sku)).pallet
            If ppal > 0 Then
                s = ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                s = s & Format(ds!sales / ppal, "#.00") & Chr(9)
                s = s & Format(ds!onhand / ppal, "#.00") & Chr(9)
                s = s & Format(ds!onorder / ppal, "#.00") & Chr(9)
                Grid1.AddItem s
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            sqty = Val(Grid1.TextMatrix(i, 2)) - Val(Grid1.TextMatrix(i, 3)) - Val(Grid1.TextMatrix(i, 4))
            Grid1.TextMatrix(i, 5) = Format(sqty, "#.00")
            Grid1.TextMatrix(i, 6) = CInt(sqty)
        Next i
    End If
    s = vbCrLf & "Recent Sales History indicates the following pallet orders are needed" & vbCrLf
    s = s & " to maintain suffecient inventory for the upcoming 2 week period."
    Text1 = s
    Grid1.FormatString = "^SKU|<Product|^Sales|^OnHand|^OnOrder|^Short|^New Order"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub oracle_onhand_rpt()
    Dim ds As ADODB.Recordset, s As String, psku As String, tp As Long, tu As Long, tn As Long
    Dim wcap As Integer, t3gal As Long
    psku = "0": tp = 0: tu = 0: tn = 0: t3gal = 0
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4: Grid1.FixedCols = 1
    wcap = branchrec(Val(whsno)).capacity
    s = "select * from bimp where branchwhs = '" & whsno & "' and onhand <> 0 order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!sku <> psku Then
                s = ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                If skurec(Val(ds!sku)).unit = "3GAL" Then                                       'jv022316
                    t3gal = t3gal + ds!onhand                                                   'jv022316
                    s = s & Chr(9)                                                              'jv022316
                Else                                                                            'jv022316
                    If skurec(Val(ds!sku)).pallet <> 0 And ds!onhand > 0 Then
                        's = s & Format((Int(ds!onhand / skurec(Val(ds!sku)).pallet)) + 0.999, "0") & Chr(9)
                        s = s & pallet_space(ds!sku, ds!onhand) & Chr(9)        'jv022516
                    Else
                        s = s & Chr(9)
                    End If
                End If                                                                          'jv022316
                s = s & ds!onhand
                Grid1.AddItem s
                psku = ds!sku
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    'If t3gal > 0 Then tp = Int(t3gal / 60) + 0.999                                              'jv022316
    If t3gal > 0 Then tp = pallet_space("507", t3gal)                           'jv022516
    'MsgBox tp, vbOKOnly, t3gal
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            tp = tp + Val(Grid1.TextMatrix(i, 2))
            tu = tu + Val(Grid1.TextMatrix(i, 3))
            If Val(Grid1.TextMatrix(i, 3)) < 0 Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ycolor.BackColor
                tn = tn + 1
            End If
        Next i
        Grid1.Row = 1
    End If
    s = "Total Units: " & tu '& vbCrLf
    '----- Removed for Pallet turnover grid         'jv112916
    's = s & " Total Racks in Use: " & tp & vbCrLf                                  'jv022516
    's = s & "Usable Capacity: "
    'If wcap <> 0 Then
    '    s = s & wcap '& vbCrLf
    'Else
    '    s = s & "Undefined" '& vbCrLf
    'End If
    'If wcap > 0 And tp > 0 Then
    '    s = s & "  Pct. " & Format(tp / wcap, "0.000")
    'End If
    Text1 = s
    'Grid1.FormatString = "^SKU|<Product|^Pallets|^Units"
    Grid1.FormatString = "^SKU|<Product|^Racks|^Units"                      'jv022516
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 4500
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.Redraw = True
    If whsno <> "001" And whsno <> "047" And whsno <> "052" Then
        Call branch_turnover(whsno)                                             'jv112916
    End If
    Screen.MousePointer = 0
End Sub

Private Sub sales_vs_inventory()
    Dim ds As ADODB.Recordset, s As String, i As Integer
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 13: Grid1.FixedCols = 2
    s = "select * from bimp where branchwhs = '" & whsno & "' and plantwhs <> 'DRY'"
    If allskus.Checked = True Then s = s & " order by sku, plantwhs"
    If wskus.Checked Then s = s & " and bimpstatus = 'W' order by sku, plantwhs"
    If yskus.Checked Then s = s & " and bimpstatus = 'Y' order by sku, plantwhs"
    If bskus.Checked Then s = s & " and bimpstatus = 'B' order by sku, plantwhs"
    If gskus.Checked Then s = s & " and bimpstatus = 'G' order by sku, plantwhs"
    If promos.Checked Then s = s & " and promoflag = 'Y' order by sku, plantwhs"
    If discprod.Checked Then s = s & " and discflag = 'Y' order by sku, plantwhs"
    If nosales.Checked Then s = s & " and sales = 0 order by sku, plantwhs"
    If ordskus.Checked Then s = s & " and onorder > 0 order by sku, plantwhs"
    If ushg.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '1/2') order by sales desc"
    If us48.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '48 OZ') order by sales desc"
    If uspt.Checked Then s = s & " and sku in (select sku from skumast where fgunit = 'PT') order by sales desc"
    If usqt.Checked Then s = s & " and sku in (select sku from skumast where fgunit = 'QT') order by sales desc"
    If us3g.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '3GAL') order by sales desc"
    If us6p.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '6PK') order by sales desc"
    If us12p.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '12PK') order by sales desc"
    If us24p.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '24PK') order by sales desc"
    If uscup.Checked Then s = s & " and sku in (select sku from skumast where fgunit = 'CUP') order by sales desc"
    If usbulk.Checked Then s = s & " and sku in (select sku from skumast where fgunit = 'BULK') order by sales desc"
    's = s & " Order by sku, plantwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            If ds!lastrecpt > "1" Then
                s = s & DateDiff("d", ds!lastrecpt, Now)
            End If
            s = s & Chr(9)
            If ds!roqty > 0 Then s = s & CInt(ds!onhand / ds!roqty)
            s = s & Chr(9)
            s = s & ds!onhand & Chr(9)
            s = s & ds!onorder & Chr(9)
            s = s & ds!sales & Chr(9)
            s = s & ds!undiff & Chr(9)
            s = s & ds!paldiff & Chr(9)
            s = s & ds!plantwhs & Chr(9)
            s = s & " " & Chr(9)
            s = s & ds!bimpstatus
            s = s & Chr(9) & Format(ds!ohpct * 30, "#")         'jv012916
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            Grid1.Row = i: Grid1.RowSel = 1
            Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
            If Grid1.TextMatrix(i, 11) = "W" Then Grid1.CellBackColor = wcolor.BackColor
            If Grid1.TextMatrix(i, 11) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
            If Grid1.TextMatrix(i, 11) = "B" Then Grid1.CellBackColor = bcolor.BackColor
            If Grid1.TextMatrix(i, 11) = "G" Then Grid1.CellBackColor = gcolor.BackColor
        Next i
        Grid1.Row = 1
        Grid1.RowHeight(0) = Grid1.RowHeight(1) * 2
    End If
    Text1 = "Sales vs Inventory"
    Text1.Visible = False
    s = "^SKU|<Product|^Days in Stock|^Pallets OnHand|^Units OnHand|^Units OnOrder|^Sales Last 30 Days"
    s = s & "|^Units Diff|^Pallet Diff|^Source|||^Days Supply"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 0 '1000
    Grid1.ColWidth(11) = 0
    Grid1.ColWidth(12) = 1000
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub sales_vs_inventory2()
    Dim ds As ADODB.Recordset, s As String, i As Integer, psku As String, sortflag As Boolean
    psku = " ": sortflag = False
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 13: Grid1.FixedCols = 2
    s = "select * from bimp where branchwhs = '" & whsno & "' and plantwhs <> 'DRY'"
    'If allskus.Checked = True Then s = s & " order by sku, onorder desc"
    If wskus.Checked Then s = s & " and bimpstatus = 'W'"
    If yskus.Checked Then s = s & " and bimpstatus = 'Y'"
    If bskus.Checked Then s = s & " and bimpstatus = 'B'"
    If gskus.Checked Then s = s & " and bimpstatus = 'G'"
    If promos.Checked Then s = s & " and promoflag = 'Y'"
    If discprod.Checked Then s = s & " and discflag = 'Y'"
    If nosales.Checked Then s = s & " and sales = 0"
    If ordskus.Checked Then s = s & " and onorder > 0"
    If ushg.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '1/2')"
    If us48.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '48 OZ')"
    If uspt.Checked Then s = s & " and sku in (select sku from skumast where fgunit = 'PT')"
    If usqt.Checked Then s = s & " and sku in (select sku from skumast where fgunit = 'QT')"
    If us3g.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '3GAL')"
    If us6p.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '6PK')"
    If us12p.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '12PK')"
    If us24p.Checked Then s = s & " and sku in (select sku from skumast where fgunit = '24PK')"
    If uscup.Checked Then s = s & " and sku in (select sku from skumast where fgunit = 'CUP')"
    If usbulk.Checked Then s = s & " and sku in (select sku from skumast where fgunit = 'BULK')"
    s = s & " Order by sku, onorder desc, plantwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!sku <> psku Then
                s = ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                If ds!lastrecpt > "1" Then
                    s = s & DateDiff("d", ds!lastrecpt, Now)
                End If
                s = s & Chr(9)
                If ds!roqty > 0 Then s = s & CInt(ds!onhand / ds!roqty)
                s = s & Chr(9)
                s = s & ds!onhand & Chr(9)
                s = s & ds!onorder & Chr(9)
                s = s & ds!sales & Chr(9)
                s = s & ds!undiff & Chr(9)
                s = s & ds!paldiff & Chr(9)
                s = s & ds!plantwhs & Chr(9)
                s = s & " " & Chr(9)
                s = s & ds!bimpstatus
                s = s & Chr(9) & Format(ds!ohpct * 30, "#")         'jv012916
                Grid1.AddItem s
                psku = ds!sku
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    If ushg.Checked Then sortflag = True
    If us48.Checked Then sortflag = True
    If uspt.Checked Then sortflag = True
    If usqt.Checked Then sortflag = True
    If us3g.Checked Then sortflag = True
    If us6p.Checked Then sortflag = True
    If us12p.Checked Then sortflag = True
    If us24p.Checked Then sortflag = True
    If uscup.Checked Then sortflag = True
    If usbulk.Checked Then sortflag = True
    If sortflag = True Then
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 6: Grid1.ColSel = 6
        Grid1.Sort = 4
    Else
        Grid1.RowSel = Grid1.Row                        'jv040716
        Grid1.Col = 1: Grid1.ColSel = 1                 'jv040716
        Grid1.Sort = 5                                  'jv040716
    End If
    
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            Grid1.Row = i: Grid1.RowSel = 1
            Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
            If Grid1.TextMatrix(i, 11) = "W" Then Grid1.CellBackColor = wcolor.BackColor
            If Grid1.TextMatrix(i, 11) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
            If Grid1.TextMatrix(i, 11) = "B" Then Grid1.CellBackColor = bcolor.BackColor
            If Grid1.TextMatrix(i, 11) = "G" Then Grid1.CellBackColor = gcolor.BackColor
        Next i
        Grid1.Row = 1
        Grid1.RowHeight(0) = Grid1.RowHeight(1) * 2
    End If
    Text1 = "Sales vs Inventory"
    Text1.Visible = False
    s = "^SKU|<Product|^Days in Stock|^Pallets OnHand|^Units OnHand|^Units OnOrder|^Sales Last 30 Days"
    s = s & "|^Units Diff|^Pallet Diff|^Source|||^Days Supply"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 0 '1000
    Grid1.ColWidth(11) = 0
    Grid1.ColWidth(12) = 1000
    Grid1.Redraw = True
    Screen.MousePointer = 0
    If whsno <> "001" And whsno <> "047" And whsno <> "052" Then
        Call branch_turnover(whsno)                                             'jv112916
    End If
End Sub

Private Sub allskus_Click()
    clear_list_checks
    allskus.Checked = True
    qstr_Change
End Sub

Private Sub bskus_Click()
    clear_list_checks
    bskus.Checked = True
    qstr_Change
End Sub

Private Sub Command1_Click()
    Dim hfile As String, rt As String, rh As String, rf As String, msg As String
    Grid1.Redraw = False
    hfile = localAppDataPath & "\htmltemp.htm"
    'msg = MsgBox(hfile, vbOKOnly)
    rt = Me.Caption
    rh = Text1
    'rf = "Printed: " & Format(Now, "M-d-yyyy h:mm am/pm")
    rf = "Last Update @ " & bimp_status_time                                            'jv022316
    Call htmlcolorgrid(Me, hfile, Grid1, rt, rh, rf, "lemonchiffon", "linen", "white")
    'msg = MsgBox(hfile, vbOKOnly)
    Form1.WebBrowser1.Navigate hfile
    Grid1.Redraw = True
    Unload Me
End Sub

Private Sub discprod_Click()
    clear_list_checks
    discprod.Checked = True
    qstr_Change
End Sub

Private Sub Form_Load()
    Me.Width = Form1.Width
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    Text1.Width = Grid1.Width
    If Me.Height > 2000 Then
        If Text1.Visible = True Then
            Grid1.Height = Me.Height - (Text1.Height * 1.5)
        Else
            Grid1.Height = Me.Height - (Text1.Height * 1.7)
        End If
    End If
    Command1.Left = Me.Width - (Command1.Width + 300)
    hgrid.Left = Me.Width - (hgrid.Width + 500)                     'jv112916
    'Text2.Left = Command1.Left                                      'jv022316
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.saleinv.Checked = False
End Sub

Private Sub gskus_Click()
    clear_list_checks
    gskus.Checked = True
    qstr_Change
End Sub

Private Sub nosales_Click()
    clear_list_checks
    nosales.Checked = True
    qstr_Change
End Sub

Private Sub ordskus_Click()
    clear_list_checks
    ordskus.Checked = True
    qstr_Change
End Sub

Private Sub promos_Click()
    clear_list_checks
    promos.Checked = True
    qstr_Change
End Sub

Private Sub qstr_Change()
    hgrid.Visible = False                                               'jv112916
    Text2.BackColor = Text1.BackColor
    Text2 = "Updated: " & bimp_status_time                              'jv022316
    listmenu.Visible = False
    Text1.Visible = True
    If qstr = "tstatrpt" Then trailer_status_rpt
    If qstr = "tktrpt" Then ticket_barcodes                             'jv020817
    If qstr = "broos" Then                                              'jv032318
        Text2 = "Updated: " & Format(FileDateTime("S:\wd\html\boutstock.csv"), "M-d-yyyy h:mm am/pm")  'jv032318
        branch_out_of_stock                                             'jv032318
    End If                                                              'jv032318
    If qstr = "gemminv" Then oracle_onhand_rpt
    If qstr = "rco50" Then Call reco_rpt("T10")
    If qstr = "rco51" Then Call reco_rpt("K10")
    If qstr = "rco52" Then Call reco_rpt("A10")
    If qstr = "salevinv" Then
        listmenu.Visible = True
        Text1.Visible = False
        'sales_vs_inventory
        sales_vs_inventory2
        Text2.BackColor = Me.BackColor
    End If
    If qstr = "tstation" Then tstation_rpt                              'jv072718
    DoEvents
End Sub

Private Sub us12p_Click()
    clear_list_checks
    us12p.Checked = True
    qstr_Change
End Sub

Private Sub us24p_Click()
    clear_list_checks
    us24p.Checked = True
    qstr_Change
End Sub

Private Sub us3g_Click()
    clear_list_checks
    us3g.Checked = True
    qstr_Change
End Sub

Private Sub us48_Click()
    clear_list_checks
    us48.Checked = True
    qstr_Change
End Sub

Private Sub us6p_Click()
    clear_list_checks
    us6p.Checked = True
    qstr_Change
End Sub

Private Sub usbulk_Click()
    clear_list_checks
    usbulk.Checked = True
    qstr_Change
End Sub

Private Sub uscup_Click()
    clear_list_checks
    uscup.Checked = True
    qstr_Change
End Sub

Private Sub ushg_Click()
    clear_list_checks
    ushg.Checked = True
    qstr_Change
End Sub

Private Sub uspt_Click()
    clear_list_checks
    uspt.Checked = True
    qstr_Change
End Sub

Private Sub usqt_Click()
    clear_list_checks
    usqt.Checked = True
    qstr_Change
End Sub

Private Sub usthome_Click()
    clear_list_checks
    usthome.Checked = True
    qstr_Change
End Sub

Private Sub whsno_Change()
    qstr_Change
End Sub

Private Sub wskus_Click()
    clear_list_checks
    wskus.Checked = True
    qstr_Change
End Sub

Private Sub yskus_Click()
    clear_list_checks
    yskus.Checked = True
    qstr_Change
End Sub
