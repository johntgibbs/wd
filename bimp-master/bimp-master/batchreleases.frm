VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form batchreleases 
   Caption         =   "Batch Releases"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15900
   LinkTopic       =   "Form1"
   ScaleHeight     =   9555
   ScaleWidth      =   15900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Add Batch to List"
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
      Left            =   14640
      TabIndex        =   12
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
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
      Left            =   12240
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   10920
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   8640
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8415
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   14843
      _Version        =   327680
      ForeColor       =   128
      BackColorFixed  =   12648447
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh Date(s)"
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
      Left            =   9840
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8280
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Production Dates:"
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
      Left            =   4320
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "thru"
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
      Left            =   7680
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "..."
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
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Plant:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "batchreleases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ship_history(grow As Integer)                  'jv012016
    Dim plit As String, mfile As String
    Dim spath As String, sdir As String, sqlx As String, fdate As String
    Dim sdate As String, edate As String, wsku As String, wlot As String
    Dim wzone As String, wstat As String, wgma As Integer, wside As String
    Dim waisle As String, wrack As String, hrow As Boolean, r12flag As Boolean, ocode As String
    Dim cfile As String, s As String, bc As String, srflag As Boolean
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim dl As Long, wbc As String, e As Long
    Dim k10path As String, a10path As String, t10path As String, opcode As String
    Dim syear As Integer, eyear As Integer, i As Integer
    Dim logpath As String
    'logpath = Form1.logdir
    logpath = "\\bbc-01-prodtrk\wd\pallogs\"
    k10path = "\\bbba-03-dc\f\user\waredist\data\pallogs\"
    a10path = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
    t10path = "\\bbc-01-prodtrk\wd\pallogs\"
    wbc = Grid1.TextMatrix(grow, 11) & "001"
    If Len(wbc) = 0 Then Exit Sub
    wsku = Trim(Left(wbc, 4))
    wlot = barcode_to_lotnum(wbc)
    opcode = Mid(wbc, 11, 3)
    s = wbc                                                             'jv012116
    sdate = Format(Val(Mid(s, 9, 2)) - 2, "00")                         'jv012116
    sdate = "20" & sdate & Mid(s, 5, 4)                                 'jv012116
    edate = Format(Now, "yyyymmdd")                                     'jv012116
    
    syear = Val(Left(sdate, 4))                                             'jv061215
    eyear = Val(Left(edate, 4))                                             'jv061215
    
    If Val(opcode) >= 200 And Val(opcode) <= 299 Then logpath = a10path
    If Val(opcode) >= 100 And Val(opcode) <= 199 Then logpath = k10path
    If Val(opcode) >= 300 And Val(opcode) <= 599 Then logpath = t10path
    
    On Error Resume Next
    
    spath = logpath & "bill*.txt"
    sdir = Dir$(spath)
    Do While sdir <> ""
        s = Right(sdir, 12)                                                 'jv061215
        mfile = logpath & "move" & s
        s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
        fdate = s                                                           'jv061215
        If fdate >= sdate And fdate <= edate Then
            
            Open logpath & sdir For Input Shared As #1
            Do Until EOF(1)
                Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                If f9 = wlot And bc < wsku Then bc = f6
                If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                    If f9 = wlot And bc < wsku Then bc = f6
                    s = Mid(f15, 3, 2) & "-"
                    s = s & Mid(f15, 5, 2) & "-20" & Mid(f15, 1, 2)
                    s = s & Mid(f15, 7, 8)
                    Grid1.TextMatrix(grow, 12) = Format(s, "M-dd-yyyy hh:mm am/pm")
                    Grid1.TextMatrix(grow, 13) = f3
                    Grid1.TextMatrix(grow, 14) = f4
                    Grid1.TextMatrix(grow, 15) = f6
                    Grid1.TextMatrix(grow, 0) = ""
                    Grid1.TextMatrix(grow, 5) = ""
                    Grid1.TextMatrix(grow, 11) = ""
                    Close #1
                    Exit Sub
                End If
            Loop
            Close #1
                       
            e = FileLen(mfile)
            If Err.Number = 0 Then
                Open mfile For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f9 = wlot And bc < wsku Then bc = f6
                    If f4 = "ORDER PICK" Or f4 = "M-OP" Then
                        If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                            If f9 = wlot And bc < wsku Then bc = f6
                            s = Mid(f15, 3, 2) & "-"
                            s = s & Mid(f15, 5, 2) & "-20" & Mid(f15, 1, 2)
                            s = s & Mid(f15, 7, 8)
                            Grid1.TextMatrix(grow, 12) = Format(s, "M-dd-yyyy hh:mm am/pm")
                            Grid1.TextMatrix(grow, 13) = f3
                            Grid1.TextMatrix(grow, 14) = f4
                            Grid1.TextMatrix(grow, 15) = f6
                            Grid1.TextMatrix(grow, 0) = ""
                            Grid1.TextMatrix(grow, 5) = ""
                            Grid1.TextMatrix(grow, 11) = ""
                            Close #1
                            Exit Sub
                        End If
                    End If
                Loop
                Close #1
            End If
            
            
        End If
        sdir = Dir$
        DoEvents
    Loop
    For i = syear To eyear                                                      'jv061215
        spath = logpath & Format(i, "0000") & "\" & "\bill*.txt"                'jv061215
        sdir = Dir$(spath)
        Do While sdir <> ""
            s = Right(sdir, 12)                                                 'jv061215
            s = Mid(s, 5, 4) & Mid(s, 1, 4)                                     'jv061215
            fdate = s                                                           'jv061215
            mfile = logpath & Format(i, "0000") & "\move" & Right(sdir, 12)
            If fdate >= sdate And fdate <= edate Then
                Open logpath & Format(i, "0000") & "\" & sdir For Input Shared As #1
                Do Until EOF(1)
                    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                    If f9 = wlot And bc < wsku Then bc = f6
                    If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                        If f9 = wlot And bc < wsku Then bc = f6
                        s = Mid(f15, 3, 2) & "-"
                        s = s & Mid(f15, 5, 2) & "-20" & Mid(f15, 1, 2)
                        s = s & Mid(f15, 7, 8)
                        Grid1.TextMatrix(grow, 12) = Format(s, "M-dd-yyyy hh:mm am/pm")
                        Grid1.TextMatrix(grow, 13) = f3
                        Grid1.TextMatrix(grow, 14) = f4
                        Grid1.TextMatrix(grow, 15) = f6
                        Grid1.TextMatrix(grow, 0) = ""
                        Grid1.TextMatrix(grow, 5) = ""
                        Grid1.TextMatrix(grow, 11) = ""
                        Close #1
                        Exit Sub
                    End If
                Loop
                Close #1
                
                e = FileLen(mfile)
                If Err.Number = 0 Then
                'If FileLen(mfile) > 0 Then
                    'MsgBox mfile
                    Open mfile For Input Shared As #1
                    Do Until EOF(1)
                        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
                        If f9 = wlot And bc < wsku Then bc = f6
                        If f4 = "ORDER PICK" Or f4 = "M-OP" Then
                            If Left(f6, 13) = Left(wbc, 13) Or (Trim(Left(f6, 4)) = wsku And (f11 = wlot & Mid(wbc, 11, 3))) Then       'jv080315
                                If f9 = wlot And bc < wsku Then bc = f6
                                s = Mid(f15, 3, 2) & "-"
                                s = s & Mid(f15, 5, 2) & "-20" & Mid(f15, 1, 2)
                                s = s & Mid(f15, 7, 8)
                                Grid1.TextMatrix(grow, 12) = Format(s, "M-dd-yyyy hh:mm am/pm")
                                Grid1.TextMatrix(grow, 13) = f3
                                Grid1.TextMatrix(grow, 14) = f4
                                Grid1.TextMatrix(grow, 15) = f6
                                Grid1.TextMatrix(grow, 0) = ""
                                Grid1.TextMatrix(grow, 5) = ""
                                Grid1.TextMatrix(grow, 11) = ""
                                Close #1
                                Exit Sub
                            End If
                        End If
                    Loop
                    Close #1
                End If
                
            End If
            sdir = Dir$
            DoEvents
        Loop
    Next i
End Sub



Private Sub refresh_grid(bno As String)
    Dim q As String, i As Integer
    Dim db As ADODB.Connection, ds As ADODB.Recordset, s As String, hs As ADODB.Recordset
    Dim t6 As Long, t7 As Long, t8 As Long                      'jv121415
    Dim t9 As Long, t10 As Long, t11 As Long, t13 As Long, t14 As Long, t15 As Long
    Dim k As Long, nl As Boolean, wdlot As String, pflag As String
    Dim psku As String, plot As String, sp As String, ep As String
    
    wdlot = Right(Text1, 2)
    wdlot = wdlot & Format(DateDiff("d", "1-1-" & Right(Text1, 4), Text1) + 1, "000")
    
    If plant_server_status(prodbatches.Combo1) = False Then                             'jv010417
        s = "Sorry, The server for Warehouse " & prodbatches.Combo1 & " has been flagged to be offline."
        MsgBox s, vbOKOnly + vbInformation, "sorry, try again later..."                 'jv010417
        Exit Sub                                                                        'jv010417
    End If                                                                              'jv010417
    
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    
    'On Error GoTo vberror
    Screen.MousePointer = 11
    If bno = "0" Then
        Grid1.Redraw = False
        Grid1.FontName = "Arial"
        Grid1.FontBold = True
        Grid1.FontSize = 8
        Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 16: Grid1.FixedCols = 2
    End If
    q = "select h.batch_id,h.batch_no,TO_CHAR(h.plan_start_date,'YYYY-MM-DD'),h.batch_status,"         'jv010516
    q = q & "h.attribute1,i.inventory_item_id,i.segment1,i.description,d.plan_qty,"
    q = q & "d.actual_qty"
    q = q & " from apps.gme_batch_header h, apps.gme_material_details d, apps.mtl_system_items_b i"
    If Val(List1) = 500 Then
        q = q & " where h.organization_id in (select organization_id from mtl_parameters where organization_code in ('500','503'))"
    Else
        q = q & " where h.organization_id in (select organization_id from mtl_parameters where organization_code in ('" & Format(Val(List1), "000") & "'))"
    End If
    If bno > "0" Then
        q = q & " and h.batch_no = " & bno
    Else
        q = q & " and h.plan_start_date >= TO_DATE('" & Format(Text1, "DD-MMM-YYYY") & "')"
        q = q & " and h.plan_start_date <= TO_DATE('" & Format(DateAdd("d", 1, Text2), "DD-MMM-YYYY") & "')"
    End If
    q = q & " and h.delete_mark = 0"
    q = q & " and h.batch_id = d.batch_id"
    q = q & " and h.batch_status in (1, 2, 3, 4)"
    q = q & " and d.line_type = 1"
    q = q & " and i.organization_id = d.organization_id"
    q = q & " and i.inventory_item_id = d.inventory_item_id"
    q = q & " and i.segment1 >= '100' and i.segment1 <= '9999'"             'jv082415
    'If sortsku.Checked Then
        q = q & " order by i.segment1, 2, d.plan_qty desc, h.attribute1"
    'Else
    '    q = q & " order by 2, i.segment1, d.plan_qty desc, h.attribute1"
    'End If
    'MsgBox q
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(0) & Chr(9)                              'Batch ID
            s = s & ds(1) & Chr(9)                          'Batch No
            s = s & Format(ds(2), "M-dd-yyyy") & Chr(9)     'Date
            If ds(3) = 1 Then s = s & "PEND" & Chr(9)       'Status
            If ds(3) = 2 Then s = s & "WIP" & Chr(9)
            If ds(3) = 3 Then s = s & "CERT" & Chr(9)
            If ds(3) = 4 Then s = s & "Closed" & Chr(9)
            s = s & ds(4) & Chr(9)                          'Location
            s = s & ds(5) & Chr(9)                          'Inventory Item Id
            s = s & ds(6) & Chr(9)                          'SKU
            s = s & ds(7) & Chr(9)                          'Product Name
            s = s & ds(8) & Chr(9)                          'Planned Qty
            s = s & ds(9) & Chr(9)                          'Actual Qty
            's = s & Format(ds(7) - ds(6), "0")              'Qty Diff
            pflag = Trim(ds(6))
            If Len(pflag) = 3 Then pflag = pflag & " "
            If Format(ds(2), "M-dd-yyyy") = "2-29-2016" Then                    'jv030416
                pflag = pflag & "022918"                                        'jv030416
            Else                                                                'jv030416
                pflag = pflag & Format(DateAdd("yyyy", 2, Format(ds(2), "M-dd-yyyy")), "MMddyy")
            End If
            pflag = pflag & Right(ds(4), 3)
            s = s & Chr(9) & pflag
            'MsgBox s
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 0) > "0" Then
                q = "select transaction_quantity, transaction_date, last_update_date, creation_date"
                q = q & " from mtl_material_transactions"
                q = q & " where transaction_source_id = " & Grid1.TextMatrix(i, 0)
                q = q & " and inventory_item_id = " & Grid1.TextMatrix(i, 5)
                q = q & " and transaction_quantity > " & skurec(Val(Grid1.TextMatrix(i, 6))).pallet
                q = q & " order by creation_date"
                'MsgBox q
                Set ds = r12db.Execute(q)
                If ds.BOF = False Then
                    ds.MoveFirst
                    Grid1.TextMatrix(i, 10) = Format(ds(3), "M-dd-yyyy hh:mm am/pm")
                End If
                ds.Close
                Call ship_history(i)
            End If
        Next i
    End If

    
    Screen.MousePointer = 0
    'Grid1.Redraw = True
    'Exit Sub
    
    'If Combo1 = "A10" Then
    '    Grid1.FormatString = "^Batch ID|^Batch No|^Plan Start|^Status|<Location|^Item|^SKU|<Description|^Planned|^Released|^Release Date|^Flag|^Ship Date|^Source|<Target|^Pallet"
    'Else
    '    Grid1.FormatString = "^Batch ID|^Batch No|^Plan Start|^Status|<Location|^Item|^SKU|<Description|^Planned|^Released|^Release Date|^BarCode|^Ship Date|^Source|<Target|^Pallet"
    'End If
    If Combo1 = "A10" Then
        Grid1.FormatString = "|^Batch No|^Plan Start|^Status|<Location||^SKU|<Description|^Planned|^Released|^Release Date||^Ship Date|^Source|<Target|^Pallet"
    Else
        Grid1.FormatString = "|^Batch No|^Plan Start|^Status|<Location||^SKU|<Description|^Planned|^Released|^Release Date||^Ship Date|^Source|<Target|^Pallet"
    End If
    
    Grid1.ColWidth(0) = 0 '900
    Grid1.ColWidth(1) = 900 '1100
    Grid1.ColWidth(2) = 1100 '800
    Grid1.ColWidth(3) = 800 '2000
    Grid1.ColWidth(4) = 2000 '700
    Grid1.ColWidth(5) = 0 '700 '2200
    Grid1.ColWidth(6) = 700 '900
    Grid1.ColWidth(7) = 2200 '900
    Grid1.ColWidth(8) = 900 '900
    Grid1.ColWidth(9) = 900 '800
    Grid1.ColWidth(10) = 1700 '800
    Grid1.ColWidth(11) = 0 '1300 '800
    Grid1.ColWidth(12) = 1700 '1400
    Grid1.ColWidth(13) = 1000 '900
    Grid1.ColWidth(14) = 1900
    Grid1.ColWidth(15) = 1700
    'Grid1.ColWidth(16) = 1100
    
    Grid1.FillStyle = flexFillRepeat
    'Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 7
    DoEvents
    Grid1.Row = 1: Grid1.Col = 1
    Grid1.Redraw = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.Description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_tickets", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_tickets - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_vlists()
    Combo1.Clear: List1.Clear: List2.Clear
    Combo1.AddItem "T10": List1.AddItem "500": List2.AddItem "Brenham"
    Combo1.AddItem "K10": List1.AddItem "501": List2.AddItem "Broken Arrow"
    Combo1.AddItem "A10": List1.AddItem "502": List2.AddItem "Sylacauga"
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    List2.ListIndex = Combo1.ListIndex
    Label2 = List2
    'Grid1_RowColChange
End Sub

Private Sub Command1_Click()
    Call refresh_grid("0")
End Sub

Private Sub Command2_Click()
    Dim rt As String, rh As String, rf As String
    rt = Me.Caption & " " & Combo1 & "-" & Label2.Caption
    rh = Text1 & " thru " & Text2
    rf = "printed:  " & Format(Now, "M-d-yyyy h:mm am/pm")
    Grid1.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "u:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\internet explorer\iexplore.exe u:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe u:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If

End Sub

Private Sub Command3_Click()
    Dim s As String
    s = InputBox("Batch #:", "Add batch to list....", " ")
    If Len(s) = 0 Then Exit Sub
    Call refresh_grid(s)
End Sub

Private Sub Form_Load()
    refresh_vlists
    Combo1.ListIndex = 0
    Text1 = Format(DateAdd("d", -14, Now), "M-d-yyyy")
    'Text2 = Format(Now, "M-d-yyyy")
    Text2 = Text1
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 16: Grid1.FixedCols = 2
    Grid1.Rows = 1
    If Combo1 = "A10" Then
        Grid1.FormatString = "|^Batch No|^Plan Start|^Status|<Location||^SKU|<Description|^Planned|^Released|^Release Date||^Ship Date|^Source|<Target|^Pallet"
    Else
        Grid1.FormatString = "|^Batch No|^Plan Start|^Status|<Location||^SKU|<Description|^Planned|^Released|^Release Date||^Ship Date|^Source|<Target|^Pallet"
    End If
    
    Grid1.ColWidth(0) = 0 '900
    Grid1.ColWidth(1) = 900 '1100
    Grid1.ColWidth(2) = 1100 '800
    Grid1.ColWidth(3) = 800 '2000
    Grid1.ColWidth(4) = 2000 '700
    Grid1.ColWidth(5) = 0 '700 '2200
    Grid1.ColWidth(6) = 700 '900
    Grid1.ColWidth(7) = 2200 '900
    Grid1.ColWidth(8) = 900 '900
    Grid1.ColWidth(9) = 900 '800
    Grid1.ColWidth(10) = 1700 '800
    Grid1.ColWidth(11) = 0 '1300 '800
    Grid1.ColWidth(12) = 1700 '1400
    Grid1.ColWidth(13) = 1000 '900
    Grid1.ColWidth(14) = 1900
    Grid1.ColWidth(15) = 1700
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 400
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 4)
End Sub
