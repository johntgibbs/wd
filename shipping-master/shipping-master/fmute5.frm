VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Form3"
   ClientHeight    =   8610
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form3"
   ScaleHeight     =   8610
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   1695
      Left            =   0
      TabIndex        =   10
      Top             =   6240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   2990
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   1335
      Left            =   0
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2355
      _Version        =   327680
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3255
      LargeChange     =   1000
      Left            =   8280
      SmallChange     =   1000
      TabIndex        =   6
      Top             =   0
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   8775
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   8415
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   6000
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Printer"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2625
      ScaleWidth      =   3225
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label polno 
      Height          =   255
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub prtpage2(pd As Control)
    Dim dl As String
    Dim xs As Long, xe As Long, st As Long
    xs = 1440 * 0.25
    xe = 1440 * 8
    dl = "_________________________"
    'pd.Height = 1440 * 11
    'pd.Width = 1440 * 8.5
    pd.FontName = "Arial"
    pd.FontSize = 10
    If TypeOf pd Is Printer Then
        pd.DrawWidth = 6
    Else
        pd.DrawWidth = 1
    End If
    pd.Print " ": pd.Print " "
    pd.Print " ": pd.Print " "
    pd.Print " ": pd.Print " "
    st = pd.CurrentY
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 2.5: pd.Print "Driver #1";
    pd.CurrentX = 1440 * 4.5: pd.Print "Driver #2";
    pd.CurrentX = 1440 * 6.5: pd.Print "Driver #3"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Driver Name"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Starting Location"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Date"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Destination"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Depart temp."
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Print " "
    pd.CurrentX = 1440 * 0.5: pd.Print "Signature"
    pd.Print " "
    pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
    pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 2: pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 4: pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 6: pd.Line (xs, st)-(xs, pd.CurrentY)
    xs = 1440 * 8: pd.Line (xs, st)-(xs, pd.CurrentY)
    pd.Print " "
    
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Arrival Date";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "Arrival temperature";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Seal #";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "Verified by:";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Time Arrived";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "Time Departed";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "# Pallets returned";
    pd.CurrentX = 1440 * 2: pd.Print dl;
    pd.CurrentX = 1440 * 4.5: pd.Print "# Sleeves returned";
    pd.CurrentX = 1440 * 6: pd.Print dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Returns";
    pd.CurrentX = 1440 * 2: pd.Print dl & dl & dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Comments";
    pd.CurrentX = 1440 * 2: pd.Print dl & dl & dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Corrections";
    pd.CurrentX = 1440 * 2: pd.Print dl & dl & dl
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Form completed by";
    pd.CurrentX = 1440 * 2: pd.Print dl
    'If TypeOf pd Is Printer Then pd.EndDoc
End Sub
Private Sub prtpol_Click(pd As Control)
    Dim ds As adodb.Recordset, sqlx As String, s As String
    Dim js As adodb.Recordset, jobtrail As Boolean
    Dim j1 As String, j2 As String, j3 As String, j4 As String, j5 As String
    Dim ss As adodb.Recordset, lc As Integer
    Dim fcode As String, bcode As String, i As Integer
    Dim scode As String, stot As Currency, gtot As Currency
    Dim pno As Integer, tu As Long, tw As Integer, tp As Integer
    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long, p5 As Long
    Dim dbranch As String, daddr1 As String, daddr2 As String, dphone As String, dfax As String
    Dim oplant As String, oaddr1 As String, oaddr2 As String, ophone As String, ofax As String
    On Error GoTo vberror
    pno = 1: jobtrail = False
    Combo1.Clear: List1.Clear
    Combo1.AddItem "Page 1"
    List1.AddItem localAppDataPath & "\dec00001.bmp"
    oplant = Form1.plantno
    If oplant = "50" Then sqlx = "select * from branches where branch = 1"
    If oplant = "51" Then sqlx = "select * from branches where branch = 47"
    If oplant = "52" Then sqlx = "select * from branches where branch = 52"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        oaddr1 = ds!addr1
        oaddr2 = ds!addr2
        ophone = ds!brphone
        ofax = ds!brfax
    End If
    ds.Close
    sqlx = "select * from branches where branch = " & Edittrl.bno
    Set ds = Sdb.Execute(sqlx)
    ds.MoveFirst
    pd.Height = 1440 * 11
    pd.Width = 1440 * 8.5
    'Form5.Picture1.Cls
    'pd.FontName = "Times New Roman"
    pd.FontName = "Arial"
    pd.FontSize = 14
    pd.FontBold = True
    pd.Print Tab(32); " " '"B i l l   O f   L a d i n g"
    pd.FontSize = 10
    pd.FontBold = True
    pd.CurrentX = 720: pd.Print "Origination:";
    pd.FontBold = False
    pd.CurrentX = 1440 * 1.5: pd.Print "Blue Bell Creameries L.P.";
    pd.FontBold = True
    pd.CurrentX = 1440 * 4.5: pd.Print "Destination: ";
    pd.FontBold = False
    pd.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        jobtrail = True
        sqlx = "select * from jobbing where branch = " & Edittrl.bno
        sqlx = sqlx & " and account = '" & Edittrl.ano & "'"
        Set js = Sdb.Execute(sqlx)
        If js.BOF = False Then
            js.MoveFirst
            j1 = js!acctdesc & " "
            j2 = js!addr1 & " "
            j3 = js!addr2 & " "
            j4 = js!addr3 & " " & js!jzip
            j5 = js!jphone & " "
        Else
            j1 = " ": j2 = " ": j3 = " ": j4 = " ": j5 = " "
        End If
        js.Close
        If j2 <= " " Then
            j2 = j3: j3 = j4: j4 = j5
        End If
        If j3 <= " " Then
            j3 = j4: j4 = j5
        End If
        If j4 <= " " Then
            j4 = j5
        End If
        pd.Print "Jobbing Account # "; Edittrl.bno; "-"; Edittrl.ano; " "
    Else
        pd.Print Format(Edittrl.bno, "00"); " "; ds!branchname; " "; Right(Edittrl.Combo1, 2)
    End If
    
    pd.CurrentX = 1440 * 1.5: pd.Print oaddr1; '"1101 S. Blue Bell Road";
    pd.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        pd.Print j1
    Else
        pd.Print ds!addr1
    End If
    pd.CurrentX = 1440 * 1.5: pd.Print oaddr2; '"Brenham, Texas  77834-1807";
    pd.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        pd.Print j2
    Else
        pd.Print ds!addr2
    End If
    pd.CurrentX = 1440 * 1.5: pd.Print ophone; '"(979) 836-7977";
    pd.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        pd.Print j3
    Else
        pd.Print ds!brphone
    End If
    pd.CurrentX = 1440 * 1.5: pd.Print "Fax: " & ofax; '"Fax: (979) 830-7398";
    pd.CurrentX = 1440 * 5.5
    If ds!branch = 16 Or ds!branch = 15 Then
        pd.Print j4
    Else
        pd.Print "Fax: " & ds!brfax
    End If
    pd.Print String(130, "_")
    ds.Close
    tu = 0: tw = 0: tp = 0
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    sqlx = "select sku, sum(units) from trailers where runid = " & Left(Edittrl.List1, Len(Edittrl.List1) - 6)
    sqlx = sqlx & " group by sku having sum(units) > 0 order by sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = "select * from skumast where sku = '" & ds!sku & "'"
            Set ss = Sdb.Execute(sqlx)
            If ss.BOF = False Then
                ss.MoveFirst
                s = ds!sku & Chr(9)
                s = s & ss!fgunit & " " & ss!fgdesc & Chr(9)
                s = s & Format(ds(1) / ss!pallet, ".00") & Chr(9)
                s = s & Format(ds(1) / ss!numwrap, "0") & Chr(9)
                s = s & ds(1)
            End If
            ss.Close
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close: ss.Close
    
    If Grid1.Rows > 37 Then '46 Then
        p1 = 1440 * 2.75
        p2 = 1440 * 3.25
        p3 = 1440 * 4.75
        p4 = 1440 * 7.25
        p5 = 1440 * 7.75
        'pd.FontName = "MS Sans Serif"
        pd.FontName = "Arial"
        'pd.FontSize = 8
        pd.FontBold = True
        pd.CurrentX = 360: pd.Print "SKU";
        If jobtrail Then
            pd.CurrentX = p1 - pd.TextWidth("Wraps")
            pd.Print "Wraps";
        End If
        pd.CurrentX = p2 - pd.TextWidth("Units"): pd.Print "Units";
        pd.CurrentX = p3: pd.Print ("SKU");
        If jobtrail Then
            pd.CurrentX = p4 - pd.TextWidth("Wraps")
            pd.Print "Wraps";
        End If
        pd.CurrentX = p5 - pd.TextWidth("Units"): pd.Print "Units"
        pd.FontBold = False
        lc = 8
        pgrid.Clear: pgrid.Rows = Int(Grid1.Rows / 2) + 1: pgrid.Cols = 8
        For i = 1 To pgrid.Rows - 1
            k = i + pgrid.Rows - 1
            pgrid.TextMatrix(i, 0) = Grid1.TextMatrix(i, 0)
            pgrid.TextMatrix(i, 1) = Grid1.TextMatrix(i, 1)
            pgrid.TextMatrix(i, 2) = CInt(Val(Grid1.TextMatrix(i, 3)))
            pgrid.TextMatrix(i, 3) = Grid1.TextMatrix(i, 4)
            tu = tu + Val(Grid1.TextMatrix(i, 4))
            tw = tw + Val(Grid1.TextMatrix(i, 3))
            tp = tp + Val(Grid1.TextMatrix(i, 2))
            If k < Grid1.Rows Then
                pgrid.TextMatrix(i, 4) = Grid1.TextMatrix(k, 0)
                pgrid.TextMatrix(i, 5) = Grid1.TextMatrix(k, 1)
                pgrid.TextMatrix(i, 6) = CInt(Val(Grid1.TextMatrix(k, 3)))
                pgrid.TextMatrix(i, 7) = Grid1.TextMatrix(k, 4)
                tu = tu + Val(Grid1.TextMatrix(k, 4))
                tw = tw + Val(Grid1.TextMatrix(k, 3))
                tp = tp + Val(Grid1.TextMatrix(k, 2))
            End If
        Next i
        For i = 1 To pgrid.Rows - 1
            'pd.FontName = "MS Sans Serif"
            pd.FontName = "Arial"
            'pd.FontSize = 8
            pd.CurrentX = 360: pd.Print pgrid.TextMatrix(i, 0); " ";
            pd.Print StrConv(pgrid.TextMatrix(i, 1), vbProperCase); " ";
            If jobtrail Then
                pd.CurrentX = p1 - pd.TextWidth(pgrid.TextMatrix(i, 2))
                pd.Print pgrid.TextMatrix(i, 2);
            End If
            pd.CurrentX = p2 - pd.TextWidth(pgrid.TextMatrix(i, 3))
            pd.Print pgrid.TextMatrix(i, 3);
            pd.CurrentX = p3
            pd.Print pgrid.TextMatrix(i, 4); " ";
            pd.Print StrConv(pgrid.TextMatrix(i, 5), vbProperCase); " ";
            If jobtrail Then
                pd.CurrentX = p4 - pd.TextWidth(pgrid.TextMatrix(i, 6))
                pd.Print pgrid.TextMatrix(i, 6);
            End If
            pd.CurrentX = p5 - pd.TextWidth(pgrid.TextMatrix(i, 7))
            pd.Print pgrid.TextMatrix(i, 7)
        
            If lc > 54 Then
                If TypeOf pd Is Printer Then
                    pd.NewPage
                Else
                    rstr = localAppDataPath & "\dec" & Format(pno, "00000") & ".bmp"
                    SavePicture pd.Image, rstr
                    pd.Cls
                End If
                pno = pno + 1
                Combo1.AddItem "Page " & pno
                List1.AddItem localAppDataPath & "\dec" & Format(pno, "00000") & ".bmp"
                pd.Print "Page "; pno;
                pd.CurrentX = 8600: pd.Print "Policy Number ";
                pd.FontBold = True
                pd.FontUnderline = True
                'pd.Print ds!policyno
                pd.FontBold = False
                pd.FontUnderline = False
                pd.Print " "
                lc = 2: scode = " ": bcode = "N": fcode = "N"
            End If
            lc = lc + 1
        Next i
    Else
        p1 = 1440 * 2.25 '2.75
        p2 = 1440 * 5.25
        p3 = 1440 * 6.25 '5.75
        'pd.FontName = "MS Sans Serif"
        pd.FontName = "Arial"
        pd.FontSize = 12
        pd.FontBold = True
        pd.CurrentX = p1:  pd.Print "SKU";
        If jobtrail Then
            pd.CurrentX = p2 - pd.TextWidth("Wraps")
            pd.Print "Wraps";
        End If
        pd.CurrentX = p3 - pd.TextWidth("Units"): pd.Print "Units"
        pd.FontBold = False
        lc = 8
        pgrid.Clear: pgrid.Rows = Int(Grid1.Rows / 2) + 1: pgrid.Cols = 8
        For i = 1 To pgrid.Rows - 1
            k = i + pgrid.Rows - 1
            pgrid.TextMatrix(i, 0) = Grid1.TextMatrix(i, 0)
            pgrid.TextMatrix(i, 1) = Grid1.TextMatrix(i, 1)
            pgrid.TextMatrix(i, 2) = CInt(Val(Grid1.TextMatrix(i, 3)))
            pgrid.TextMatrix(i, 3) = Grid1.TextMatrix(i, 4)
            tu = tu + Val(Grid1.TextMatrix(i, 4))
            tw = tw + Val(Grid1.TextMatrix(i, 3))
            tp = tp + Val(Grid1.TextMatrix(i, 2))
            If k < Grid1.Rows Then
                pgrid.TextMatrix(i, 4) = Grid1.TextMatrix(k, 0)
                pgrid.TextMatrix(i, 5) = Grid1.TextMatrix(k, 1)
                pgrid.TextMatrix(i, 6) = CInt(Val(Grid1.TextMatrix(k, 3)))
                pgrid.TextMatrix(i, 7) = Grid1.TextMatrix(k, 4)
                tu = tu + Val(Grid1.TextMatrix(k, 4))
                tw = tw + Val(Grid1.TextMatrix(k, 3))
                tp = tp + Val(Grid1.TextMatrix(k, 2))
            End If
        Next i
        For i = 1 To Grid1.Rows - 1
            'pd.FontName = "MS Sans Serif"
            pd.FontName = "Arial"
            'pd.FontSize = 8
            pd.CurrentX = p1
            pd.Print Grid1.TextMatrix(i, 0); " ";
            pd.Print StrConv(Grid1.TextMatrix(i, 1), vbProperCase); " ";
            If jobtrail Then
                pd.CurrentX = p2 - pd.TextWidth(Grid1.TextMatrix(i, 3))
                pd.Print Grid1.TextMatrix(i, 3);
            End If
            pd.CurrentX = p3 - pd.TextWidth(Grid1.TextMatrix(i, 4))
            pd.Print Grid1.TextMatrix(i, 4)
        
            If lc > 54 Then
                If TypeOf pd Is Printer Then
                    pd.NewPage
                Else
                    rstr = localAppDataPath & "\dec" & Format(pno, "00000") & ".bmp"
                    SavePicture pd.Image, rstr
                    pd.Cls
                End If
                pno = pno + 1
                Combo1.AddItem "Page " & pno
                List1.AddItem localAppDataPath & "\dec" & Format(pno, "00000") & ".bmp"
                pd.Print "Page "; pno;
                pd.CurrentX = 8600: pd.Print "Policy Number ";
                pd.FontBold = True
                pd.FontUnderline = True
                'pd.Print ds!policyno
                pd.FontBold = False
                pd.FontUnderline = False
                pd.Print " "
                lc = 2: scode = " ": bcode = "N": fcode = "N"
            End If
            lc = lc + 1
        Next i
    End If
    pd.Print " "
    If Grid1.Rows > 37 Then '46 Then
        pd.CurrentX = p3: pd.Print "Total Units";
        If jobtrail Then
            pd.CurrentX = p4 - pd.TextWidth(Format(tw, "#,###,###"))
            pd.Print Format(tw, "#,###,###");
        End If
        pd.CurrentX = p5 - pd.TextWidth(Format(tu, "#,###,###")): pd.Print Format(tu, "#,###,###")
    Else
        pd.CurrentX = p1: pd.Print "Total Units";
        If jobtrail Then
            pd.CurrentX = p2 - pd.TextWidth(Format(tw, "#,###,###"))
            pd.Print Format(tw, "#,###,###");
        End If
        pd.CurrentX = p3 - pd.TextWidth(Format(tu, "#,###,###")): pd.Print Format(tu, "#,###,###")
    End If
    lc = lc + 2
    pd.FontName = "Arial"
    pd.FontSize = 10
    pd.FontBold = False
    pd.CurrentY = 1440 * 9
    'For i = lc To 50 '54 '45 '50 '57
    '    pd.Print " "
    'Next i
    pd.CurrentX = 720: pd.Print "Ship Date:";
    pd.CurrentX = 1440 * 1.5: pd.Print Edittrl.sd;
    pd.CurrentX = 1440 * 3: pd.Print "Trailer #";
    pd.CurrentX = 1440 * 4: pd.Print tc;
    pd.CurrentX = 1440 * 5: pd.Print "Total Pallets:";
    pd.CurrentX = 1440 * 6: pd.Print Int(tp + 0.8)
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Inspected By:";                                'jv082415
    pd.CurrentX = 1440 * 1.5: pd.Print "_____________________________";
    pd.CurrentX = 1440 * 4: pd.Print "Driver:";
    pd.CurrentX = 1440 * 5: pd.Print "_____________________________"
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Seal #";
    pd.CurrentX = 1440 * 1.5: pd.Print "_____________________________";
    pd.CurrentX = 1440 * 4: pd.Print "Sealed By:";
    pd.CurrentX = 1440 * 5: pd.Print "_____________________________"
    pd.Print " "
    'pd.Print "                                  "; UCase(Right(Form1.Caption, Len(Form1.Caption) - 8)); " TRAILER"
    pd.CurrentX = 720: pd.Print "Freight:";
    pd.CurrentX = 1440 * 1.5: pd.Print "__________________________________________________________________________"
    pd.Print " "
    pd.CurrentX = 720: pd.Print "Special Instructions:";
    pd.CurrentX = 1440 * 2: pd.Print "____________________________________________________________________"
    'If TypeOf pd Is Printer Then
    '    pd.EndDoc
    'Else
    '    rstr = "c:\dec" & Format(pno, "00000") & ".bmp"
    '    SavePicture pd.Image, rstr
    '    DoEvents
    '    Combo1.ListIndex = 0
    '    HScroll2.Max = pno
    'End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "prtpol_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " prtpol_click - Error Number: " & eno
        End
    End If
End Sub


Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
    Printer.Duplex = 3
    Printer.Orientation = 1
    Call prtpol_Click(Printer)
    Printer.NewPage
    Call prtpage2(Printer)
    Printer.EndDoc
    Printer.Duplex = 1
End Sub


Private Sub Form_Load()
    Picture1.Width = 1440 * 8.5
    Picture1.Height = 1440 * 11
    If Len(Dir(localAppDataPath & "\blnk8x11.bmp")) = 0 Then SavePicture Picture1.Image, localAppDataPath & "\blnk8x11.bmp"
    VScroll1.Max = 1140 * 11
End Sub

Private Sub Form_Resize()
    'Picture1.Width = Me.Width
    'Picture1.Height = Me.Height
    VScroll1.Height = Me.Height - 400
    'VScroll1.SmallChange = VScroll1.Height * 4
    VScroll1.Left = Me.Width - (VScroll1.Width + 100)
    If Me.Height > 2000 Then Frame1.Top = Me.Height - 900
    Frame1.Width = Me.Width
    HScroll1.Width = Me.Width - (VScroll1.Width + 100)
    HScroll1.Max = 12240 - HScroll1.Width
    HScroll1.LargeChange = Frame1.Width / 2
End Sub

Private Sub HScroll1_Change()
    Picture1.Move 0 - HScroll1.Value
End Sub

Private Sub List1_Click()
    Picture1.Picture = LoadPicture(localAppDataPath & "\blnk8x11.bmp")
    If Len(Dir(List1)) > 0 Then
        Picture1.Picture = LoadPicture(List1)
    End If
End Sub

Private Sub polno_Change()
    Me.Caption = "Policy " & polno & " Declarations"
    Picture1.CurrentY = 0
    Picture1.Picture = LoadPicture(localAppDataPath & "\blnk8x11.bmp")
    Call prtpol_Click(Picture1)
    'Call prtpage2(Picture1)
End Sub

Private Sub VScroll1_Change()
    Picture1.Move Picture1.Left, 0 - VScroll1.Value
End Sub
