VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   6030
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11940
   ForeColor       =   &H00C00000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form10"
   ScaleHeight     =   6030
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   450
      Left            =   2280
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   1095
      Left            =   0
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1931
      _Version        =   327680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Paste Oracle Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid gemmies 
      Height          =   1215
      Left            =   0
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2143
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid racks 
      Height          =   2895
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5106
      _Version        =   327680
      Cols            =   4
      BackColorFixed  =   12648447
      BackColorSel    =   16711680
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.ListBox sortlist 
      Height          =   1815
      Left            =   5040
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   2280
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2535
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
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label brcode 
      Caption         =   "Label2"
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
      Left            =   10920
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Countsheet:"
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
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
      Begin VB.Menu pcountsht 
         Caption         =   "Count Sheet"
      End
      Begin VB.Menu sprodtot 
         Caption         =   "Warehouse Product Totals"
      End
      Begin VB.Menu xitmenu 
         Caption         =   "Exit"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu instag 
         Caption         =   "Insert Tag - F10"
      End
      Begin VB.Menu delrec 
         Caption         =   "Clear Item - F9"
      End
      Begin VB.Menu deltag 
         Caption         =   "Delete Tag - Shift F9"
      End
      Begin VB.Menu delalltag 
         Caption         =   "Delete Tag - All"
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String, edrow As Integer
Dim savewhs As Boolean
Dim rcw(0 To 38) As Long
Private Sub print_racks(r1 As Integer, r2 As Integer, c1 As Integer, c2 As Integer)
    Dim i As Integer, k As Integer, j As Integer
    '  Print Check Off
    Printer.FontTransparent = True
    Printer.FillStyle = 0
    Printer.FillColor = QBColor(15)
    Printer.DrawMode = 1
    Printer.ForeColor = QBColor(0)
    
    Printer.FontName = "MS Serif"
    Printer.FontTransparent = True
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    Printer.Print "Count Sheet"
    Printer.Print Combo2.Text & " - " & brcode
    Printer.Print Format(Now, "mmmm d, yyyy")
    Dim xs As Long, xe As Long, xm As Long
    Dim ys As Long, ye As Long

    Printer.FontSize = 8
    xs = 0: xe = xs
    For i = c1 To c2
        xe = xe + racks.ColWidth(i)
    Next i
    Printer.Line (xs, 1200)-(xe, 1200)
    Printer.Line (xs, 1440)-(xe, 1440)
    Printer.FillColor = QBColor(15)
    Printer.DrawWidth = 3
    j = 0
    For i = r1 To r2 + 1
        ye = j * 240 + 1440
        Printer.Line (xs, ye)-(xe, ye)
        j = j + 1
    Next i
    Printer.DrawWidth = 1
    Printer.FontBold = False
    xm = xs + 100
    For k = c1 To c2
        Printer.PSet (xm, 1230)
        Printer.Print racks.TextMatrix(0, k)
        xm = xm + racks.ColWidth(k)
    Next k
    j = 1
    For i = r1 To r2
        xm = xs + 100
        For k = c1 To c2
            Printer.PSet (xm, j * 240 + 1230)
            Printer.Print racks.TextMatrix(i, k)
            xm = xm + racks.ColWidth(k)
        Next k
        j = j + 1
    Next i
    ys = 1200
    xm = xs
    Printer.DrawWidth = 6
    For i = c1 To c2
        Printer.Line (xm, ys)-(xm, ye)
        xm = xm + racks.ColWidth(i)
    Next i
    Printer.Line (xm, ys)-(xm, ye)
    Printer.EndDoc
End Sub
Private Sub print_pgrid(r1 As Integer, r2 As Integer, c1 As Integer, c2 As Integer)
    Dim i As Integer, k As Integer
    Printer.FontTransparent = True
    Printer.FillStyle = 0
    Printer.FillColor = QBColor(15)
    Printer.DrawMode = 1
    Printer.ForeColor = QBColor(0)
    
    Printer.Orientation = 2
    Printer.FontName = "MS Serif"
    Printer.FontTransparent = True
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    Printer.Print "Product Totals - " & Combo2
    Printer.Print "Branch " & brcode
    Printer.Print Format(Now, "mmmm d, yyyy")
    Dim xs As Long, xe As Long, xm As Long, gl As Integer
    Dim ys As Long, ye As Long, pg1 As Integer

    Printer.FontSize = 8
    xs = 0: xe = xs
    For i = c1 To c2
        xe = xe + pgrid.ColWidth(i)
    Next i
    Printer.Line (xs, 1200)-(xe, 1200)
    Printer.Line (xs, 1440)-(xe, 1440)
    Printer.FillColor = QBColor(15)
    Printer.DrawWidth = 3
    j = 0
    For i = r1 To r2
        ye = j * 240 + 1440
        k = j Mod 3
        Printer.Line (xs, ye)-(xe, ye)
        If k = 0 Then
            If Printer.FillColor = QBColor(14) Then
                Printer.FillColor = QBColor(15)
                Printer.Line (xs, ye)-(xe, ye + 720), Printer.FillColor, BF
            Else
                Printer.FillColor = QBColor(14)
                Printer.Line (xs, ye)-(xe, ye + 720), &HFFFF&, BF
            End If
        End If
        j = j + 1
    Next i
    Printer.DrawWidth = 1
    Printer.FontBold = False
    j = 0
    For i = r1 To r2
        xm = xs + 100
        For k = c1 To c2
            Printer.PSet (xm, j * 240 + 1230)
            Printer.Print pgrid.TextMatrix(i, k)
            xm = xm + pgrid.ColWidth(k)
        Next k
        j = j + 1
    Next i
    ys = 1200
    xm = xs
    Printer.DrawWidth = 6
    For i = c1 To c2
        If i Mod 2 = 0 Then Printer.Line (xm, ys)-(xm, ye)
        xm = xm + pgrid.ColWidth(i)
    Next i
    Printer.Line (xm, ys)-(xm, ye)
    Printer.EndDoc
End Sub

Private Sub refresh_tots()
    Dim i As Long, k As Integer, j As Long
    Dim pg As Integer, gc As Integer, gl As Integer
    Dim ssku As String, y As Integer, sy As Integer
    Dim stot As Long, ty As Integer, tx As Integer
    sort_item
    DoEvents
    pgrid.Clear: pgrid.Cols = 6: pgrid.Rows = 1
    pgrid.FormatString = "<|^|<|^|<|^"
    pgrid.ColWidth(0) = 4300: pgrid.ColWidth(1) = 700
    pgrid.ColWidth(2) = 4300: pgrid.ColWidth(3) = 700
    pgrid.ColWidth(4) = 4300: pgrid.ColWidth(5) = 700
    pg = 0: gc = 0: gl = 1: ssku = "s,swdjw"
    stot = 0: ty = 0: tx = 1: sy = 0
    For i = 0 To sortlist.ListCount - 1
        j = Val(Right(sortlist.List(i), 7))
        If ssku <> racks.TextMatrix(j, 2) Then
            If gl > 40 Then gc = 2
            If gl > 80 Then gc = 4
            If gl > 120 Then
                pg = pg + 1
                pgrid.AddItem "Page " & pg + 1
                pgrid.Row = pgrid.Rows - 1
                pgrid.RowSel = pgrid.Row
                pgrid.Col = 0
                pgrid.ColSel = pgrid.Cols - 1
                pgrid.FillStyle = flexFillRepeat
                pgrid.CellBackColor = pgrid.BackColorFixed
                gc = 0: gl = 1: sy = pgrid.Row
            End If
            pgrid.TextMatrix(ty, tx) = " " & Format(stot, "#")
            stot = 0
            ssku = racks.TextMatrix(j, 2)
            If gc = 0 Then
                pgrid.AddItem ssku & " " & racks.TextMatrix(j, 3)
                ty = pgrid.Rows - 1: tx = gc + 1
            Else
                y = sy + gl '(pg * 40) + gl
                If gl > 40 Then y = y - 40
                If gl > 80 Then y = y - 40
                pgrid.TextMatrix(y, gc) = ssku & " " & racks.TextMatrix(j, 3)
                ty = y: tx = gc + 1
            End If
            gl = gl + 1
        End If
        
        If gl > 40 Then gc = 2
        If gl > 80 Then gc = 4
        If gl > 120 Then
            pg = pg + 1
            pgrid.AddItem "Page " & pg + 1
            pgrid.Row = pgrid.Rows - 1
            pgrid.RowSel = pgrid.Row
            pgrid.Col = 0
            pgrid.ColSel = pgrid.Cols - 1
            pgrid.FillStyle = flexFillRepeat
            pgrid.CellBackColor = pgrid.BackColorFixed
            gc = 0: gl = 1: sy = pgrid.Row
        End If
        If gc = 0 Then
            pgrid.AddItem "        " & racks.TextMatrix(j, 1) & Chr(9) & " " & racks.TextMatrix(j, 7)
        Else
            y = sy + gl '(pg * 40) + gl
            If gl > 40 Then y = y - 40
            If gl > 80 Then y = y - 40
            pgrid.TextMatrix(y, gc) = "        " & racks.TextMatrix(j, 1)
            pgrid.TextMatrix(y, gc + 1) = " " & racks.TextMatrix(j, 7)
        End If
        stot = stot + Val(racks.TextMatrix(j, 7))
        gl = gl + 1
    Next i
    pgrid.TextMatrix(ty, tx) = " " & Format(stot, "#")
End Sub
Private Sub refresh_whs()
    Dim f0 As String, f1 As String, f2 As String
    Dim f3 As String, f4 As String, f5 As String
    Dim f6 As String, f7 As String, cfile As String
    racks.Redraw = False
    racks.FontName = "Arial"
    racks.FontBold = True
    racks.FontSize = 8
    racks.Clear: racks.Rows = 1: racks.Cols = 9
    racks.FormatString = "^|^Tag|^Item|<Contents|^Pallets|^Wraps|^Units|^Total|^LastDate"
    racks.ColWidth(0) = rcw(0)  '300
    racks.ColWidth(1) = rcw(1)  '1500
    racks.ColWidth(2) = rcw(2)  '2000
    racks.ColWidth(3) = rcw(3)  '3500
    racks.ColWidth(4) = rcw(4)  '700
    racks.ColWidth(5) = rcw(5)  '700
    racks.ColWidth(6) = rcw(6)  '700
    racks.ColWidth(7) = rcw(7)  '700
    racks.ColWidth(8) = rcw(8)  '900
    If Combo2 = "All Warehouses" Then
        For i = 0 To Combo2.ListCount - 1
            If Combo2.List(i) <> "All Warehouses" And Combo2.List(i) <> "Cycle Count Items" Then
                cfile = List2.List(i)
                If Len(Dir(cfile)) = 0 Then cfile = List1.List(i)
                If Len(Dir(cfile)) > 0 Then
                    Open cfile For Input As #1
                    Do Until EOF(1)
                        Input #1, f0, f1, f2, f3, f4, f5, f6, f7
                        'f0 = f0 & " " & Combo2.List(i)
                        sqlx = Chr(9) & f0 & Chr(9) & f1 & Chr(9)
                        sqlx = sqlx & f2 & Chr(9) & f3 & Chr(9)
                        sqlx = sqlx & f4 & Chr(9) & f5 & Chr(9)
                        sqlx = sqlx & f6 & Chr(9) & f7
                        racks.AddItem sqlx
                    Loop
                    Close #1
                End If
            End If
        Next i
    Else
        cfile = List2
        If Len(Dir(cfile)) = 0 Then cfile = List1
        If Len(Dir(cfile)) = 0 Then
            racks.AddItem "..."
            Exit Sub
        End If
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7
            sqlx = Chr(9) & f0 & Chr(9) & f1 & Chr(9)
            sqlx = sqlx & f2 & Chr(9) & f3 & Chr(9)
            sqlx = sqlx & f4 & Chr(9) & f5 & Chr(9)
            sqlx = sqlx & f6 & Chr(9) & f7
            racks.AddItem sqlx
        Loop
        Close #1
    End If
    racks.AddItem "..."
    racks.Redraw = True
End Sub
Private Sub sort_whse()
    Dim i As Long, sl As String
    sortlist.Clear
    If racks.Rows > 2 Then
        For i = 1 To racks.Rows - 2
            If Val(racks.TextMatrix(i, 1)) > 0 Then
                sl = Format(Val(racks.TextMatrix(i, 1)), "0000000")
            Else
                sl = racks.TextMatrix(i, 1)
            End If
            sl = Left(sl, 20)
            sl = sl & Space(20 - Len(sl))
            sl = sl & Format(i, "0000000")
            sortlist.AddItem sl
        Next i
    End If
End Sub
Private Sub sort_item()
    Dim i As Long, sl As String
    sortlist.Clear
    If racks.Rows > 2 Then
        For i = 1 To racks.Rows - 2
            If Val(racks.TextMatrix(i, 2)) > 0 Then
                sl = Format(Val(racks.TextMatrix(i, 2)), "0000000")
            Else
                sl = racks.TextMatrix(i, 2)
            End If
            sl = Left(sl, 20)
            sl = sl & Space(20 - Len(sl))
            sl = sl & Format(i, "0000000")
            sortlist.AddItem sl
        Next i
    End If
End Sub
Private Sub save_whs()
    Dim i As Long, k As Integer, j As Long
    'MsgBox "save_whs"
    'Exit Sub
    sort_whse
    DoEvents
    'Save on Local Server
    Open List2 For Output As #1
    For i = 0 To sortlist.ListCount - 1
        j = Val(Right(sortlist.List(i), 7))
        For k = 1 To racks.Cols - 2
            Write #1, racks.TextMatrix(j, k);
        Next k
        Write #1, racks.TextMatrix(j, racks.Cols - 1)
    Next i
    Close #1
    'Save on Web Server
    If List1 <> List2 Then
        Open List1 For Output As #1
        For i = 0 To sortlist.ListCount - 1
            j = Val(Right(sortlist.List(i), 7))
            For k = 1 To racks.Cols - 2
                Write #1, racks.TextMatrix(j, k);
            Next k
            Write #1, racks.TextMatrix(j, racks.Cols - 1)
        Next i
        Close #1
    End If
    savewhs = False
End Sub
Private Sub update_rec()
    Dim i As Integer, tu As Long
    Dim pdesc As String, pconv As Integer, wconv As Integer
    If edcell = "palqty" Then racks.TextMatrix(edrow, 4) = Format(Val(racks.TextMatrix(edrow, 4)), "#")
    If edcell = "wrpqty" Then racks.TextMatrix(edrow, 5) = Format(Val(racks.TextMatrix(edrow, 5)), "#")
    If edcell = "unqty" Then racks.TextMatrix(edrow, 6) = Format(Val(racks.TextMatrix(edrow, 6)), "#")
    If racks.TextMatrix(edrow, 2) <> gemmies.TextMatrix(gemmies.Row, 0) Then
        For i = 1 To gemmies.Rows - 1
            'If LCase(gemmies.TextMatrix(i, 0)) >= LCase(racks.TextMatrix(edrow, 2)) Then
            If LCase(gemmies.TextMatrix(i, 0)) = LCase(racks.TextMatrix(edrow, 2)) Then         'jv022416
                gemmies.Row = i
                Exit For
            End If
        Next i
    End If
    If LCase(racks.TextMatrix(edrow, 2)) = LCase(gemmies.TextMatrix(gemmies.Row, 0)) Then
        racks.TextMatrix(edrow, 2) = gemmies.TextMatrix(gemmies.Row, 0)
        pdesc = gemmies.TextMatrix(gemmies.Row, 1)
        pconv = Val(gemmies.TextMatrix(gemmies.Row, 2))
        wconv = Val(gemmies.TextMatrix(gemmies.Row, 3))
    Else
        pdesc = ""
        pconv = 0
        wconv = 0
    End If
    If edcell = "item" Then racks.TextMatrix(edrow, 3) = pdesc
    tu = pconv * Val(racks.TextMatrix(edrow, 4))
    tu = tu + (wconv * Val(racks.TextMatrix(edrow, 5)))
    tu = tu + Val(racks.TextMatrix(edrow, 6))
    racks.TextMatrix(edrow, 7) = Format(tu, "0")
    racks.TextMatrix(edrow, 8) = Format(Now, "m-d-yyyy")
    savewhs = True
    edcell = ""
End Sub
Private Sub refresh_gemmies()
    'Dim f0 As String, f1 As String
    'Dim f2 As String, f3 As String
    Dim i As Integer, pdesc As String, s As String
    ''gemmies.Visible = False
    gemmies.Clear: gemmies.Rows = 1
    gemmies.Cols = 4
    gemmies.FormatString = "<Item|<Description|^PalConv|^WrpConv"
    gemmies.ColWidth(0) = 2000
    gemmies.ColWidth(1) = 4000
    gemmies.ColWidth(2) = 600
    gemmies.ColWidth(3) = 600
    Combo1.Clear
    'f0 = Form1.webdir & "\counts\gemmies.txt"
    'Open f0 For Input As #1
    'Do Until EOF(1)
    '    Input #1, f0, f1, f2, f3
    '    sqlx = f0 & Chr(9) & StrConv(f1, vbProperCase) & Chr(9) & f2 & Chr(9)
    '    sqlx = sqlx & f3
    '    gemmies.AddItem sqlx
    '    If Val(f0) > 0 And Val(f0) < 999 Then
    '        Combo1.AddItem f0 & "   " & StrConv(f1, vbProperCase)
    '    'Else
    '    '    Combo1.AddItem f0
    '    End If
    'Loop
    'Close #1
    'gemmies.Visible = True
    For i = 1 To 9999
        If skurec(i).sku > " " Then
            pdesc = skurec(i).unit & " " & skurec(i).desc
            Combo1.AddItem skurec(i).sku & "   " & StrConv(pdesc, vbProperCase)
            s = skurec(i).sku & Chr(9) & pdesc
            s = s & Chr(9) & skurec(i).pallet & Chr(9) & skurec(i).wrapunits
            gemmies.AddItem s
        End If
    Next i
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End Sub

Private Sub brcode_Change()
    Combo2.Clear: List1.Clear: List2.Clear
    Combo2.AddItem "Cycle Count Items"
    List1.AddItem Form1.webdir & "\counts\cyclecnt." & brcode
    List2.AddItem Form1.locdir & "\cyclecnt." & brcode
    Combo2.AddItem "Order Pick"
    List1.AddItem Form1.webdir & "\counts\op." & brcode
    List2.AddItem Form1.locdir & "\op." & brcode
    Combo2.AddItem "Pallet Racks"
    List1.AddItem Form1.webdir & "\counts\racks." & brcode
    List2.AddItem Form1.locdir & "\racks." & brcode
    
    If Len(Dir(Form1.webdir & "\counts\tstation." & brcode)) > 0 Then       'jv061818
        Combo2.AddItem "Transfer Station"                                   'jv061818
        List1.AddItem Form1.webdir & "\counts\tstation." & brcode           'jv061818
        List2.AddItem Form1.locdir & "\tstation." & brcode                  'jv061818
    End If                                                                  'jv061818
    
    If brcode = "001" Or brcode = "047" Or brcode = "052" Then      'jv122817
        Combo2.AddItem "Routes"                                     'jv010716
        List1.AddItem Form1.webdir & "\counts\routes." & brcode     'jv010716
        List2.AddItem Form1.locdir & "\routes." & brcode            'jv010716
    End If                                                          'jv010716
    
    Combo2.AddItem "All Warehouses"
    List1.AddItem "allwhs.txt"
    List2.AddItem "allwhs.txt"
    Me.Caption = "Branch " & brcode & " Countsheets"
    Combo2.ListIndex = 0
End Sub

Private Sub Combo1_Click()
    If gemmies.Rows > Combo1.ListIndex Then
        gemmies.Row = Combo1.ListIndex + 1
    End If
End Sub

Private Sub Combo2_Click()
    Dim i As Integer
    If savewhs = True Then save_whs
    If racks.Rows > 2 Then
        For i = 0 To racks.Cols - 1
            rcw(i) = racks.ColWidth(i)
        Next i
    End If
    List2.ListIndex = Combo2.ListIndex
    List1.ListIndex = Combo2.ListIndex
    If Combo2 = "All Warehouses" Then
        edmenu.Visible = False
        Combo1.Visible = False
        Command1.Visible = False
    Else
        edmenu.Visible = True
        Combo1.Visible = True
        Command1.Visible = True
    End If
End Sub

Private Sub Command1_Click()
    edrow = racks.Row
    If Len(edcell) > 0 Then update_rec
    If racks.Row > 0 Then
        racks.TextMatrix(racks.Row, 2) = gemmies.TextMatrix(gemmies.Row, 0)
        edcell = "item"
        update_rec
    End If
End Sub

Private Sub delalltag_Click()
    Dim mok As String, i As Integer
    If edmenu.Visible = False Then Exit Sub
    If racks.TextMatrix(racks.Row, 0) > " " Then Exit Sub
    mok = "Ok to remove all lines with tag " & racks.TextMatrix(racks.Row, 1)
    mok = Trim(mok) & "?"
    If MsgBox(mok, vbYesNo + vbQuestion, "remove tag lines....") = vbNo Then Exit Sub
    mok = racks.TextMatrix(racks.Row, 1)
    For i = racks.Rows - 1 To 1 Step -1
        If racks.TextMatrix(i, 1) = mok And racks.TextMatrix(i, 0) < ".." Then
            racks.RemoveItem i
        End If
    Next i
    savewhs = True
End Sub

Private Sub delrec_Click()
    Dim mok As String, i As Integer
    If edmenu.Visible = False Then Exit Sub
    mok = "Ok to clear item " & racks.TextMatrix(racks.Row, 2)
    mok = mok & " " & racks.TextMatrix(racks.Row, 3)
    mok = Trim(mok) & "?"
    If MsgBox(mok, vbYesNo + vbQuestion, "clear item....") = vbNo Then Exit Sub
    For i = 2 To racks.Cols - 1
        racks.TextMatrix(racks.Row, i) = ""
    Next i
    racks.TextMatrix(racks.Row, racks.Cols - 1) = Format(Now, "m-d-yyyy")
    savewhs = True
End Sub

Private Sub deltag_Click()
    Dim mok As String, i As Integer
    If edmenu.Visible = False Then Exit Sub
    If racks.TextMatrix(racks.Row, 0) > " " Then Exit Sub
    mok = "Ok to remove tag " & racks.TextMatrix(racks.Row, 1)
    mok = Trim(mok) & " line?"
    If MsgBox(mok, vbYesNo + vbQuestion, "remove tag line....") = vbNo Then Exit Sub
    racks.RemoveItem racks.Row
    savewhs = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 120 Then   'F9
        If Shift = 0 Then delrec_Click
        If Shift = 1 Then deltag_Click      'Shift - F9
        If Shift = 2 Then delalltag_Click   'Ctl-F9
    End If
    If KeyCode = 121 Then   'F10
        KeyCode = 0
        instag_Click
    End If
End Sub

Private Sub Form_Load()
    Dim f As String, i As Integer, k As Integer
    savewhs = False
    rcw(0) = 300: rcw(1) = 1500: rcw(2) = 1000 '2000
    rcw(3) = 3500: rcw(4) = 1200: rcw(5) = 1200
    rcw(6) = 1200: rcw(7) = 1200: rcw(8) = 1200
    'If Len(Dir("c:\gswidth.ini")) > 0 Then
    '    If FileLen("c:\gswidth.ini") > 0 Then
    '        Open "c:\gswidth.ini" For Input As #1
    '        Input #1, f: Me.Height = Val(f)
    '        Input #1, f: Me.Width = Val(f)
    '        Input #1, f: Me.Top = Val(f)
    '        i = 0
    '        Do Until EOF(1)
    '        'For i = 0 To 8
    '            Input #1, f
    '            rcw(i) = Val(f)
    '            i = i + 1
    '        'Next i
    '        Loop
    '        Close #1
    '    End If
    'End If
    refresh_gemmies
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
    Me.Width = Form1.wdbanner.Width
End Sub

Private Sub Form_Resize()
    'gemmies.Width = Me.Width - 80
    racks.Width = Me.Width - 80
    If Me.Height > 3000 Then racks.Height = Me.Height - 1500 '30
End Sub

Private Sub Form_Terminate()
    Call xitmenu_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.csc.Checked = False
    Call xitmenu_Click
End Sub

Private Sub instag_Click()
    Dim ntag As String
    ntag = InputBox("Tag:", "Add tag to list", racks.TextMatrix(racks.Row, 1))
    If Len(ntag) = 0 Then Exit Sub
    If racks.Row > 0 Then
        racks.AddItem "" & Chr(9) & ntag, racks.Row
    Else
        i = racks.Rows - 1
        racks.TextMatrix(i, 0) = ""
        racks.TextMatrix(i, 1) = ntag
        racks.AddItem "..."
    End If
    savewhs = True
End Sub

Private Sub List1_Click()
    refresh_whs
End Sub

Private Sub pcountsht_Click()
    Dim i As Integer, k As Integer, j As Integer
    sort_whse
    DoEvents
    Screen.MousePointer = 11
    k = 1: j = 1
    For i = 1 To racks.Rows - 1
        If j > 55 Then
            Call print_racks(k, i, 1, 7)
            k = i + 1
            j = 0
        End If
        j = j + 1
    Next i
    If k < racks.Rows - 1 Then
        Call print_racks(k, racks.Rows - 1, 1, 7)
    End If
    Screen.MousePointer = 0
End Sub

Private Sub racks_KeyPress(KeyAscii As Integer)
    If Combo2 = "All Warehouses" Then
        MsgBox "No editing allowed while viewing 'All Warehouses'.", vbOKOnly + vbInformation, "Read-only data..."
        Exit Sub
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        If racks.Col < racks.Cols - 1 Then
            racks.Col = racks.Col + 1
        Else
            racks.Col = 1
        End If
        Exit Sub
    End If
    If racks.Col = 0 Then Exit Sub
    If racks.Col > 6 Then Exit Sub
    If Len(edcell) = 0 Then racks.Text = ""
    If racks.Col = 1 Then edcell = "tag"
    If racks.Col = 2 Then edcell = "item"
    If racks.Col = 3 Then edcell = "contents"
    If racks.Col = 4 Then edcell = "palqty"
    If racks.Col = 5 Then edcell = "wrpqty"
    If racks.Col = 6 Then edcell = "unqty"
    'edcell = racks.TextMatrix(0, racks.Col)
    If KeyAscii = 8 Then
        If Len(racks.Text) <= 1 Then
            racks.Text = ""
        Else
            racks.Text = Left(racks.Text, Len(racks.Text) - 1)
        End If
        edrow = racks.Row
    End If
    If KeyAscii > 31 And KeyAscii < 127 Then
        racks.Text = racks.Text + Chr(KeyAscii)
        If racks.TextMatrix(racks.Row, 0) = "..." Then
            racks.TextMatrix(racks.Row, 0) = ""
            racks.AddItem "..."
        End If
        edrow = racks.Row
    End If
End Sub

Private Sub racks_RowColChange()
    If Len(edcell) > 0 Then
        update_rec
        DoEvents
    End If
End Sub

Private Sub sprodtot_Click()
    refresh_tots
    DoEvents
    'Form2.Caption = "Product Totals - " & Combo2
    'Form2.Show
    Dim i As Integer, k As Integer
    Screen.MousePointer = 11
    k = 1
    For i = 1 To pgrid.Rows - 1
        If Left(pgrid.TextMatrix(i, 0), 4) = "Page" Then
            Call print_pgrid(k, i - 1, 0, 5)
            k = i + 1
        End If
    Next i
    If k < pgrid.Rows - 1 Then
        Call print_pgrid(k, pgrid.Rows - 1, 0, 5)
    End If
    Screen.MousePointer = 0
    
End Sub

Private Sub xitmenu_Click()
    Dim i As Integer
    If savewhs = True Then save_whs
    DoEvents
    'Open "c:\gswidth.ini" For Output As #7
    'If Me.WindowState = 0 Then
    '    Write #7, Me.Height
    '    Write #7, Me.Width
    '    Write #7, Me.Top
    'Else
    '    Write #7, 4860
    '    Write #7, 7425
    '    Write #7, 105
    'End If
    'For i = 0 To racks.Cols - 1
    '    Write #7, racks.ColWidth(i)
    'Next i
    'Close #7
    'End
End Sub

