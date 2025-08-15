Attribute VB_Name = "PrntGrid"
Global htdc(0 To 8) As String
Global gndc(0 To 8) As Long

Sub printstringwrap(pd As Control, ps As String, pl As Long, xs As Long)
    Dim Line2 As Boolean, os As String, sp As Integer
    Dim c As String, n As Integer
    If xs = 0 Then xs = 1440
    If pl = 0 Then pl = 8640
    Do
        Line2 = False
        pd.CurrentX = xs
        If pd.TextWidth(ps) > pl Then
            os = "": sp = Len(ps)
            For n = 1 To Len(ps)
                c = Mid(ps, n, 1)
                If c = " " Or c = "-" Or c = "/" Then sp = n
                os = os & c
                If pd.TextWidth(os) > pl Then
                    Line2 = True
                    pd.Print Left(ps, sp)
                    ps = Right(ps, Len(ps) - sp)
                    Exit For
                End If
            Next n
        Else
            pd.Print ps
        End If
    Loop While Line2 = True
End Sub

Sub printflexgrid(pd As Control, gn As Control, rt As String, rh As String, rf As String)
    'pd - output control
    'gn - grid control
    'rt - report title
    'rh - report header
    'rf - report footer
    Dim cw(0 To 128) As Long
    Dim i As Integer, k As Integer, j As Integer
    Dim xs As Long, xe As Long, xm As Long
    Dim ys As Long, ye As Long
    Dim header As Boolean, plim As Long
    Dim Line2 As Boolean, os As String
    Dim n As Integer, sp As Integer, c As String
    Dim gc(0 To 128) As String
    If gn.Rows < 2 Then Exit Sub
    If gn.Cols > 129 Then
        MsgBox "Too many columns in this grid.", vbOKOnly + vbInformation, "cannot print...."
        Exit Sub
    End If
    plim = 14400
    For i = 0 To gn.Cols - 1
        cw(i) = gn.ColWidth(i)
    Next i
    xs = 0: xe = xs
    For i = 0 To gn.Cols - 1
        If cw(i) > 100 Then xe = xe + cw(i)
    Next i
    If TypeOf pd Is Printer Then
        If xe > 12240 Then
            Printer.Orientation = 2
            plim = 10800
        Else
            Printer.Orientation = 1
            plim = 14400
        End If
    End If
    pd.FontName = gn.FontName
    pd.DrawWidth = 6
    header = True
    For i = 1 To gn.Rows - 1
        If header Then
            If TypeOf pd Is Printer And gn.Row > 1 Then pd.NewPage
            'pd.CurrentX = 0: pd.CurrentY = 0
            pd.FontSize = 14
            pd.Print " "
            pd.Print rt
            pd.Print rh
            pd.Print " "
            pd.FontSize = 8
            pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
            ys = pd.CurrentY
            xm = xs + 100
            pd.CurrentY = pd.CurrentY + 30
            For k = 0 To gn.Cols - 1
                If cw(k) > 100 Then
                    pd.CurrentX = xm
                    pd.Print gn.TextMatrix(0, k);
                    xm = xm + cw(k)
                End If
            Next k
            pd.Print ""
            pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
            header = False
        End If
        'Multi line
        For k = 0 To gn.Cols - 1
            If cw(k) > 100 Then gc(k) = gn.TextMatrix(i, k)
        Next k
        Do
            Line2 = False
            xm = xs + 100
            pd.CurrentY = pd.CurrentY + 30
            For k = 0 To gn.Cols - 1
                If cw(k) > 100 Then
                    pd.CurrentX = xm
                    xm = xm + cw(k)
                    If pd.TextWidth(gc(k)) > cw(k) - 100 Then
                        os = "": sp = Len(gc(k))
                        For n = 1 To Len(gc(k))
                            c = Mid(gc(k), n, 1)
                            If c = " " Or c = "/" Or c = "-" Then sp = n
                            os = os & c
                            If pd.TextWidth(os) > cw(k) - 100 Then
                                Line2 = True
                                pd.Print Left(gc(k), sp);
                                gc(k) = Right(gc(k), Len(gc(k)) - sp)
                                Exit For
                            End If
                        Next n
                    Else
                        pd.Print gc(k);
                        gc(k) = " "
                    End If
                End If
            Next k
            pd.Print " "
        Loop While Line2 = True
        pd.Line (xs, pd.CurrentY)-(xe, pd.CurrentY)
        If pd.CurrentY >= plim Then
            pd.Line (xs, ys)-(xs, pd.CurrentY)
            xm = xs
            For k = 0 To gn.Cols - 1
                If cw(k) > 100 Then
                    xm = xm + cw(k)
                    pd.Line (xm, ys)-(xm, pd.CurrentY)
                End If
            Next k
            pd.Print ""
            pd.Print rf
            header = True
        End If
    Next i
    
    pd.Line (xs, ys)-(xs, pd.CurrentY)
    xm = xs
    For k = 0 To gn.Cols - 1
        If cw(k) > 100 Then
            xm = xm + cw(k)
            pd.Line (xm, ys)-(xm, pd.CurrentY)
        End If
    Next k
    pd.Print ""
    pd.Print rf
    If TypeOf pd Is Printer Then pd.EndDoc
End Sub

Sub htmlgrid(fn As Form, hf As String, gn As Control, rt As String, rh As String, rf As String, pc As String, hc As String, dc As String)
    Dim i As Integer, k As Integer, tw As Long
    If Len(pc) = 0 Then pc = "linen"
    If Len(hc) = 0 Then hc = "lightgrey"
    If Len(dc) = 0 Then dc = "white"
    If Len(hf) = 0 Then hf = localAppDataPath & "\htmlgrid.htm"
    Open hf For Output As #1
    Print #1, "<html>"
    Print #1, "<head><title>"; rt; "</title></head>"
    Print #1, "<body bgcolor="; pc; ">"
    Print #1, "<font face=" & Chr(34) & gn.Font & Chr(34) & "SIZE=4>"; rt; "</font>"
    Print #1, "<BR>"
    Print #1, "<font face=" & Chr(34) & gn.Font & Chr(34) & "SIZE=2>" & rf & "</font>"
    Print #1, "<HR><font face=" & Chr(34) & gn.Font & Chr(34) & "SIZE=1>"
    Print #1, "<TABLE BORDER=" & Chr(34) & "1" & Chr(34) & " WIDTH=" & Chr(34) & Int((gn.Width / fn.Width) * 90); "%" & Chr(34);
    Print #1, " CELLPADDING=" & Chr(34) & "2" & Chr(34) & " CELLSPACING=" & Chr(34) & "1" & Chr(34) & ">"
    If Len(rh) > 0 Then
        'Print #1, "<CAPTION>"; rh; "</CAPTION>"
        k = 0
        For i = 0 To gn.Cols - 1
            If gn.ColWidth(i) > 10 Then k = k + 1
        Next i
        Print #1, "<TR><TH COLSPAN=" & k & " BGCOLOR = "; hc; "><font size=2>"; rh; "</TH></TR>"
    End If
    tw = 0
    For i = 0 To gn.Cols - 1
        If gn.ColWidth(i) > 10 Then tw = tw + gn.ColWidth(i)
    Next i
    For i = 0 To gn.Cols - 1
        If gn.ColWidth(i) > 10 Then
            Print #1, "<COLGROUP WIDTH="; Chr(34) & Int((gn.ColWidth(i) / tw) * 90) & "%" & Chr(34);
            If i < gn.FixedCols Then
                Print #1, " BGCOLOR="; hc;
            Else
                Print #1, " BGCOLOR="; dc;
            End If
            If gn.ColAlignment(i) = 1 Then Print #1, " ALIGN=" & Chr(34) & "LEFT" & Chr(34);
            If gn.ColAlignment(i) = 4 Then Print #1, " ALIGN=" & Chr(34) & "CENTER" & Chr(34);
            If gn.ColAlignment(i) = 7 Then Print #1, " ALIGN=" & Chr(34) & "RIGHT" & Chr(34);
            Print #1, ">"
        End If
    Next i
    Print #1, "<TR>"
    For i = 0 To gn.Cols - 1
        If gn.ColWidth(i) > 100 Then
            'Print #1, "<TH BGCOLOR=" & hc & " ALIGN=" & Chr(34) & "CENTER" & Chr(34) & "><font size=2>"; gn.TextMatrix(0, i); "</TH>"
            Print #1, "<TH BGCOLOR=" & hc & "><font size=2>"; gn.TextMatrix(0, i); "</TH>"
        End If
    Next i
    Print #1, "</TR>"
    For i = 1 To gn.Rows - 1
        Print #1, "<TR>";
        For k = 0 To gn.Cols - 1
            If gn.ColWidth(k) > 100 Then
                If Len(gn.TextMatrix(i, k)) > 0 Then
                    Print #1, "<TD><font size=1>"; gn.TextMatrix(i, k); "</TD>";
                Else
                    Print #1, "<TD><font size=1>.</TD>";
                End If
            End If
        Next k
        Print #1, "</TR>"
    Next i
    Print #1, "</TABLE>"
    Print #1, "</CENTER></font></body></html>"
    Close #1
End Sub

Sub htmlcolorgrid(fn As Form, hf As String, gn As Control, rt As String, rh As String, rf As String, pc As String, hc As String, dc As String)
    Dim i As Integer, k As Integer, tw As Long, j As Integer
    Dim cc As Long
    If Len(pc) = 0 Then pc = "linen"
    If Len(hc) = 0 Then hc = "lightgrey"
    If Len(dc) = 0 Then dc = "white"
    If Len(hf) = 0 Then hf = localAppDataPath & "\htmlgrid.htm"
    Open hf For Output As #1
    Print #1, "<html>"
    Print #1, "<head><title>"; rt; "</title></head>"
    Print #1, "<body bgcolor="; pc; ">"
    Print #1, "<font face=" & Chr(34) & gn.Font & Chr(34) & "SIZE=4>"; rt; "</font>"
    Print #1, "<BR>"
    Print #1, "<font face=" & Chr(34) & gn.Font & Chr(34) & "SIZE=2>" & rf & "</font>"
    Print #1, "<HR><font face=" & Chr(34) & gn.Font & Chr(34) & "SIZE=1>"
    Print #1, "<TABLE BORDER=" & Chr(34) & "1" & Chr(34) & " WIDTH=" & Chr(34) & Int((gn.Width / fn.Width) * 90); "%" & Chr(34);
    Print #1, " CELLPADDING=" & Chr(34) & "2" & Chr(34) & " CELLSPACING=" & Chr(34) & "1" & Chr(34) & ">"
    If Len(rh) > 0 Then
        'Print #1, "<CAPTION>"; rh; "</CAPTION>"
        k = 0
        For i = 0 To gn.Cols - 1
            If gn.ColWidth(i) > 10 Then k = k + 1
        Next i
        Print #1, "<TR><TH COLSPAN=" & k & " BGCOLOR = "; hc; "><font size=2>"; rh; "</TH></TR>"
    End If
    tw = 0
    For i = 0 To gn.Cols - 1
        If gn.ColWidth(i) > 10 Then tw = tw + gn.ColWidth(i)
    Next i
    For i = 0 To gn.Cols - 1
        If gn.ColWidth(i) > 10 Then
            Print #1, "<COLGROUP WIDTH="; Chr(34) & Int((gn.ColWidth(i) / tw) * 90) & "%" & Chr(34);
            If i < gn.FixedCols Then
                Print #1, " BGCOLOR="; hc;
            Else
                Print #1, " BGCOLOR="; dc;
            End If
            If gn.ColAlignment(i) = 1 Then Print #1, " ALIGN=" & Chr(34) & "LEFT" & Chr(34);
            If gn.ColAlignment(i) = 4 Then Print #1, " ALIGN=" & Chr(34) & "CENTER" & Chr(34);
            If gn.ColAlignment(i) = 7 Then Print #1, " ALIGN=" & Chr(34) & "RIGHT" & Chr(34);
            Print #1, ">"
        End If
    Next i
    Print #1, "<TR>"
    For i = 0 To gn.Cols - 1
        If gn.ColWidth(i) > 100 Then
            'Print #1, "<TH BGCOLOR=" & hc & " ALIGN=" & Chr(34) & "CENTER" & Chr(34) & "><font size=2>"; gn.TextMatrix(0, i); "</TH>"
            Print #1, "<TH BGCOLOR=" & hc & "><font size=2>"; gn.TextMatrix(0, i); "</TH>"
        End If
    Next i
    Print #1, "</TR>"
    For i = 1 To gn.Rows - 1
        Print #1, "<TR>";
        For k = 0 To gn.Cols - 1
            If gn.ColWidth(k) > 100 Then
                If Len(gn.TextMatrix(i, k)) > 0 Then
                    Print #1, "<TD";
                    gn.Row = i: gn.Col = k
                    cc = gn.CellBackColor
                    For j = 0 To 8
                        If gndc(j) = 0 Then Exit For
                        If cc = gndc(j) Then
                            Print #1, " BGCOLOR="; htdc(j);
                            Exit For
                        End If
                    Next j
                    Print #1, "><font size=1>"; gn.TextMatrix(i, k); "</TD>";
                Else
                    Print #1, "<TD><font size=1>.</TD>";
                End If
            End If
        Next k
        Print #1, "</TR>"
    Next i
    Print #1, "</TABLE>"
    Print #1, "</CENTER></font></body></html>"
    Close #1
End Sub
