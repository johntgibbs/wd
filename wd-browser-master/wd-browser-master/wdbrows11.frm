VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form11 
   Caption         =   "Oracle vs Countsheet"
   ClientHeight    =   7470
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11640
   LinkTopic       =   "Form11"
   ScaleHeight     =   7470
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Cycle Count Items"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.OptionButton Option1 
      Caption         =   "End of Period (All Warehouses)"
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
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3975
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7011
      _Version        =   327680
      Cols            =   4
      BackColorFixed  =   12648384
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label brcode 
      Caption         =   "..."
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   855
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
   End
   Begin VB.Menu postadj 
      Caption         =   "Post Adjustments"
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub refresh_cycle()
    Dim cfile As String, i As Integer
    Dim f0 As String, f1 As String, f2 As String
    Dim f3 As String, f4 As String, f5 As String
    Dim f6 As String, f7 As String, f8 As String
    Dim f9 As String, sqlx As String, fil As String
    Dim t0 As Single, t1 As Single, t2 As Single
    Dim ds As ADODB.Recordset
    t0 = 0: t1 = 0: t2 = 0
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    cfile = Form1.locdir & "\cyclecnt." & brcode
    If Len(Dir(cfile)) = 0 Then cfile = Form1.webdir & "\counts\cyclecnt." & brcode
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7
            sqlx = f1 & Chr(9) & f2 & Chr(9) & Chr(9) & f6
            If Val(f1) > 0 Then Grid1.AddItem sqlx
        Loop
        Close #1
    End If
    
    ''cfile = Form1.webdir & "\counts\gemmeop." & brcode
    'cfile = Form1.webdir & "\stock\goh." & Format(Val(brcode), "00")
    'If Len(Dir(cfile)) > 0 Then
    '    Open cfile For Input As #1
    '    'Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
    '    Line Input #1, fil
    '    f9 = Format(FileDateTime(cfile), "m-d-yyyy")
    '    sqlx = "^Item|<Description|^" & f9 & "|^Count|^Adj Qty"
    '    Grid1.FormatString = sqlx
    '    Do Until EOF(1)
    '        'Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
    '        Line Input #1, fil
    '        If Len(fil) > 50 Then
    '            'f1 = Left(fil, 3)
    '            f1 = Trim(Left(fil, 4))                             'jv012816
    '            f2 = Trim(Mid(fil, 5, 40))
    '            f9 = Val(Right(fil, 7))
    '            For i = 0 To Grid1.Rows - 1
    '                If Grid1.TextMatrix(i, 0) = f1 Then
    '                    Grid1.TextMatrix(i, 1) = f2
    '                    Grid1.TextMatrix(i, 2) = f9
    '                    Grid1.TextMatrix(i, 4) = Format(Val(Grid1.TextMatrix(i, 3)) - Val(Grid1.TextMatrix(i, 2)), "######.00")
    '                    Exit For
    '                End If
    '            Next i
    '        End If
    '    Loop
    '    Close #1
    'End If
    
    Grid1.FormatString = "^Item|<Description|^" & Format(Now, "M-dd-yyyy") & "|^Count|^Adj Qty"     'jv012916
    sqlx = "select sku, onhand from bimp where branchwhs = '" & Format(Val(brcode), "000") & "'"    'jv012916
    Set ds = wdb.Execute(sqlx)                                                                      'jv012916
    If ds.BOF = False Then                                                                          'jv012916
        ds.MoveFirst                                                                                'jv012916
        Do Until ds.EOF                                                                             'jv012916
            For i = 0 To Grid1.Rows - 1                                                             'jv012916
                If Grid1.TextMatrix(i, 0) = ds!sku Then                                             'jv012916
                    Grid1.TextMatrix(i, 2) = ds!onhand                                              'jv012916
                    Grid1.TextMatrix(i, 4) = Format(Val(Grid1.TextMatrix(i, 3)) - Val(Grid1.TextMatrix(i, 2)), "######.00")
                    Exit For                                                                        'jv012916
                End If                                                                              'jv012916
            Next i                                                                                  'jv012916
            ds.MoveNext                                                                             'jv012916
        Loop                                                                                        'jv012916
    End If                                                                                          'jv012916
    ds.Close                                                                                        'jv012916
                    
    
    Grid1.ColWidth(0) = 700
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            t0 = t0 + Val(Grid1.TextMatrix(i, 2))
            t1 = t1 + Val(Grid1.TextMatrix(i, 3))
            t2 = t2 + Val(Grid1.TextMatrix(i, 4))
        Next i
        sqlx = Chr(9) & "Totals" & Chr(9) & Format(t0, "#,###,###.00") & Chr(9)
        sqlx = sqlx & Format(t1, "#,###,###.00") & Chr(9)
        sqlx = sqlx & Format(t2, "#,###,###.00")
        Grid1.AddItem sqlx
    End If
    
End Sub
Private Sub refresh_grid()
    Dim cfile As String, i As Integer
    Dim f0 As String, f1 As String, f2 As String
    Dim f3 As String, f4 As String, f5 As String
    Dim f6 As String, f7 As String, f8 As String
    Dim f9 As String, sqlx As String
    Dim t0 As Single, t1 As Single, t2 As Single
    If Option2.Value = True Then
        refresh_cycle
        Exit Sub
    End If
    t0 = 0: t1 = 0: t2 = 0
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    cfile = Form1.webdir & "\counts\gemmeop." & brcode
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
        sqlx = "^" & f1 & "|<" & f2 & "|^" & f9 & "|^Count|^Adj Qty"
        Grid1.FormatString = sqlx
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
            If Val(f0) > 0 Then
                sqlx = f1 & Chr(9) & f2 & Chr(9) & f9
                Grid1.AddItem sqlx
            End If
        Loop
        Close #1
    End If
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 1200
    DoEvents
    cfile = Form1.locdir & "\op." & brcode
    If Len(Dir(cfile)) = 0 Then cfile = Form1.webdir & "\counts\op." & brcode
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7
            'If Val(f6) > 0 Then
                For i = 1 To Grid1.Rows - 1
                    If f1 = Grid1.TextMatrix(i, 0) Then
                        Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + Val(f6)
                        Exit For
                    End If
                Next i
            'End If
        Loop
        Close #1
        DoEvents
    End If
    cfile = Form1.locdir & "\racks." & brcode
    If Len(Dir(cfile)) = 0 Then cfile = Form1.webdir & "\counts\racks." & brcode
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7
            'If Val(f6) > 0 Then
                For i = 1 To Grid1.Rows - 1
                    If f1 = Grid1.TextMatrix(i, 0) Then
                        Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + Val(f6)
                        Exit For
                    End If
                Next i
            'End If
        Loop
        Close #1
        DoEvents
    End If
    cfile = Form1.locdir & "\routes." & brcode                                          'jv010716
    If Len(Dir(cfile)) = 0 Then cfile = Form1.webdir & "\counts\routes." & brcode       'jv010716
    If Len(Dir(cfile)) > 0 Then                                                         'jv010716
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7
            'If Val(f6) > 0 Then
                For i = 1 To Grid1.Rows - 1
                    If f1 = Grid1.TextMatrix(i, 0) Then
                        Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + Val(f6)
                        Exit For
                    End If
                Next i
            'End If
        Loop
        Close #1
        DoEvents
    End If
    
    
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            Grid1.TextMatrix(i, 4) = Format(Val(Grid1.TextMatrix(i, 3)) - Val(Grid1.TextMatrix(i, 2)), "######.00")
            t0 = t0 + Val(Grid1.TextMatrix(i, 2))
            t1 = t1 + Val(Grid1.TextMatrix(i, 3))
            t2 = t2 + Val(Grid1.TextMatrix(i, 4))
        Next i
        sqlx = Chr(9) & "Totals" & Chr(9) & Format(t0, "#,###,###.00") & Chr(9)
        sqlx = sqlx & Format(t1, "#,###,###.00") & Chr(9)
        sqlx = sqlx & Format(t2, "#,###,###.00")
        Grid1.AddItem sqlx
    End If
    Grid1.Redraw = True
End Sub

Private Sub brcode_Change()
    refresh_grid
End Sub

Private Sub Form_Load()
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1200 '980 '680
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.gemmvc.Checked = False
End Sub

Private Sub Option1_Click()
    refresh_grid
End Sub

Private Sub Option2_Click()
    refresh_grid
End Sub

Private Sub postadj_Click()
    Dim tfile As String, i As Integer, k As Integer
    Dim tdate As String, treas As String
    tfile = Form1.webdir & "\counts\whsadj." & Form11.brcode
    If Len(Dir(tfile)) <> 0 Then
        If FileLen(tfile) > 0 Then
            If MsgBox("Adjustments are pending.", vbYesNo + vbQuestion, "are you sure...") = vbNo Then Exit Sub
        End If
    End If
    tdate = Grid1.TextMatrix(0, 2)
    If Option1 = True Then
        treas = "CHIN"
    Else
        treas = "CYCL"
    End If
    Screen.MousePointer = 11
    Open tfile For Append As #1
    k = 0
    For i = 0 To Grid1.Rows - 1
         If Val(Grid1.TextMatrix(i, 0)) <> 0 And Val(Grid1.TextMatrix(i, 4)) <> 0 Then
            Write #1, tdate;
            Write #1, Form11.brcode;            'Orgn Code
            Write #1, Form11.brcode;            'Whse Code
            Write #1, "LOT1";                   'Lot Code
            Write #1, Grid1.TextMatrix(i, 0);
            Write #1, Grid1.TextMatrix(i, 1);
            Write #1, Grid1.TextMatrix(i, 4);
            Write #1, treas;
            Write #1, Form1.wduser;
            Write #1, Format(Now, "m-d-yyyy")
            k = k + 1
        End If
    Next i
    Close #1
    Screen.MousePointer = 0
    MsgBox Format$(k, "0") & " Adjustment qtys have been posted.", vbOKOnly + vbInformation, "Reason code " & treas & "..."
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    rt = Me.Caption
    rh = "Branch " & brcode.Caption
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Call printflexgrid(Printer, Grid1, rt, rh, rf)
End Sub

