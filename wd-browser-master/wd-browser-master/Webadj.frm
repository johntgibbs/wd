VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Webadj 
   Caption         =   "Inventory Adjustments"
   ClientHeight    =   9045
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form13"
   ScaleHeight     =   9045
   ScaleWidth      =   13365
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4335
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7646
      _Version        =   327680
      ForeColor       =   4194368
      BackColorFixed  =   12648447
      BackColorSel    =   12582912
      FocusRect       =   0
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
      Left            =   720
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
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
      Left            =   3960
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   240
      Width           =   3255
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
      Left            =   7440
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   4575
   End
   Begin MSFlexGridLib.MSFlexGrid gemmies 
      Height          =   3615
      Left            =   960
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6376
      _Version        =   327680
   End
   Begin VB.Label whscode 
      Caption         =   "Label3"
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
      Left            =   12120
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label adjfile 
      Caption         =   "Label3"
      Height          =   255
      Left            =   9840
      TabIndex        =   7
      Top             =   6360
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Reason Code:"
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
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
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
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.Menu filemenu 
      Caption         =   "&File"
      Begin VB.Menu prtmenu 
         Caption         =   "Print"
      End
      Begin VB.Menu save_adj 
         Caption         =   "Save"
      End
      Begin VB.Menu xitmenu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edmenu 
      Caption         =   "&Edit"
      Begin VB.Menu insrec 
         Caption         =   "Insert Record - F10"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete Record - F9"
      End
   End
End
Attribute VB_Name = "Webadj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String, edrow As Integer
Dim savewhs As Boolean

Private Sub update_rec()
    Dim i As Integer
    Dim pdesc As String
    If edcell = "date" Then
        If IsDate(Grid1.TextMatrix(edrow, 1)) Then
            Grid1.TextMatrix(edrow, 1) = Format(Grid1.TextMatrix(edrow, 1), "m-d-yyyy")
        Else
            If IsDate(Text1) Then
                Grid1.TextMatrix(edrow, 1) = Format(Text1, "m-d-yyyy")
            Else
                Grid1.TextMatrix(edrow, 1) = Format(Now, "m-d-yyyy")
            End If
        End If
    End If
    If edcell = "qty" Then Grid1.TextMatrix(edrow, 4) = Format(Val(Grid1.TextMatrix(edrow, 4)), "#")
    If Grid1.TextMatrix(edrow, 2) <> gemmies.TextMatrix(gemmies.Row, 0) Then
        For i = 1 To gemmies.Rows - 1
            'If LCase(gemmies.TextMatrix(i, 0)) >= LCase(Grid1.TextMatrix(edrow, 2)) Then
            If LCase(gemmies.TextMatrix(i, 0)) = LCase(Grid1.TextMatrix(edrow, 2)) Then
                gemmies.Row = i
                Exit For
            End If
        Next i
    End If
    If LCase(Grid1.TextMatrix(edrow, 2)) = LCase(gemmies.TextMatrix(gemmies.Row, 0)) Then
        Grid1.TextMatrix(edrow, 2) = gemmies.TextMatrix(gemmies.Row, 0)
        pdesc = gemmies.TextMatrix(gemmies.Row, 1)
    Else
        pdesc = "Item # Not on File"
    End If
    If edcell = "item" Then Grid1.TextMatrix(edrow, 3) = pdesc
    If edcell = "reason" Then
        Grid1.TextMatrix(edrow, 5) = UCase(Grid1.TextMatrix(edrow, 5))
        For i = 0 To Combo2.ListCount - 1
            'If Left(Combo2.List(i), 4) = Grid1.TextMatrix(edrow, 5) Then
            If Trim(Left(Combo2.List(i), 4)) = Grid1.TextMatrix(edrow, 5) Then                      'jv031116
                Combo2.ListIndex = i
                Exit For
            End If
        Next i
        'Grid1.TextMatrix(edrow, 5) = Left(Combo2, 4)
        Grid1.TextMatrix(edrow, 5) = Trim(Left(Combo2, 4))                                          'jv031116
    End If
    If Len(Grid1.TextMatrix(edrow, 1)) = 0 Then Grid1.TextMatrix(edrow, 1) = Text1
    'If Len(Grid1.TextMatrix(edrow, 5)) = 0 Then Grid1.TextMatrix(edrow, 5) = Left(Combo2, 4)
    If Len(Grid1.TextMatrix(edrow, 5)) = 0 Then Grid1.TextMatrix(edrow, 5) = Trim(Left(Combo2, 4))  'jv031116
    Grid1.TextMatrix(edrow, 6) = Form1.wduser
    Grid1.TextMatrix(edrow, 7) = Format(Now, "m-d-yyyy")
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

Private Sub adjfile_Change()
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    If Len(Dir(adjfile)) > 0 Then
        Open adjfile For Input As #1
        Do Until EOF(1)
            Input #1, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10
            f0 = Chr(9) & f1 & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9)
            f0 = f0 & f8 & Chr(9) & f9 & Chr(9) & f10
            Grid1.AddItem f0
        Loop
        Close #1
    End If
    Grid1.FormatString = "^|^Tran Date|^Item|<Product|^Adj Qty|^Reason|^User|^Posted"
    Grid1.ColWidth(0) = 250
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 4000
    Grid1.ColWidth(4) = 1200
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 2000
    Grid1.ColWidth(7) = 1200
    Grid1.AddItem "..."
    Grid1.Redraw = True
End Sub

Private Sub delrec_Click()
    If Grid1.Row = 0 Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 0) = "..." Then Exit Sub
    Grid1.RemoveItem Grid1.Row
    savewhs = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 120 Then   'F9
        If Shift = 0 Then delrec_Click
        'If Shift = 1 Then deltag_Click      'Shift - F9
        'If Shift = 2 Then delalltag_Click   'Ctl-F9
    End If
    If KeyCode = 121 Then   'F10
        KeyCode = 0
        insrec_Click
    End If
End Sub

Private Sub Form_Load()
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim ds As ADODB.Recordset, s As String                              'jv031116
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
    Text1 = Format(Now, "m-d-yyyy")
    Combo2.Clear
    s = "select listdisplay from valuelists where listname = 'r12adjreasons' order by listdisplay"  'jv031116
    Set ds = wdb.Execute(s)                                                                         'jv031116
    If ds.BOF = False Then                                                                          'jv031116
        ds.MoveFirst                                                                                'jv031116
        Do Until ds.EOF                                                                             'jv031116
            Combo2.AddItem ds!listdisplay                                                           'jv031116
            ds.MoveNext                                                                             'jv031116
        Loop                                                                                        'jv031116
    End If                                                                                          'jv031116
    ds.Close                                                                                        'jv031116
    'Combo2.AddItem "CYCL - Cycle Counts"
    'Combo2.AddItem "CHIN - Change Inventory"
    'Combo2.AddItem "BKRM - Break Room"
    'Combo2.AddItem "DMGE - Damaged"
    Combo2.ListIndex = 0
    refresh_gemmies
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1580
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If savewhs = True Then save_adj_Click
    Form1.invadj.Checked = False
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Grid1.Col = 2 Then
            Grid1.Col = 4
        Else
            Grid1.Col = 2
            If Grid1.Row < Grid1.Rows - 1 Then Grid1.Row = Grid1.Row + 1
        End If
        Exit Sub
    End If
    If Grid1.Col = 0 Then Exit Sub
    If Grid1.Col = 3 Then Exit Sub
    If Grid1.Col > 5 Then Exit Sub
    If Len(edcell) = 0 Then Grid1.Text = ""
    If Grid1.Col = 1 Then edcell = "date"
    If Grid1.Col = 2 Then edcell = "item"
    If Grid1.Col = 4 Then edcell = "qty"
    If Grid1.Col = 5 Then edcell = "reason"
    If KeyAscii = 8 Then
        If Len(Grid1.Text) <= 1 Then
            Grid1.Text = ""
        Else
            Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
        End If
        edrow = Grid1.Row
    End If
    If KeyAscii > 31 And KeyAscii < 127 Then
        Grid1.Text = Grid1.Text + Chr(KeyAscii)
        If Grid1.TextMatrix(Grid1.Row, 0) = "..." Then
            Grid1.TextMatrix(Grid1.Row, 0) = ""
            Grid1.AddItem "..."
        End If
        edrow = Grid1.Row
    End If
End Sub

Private Sub Grid1_RowColChange()
    If Len(edcell) > 0 Then
        update_rec
        DoEvents
    End If
End Sub

Private Sub insrec_Click()
    Dim s As String
    s = Chr(9) & Format(Text1, "m-d-yyyy") & Chr(9)
    s = s & Chr(9)
    s = s & Chr(9)
    s = s & Chr(9)
    's = s & Left(Combo2, 4) & Chr(9)
    s = s & Trim(Left(Combo2, 4)) & Chr(9)                          'jv031116
    s = s & Form1.wduser & Chr(9)
    s = s & Format(Now, "m-d-yyyy")
    Grid1.AddItem s, Grid1.Row
    savewhs = True
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    save_adj_Click
    rt = Me.Caption & " - Branch " & Right(adjfile, 2)
    rh = Format(Now, "mmmm d, yyyy")
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Call printflexgrid(Printer, Grid1, rt, rh, rf)
End Sub

Private Sub save_adj_Click()
    Dim i As Integer, k As Integer
    Open adjfile For Output As #1
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 4)) <> 0 Then
            Write #1, Grid1.TextMatrix(i, 1);
            Write #1, whscode.Caption;      'Orgn Code
            Write #1, whscode.Caption;      'Whse Code
            Write #1, "LOT1";               'Default Lot Code
            For k = 2 To 6
                Write #1, Grid1.TextMatrix(i, k);
            Next k
            Write #1, Grid1.TextMatrix(i, 7)
        End If
    Next i
    Close #1
    savewhs = False
End Sub

Private Sub xitmenu_Click()
    Unload Me
End Sub
