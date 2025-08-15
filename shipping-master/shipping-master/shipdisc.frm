VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form shipdisc 
   Caption         =   "Discontinued Products"
   ClientHeight    =   5235
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11340
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5235
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh"
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
      Left            =   9240
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit Comment"
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
      Left            =   7560
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3135
      Left            =   -120
      TabIndex        =   2
      Top             =   360
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5530
      _Version        =   327680
      Cols            =   4
      FixedCols       =   2
      ForeColor       =   16384
      BackColorFixed  =   12648447
      BackColorSel    =   255
      BackColorBkg    =   -2147483633
      FocusRect       =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete SKU"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert SKU"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label ycolor 
      BackColor       =   &H0080FFFF&
      Caption         =   "ycolor"
      Height          =   255
      Left            =   9000
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu insrec 
         Caption         =   "Insert Record"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete Record"
      End
      Begin VB.Menu edcomm 
         Caption         =   "Edit Comment"
      End
   End
End
Attribute VB_Name = "shipdisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modified As Boolean
Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    sqlx = "select discont.id,discont.sku,fgunit,fgdesc,discdate,discomm,plantwhs"  'jv070317
    sqlx = sqlx & " from discont,skumast"
    sqlx = sqlx & " Where discont.sku = skumast.sku"
    'sqlx = sqlx & " Order by discdate desc,discont.sku"
    sqlx = sqlx & " Order by plantwhs desc,discont.sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds(0) & Chr(9) & " " & ds(1) & Chr(9)
            sqlx = sqlx & " " & ds(2) & " " & ds(3) & Chr(9)
            sqlx = sqlx & Format(ds(4), "m-d-yyyy") & Chr(9)
            sqlx = sqlx & ds(5)
            sqlx = sqlx & Chr(9) & ds(6)                                    'jv070317
            If ds(6) = "T10" Then sqlx = sqlx & Chr(9) & "Brenham"          'jv070317
            If ds(6) = "K10" Then sqlx = sqlx & Chr(9) & "Broken Arrow"     'jv070317
            If ds(6) = "A10" Then sqlx = sqlx & Chr(9) & "Sylacauga"        'jv070317
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 5) = "K10" Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ycolor.BackColor
            End If
        Next i
        Grid1.Row = 1
    End If
    Grid1.FormatString = "ID|^SKU|Product|^Date|<Comments|^Plant|<Location"           'jv070317
    Grid1.ColWidth(0) = 1: Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 4000: Grid1.ColWidth(3) = 1400
    Grid1.ColWidth(4) = 4500
    Grid1.ColWidth(5) = 1000                                                'jv070317
    Grid1.ColWidth(6) = 1500                                                'jv070317
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid - Error Number: " & eno
        End
    End If
End Sub
Private Sub Command1_Click()                'Add SKU
    Dim ds As adodb.Recordset, sqlx As String, msku As String
    Dim pdesc As String, pkey As Long, pwhs As String, pdate As String
    On Error GoTo vberror
    msku = InputBox$("SKU #", "Add product to Discontinued List", "777")
    If Len(msku) = 0 Then Exit Sub
    sqlx = "select * from skumast where sku = '" & msku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid SKU #", vbOKOnly + vbExclamation, "Sorry"
        ds.Close
        Exit Sub
    Else
        pdesc = " " & ds!fgunit & " " & ds!fgdesc
    End If
    ds.Close
    pwhs = InputBox("Plant Warehouse (T10, K10, A10):", "Plant Warehouse", "T10")       'jv070317
    If Len(pwhs) = 0 Then Exit Sub                                                      'jv070317
    pwhs = UCase(pwhs)                                                                  'jv070317
    If pwhs <> "T10" And pwhs <> "K10" And pwhs <> "A10" Then                           'jv070317
        MsgBox "Invalid Warehouse entered: " & pwhs, vbOKOnly + vbInformation, "try again..."   'jv070317
        Exit Sub                                                                        'jv070317
    End If                                                                              'jv070317
    pdate = InputBox("Date:", "Discontinued Date...", Format(Now, "M-d-yyyy"))          'jv070317
    If Len(pdate) = 0 Then Exit Sub                                                     'jv070317
    If IsDate(pdate) = False Then                                                       'jv070317
        MsgBox "Invalid date entered: " & pdate, vbOKOnly + vbInformation, "try again..."   'jv070317
        Exit Sub                                                                        'jv070317
    End If                                                                              'jv070317
    'sqlx = "select * from discont where sku = '" & msku & "'"
    sqlx = "select * from discont where sku = '" & msku & "' and plantwhs = '" & pwhs & "'"     'jv070317
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        MsgBox "SKU " & msku & " already in " & pwhs & " list..", vbOKOnly + vbExclamation, "Sorry" 'jv070317
        ds.Close
        Exit Sub
    Else
        pkey = wd_seq("discont", Form1.shipdb)
        'sqlx = "Insert into discont (id, sku, discdate, discomm) Values (" & pkey
        sqlx = "Insert into discont (id, sku, discdate, discomm, plantwhs) Values (" & pkey 'jv070317
        sqlx = sqlx & ", '" & msku & "'"
        'sqlx = sqlx & ", '" & Format(Now, "m-d-yyyy") & "'"
        sqlx = sqlx & ", '" & Format(pdate, "m-d-yyyy") & "'"                           'jv070317
        'sqlx = sqlx & ", ' ')"
        sqlx = sqlx & ", ' ', '" & pwhs & "')"                                          'jv070317
        Sdb.Execute sqlx
        'MsgBox sqlx
    End If
    ds.Close
    modified = True
    sqlx = pkey & Chr(9) & msku & Chr(9) & pdesc & Chr(9) & pdate & Chr(9) & Chr(9) & pwhs & Chr(9) 'jv070317
    If pwhs = "T10" Then sqlx = sqlx & "Brenham"                                        'jv070317
    If pwhs = "K10" Then sqlx = sqlx & "Broken Arrow"                                   'jv070317
    If pwhs = "A10" Then sqlx = sqlx & "Sylacauga"                                      'jv070317
    Grid1.AddItem sqlx, 1                                                               'jv070317
    'Grid1.AddItem pkey & Chr(9) & msku & Chr(9) & pdesc & Chr(9) & Format(Now, "m-d-yyyy"), 1
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command1_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command1_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command2_Click()                    'Delete SKU
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    If Grid1.Row = 0 Then Exit Sub
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    If MsgBox("Delete " & Grid1.TextMatrix(Grid1.Row, 2) & " From List", vbYesNo + vbQuestion, "Are you sure?") = vbNo Then Exit Sub
    sqlx = "select * from discont where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = "Delete from discont where id = " & ds!id
            Sdb.Execute sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    modified = True
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_grid
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command2_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command2_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command3_Click()                'Edit Comments
    Dim s As String, ds As adodb.Recordset
    On Error GoTo vberror
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    s = Grid1.TextMatrix(Grid1.Row, 4)
    s = InputBox("Comments:", "Comments....", s)
    If Len(s) = 0 Then Exit Sub
    Grid1.TextMatrix(Grid1.Row, 4) = Left(s, 50)
    s = "select * from discont where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update discont set discomm = '" & Grid1.TextMatrix(Grid1.Row, 4) & "' Where id = " & ds!id
        Sdb.Execute s
    End If
    ds.Close
    modified = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command3_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command3_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command4_Click()                                        'jv070717
    Dim rt As String, rf As String, rh As String, hf As String
    Call refresh_grid
    DoEvents
    'rt = "<center><img src=" & Chr(34) & "images/wdlogo.gif" & Chr(34) & ">"
    rt = "<center><img src=" & Chr(34) & "images/bbcolor.jpg" & Chr(34) & ">"
    rt = rt & "<BR>Discontinued Products"
    rh = "Discontinued Products"
    rf = "Updated: " & Format(Now, "m-d-yyyy h:mm am/pm")
    hf = Form1.webdir & "\discont.htm"
    'hf = "u:\discont.htm"
    Grid1.Redraw = False
    htdc(0) = "cyan": gndc(0) = Me.Grid1.BackColorFixed
    htdc(1) = "yellow": gndc(1) = Me.ycolor.BackColor
    Call htmlcolorgrid(Me, hf, Grid1, rt, rh, rf, "lemonchiffon", "Linen", "White")
    Grid1.Redraw = True
End Sub

Private Sub delrec_Click()
    Command2_Click
End Sub

Private Sub edcomm_Click()
    Command3_Click
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer, psku As String, pdesc As String, pdate As String
    Dim rt As String, rf As String, rh As String, hf As String
    If modified Then
        'rt = "<center><img src=" & Chr(34) & "images/wdlogo.gif" & Chr(34) & ">"
        rt = "<center><img src=" & Chr(34) & "images/bbcolor.jpg" & Chr(34) & ">"
        rt = rt & "<BR>Discontinued Products"
        rh = "Discontinued Products"
        rf = "Updated: " & Format(Now, "m-d-yyyy h:mm am/pm")
        hf = Form1.webdir & "\discont.htm"
        'hf = "u:\discont.htm"
        Grid1.Redraw = False
        htdc(0) = "cyan": gndc(0) = Me.Grid1.BackColorFixed
        htdc(1) = "yellow": gndc(1) = Me.ycolor.BackColor
        Call htmlcolorgrid(Me, hf, Grid1, rt, rh, rf, "lemonchiffon", "Linen", "White")
        Grid1.Redraw = True
        'Open Form1.tempdir & "\discont.txt" For Output As #1
        'Print #1, "+--------------------------------------------------+"
        'Print #1, "|       Discontinued Product Listing               |"
        'Print #1, "+--------------------------------------------------+"
        'Print #1, " "
        'For i = 1 To Grid1.Rows - 1
        '    psku = Grid1.TextMatrix(i, 1)
        '    pdesc = Grid1.TextMatrix(i, 2)
        '    pdate = Grid1.TextMatrix(i, 3)
        '    Print #1, "  " & psku & " " & pdesc & " (" & pdate & ")"
        'Next i
        'Print #1, " "
        'Close #1
    End If
    If shipdisc.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            If Form1.FrmGrid.TextMatrix(i, 0) = "shipdisc" Then
                Form1.FrmGrid.TextMatrix(i, 1) = shipdisc.Top
                Form1.FrmGrid.TextMatrix(i, 2) = shipdisc.Left
                Form1.FrmGrid.TextMatrix(i, 3) = shipdisc.Height
                Form1.FrmGrid.TextMatrix(i, 4) = shipdisc.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If shipdisc.ActiveControl.Name = "Grid1" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command1_Click
        If KeyCode = 46 Or KeyCode = 120 Then Call Command2_Click
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    modified = False
    For i = 1 To Form1.FrmGrid.Rows - 1
        If Form1.FrmGrid.TextMatrix(i, 0) = "shipdisc" Then
            shipdisc.Top = Val(Form1.FrmGrid.TextMatrix(i, 1))
            shipdisc.Left = Val(Form1.FrmGrid.TextMatrix(i, 2))
            shipdisc.Height = Val(Form1.FrmGrid.TextMatrix(i, 3))
            shipdisc.Width = Val(Form1.FrmGrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Call refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 110
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1220
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub insrec_Click()
    Command1_Click
End Sub
