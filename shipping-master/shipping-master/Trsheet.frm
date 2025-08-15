VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Trsheet 
   Caption         =   "Trailer History Sheets"
   ClientHeight    =   5700
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5700
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2415
      Left            =   0
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4260
      _Version        =   327680
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Drop Record"
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add Record"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2895
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5106
      _Version        =   327680
      Cols            =   11
      FixedCols       =   2
      ForeColor       =   12582912
      BackColorFixed  =   12648384
      BackColorSel    =   128
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin VB.ListBox List2 
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
      Top             =   0
      Visible         =   0   'False
      Width           =   735
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
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   3375
   End
   Begin VB.ListBox List1 
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
      Left            =   1560
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   735
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
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Import Rbase"
      Height          =   255
      Left            =   7800
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Dates"
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
      Left            =   8160
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Left            =   6720
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Branch:"
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
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   735
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
      Width           =   1215
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu addrec 
         Caption         =   "Add Record"
      End
      Begin VB.Menu droprec 
         Caption         =   "Drop Record"
      End
   End
End
Attribute VB_Name = "Trsheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String
Private Sub update_ts()
    Dim ds As adodb.Recordset, sqlx As String, i As Integer
    On Error GoTo vberror
    sqlx = "select * from trhist where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Grid1.Text = Val(Grid1.Text)
        If Val(Grid1.Text) = 0 Then Grid1.Text = ""
        If edcell = "trl1" Then sqlx = "Update trhist set trl1 = " & Val(Grid1.Text) & " Where id = " & ds!id
        If edcell = "trl2" Then sqlx = "Update trhist set trl2 = " & Val(Grid1.Text) & " Where id = " & ds!id
        If edcell = "trl3" Then sqlx = "Update trhist set trl3 = " & Val(Grid1.Text) & " Where id = " & ds!id
        If edcell = "trl4" Then sqlx = "Update trhist set trl4 = " & Val(Grid1.Text) & " Where id = " & ds!id
        If edcell = "trl5" Then sqlx = "Update trhist set trl5 = " & Val(Grid1.Text) & " Where id = " & ds!id
        If edcell = "trl6" Then sqlx = "Update trhist set trl6 = " & Val(Grid1.Text) & " Where id = " & ds!id
        If edcell = "trladj" Then sqlx = "Update trhist set trladj = " & Val(Grid1.Text) & " Where id = " & ds!id
        If edcell = "trlbob" Then sqlx = "Update trhist set trlbob = " & Val(Grid1.Text) & " Where id = " & ds!id
        Sdb.Execute sqlx
    End If
    ds.Close
    edcell = ""
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "update_ts", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " update_ts - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String, tot As Long
    On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1
    sqlx = "select * from trhist where plant = " & List1
    sqlx = sqlx & " and branch = " & List2
    sqlx = sqlx & " order by shipdate DESC"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!id & Chr(9)
            sqlx = sqlx & Format(ds!shipdate, "m-d-yyyy") & Chr(9)
            sqlx = sqlx & Format(ds!trl1, "#####") & Chr(9)
            sqlx = sqlx & Format(ds!trl2, "#####") & Chr(9)
            sqlx = sqlx & Format(ds!trl3, "#####") & Chr(9)
            sqlx = sqlx & Format(ds!trl4, "#####") & Chr(9)
            sqlx = sqlx & Format(ds!trl5, "#####") & Chr(9)
            sqlx = sqlx & Format(ds!trl6, "#####") & Chr(9)
            sqlx = sqlx & Format(ds!trladj, "#####") & Chr(9)
            sqlx = sqlx & Format(ds!trlbob, "#####") & Chr(9)
            tot = ds!trl1 + ds!trl2 + ds!trl3 + ds!trl4 + ds!trl5
            If Len(ds!trladj) > 0 Then tot = tot + ds!trladj
            If Len(ds!trlbob) > 0 Then tot = tot + ds!trlbob
            'tot = tot + ds!trl6 + ds!trladj + ds!trlbob
            sqlx = sqlx & Format(tot, "######")
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "ID|^Date|^#1|^#2|^#3|^#4|^#5|^#6|^Adjust|^BobTail|^Net"
    Grid1.ColWidth(0) = 1: Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 900: Grid1.ColWidth(3) = 900
    Grid1.ColWidth(4) = 900: Grid1.ColWidth(5) = 900
    Grid1.ColWidth(6) = 900: Grid1.ColWidth(7) = 900
    Grid1.ColWidth(8) = 900: Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 1000
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

Private Sub addrec_Click()
    Command4_Click
End Sub

Private Sub Combo1_Click()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    List1.ListIndex = Combo1.ListIndex
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "combo1_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " combo1_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo2_Click()
    List2.ListIndex = Combo2.ListIndex
    Call refresh_grid
End Sub

Private Sub Command1_Click()
    Dim sqlx As String, sdate As String, edate As String
    Dim rf As String, rt As String, rh As String
    Dim ds As adodb.Recordset, s As String
    Dim i As Integer, k As Integer, j As Long, t As Long
    Dim t1 As Long, t2 As Long, t3 As Long, t4 As Long
    Dim t5 As Long, t6 As Long, t7 As Long, t8 As Long, t9 As Long
    On Error GoTo vberror
    sdate = InputBox("Starting Date:", "Start Date", "1-1-1999")
    If Len(sdate) = 0 Then Exit Sub
    If IsDate(sdate) = False Then
        MsgBox "Invalid Date Format!", vbOKOnly + vbExclamation, "Try again."
        Exit Sub
    End If
    edate = InputBox("Ending Date:", "End Date", "1-31-1999")
    If Len(edate) = 0 Then Exit Sub
    If IsDate(edate) = False Then
        MsgBox "Invalid Date Format!", vbOKOnly + vbExclamation, "Try Again."
        Exit Sub
    End If
    If Format(edate, "yyyymmdd") < Format(sdate, "yyyymmdd") Then
        MsgBox "Date Range Error..", vbOKOnly + vbExclamation, "Try Again."
        Exit Sub
    End If
    If MsgBox("Print all branches?", vbYesNo + vbQuestion, "All " & Combo1 & " Branches?") = vbNo Then
        j = List2.ListIndex
        k = List2.ListIndex
    Else
        j = 0
        k = List2.ListCount - 1
    End If
        
    For i = j To k
        Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 10
        t1 = 0: t2 = 0: t3 = 0: t4 = 0: t5 = 0
        t6 = 0: t7 = 0: t8 = 0: t9 = 0
        s = "select * from trhist where shipdate >= '" & sdate & "'"
        s = s & " and shipdate <= '" & edate & "'"
        s = s & " and plant = " & List1
        s = s & " and branch = " & List2.List(i)
        s = s & " order by shipdate"
        Set ds = Sdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = ds!shipdate & Chr(9)
                s = s & Format(ds!trl1, "#") & Chr(9)
                s = s & Format(ds!trl2, "#") & Chr(9)
                s = s & Format(ds!trl3, "#") & Chr(9)
                s = s & Format(ds!trl4, "#") & Chr(9)
                s = s & Format(ds!trl5, "#") & Chr(9)
                s = s & Format(ds!trl6, "#") & Chr(9)
                s = s & Format(ds!trladj, "#") & Chr(9)
                s = s & Format(ds!trlbob, "#") & Chr(9)
                t = ds!trl1 + ds!trl2 + ds!trl3 + ds!trl4 + ds!trl5 + ds!trl6 + ds!trladj + ds!trlbob
                s = s & t
                t1 = t1 + ds!trl1
                t2 = t2 + ds!trl2
                t3 = t3 + ds!trl3
                t4 = t4 + ds!trl4
                t5 = t5 + ds!trl5
                t6 = t6 + ds!trl6
                t7 = t7 + ds!trladj
                t8 = t8 + ds!trlbob
                t9 = t9 + t
                Grid2.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
        Grid2.AddItem " "
        s = Chr(9) & Format(t1, "#") & Chr(9)
        s = s & Format(t2, "#") & Chr(9)
        s = s & Format(t3, "#") & Chr(9)
        s = s & Format(t4, "#") & Chr(9)
        s = s & Format(t5, "#") & Chr(9)
        s = s & Format(t6, "#") & Chr(9)
        s = s & Format(t7, "#") & Chr(9)
        s = s & Format(t8, "#") & Chr(9)
        s = s & Format(t9, "#")
        Grid2.AddItem s
        Grid2.FormatString = "^ShipDate|^Trailer 1|^Trailer 2|^Trailer 3|^Trailer 4|^Trailer 5|^Trailer 6|^Corrections|^BobTail|^Total"
        Grid2.ColWidth(0) = 1000
        Grid2.ColWidth(1) = 1000
        Grid2.ColWidth(2) = 1000
        Grid2.ColWidth(3) = 1000
        Grid2.ColWidth(4) = 1000
        Grid2.ColWidth(5) = 1000
        Grid2.ColWidth(6) = 1000
        Grid2.ColWidth(7) = 1000
        Grid2.ColWidth(8) = 1000
        Grid2.ColWidth(9) = 1000
        rt = "Trailer History Sheet"
        rh = Combo1 & " --> " & Combo2.List(i) & "  " & sdate & " thru " & edate
        rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
        If j = k Then
            If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
                Call printflexgrid(Printer, Grid2, rt, rh, rf)
            Else
                Call htmlcolorgrid(Me, htmlTempFile, Grid2, rt, rh, rf, "linen", "lemonchiffon", "white")
                If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
                    i = Shell("C:\program files\internet explorer\iexplore.exe " & htmlTempFile, vbNormalFocus)
                    Exit Sub
                End If
                If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
                    i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & htmlTempFile, vbNormalFocus)
                    Exit Sub
                End If
            End If
        Else
            Call printflexgrid(Printer, Grid2, rt, rh, rf)
        End If
    Next i
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

Private Sub Command2_Click()
    Dim sqlx As String, sdate As String, edate As String
    On Error GoTo vberror
    sdate = InputBox("Starting Date:", "Start Date", "1-1-1999")
    If Len(sdate) = 0 Then Exit Sub
    If IsDate(sdate) = False Then
        MsgBox "Invalid Date Format!", vbOKOnly + vbExclamation, "Try again."
        Exit Sub
    End If
    edate = InputBox("Ending Date:", "End Date", "1-31-1999")
    If Len(edate) = 0 Then Exit Sub
    If IsDate(edate) = False Then
        MsgBox "Invalid Date Format!", vbOKOnly + vbExclamation, "Try Again."
        Exit Sub
    End If
    If Format(edate, "yyyymmdd") < Format(sdate, "yyyymmdd") Then
        MsgBox "Date Range Error..", vbOKOnly + vbExclamation, "Try Again."
        Exit Sub
    End If
    sqlx = "delete from trhist where plant = " & List1
    If MsgBox("Clear all branches?", vbYesNo + vbQuestion, "All " & Combo1 & " Branches") = vbNo Then
        sqlx = sqlx & " and branch = " & List2
    End If
    sqlx = sqlx & " and shipdate >= '" & sdate & "'"
    sqlx = sqlx & " and shipdate <= '" & edate & "'"
    Sdb.Execute sqlx
    Call refresh_grid
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

Private Sub Command3_Click()
    Dim ds As adodb.Recordset, pkey As Long, sqlx As String
    Dim f1, f2, f3, f4, f5, f6, f7, f8, f9
    On Error GoTo vberror
    Screen.MousePointer = 11
    Sdb.Execute "delete from trhist"
    Open "C:\trhist.txt" For Input As #1
    Input #1, f1, f2, f3, f4, f5, f6, f7, f8, f9
    Do Until EOF(1)
        pkey = wd_seq("trhist", Form1.shipdb)
        sqlx = "Insert into trhist (id, plant, branch, shipdate, trl1, trl2, trl3, trl4, trl5, trl6, trladj"
        sqlx = sqlx & ", trlbob) Values (" & pkey
        sqlx = sqlx & ", 50"
        sqlx = sqlx & ", " & Val(f1)
        sqlx = sqlx & ", '" & f2 & "'"
        sqlx = sqlx & ", " & Val(f3)
        sqlx = sqlx & ", " & Val(f4)
        sqlx = sqlx & ", " & Val(f5)
        sqlx = sqlx & ", " & Val(f6)
        sqlx = sqlx & ", " & Val(f7)
        sqlx = sqlx & ", 0"
        sqlx = sqlx & ", " & Val(f8)
        sqlx = sqlx & ", " & Val(f9) & ")"
        Sdb.Execute sqlx
        Input #1, f1, f2, f3, f4, f5, f6, f7, f8, f9
    Loop
    Close #1
    Screen.MousePointer = 0
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


Private Sub Command4_Click()
    Dim pdate As String, ds As adodb.Recordset, sqlx As String
    Dim pkey As Long, y As Integer, i As Integer, j As Integer
    On Error GoTo vberror
    pdate = InputBox("Trailer Date:", "Add Record.", Format(Now, "m-d-yyyy"))
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Invalid Date Format!", vbOKOnly + vbExclamation, "Sorry, cannot add..."
        Exit Sub
    End If
    pkey = wd_seq("trhist", Form1.shipdb)
    sqlx = "Insert into trhist (id, plant, branch, shipdate, trl1, trl2, trl3, trl4, trl5, trl6, trladj, trlbob)"
    sqlx = sqlx & " Values (" & pkey
    sqlx = sqlx & ", " & Val(List1)
    sqlx = sqlx & ", " & Val(List2)
    sqlx = sqlx & ", '" & pdate & "', 0, 0, 0, 0, 0, 0, 0, 0)"
    Sdb.Execute sqlx
    If Grid1.Rows < 2 Then
        Grid1.AddItem pkey & Chr(9) & Format(pdate, "m-d-yyyy")
        Exit Sub
    End If
    y = 1
    For i = 1 To Grid1.Rows - 1
        y = i
        If Format$(pdate, "yyyymmdd") >= Format$(Grid1.TextMatrix(i, 1), "yyyymmdd") Then Exit For
    Next i
    Grid1.AddItem " "
    For i = Grid1.Rows - 2 To y Step -1
        For j = 0 To Grid1.Cols - 1
            Grid1.TextMatrix(i + 1, j) = Grid1.TextMatrix(i, j)
        Next j
    Next i
    Grid1.TextMatrix(y, 0) = pkey
    Grid1.TextMatrix(y, 1) = Format(pdate, "m-d-yyyy")
    Grid1.TextMatrix(y, 2) = "": Grid1.TextMatrix(y, 3) = "0"
    Grid1.TextMatrix(y, 4) = "0": Grid1.TextMatrix(y, 5) = "0"
    Grid1.TextMatrix(y, 6) = "0": Grid1.TextMatrix(y, 7) = "0"
    Grid1.TextMatrix(y, 8) = "0": Grid1.TextMatrix(y, 9) = "0"
    Grid1.Row = y: Grid1.Col = 2
    Grid1.SetFocus
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command4_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command4_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command5_Click()
    Dim sqlx As String
    On Error GoTo vberror
    If Grid1.Row = 0 Then Exit Sub
    If MsgBox("Ok to drop record?", vbYesNo + vbQuestion, "Drop Record") = vbNo Then Exit Sub
    sqlx = "delete from trhist where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Sdb.Execute sqlx
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_grid
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command5_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command5_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub droprec_Click()
    Command5_Click
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Len(edcell) > 0 Then
        If MsgBox("Update trailer record?", vbYesNo + vbQuestion, "Save changes...") = vbYes Then
            Call update_ts
        Else
            edcell = ""
        End If
    End If
    If Trsheet.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
            If Form1.FrmGrid.Text = "trsheet" Then
                Form1.FrmGrid.Col = 1: Form1.FrmGrid.Text = Trsheet.Top
                Form1.FrmGrid.Col = 2: Form1.FrmGrid.Text = Trsheet.Left
                Form1.FrmGrid.Col = 3: Form1.FrmGrid.Text = Trsheet.Height
                Form1.FrmGrid.Col = 4: Form1.FrmGrid.Text = Trsheet.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Trsheet.ActiveControl.Name = "Grid1" Then
        If KeyCode = 45 Then Call Command4_Click
        If KeyCode = 46 Then Call Command5_Click
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer, ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    For i = 1 To Form1.FrmGrid.Rows - 1
        Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
        If Form1.FrmGrid.Text = "trsheet" Then
            Form1.FrmGrid.Col = 1: Trsheet.Top = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 2: Trsheet.Left = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 3: Trsheet.Height = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 4: Trsheet.Width = Val(Form1.FrmGrid.Text)
            Exit For
        End If
    Next i
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Combo1.Clear: Combo2.Clear: List1.Clear: List2.Clear
    sqlx = "select * from plants order by plant"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds!plantname
            List1.AddItem ds!plant
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "SELECT * FROM branches WHERE branch NOT IN (97, 98, 99) ORDER BY branch"      'jv012016
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            List2.AddItem ds!branch
            Combo2.AddItem ds!branchname
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "form_load", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " form_load - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 110
    If Trsheet.Height > 4000 Then
        Grid1.Height = Trsheet.Height - 1600
    End If
    Grid2.Width = Me.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_GotFocus()
    Grid1.FocusRect = flexFocusNone
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Grid1.Col = Grid1.Cols - 1 Then
            SendKeys "{HOME}{DOWN}"
        Else
            SendKeys "{RIGHT}"
        End If
        Exit Sub
    End If
    If Grid1.Row = 0 Then Exit Sub
    If Grid1.Col < 2 Or Grid1.Col > 9 Then Exit Sub
    If Len(edcell) = 0 Then Grid1.Text = ""
    If Grid1.Col = 2 Then edcell = "trl1"
    If Grid1.Col = 3 Then edcell = "trl2"
    If Grid1.Col = 4 Then edcell = "trl3"
    If Grid1.Col = 5 Then edcell = "trl4"
    If Grid1.Col = 6 Then edcell = "trl5"
    If Grid1.Col = 7 Then edcell = "trl6"
    If Grid1.Col = 8 Then edcell = "trladj"
    If Grid1.Col = 9 Then edcell = "trlbob"
    If KeyAscii = 8 Then
        If Len(Grid1.Text) > 1 Then
            Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
        Else
            Grid1.Text = ""
        End If
    End If
    If KeyAscii > 31 And KeyAscii < 127 Then
        Grid1.Text = Grid1.Text & Chr(KeyAscii)
    End If
    Grid1.TextMatrix(Grid1.Row, 10) = Grid1.TextMatrix(Grid1.Row, 2)
    For i = 3 To 9
        Grid1.TextMatrix(Grid1.Row, 10) = Val(Grid1.TextMatrix(Grid1.Row, 10)) + Val(Grid1.TextMatrix(Grid1.Row, i))
    Next i
End Sub

Private Sub Grid1_LeaveCell()
    If Len(edcell) > 0 Then Call update_ts
End Sub

Private Sub Grid1_LostFocus()
    If Len(edcell) > 0 Then Call update_ts
    Grid1.FocusRect = flexFocusLight
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub
