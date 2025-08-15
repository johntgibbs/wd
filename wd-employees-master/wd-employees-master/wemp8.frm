VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form8 
   Caption         =   "Employees Report"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   9210
   LinkTopic       =   "Form8"
   ScaleHeight     =   6495
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox qstr 
      Height          =   1215
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   6615
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8070
      _Version        =   327680
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label qtrig 
      Caption         =   "0"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Menu showmenu 
      Caption         =   "Show Fields"
      Begin VB.Menu sf 
         Caption         =   "BB Number"
         Index           =   0
      End
      Begin VB.Menu sf 
         Caption         =   "SS Number"
         Index           =   1
      End
      Begin VB.Menu sf 
         Caption         =   "First Name"
         Index           =   2
      End
      Begin VB.Menu sf 
         Caption         =   "Middle Name"
         Index           =   3
      End
      Begin VB.Menu sf 
         Caption         =   "Last Name"
         Index           =   4
      End
      Begin VB.Menu sf 
         Caption         =   "Maiden Name"
         Index           =   5
      End
      Begin VB.Menu sf 
         Caption         =   "Nickname"
         Index           =   6
      End
      Begin VB.Menu sf 
         Caption         =   "SS Name"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu sf 
         Caption         =   "DL Name"
         Index           =   8
      End
      Begin VB.Menu sf 
         Caption         =   "DL Number"
         Index           =   9
      End
      Begin VB.Menu sf 
         Caption         =   "Home Phone"
         Index           =   10
      End
      Begin VB.Menu sf 
         Caption         =   "Work Phone"
         Index           =   11
      End
      Begin VB.Menu sf 
         Caption         =   "Street"
         Index           =   12
      End
      Begin VB.Menu sf 
         Caption         =   "City"
         Index           =   13
      End
      Begin VB.Menu sf 
         Caption         =   "State"
         Index           =   14
      End
      Begin VB.Menu sf 
         Caption         =   "Zip Code"
         Index           =   15
      End
      Begin VB.Menu sf 
         Caption         =   "County"
         Index           =   16
      End
      Begin VB.Menu sf 
         Caption         =   "Veteran"
         Index           =   17
      End
      Begin VB.Menu sf 
         Caption         =   "Vet Years"
         Index           =   18
      End
      Begin VB.Menu sf 
         Caption         =   "Vietnam Vet"
         Index           =   19
      End
      Begin VB.Menu sf 
         Caption         =   "Birthday"
         Index           =   20
      End
      Begin VB.Menu sf 
         Caption         =   "Date Employed"
         Index           =   21
      End
      Begin VB.Menu sf 
         Caption         =   "Full Time"
         Index           =   22
      End
      Begin VB.Menu sf 
         Caption         =   "Full time Date"
         Index           =   23
      End
      Begin VB.Menu sf 
         Caption         =   "Termination Date"
         Index           =   24
      End
      Begin VB.Menu sf 
         Caption         =   "Term Reason"
         Index           =   25
      End
      Begin VB.Menu sf 
         Caption         =   "Marital Status"
         Index           =   26
      End
      Begin VB.Menu sf 
         Caption         =   "Parent Status"
         Index           =   27
      End
      Begin VB.Menu sf 
         Caption         =   "Radio Code"
         Index           =   28
      End
      Begin VB.Menu sf 
         Caption         =   "Cell Phone"
         Index           =   29
      End
      Begin VB.Menu sf 
         Caption         =   "Department"
         Index           =   30
      End
      Begin VB.Menu sf 
         Caption         =   "All"
         Index           =   31
      End
   End
   Begin VB.Menu hidef 
      Caption         =   "&Hide Field"
   End
   Begin VB.Menu prtgrid 
      Caption         =   "&Print"
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gc(0 To 30) As Integer
Private Sub print_grid(r1 As Integer, r2 As Integer)
    Dim i As Integer, k As Integer, j As Integer
    Dim xs As Long, xe As Long, xm As Long
    Dim ys As Long, ye As Long
    
    '  Print Check Off
    xs = 0: xe = xs
    For i = 0 To Grid1.Cols - 1
        If Grid1.ColWidth(i) > 10 Then xe = xe + Grid1.ColWidth(i)
    Next i
    If xe > 11600 Then
        Printer.Orientation = 2
    Else
        Printer.Orientation = 1
    End If
    
    Printer.FontTransparent = True
    Printer.FillStyle = 0
    Printer.FillColor = QBColor(15)
    Printer.DrawMode = 1
    Printer.ForeColor = QBColor(0)
    
    Printer.FontName = "MS Serif"
    Printer.FontTransparent = True
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    Printer.Print "W/D Employee List"
    Printer.Print Form8.Caption
    Printer.Print Format(Now, "mmmm d, yyyy")

    Printer.FontSize = 8
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
    For k = 0 To Grid1.Cols - 1
        If Grid1.ColWidth(k) > 10 Then
            Printer.PSet (xm, 1230)
            Printer.Print Grid1.TextMatrix(0, k)
            xm = xm + Grid1.ColWidth(k)
        End If
    Next k
    j = 1
    For i = r1 To r2
        xm = xs + 100
        For k = 0 To Grid1.Cols - 1
            If Grid1.ColWidth(k) > 10 Then
                Printer.PSet (xm, j * 240 + 1230)
                Printer.Print Grid1.TextMatrix(i, k)
                xm = xm + Grid1.ColWidth(k)
            End If
        Next k
        j = j + 1
    Next i
    ys = 1200
    xm = xs
    Printer.DrawWidth = 6
    For i = 0 To Grid1.Cols - 1
        If Grid1.ColWidth(i) > 10 Then
            Printer.Line (xm, ys)-(xm, ye)
            xm = xm + Grid1.ColWidth(i)
        End If
    Next i
    Printer.Line (xm, ys)-(xm, ye)
    Printer.EndDoc
End Sub

Private Sub refresh_grid()
    Dim db As Database, ds As Recordset, sqlx As String
    Dim ss As Recordset
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 30
    Grid1.FixedCols = 0
    sqlx = "^BB Number|"
    sqlx = sqlx & "^SS Number|"
    sqlx = sqlx & "<1st Name|"
    sqlx = sqlx & "^MI|"
    sqlx = sqlx & "<Last Name|"
    sqlx = sqlx & "<Maiden Name|"
    sqlx = sqlx & "<Nickname|"
    sqlx = sqlx & "<SS Name|"
    sqlx = sqlx & "<DL Name|"
    sqlx = sqlx & "<DL Number|"
    sqlx = sqlx & "^Home Phone|"
    sqlx = sqlx & "^Work Phone|"
    sqlx = sqlx & "<Street|"
    sqlx = sqlx & "<City|"
    sqlx = sqlx & "<State|"
    sqlx = sqlx & "<ZipCode|"
    sqlx = sqlx & "<County|"
    sqlx = sqlx & "^Veteran|"
    sqlx = sqlx & "^Vet Years|"
    sqlx = sqlx & "^Viet Vet|"
    sqlx = sqlx & "^Birthday|"
    sqlx = sqlx & "^Date Employed|"
    sqlx = sqlx & "^Full Time|"
    sqlx = sqlx & "^Date Full Time|"
    sqlx = sqlx & "^Date Termed|"
    sqlx = sqlx & "<Term Reason|"
    sqlx = sqlx & "^Marital Status|"
    sqlx = sqlx & "^Parent Status|"
    sqlx = sqlx & "^Radio Code|"
    sqlx = sqlx & "^Cell Phone|"
    sqlx = sqlx & "<Department"
    Grid1.FormatString = sqlx
    For i = 0 To 30
        If sf(i).Visible = False Then
            Grid1.ColWidth(i) = gc(i)
        Else
            Grid1.ColWidth(i) = 1
        End If
    Next i
    'Grid1.ColWidth(0) = 1 ' 1000
    'Grid1.ColWidth(1) = 1 '1100
    'Grid1.ColWidth(2) = 1 '1000
    'Grid1.ColWidth(3) = 1 '1000
    'Grid1.ColWidth(4) = 1 '1400
    'Grid1.ColWidth(5) = 1 '1400
    'Grid1.ColWidth(6) = 1 '1800
    'Grid1.ColWidth(7) = 2300
    'Grid1.ColWidth(8) = 1 '2300
    'Grid1.ColWidth(9) = 1 '1000
    'Grid1.ColWidth(10) = 1 '1100
    'Grid1.ColWidth(11) = 1 '1100
    'Grid1.ColWidth(12) = 1 '2200
    'Grid1.ColWidth(13) = 1 '1500
    'Grid1.ColWidth(14) = 1 '600
    'Grid1.ColWidth(15) = 1 '800
    'Grid1.ColWidth(16) = 1 '1000
    'Grid1.ColWidth(17) = 1 '800
    'Grid1.ColWidth(18) = 1 '900
    'Grid1.ColWidth(19) = 1 '900
    'Grid1.ColWidth(20) = 1 '1000
    'Grid1.ColWidth(21) = 1 '1200
    'Grid1.ColWidth(22) = 1 '900
    'Grid1.ColWidth(23) = 1 '1200
    'Grid1.ColWidth(24) = 1 '1000
    'Grid1.ColWidth(25) = 1 '1500
    'Grid1.ColWidth(26) = 1 '1500
    'Grid1.ColWidth(27) = 1 '1500
    'Grid1.ColWidth(28) = 1 '1000
    'Grid1.ColWidth(29) = 1 '1200
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    'Set ss = db.OpenRecordset("select * from departments")
    sqlx = "select * from employees"
    If Len(qstr) > 0 Then sqlx = sqlx & " where " & qstr
    sqlx = sqlx & " order by last_name,first_name,middle_name"
    'MsgBox sqlx
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ""
            For i = 1 To 28
                sqlx = sqlx & ds(i) & Chr(9)
            Next i
            sqlx = sqlx & ds!radiocode & Chr(9)
            sqlx = sqlx & ds!cellphone & Chr(9)
            Set ss = db.OpenRecordset("select deptdesc from departments where id = " & ds!deptcode)
            'ss.FindFirst "id = " & ds!deptcode
            'If ss.NoMatch Then
            If ss.BOF = True Then
                sqlx = sqlx & " "
            Else
                ss.MoveFirst
                sqlx = sqlx & ss!deptdesc
            End If
            ss.Close
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    'ds.Close: ss.Close: db.Close
    ds.Close: db.Close
End Sub
Private Sub Form_Deactivate()
    Dim i As Integer, x As Integer
    If Form8.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form8" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form8.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form8.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form8.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form8.Width
                x = 2
                Exit For
            End If
        Next i
        If x <> 2 Then Form1.frmgrid.AddItem "form8" & Chr(9) & Form8.Top & Chr(9) & Form8.Left & Chr(9) & Form8.Height & Chr(9) & Form8.Width
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form8" Then
            Form8.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form8.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form8.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form8.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i

    gc(0) = 1000: gc(1) = 1100: gc(2) = 1000
    gc(3) = 1000: gc(4) = 1400: gc(5) = 1400
    gc(6) = 1800: gc(7) = 2300: gc(8) = 2300
    gc(9) = 1000: gc(10) = 1100: gc(11) = 1100
    gc(12) = 2200: gc(13) = 1500: gc(14) = 600
    gc(15) = 800: gc(16) = 1000: gc(17) = 800
    gc(18) = 900: gc(19) = 900: gc(20) = 1000
    gc(21) = 1200: gc(22) = 900: gc(23) = 1200
    gc(24) = 1000: gc(25) = 1500: gc(26) = 1500
    gc(27) = 1500: gc(28) = 1000: gc(29) = 1200
    gc(30) = 1200
    'refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Form8.Width - 80
    If Form8.Height > 2000 Then Grid1.Height = Form8.Height - 680
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Deactivate
End Sub

Private Sub hidef_Click()
    Dim i As Integer, k As Integer
    If Grid1.ColWidth(Grid1.Col) > 100 Then
        gc(Grid1.Col) = Grid1.ColWidth(Grid1.Col)
    Else
        gc(Grid1.Col) = 400
    End If
    Grid1.ColWidth(Grid1.Col) = 1
    sf(Grid1.Col).Visible = True
    For i = 0 To Grid1.Cols - 1
        If Grid1.ColWidth(i) > 10 Then
            k = i
            If i > Grid1.Col Then Exit For
        End If
    Next i
    Grid1.Col = k
End Sub

Private Sub prtgrid_Click()
    Dim i As Integer, k As Integer, j As Integer
    Screen.MousePointer = 11
    k = 1: j = 1
    For i = 1 To Grid1.Rows - 1
        If j > 55 Then
            Call print_grid(k, i)
            k = i + 1
            j = 0
        End If
        j = j + 1
    Next i
    If k < Grid1.Rows - 1 Then
        Call print_grid(k, Grid1.Rows - 1)
    End If
    Screen.MousePointer = 0
End Sub

Private Sub qtrig_Change()
    refresh_grid
End Sub

Public Sub sf_Click(Index As Integer)
    If sf(Index).Caption = "All" Then
        For i = 0 To Grid1.Cols - 1
            If Grid1.ColWidth(i) < 100 Then Grid1.ColWidth(i) = gc(i)
            sf(i).Visible = False
        Next i
    Else
        sf(Index).Visible = False
        Grid1.ColWidth(Index) = gc(Index)
    End If
End Sub
