VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "W/D Calendar"
   ClientHeight    =   5295
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   6975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   ScaleHeight     =   5295
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Anniversaries"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Birthdays"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3480
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4560
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5530
      _Version        =   327680
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
   End
   Begin VB.PictureBox Calendar1 
      BackColor       =   &H00FFFFC0&
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   3195
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Days to view:"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu insrec 
         Caption         =   "Insert - F10"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete - F9"
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim edcell As String
Private Sub update_item()
    Dim db As Database, ds As Recordset, sqlx As String, s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then
        edcell = "": Exit Sub
    End If
    sqlx = "select * from wdevents where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    s = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, s)
    'Set db = OpenDatabase(Form1.empdb)
    Set ds = db.OpenRecordset(sqlx)
    ds.MoveFirst
    ds.Edit
    Grid1.Text = Trim(Grid1.Text)
    If edcell = "Description" Then ds!Desc = Grid1.Text
    If edcell = "Start" Then
        Grid1.Text = Format(Grid1.Text, "m-d-yyyy")
        If IsDate(Grid1.Text) = False Then
            Grid1.Text = Format(ds!sdate, "m-d-yyyy")
        Else
            ds!sdate = Grid1.Text
        End If
    End If
    If edcell = "End" Then
        Grid1.Text = Format(Grid1.Text, "m-d-yyyy")
        If IsDate(Grid1.Text) = False Then
            Grid1.Text = Format(ds!edate, "m-d-yyyy")
        Else
            ds!edate = Grid1.Text
        End If
    End If
    
    ds.Update
    ds.Close: db.Close
    edcell = ""
End Sub

Private Sub refresh_anniversary()
    Dim db As Database, ds As Recordset, sqlx As String
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select first_name,last_name,doe from employees"
    sqlx = sqlx & " where doe <> dfulltime"
    'sqlx = sqlx & " and month(doe) = " & Calendar1.Month
    sqlx = sqlx & " and employees.dot < '0'"                        'jv081815
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = "0" & Chr(9)
            sqlx = sqlx & ds!first_name & " " & ds!last_name & Chr(9)
            sqlx = sqlx & Format(ds!doe, "m-d-yyyy") & Chr(9) & Chr(9)
            sqlx = sqlx & "Anniversary - Date Employed "
            'sqlx = sqlx & DateDiff("yyyy", ds!doe, Calendar1.Value)
            sqlx = sqlx & " Years"
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "select first_name,last_name,dfulltime from employees"
    'sqlx = sqlx & " where month(dfulltime) = " & Calendar1.Month
    sqlx = sqlx & " and employees.dot < '0'"                        'jv081815
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = "0" & Chr(9)
            sqlx = sqlx & ds!first_name & " " & ds!last_name & Chr(9)
            sqlx = sqlx & Format(ds!dfulltime, "m-d-yyyy") & Chr(9) & Chr(9)
            sqlx = sqlx & "Anniversary - Date Full Time "
            'sqlx = sqlx & DateDiff("yyyy", ds!dfulltime, Calendar1.Value)
            sqlx = sqlx & " Years"
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "select spouses.first_name,anniversary,employees.first_name,"
    sqlx = sqlx & "employees.last_name from spouses,employees"
    'sqlx = sqlx & " where month(anniversary) = " & Calendar1.Month
    sqlx = sqlx & " and spouses.empkey = employees.id"
    sqlx = sqlx & " and employees.dot < '0'"                        'jv081815
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = "0" & Chr(9)
            sqlx = sqlx & ds(2) & " " & ds(3) & Chr(9)
            sqlx = sqlx & Format(ds!anniversary, "m-d-yyyy") & Chr(9) & Chr(9)
            sqlx = sqlx & "Wedding Anniversary - " & ds(0) & " "
            'sqlx = sqlx & DateDiff("yyyy", ds(1), Calendar1.Value)
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
End Sub
Private Sub refresh_birthdays()
    Dim db As Database, ds As Recordset, sqlx As String
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select first_name,last_name,dob from employees"
    'sqlx = sqlx & " where month(dob) = " & Calendar1.Month
    sqlx = sqlx & " and dot < '0'"                                  'jv081815
    sqlx = sqlx & " order by last_name,first_name"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = "0" & Chr(9)
            sqlx = sqlx & ds!first_name & " " & ds!last_name & Chr(9)
            sqlx = sqlx & Format(ds!dob, "m-d-yyyy") & Chr(9) & Chr(9)
            sqlx = sqlx & "Employee Birthday "
            'sqlx = sqlx & DateDiff("yyyy", ds!dob, Calendar1.Value)
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "select spouses.first_name,spouses.dob,"
    sqlx = sqlx & "employees.first_name,employees.last_name "
    sqlx = sqlx & "from spouses,employees"
    'sqlx = sqlx & " where month(spouses.dob) = " & Calendar1.Month
    sqlx = sqlx & " and spouses.empkey = employees.id"
    sqlx = sqlx & " and employees.dot < '0'"                            'jv081815
    sqlx = sqlx & " order by employees.last_name,employees.first_name"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = "0" & Chr(9)
            sqlx = sqlx & ds(2) & " " & ds(3) & Chr(9)
            sqlx = sqlx & Format(ds(1), "m-d-yyyy") & Chr(9) & Chr(9)
            sqlx = sqlx & "Spouse Birthday - " & ds(0) & " "
            'sqlx = sqlx & DateDiff("yyyy", ds(1), Calendar1.Value)
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "select children.first_name,children.dob,"
    sqlx = sqlx & "employees.first_name,employees.last_name "
    sqlx = sqlx & "from children,employees"
    'sqlx = sqlx & " where month(children.dob) = " & Calendar1.Month
    sqlx = sqlx & " and children.empkey = employees.id"
    sqlx = sqlx & " and employees.dot < '0'"                            'jv081815
    sqlx = sqlx & " order by employees.last_name,employees.first_name"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = "0" & Chr(9)
            sqlx = sqlx & ds(2) & " " & ds(3) & Chr(9)
            sqlx = sqlx & Format(ds(1), "m-d-yyyy") & Chr(9) & Chr(9)
            sqlx = sqlx & "Child Birthday - " & ds(0) & " "
            'sqlx = sqlx & DateDiff("yyyy", ds(1), Calendar1.Value)
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    
End Sub
Private Sub refresh_employees()
    Dim db As Database, ds As Recordset, sqlx As String
    Combo1.Clear: List1.Clear
    Combo1.AddItem "Scheduled Event"
    List1.AddItem "0"
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select id,first_name,last_name from employees"
    sqlx = sqlx & " where employees.dot < '0'"                      'jv081815
    sqlx = sqlx & " order by last_name,first_name"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds!last_name & ", " & ds!first_name
            List1.AddItem ds!id
            ds.MoveNext
        Loop
        Combo1.ListIndex = 0
    End If
    ds.Close: db.Close
End Sub

Private Sub refresh_grid()
    Dim db As Database, ds As Recordset, sqlx As String, s As String
    Dim ns As Recordset, ename As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    Grid1.FixedCols = 2
    'sqlx = "select * from wdevents where sdate <= #" & Format(DateAdd("d", Val(Text1) - 1, Calendar1.Value)) & "#"
    'sqlx = sqlx & " and edate >= #" & Calendar1.Value & "#"
    sqlx = sqlx & " order by sdate"
    s = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, s)
    'Set db = OpenDatabase(Form1.empdb)
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!empkey = 0 Then
                ename = "Scheduled Event"
            Else
                sqlx = "select first_name,last_name from employees"
                sqlx = sqlx & " where id = " & ds!empkey
                Set ns = db.OpenRecordset(sqlx)
                If ns.BOF = False Then
                    ename = ns!first_name & " " & ns!last_name
                Else
                    ename = "employee #" & ds!empkey
                End If
                ns.Close
            End If
            sqlx = ds!id & Chr(9)
            sqlx = sqlx & ename & Chr(9)
            sqlx = sqlx & Format(ds!sdate, "m-d-yyyy") & Chr(9)
            sqlx = sqlx & Format(ds!edate, "m-d-yyyy") & Chr(9)
            sqlx = sqlx & ds!Desc
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    Grid1.FormatString = "^ID|<Name|^Start|^End|<Description"
    Grid1.ColWidth(0) = 0
    Grid1.ColWidth(1) = 1700
    Grid1.ColWidth(2) = 900
    Grid1.ColWidth(3) = 900
    Grid1.ColWidth(4) = 3000
    If Check1 = 1 Then refresh_birthdays
    If Check2 = 1 Then refresh_anniversary
    If Grid1.Rows > 1 Then
        Grid1.RowHeight(-1) = Grid1.RowHeight(1) * 2
    End If
End Sub

Private Sub Calendar1_AfterUpdate()
    Call refresh_grid
End Sub

Private Sub Check1_Click()
    Call refresh_grid
End Sub

Private Sub Check2_Click()
    Call refresh_grid
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub delrec_Click()
    Dim db As Database, ds As Recordset, sqlx As String, s As String
    If Grid1.Row = 0 Then Exit Sub
    sqlx = "Ok to delete event record for " & Grid1.TextMatrix(Grid1.Row, 3)
    If MsgBox(sqlx, vbYesNo + vbQuestion, "Delete Event Record...") = vbNo Then Exit Sub
    sqlx = "select * from wdevents where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    s = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, s)
    'Set db = OpenDatabase(Form1.empdb)
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        ds.Delete
    End If
    ds.Close: db.Close
    Call refresh_grid
End Sub
Private Sub Form_Deactivate()
    Dim i As Integer, x As Integer
    If Len(edcell) > 0 Then
        If MsgBox("Save Changes?", vbYesNo + vbQuestion, "Save changes..") = vbYes Then
            Call update_item
        Else
            edcell = ""
        End If
    End If
    If Form7.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form7" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form7.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form7.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form7.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form7.Width
                x = 7
                Exit For
            End If
        Next i
        If x <> 7 Then Form1.frmgrid.AddItem "form7" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 5985 & Chr(9) & 7095
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 121 Then
        KeyCode = 0
        Call insrec_Click     'F10
    End If
    If KeyCode = 120 Then Call delrec_Click     'F9
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form7" Then
            Form7.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form7.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form7.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form7.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    'Calendar1.Month = Month(Now)
    'Calendar1.Day = Day(Now)
    'Calendar1.Year = Year(Now)
    Call refresh_employees
    Call refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Form7.Width - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Grid1.Row <> Grid1.Rows - 1 Then Grid1.Row = Grid1.Row + 1
        Exit Sub
    End If
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    If Grid1.Row = 0 Or Grid1.Col < 2 Then Exit Sub
    If Len(edcell) = 0 And (Grid1.Col = 2 Or Grid1.Col = 3) Then Grid1.Text = ""
    edcell = Grid1.TextMatrix(0, Grid1.Col)
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
End Sub

Private Sub Grid1_LeaveCell()
    If Len(edcell) > 0 Then Call update_item
End Sub

Private Sub Grid1_LostFocus()
    If Len(edcell) > 0 Then Call update_item
End Sub

Private Sub insrec_Click()
    Dim db As Database, ds As Recordset, sqlx As String
    Dim pkey As Long
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    
    sqlx = "select sequence_id from sequences where seq = 'WDEvent'"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        pkey = ds(0) + 1
    Else
        pkey = 1
    End If
    sqlx = "Insert into wdevents (id) values (" & pkey & ")"
    db.Execute sqlx
    sqlx = "select * from wdevents where id = " & pkey
    Set ds = db.OpenRecordset(sqlx)
    ds.Edit
    ds!empkey = Val(List1)
    ds.Update
    sqlx = "Update sequences set sequence_id = " & pkey & " where seq = 'WDEvents'"
    db.Execute sqlx
    
    
    
    'sqlx = "select * from wdevents where id = 0"
    'Set ds = db.OpenRecordset(sqlx)
    'ds.AddNew
    'ds!empkey = Val(List1)
    ''ds!sdate = Calendar1.Value
    ''ds!edate = Calendar1.Value
    'pkey = ds!id
    'ds.Update
    
    
    ds.Close: db.Close
    sqlx = pkey & Chr(9)
    sqlx = sqlx & Combo1 & Chr(9)
    'sqlx = sqlx & Format(Calendar1.Value, "m-d-yyyy") & Chr(9)
    'sqlx = sqlx & Format(Calendar1.Value, "m-d-yyyy") & Chr(9)
    Grid1.AddItem sqlx
    Grid1.Row = Grid1.Rows - 1
    Grid1.RowHeight(Grid1.Row) = Grid1.RowHeight(0)
    Grid1.TopRow = Grid1.Row
End Sub

Private Sub Text1_Change()
    Call refresh_grid
End Sub
