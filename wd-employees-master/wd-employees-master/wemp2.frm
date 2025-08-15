VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4350
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   5280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   4350
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   7223
      _Version        =   327680
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label ftype 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   4215
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String
Private Sub refresh_grid()
    Dim db As Database, ds As Recordset, sqlx As String, s As String
    If ftype = "Departments" Then
        sqlx = "select * from departments order by deptdesc"
    End If
    If ftype = "Skills" Then
        sqlx = "select * from skills order by skill"
    End If
    If Len(sqlx) = 0 Then Exit Sub
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 2
    Grid1.FixedCols = 1
    s = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, s)
    'Set db = OpenDatabase(Form1.empdb)
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds(0) & Chr(9) & ds(1)
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    Grid1.FormatString = "^ID|<" & ftype
    Grid1.ColWidth(0) = 2
    Grid1.ColWidth(1) = 5000
End Sub
Private Sub update_item()
    Dim db As Database, ds As Recordset, sqlx As String, s As String
    sqlx = "select * from " & Form2.ftype & " where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    s = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, s)
    'Set db = OpenDatabase(Form1.empdb)
    Set ds = db.OpenRecordset(sqlx)
    ds.MoveFirst
    ds.Edit
    Grid1.Text = Trim(Grid1.Text)
    If edcell = "desc" Then
        If Len(Grid1.Text) > 50 Then Grid1.Text = Left(Grid1.Text, 50)
        If Len(Grid1.Text) > 0 Then
            ds(1) = Grid1.Text
        Else
            ds(1) = " "
        End If
    End If
    ds.Update
    ds.Close: db.Close
    edcell = ""
End Sub

Private Sub delrec_Click()
    Dim db As Database, ds As Recordset, sqlx As String
    If Grid1.Row = 0 Then Exit Sub
    sqlx = "Ok to delete " & Grid1.TextMatrix(Grid1.Row, 1) & "?"
    If MsgBox(sqlx, vbYesNo + vbQuestion, "Delete Record...") = vbNo Then Exit Sub
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select * from " & Form2.Caption
    sqlx = sqlx & " where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        ds.Delete
    End If
    ds.Close: db.Close
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        refresh_grid
    End If
    If Form1.Command4.FontBold = True Then Form6.skey = Val(Form6.skey) + 1
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
    If Form2.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form2" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form2.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form2.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form2.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form2.Width
                x = 2
                Exit For
            End If
        Next i
        If x <> 2 Then Form1.frmgrid.AddItem "form2" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 5040 & Chr(9) & 5400
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
        If Form1.frmgrid.TextMatrix(i, 0) = "form2" Then
            Form2.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form2.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form2.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form2.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Resize()
    Grid1.Width = Form2.Width - 80
    If Form2.Height > 2000 Then
        Grid1.Height = Form2.Height - 680
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub ftype_Change()
    Form2.Caption = ftype
    Call refresh_grid
End Sub
Private Sub Grid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Grid1.Row <> Grid1.Rows - 1 Then Grid1.Row = Grid1.Row + 1
        Exit Sub
    End If
    If Grid1.Row = 0 Or Grid1.Col = 0 Then Exit Sub
    If Len(edcell) = 0 And Grid1.Col = 1 Then Grid1.Text = ""
    If Grid1.Col = 1 Then edcell = "desc"
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
    Dim mdesc As String, pkey As Long
    sqlx = "insert a new record for " & Form2.Caption & ":"
    mdesc = InputBox(sqlx, "Insert " & Form2.Caption)
    If Len(mdesc) = 0 Then Exit Sub
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    
    sqlx = "select sequence_id from sequences where seq = '" & Form2.Caption & "'"
    MsgBox sqlx
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        pkey = ds(0) + 1
    Else
        pkey = 1
    End If
    sqlx = "Insert into " & Form2.Caption & " values (" & pkey & ", '" & mdesc & "')"
    MsgBox sqlx
    db.Execute sqlx
    
    'sqlx = "select * from employees where id = " & pkey
    'Set ds = db.OpenRecordset(sqlx)
    'ds.Edit
    'ds!first_name = pfirst
    'ds!last_name = plast
    'ds!State = "TX"
    'ds!lastmod = Format(Now, "m-d-yyyy h:mm am/pm")
    'ds!crt = crt.Caption
    'ds.Update
        
    sqlx = "Update sequences set sequence_id = " & pkey & " where seq = '" & Form2.Caption & "'"
    MsgBox sqlx
    db.Execute sqlx
    
    
    
    'sqlx = "select * from " & Form2.Caption
    'Set ds = db.OpenRecordset(sqlx)
    'ds.AddNew
    'ds(1) = mdesc
    'pkey = ds!id
    'ds.Update
    
    ds.Close: db.Close
    Grid1.AddItem pkey & Chr(9) & mdesc
    If Form1.Command4.FontBold = True Then Form6.skey = Val(Form6.skey) + 1
End Sub
