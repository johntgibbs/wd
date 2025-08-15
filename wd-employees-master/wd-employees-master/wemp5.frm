VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   1920
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   ScaleHeight     =   1920
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2778
      _Version        =   327680
      ScrollTrack     =   -1  'True
      FocusRect       =   0
   End
   Begin VB.Label ekey 
      Caption         =   "ekey"
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   960
      Width           =   1575
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
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String
Private Sub update_item()
    Dim db As Database, ds As Recordset, sqlx As String, s As String
    sqlx = "select * from econtacts where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    s = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, s)
    'Set db = OpenDatabase(Form1.empdb)
    Set ds = db.OpenRecordset(sqlx)
    ds.MoveFirst
    ds.Edit
    Grid1.Text = Trim(Grid1.Text)
    If Len(Grid1.Text) > 25 Then Grid1.Text = Left(Grid1.Text, 25)
    If edcell = "Full Name" Then ds!fullname = Grid1.Text
    If edcell = "Relationship" Then ds!relashun = Grid1.Text
    If edcell = "Home Phone" Then ds!home_phone = Grid1.Text
    If edcell = "Work Phone" Then ds!work_phone = Grid1.Text
    ds.Update
    ds.Close: db.Close
    edcell = ""
End Sub

Private Sub refresh_grid()
    Dim db As Database, ds As Recordset, sqlx As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select * from econtacts where empkey = " & Val(ekey)
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!id & Chr(9)
            sqlx = sqlx & ds!fullname & Chr(9)
            sqlx = sqlx & ds!relashun & Chr(9)
            sqlx = sqlx & ds!home_phone & Chr(9)
            sqlx = sqlx & ds!work_phone
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    Grid1.FormatString = "^ID|^Full Name|^Relationship|^Home Phone|^Work Phone"
    Grid1.ColWidth(0) = 2
    Grid1.ColWidth(1) = 3300
    Grid1.ColWidth(2) = 2000
    Grid1.ColWidth(3) = 1100
    Grid1.ColWidth(4) = 1100
End Sub

Private Sub delrec_Click()
    Dim db As Database, ds As Recordset, sqlx As String, s As String
    If Grid1.Row = 0 Then Exit Sub
    sqlx = "Ok to delete contact record for " & Grid1.TextMatrix(Grid1.Row, 1)
    If MsgBox(sqlx, vbYesNo + vbQuestion, "Delete Contact Record...") = vbNo Then Exit Sub
    sqlx = "select * from econtacts where id = " & Grid1.TextMatrix(Grid1.Row, 0)
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

Private Sub ekey_Change()
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
    If Form5.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form5" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form5.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form5.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form5.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form5.Width
                x = 5
                Exit For
            End If
        Next i
        If x <> 5 Then Form1.frmgrid.AddItem "form5" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 2610 & Chr(9) & 7335
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
        If Form1.frmgrid.TextMatrix(i, 0) = "form5" Then
            Form5.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form5.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form5.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form5.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Resize()
    Grid1.Width = Form5.Width - 80
    If Form5.Height > 1200 Then Grid1.Height = Form5.Height - 680
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Command3.FontBold = False
    Call Form_Deactivate
End Sub
Private Sub Grid1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Grid1.Row <> Grid1.Rows - 1 Then Grid1.Row = Grid1.Row + 1
        Exit Sub
    End If
    If Grid1.Row = 0 Or Grid1.Col = 0 Then Exit Sub
    If Len(edcell) = 0 And Grid1.Col > 0 Then Grid1.Text = ""
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
    Dim plast As String, pfirst As String, pkey As Long
    If Val(ekey) = 0 Then Exit Sub
    pfirst = InputBox("Full Name: ", "Full Name.....")
    If Len(pfirst) = 0 Then Exit Sub
    plast = Form1.Text3
    'plast = InputBox("Last Name: ", "Last Name......")
    'If Len(plast) = 0 Then Exit Sub
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    
    sqlx = "select sequence_id from sequences where seq = 'Econtacts'"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        pkey = ds(0) + 1
    Else
        pkey = 1
    End If
    sqlx = "Insert into econtacts (id) values (" & pkey & ")"
    db.Execute sqlx
    sqlx = "select * from econtacts where id = " & pkey
    Set ds = db.OpenRecordset(sqlx)
    ds.Edit
    ds!empkey = Val(ekey)
    ds!fullname = pfirst
    ds.Update
    sqlx = "Update sequences set sequence_id = " & pkey & " where seq = 'Econtacts'"
    db.Execute sqlx
    
    
    
    'sqlx = "select * from econtacts where id = 0"
    'Set ds = db.OpenRecordset(sqlx)
    'ds.AddNew
    'ds!empkey = Val(ekey)
    'ds!fullname = pfirst
    'pkey = ds!id
    'ds.Update
    
    
    ds.Close: db.Close
    Grid1.AddItem pkey & Chr(9) & pfirst
    Grid1.Row = Grid1.Rows - 1
End Sub

