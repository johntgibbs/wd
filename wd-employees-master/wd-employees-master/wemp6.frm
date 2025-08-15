VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   ScaleHeight     =   4035
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   6000
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4683
      _Version        =   327680
      BackColorSel    =   128
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label skey 
      Caption         =   "skey"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label ekey 
      Caption         =   "ekey"
      Height          =   255
      Left            =   6120
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
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
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_skills()
    Dim db As Database, ds As Recordset, sqlx As String
    Combo1.Clear: List1.Clear
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select * from skills order by skill"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds!skill
            List1.AddItem ds!id
            ds.MoveNext
        Loop
        Combo1.ListIndex = 0
    End If
    ds.Close: db.Close
End Sub
Private Sub refresh_grid()
    Dim db As Database, ds As Recordset, sqlx As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 2
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select empskills.id,skill from empskills,skills"
    sqlx = sqlx & " where empkey = " & Val(ekey)
    sqlx = sqlx & " and empskills.skillkey = skills.id"
    sqlx = sqlx & " order by skill"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds(0) & Chr(9) & ds!skill
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    Grid1.FormatString = "^ID|<Skill"
    Grid1.ColWidth(0) = 2
    Grid1.ColWidth(1) = 5000
End Sub


Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub delrec_Click()
    Dim db As Database, ds As Recordset, sqlx As String, s As String
    If Grid1.Row = 0 Then Exit Sub
    sqlx = "Ok to delete skill record for " & Grid1.TextMatrix(Grid1.Row, 1)
    If MsgBox(sqlx, vbYesNo + vbQuestion, "Delete Skill Record...") = vbNo Then Exit Sub
    sqlx = "select * from empskills where id = " & Grid1.TextMatrix(Grid1.Row, 0)
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
    If Form6.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form6" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form6.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form6.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form6.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form6.Width
                x = 6
                Exit For
            End If
        Next i
        If x <> 6 Then Form1.frmgrid.AddItem "form6" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 4725 & Chr(9) & 7680
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
        If Form1.frmgrid.TextMatrix(i, 0) = "form6" Then
            Form6.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form6.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form6.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form6.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Call refresh_skills
End Sub

Private Sub Form_Resize()
    Grid1.Width = Form6.Width - 80
    If Form6.Height > 2800 Then Grid1.Height = Form6.Height - 700
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Command4.FontBold = False
    Call Form_Deactivate
End Sub

Private Sub insrec_Click()
    Dim db As Database, ds As Recordset, sqlx As String
    Dim pkey As Long
    If Val(ekey) = 0 Then Exit Sub
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    
    sqlx = "select sequence_id from sequences where seq = 'Empskills'"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        pkey = ds(0) + 1
    Else
        pkey = 1
    End If
    sqlx = "Insert into empskills (id) values (" & pkey & ")"
    db.Execute sqlx
    sqlx = "select * from empskills where id = " & pkey
    Set ds = db.OpenRecordset(sqlx)
    ds.Edit
    ds!empkey = Val(ekey)
    ds!skillkey = Val(List1)
    ds.Update
    sqlx = "Update sequences set sequence_id = " & pkey & " where seq = 'Empskills'"
    db.Execute sqlx
    
    
    'sqlx = "select * from empskills where id = 0"
    'Set ds = db.OpenRecordset(sqlx)
    'ds.AddNew
    'ds!empkey = Val(ekey)
    'ds!skillkey = Val(List1)
    'pkey = ds!id
    'ds.Update
    
    ds.Close: db.Close
    Grid1.AddItem pkey & Chr(9) & Combo1
    Grid1.Row = Grid1.Rows - 1
End Sub

Private Sub skey_Change()
    Call refresh_skills
    'MsgBox "New skills"
End Sub

