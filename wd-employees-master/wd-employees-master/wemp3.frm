VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4335
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   5265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   4335
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1320
      TabIndex        =   28
      Text            =   "Text13"
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   4440
      TabIndex        =   27
      Text            =   "Text12"
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1320
      TabIndex        =   26
      Text            =   "Text11"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1320
      TabIndex        =   25
      Text            =   "Text10"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1320
      TabIndex        =   24
      Text            =   "Text9"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1320
      TabIndex        =   23
      Text            =   "Text8"
      Top             =   1920
      Width           =   3495
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ex"
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1320
      TabIndex        =   21
      Text            =   "Text7"
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Text            =   "Text6"
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Text            =   "Text5"
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Text            =   "Text4"
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Text            =   "Text3"
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Text            =   "Text2"
      Top             =   240
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   0
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   3240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1931
      _Version        =   327680
      FocusRect       =   0
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   6840
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label label1 
      Caption         =   "Anniversay:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   20
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label label1 
      Caption         =   "Zip Code:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label label1 
      Caption         =   "State:"
      Height          =   255
      Index           =   10
      Left            =   3960
      TabIndex        =   12
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label label1 
      Caption         =   "City:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label label1 
      Caption         =   "Street Address:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label label1 
      Caption         =   "Work Phone:"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label label1 
      Caption         =   "Employer:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label label1 
      Caption         =   "Birthday:"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label label1 
      Caption         =   "Nickname:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label label1 
      Caption         =   "Maiden Name:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label label1 
      Caption         =   "Last Name:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label label1 
      Caption         =   "Middle Name:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label label1 
      Caption         =   "1st Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label ekey 
      Caption         =   "ekey"
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   1560
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String, rflag As Boolean
Private Sub refresh_spouse()
    Dim db As Database, ds As Recordset, sqlx As String
    rflag = True
    Text1 = "": Text2 = "": Text3 = "": Text4 = ""
    Text5 = "": Text6 = "": Text7 = "": Text8 = ""
    Text9 = "": Text10 = "": Text11 = ""
    Text12 = "": Text13 = "": Check1 = 0
    Text1.Enabled = False: Text2.Enabled = False
    Text3.Enabled = False: Text4.Enabled = False
    Text5.Enabled = False: Text6.Enabled = False
    Text7.Enabled = False: Text8.Enabled = False
    Text9.Enabled = False: Text10.Enabled = False
    Text11.Enabled = False: Text12.Enabled = False
    Text13.Enabled = False: Check1.Enabled = False
    
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    Text1.Enabled = True: Text2.Enabled = True
    Text3.Enabled = True: Text4.Enabled = True
    Text5.Enabled = True: Text6.Enabled = True
    Text7.Enabled = True: Text8.Enabled = True
    Text9.Enabled = True: Text10.Enabled = True
    Text11.Enabled = True: Text12.Enabled = True
    Text13.Enabled = True: Check1.Enabled = True
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select * from spouses where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If ds!xflag = "Y" Then Check1 = 1                                       'jv081815
        If IsNull(ds!first_name) = False Then Text1 = ds!first_name
        If IsNull(ds!middle_name) = False Then Text2 = ds!middle_name
        If IsNull(ds!last_name) = False Then Text3 = ds!last_name
        If IsNull(ds!maiden_name) = False Then Text4 = ds!maiden_name
        If IsNull(ds!nickname) = False Then Text5 = ds!nickname
        If IsNull(ds!dob) = False Then Text6 = Format(ds!dob, "m-d-yyyy")
        If IsNull(ds!employer) = False Then Text8 = ds!employer
        If IsNull(ds!work_phone) = False Then Text9 = ds!work_phone
        If IsNull(ds!work_address) = False Then Text10 = ds!work_address
        If IsNull(ds!work_city) = False Then Text11 = ds!work_city
        If IsNull(ds!work_state) = False Then Text12 = ds!work_state
        If IsNull(ds!work_zip) = False Then Text13 = ds!work_zip
        If IsNull(ds!anniversary) = False Then Text7 = Format(ds!anniversary, "m-d-yyyy")
    End If
    ds.Close: db.Close
    DoEvents
    rflag = False
End Sub

Private Sub update_rec()
    Dim db As Database, ds As Recordset, sqlx As String
    If rflag = True Then Exit Sub
    If Len(edcell) = 0 Then Exit Sub
    'MsgBox "update rec " & edcell
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select * from spouses where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        ds.Edit
        If edcell = "first_name" Then ds!first_name = Text1
        If edcell = "middle_name" Then ds!middle_name = Text2
        If edcell = "last_name" Then ds!last_name = Text3
        Grid1.TextMatrix(Grid1.Row, 1) = Text1 & " " & Text3
        If edcell = "maiden_name" Then ds!maiden_name = Text4
        If edcell = "nickname" Then ds!nickname = Text5
        If edcell = "employer" Then ds!employer = Text8
        If edcell = "work_phone" Then ds!work_phone = Text9
        If edcell = "work_address" Then ds!work_address = Text10
        If edcell = "work_city" Then ds!work_city = Text11
        If edcell = "work_state" Then ds!work_state = Text12
        If edcell = "work_zip" Then ds!work_zip = Text13
        If edcell = "xflag" Then
            If Check1 = 1 Then
                ds!xflag = "Y"                                  'jv081815
            Else
                ds!xflag = "N"                                  'jv081815
            End If
        End If
        If edcell = "dob" Then
            If IsDate(Text6) Then
                ds!dob = Format(Text6, "m-d-yyyy")
            Else
                Beep
                Text6 = Format(ds!dob, "m-d-yyyy")
            End If
        End If
        If edcell = "anniversary" Then
            If IsDate(Text7) Then
                ds!anniversary = Format(Text7, "m-d-yyyy")
            Else
                Beep
                Text7 = Format(ds!anniversary, "m-d-yyyy")
            End If
        End If
        ds.Update
    End If
    ds.Close: db.Close
    edcell = ""
End Sub

Private Sub refresh_grid()
    Dim db As Database, ds As Recordset, sqlx As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 2
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select id, first_name, last_name from spouses"
    sqlx = sqlx & " where empkey = " & Val(ekey)
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!id & Chr(9) & ds!first_name & " " & ds!last_name
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    Grid1.FormatString = "^ID|<Spouse"
    Grid1.ColWidth(0) = 2: Grid1.ColWidth(1) = 5000
    If Grid1.Rows > 1 Then Grid1.Row = 1
    Call Grid1_RowColChange
End Sub

Private Sub Check1_Click()
    If rflag = False Then
        If Len(edcell) > 0 Then DoEvents
        edcell = "xflag"
        Call update_rec
        DoEvents
        Text8.SetFocus
    End If
End Sub

Private Sub delrec_Click()
    Dim db As Database, ds As Recordset, sqlx As String, s As String
    If Grid1.Row = 0 Then Exit Sub
    sqlx = "Ok to delete spouse record for " & Grid1.TextMatrix(Grid1.Row, 1)
    If MsgBox(sqlx, vbYesNo + vbQuestion, "Delete Spouse Record...") = vbNo Then Exit Sub
    sqlx = "select * from spouses where id = " & Grid1.TextMatrix(Grid1.Row, 0)
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
            Call update_rec
        Else
            edcell = ""
        End If
    End If
    If Form3.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form3" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form3.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form3.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form3.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form3.Width
                x = 3
                Exit For
            End If
        Next i
        If x <> 3 Then Form1.frmgrid.AddItem "form3" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 5025 & Chr(9) & 5385
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
        If Form1.frmgrid.TextMatrix(i, 0) = "form3" Then
            Form3.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form3.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form3.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form3.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Command1.FontBold = False
    Call Form_Deactivate
End Sub

Private Sub Grid1_RowColChange()
    Call refresh_spouse
End Sub
Private Sub insrec_Click()
    Dim db As Database, ds As Recordset, sqlx As String
    Dim plast As String, pfirst As String, pkey As Long
    If Val(ekey) = 0 Then Exit Sub
    pfirst = InputBox("First Name: ", "First Name.....")
    If Len(pfirst) = 0 Then Exit Sub
    plast = InputBox("Last Name: ", "Last Name......")
    If Len(plast) = 0 Then Exit Sub
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    
    sqlx = "select sequence_id from sequences where seq = 'Spouses'"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        pkey = ds(0) + 1
    Else
        pkey = 1
    End If
    sqlx = "Insert into spouses (id) values (" & pkey & ")"
    db.Execute sqlx
    sqlx = "select * from spouses where id = " & pkey
    Set ds = db.OpenRecordset(sqlx)
    ds.Edit
    ds!empkey = Val(ekey)
    ds!first_name = pfirst
    ds!last_name = plast
    ds.Update
    sqlx = "Update sequences set sequence_id = " & pkey & " where seq = 'Spouses'"
    db.Execute sqlx
    
    
    
    
    'sqlx = "select * from spouses where id = 0"
    'Set ds = db.OpenRecordset(sqlx)
    'ds.AddNew
    'ds!empkey = Val(ekey)
    'ds!first_name = pfirst
    'ds!last_name = plast
    'pkey = ds!id
    'ds.Update
    
    
    ds.Close: db.Close
    Grid1.AddItem pkey & Chr(9) & pfirst & " " & plast
    Grid1.Row = Grid1.Rows - 1
    Call refresh_spouse
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0: Text2.SetFocus
    Else
        edcell = "first_name"
    End If
End Sub

Private Sub Text1_LostFocus()
    If edcell = "first_name" Then Call update_rec
End Sub

Private Sub Text10_GotFocus()
    Text10.SelStart = 0
    Text10.SelLength = Len(Text10)
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text11.SetFocus
    Else
        edcell = "work_address"
    End If
End Sub

Private Sub Text10_LostFocus()
    If edcell = "work_address" Then Call update_rec
End Sub
Private Sub Text11_GotFocus()
    Text11.SelStart = 0
    Text11.SelLength = Len(Text11)
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text12.SetFocus
    Else
        edcell = "work_city"
    End If
End Sub

Private Sub Text11_LostFocus()
    If edcell = "work_city" Then Call update_rec
End Sub

Private Sub Text12_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Len(Text12)
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text13.SetFocus
    Else
        edcell = "work_state"
    End If
End Sub

Private Sub Text12_LostFocus()
    If edcell = "work_state" Then Call update_rec
End Sub

Private Sub Text13_GotFocus()
    Text13.SelStart = 0
    Text13.SelLength = Len(Text13)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text1.SetFocus
    Else
        edcell = "work_zip"
    End If
End Sub

Private Sub Text13_LostFocus()
    If edcell = "work_zip" Then Call update_rec
End Sub

Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text3.SetFocus
    Else
        edcell = "middle_name"
    End If
End Sub

Private Sub Text2_LostFocus()
    If edcell = "middle_name" Then Call update_rec
End Sub

Private Sub Text3_GotFocus()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text4.SetFocus
    Else
        edcell = "last_name"
    End If
End Sub

Private Sub Text3_LostFocus()
    If edcell = "last_name" Then Call update_rec
End Sub

Private Sub Text4_GotFocus()
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text5.SetFocus
    Else
        edcell = "maiden_name"
    End If
End Sub

Private Sub Text4_LostFocus()
    If edcell = "maiden_name" Then Call update_rec
End Sub

Private Sub Text5_GotFocus()
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text6.SetFocus
    Else
        edcell = "nickname"
    End If
End Sub

Private Sub Text5_LostFocus()
    If edcell = "nickname" Then Call update_rec
End Sub

Private Sub Text6_GotFocus()
    Text6.SelStart = 0
    Text6.SelLength = Len(Text6)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text7.SetFocus
    Else
        edcell = "dob"
    End If
End Sub

Private Sub Text6_LostFocus()
    If edcell = "dob" Then Call update_rec
End Sub

Private Sub Text7_GotFocus()
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text8.SetFocus
    Else
        edcell = "anniversary"
    End If
End Sub

Private Sub Text7_LostFocus()
    If edcell = "anniversary" Then Call update_rec
End Sub

Private Sub Text8_GotFocus()
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text9.SetFocus
    Else
        edcell = "employer"
    End If
End Sub

Private Sub Text8_LostFocus()
    If edcell = "employer" Then Call update_rec
End Sub

Private Sub Text9_GotFocus()
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9)
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text10.SetFocus
    Else
        edcell = "work_phone"
    End If
End Sub

Private Sub Text9_LostFocus()
    If edcell = "work_phone" Then Call update_rec
End Sub

