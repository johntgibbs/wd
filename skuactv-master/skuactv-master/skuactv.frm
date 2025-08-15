VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SKU Activity"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Racks"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Cranes"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid frmgrid 
      Height          =   3495
      Left            =   4680
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6165
      _Version        =   327680
      Cols            =   5
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Orders"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Receipts"
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
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Locations"
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
      Left            =   2160
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Lots"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Totals"
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
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox bbsr 
      Height          =   285
      Left            =   4920
      TabIndex        =   1
      Text            =   "\\bbc-01-msg\sharedgroups\wd\data\bbsr.mdb"
      Top             =   6120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   4830
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "12.17.2021"
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
      Left            =   7080
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label pallogs 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   6600
      Width           =   4335
   End
   Begin VB.Label plantno 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7320
      TabIndex        =   11
      Top             =   480
      Width           =   855
   End
   Begin VB.Label plantdesc 
      Caption         =   "..."
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   5040
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public wdb As ADODB.Connection
Public db5 As ADODB.Connection

Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Command2_Click()
    Form3.Show
End Sub

Private Sub Command3_Click()
    Form4.Show
End Sub

Private Sub Command4_Click()
    Form6.Show
End Sub

Private Sub Command5_Click()
    Form5.Show
End Sub

Private Sub Form_Load()
    Dim ds As ADODB.Recordset, sqlx As String
    Dim f As String
    If UCase(Command()) = "BAUSER" Then
        Open "\\bbba-03-dc\f\user\waredist\bin\wd.ini" For Input As #1
    Else
        If UCase(Command()) = "SYUSER" Then
            Open "\\bbsy-02-dc\f\user\waredist\bin\wd.ini" For Input As #1
        Else
            'Open "wd.ini" For Input As #1
            'Open "\\bbsy-02-dc\f\user\waredist\bin\wd.ini" For Input As #1
            'Open "\\bbba-03-dc\f\user\waredist\bin\wd.ini" For Input As #1
            Open "\\bbc-01-prodtrk\wd\bin\wd.ini" For Input As #1
        End If
    End If
    Line Input #1, f
    Do Until EOF(1)
        Line Input #1, f
        f = LCase(f): f = Trim(f)
        If Left(f, 5) = "bbsr=" Then bbsr = Right(f, Len(f) - 5)
        If Left(f, 6) = "plant=" Then plantdesc.Caption = Right(f, Len(f) - 6)
        If Left(f, 8) = "plantno=" Then plantno.Caption = Right(f, 2)
        If f = "cranes=yes" Then Check1.Value = 1
        If f = "racks=yes" Then Check2.Value = 1
        If Left(f, 8) = "pallogs=" Then pallogs.Caption = Right(f, Len(f) - 8)
    Loop
    Close #1
    Form1.Caption = Form1.Caption & " " & plantdesc
    frmgrid.FormatString = "^Form|^Top|^Left|^Height|^Width"
    frmgrid.ColWidth(0) = 800: frmgrid.ColWidth(1) = 800
    frmgrid.ColWidth(2) = 800: frmgrid.ColWidth(3) = 800
    frmgrid.ColWidth(4) = 800
    frmgrid.Rows = 1
    On Error Resume Next
    Open "c:\windows\skuactv.ini" For Input As #1
    If Err = 53 Then
        frmgrid.AddItem "form1" & Chr(9) & Form1.Top & Chr(9) & Form1.Left & Chr(9) & Form1.Height & Chr(9) & Form1.Width
        frmgrid.AddItem "form2" & Chr(9) & Form2.Top & Chr(9) & Form2.Left & Chr(9) & Form2.Height & Chr(9) & Form2.Width
        frmgrid.AddItem "form3" & Chr(9) & Form3.Top & Chr(9) & Form3.Left & Chr(9) & Form3.Height & Chr(9) & Form3.Width
        frmgrid.AddItem "form4" & Chr(9) & Form4.Top & Chr(9) & Form4.Left & Chr(9) & Form4.Height & Chr(9) & Form4.Width
        frmgrid.AddItem "form5" & Chr(9) & Form5.Top & Chr(9) & Form5.Left & Chr(9) & Form5.Height & Chr(9) & Form5.Width
        frmgrid.AddItem "form6" & Chr(9) & Form6.Top & Chr(9) & Form6.Left & Chr(9) & Form6.Height & Chr(9) & Form6.Width
    Else
        Do Until EOF(1)
            Input #1, f, t, l, h, w
            frmgrid.AddItem f & Chr(9) & t & Chr(9) & l & Chr(9) & h & Chr(9) & w
        Loop
    End If
    Close #1
    On Error GoTo 0
    Form1.Top = Val(Form1.frmgrid.TextMatrix(1, 1))
    Form1.Left = Val(Form1.frmgrid.TextMatrix(1, 2))
    Form1.Height = Val(Form1.frmgrid.TextMatrix(1, 3))
    Form1.Width = Val(Form1.frmgrid.TextMatrix(1, 4))
    Set wdb = CreateObject("ADODB.Connection")                          'jv123114
    wdb.Open Me.bbsr                                                     'jv123114
    If Me.plantno = "52" Then
        Set db5 = CreateObject("ADODB.Connection")                          'jv123114
        'db5.Open "ODBC;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907;DSN=wdsqlcs5"
        'db5.Open "Driver={SQL Server};Server=bbsy-01-sqlsvr;DATABASE=BBC_WMS;UID=bbcwdcs5;PWD=bbclp1907" ' Dead database
        db5.Open "Driver={SQL Server};Server=BBSY-01-WESTFALIA;DATABASE=BlueBell_WMS;UID=bbcwdcs5;PWD=bbclp1907" ' New Westfalia database
    End If
    's = "ODBC;DATABASE=WDRacks;DSN=wdracks"
    's = "ODBC;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500;DSN=wdsql500"
    'Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, True, Form1.bbsr)
    sqlx = "select * from sku_config"
    'If Form1.plantno = "50" Then
        sqlx = sqlx & " where sku in (select sku from lane)"
        sqlx = sqlx & " or sku in (select sku from rackpos)"
    'End If
    sqlx = sqlx & " order by sku"
    'Set ds = db.OpenRecordset(sqlx)
    Set ds = wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!sku & "  " & StrConv(ds!Description, vbProperCase) & " "
            sqlx = sqlx & StrConv(ds!uom_type, vbProperCase)
            List1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close ': db.Close
    If List1.ListCount > 0 Then List1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    If Form1.Height > 2000 Then List1.Height = Form1.Height - 835
    If Form1.Width > 2000 Then List1.Width = Form1.Width - 100
End Sub

Private Sub Form_Terminate()
    Dim i As Integer, f As String
    Dim t As Long, l As Long, h As Long, w As Long
    wdb.Close
    If Me.plantno = "52" Then db5.Close
    If Form1.WindowState = 0 Then
        Form1.frmgrid.TextMatrix(1, 1) = Form1.Top
        Form1.frmgrid.TextMatrix(1, 2) = Form1.Left
        Form1.frmgrid.TextMatrix(1, 3) = Form1.Height
        Form1.frmgrid.TextMatrix(1, 4) = Form1.Width
    End If
    Open "c:\windows\skuactv.ini" For Output As #1
    For i = 1 To frmgrid.Rows - 1
        f = Form1.frmgrid.TextMatrix(i, 0)
        t = Val(Form1.frmgrid.TextMatrix(i, 1))
        l = Val(Form1.frmgrid.TextMatrix(i, 2))
        h = Val(Form1.frmgrid.TextMatrix(i, 3))
        w = Val(Form1.frmgrid.TextMatrix(i, 4))
        Write #1, f, t, l, h, w
    Next i
    Close #1
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Terminate
End Sub

Private Sub List1_Click()
    Form2.tsku = Trim(Left(List1, 4))                                       'jv082415
    Form2.tprod = Right(List1, Len(List1) - (Len(Form2.tsku) + 1))          'jv082415
    Form3.lsku = Trim(Left(List1, 4))                                       'jv082415
    Form3.lprod = Right(List1, Len(List1) - (Len(Form3.lsku) + 1))          'jv082415
    Form3.lwhs = "0"
    Form4.bsku = Trim(Left(List1, 4))                                       'jv082415
    Form4.bprod = Right(List1, Len(List1) - (Len(Form4.bsku) + 1))          'jv082415
    Form4.blot = "0"
    Form4.bwhs = "0"
    Form5.osku = Trim(Left(List1, 4))                                       'jv082415
    Form5.oprod = Right(List1, Len(List1) - (Len(Form5.osku) + 1))          'jv082415
    Form6.rsku = Trim(Left(List1, 4))                                       'jv082415
    Form6.rprod = Right(List1, Len(List1) - (Len(Form6.rsku) + 1))          'jv082415
End Sub
