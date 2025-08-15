VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "W/D Employee List"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   7965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   6360
      TabIndex        =   69
      Text            =   "Text25"
      Top             =   5400
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid frmgrid 
      Height          =   615
      Left            =   120
      TabIndex        =   65
      Top             =   6120
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1085
      _Version        =   327680
      Cols            =   5
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   6240
      TabIndex        =   64
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   6360
      MaxLength       =   15
      TabIndex        =   30
      Text            =   "Text24"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox empdb 
      Height          =   285
      Left            =   120
      TabIndex        =   62
      Text            =   "c:\wdemp.mdb"
      Top             =   7200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   6360
      TabIndex        =   61
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4320
      TabIndex        =   27
      Text            =   "Combo2"
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Skills"
      Height          =   375
      Left            =   4680
      TabIndex        =   59
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Emergency Contacts"
      Height          =   375
      Left            =   2760
      TabIndex        =   58
      Top             =   5760
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Children"
      Height          =   375
      Left            =   1440
      TabIndex        =   57
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Spouse"
      Height          =   375
      Left            =   120
      TabIndex        =   56
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   4320
      MaxLength       =   5
      TabIndex        =   29
      Text            =   "Text23"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   28
      Text            =   "Text22"
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CheckBox Check5 
      Alignment       =   1  'Right Justify
      Caption         =   "Parent:"
      Height          =   255
      Left            =   1920
      TabIndex        =   26
      Top             =   4680
      Width           =   975
   End
   Begin VB.CheckBox Check4 
      Alignment       =   1  'Right Justify
      Caption         =   "Married:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   3480
      MaxLength       =   50
      TabIndex        =   24
      Text            =   "Text21"
      Top             =   4320
      Width           =   4335
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   1320
      TabIndex        =   23
      Text            =   "Text20"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   5160
      TabIndex        =   22
      Text            =   "Text19"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   1320
      TabIndex        =   20
      Text            =   "Text18"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Caption         =   "Full Time:"
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Vietnam Era:"
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   18
      Text            =   "Text17"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Veteran:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   4320
      MaxLength       =   20
      TabIndex        =   16
      Text            =   "Text16"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   6360
      MaxLength       =   12
      TabIndex        =   14
      Text            =   "Text15"
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   13
      Text            =   "Text14"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   12
      Text            =   "Text13"
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   11
      Text            =   "Text12"
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   15
      Text            =   "Text11"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   6360
      MaxLength       =   15
      TabIndex        =   10
      Text            =   "Text10"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   9
      Text            =   "Text9"
      Top             =   2280
      Width           =   3975
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Text            =   "Text8"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   6360
      MaxLength       =   15
      TabIndex        =   8
      Text            =   "Text7"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   7
      Text            =   "Text6"
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Cell Phone:"
      Height          =   255
      Left            =   5280
      TabIndex        =   68
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label tcolor 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Terminated"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3840
      TabIndex        =   67
      Top             =   600
      Width           =   975
   End
   Begin VB.Label ncolor 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Active"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3840
      TabIndex        =   66
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Work Phone:"
      Height          =   255
      Index           =   24
      Left            =   5280
      TabIndex        =   63
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lastlit 
      Caption         =   "Last Modified:"
      Height          =   255
      Left            =   3000
      TabIndex        =   60
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Radio Code:"
      Height          =   255
      Index           =   23
      Left            =   3240
      TabIndex        =   55
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "BB Employee Number:"
      Height          =   255
      Index           =   22
      Left            =   120
      TabIndex        =   54
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Department:"
      Height          =   255
      Index           =   21
      Left            =   3240
      TabIndex        =   53
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Reason:"
      Height          =   255
      Index           =   20
      Left            =   2760
      TabIndex        =   52
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Date Termed:"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   51
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Date Full Time:"
      Height          =   255
      Index           =   18
      Left            =   3960
      TabIndex        =   50
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Date Employed:"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   49
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Years:"
      Height          =   255
      Index           =   16
      Left            =   1800
      TabIndex        =   48
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "County:"
      Height          =   255
      Index           =   15
      Left            =   3720
      TabIndex        =   47
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Zip Code:"
      Height          =   255
      Index           =   14
      Left            =   5400
      TabIndex        =   46
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "State:"
      Height          =   255
      Index           =   13
      Left            =   3840
      TabIndex        =   45
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "City:"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   44
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Street:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   43
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nickname:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   42
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Maiden Name:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   41
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Home Phone:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   40
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "DL Number:"
      Height          =   255
      Index           =   7
      Left            =   5400
      TabIndex        =   39
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "DL Name:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   38
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Date of Birth:"
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   37
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "SS Number:"
      Height          =   255
      Index           =   4
      Left            =   5400
      TabIndex        =   36
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "SS Name:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   35
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   34
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Middle:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   33
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "1st Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   600
      Width           =   855
   End
   Begin VB.Label crt 
      Caption         =   "crt"
      Height          =   255
      Left            =   6120
      TabIndex        =   31
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu filemenu 
      Caption         =   "&File"
      Begin VB.Menu xitmenu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edmenu 
      Caption         =   "&Edit"
      Begin VB.Menu insrec 
         Caption         =   "Insert - F10"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete - F9"
      End
   End
   Begin VB.Menu calmenu 
      Caption         =   "Calendar"
   End
   Begin VB.Menu offday 
      Caption         =   "Off Days"
   End
   Begin VB.Menu repmenu 
      Caption         =   "Reports"
      Begin VB.Menu repgrid 
         Caption         =   "Employee Query"
         Begin VB.Menu qact 
            Caption         =   "Active Employees"
            Begin VB.Menu qactv 
               Caption         =   "All Active Employees"
            End
            Begin VB.Menu qdepts 
               Caption         =   "Departments"
               Begin VB.Menu qdept 
                  Caption         =   "0"
                  Index           =   0
               End
               Begin VB.Menu qdept 
                  Caption         =   "1"
                  Index           =   1
                  Visible         =   0   'False
               End
               Begin VB.Menu qdept 
                  Caption         =   "2"
                  Index           =   2
                  Visible         =   0   'False
               End
               Begin VB.Menu qdept 
                  Caption         =   "3"
                  Index           =   3
                  Visible         =   0   'False
               End
               Begin VB.Menu qdept 
                  Caption         =   "4"
                  Index           =   4
                  Visible         =   0   'False
               End
               Begin VB.Menu qdept 
                  Caption         =   "5"
                  Index           =   5
                  Visible         =   0   'False
               End
               Begin VB.Menu qdept 
                  Caption         =   "6"
                  Index           =   6
                  Visible         =   0   'False
               End
               Begin VB.Menu qdept 
                  Caption         =   "7"
                  Index           =   7
                  Visible         =   0   'False
               End
               Begin VB.Menu qdept 
                  Caption         =   "8"
                  Index           =   8
                  Visible         =   0   'False
               End
               Begin VB.Menu qdept 
                  Caption         =   "9"
                  Index           =   9
                  Visible         =   0   'False
               End
            End
            Begin VB.Menu qdbdays 
               Caption         =   "Birthdays"
               Begin VB.Menu qbd 
                  Caption         =   "January"
                  Index           =   1
               End
               Begin VB.Menu qbd 
                  Caption         =   "February"
                  Index           =   2
               End
               Begin VB.Menu qbd 
                  Caption         =   "March"
                  Index           =   3
               End
               Begin VB.Menu qbd 
                  Caption         =   "April"
                  Index           =   4
               End
               Begin VB.Menu qbd 
                  Caption         =   "May"
                  Index           =   5
               End
               Begin VB.Menu qbd 
                  Caption         =   "June"
                  Index           =   6
               End
               Begin VB.Menu qbd 
                  Caption         =   "July"
                  Index           =   7
               End
               Begin VB.Menu qbd 
                  Caption         =   "August"
                  Index           =   8
               End
               Begin VB.Menu qbd 
                  Caption         =   "September"
                  Index           =   9
               End
               Begin VB.Menu qbd 
                  Caption         =   "October"
                  Index           =   10
               End
               Begin VB.Menu qbd 
                  Caption         =   "November"
                  Index           =   11
               End
               Begin VB.Menu qbd 
                  Caption         =   "December"
                  Index           =   12
               End
            End
            Begin VB.Menu qdoes 
               Caption         =   "Date of Employment Anniversaries"
               Begin VB.Menu qdoe 
                  Caption         =   "January"
                  Index           =   1
               End
               Begin VB.Menu qdoe 
                  Caption         =   "February"
                  Index           =   2
               End
               Begin VB.Menu qdoe 
                  Caption         =   "March"
                  Index           =   3
               End
               Begin VB.Menu qdoe 
                  Caption         =   "April"
                  Index           =   4
               End
               Begin VB.Menu qdoe 
                  Caption         =   "May"
                  Index           =   5
               End
               Begin VB.Menu qdoe 
                  Caption         =   "June"
                  Index           =   6
               End
               Begin VB.Menu qdoe 
                  Caption         =   "July"
                  Index           =   7
               End
               Begin VB.Menu qdoe 
                  Caption         =   "August"
                  Index           =   8
               End
               Begin VB.Menu qdoe 
                  Caption         =   "September"
                  Index           =   9
               End
               Begin VB.Menu qdoe 
                  Caption         =   "October"
                  Index           =   10
               End
               Begin VB.Menu qdoe 
                  Caption         =   "November"
                  Index           =   11
               End
               Begin VB.Menu qdoe 
                  Caption         =   "December"
                  Index           =   12
               End
            End
            Begin VB.Menu qdfull 
               Caption         =   "Full Time Employee Anniversaries"
               Begin VB.Menu qdf 
                  Caption         =   "January"
                  Index           =   1
               End
               Begin VB.Menu qdf 
                  Caption         =   "February"
                  Index           =   2
               End
               Begin VB.Menu qdf 
                  Caption         =   "March"
                  Index           =   3
               End
               Begin VB.Menu qdf 
                  Caption         =   "April"
                  Index           =   4
               End
               Begin VB.Menu qdf 
                  Caption         =   "May"
                  Index           =   5
               End
               Begin VB.Menu qdf 
                  Caption         =   "June"
                  Index           =   6
               End
               Begin VB.Menu qdf 
                  Caption         =   "July"
                  Index           =   7
               End
               Begin VB.Menu qdf 
                  Caption         =   "August"
                  Index           =   8
               End
               Begin VB.Menu qdf 
                  Caption         =   "September"
                  Index           =   9
               End
               Begin VB.Menu qdf 
                  Caption         =   "October"
                  Index           =   10
               End
               Begin VB.Menu qdf 
                  Caption         =   "November"
                  Index           =   11
               End
               Begin VB.Menu qdf 
                  Caption         =   "December"
                  Index           =   12
               End
            End
            Begin VB.Menu qvets 
               Caption         =   "Veterans"
            End
            Begin VB.Menu qms 
               Caption         =   "Marital Status"
               Begin VB.Menu qsingle 
                  Caption         =   "Single"
               End
               Begin VB.Menu qmarried 
                  Caption         =   "Married"
               End
            End
            Begin VB.Menu qparent 
               Caption         =   "Parent"
            End
            Begin VB.Menu qradio 
               Caption         =   "Radio Codes"
            End
            Begin VB.Menu qcells 
               Caption         =   "Cell Phone Numbers"
            End
         End
         Begin VB.Menu qterm 
            Caption         =   "Terminated Employees"
         End
         Begin VB.Menu qall 
            Caption         =   "All Employees"
         End
      End
   End
   Begin VB.Menu confmenu 
      Caption         =   "Configure"
      Begin VB.Menu skilltype 
         Caption         =   "Skill Types"
      End
      Begin VB.Menu wddepts 
         Caption         =   "W/D Departments"
      End
      Begin VB.Menu vallists 
         Caption         =   "Value Lists"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String, rflag As Boolean

Private Function check_userid(uname As String) As Boolean
    Dim db As Database, ds As Recordset, sqlx As String, cflag As Boolean
    cflag = False
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select * from valuelists where listname = 'wdempuser' and listreturn = '" & uname & "'"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        cflag = True
    End If
    ds.Close: db.Close
    check_userid = cflag
End Function

Public Function fixquotes(s As String) As String
    Dim i As Integer, k As Integer, rs As String
    rs = ""
    i = 1: k = 1
    Do Until i = 0
        i = InStr(k, s, "'")
        If i = 0 Then
            rs = rs & Mid(s, k, Len(s))
            Exit Do
        Else
            rs = rs & Mid(s, k, i - k) & "''"
            k = i + 1
        End If
    Loop
    fixquotes = rs
End Function

Private Sub update_value_lists(bbnum As String, ename As String)
    Dim db As Database, ds As Recordset, ss As Recordset, s As String, k As Long
    'MsgBox bbnum & ".." & ename
    If Val(bbnum) = 0 Then Exit Sub
    If Len(bbnum) <> 6 Then Exit Sub
    If Len(ename) < 1 Then Exit Sub
    's = "ODBC;DATABASE=WDship;UID=bbcship500;PWD=brenham500;DSN=wdship500"
    s = "Driver={SQL Server};Server=BBC-08-SQLSVR;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, s)
    s = "select * from valuelists where listname = 'wdempid'"
    s = s & " and listreturn = '" & bbnum & "'"
    Set ds = db.OpenRecordset(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update valuelists set listdisplay = '" & fixquotes(ename) & "' where id = " & ds!id
        'MsgBox s
        db.Execute s
    Else
        s = "select sequence_id from sequences where seq = 'valuelists'"
        Set ss = db.OpenRecordset(s)
        If ss.BOF = False Then
            ss.MoveFirst
            k = ss(0)
        Else
            k = 0
        End If
        ss.Close
        k = k + 1
        s = "Insert into valuelists (id, listname, listreturn, listdisplay) Values ("
        s = s & k & ", 'wdempid', '" & bbnum & "', '" & fixquotes(ename) & "')"
        'MsgBox s
        db.Execute s
        s = "Update sequences set sequence_id = " & k & " Where seq = 'valuelists'"
        'MsgBox s
        db.Execute s
    End If
    ds.Close: db.Close
End Sub

Private Sub fetch_employee()
    Dim db As Database, ds As Recordset, sqlx As String
    rflag = True
    Text1 = "": Text7 = "": Text13 = "": Text19 = ""
    Text2 = "": Text8 = "": Text14 = "": Text20 = ""
    Text3 = "": Text9 = "": Text15 = "": Text21 = ""
    Text4 = "": Text10 = "": Text16 = "": Text22 = ""
    Text5 = "": Text11 = "": Text17 = "": Text23 = ""
    Text6 = "": Text12 = "": Text18 = "": Text24 = ""
    Text25 = ""
    Check1.Value = 0: Check2.Value = 0: Check3.Value = 0
    Check4.Value = 0: Check5.Value = 0
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select * from employees where id = " & List1
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If IsNull(ds!bb_num) = False Then Text22 = ds!bb_num
        If IsNull(ds!ss_num) = False Then Text7 = ds!ss_num
        If IsNull(ds!first_name) = False Then Text1 = ds!first_name
        If IsNull(ds!middle_name) = False Then Text2 = ds!middle_name
        If IsNull(ds!last_name) = False Then Text3 = ds!last_name
        If IsNull(ds!maiden_name) = False Then Text4 = ds!maiden_name
        If IsNull(ds!nickname) = False Then Text5 = ds!nickname
        If IsNull(ds!ss_name) = False Then Text6 = ds!ss_name
        If IsNull(ds!dl_name) = False Then Text9 = ds!dl_name
        If IsNull(ds!dl_num) = False Then Text10 = ds!dl_num
        If IsNull(ds!home_phone) = False Then Text11 = ds!home_phone
        If IsNull(ds!work_phone) = False Then Text24 = ds!work_phone
        If IsNull(ds!cellphone) = False Then Text25 = ds!cellphone
        If IsNull(ds!street) = False Then Text12 = ds!street
        If IsNull(ds!city) = False Then Text13 = ds!city
        If IsNull(ds!State) = False Then Text14 = ds!State
        If IsNull(ds!zipcode) = False Then Text15 = ds!zipcode
        If IsNull(ds!county) = False Then Text16 = ds!county
        If ds!veteran = "Y" Then Check1 = 1                                 'jv081815
        If IsNull(ds!vet_years) = False Then Text17 = ds!vet_years
        If ds!vietvet = "Y" Then Check2 = 1                                 'jv081815
        If IsNull(ds!dob) = False Then Text8 = Format(ds!dob, "m-d-yyyy")
        If IsNull(ds!doe) = False Then Text18 = Format(ds!doe, "m-d-yyyy")
        If ds!fulltime = "Y" Then Check3 = 1                                'jv081815
        If IsNull(ds!dfulltime) = False Then Text19 = Format(ds!dfulltime, "m-d-yyyy")
        If IsNull(ds!dot) = False Then Text20 = Format(ds!dot, "m-d-yyyy")
        If IsNull(ds!termreason) = False Then Text21 = ds!termreason
        If ds!married = "Y" Then Check4 = 1                                 'jv081815
        If ds!Parent = "Y" Then Check5 = 1                                  'jv081815
        If IsNull(ds!radiocode) = False Then Text23 = ds!radiocode
        If IsNull(ds!lastmod) = False Then lastlit = "Last modified: " & Format(ds!lastmod, "m-d-yyyy h:mm am/pm") & " "
        If IsNull(ds!crt) = False Then lastlit = lastlit & " By " & ds!crt & "."
        For i = 0 To List2.ListCount - 1
            If List2.List(i) = ds!deptcode Then
                Combo2.ListIndex = i
                Exit For
            End If
        Next i
    End If
    ds.Close: db.Close
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
    sqlx = "select * from employees where id = " & List1
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        ds.Edit
        If edcell = "bb_num" Then ds!bb_num = Text22
        If edcell = "ss_num" Then ds!ss_num = Text7
        If edcell = "first_name" Then ds!first_name = Text1
        If edcell = "middle_name" Then ds!middle_name = Text2
        If edcell = "last_name" Then ds!last_name = Text3
        If edcell = "maiden_name" Then ds!maiden_name = Text4
        If edcell = "nickname" Then ds!nickname = Text5
        If edcell = "ss_name" Then ds!ss_name = Text6
        If edcell = "dl_name" Then ds!dl_name = Text9
        If edcell = "dl_num" Then ds!dl_num = Text10
        If edcell = "home_phone" Then ds!home_phone = Text11
        If edcell = "work_phone" Then ds!work_phone = Text24
        If edcell = "cellphone" Then ds!cellphone = Text25
        If edcell = "street" Then ds!street = Text12
        If edcell = "city" Then ds!city = Text13
        If edcell = "state" Then ds!State = Text14
        If edcell = "zipcode" Then ds!zipcode = Text15
        If edcell = "county" Then ds!county = Text16
        If edcell = "veteran" Then
            If Check1 = 1 Then
                ds!veteran = "Y"                                'jv081815
            Else
                ds!veteran = "N"                                'jv081815
            End If
        End If
        If edcell = "vet_years" Then ds!vet_years = Text17
        If edcell = "viet_vet" Then
            If Check2 = 1 Then
                ds!vietvet = "Y"                                'jv081815
            Else
                ds!vietvet = "N"                                'jv081815
            End If
        End If
        If edcell = "dob" Then
            If IsDate(Text8) Then
                ds!dob = Format(Text8, "m-d-yyyy")
            Else
                Beep
                Text8 = Format(ds!dob, "m-d-yyyy")
            End If
        End If
        If edcell = "doe" Then
            If IsDate(Text18) Then
                ds!doe = Format(Text18, "m-d-yyyy")
            Else
                Beep
                Text18 = Format(ds!doe, "m-d-yyyy")
            End If
        End If
        If edcell = "fulltime" Then
            If Check3 = 1 Then
                ds!fulltime = "Y"                               'jv081815
            Else
                ds!fulltime = "N"                               'jv081815
            End If
        End If
        If edcell = "dfulltime" Then
            If IsDate(Text19) Then
                ds!dfulltime = Format(Text19, "m-d-yyyy")
            Else
                Beep
                Text19 = Format(ds!dfulltime, "m-d-yyyy")
            End If
        End If
        If edcell = "dot" Then
            If IsDate(Text20) Then
                ds!dot = Format(Text20, "m-d-yyyy")
            Else
                Beep
                Text20 = Format(ds!dot, "m-d-yyyy")
            End If
        End If
        If edcell = "termreason" Then ds!termreason = Text21
        If edcell = "married" Then
            If Check4 = 1 Then
                ds!married = "Y"                                'jv081815
            Else
                ds!married = "N"                                'jv081815
            End If
        End If
        If edcell = "parent" Then
            If Check5 = 1 Then
                ds!Parent = "Y"                                 'jv081815
            Else
                ds!Parent = "N"                                 'jv081815
            End If
        End If
        If edcell = "deptcode" Then ds!deptcode = Val(List2)
        If edcell = "radiocode" Then ds!radiocode = Text23
        ds!lastmod = Format(Now, "m-d-yyyy h:mm am/pm")
        ds!crt = crt.Caption
        ds.Update
    End If
    ds.Close: db.Close
    If Combo2 > " " And Combo2 <> "Transport Driver" Then
        If edcell = "bb_num" Or edcell = "first_name" Or edcell = "last_name" Or edcell = "deptcode" Then
            Call update_value_lists(Text22, Text1 & " " & Text3)
        End If
    End If
    edcell = ""
End Sub

Private Sub calmenu_Click()
    'Form7.Show
End Sub

Private Sub Check1_Click()
    If rflag = False Then
        If Len(edcell) > 0 Then DoEvents
        edcell = "veteran"
        Call update_rec
        DoEvents
        Text17.SetFocus
    End If
End Sub

Private Sub Check2_Click()
    If rflag = False Then
        If Len(edcell) > 0 Then DoEvents
        edcell = "viet_vet"
        Call update_rec
        DoEvents
        Text18.SetFocus
    End If
End Sub

Private Sub Check3_Click()
    If rflag = False Then
        If Len(edcell) > 0 Then DoEvents
        edcell = "fulltime"
        Call update_rec
        DoEvents
        Text19.SetFocus
    End If
End Sub

Private Sub Check4_Click()
    If rflag = False Then
        If Len(edcell) > 0 Then DoEvents
        edcell = "married"
        Call update_rec
        DoEvents
        Check5.SetFocus
    End If
End Sub

Private Sub Check5_Click()
    If rflag = False Then
        If Len(edcell) > 0 Then DoEvents
        edcell = "parent"
        Call update_rec
        DoEvents
        Combo2.SetFocus
    End If
End Sub

Private Sub Combo1_Click()
    If Len(edcell) > 0 Then
        Call update_rec
        DoEvents
    End If
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Combo2_Click()
    If rflag = False Then
        If Len(edcell) > 0 Then
            Call update_rec
            DoEvents
        End If
        List2.ListIndex = Combo2.ListIndex
        DoEvents
        'Text22.SetFocus
    End If
End Sub

Private Sub Command1_Click()
    Command1.FontBold = True
    Form3.ekey = List1
    Form3.Caption = "Spouse Information: " & Text1 & " " & Text3
    Form3.Show
End Sub

Private Sub Command2_Click()
    Command2.FontBold = True
    Form4.ekey = List1
    Form4.Caption = "Children: " & Text1 & " " & Text3
    Form4.Show
End Sub

Private Sub Command3_Click()
    Command3.FontBold = True
    Form5.ekey = List1
    Form5.Caption = "Emergency Contacts: " & Text1 & " " & Text3
    Form5.Show
End Sub

Private Sub Command4_Click()
    Command4.FontBold = True
    Form6.ekey = List1
    Form6.Caption = "Skills: " & Text1 & " " & Text3
    Form6.Show
End Sub

Private Sub delrec_Click()
    Dim db As Database, ds As Recordset, sqlx As String
    sqlx = "Ok to delete this employee record?"
    If List1.ListCount = 0 Then Exit Sub
    If MsgBox(sqlx, vbYesNo + vbQuestion, "Are you sure....") = vbNo Then Exit Sub
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select * from children where empkey = " & List1
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            ds.Delete
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "select * from econtacts where empkey = " & List1
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            ds.Delete
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "select * from empskills where empkey = " & List1
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            ds.Delete
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "select * from spouses where empkey = " & List1
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            ds.Delete
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "select * from employees where id = " & List1
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        ds.Delete
    End If
    ds.Close: db.Close
    DoEvents
    i = List1.ListIndex
    List1.RemoveItem i
    Combo1.RemoveItem i
    If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 121 Then
        KeyCode = 0
        Call insrec_Click     'F10
    End If
    If KeyCode = 120 Then Call delrec_Click     'F9
End Sub

Private Sub Form_Load()
    Dim lpbuff As String * 25, ret As Long
    Dim f As String, t As String, l As String
    Dim h As String, w As String, i As Integer
    Dim db As Database, ds As Recordset, sqlx As String
    check_hax
    ret = GetUserName(lpbuff, 25)
    crt.Caption = Left(lpbuff, InStr(lpbuff, Chr(0)) - 1)
    If check_userid(crt.Caption) = False Then
    'If check_userid("wduser") = False Then
        MsgBox "Who are you?", vbOKOnly + vbQuestion, "Access denied....."
        End
    End If
    frmgrid.Rows = 1
    If Len(Dir("c:\windows\wdemp.ini")) > 0 Then
        Open "c:\windows\wdemp.ini" For Input As #1
        Do Until EOF(1)
            Input #1, f, t, l, h, w
            frmgrid.AddItem f & Chr(9) & t & Chr(9) & l & Chr(9) & h & Chr(9) & w
        Loop
        Close #1
    Else
        frmgrid.AddItem "form1" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 7035 & Chr(9) & 8085
        frmgrid.AddItem "form2" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 5040 & Chr(9) & 5400
        frmgrid.AddItem "form3" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 5025 & Chr(9) & 5385
        frmgrid.AddItem "form4" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 3735 & Chr(9) & 7530
        frmgrid.AddItem "form5" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 2610 & Chr(9) & 7335
        frmgrid.AddItem "form6" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 4725 & Chr(9) & 7680
        frmgrid.AddItem "form7" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 5985 & Chr(9) & 7095
        frmgrid.AddItem "form8" & Chr(9) & 105 & Chr(9) & 105 & Chr(9) & 5985 & Chr(9) & 7095
    End If
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form1" Then
            Form1.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form1.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form1.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form1.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    
    Combo1.Clear: List1.Clear
    Combo2.Clear: List2.Clear
    rflag = True
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select * from departments order by deptdesc"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo2.AddItem ds!deptdesc
            List2.AddItem ds!id
            ds.MoveNext
        Loop
        Combo2.ListIndex = 0
    End If
    ds.Close
    sqlx = "select id,last_name,first_name from employees"
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
    For i = 0 To List2.ListCount - 1
        If i < 10 Then
            qdept(i).Caption = Combo2.List(i)
            qdept(i).Visible = True
        End If
    Next i
    DoEvents
    rflag = False
End Sub

Private Sub Form_Terminate()
    Call xitmenu_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call xitmenu_Click
End Sub

Private Sub insrec_Click()
    Dim db As Database, ds As Recordset, sqlx As String
    Dim plast As String, pfirst As String, pkey As Long
    pfirst = InputBox("First Name: ", "First Name.....")
    If Len(pfirst) = 0 Then Exit Sub
    plast = InputBox("Last Name: ", "Last Name......")
    If Len(plast) = 0 Then Exit Sub
    sqlx = "Driver={SQL Server};Server=BBC-08-SQLSVR;database=wdemployees;uid=wdemployee500;pwd=brenham500;"
    Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, sqlx)
    'Set db = OpenDatabase(Form1.empdb)
    sqlx = "select sequence_id from sequences where seq = 'Employees'"
    Set ds = db.OpenRecordset(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        pkey = ds(0) + 1
    Else
        pkey = 1
    End If
    sqlx = "Insert into employees (id) values (" & pkey & ")"
    db.Execute sqlx
    sqlx = "select * from employees where id = " & pkey
    Set ds = db.OpenRecordset(sqlx)
    ds.Edit
    ds!first_name = pfirst
    ds!last_name = plast
    ds!State = "TX"
    ds!lastmod = Format(Now, "m-d-yyyy h:mm am/pm")
    ds!crt = crt.Caption
    ds!dot = " "
    ds!termreason = " "
    ds!married = "N"
    ds!Parent = "N"
    ds!radiocode = " "
    ds!ss_num = " "
    ds!maiden_name = " "
    ds!nickname = " "
    ds!work_phone = " "
    ds!veteran = "N"
    ds!vet_years = " "
    ds!vietvet = "N"
    ds!deptcode = 1
    ds.Update
    sqlx = "Update sequences set sequence_id = " & pkey & " where seq = 'Employees'"
    db.Execute sqlx
    
    'sqlx = "select * from employees where id = 0"
    'Set ds = db.OpenRecordset(sqlx)
    'ds.AddNew
    'ds!first_name = pfirst
    'ds!last_name = plast
    'ds!State = "TX"
    'ds!lastmod = Format(Now, "m-d-yyyy h:mm am/pm")
    'ds!crt = crt.Caption
    'pkey = ds!id
    'ds.Update
    
    ds.Close: db.Close
    Combo1.AddItem plast & ", " & pfirst
    List1.AddItem pkey
    Combo1.ListIndex = Combo1.ListCount - 1
End Sub

Private Sub List1_Click()
    Call fetch_employee
    If Command1.FontBold = True Then
        Form3.ekey = List1
        Form3.Caption = "Spouse Information: " & Text1 & " " & Text3
    End If
    If Command2.FontBold = True Then
        Form4.ekey = List1
        Form4.Caption = "Children: " & Text1 & " " & Text3
    End If
    If Command3.FontBold = True Then
        Form5.ekey = List1
        Form5.Caption = "Emergency Contacts: " & Text1 & " " & Text3
    End If
    If Command4.FontBold = True Then
        Form6.ekey = List1
        Form6.Caption = "Skills: " & Text1 & " " & Text3
    End If
End Sub

Private Sub List2_Click()
    If rflag = False Then
        edcell = "deptcode"
        Call update_rec
    End If
End Sub

Private Sub offday_Click()
    'Form9.Show
End Sub

Private Sub qactv_Click()
    Dim sqlx As String
    sqlx = "dot < '0'"                          'jv081815
    Form8.Caption = "Active Employees"
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
End Sub

Private Sub qall_Click()
    Form8.Caption = "Active & Terminated"
    Form8.qstr = ""
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
End Sub

Private Sub qbd_Click(Index As Integer)
    Dim sqlx As String
    sqlx = "dot < '0'"                          'jv081815
    sqlx = sqlx & " and Month(dob) = " & Index
    Form8.Caption = qbd(Index).Caption & " Birthdays"
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
    Call Form8.sf_Click(20)
End Sub

Private Sub qcells_Click()
    Dim sqlx As String
    sqlx = "dot < '0'"                          'jv081815
    sqlx = sqlx & " and cellphone > '00'"
    Form8.Caption = "Cell Phone Numbers"
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
    Call Form8.sf_Click(29)
End Sub

Private Sub qdept_Click(Index As Integer)
    Dim sqlx As String
    sqlx = "dot < '0'"                          'jv081815
    sqlx = sqlx & " and deptcode = " & List2.List(Index)
    Form8.Caption = qdept(Index).Caption
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
    Call Form8.sf_Click(30)
End Sub

Private Sub qdf_Click(Index As Integer)
    Dim sqlx As String
    sqlx = "dot < '0'"                          'jv081815
    sqlx = sqlx & " and Month(dfulltime) = " & Index
    Form8.Caption = qbd(Index).Caption & " Full Time Employee Anniversaries"
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
    Call Form8.sf_Click(23)
End Sub

Private Sub qdoe_Click(Index As Integer)
    Dim sqlx As String
    sqlx = "dot < '0'"                          'jv081815
    sqlx = sqlx & " and Month(doe) = " & Index
    Form8.Caption = qbd(Index).Caption & " Date of Employment Anniversaries"
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
    Call Form8.sf_Click(21)
End Sub

Private Sub qmarried_Click()
    Dim sqlx As String
    sqlx = "dot < '0'"                          'jv081815
    sqlx = sqlx & " and married = 'Y'"          'jv081815
    Form8.Caption = "Married Employees"
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
    Call Form8.sf_Click(26)
End Sub

Private Sub qparent_Click()
    Dim sqlx As String
    sqlx = "dot < '0'"                          'jv081815
    sqlx = sqlx & " and parent = 'Y'"           'jv081815
    Form8.Caption = "Parents"
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
    Call Form8.sf_Click(27)
End Sub

Private Sub qradio_Click()
    Dim sqlx As String
    sqlx = "dot < '0'"                          'jv081815
    sqlx = sqlx & " and radiocode > '00'"
    Form8.Caption = "Radio Codes"
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
    Call Form8.sf_Click(28)
End Sub

Private Sub qsingle_Click()
    Dim sqlx As String
    sqlx = "dot < '0'"                          'jv081815
    sqlx = sqlx & " and married = 'N'"          'jv081815
    Form8.Caption = "Bachelors & Bachelorettes"
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
    Call Form8.sf_Click(26)
End Sub

Private Sub qterm_Click()
    Dim sqlx As String
    sqlx = "dot >= '0'"                         'jv081815
    Form8.Caption = "Terminated Employees"
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
    Call Form8.sf_Click(24)
    Call Form8.sf_Click(25)
End Sub

Private Sub qvets_Click()
    Dim sqlx As String
    sqlx = "dot < '0'"                          'jv081815
    sqlx = sqlx & " and veteran = 'Y'"          'jv081815
    Form8.Caption = "Veterans"
    Form8.qstr = sqlx
    Form8.qtrig = Val(Form8.qtrig) + 1
    Form8.Show
    Call Form8.sf_Click(17)
    Call Form8.sf_Click(18)
    Call Form8.sf_Click(19)
End Sub

Private Sub repgrid_Click()
    'Form8.Show
End Sub

Private Sub skilltype_Click()
    Form2.ftype = "Skills"
    Form2.Show
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0: Text2.SetFocus
    Else
        'If KeyAscii > 64 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
        KeyAscii = 0: Text12.SetFocus
    Else
        edcell = "dl_num"
    End If
End Sub

Private Sub Text10_LostFocus()
    If edcell = "dl_num" Then Call update_rec
End Sub
Private Sub Text11_GotFocus()
    Text11.SelStart = 0
    Text11.SelLength = Len(Text11)
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text16.SetFocus
    Else
        edcell = "home_phone"
    End If
End Sub

Private Sub Text11_LostFocus()
    If edcell = "home_phone" Then Call update_rec
End Sub

Private Sub Text12_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Len(Text12)
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text13.SetFocus
    Else
        edcell = "street"
    End If
End Sub

Private Sub Text12_LostFocus()
    If edcell = "street" Then Call update_rec
End Sub

Private Sub Text13_GotFocus()
    Text13.SelStart = 0
    Text13.SelLength = Len(Text13)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text14.SetFocus
    Else
        edcell = "city"
    End If
End Sub

Private Sub Text13_LostFocus()
    If edcell = "city" Then Call update_rec
End Sub

Private Sub Text14_GotFocus()
    Text14.SelStart = 0
    Text14.SelLength = Len(Text14)
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text15.SetFocus
    Else
        edcell = "state"
    End If
End Sub

Private Sub Text14_LostFocus()
    If edcell = "state" Then Call update_rec
End Sub

Private Sub Text15_GotFocus()
    Text15.SelStart = 0
    Text15.SelLength = Len(Text15)
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text11.SetFocus
    Else
        edcell = "zipcode"
    End If
End Sub

Private Sub Text15_LostFocus()
    If edcell = "zipcode" Then Call update_rec
End Sub

Private Sub Text16_GotFocus()
    Text16.SelStart = 0
    Text16.SelLength = Len(Text16)
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Check1.SetFocus
    Else
        edcell = "county"
    End If
End Sub

Private Sub Text16_LostFocus()
    If edcell = "county" Then Call update_rec
End Sub

Private Sub Text17_GotFocus()
    Text17.SelStart = 0
    Text17.SelLength = Len(Text17)
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Check2.SetFocus
    Else
        edcell = "vet_years"
    End If
End Sub

Private Sub Text17_LostFocus()
    If edcell = "vet_years" Then Call update_rec
End Sub

Private Sub Text18_GotFocus()
    Text18.SelStart = 0
    Text18.SelLength = Len(Text18)
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Check3.SetFocus
    Else
        edcell = "doe"
    End If
End Sub

Private Sub Text18_LostFocus()
    If edcell = "doe" Then Call update_rec
End Sub

Private Sub Text19_GotFocus()
    Text19.SelStart = 0
    Text19.SelLength = Len(Text19)
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text20.SetFocus
    Else
        edcell = "dfulltime"
    End If
End Sub

Private Sub Text19_LostFocus()
    If edcell = "dfulltime" Then Call update_rec
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

Private Sub Text20_Change()
    If IsDate(Text20) Then
        Text1.ForeColor = tcolor.ForeColor
        Text1.BackColor = tcolor.BackColor
        Text2.ForeColor = tcolor.ForeColor
        Text2.BackColor = tcolor.BackColor
        Text3.ForeColor = tcolor.ForeColor
        Text3.BackColor = tcolor.BackColor
        Text4.ForeColor = tcolor.ForeColor
        Text4.BackColor = tcolor.BackColor
        Text5.ForeColor = tcolor.ForeColor
        Text5.BackColor = tcolor.BackColor
        Text8.ForeColor = tcolor.ForeColor
        Text8.BackColor = tcolor.BackColor
        tcolor.Visible = True: ncolor.Visible = False
    Else
        Text1.ForeColor = ncolor.ForeColor
        Text1.BackColor = ncolor.BackColor
        Text2.ForeColor = ncolor.ForeColor
        Text2.BackColor = ncolor.BackColor
        Text3.ForeColor = ncolor.ForeColor
        Text3.BackColor = ncolor.BackColor
        Text4.ForeColor = ncolor.ForeColor
        Text4.BackColor = ncolor.BackColor
        Text5.ForeColor = ncolor.ForeColor
        Text5.BackColor = ncolor.BackColor
        Text8.ForeColor = ncolor.ForeColor
        Text8.BackColor = ncolor.BackColor
        tcolor.Visible = False: ncolor.Visible = True
    End If
        
End Sub

Private Sub Text20_GotFocus()
    Text20.SelStart = 0
    Text20.SelLength = Len(Text20)
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text21.SetFocus
    Else
        edcell = "dot"
    End If
End Sub

Private Sub Text20_LostFocus()
    If edcell = "dot" Then Call update_rec
End Sub

Private Sub Text21_GotFocus()
    Text21.SelStart = 0
    Text21.SelLength = Len(Text21)
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Check4.SetFocus
    Else
        edcell = "termreason"
    End If
End Sub

Private Sub Text21_LostFocus()
    If edcell = "termreason" Then Call update_rec
End Sub

Private Sub Text22_GotFocus()
    Text22.SelStart = 0
    Text22.SelLength = Len(Text22)
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text23.SetFocus
    Else
        edcell = "bb_num"
    End If
End Sub

Private Sub Text22_LostFocus()
    If edcell = "bb_num" Then Call update_rec
End Sub

Private Sub Text23_GotFocus()
    Text23.SelStart = 0
    Text23.SelLength = Len(Text23)
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text24.SetFocus
    Else
        edcell = "radiocode"
    End If
End Sub

Private Sub Text23_LostFocus()
    If edcell = "radiocode" Then Call update_rec
End Sub

Private Sub Text24_GotFocus()
    Text24.SelStart = 0
    Text24.SelLength = Len(Text24)
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text25.SetFocus
    Else
        edcell = "work_phone"
    End If
End Sub

Private Sub Text24_LostFocus()
    If edcell = "work_phone" Then Call update_rec
End Sub

Private Sub Text25_GotFocus()
    Text25.SelStart = 0
    Text25.SelLength = Len(Text25)
End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text1.SetFocus
    Else
        edcell = "cellphone"
    End If
End Sub

Private Sub Text25_LostFocus()
    If edcell = "cellphone" Then Call update_rec
End Sub

Private Sub Text3_GotFocus()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text8.SetFocus
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
        edcell = "ss_name"
    End If
End Sub

Private Sub Text6_LostFocus()
    If edcell = "ss_name" Then Call update_rec
End Sub

Private Sub Text7_GotFocus()
    Text7.SelStart = 0
    Text7.SelLength = Len(Text7)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text9.SetFocus
    Else
        edcell = "ss_num"
    End If
End Sub

Private Sub Text7_LostFocus()
    If edcell = "ss_num" Then Call update_rec
End Sub

Private Sub Text8_GotFocus()
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8)
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text4.SetFocus
    Else
        edcell = "dob"
    End If
End Sub

Private Sub Text8_LostFocus()
    If edcell = "dob" Then Call update_rec
End Sub

Private Sub Text9_GotFocus()
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9)
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Text10.SetFocus
    Else
        edcell = "dl_name"
    End If
End Sub

Private Sub Text9_LostFocus()
    If edcell = "dl_name" Then Call update_rec
End Sub

Private Sub vallists_Click()
    Form9.Show
End Sub

Private Sub wddepts_Click()
    Form2.ftype = "Departments"
    Form2.Show
End Sub

Private Sub xitmenu_Click()
    Dim i As Integer, f As String
    Dim t As Long, l As Long, h As Long, w As Long
    If Len(edcell) > 0 Then
        If MsgBox("Save Changes?", vbYesNo + vbQuestion, "Save changes..") = vbYes Then
            Call update_rec
        Else
            edcell = ""
        End If
    End If
    If Form1.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form1" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form1.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form1.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form1.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form1.Width
                Exit For
            End If
        Next i
    End If
    Open "c:\windows\wdemp.ini" For Output As #1
    For i = 1 To frmgrid.Rows - 1
        f = frmgrid.TextMatrix(i, 0)
        t = Val(frmgrid.TextMatrix(i, 1))
        l = Val(frmgrid.TextMatrix(i, 2))
        h = Val(frmgrid.TextMatrix(i, 3))
        w = Val(frmgrid.TextMatrix(i, 4))
        Write #1, f, t, l, h, w
    Next i
    Close #1
    End
End Sub
