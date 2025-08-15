VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Tri-Level Traffic Master"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox daisqldb 
      Height          =   285
      Left            =   360
      TabIndex        =   23
      Text            =   "Text5"
      Top             =   600
      Width           =   8415
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Text            =   "Text4"
      Top             =   6000
      Width           =   8175
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2055
      Left            =   0
      TabIndex        =   21
      Top             =   6960
      Width           =   9615
      ExtentX         =   16960
      ExtentY         =   3625
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   7440
      TabIndex        =   20
      Text            =   "Text3"
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox daioradb 
      Height          =   285
      Left            =   360
      TabIndex        =   19
      Text            =   "Text3"
      Top             =   0
      Width           =   8415
   End
   Begin VB.TextBox tbbsr 
      Height          =   285
      Left            =   360
      TabIndex        =   18
      Text            =   "Text3"
      Top             =   1320
      Width           =   8415
   End
   Begin VB.TextBox ws3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   17
      Text            =   "38050"
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox ws2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   16
      Text            =   "22430"
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox ws1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   15
      Text            =   "10501"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   8040
      TabIndex        =   11
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7440
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   4200
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7440
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Conveyors OnLine "
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   600
      TabIndex        =   3
      Top             =   3720
      Width           =   2655
      Begin VB.CheckBox srstat5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SR-5"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox srstat4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SR-4"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox srstat3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SR-3"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox srstat2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SR-2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox srstat1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SR-1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3201
      _Version        =   327680
   End
   Begin VB.TextBox bbsr 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   8415
   End
   Begin VB.TextBox oradb 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   8415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wrapper 3 BC Sequence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   14
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wrapper 2 BC Sequence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   13
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wrapper 1 BC Sequence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   12
      Top             =   3960
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub poll_wrappers()
    Dim cfile As String
    Do While True
        Text1 = Format(Now, "MMddyyyy")
        cfile = "\\BBC-01-PRODTRK\wd\pallogs\recv" & Text1 & ".txt"
        'Text2 = FileLen(cfile)
        DoEvents
        cfile = "\\BBC-01-PRODTRK\wd\pallogs\move" & Text1 & ".txt"
        Text3 = FileLen(cfile)
        DoEvents
    Loop
End Sub

Sub bhsp_pallets()
    Dim db As ADODB.Connection, ds As Recordset, sqlx As String
    Dim d As daiexprct, rkey As Long
    Dim cfile As String, f0 As String
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.bbsr
    sqlx = "select * from queue_infc where queue_num <> 0"
    sqlx = sqlx & " and source = 'FG3'"
    Set ds = db.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            lid = check_plateno(ds!palletid)
            If lid > "0" Then
                d.action = "ADD"
                d.sOrderID = lid
                d.dExpectedDate = Format(Now, "MM/dd/yyyy hh:mm:ss")
                d.sItem = ds!SKU
                d.sLot = ds!lot_num
                d.fExpectedQuantity = ds!units + ds!units2
                d.sStoreDestination = "3"
                Text4.Text = Dai_expected_receipt(d)
                Open "c:\jvwork\daiExpectedReceiptMessage.xml" For Output As #1
                Print #1, Text4.Text
                Close #1
                DoEvents
                rkey = wd_seq("DAIRequests")
                'Call write_oracle_request("ExpectedReceiptMessage", Val(d.sOrderID))
                Call write_oracle_request("ExpectedReceiptMessage", rkey)
                WebBrowser1.Navigate2 "c:\jvwork\daiExpectedReceiptMessage.xml"
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
End Sub

Function check_plateno(bc As String) As String
    Dim db As ADODB.Connection, ds As Recordset, sqlx As String
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.tbbsr
    sqlx = "select id, plateno from pallets where barcode = '" & bc & "'"
    sqlx = sqlx & " and status = 'Wrapper'"
    Set ds = db.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        check_plateno = ds!plateno
        MsgBox "Barcode: " & bc & " = Plateno " & ds!plateno
        sqlx = "Update pallets set status = 'Warehouse' where id = " & ds!id
        db.Execute sqlx
    Else
        check_plateno = " "
        'MsgBox "Barcode: " & bc & " = No Plateno "
    End If
    ds.Close: db.Close
End Function

Private Sub Command1_Click()
    poll_wrappers
End Sub

Private Sub Form_Load()
    'bbsr = "odbc;database=wdracks;uid=bbcwd500;pwd=brenham500;dsn=wdsql500"
    bbsr = "Driver={SQL Server};Server=bbc-08-sqlsvr;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500;"
    'tbbsr = "odbc;database=wdracks;dsn=wdracks"
    tbbsr = "Driver={SQL Server};Server=bbc-08-sqlsvr;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500;"
    oradb = "odbc;database=pbelle;uid=Apps;pwd=pb3113tx;dsn=pbelle"
    'oradb = "odbc;database=tbelle;uid=bolinf;pwd=euge_pbbcri;dsn=tbelle"
    daioradb = "odbc;database=bluebell;uid=wrxjhost;pwd=asrs;dsn=bluebell"
    daisqldb = "DRIVER={SQL Server};Server=BBC-08-SQLSVR;database=dbDaifuku;uid=asrs;pwd=rdP4fOpiKkqknoi0PwDlw6QevTXX2bdu;"
    
    localAppDataPath = Environ("LOCALAPPDATA") & "\TrafficMaster"
    If DirExists(localAppDataPath) <> True Then
        MkDir (localAppDataPath)
    End If
    allocations.Show
    queue_infc.Show
    tmtasks.Show
    daimessage.Show
    'saerequests.Show
    daiship.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Text2_Change()
    tmtasks.ttrig = Text2
    DoEvents
    Call daimessage.dai_poll_messages
End Sub

Private Sub Text3_Change()
    Call bhsp_pallets
End Sub

Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    Dim RetVal As Boolean
    'RetVal = (GetAttr(DirName) = vbDirectory)
    RetVal = (FileLen(DirName) >= 0)
    
    DirExists = RetVal
    Exit Function
ErrorHandler:
    If (Err = 53) Then ' 53 means file was not found at all
        DirExists = False
    End If
    DirExists = False
End Function
