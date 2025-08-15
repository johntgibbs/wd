VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Blue Bell Warehouse Client"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox bbsr 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sign In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox userid 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "V2025.01.17"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   3240
      Picture         =   "r0barcode1.frx":0000
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label emess 
      Alignment       =   2  'Center
      Caption         =   "Invalid User ID!!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   4800
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   1920
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim uname As String
    Me.Caption = "Blue Bell Warehouse Client"
    If Len(userid) <> 6 Then
        emess.Visible = True
    Else
        uname = wdempname(userid)
        If userid = uname Then
            emess.Visible = True
        Else
            WDUserId = Me.userid                        'jv111915
            Me.Caption = uname
            emess.Visible = False
            Form2.Show
        End If
    End If
End Sub

Private Sub Form_Activate()
    userid = ""
    emess.Visible = False
    userid.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub Form_Load()
    vberror_log = "\\bbc-01-prodtrk\wd\temp\sqlerrors.txt"
    userid = ""
    emess.Visible = False
    'Me.bbsr = "ODBC;DATABASE=WDRacks;DSN=wdracks"
    'Me.bbsr = "ODBC;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500;DSN=wdsql500"
    Me.bbsr = "Driver={SQL Server};Server=bbc-08-sqlsvr;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500"
    WDbbsr = Me.bbsr
    'logdir = "v:\testlogs\"
    logdir = "\\bbc-01-prodtrk\wd\pallogs\"
    
    Set Wdb = CreateObject("ADODB.Connection")
    Wdb.Open WDbbsr
    
    labfmtfile = "\\BBC-03-FILESVR\SharedGroups\wd\bin\labfmt.txt"
    load_labpics
    'labpics.Show
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    'Call vb_elog(eno, edesc, "wmsmobile.bas", "add_alternate_dock_pallet", wduserid)
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, Me.Name, "form_load", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: form_load: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    Dim msz As Long
    If Me.WindowState = 1 Then Exit Sub
    Label1.Left = (Me.Width - Label1.Width) * 0.5
    userid.Left = Label1.Left
    Command1.Left = Label1.Left
    emess.Left = (Me.Width - emess.Width) * 0.5
    Image1.Left = (Me.Width - Image1.Width) * 0.5
    msz = Label1.Height + userid.Height + Command1.Height + emess.Height + Image1.Height
    If Me.Height > 4500 Then 'msz + 1000 Then
        userid.Top = (Me.Height - (userid.Height / 2)) * 0.5
        Image1.Top = userid.Top - 2040
        Label1.Top = userid.Top - 600
        Command1.Top = userid.Top + 720
        emess.Top = userid.Top + 2280
    End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub userid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
