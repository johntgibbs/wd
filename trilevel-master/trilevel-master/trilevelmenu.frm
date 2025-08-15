VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tri-Level Loop"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   1680
      TabIndex        =   0
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "V07.23.18"
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
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label dailogs 
      Caption         =   "Label2"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   6960
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label srserv 
      Caption         =   "v:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label plantno 
      Caption         =   "50"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label pbelle 
      Caption         =   "Label2"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.Label bbsr 
      Caption         =   "Label2"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   2760
      Picture         =   "trilevelmenu.frx":0000
      Top             =   480
      Width           =   2235
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tri-Level Options Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2640
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    List1.AddItem "Crane Traffic"
    List1.AddItem "Change Plate"
    List1.AddItem "Receiving List"
    List1.AddItem "Movement Logs"
    List1.AddItem "Exit"
    List1.ListIndex = 0
    WDUserId = "TMaster"
    Me.bbsr = "odbc;database=wdracks;uid=bbcwd500;pwd=brenham500;dsn=wdsql500"
    'Me.bbsr = "odbc;database=wdracks;dsn=wdracks"
    Me.pbelle = "odbc;database=pbelle;uid=apps;pwd=papps;dsn=pbelle"
    WDbbsr = Me.bbsr
    daioradb = "odbc;database=bluebell;uid=wrxjhost;pwd=asrs;dsn=bluebell"
    'logdir = "V:\testlogs\"
    logdir = "\\bbc-01-prodtrk\wd\pallogs\"
    Me.dailogs = "\\bbc-01-prodtrk\wd\sr5\bin\"
    vberror_log = "\\bbc-01-prodtrk\wd\temp\sqlerrors.txt"
    Set Wdb = CreateObject("ADODB.Connection")
    Wdb.Open WDbbsr
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    Label1.Left = (Me.Width - Label1.Width) * 0.5
    List1.Left = Label1.Left
    Image1.Left = (Me.Width - Image1.Width) * 0.5
    If Me.Height > 6000 Then
        List1.Top = (Me.Height - (List1.Height / 2)) * 0.5
        Image1.Top = List1.Top - 3120
        Label1.Top = List1.Top - 480
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Wdb.Close
    End
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If List1 = "Crane Traffic" Then cranetraffic.Show
        If List1 = "Change Plate" Then
            Form5.Option4.Value = True
            Form5.Show
        End If
        If List1 = "Receiving List" Then tlreceiving.Show
        If List1 = "Movement Logs" Then traffmoves.Show
        If List1 = "Exit" Then Unload Me
    End If
    If KeyAscii = 27 Then Unload Me
End Sub

