VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form22 
   Caption         =   "Oracle Daifuku Messages"
   ClientHeight    =   11805
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12810
   LinkTopic       =   "Form22"
   ScaleHeight     =   11805
   ScaleWidth      =   12810
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5055
      Left            =   0
      TabIndex        =   3
      Top             =   6480
      Width           =   6135
      ExtentX         =   10821
      ExtentY         =   8916
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
      Location        =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5775
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   10186
      _Version        =   327680
      BackColorFixed  =   16777152
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
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
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label daioradb 
      Caption         =   "Label1"
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
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   8175
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu delrec 
         Caption         =   "Delete Message"
      End
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim odb As adodb.Connection

Private Sub refresh_grid()
    Dim ds As adodb.Recordset
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4
    s = "select * from hosttowrx order by imessagesequence"
    Set ds = odb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(0) & Chr(9) & ds(1) & Chr(9) & ds(2) & Chr(9) & ds(3)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "<Time|^ID|<Type|<Message"
    Grid1.ColWidth(0) = 1800
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 2000
    Grid1.ColWidth(3) = 7000
    Grid1.Redraw = True
    Grid1_Click
    Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub delrec_Click()              'Delete Message Record
    Dim s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 1)) = 0 Then Exit Sub
    s = "Delete from HostToWrx Where imessagesequence = " & Grid1.TextMatrix(Grid1.Row, 1)
    odb.Execute s
    refresh_grid
End Sub

Private Sub Form_Load()
    Me.daioradb = "odbc;database=bluebell;uid=wrxjhost;pwd=asrs;dsn=bluebell"
    'Me.daioradb = "odbc;database=wdracks;dsn=wdracks"
    Set odb = CreateObject("ADODB.Connection")
    odb.Open Me.daioradb
    refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    WebBrowser1.Width = Me.Width - 180
End Sub

Private Sub Form_Unload(Cancel As Integer)
    odb.Close
End Sub

Private Sub Grid1_Click()
    cfile = "U:\xlook.xml"
    If Len(Grid1.TextMatrix(Grid1.Row, 3)) > 100 Then
        Open cfile For Output As #1
        Print #1, Grid1.TextMatrix(Grid1.Row, 3)
        Close #1
    End If
    WebBrowser1.Navigate cfile
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
    If Button = 1 Then Grid1_Click
End Sub

Private Sub Grid1_RowColChange()
    If Grid1.Redraw = True Then Grid1_Click
End Sub

