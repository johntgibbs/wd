VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form20 
   Caption         =   "BlueBell SQL Daifuku Requests"
   ClientHeight    =   11520
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14325
   LinkTopic       =   "Form20"
   ScaleHeight     =   11520
   ScaleWidth      =   14325
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6135
      Left            =   0
      TabIndex        =   5
      Top             =   5280
      Width           =   9735
      ExtentX         =   17171
      ExtentY         =   10821
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
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   8070
      _Version        =   327680
      BackColorFixed  =   16777152
      BackColorSel    =   4210752
      WordWrap        =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label jcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "jcolor"
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
      Left            =   9240
      TabIndex        =   4
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label ccolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ccolor"
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
      Left            =   6480
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label pcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "pcolor"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu markpend 
         Caption         =   "Mark - PEND"
      End
      Begin VB.Menu markcomp 
         Caption         =   "Mark - COMP"
      End
      Begin VB.Menu markjunk 
         Caption         =   "Mark - JUNK"
      End
      Begin VB.Menu markcompall 
         Caption         =   "Mark COMP - All"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete Record"
      End
   End
   Begin VB.Menu procmenu 
      Caption         =   "Process"
      Begin VB.Menu delallcomp 
         Caption         =   "Clear All Complete Records"
      End
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refresh_grid1()
    Dim ds As adodb.Recordset, s As String
    Dim pq As Integer, cq As Integer, jq As Integer
    pq = 0: cq = 0: jq = 0
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    s = "select * from BBC_HostToWrx order by imessagesequence"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = Format(ds(0), "M-d-yy h:mm:ss am/pm") & Chr(9)
            s = s & ds(1) & Chr(9) & ds(2) & Chr(9) & ds(3) & Chr(9) & ds(4) & Chr(9) & ds(5)
            If ds(5) = "PEND" Then
                pq = pq + 1
            Else
                If ds(5) = "COMP" Then
                    cq = cq + 1
                Else
                    jq = jq + 1
                End If
            End If
            Grid1.AddItem s
            ds.MoveNext
        Loop
    Else
        s = "Update sequences set sequence_id = 0 where seq = 'BBC_HostToWrx'"
        Wdb.Execute s
    End If
    ds.Close
    pcolor.Visible = False: ccolor.Visible = False: jcolor.Visible = False
    If pq > 0 Then
        pcolor.Caption = pq & " Pending Records"
        pcolor.Visible = True
    End If
    If cq > 0 Then
        ccolor.Caption = cq & " Completed Records"
        ccolor.Visible = True
    End If
    If jq > 0 Then
        jcolor.Caption = jq & " Junk Records"
        jcolor.Visible = True
    End If
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 3: Grid1.ColSel = 5
            If Grid1.TextMatrix(i, 5) = "PEND" Then
                Grid1.CellBackColor = pcolor.BackColor
            Else
                If Grid1.TextMatrix(i, 5) = "COMP" Then
                    Grid1.CellBackColor = ccolor.BackColor
                Else
                    Grid1.CellBackColor = jcolor.BackColor
                End If
            End If
        Next i
    End If
            
    Grid1.FormatString = "<Time|^ID|<Type|<Message|<BB ID|^Status"
    Grid1.ColWidth(0) = 1800
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 2000
    Grid1.ColWidth(3) = 4000
    Grid1.ColWidth(4) = 2000
    Grid1.ColWidth(5) = 1000
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
    refresh_grid1
End Sub

Private Sub delallcomp_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 1)) = 0 Then Exit Sub
    s = "Delete from BBC_HostToWrx where bbcstatus = 'COMP'"
    Wdb.Execute s
    refresh_grid1
    If i < Grid1.Rows - 1 Then Grid1.Row = i
End Sub

Private Sub delrec_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 1)) = 0 Then Exit Sub
    s = "Delete from BBC_HostToWrx where imessagesequence = "
    s = s & Grid1.TextMatrix(i, 1)
    Wdb.Execute s
    refresh_grid1
    If i < Grid1.Rows - 1 Then Grid1.Row = i
End Sub

Private Sub Form_Load()
    refresh_grid1
    Grid1_Click
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    WebBrowser1.Width = Me.Width - 100
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

Private Sub markcomp_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 1)) = 0 Then Exit Sub
    s = "Update BBC_HostToWrx set bbcstatus = 'COMP' where imessagesequence = "
    s = s & Grid1.TextMatrix(i, 1)
    Wdb.Execute s
    refresh_grid1
    If i < Grid1.Rows - 1 Then Grid1.Row = i
End Sub

Private Sub markcompall_Click()
    Dim i As Integer
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 1)) > 0 Then
            s = "Update BBC_HostToWrx set bbcstatus = 'COMP' where imessagesequence = "
            s = s & Grid1.TextMatrix(i, 1)
            Wdb.Execute s
        End If
    Next i
    refresh_grid1
End Sub

Private Sub markjunk_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 1)) = 0 Then Exit Sub
    s = "Update BBC_HostToWrx set bbcstatus = 'JUNK' where imessagesequence = "
    s = s & Grid1.TextMatrix(i, 1)
    Wdb.Execute s
    refresh_grid1
    If i < Grid1.Rows - 1 Then Grid1.Row = i
End Sub

Private Sub markpend_Click()
    Dim s As String, i As Integer
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 1)) = 0 Then Exit Sub
    s = "Update BBC_HostToWrx set bbcstatus = 'PEND' where imessagesequence = "
    s = s & Grid1.TextMatrix(i, 1)
    Wdb.Execute s
    refresh_grid1
    If i < Grid1.Rows - 1 Then Grid1.Row = i
End Sub
