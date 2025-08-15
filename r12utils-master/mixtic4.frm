VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Mixer Batch Tickets"
   ClientHeight    =   11655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11040
   LinkTopic       =   "Form4"
   ScaleHeight     =   11655
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4575
      Left            =   0
      TabIndex        =   10
      Top             =   7080
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8070
      _Version        =   327680
      BackColor       =   12648447
      ForeColor       =   16711680
      BackColorFixed  =   16777152
      FocusRect       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton Option6 
         Caption         =   "All"
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
         Left            =   2520
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "FG"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Mixer"
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
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton Option3 
         Caption         =   "Complete"
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
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Pending"
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
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Errors"
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
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
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
      Left            =   9360
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9763
      _Version        =   327680
      ForeColor       =   8421376
      BackColorFixed  =   12648447
      WordWrap        =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label elit 
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
      Left            =   1560
      TabIndex        =   12
      Top             =   6840
      Width           =   7815
   End
   Begin VB.Label seqkey 
      Caption         =   "seqkey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6840
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 14
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.oradb
    s = "select * from mixer_batch_hdr"
    If Option1 = True Then s = s & " where upload_flag = 2"
    If Option2 = True Then s = s & " where upload_flag = 0"
    If Option3 = True Then s = s & " where upload_flag = 1"
    If Option4 = True Then s = s & " and type = 'MIXER'"
    If Option5 = True Then s = s & " and type = 'FG'"
    s = s & " and produce_date > SYSDATE - 60"
    s = s & " order by upload_time desc"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!seq_id & Chr(9)
            s = s & ds!orgn_code & Chr(9)
            s = s & ds!queue_id & Chr(9)
            s = s & ds!batch_id & Chr(9)
            s = s & ds!p_system & Chr(9)
            s = s & ds!formula_id & Chr(9)
            s = s & ds!time_started & Chr(9)
            s = s & ds!time_finished & Chr(9)
            s = s & ds!added_by & Chr(9)
            s = s & ds!upload_time & Chr(9)
            s = s & ds!upload_flag & Chr(9)
            s = s & ds!Type & Chr(9)
            s = s & ds!Error & Chr(9)
            s = s & ds!produce_date
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    s = "^Seq_id|^Org|^Queue|^Batch_id|^P_system|^FormulaId|<Time_started|<Time_finished"
    s = s & "|^added_by|<Upload_time|^Upload_Flag|^Type|<Error|<Produce_date"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 2300
    Grid1.ColWidth(7) = 2300
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 2300
    Grid1.ColWidth(10) = 1200
    Grid1.ColWidth(11) = 800
    Grid1.ColWidth(12) = 1000
    Grid1.ColWidth(13) = 2300
    Call Grid1_RowColChange
End Sub

Private Sub refresh_grid2()
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 9
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.oradb
    s = "select * from mixer_batch_dtl where seq_id = " & seqkey.Caption
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!seq_id & Chr(9)
            s = s & ds!item_id & Chr(9)
            s = s & ds!item_qty & Chr(9)
            s = s & ds!item_type & Chr(9)
            s = s & ds!whse_code & Chr(9)
            s = s & ds!loct_code & Chr(9)
            s = s & ds!lot & Chr(9)
            s = s & ds!Line
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid2.Rows > 1 Then
        For i = 1 To Grid2.Rows - 1
            s = "select segment1, description from mtl_system_items_b"
            s = s & " where inventory_item_id = " & Grid2.TextMatrix(i, 1)
            Set ds = db.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                Grid2.TextMatrix(i, 8) = ds!segment1 & " " & ds!Description
            End If
            ds.Close
        Next i
    End If
    db.Close
    s = "^Seq_id|^Item_id|^Item_qty|^Item_type|^Whse_code|^Loct_code|<Lot|^Line|<Item"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 1200
    Grid2.ColWidth(1) = 1200
    Grid2.ColWidth(2) = 1200
    Grid2.ColWidth(3) = 1200
    Grid2.ColWidth(4) = 1200
    Grid2.ColWidth(5) = 1500
    Grid2.ColWidth(6) = 2400
    Grid2.ColWidth(7) = 1200
    Grid2.ColWidth(8) = 5500
End Sub

Private Sub Command1_Click()
    refresh_grid1
End Sub

Private Sub Form_Load()
    refresh_grid1
    Grid1_RowColChange
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    Grid2.Width = Me.Width - 100
End Sub

Private Sub Grid1_RowColChange()
    seqkey.Caption = Val(Grid1.TextMatrix(Grid1.Row, 0))
    elit.Caption = Grid1.TextMatrix(Grid1.Row, 12)
End Sub

Private Sub seqkey_Change()
    refresh_grid2
End Sub

