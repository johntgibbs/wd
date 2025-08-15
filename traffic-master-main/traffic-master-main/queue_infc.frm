VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form queue_infc 
   Caption         =   "SR Queues"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   LinkTopic       =   "Form2"
   ScaleHeight     =   5910
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8070
      _Version        =   327680
   End
   Begin VB.Label ques5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ques5"
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
      Left            =   6720
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Label ques3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ques3"
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
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label ques2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ques2"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label ques1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ques1"
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
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label srlabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR-5"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label srlabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR-3"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label srlabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR-2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label srlabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR-1"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label trigkey 
      Caption         =   "trigkey"
      Height          =   255
      Left            =   7560
      TabIndex        =   0
      Top             =   5400
      Width           =   1455
   End
End
Attribute VB_Name = "queue_infc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function sr_conveyor_count(whs As String) As Integer
    Dim i As Integer, c As Integer
    c = 0
    If Grid1.Rows < 2 Then
        sr_conveyor_count = 0
    Else
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 1) = whs Then c = c + 1
        Next i
        sr_conveyor_count = c
    End If
End Function

Function sr_single_sku(pwhs As String, psku As String) As String
    Dim i As Integer, c As Integer
    c = 0
    If Grid1.Rows < 2 Then
        c = 0
    Else
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 1) = pwhs And Grid1.TextMatrix(i, 2) = psku Then
                c = c + 1
            End If
        Next i
    End If
    If c = 1 Or c = 3 Or c = 5 Or c = 7 Then
        sr_single_sku = "1"
    Else
        sr_single_sku = "0"
    End If
    'MsgBox sr_single_sku & " " & pwhs & " " & psku & " " & c
End Function

Sub refresh_grid1()
    Dim db As ADODB.Connection, ds As Recordset, s As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    db.Open Form1.tbbsr
    s = "select * from queue_infc"
    's = s & " where queue_num > 0 and source = 'TML' order by queue_num"
    s = s & " where queue_num > 0 and source in ('TML', 'FG3') order by queue_num"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!whse_num & Chr(9)
            s = s & ds!SKU & Chr(9)
            s = s & ds!lot_num & Chr(9)
            s = s & ds!drop_flag & Chr(9)
            s = s & ds!queue_num & Chr(9)
            s = s & ds!rack_num & Chr(9)
            s = s & ds!units & Chr(9)
            s = s & ds!lot_num2 & Chr(9)
            s = s & ds!units2 & Chr(9)
            s = s & ds!palletid & Chr(9)
            s = s & ds!Source
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    s = "^ID|^Whs|^SKU|^Lot|^Dflag|^Queue|^Rack|^Units|^Lot2|^Units2|<Pallet|^Source"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 800
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 800
    Grid1.ColWidth(10) = 1600
    Grid1.ColWidth(11) = 800
    ques1.Caption = sr_conveyor_count("1")
    ques2.Caption = sr_conveyor_count("2")
    ques3.Caption = sr_conveyor_count("3")
    ques5.Caption = sr_conveyor_count("5")
End Sub

Private Sub Form_Load()
    refresh_grid1
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
End Sub

Private Sub trigkey_Change()
    refresh_grid1
End Sub

