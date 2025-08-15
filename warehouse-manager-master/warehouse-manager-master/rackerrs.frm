VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form17 
   Caption         =   "Rack Error Log"
   ClientHeight    =   8925
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11040
   LinkTopic       =   "Form17"
   ScaleHeight     =   8925
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   8055
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read Log File"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6495
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   11456
      _Version        =   327680
      BackColorFixed  =   8454016
      BackColorSel    =   33023
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Menu usermenu 
      Caption         =   "User"
      Visible         =   0   'False
      Begin VB.Menu emplook 
         Caption         =   "Lookup Employee Name"
      End
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim cfile As String, s As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim f16 As String, logpath As String
    logpath = Form1.logdir
    List1.Clear
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 14
    cfile = logpath & "elog" & Format(Text1, "mmddyyyy") & ".txt"
    If Len(Dir(cfile)) = 0 Then
        MsgBox "No error log exists for this date.", vbOKOnly + vbExclamation, cfile
        Exit Sub
    End If
    Open cfile For Input Shared As #1
    Do Until EOF(1)
        Line Input #1, s
        List1.AddItem s
        If Mid(s, 2, 6) = "failed" Then
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
            's = f0 & Chr(9)                 'id
            s = f1 & Chr(9)             'area
            's = s & f2 & Chr(9)             'description
            s = s & f3 & Chr(9)             'source
            s = s & f4 & Chr(9)             'target
            s = s & f5 & Chr(9)             'product
            s = s & f6 & Chr(9)             'pallet
            s = s & f7 & Chr(9)             'qty
            s = s & f8 & Chr(9)             'uom
            s = s & f9 & Chr(9)             'lot
            s = s & f10 & Chr(9)            'units
            s = s & f11 & Chr(9)            'lot2
            s = s & f12 & Chr(9)            'units2
            s = s & f13 & Chr(9)            'status
            s = s & f14 & Chr(9)            'user
            s = s & f15 '& Chr(9)            'time
            's = s & f16                     'reqid
            Grid1.AddItem s
        End If
    Loop
    Close #1
    s = "^Area|^Source|^Target|<Product|<Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum2|^Units2|^Status|^User|<Time"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1400
    Grid1.ColWidth(1) = 1600
    Grid1.ColWidth(2) = 1600
    Grid1.ColWidth(3) = 3000
    Grid1.ColWidth(4) = 2000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1000
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 1000
    Grid1.ColWidth(11) = 1000
    Grid1.ColWidth(12) = 1000
    Grid1.ColWidth(13) = 1600
    'Grid1.ColWidth(14) = 1000
    'Grid1.ColWidth(15) = 1000
    'Grid1.ColWidth(16) = 1000
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub emplook_Click()                 'Lookup Employee
    Dim ds As adodb.Recordset, s As String
    If Len(Grid1.Text) = 0 Then Exit Sub
    'SQL Database - bbsr
    s = "select * from valuelists where listname = 'wdempid'"
    s = s & " and listreturn = '" & Grid1.Text & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = ds!listdisplay
    Else
        s = "Employee #: " & Grid1.Text & " is not in WdEmp database."
    End If
    ds.Close
    MsgBox s, vbOKOnly + vbInformation, "WMS SQL Employee " & Grid1.Text & " ...."
End Sub

Private Sub Form_Load()
    Text1 = Format(Now, "mm-dd-yyyy")
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    List1.Width = Me.Width - 100
    If Me.Height > 2000 Then List1.Height = Me.Height - 1240 '1020
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If Grid1.TextMatrix(0, Grid1.Col) = "User" Then
            PopupMenu usermenu
        End If
    End If
End Sub

