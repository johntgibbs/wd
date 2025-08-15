VERSION 5.00
Begin VB.Form branchtrans 
   Caption         =   "Branch Transfers"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form23"
   ScaleHeight     =   7425
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Post To Logs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   34
      Top             =   6600
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   32
      Text            =   "reqid"
      Top             =   5880
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   14
      Left            =   2040
      TabIndex        =   31
      Text            =   "time"
      Top             =   5520
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   13
      Left            =   2040
      TabIndex        =   30
      Text            =   "user"
      Top             =   5160
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   2040
      TabIndex        =   29
      Text            =   "status"
      Top             =   4800
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   28
      Text            =   "units"
      Top             =   4440
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   27
      Text            =   "lot2"
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   26
      Text            =   "units"
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   25
      Text            =   "lotnum"
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   24
      Text            =   "uom"
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   23
      Text            =   "qty"
      Top             =   2640
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   22
      Text            =   "pallet"
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   21
      Text            =   "product"
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   20
      Text            =   "source"
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   19
      Text            =   "description"
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   18
      Text            =   "area"
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox cval 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   17
      Text            =   "cval"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   16
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   15
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   14
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   13
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   12
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label tagname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "tagname"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "branchtrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_branches()
    Dim ds As ADODB.Recordset, s As String
    Combo1.Clear
    s = "select branchname from branches where branch NOT IN (97, 98, 99) order by branchname"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem UCase(ds!branchname)
            ds.MoveNext
        Loop
    End If
    ds.Close
    Combo1.ListIndex = 0
End Sub

Private Sub Command1_Click()                'Post To Logs
    Dim cfile As String, dt As String, t1 As String, t2 As String, s As String
    Dim a10logs As String, k10logs As String, opcode As String
    opcode = Mid(cval(5).Text, 11, 3)
    MsgBox opcode
    a10logs = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
    k10logs = "\\bbba-03-dc\f\user\waredist\data\pallogs\"
    t1 = Form15.Grid1.TextMatrix(Form15.Grid1.Row, 16)
    t1 = Mid(t1, 3, 2) & "-" & Mid(t1, 5, 2) & "-20" & Mid(t1, 1, 2)
    t2 = Format(cval(14).Text, "MM-dd-yyyy")
    If DateDiff("d", t1, t2) < 0 Then
        s = "The log date entered (" & t2 & ") cannot be earlier than the previous logged ship date (" & t1 & ")."
        MsgBox s, vbOKOnly + vbExclamation, "sorry, try another date..."
        Exit Sub
    End If
    dt = Format(cval(14).Text, "yyMMdd") & " 14:00:00"
    cfile = Form1.logdir & "ship" & Format(cval(14).Text, "MMddyyyy") & ".txt"
    'cfile = "v:\testlogs\ship" & Format(cval(14).Text, "MMddyyyy") & ".txt"
    If MsgBox("Ok to post to " & cfile, vbYesNo + vbQuestion, "post to logs....") = vbNo Then Exit Sub
    s = "S" & Chr(9)
    s = s & cval(0).Text & Chr(9)   'Recid
    s = s & cval(1).Text & Chr(9)   'Area
    s = s & cval(2).Text & Chr(9)   'Description
    s = s & cval(3).Text & Chr(9)   'Source
    s = s & Combo1 & Chr(9)         'Target
    s = s & cval(4).Text & Chr(9)   'Product
    s = s & bc000(cval(5).Text) & Chr(9)   'Pallet
    s = s & cval(6).Text & Chr(9)   'Qty
    s = s & cval(7).Text & Chr(9)   'Uom
    s = s & cval(8).Text & Chr(9)   'Lotnum
    s = s & cval(9).Text & Chr(9)   'units
    s = s & cval(10).Text & Chr(9)  'Lot2
    s = s & cval(11).Text & Chr(9)  'units
    s = s & cval(12).Text & Chr(9)  'Status
    s = s & cval(13).Text & Chr(9)  'Userid
    s = s & dt & Chr(9)             'time
    s = s & cval(15).Text & Chr(9)  'reqid
    s = s & bc000(cval(5).Text) & dt    'sortcolumn
    Form15.Grid1.AddItem s
        
    Open cfile For Append Shared As #1
    Write #1, cval(0).Text;         'Recid
    Write #1, cval(1).Text;         'Area
    Write #1, cval(2).Text;         'Description
    Write #1, cval(3).Text;         'Source
    Write #1, Combo1;               'Target
    Write #1, cval(4).Text;         'Product
    Write #1, cval(5).Text;         'Pallet
    Write #1, cval(6).Text;         'Qty
    Write #1, cval(7).Text;         'Uom
    Write #1, cval(8).Text;         'Lotnum
    Write #1, cval(9).Text;         'Units
    Write #1, cval(10).Text;        'Lot2
    Write #1, cval(11).Text;        'units
    Write #1, cval(12).Text;        'Status
    Write #1, cval(13).Text;        'Userid
    Write #1, dt;                   'time
    Write #1, cval(15).Text         'reqid
    Close #1
    Form15.sortshiptrig.Caption = Val(Form15.sortshiptrig.Caption) + 1
    MsgBox "posted to " & cfile, vbOKOnly + vbInformation, "posted....."
        
        
    If opcode >= "100" And opcode <= "199" Then
        cfile = k10logs & "ship" & Format(cval(14).Text, "MMddyyyy") & ".txt"
        'cfile = "v:\testlogs\ship" & Format(cval(14).Text, "MMddyyyy") & ".txt"
        Open cfile For Append Shared As #1
        Write #1, cval(0).Text;         'Recid
        Write #1, cval(1).Text;         'Area
        Write #1, cval(2).Text;         'Description
        Write #1, cval(3).Text;         'Source
        Write #1, Combo1;               'Target
        Write #1, cval(4).Text;         'Product
        Write #1, cval(5).Text;         'Pallet
        Write #1, cval(6).Text;         'Qty
        Write #1, cval(7).Text;         'Uom
        Write #1, cval(8).Text;         'Lotnum
        Write #1, cval(9).Text;         'Units
        Write #1, cval(10).Text;        'Lot2
        Write #1, cval(11).Text;        'units
        Write #1, cval(12).Text;        'Status
        Write #1, cval(13).Text;        'Userid
        Write #1, dt;                   'time
        Write #1, cval(15).Text         'reqid
        Close #1
        MsgBox "posted to " & cfile, vbOKOnly + vbInformation, "posted....."
    End If
                        
    If opcode >= "200" And opcode <= "299" Then
        cfile = a10logs & "ship" & Format(cval(14).Text, "MMddyyyy") & ".txt"
        'cfile = "v:\testlogs\ship" & Format(cval(14).Text, "MMddyyyy") & ".txt"
        Open cfile For Append Shared As #1
        Write #1, cval(0).Text;         'Recid
        Write #1, cval(1).Text;         'Area
        Write #1, cval(2).Text;         'Description
        Write #1, cval(3).Text;         'Source
        Write #1, Combo1;               'Target
        Write #1, cval(4).Text;         'Product
        Write #1, cval(5).Text;         'Pallet
        Write #1, cval(6).Text;         'Qty
        Write #1, cval(7).Text;         'Uom
        Write #1, cval(8).Text;         'Lotnum
        Write #1, cval(9).Text;         'Units
        Write #1, cval(10).Text;        'Lot2
        Write #1, cval(11).Text;        'units
        Write #1, cval(12).Text;        'Status
        Write #1, cval(13).Text;        'Userid
        Write #1, dt;                   'time
        Write #1, cval(15).Text         'reqid
        Close #1
        MsgBox "posted to " & cfile, vbOKOnly + vbInformation, "posted....."
    End If
        
    'Unload Me
End Sub

Private Sub Form_Load()
    refresh_branches
    tagname(15).Caption = "Date"
    cval(14).Text = Format(Now, "MM-dd-yyyy")
End Sub
