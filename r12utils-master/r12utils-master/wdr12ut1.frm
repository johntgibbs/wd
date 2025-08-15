VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "W/D R12 Utilities"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   4335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "2024.02.05"
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label oradb 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label pallogs 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label plantname 
      Caption         =   "Label2"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Plant Code:"
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
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    If Combo1 = "500" Then
        plantname = "Brenham Package Plant"
        pallogs = "\\bbc-01-prodtrk\wd\pallogs\"
    End If
    If Combo1 = "501" Then
        plantname = "Broken Arrow"
        pallogs = "\\bbba-03-dc\f\user\waredist\data\pallogs\"
    End If
    If Combo1 = "502" Then
        plantname = "Sylacauga"
        pallogs = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
    End If
End Sub

Private Sub Form_Load()
    Dim p As String
    Dim site As String
    
    site = UCase(Left(Environ$("computername"), 2))
    If site = "BR" Then
        Open "\\bbc-01-prodtrk\wd\bin\wd.ini" For Input As #1
    ElseIf site = "BA" Then
        Open "\\bbba-03-dc\f\user\waredist\bin\wd.ini" For Input As #1
    ElseIf site = "SY" Then
        Open "\\bbsy-02-dc\f\user\waredist\bin\wd.ini" For Input As #1
    Else
        Open "wd.ini" For Input As #1
    End If
    
    Line Input #1, f
    Do Until EOF(1)
        f = LCase(f): f = Trim(f)
        'If Left$(f, 6) = "plant=" Then p = Right(f, Len(f) - 6)
        If Left$(f, 8) = "plantno=" Then p = Right(f, Len(f) - 8)
        Line Input #1, f
    Loop
    Close #1
    'MsgBox p
    Combo1.Clear
    If p = 50 Then
        Combo1.AddItem "500"
        Combo1.AddItem "501"
        Combo1.AddItem "502"
    End If
    If p = 51 Then Combo1.AddItem "501"
    If p = 52 Then Combo1.AddItem "502"
    Combo1.ListIndex = 0
    List1.AddItem "Branch Trailer Tickets"
    List1.AddItem "Finished Goods Batch Tickets"
    List1.AddItem "Process Order Pick Pallets"
    List1.AddItem "Process Pick Orders"
    If p = 50 Then List1.AddItem "Dry Goods Tickets"
    If p = 50 Then List1.AddItem "Mixer Batch Tickets"
    Form1.oradb = "odbc;database=pbelle;uid=Apps;pwd=pb3113tx;dsn=pbelle"
End Sub

Private Sub Form_Resize()
    List1.Left = (Me.Width - List1.Width) * 0.5
    Label1.Left = List1.Left
    Combo1.Left = Label1.Left + Label1.Width + 50
    plantname.Left = Combo1.Left + Combo1.Width + 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub List1_Click()
    If List1.ListIndex = 0 Then r12trlmonit.Show 'Branch Trailer Tickets
    'If List1.ListIndex = 1 Then r12batpost.Show
    If List1.ListIndex = 1 Then r12wbatpost.Show 'Finished Goods Batch Tickets
    If List1.ListIndex = 2 Then r12oppost.Show 'Process Order Pick Pallets
    If List1.ListIndex = 3 Then r12pickorders.Show 'Process Pick Orders
    If List1.ListIndex = 4 Then r12drytickets.Show 'Dry Goods Tickets
    If List1.ListIndex = 5 Then Form4.Show 'Mixer Batch Tickets
End Sub
