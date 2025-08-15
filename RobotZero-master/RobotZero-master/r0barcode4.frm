VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Return to Wrapper"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   ScaleHeight     =   6915
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Return to Wrapper"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   3255
   End
   Begin VB.TextBox Text1 
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
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Label emess 
      Alignment       =   2  'Center
      Caption         =   "Label2"
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
      Left            =   1440
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Barcode:"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Call return_to_wrapper(Text1, Form1.userid, "ROBOT ZERO", "0")
    emess = return_to_wrapper(Text1, Form1.userid, "ROBOT ZERO", "0")
    emess.Visible = True
    Text1 = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Text1 = ""
    Command1.Enabled = False
End Sub

Private Sub Form_Resize()
    Label1.Left = (Me.Width - Label1.Width) * 0.5
    Text1.Left = Label1.Left
    Command1.Left = Label1.Left '(Me.Width - Command1.Width) * 0.5
    emess.Left = Label1.Left
End Sub

Private Sub Text1_Change()
    If Len(Text1) > 15 Then
        Text1 = UCase(Text1)
        If barcode_profile(Text1) = True Then
            Command1.Enabled = True
            emess.Visible = False
            Command1.SetFocus
        Else
            emess = "Invalid Barcode.."
            emess.Visible = True
        End If
    Else
        Command1.Enabled = False
    End If
End Sub

