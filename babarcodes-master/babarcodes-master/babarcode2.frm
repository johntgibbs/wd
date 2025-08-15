VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "BA Wrapper Menu"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   LinkTopic       =   "Form2"
   ScaleHeight     =   5820
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Options"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub list1_KeyPress(KeyAscii As Integer)
    'MsgBox KeyAscii & " - " & list1
    If KeyAscii = 27 Then Unload Me
    If KeyAscii = 13 Then
        If List1 = "Blue Bell Pallets" Then Form7.Show
        'If List1 = "Move From Rack" Then Form3.Show
        If List1 = "Return to Wrapper" Then Form4.Show
        'If List1 = "Change Plate" Then Form5.Show
        If List1 = "Exit" Then Unload Me
    End If
End Sub

Private Sub Form_Load()
    List1.Clear
    List1.AddItem "Blue Bell Pallets"
    'List1.AddItem "Move From Rack"
    List1.AddItem "Return to Wrapper"
    'List1.AddItem "Change Plate"
    List1.AddItem "Exit"
    List1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Label1.Left = (Me.Width - Label1.Width) * 0.5
    List1.Left = Label1.Left
End Sub
