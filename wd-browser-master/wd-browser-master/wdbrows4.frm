VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Notes to Brenham"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   LinkTopic       =   "Form4"
   ScaleHeight     =   5220
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
   Begin VB.Label Label2 
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
      Left            =   1440
      TabIndex        =   3
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Branch:"
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
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Label brcode 
      Caption         =   "00"
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
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub brcode_Change()
    Dim flen As Long, t1 As String
    Text1 = ""
    Label2.Caption = branchrec(Val(brcode)).branchname
    If Len(Dir(Form1.webdir & "\orders\notes." & brcode)) > 0 Then
        flen = FileLen(Form1.webdir & "\orders\notes." & brcode)
        Open Form1.webdir & "\orders\notes." & brcode For Input As #1
        t1 = Input(flen, #1)
        Close #1
        Text1 = Trim(t1)
    End If
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Val(brcode) > 0 Then
        Open Form1.webdir & "\orders\notes." & Format(Val(brcode), "00") For Output As #1
        Print #1, Trim(Text1);
        Close #1
    End If
    If Form2.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form4" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form4.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form4.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form4.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form4.Width
                Exit For
            End If
        Next i
    End If
End Sub
Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form4" Then
            Form4.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form4.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form4.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form4.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
End Sub

Private Sub Form_Resize()
    Text1.Width = Form4.Width - 80
    If Form4.Height > 2000 Then
        Text1.Height = Form4.Height - 600
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub
