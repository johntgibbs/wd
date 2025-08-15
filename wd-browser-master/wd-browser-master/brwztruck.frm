VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form brwztruck 
   Caption         =   "PLTTRUCK"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13260
   LinkTopic       =   "Form13"
   ScaleHeight     =   7275
   ScaleWidth      =   13260
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1935
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
      _Version        =   327680
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Paste Back Original"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3495
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6165
      _Version        =   327680
      BackColorFixed  =   16777152
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
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
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
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
      Left            =   4080
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label schdir 
      Caption         =   "s:\wd\html\schedule\bawks."
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   6480
      Width           =   3855
   End
   Begin VB.Label orgdir 
      Caption         =   "s:\wd\html\schedule\baorg."
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   5880
      Width           =   3495
   End
   Begin VB.Label brcode 
      Caption         =   "brcode"
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Changes"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Week #:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "brwztruck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String
Private Sub refresh_grid()
    Dim filler As String, i As Integer, k As Integer
    Label2.Visible = False: Command1.Visible = False
    Command2.Visible = False
    Grid1.Font = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 8
    If Len(Dir(schdir.Caption & Combo1)) = 0 Then
        MsgBox "File does not exist for week " & Combo1 & ".", vbOKOnly + vbInformation, "Sorry, try another week.."
        Exit Sub
    End If
    If FileLen(schdir.Caption & Combo1) = 0 Then
        MsgBox "File exists, but it is blank.", vbOKOnly + vbInformation, "Sorry, try again..."
        Exit Sub
    End If
    If Len(Dir(orgdir.Caption & Combo1)) = 0 Then
        MsgBox "File does not exist for week " & Combo1 & ".", vbOKOnly + vbInformation, "Sorry, try another week.."
        Exit Sub
    End If
    If FileLen(orgdir.Caption & Combo1) = 0 Then
        MsgBox "File exists, but it is blank.", vbOKOnly + vbInformation, "Sorry, try again..."
        Exit Sub
    End If
    
    Grid1.Rows = 1
    Grid1.FormatString = "^|^|^|^|^|^|^|^"
    Open schdir.Caption & Combo1 For Input As #1
    Do Until EOF(1)
        Line Input #1, filler
        Grid1.AddItem filler
    Loop
    Close #1
    Grid1.ColWidth(0) = 2000
    For i = 1 To 6
        'Grid1.ColWidth(i) = 1500
        Grid1.ColWidth(i) = 1800
    Next i
    Grid1.ColWidth(Grid1.Cols - 1) = 1
    If Grid1.Rows > 1 Then
        Grid1.FixedRows = 1
        'Grid1.RowHeight(-1) = Grid1.RowHeight(1) * 2
        Grid1.RowHeight(-1) = Grid1.RowHeight(1) * 3
        For i = 0 To Grid1.Cols - 1
            Grid1.TextMatrix(0, i) = Grid1.TextMatrix(1, i)
        Next i
        Grid1.RemoveItem 1
    End If
    
    Grid2.Rows = 1
    Grid2.FormatString = "^|^|^|^|^|^|^|^"
    Open orgdir.Caption & Combo1 For Input As #1
    Do Until EOF(1)
        Line Input #1, filler
        Grid2.AddItem filler
    Loop
    Close #1
    Grid2.ColWidth(0) = 2000
    For i = 1 To 6
        Grid2.ColWidth(i) = 1500
    Next i
    Grid2.ColWidth(Grid2.Cols - 1) = 1
    If Grid2.Rows > 1 Then
        Grid2.FixedRows = 1
        Grid2.RowHeight(-1) = Grid2.RowHeight(1) * 2
        For i = 0 To Grid1.Cols - 1
            Grid2.TextMatrix(0, i) = Grid2.TextMatrix(1, i)
        Next i
        Grid2.RemoveItem 1
    End If
        
    If Grid1.Rows = Grid2.Rows And Grid1.Cols = Grid2.Cols Then
        For i = 0 To Grid1.Rows - 1
            For k = 0 To Grid1.Cols - 1
                If Grid1.TextMatrix(i, k) <> Grid2.TextMatrix(i, k) Then
                    Grid1.Row = i: Grid1.Col = k
                    Grid1.CellBackColor = Label2.BackColor
                    Label2.Visible = True
                End If
            Next k
        Next i
    End If
    If Grid1.Rows > 1 Then
        Grid1.Row = 1: Grid1.Col = 1
        Command2.Visible = True
    End If
    Me.Caption = Grid1.TextMatrix(0, 0) & " originally sent: " & Format(FileDateTime(Me.orgdir & Combo1), "ddd m-d-yy h:mm am/pm")
End Sub

Private Sub brcode_Change()
    Dim spath As String, sdir As String, w As Integer, s As String
    Dim i As Integer
    If Me.brcode = "500" Then
        Me.schdir = Form1.webdir & "\schedule\txwks."
        Me.orgdir = Form1.webdir & "\schedule\txorg."
    End If
    If Me.brcode = "501" Then
        Me.schdir = Form1.webdir & "\schedule\bawks."
        Me.orgdir = Form1.webdir & "\schedule\baorg."
    End If
    If Me.brcode = "502" Then
        Me.schdir = Form1.webdir & "\schedule\sywks."
        Me.orgdir = Form1.webdir & "\schedule\syorg."
    End If
    Combo1.Clear
    spath = Me.schdir & "*"
    sdir = Dir$(spath)
    Do While sdir <> ""
        If Val(Right(sdir, 2)) > 0 Then Combo1.AddItem Right(sdir, 2)
        sdir = Dir$
    Loop
    s = "1-1-" & Format(Now, "yyyy")
    w = DateDiff("ww", s, Now, , vbFirstJan1) + 1
    If Combo1.ListCount > 0 Then
        For i = 0 To Combo1.ListCount - 1
            If Val(Combo1.List(i)) = w Then
                Combo1.ListIndex = i
                Exit For
            End If
        Next i
        If Val(Combo1.List(Combo1.ListIndex)) <> w Then
            Combo1 = Combo1.List(Combo1.ListCount - 1)
            refresh_grid
        End If
    End If
    
End Sub

Private Sub Combo1_Click()
    refresh_grid
    'DoEvents
    'Grid1.SetFocus
End Sub

Private Sub Command1_Click()
    Dim i As Integer, k As Integer, lo As String
    k = Len(Grid1.TextMatrix(0, 0))
    If Left(Me.Caption, k) <> Grid1.TextMatrix(0, 0) Then Exit Sub
    'MsgBox "saving"
    Screen.MousePointer = 11
    Open Me.schdir & Combo1 For Output As #1
    For i = 0 To Grid1.Rows - 1
        lo = ""
        For k = 0 To Grid1.Cols - 2
            lo = lo & Grid1.TextMatrix(i, k) & Chr(9)
        Next k
        Print #1, lo
    Next i
    Close #1
    Screen.MousePointer = 0
    Command1.Visible = False
End Sub

Private Sub Command2_Click()
    Dim i As Integer, k As Integer
    Dim lo As String, lw As String, dc As Integer
    If Grid1.Rows < 2 Then Exit Sub
    If Command1.Visible = True Then
        Call Command1_Click
        lo = "Changes Made to the original schedule have been saved."
        MsgBox lo, vbOKOnly + vbInformation, "Changes detected..."
        'Command1.Visible = False
    End If
    Screen.MousePointer = 11
    dc = 0
    Printer.Font = "Courier New"
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Orientation = 2
    lo = Grid1.TextMatrix(0, 0) & "   Week: " & Combo1
    Printer.Print lo
    Printer.Print " "
    Printer.FontSize = 7
    lo = Space(20)
    For i = 1 To 6
        lo = lo & Grid1.TextMatrix(0, i)
        lo = lo & Space(25 - Len(Grid1.TextMatrix(0, i)))
    Next i
    Printer.Print lo
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) > ".." Then
            If dc > 12 Then
                Printer.NewPage
                Printer.FontSize = 12
                Printer.FontBold = True
                lo = Grid1.TextMatrix(0, 0) & "   Week: " & Combo1
                Printer.Print lo
                Printer.Print " "
                Printer.FontSize = 7
                lo = Space(20)
                For k = 1 To 6
                    lo = lo & Grid1.TextMatrix(0, k)
                    lo = lo & Space(25 - Len(Grid1.TextMatrix(0, k)))
                Next k
                Printer.Print lo
                dc = 0
            End If
            Printer.Print Grid1.TextMatrix(i, 0)
            dc = dc + 1
        End If
        If Val(Grid1.TextMatrix(i, 0)) > 0 Then
            lo = Grid1.TextMatrix(i, 0) & Space(20 - Len(Grid1.TextMatrix(i, 0)))
        Else
            lo = Space(20)
        End If
        For k = 1 To 6
            lw = Left(Grid1.TextMatrix(i, k), 24)
            lo = lo & lw & Space(25 - Len(lw))
        Next k
        Printer.Print lo
    Next i
    Printer.Print " "
    lo = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    lo = lo & "  " & Me.schdir & Combo1
    Printer.Print lo
    Printer.EndDoc
    Screen.MousePointer = 0
        
End Sub

Private Sub Command3_Click()
    If Grid1.Row = Grid2.Row And Grid1.Col = Grid2.Col Then
        Grid1.Text = Grid2.Text
        Grid1_LeaveCell
    End If
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
    Grid1.Width = Me.Width - 80
    If Me.Height > 3000 Then
        'Grid1.Height = Me.Height - 680
        Grid1.Height = Me.Height - (Command1.Height * 2)
    End If
End Sub

Private Sub Grid1_GotFocus()
    Grid1.FocusRect = flexFocusNone
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    If Grid1.Col < 1 Then Exit Sub
    If Grid1.Row < 1 Then Exit Sub
    If KeyAscii = 8 Then
        If Len(Grid1.Text) > 1 Then
            Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
            Grid1_RowColChange
        Else
            Grid1.Text = ""
        End If
        edcell = Grid1.TextMatrix(0, Grid1.Col)
        Command1.Visible = True
    Else
        If KeyAscii >= 31 And KeyAscii <= 127 Then
            If edcell = "" Then
                Grid1.Text = ""
                Grid1_RowColChange
            End If
            Grid1.Text = Grid1.Text & Chr(KeyAscii)
            edcell = Grid1.TextMatrix(0, Grid1.Col)
            Command1.Visible = True
        End If
    End If
End Sub

Private Sub Grid1_LeaveCell()
    If Grid2.Row = Grid1.Row And Grid2.Col = Grid1.Col Then
        If Grid2.Text <> Grid1.Text Then
            Grid1.CellBackColor = Label2.BackColor
            Label2.Visible = True
            'Command1.Visible = True
        Else
            Grid1.CellBackColor = Grid1.BackColor
        End If
    End If
    edcell = ""
End Sub

Private Sub Grid1_LostFocus()
    Grid1_LeaveCell
    Grid1.FocusRect = flexFocusHeavy
End Sub

Private Sub Grid1_RowColChange()
    If Grid1.Cols = Grid2.Cols And Grid1.Rows = Grid2.Rows Then
        Grid2.Row = Grid1.Row
        Grid2.Col = Grid1.Col
        If LCase(Grid1.Text) <> LCase(Grid2.Text) Then
            Label3.Caption = "originally: " & Grid2.Text
            Grid1.ToolTipText = "originally: " & Grid2.Text
            Command3.Visible = True
        Else
            Label3.Caption = ""
            Grid1.ToolTipText = ""
            Command3.Visible = False
        End If
    End If
End Sub



