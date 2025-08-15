VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form pallabcfg 
   Caption         =   "Product Descriptions"
   ClientHeight    =   9930
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11370
   LinkTopic       =   "pallabcfg"
   ScaleHeight     =   9930
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Post to Plants"
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   7680
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   1455
      Left            =   120
      TabIndex        =   22
      Top             =   8400
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2566
      _Version        =   327680
      BackColorFixed  =   12648447
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " View Label "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   5640
      TabIndex        =   14
      Top             =   4320
      Width           =   2775
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   360
         TabIndex        =   21
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   19
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   17
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   16
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save Changes"
      Height          =   495
      Left            =   3480
      TabIndex        =   13
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete SKU"
      Height          =   495
      Left            =   1800
      TabIndex        =   12
      Top             =   6840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New SKU"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   1560
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   1560
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   1560
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4560
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7435
      _Version        =   327680
      BackColorFixed  =   16777152
      FocusRect       =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Menu findsku 
      Caption         =   "Find SKU"
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print List"
   End
End
Attribute VB_Name = "pallabcfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function labfield(lsku As String, lfield As String) As String
    Dim i As Integer
    labfield = " "
    lfield = LCase(lfield)
    If Grid1.Rows <= 2 Then
        refresh_grid1
        DoEvents
    End If
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = lsku Then
            If lfield = "pkg" Then labfield = Grid1.TextMatrix(i, 1)
            If lfield = "name1" Then labfield = Grid1.TextMatrix(i, 2)
            If lfield = "name2" Then labfield = Grid1.TextMatrix(i, 3)
            If lfield = "name3" Then labfield = Grid1.TextMatrix(i, 4)
            Exit For
        End If
    Next i
End Function

Function checkamp(s1 As String) As String
    Dim i As String, s As String, t As String
    i = InStr(1, s1, "&", vbBinaryCompare)
    If i <> 0 Then
        t = s1
        s = Left(t, i) & "&"
        If Len(t) > i Then s = s & Right(t, Len(t) - i)
        s1 = s
    End If
    checkamp = s1
End Function

Private Sub refresh_grid1()
    Dim s As String, cfile As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    cfile = Form1.fmtfile
    'cfile = "U:\jvlook.txt"
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    If Len(Dir(cfile)) = 0 Then
        For i = 100 To 965
            Grid1.AddItem i
        Next i
    Else
        Open cfile For Input As #1
        Do Until EOF(1)
            'Input #1, f0, f1, f2, f3, f4
            's = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & f4
            Line Input #1, s
            Grid1.AddItem s
        Loop
        Close #1
    End If
    Grid1.FormatString = "^SKU|^Package|^Name 1|^Name 2|^Name 3"
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 1600
    Grid1.ColWidth(2) = 1600
    Grid1.ColWidth(3) = 1600
    Grid1.ColWidth(4) = 1600
    Call Grid1_RowColChange
End Sub

Private Sub Command1_Click()
    Dim s As String, i As Integer
    s = InputBox("SKU:", "New SKU......", "")
    If Len(s) = 0 Then Exit Sub
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = s Then
            MsgBox s & " is already being used...", vbOKOnly + vbInformation, "try again..."
            Exit Sub
        End If
    Next i
    Grid1.AddItem s
    Grid1.Row = Grid1.Rows - 1
End Sub

Private Sub Command2_Click()
    Dim s As String
    If Grid1.Row = 0 Then Exit Sub
    s = Grid1.TextMatrix(Grid1.Row, 0)
    If MsgBox("Ok to delete SKU: " & s, vbYesNo + vbQuestion, "are you sure..") = vbNo Then Exit Sub
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Grid1.Rows = 1
    End If
End Sub

Private Sub Command3_Click()
    Dim cfile As String, i As Integer, s As String
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 0
    Grid1.Sort = 5
    cfile = Form1.fmtfile
    'cfile = "U:\jvlook.txt"
    Open cfile For Output As #1
    For i = 1 To Grid1.Rows - 1
        s = Grid1.TextMatrix(i, 0) & Chr(9)
        s = s & Grid1.TextMatrix(i, 1) & Chr(9)
        s = s & Grid1.TextMatrix(i, 2) & Chr(9)
        s = s & Grid1.TextMatrix(i, 3) & Chr(9)
        s = s & Grid1.TextMatrix(i, 4)
        Print #1, s
        'Write #1, Grid1.TextMatrix(i, 0);
        'Write #1, Grid1.TextMatrix(i, 1);
        'Write #1, Grid1.TextMatrix(i, 2);
        'Write #1, Grid1.TextMatrix(i, 3);
        'Write #1, Grid1.TextMatrix(i, 4)
    Next i
    Close #1
    If Form1.Caption = "SKU Master Maintenance" Then
        'MsgBox "labpics"
        Call load_labpics
    End If
End Sub

Private Sub Command4_Click()
    Dim i As Long, brendir As String
    If pgrid.Rows < 2 Then
        MsgBox "Don't know where to post..."
        Exit Sub
    Else
        brendir = pgrid.TextMatrix(1, 1)
    End If
    If Grid1.Rows < 5 Then
        MsgBox "No finished goods"
        Exit Sub
    End If
    Screen.MousePointer = 11
    'Open brendir For Output As #1
    'For i = 1 To Grid1.Rows - 1
    '    Write #1, Grid1.TextMatrix(i, 0);
    '    Write #1, Grid1.TextMatrix(i, 1);
    '    Write #1, Grid1.TextMatrix(i, 2);
    '    Write #1, Grid1.TextMatrix(i, 3)
    'Next i
    'Close #1
    Command3_Click
    If pgrid.Rows > 2 Then
        For i = 2 To pgrid.Rows - 1
            FileCopy brendir, pgrid.TextMatrix(i, 1)
        Next i
    End If
    Screen.MousePointer = 0
End Sub

Private Sub findsku_Click()
    Dim s As String
    s = Grid1.TextMatrix(Grid1.Row, 0)
    s = InputBox("SKU:", "Find SKU.....", s)
    If Len(s) = 0 Then Exit Sub
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = s Then
            Grid1.Row = i
            Grid1.TopRow = i
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Load()
    refresh_grid1
    pgrid.Clear: pgrid.Cols = 2: pgrid.Rows = 1
    pgrid.AddItem "Brenham" & Chr(9) & "s:\wd\bin\labfmt.txt"
    'pgrid.AddItem "TX Cycle" & Chr(9) & "s:\wd\counts\txcycle\labfmt.txt"
    'pgrid.AddItem "Snack Plant" & Chr(9) & "s:\wd\counts\snackplant\labfmt.txt"
    'pgrid.AddItem "TX Dmgs" & Chr(9) & "s:\wd\counts\txdmgs\labfmt.txt"
    pgrid.AddItem "Broken Arrow" & Chr(9) & "\\bbba-02-dc\f\user\waredist\bin\labfmt.txt"
    pgrid.AddItem "Sylacauga" & Chr(9) & "\\bbsy-02-dc\f\user\waredist\bin\labfmt.txt"
    pgrid.FormatString = "<Plant|<File Name"
    pgrid.ColWidth(0) = 1200
    pgrid.ColWidth(1) = 6000
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    i = Grid1.Col
    'MsgBox KeyAscii & " = " & Chr(KeyAscii)
    If KeyAscii > 31 And KeyAscii < 128 Then
        If Text1(i).Enabled Then
            Text1(i).Text = Text1(i).Text & Chr(KeyAscii)
        End If
    End If
    If KeyAscii = 8 Then
        If Text1(i).Enabled Then
            If Len(Text1(i).Text) <= 1 Then
                Text1(i).Text = ""
            Else
                Text1(i).Text = Left(Text1(i).Text, Len(Text1(i).Text) - 1)
            End If
        End If
    End If

End Sub

Private Sub Grid1_RowColChange()
    Dim i As Integer
    If Grid1.Row = 0 Then Exit Sub
    For i = 0 To Grid1.Cols - 1
        Label1(i).Caption = Grid1.TextMatrix(0, i)
        Text1(i).Text = Grid1.TextMatrix(Grid1.Row, i)
    Next i
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rf As String, rh As String
    rt = "Pallet Labels"
    rh = Me.Caption
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, htmlTempFile, Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & htmlTempFile, vbNormalFocus)
            Exit Sub
        End If
    End If

End Sub

Private Sub Text1_Change(Index As Integer)
    Grid1.TextMatrix(Grid1.Row, Index) = Text1(Index).Text
    'Label2(Index).Caption = Text1(Index).Text
    Label2(Index).Caption = checkamp(Text1(Index).Text)
    If Val(Label2(6).Caption) < 80 Then
        Label2(6).Caption = Val(Label2(6).Caption) + 1
    Else
        Label2(6).Caption = "1"
    End If
    Label2(5).Caption = Format(Now, "MMddyy") & " A"
End Sub
