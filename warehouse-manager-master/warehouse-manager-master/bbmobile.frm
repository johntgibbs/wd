VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bbmobile 
   Caption         =   "Blue Bell Mobile Devices"
   ClientHeight    =   7485
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14010
   LinkTopic       =   "Form8"
   ScaleHeight     =   7485
   ScaleWidth      =   14010
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6855
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   12091
      _Version        =   327680
      BackColorFixed  =   8454016
      FocusRect       =   0
   End
   Begin VB.TextBox Text1 
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
      Left            =   2400
      TabIndex        =   1
      Text            =   "s:\wd\data\mobile.csv"
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label3 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   9960
      TabIndex        =   4
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8040
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Mobile Device Listing:"
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
      Top             =   120
      Width           =   2055
   End
   Begin VB.Menu refmenu 
      Caption         =   "Refresh Data"
   End
   Begin VB.Menu savemenu 
      Caption         =   "Save Data"
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu insrec 
         Caption         =   "Insert New Line"
      End
      Begin VB.Menu delrec 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu edcell 
         Caption         =   "Edit Cell"
      End
   End
   Begin VB.Menu sortmenu 
      Caption         =   "Sort"
      Begin VB.Menu sortser 
         Caption         =   "Serial Number"
      End
      Begin VB.Menu sortloc 
         Caption         =   "Location"
      End
      Begin VB.Menu sortmac 
         Caption         =   "Mac Address"
      End
      Begin VB.Menu sortmod 
         Caption         =   "Model"
      End
      Begin VB.Menu sortdev 
         Caption         =   "Device ID"
      End
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "bbmobile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub delrec_Click()
    If MsgBox("Are you sure?", vbYesNo + vbQuestion, "delete current row...") = vbNo Then Exit Sub
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Grid1.Rows = 1
    End If
End Sub

Private Sub edcell_Click()
    Dim s As String, t As String
    If Grid1.Row < 1 Then Exit Sub
    s = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
    t = Grid1.TextMatrix(0, Grid1.Col)
    s = InputBox(t, t & "....", s)
    If Len(s) = 0 Then Exit Sub
    Grid1.Text = s
    DoEvents
    savemenu_Click
End Sub

Private Sub Form_Load()
    refmenu_Click
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1400
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub insrec_Click()
    Grid1.AddItem " ", Grid1.Row
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rf As String, rh As String
    rt = Me.Caption
    rh = Label1 & ":" & Text1
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
End Sub

Private Sub refmenu_Click()
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim s As String, gs As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7: Grid1.FixedCols = 0
    Open Text1 For Input As #1
    Input #1, f0, f1, f2, f3, f4, f5, f6
    gs = "<" & f0 & "|"
    gs = gs & "<" & f1 & "|"
    gs = gs & "<" & f2 & "|"
    gs = gs & "<" & f3 & "|"
    gs = gs & "<" & f4 & "|"
    gs = gs & "<" & f5 & "|"
    gs = gs & "<" & f6
    
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6
        s = f0 & Chr(9)
        s = s & f1 & Chr(9)
        s = s & f2 & Chr(9)
        s = s & f3 & Chr(9)
        s = s & f4 & Chr(9)
        s = s & f5 & Chr(9)
        s = s & f6 & Chr(9)
        Grid1.AddItem s
    Loop
    Close #1
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 0
    Grid1.Sort = 5
    's = "<Device ID|<Model|<Site|<Relay Server|<Serial Number|<Mac Address|<BB Name"
    Grid1.FormatString = gs
    Grid1.ColWidth(0) = 1800
    Grid1.ColWidth(1) = 1800
    Grid1.ColWidth(2) = 1800
    Grid1.ColWidth(3) = 1800
    Grid1.ColWidth(4) = 1800
    Grid1.ColWidth(5) = 1800
    Grid1.ColWidth(6) = 1800
    'Grid1.ColWidth(0) = 1800
    s = Grid1.Rows - 1
    Label2.Caption = s & " Records"
End Sub

Private Sub savemenu_Click()
    Dim cfile As String, i As Integer, k As Integer, s As String
    If Grid1.Rows = 1 Then Exit Sub
    Open Text1 For Output As #1
    For i = 0 To Grid1.Rows - 1
        For k = 0 To Grid1.Cols - 2
            Write #1, Grid1.TextMatrix(i, k);
        Next k
        Write #1, Grid1.TextMatrix(i, Grid1.Cols - 1)
    Next i
    Close #1
    cfile = "s:\wd\html\mspupdate.csv"
    Open cfile For Output As #1
    For i = 1 To Grid1.Rows - 1
        If Len(Grid1.TextMatrix(i, 4)) >= 14 Then
            'Write #1, "identity.serial";
            'Write #1, Grid1.TextMatrix(i, 4);
            'Write #1, "userAttribute.assetname";
            'Write #1, Grid1.TextMatrix(i, 6)
            s = "identity.serial,"
            s = s & Grid1.TextMatrix(i, 4) & ","
            s = s & "userAttribute.assetname,"
            s = s & Grid1.TextMatrix(i, 6)
            Print #1, s
            
        End If
    Next i
    Close #1
    s = Grid1.Rows - 1
    Label2.Caption = s & " Records"
    Label3.Caption = "New upload file @ " & cfile
End Sub

Private Sub sortdev_Click()
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 0
    Grid1.Sort = 5
End Sub

Private Sub sortloc_Click()
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 6: Grid1.ColSel = 6
    Grid1.Sort = 5
End Sub

Private Sub sortmac_Click()
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 5: Grid1.ColSel = 5
    Grid1.Sort = 5
End Sub

Private Sub sortmod_Click()
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 1: Grid1.ColSel = 4
    Grid1.Sort = 5
End Sub

Private Sub sortser_Click()
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 4: Grid1.ColSel = 4
    Grid1.Sort = 5
End Sub
