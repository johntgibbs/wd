VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form proadj 
   Caption         =   "Oracle Adjustment Requests"
   ClientHeight    =   6840
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11085
   LinkTopic       =   "Form3"
   ScaleHeight     =   6840
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Format Dates"
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   2760
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   7223
      _Version        =   327680
      FixedCols       =   0
      BackColorFixed  =   128
      ForeColorFixed  =   16777088
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4260
      _Version        =   327680
      Cols            =   3
      ForeColor       =   8388736
      BackColorFixed  =   8454016
      BackColorSel    =   8388736
      FocusRect       =   0
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Invalid Date Format"
      Height          =   255
      Left            =   7080
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label wfile 
      Caption         =   "...."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   6735
   End
   Begin VB.Menu prtmenu 
      Caption         =   "&Print"
   End
   Begin VB.Menu procmenu 
      Caption         =   "Proc&ess"
      Begin VB.Menu reflist 
         Caption         =   "Refresh List"
      End
      Begin VB.Menu posto 
         Caption         =   "Post to Oracle"
      End
      Begin VB.Menu canfile 
         Caption         =   "Cancel File"
      End
   End
End
Attribute VB_Name = "proadj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid2()
    Dim f0 As String, f1 As String, f2 As String
    Dim f3 As String, f4 As String, f5 As String
    Dim f6 As String, f7 As String, f8 As String
    Dim f9 As String, s As String
    Screen.MousePointer = 11
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 10
    If Len(Dir(wfile)) > 0 Then
        Open wfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
            s = f0 & Chr(9)
            s = s & f1 & Chr(9)
            s = s & f2 & Chr(9)
            s = s & f3 & Chr(9)
            s = s & f4 & Chr(9)
            s = s & f5 & Chr(9)
            s = s & f6 & Chr(9)
            If UCase(f7) = "POST" Then f7 = "CYCL"
            s = s & f7 & Chr(9)
            s = s & f8 & Chr(9)
            s = s & f9
            Grid2.AddItem s
        Loop
        Close #1
    End If
    Grid2.FormatString = "^Tran Date|^Whs|^Locn|<Lot|<Item|<Description|^Qty|^Reason|^User|^Entry Date"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 600
    Grid2.ColWidth(2) = 600
    Grid2.ColWidth(3) = 600
    Grid2.ColWidth(4) = 600
    Grid2.ColWidth(5) = 3000
    Grid2.ColWidth(6) = 800
    Grid2.ColWidth(7) = 800
    Grid2.ColWidth(8) = 1600
    Grid2.ColWidth(9) = 1100
    ycolor.Visible = False
    Command1.Visible = False
    Grid2.FillStyle = flexFillRepeat
    For i = 1 To Grid2.Rows - 1
        If Format$(Grid2.TextMatrix(i, 0), "m-d-yyyy") <> Grid2.TextMatrix(i, 0) Then
            Grid2.Row = i: Grid2.RowSel = i
            Grid2.Col = 0: Grid2.ColSel = Grid2.Cols - 1
            Grid2.CellBackColor = ycolor.BackColor
            Grid2.CellForeColor = ycolor.ForeColor
            ycolor.Visible = True
            Command1.Visible = True
        End If
    Next i
            
    Screen.MousePointer = 0
End Sub

Private Sub canfile_Click()
    If Grid1.Rows < 2 Then Exit Sub
    If Len(Dir(wfile)) = 0 Then Exit Sub
    If MsgBox("Ok to remove " & wfile & "?", vbYesNo + vbQuestion, "are you sure...") = vbNo Then Exit Sub
    Kill wfile
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call reflist_Click
    End If
End Sub

Private Sub Command1_Click()
    Dim i As Long, k As Integer
    If Grid2.Rows < 2 Then Exit Sub
    For i = 1 To Grid2.Rows - 1
        Grid2.TextMatrix(i, 0) = Format(Grid2.TextMatrix(i, 0), "m-d-yyyy")
    Next i
    Open wfile For Output As #1
    For i = 1 To Grid2.Rows - 1
        For k = 0 To Grid2.Cols - 2
            Write #1, Grid2.TextMatrix(i, k);
        Next k
        Write #1, Grid2.TextMatrix(i, Grid2.Cols - 1)
    Next i
    Close #1
    DoEvents
    refresh_grid2
End Sub

Private Sub Form_Load()
    reflist_Click
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    Grid2.Width = Me.Width - 100
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu procmenu
End Sub

Private Sub Grid1_RowColChange()
    wfile.Caption = Form1.webdir & "\counts\" & Grid1.TextMatrix(Grid1.Row, 0)
End Sub

Private Sub posto_Click()
    Dim i As Long, k As Integer, ofile As String, sfile As String
    Dim fdate As String, fsz As String
    fdate = Format(FileDateTime(wfile), "mm-dd-yyyy hh:mm am/pm")
    fsz = FileLen(wfile)
    If fdate <> Grid1.TextMatrix(Grid1.Row, 1) Then
        MsgBox "File date/time has changed.  Please review the file contents again before posting.", vbOKOnly + vbExclamation, "change detected..."
        reflist_Click
        DoEvents
        refresh_grid2
        Exit Sub
    End If
    If fsz <> Grid1.TextMatrix(Grid1.Row, 2) Then
        MsgBox "File size has changed.  Please review the file contents again before posting.", vbOKOnly + vbExclamation, "change detected..."
        reflist_Click
        DoEvents
        refresh_grid2
        Exit Sub
    End If
    
    ofile = Form1.webdir & "\counts\whseadj." & Right(wfile, 3)
    'MsgBox ofile
    If ycolor.Visible Then
        MsgBox "Date formats may fail.  Please re-format dates and try again."
    Else
        Name wfile As ofile
    End If
    
    sfile = Form1.webdir & "\counts\r12adjs.win"
    Open sfile For Output As #1
    Print #1, "open pbelle.bluebell.com"
    Print #1, "infbbcri"
    Print #1, "welcome@2023"
    Print #1, "BINARY"
    'Print #1, "cd /interface/infbbcri/PBELLE/incoming"
    Print #1, "cd PBELLE/incoming"
    Print #1, "lcd " & Form1.webdir & "\counts"
    Print #1, "put whseadj." & Right(wfile, 3) & " whseadj." & Right(wfile, 3)
    Print #1, "close"
    Print #1, "bye"
    Close #1
    ftpexe = "c:\windows\system32\ftp.exe"
    x = Shell(ftpexe & " -s:" & sfile, vbNormalFocus)
    MsgBox ftpexe & " -s:" & sfile
    
    
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
        Grid1_RowColChange
    Else
        Grid1.Rows = 1
    End If
    
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    Screen.MousePointer = 11
    rt = Me.Caption
    rh = wfile
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Call printflexgrid(Printer, Grid2, rt, rh, rf)
    Screen.MousePointer = 0
End Sub

Private Sub reflist_Click()
    Dim spath As String, sdir As String, sqlx As String
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Cols = 3: Grid1.Rows = 1
    spath = Form1.webdir & "\counts\whsadj.*"
    sdir = Dir$(spath)
    Do While sdir <> ""
        sqlx = sdir & Chr(9)
        sqlx = sqlx & Format(FileDateTime(Form1.webdir & "\counts\" & sdir), "mm-dd-yyyy hh:mm am/pm") & Chr(9)
        sqlx = sqlx & FileLen(Form1.webdir & "\counts\" & sdir)
        Grid1.AddItem sqlx
        sdir = Dir$
    Loop
    Grid1.FormatString = "^File|^Time Created|^Size"
    Grid1.ColWidth(0) = 1400
    Grid1.ColWidth(1) = 1800
    Grid1.ColWidth(2) = 1200
    Screen.MousePointer = 0
    Grid1_RowColChange
End Sub

Private Sub wfile_Change()
    refresh_grid2
End Sub

