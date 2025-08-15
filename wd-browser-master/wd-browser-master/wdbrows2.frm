VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6300
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7335
   LinkTopic       =   "Form2"
   ScaleHeight     =   6300
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   2778
      _Version        =   327680
      BackColorFixed  =   12640511
      Appearance      =   0
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   3180
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label ycolor 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label1"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label wdfile 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Menu prtlist 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Deactivate()
    Dim i As Integer
    If Form2.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form2" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form2.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form2.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form2.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form2.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form2" Then
            Form2.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form2.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form2.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form2.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
End Sub

Private Sub Form_Resize()
    If Form2.Width > 200 Then List1.Width = Form2.Width - 80
    If Form2.Height > 2000 Then List1.Height = Form2.Height - 650
    If Form2.Width > 200 Then Grid1.Width = Form2.Width - 80
    If Form2.Height > 5000 Then Grid1.Height = Form2.Height - 2650
End Sub

Private Sub prtlist_Click()
    Dim i As Integer
    Printer.FontName = "Courier New"
    Printer.FontSize = 8
    Printer.FontBold = False
    Screen.MousePointer = 11
    For i = 0 To List1.ListCount - 1
        Printer.Print List1.List(i)
    Next i
    Printer.EndDoc
    Screen.MousePointer = 0
End Sub

Private Sub wdfile_Change()
    Dim mpos As Integer, i As Integer, nc As Integer
    Dim rt As String, rf As String, rh As String
    Dim hfile As String
    'If Form1.locdir = "f:\public" Then
    Screen.MousePointer = 11
    'MsgBox Len(Dir("f:\public\wdbrowse.exe"))
    hfile = localAppDataPath & "\htmltemp.htm"
    mpos = InStr(1, wdfile, "goh.", vbBinaryCompare)
    'If LCase(Left(wdfile, 3)) = "goh" Then
    If mpos > 0 Then
        Grid1.Visible = True: List1.FontBold = True: List1.Clear
        Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 4: nc = 0
        If Len(Dir(wdfile)) > 0 Then
            Open wdfile For Input As #1
            Do Until EOF(1)
                Line Input #1, filler
                If Left(filler, 22) = "Oracle Inventory Report" Then
                    s = Right(filler, Len(filler) - 25)
                    List1.AddItem s
                End If
                If Left(filler, 11) = "Total Units" Then
                    If Len(filler) > 25 Then
                        s = Left(filler, 25)
                        List1.AddItem s
                    End If
                End If
                If Left(filler, 12) = "Pallet Space" Then List1.AddItem "----- Pallet Space Summary -----"
                If Left(filler, 12) = "3 Gallons:  " Then List1.AddItem filler
                If Left(filler, 12) = "Other:      " Then List1.AddItem filler
                If Left(filler, 12) = "Total Pallet" Then List1.AddItem filler
                If Left(filler, 12) = "Usable Capac" Then List1.AddItem filler
                If Left(filler, 12) = "Pct.        " Then List1.AddItem filler
                
                If Val(Left(filler, 3)) > 0 Then
                    s = Left(filler, 3) & Chr(9)
                    s = s & Trim(Mid(filler, 4, 45)) & Chr(9)
                    s = s & Trim(Mid(filler, 51, 8)) & Chr(9)
                    s = s & Trim(Mid(filler, 59, 10))
                    If Left(filler, 3) <> "3 G" Then Grid1.AddItem s
                End If
            Loop
            Close #1
            Grid1.FillStyle = flexFillRepeat
            nc = 0
            For i = 1 To Grid1.Rows - 1
                If Val(Grid1.TextMatrix(i, Grid1.Cols - 1)) < 0 Then
                    nc = nc + 1
                    Grid1.Row = i: Grid1.RowSel = i
                    Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                    Grid1.CellBackColor = ycolor.BackColor
                End If
            Next i
            'If nc > 0 Then List1.List(1) = List1.List(1) & "  " & nc & " negative quantities"
        Else
            Grid1.AddItem wdfile & " file not found...."
        End If
        Grid1.FormatString = "^SKU|<Product|^Pallets|^Units"
        Grid1.ColWidth(0) = 1200
        Grid1.ColWidth(1) = 4000
        Grid1.ColWidth(2) = 1200
        Grid1.ColWidth(3) = 1200
        rt = "Oracle Inventory Report"
        rf = List1.List(0)
        If nc > 1 Then
            rf = rf & "<br>" & nc & " items have negative quantities."
        Else
            If nc = 1 Then
                rf = rf & "<br>" & "1 item has a negative quantity."
            End If
        End If
        rh = List1.List(1)
        For i = 2 To List1.ListCount - 1
            rh = rh & "<br>" & List1.List(i)
        Next i
        htdc(0) = "Yellow": gndc(0) = ycolor.BackColor
        'Call htmlcolorgrid(Me, "u:\htmltemp.htm", Grid1, rt, rh, rf, "lemonchiffon", "linen", "white")
        'Form1.WebBrowser1.Navigate "u:\htmltemp.htm"
        Call htmlcolorgrid(Me, hfile, Grid1, rt, rh, rf, "lemonchiffon", "linen", "white")
        Form1.WebBrowser1.Navigate hfile
        
        'If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        '    i = Shell("C:\program files\internet explorer\iexplore.exe u:\htmltemp.htm", vbNormalFocus)
        '    Exit Sub
        'End If
        'If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        '    i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe u:\htmltemp.htm", vbNormalFocus)
        '    Exit Sub
        'End If
        Screen.MousePointer = 0
        Unload Me
    Else
        Grid1.Visible = False: List1.FontBold = False
        List1.Clear
        If Len(Dir(wdfile)) > 0 Then
            Open wdfile For Input As #1
            Do Until EOF(1)
                Line Input #1, filler
                List1.AddItem filler
            Loop
            Close #1
        Else
            List1.AddItem wdfile & " file not found...."
        End If
    End If
    Screen.MousePointer = 0
End Sub

