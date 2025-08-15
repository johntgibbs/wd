VERSION 5.00
Begin VB.Form pallabprt 
   BackColor       =   &H00404000&
   Caption         =   "pallabprt"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8730
   LinkTopic       =   "pallabprt"
   ScaleHeight     =   6375
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Page"
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   6360
      TabIndex        =   7
      Top             =   720
      Width           =   255
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3240
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   2475
      ScaleWidth      =   4635
      TabIndex        =   5
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label prtdevice 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label ptrig 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   4080
      Width           =   4215
   End
   Begin VB.Label seqlab 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lotlab 
      BackColor       =   &H00FFFFFF&
      Caption         =   "000000 A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label pkglab 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "1/2 GAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   5520
      Width           =   4215
   End
   Begin VB.Label desc1lab 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CARMEL SUNDAE CRUNCH"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5040
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label skulab 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "777"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   4215
   End
End
Attribute VB_Name = "pallabprt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub split50(pd As Control, st As String)
    Dim i As Long, s As String, s1 As String, s2 As String
    For i = 1 To Len(st)
        If Mid(st, i, 1) = " " And i > 1 Then s = Left(st, i - 1)
        If Mid(st, i, 1) = "-" And i > 1 Then s = Left(st, i)
        If Mid(st, i, 1) = "&" And i > 1 Then s = Left(st, i)
        If Mid(st, i, 1) = "+" And i > 1 Then s = Left(st, i)
        If pd.TextWidth(Left(st, i)) > pd.ScaleWidth Then
            s1 = Trim(s)
            s2 = Trim(Right(st, Len(st) - Len(s)))
            Exit For
        End If
    Next i
    If pd.TextWidth(s2) > pd.ScaleWidth Then
        st = s2
        For i = 1 To Len(st)
            If Mid(st, i, 1) = " " And i > 1 Then s = Left(st, i - 1)
            If Mid(st, i, 1) = "-" And i > 1 Then s = Left(st, i)
            If Mid(st, i, 1) = "&" And i > 1 Then s = Left(st, i)
            If Mid(st, i, 1) = "+" And i > 1 Then s = Left(st, i)
            If pd.TextWidth(Left(st, i)) > pd.ScaleWidth Then
                s2 = Trim(s)
                s3 = Trim(Right(st, Len(st) - Len(s)))
                Exit For
            End If
        Next i
        halfwidth = pd.TextWidth(s1) / 2  ' Calculate one-half width.
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
        pd.Print s1
        halfwidth = pd.TextWidth(s2) / 2  ' Calculate one-half width.
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
        pd.Print s2
        halfwidth = pd.TextWidth(s3) / 2  ' Calculate one-half width.
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
        pd.Print s3
    Else
        pd.Print " "
        halfwidth = pd.TextWidth(s1) / 2  ' Calculate one-half width.
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
        pd.Print s1
        halfwidth = pd.TextWidth(s2) / 2  ' Calculate one-half width.
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
        pd.Print s2
    End If
End Sub
Private Sub view_prtall(pd As Control)
    Dim k As Integer, i As Integer, tstat As String
    Dim p As Integer, rstr As String, pflag As Boolean
    Dim halfwidth, sy As Long, s As String
    Dim prop As String
    Screen.MousePointer = 11
    On Error GoTo PrintError
    If TypeOf pd Is PictureBox Then
        rstr = localAppDataPath & "\blnk8x14.bmp"
        pd.Picture = LoadPicture(rstr)
        rstr = Dir(localAppDataPath & "\cic*.bmp")
        Do While Len(rstr) > 0
            Kill localAppDataPath & "\" & rstr
            rstr = Dir
        Loop
        DoEvents
    End If
    prop = "FontName: Arial"
    pd.FontName = "Arial"
    prop = "FontSize: 100"
    pd.FontSize = 100
    prop = "FontBold: True"
    pd.FontBold = True
    
    prop = "CurrentX: 0"
    pd.CurrentX = 0
    prop = "CurrentY: 0"
    pd.CurrentY = 0
    prop = "FontSize: 10"
    pd.FontSize = 10
    
    prop = "CurrentY: 3600"
    pd.CurrentY = 1440 * 2.5
    
    s = "----------------------------------------"
    s = s & " Fold On this Line "
    s = s & "----------------------------------------"
    
    
    halfwidth = pd.TextWidth(s) / 2
    prop = "CurrentX: 2820"
    pd.CurrentX = pd.ScaleWidth / 2 - halfwidth
    pd.Print s
    
    prop = "CurrentY: 5040"
    pd.CurrentY = 1440 * 3.5
    
    prop = "FontBold: True"
    pd.FontBold = True
    prop = "FontSize: 125"
    pd.FontSize = 125   ' Set font size.
    halfwidth = pd.TextWidth(skulab.Caption) / 2  ' Calculate one-half width.
    prop = "CurrentX: 3998"
    pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
    pd.Print skulab.Caption
    prop = "FontSize: 22"
    pd.FontSize = 22
    pd.Print " "
    prop = "FontSize: 80"
    pd.FontSize = 80
    If pd.TextWidth(desc1lab.Caption) < pd.ScaleWidth Then
        pd.Print " "
        pd.Print " "
        halfwidth = pd.TextWidth(desc1lab.Caption) / 2  ' Calculate one-half width.
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
        pd.Print desc1lab.Caption
    Else
        Call split50(pd, desc1lab.Caption)
    End If
    prop = "FontSize: 22"
    pd.FontSize = 22
    pd.Print " "
    prop = "FontSize: 48"
    pd.FontSize = 48
    halfwidth = pd.TextWidth(pkglab.Caption) / 2  ' Calculate one-half width.
    prop = "CurrentX: 4275"
    pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
    pd.Print pkglab.Caption
    'pd.FontSize = 60
    pd.Print " "
    prop = "FontSize: 80"
    pd.FontSize = 80
    sy = pd.CurrentY
    pd.Print lotlab.Caption
    prop = "CurrentY: 16740"
    pd.CurrentY = sy
    prop = "CurrentX: 9480"
    pd.CurrentX = pd.ScaleWidth - pd.TextWidth(seqlab.Caption)
    pd.Print seqlab.Caption
    
    'Bar Code - 39
    s = "!"
    If Len(skulab.Caption) = 4 Then
        s = s & skulab.Caption
    Else
        s = s & skulab.Caption & "="
    End If
    s = s & lotlab.Caption                      'jv070115
    's = s & Left(lotlab.Caption, 6) & "=" & Right(lotlab.Caption, 3) & "="
    's = s & Format(Val(seqlab.Caption), "000") & "!"
    s = s & seqlab.Caption & "!"
    
    prop = "FontName: IDAutomationHC39M"
    pd.FontName = "IDAutomationHC39M"
    prop = "FontSize: 16"
    pd.FontSize = 16
    halfwidth = pd.TextWidth(s) / 2
    prop = "CurrentX: 3255"
    pd.CurrentX = pd.ScaleWidth / 2 - halfwidth
    pd.Print s
    prop = "FontName: Arial"
    pd.FontName = "Arial"
    'End Bar Code - 39

    'If TypeOf pd Is Printer Then pd.EndDoc
    'If TypeOf pd Is Printer Then pd.NewPage
    If TypeOf pd Is PictureBox Then
        If p > 1 Then
            rstr = localAppDataPath & "\cic" & Format(p, "00000") & ".bmp"
            SavePicture pd.Image, rstr
            pd.Picture = LoadPicture(localAppDataPath & "\cic00001.bmp")
            HScroll1.Visible = True
            HScroll1.Value = 1
        Else
            HScroll1.Visible = False
        End If
        'pagelit.Caption = "Page 1 of " & p
        HScroll1.Max = p
    End If
    Screen.MousePointer = 0
    Exit Sub
PrintError:
    eno = Err.Number: edesc = Err.description: Err.Clear
    If MsgBox(edesc & vbCrLf & prop, vbRetryCancel + vbQuestion, "Invalid Printer Property") = vbRetry Then
        Resume
    Else
        End
    End If
End Sub


Private Sub Command2_Click()
    'If Form1.paplegal.Checked = True Then
        Printer.PaperSize = 5
    'Else
    '    Printer.PaperSize = 1
    'End If
    Call view_prtall(Printer)
    Printer.EndDoc
End Sub
Private Sub Form_Load()
    'refresh_grid
    'DoEvents
    'sort_grid (5)
    Picture1.Width = 12240
    'If Form1.paplegal.Checked = True Then
        Picture1.Height = 20160
        If Len(Dir(localAppDataPath & "\blnk8x14.bmp")) = 0 Then
            SavePicture Picture1.Image, localAppDataPath & "\blnk8x14.bmp"
        End If
    'Else
    '    Picture1.Height = 15840
    '    If Len(Dir("c:\blnk8x11.bmp")) = 0 Then
    '        SavePicture Picture1.Image, "c:\blnk8x11.bmp"
    '    End If
    'End If
    Me.Width = 12240 + VScroll1.Width
    'If Len(Dir("c:\blnk8x11.bmp")) = 0 Then
    '    SavePicture Picture1.Image, "c:\blnk8x11.bmp"
    'End If
    'Call view_prtall(Picture1)
    'Picture1.Picture = LoadPicture("s:\bb_cabinet\bin\cabtv.bmp")
End Sub

Private Sub Form_Resize()
    If Me.Height > 2000 Then
        VScroll1.Height = Me.Height - 1250 '1050 '920
        'If Form1.paplegal.Checked = True Then
            VScroll1.Max = 20160 - VScroll1.Height
        'Else
        '    VScroll1.Max = 15840 - VScroll1.Height
        'End If
        VScroll1.SmallChange = Int(VScroll1.Max / 8)
        VScroll1.LargeChange = Int(VScroll1.Max / 3)
        HScroll2.Top = Me.Height - HScroll2.Height * 3 '1050 '920 '880
    End If
    If Me.Width > 2000 Then
        VScroll1.Left = Me.Width - 380
        HScroll2.Width = Me.Width - 380
        HScroll2.Max = 12240 - HScroll2.Width
        If HScroll2.Max > 0 Then
            HScroll2.SmallChange = Int(HScroll2.Max / 8)
            HScroll2.LargeChange = Int(HScroll2.Max / 3)
        End If
        'Frame1.Width = Me.Width
    End If
    If Me.Width > 12000 Then
        HScroll2.Visible = False
    Else
        HScroll2.Visible = True
    End If
End Sub
Private Sub HScroll1_Change()
    'Picture1.Cls
    rstr = localAppDataPath & "\cic" & Format(HScroll1.Value, "00000") & ".bmp"
    If Len(Dir(rstr)) > 0 Then
        Picture1.Picture = LoadPicture(rstr)
        pagelit.Caption = "Page " & HScroll1.Value & " of " & HScroll1.Max
    End If
End Sub

Private Sub HScroll2_Change()
    Picture1.Move 0 - HScroll2.Value
End Sub

Private Sub ptrig_Change()
    Me.Caption = desc1lab.Caption & " " & lotlab.Caption & " " & seqlab.Caption
    'If Form1.paplegal.Checked = True Then
        Me.Caption = Me.Caption & " - Legal 8.5x14"
    'Else
    '    Me.Caption = Me.Caption & " - Letter 8.5x11"
    'End If
    Call view_prtall(Picture1)
    If prtdevice = "Printer" Then Call view_prtall(Printer)
End Sub

Private Sub VScroll1_Change()
    Picture1.Move Picture1.Left, Frame1.Height - VScroll1.Value
End Sub

