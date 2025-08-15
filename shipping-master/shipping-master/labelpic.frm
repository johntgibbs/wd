VERSION 5.00
Begin VB.Form labelpic 
   Caption         =   "Form4"
   ClientHeight    =   8835
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9360
   LinkTopic       =   "Form4"
   ScaleHeight     =   8835
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   4080
      Width           =   2415
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3255
      Left            =   5520
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3225
      ScaleWidth      =   5265
      TabIndex        =   2
      Top             =   720
      Width           =   5295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label pagelit 
         Alignment       =   2  'Center
         Caption         =   "page"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Label labpt 
      Caption         =   "labpt"
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Label ptrig 
      Caption         =   "ptrig"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.Menu sizemenu 
      Caption         =   "Paper Size"
      Begin VB.Menu paplegal 
         Caption         =   "Legal"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "labelpic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub split50(pd As Control, st As String)
    Dim i As Long, s As String, s1 As String, s2 As String
    For i = 1 To Len(st)
        If mid(st, i, 1) = " " And i > 1 Then s = Left(st, i - 1)
        If mid(st, i, 1) = "-" And i > 1 Then s = Left(st, i)
        If mid(st, i, 1) = "&" And i > 1 Then s = Left(st, i)
        If mid(st, i, 1) = "+" And i > 1 Then s = Left(st, i)
        If pd.TextWidth(Left(st, i)) > pd.ScaleWidth Then
            s1 = Trim(s)
            s2 = Trim(Right(st, Len(st) - Len(s)))
            Exit For
        End If
    Next i
    If pd.TextWidth(s2) > pd.ScaleWidth Then
        st = s2
        For i = 1 To Len(st)
            If mid(st, i, 1) = " " And i > 1 Then s = Left(st, i - 1)
            If mid(st, i, 1) = "-" And i > 1 Then s = Left(st, i)
            If mid(st, i, 1) = "&" And i > 1 Then s = Left(st, i)
            If mid(st, i, 1) = "+" And i > 1 Then s = Left(st, i)
            If pd.TextWidth(Left(st, i)) > pd.ScaleWidth Then
                s2 = Trim(s)
                s3 = Trim(Right(st, Len(st) - Len(s)))
                Exit For
            End If
        Next i
        'MsgBox s1 & " + " & s2 & " + " & s3
        'pd.FontSize = 48
        halfwidth = pd.TextWidth(s1) / 2  ' Calculate one-half width.
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
        pd.Print s1
        halfwidth = pd.TextWidth(s2) / 2  ' Calculate one-half width.
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
        pd.Print s2
        'If Left(Me.pkglab, Len(s3)) <> s3 Then
        '    Me.pkglab = s3 & " " & Me.pkglab
        'End If
        halfwidth = pd.TextWidth(s3) / 2  ' Calculate one-half width.
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
        pd.Print s3
    Else
        'MsgBox s1 & " + " & s2
        pd.Print " "
        halfwidth = pd.TextWidth(s1) / 2  ' Calculate one-half width.
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
        pd.Print s1
        halfwidth = pd.TextWidth(s2) / 2  ' Calculate one-half width.
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
        pd.Print s2
    End If
End Sub
Private Sub view_prtlist(pd As Control)
    Dim k As Integer, i As Integer, tstat As String
    Dim p As Integer, rstr As String, pflag As Boolean
    Dim halfwidth, sy As Long, s As String
    Dim palid As String, bno As Integer, spal
    Screen.MousePointer = 11
    If TypeOf pd Is PictureBox Then
        If Me.paplegal.Checked = True Then
            rstr = localAppDataPath & "\blnk8x14.bmp"
        Else
            rstr = localAppDataPath & "\blnk8x11.bmp"
        End If
        pd.Picture = LoadPicture(rstr)
        rstr = Dir(localAppDataPath & "\cic*.bmp")
        Do While Len(rstr) > 0
            Kill localAppDataPath & "\" & rstr
            rstr = Dir
        Loop
        DoEvents
    End If
    pd.FontName = "Arial"
    pd.FontSize = 12
    pd.FontBold = True
    
    pd.CurrentX = 0: pd.CurrentY = 0
    If Me.paplegal.Checked = True Then
        pd.FontSize = 12
        pd.FontUnderline = False
        pd.CurrentY = 1440 * 1
        pd.CurrentX = 1440 * 0.75: pd.Print "Partial Pallet Order";
        pd.CurrentX = 1440 * 2.75: pd.Print partlabs.Combo2;
        pd.CurrentX = 1440 * 6.25: pd.Print Left(partlabs.Combo1, 10)
        s = partlabs.Combo2
        
        pd.FontSize = 10
        pd.CurrentY = 1440 * 1.5
        pd.CurrentX = 1440 * 0.75:  pd.Print "Pallet";
        pd.CurrentX = 1440 * 1.25: pd.Print "SKU  Product";
        pd.CurrentX = 1440 * 4.25: pd.Print "Wraps";
        pd.CurrentX = 1440 * 5.25: pd.Print "Code Date(s)"
    Else
        pd.FontSize = 10
        pd.Print " "
        'pd.CurrentY = 15480 / 11 * 2.5
    End If
    pd.FontBold = False
    spal = "..."
    For i = 1 To partlabs.Grid1.Rows - 1
        If Val(partlabs.Grid1.TextMatrix(i, 0)) > 0 Then
        
            If Left(partlabs.List1, 1) = "T" Then
                bno = 16
                palid = "!"
                palid = palid & Format(bno, "00")               'branch
                palid = palid & Right(partlabs.List1, 6)         'account
                palid = palid & Format(Val(partlabs.Grid1.TextMatrix(i, 0)), "00")  'palnum
                palid = palid & Left(partlabs.Combo1, 2) & mid(partlabs.Combo1, 4, 2) & mid(partlabs.Combo1, 9, 2)
                palid = palid & "!"
            Else
                bno = Val(Left(partlabs.List1, 2))
                palid = "!"
                palid = palid & Format(bno, "000") & "="
                palid = palid & Left(partlabs.Combo1, 2) & mid(partlabs.Combo1, 4, 2) & mid(partlabs.Combo1, 9, 2) & "=B="
                palid = palid & Format(Val(partlabs.Grid1.TextMatrix(i, 0)), "000")
                palid = palid & "!"
            End If
            
            If palid <> spal Then
                pd.Print " "
                pd.Font = "IDAutomationHC39M"
                pd.Print palid
                spal = palid
                pd.Font = "Arial"
            End If
        
            pd.FontSize = 10
            pd.Print " "
            pd.CurrentX = 1440 * 1 - pd.TextWidth(partlabs.Grid1.TextMatrix(i, 0))
            pd.Print partlabs.Grid1.TextMatrix(i, 0);
            
            pd.CurrentX = 1440 * 1.25
            pd.Print partlabs.Grid1.TextMatrix(i, 2) & "  ";
            pd.Print StrConv(partlabs.Grid1.TextMatrix(i, 3), vbProperCase);
            
            pd.CurrentX = 1440 * 4.55 - pd.TextWidth(partlabs.Grid1.TextMatrix(i, 4))
            pd.Print partlabs.Grid1.TextMatrix(i, 4);
            
            pd.CurrentX = 1440 * 5.25
            pd.Print "___________________________"
            
        End If
    Next i
    
    Screen.MousePointer = 0
End Sub

Private Sub view_prtall(pd As Control)
    Dim k As Integer, i As Integer, tstat As String
    Dim p As Integer, rstr As String, pflag As Boolean
    Dim halfwidth, sy As Long, s As String
    Screen.MousePointer = 11
    If TypeOf pd Is PictureBox Then
        If Me.paplegal.Checked = True Then
            rstr = localAppDataPath & "\blnk8x14.bmp"
        Else
            rstr = localAppDataPath & "\blnk8x11.bmp"
        End If
        pd.Picture = LoadPicture(rstr)
        rstr = Dir(localAppDataPath & "\cic*.bmp")
        Do While Len(rstr) > 0
            Kill localAppDataPath & "\" & rstr
            rstr = Dir
        Loop
        DoEvents
    End If
    pd.FontName = "Arial"
    pd.FontSize = 8
    pd.CurrentX = 0: pd.CurrentY = 0
    pd.Print Me.Caption & " " & partlabs.Combo1
    pd.FontSize = 48
    'pd.FontBold = True
    pd.FontBold = True
    
    'If TypeOf pd Is Printer Then
    '    If me.paplegal.Checked = True Then
    '        'Printer.PaperBin = vbPRBNUpper
    '        'Printer.PaperBin = 1 'vbPRBNMiddle
    '        Printer.PaperSize = 5
    '    Else
    '        'Printer.PaperBin = vbPRBNLower
    '        'Printer.PaperBin = vbPRBNAuto
    '        Printer.PaperSize = 1
    '    End If
    'End If
    pd.CurrentX = 0: pd.CurrentY = 0
    If Me.paplegal.Checked = True Then
        'pd.FontSize = 100
        'pd.Print " "
        pd.FontSize = 48
        pd.FontUnderline = True
        'halfwidth = pd.TextWidth("(Fold at dotted line)") / 2
        'pd.CurrentX = pd.ScaleWidth / 2 - halfwidth
        'pd.Print "(Fold at dotted line)"
        'pd.FontSize = 100
        'pd.FontBold = False
        'halfwidth = pd.TextWidth("................") / 2
        'pd.CurrentX = pd.ScaleWidth / 2 - halfwidth
        'pd.Print "................"
        pd.CurrentY = 1440 * 1.5
        s = partlabs.Combo2
        
        If pd.TextWidth(s) > pd.ScaleWidth Then
            pd.FontSize = 24
            For i = Len(s) To 5 Step -1
                s = Left(s, i)
                If pd.TextWidth(s) <= pd.ScaleWidth Then Exit For
            Next i
        End If
        
        halfwidth = pd.TextWidth(s) / 2
        pd.CurrentX = pd.ScaleWidth / 2 - halfwidth
        pd.Print s
        
        
        pd.CurrentY = 1440 * 2.5
        'pd.CurrentY = 1440 * 5
    Else
        pd.FontSize = 48
        pd.Print " "
        'pd.CurrentY = 15480 / 11 * 2.5
    End If
    'pd.FontName = "Arial"
    'pd.FontSize = 10 '100
    pd.FontBold = True
    pd.FontUnderline = True
    pd.FontSize = 18
    'pd.CurrentX = 1440 * 0.5: pd.Print "  SKU";
    'pd.CurrentX = 1440: pd.Print "  SKU";
    pd.CurrentX = 1: pd.Print "  SKU  ";
    pd.CurrentX = 1440 * 3.5: pd.Print "Description";
    pd.CurrentX = 1440 * 7.2: pd.Print "Wraps"
    'pd.Print " "
    For i = 1 To partlabs.Grid1.Rows - 1
        If partlabs.Grid1.TextMatrix(i, 0) = labpt.Caption Then
            pd.FontSize = 36
            pd.FontUnderline = False
            pd.FontBold = False
            'pd.CurrentX = 1440 * 0.5
            pd.CurrentX = 1
            pd.Print partlabs.Grid1.TextMatrix(i, 2) & " ";
            pd.Print StrConv(partlabs.Grid1.TextMatrix(i, 3), vbProperCase);
            s = "  " & partlabs.Grid1.TextMatrix(i, 4)
            'pd.CurrentX = pd.ScaleWidth - pd.TextWidth(s)
            pd.CurrentX = (1440 * 8) - pd.TextWidth(s)
            pd.Print s
            'pd.CurrentX = 1440 * 7
            'pd.Print " " & partlabs.Grid1.TextMatrix(i, 4)
            pd.FontSize = 16
            pd.FontBold = False
            'pd.CurrentX = 1440 * 0.5
            pd.CurrentX = 1
            pd.Print "Code Date(s):"
            pd.Print " "
        End If
    Next i
    
    Screen.MousePointer = 0
    Exit Sub
    
    pd.FontSize = 125   ' Set font size.
    halfwidth = pd.TextWidth(skulab.Caption) / 2  ' Calculate one-half width.
    pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
    pd.Print skulab.Caption
    pd.FontSize = 22
    pd.Print " "
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
    pd.FontSize = 22
    pd.Print " "
    pd.FontSize = 48
    halfwidth = pd.TextWidth(pkglab.Caption) / 2  ' Calculate one-half width.
    pd.CurrentX = pd.ScaleWidth / 2 - halfwidth   ' Set X.
    pd.Print pkglab.Caption
    'pd.FontSize = 60
    pd.Print " "
    pd.FontSize = 80
    sy = pd.CurrentY
    pd.Print lotlab.Caption
    pd.CurrentY = sy
    pd.CurrentX = pd.ScaleWidth - pd.TextWidth(seqlab.Caption)
    pd.Print seqlab.Caption
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
End Sub

Private Sub Command1_Click()
    If Me.paplegal.Checked = True Then
        Printer.PaperSize = 5
    Else
        Printer.PaperSize = 1
    End If
    'Call view_prtall(Printer)
    Call view_prtlist(Printer)
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    'refresh_grid
    'DoEvents
    'sort_grid (5)
    Picture1.Width = 12240
    If Me.paplegal.Checked = True Then
        Picture1.Height = 20160
        If Len(Dir(localAppDataPath & "\blnk8x14.bmp")) = 0 Then
            SavePicture Picture1.Image, localAppDataPath & "\blnk8x14.bmp"
        End If
    Else
        Picture1.Height = 15840
        If Len(Dir(localAppDataPath & "\blnk8x11.bmp")) = 0 Then
            SavePicture Picture1.Image, localAppDataPath & "\blnk8x11.bmp"
        End If
    End If
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
        If Me.paplegal.Checked = True Then
            VScroll1.Max = 20160 - VScroll1.Height
        Else
            VScroll1.Max = 15840 - VScroll1.Height
        End If
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
    Me.Caption = partlabs.Combo2 & " Pallet #" & labpt.Caption
    If Me.paplegal.Checked = True Then
        Me.Caption = Me.Caption & " - Legal 8.5x14"
    Else
        Me.Caption = Me.Caption & " - Letter 8.5x11"
    End If
    'Call view_prtall(Picture1)
    Call view_prtlist(Picture1)
    'If prtdevice = "Printer" Then Call view_prtall(Printer)
End Sub

Private Sub VScroll1_Change()
    Picture1.Move Picture1.Left, Frame1.Height - VScroll1.Value
End Sub


