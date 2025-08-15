VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   9585
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9765
   LinkTopic       =   "Form9"
   ScaleHeight     =   9585
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6165
      _Version        =   327680
      Cols            =   4
      ForeColor       =   4210688
      BackColorFixed  =   16777152
      BackColorSel    =   255
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label bannah 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label bcode 
      Caption         =   "..."
      Height          =   255
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub print_grid(gname As Control, r1 As Integer, r2 As Integer, rtitle As String)
    Dim i As Integer, k As Integer, j As Integer
    Dim xs As Long, xe As Long, xm As Long
    Dim ys As Long, ye As Long
    Dim cw(0 To 12) As Long
    For i = 0 To gname.Cols - 1
        cw(i) = gname.ColWidth(i)
    Next i
    'Override Grid Col Widths
    'cw(0) = 600
    'cw(1) = 1700
    'cw(2) = 1700
    'cw(3) = 1700
    'cw(4) = 1700
    'cw(5) = 1700
    'cw(VGrid.Cols - 1) = 1000
    xs = 0: xe = xs
    For i = 0 To gname.Cols - 1
        If cw(i) > 10 Then xe = xe + cw(i)
    Next i
    If xe > 11600 Then
        Printer.Orientation = 2
    Else
        Printer.Orientation = 1
    End If
    
    Printer.FontTransparent = True
    Printer.FillStyle = 0
    Printer.FillColor = QBColor(15)
    Printer.DrawMode = 1
    Printer.ForeColor = QBColor(0)
    
    Printer.FontName = "MS Serif"
    Printer.FontTransparent = True
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    Printer.Print rtitle
    Printer.Print bannah

    Printer.FontSize = 8
    Printer.Line (xs, 1200)-(xe, 1200)
    Printer.Line (xs, 1440)-(xe, 1440)
    Printer.FillColor = QBColor(15)
    Printer.DrawWidth = 3
    j = 0
    For i = r1 To r2 + 1
        ye = j * 240 + 1440
        Printer.Line (xs, ye)-(xe, ye)
        j = j + 1
    Next i
    Printer.DrawWidth = 1
    Printer.FontBold = False
    xm = xs + 100
    For k = 0 To gname.Cols - 1
        If cw(k) > 10 Then
            Printer.PSet (xm, 1230)
            Printer.Print gname.TextMatrix(0, k)
            xm = xm + cw(k)
        End If
    Next k
    j = 1
    For i = r1 To r2
        xm = xs + 100
        For k = 0 To gname.Cols - 1
            If cw(k) > 10 Then
                Printer.PSet (xm, j * 240 + 1230)
                Printer.Print gname.TextMatrix(i, k)
                xm = xm + cw(k)
            End If
        Next k
        j = j + 1
    Next i
    ys = 1200
    xm = xs
    Printer.DrawWidth = 6
    For i = 0 To gname.Cols - 1
        If cw(i) > 10 Then
            Printer.Line (xm, ys)-(xm, ye)
            xm = xm + cw(i)
        End If
    Next i
    Printer.Line (xm, ys)-(xm, ye)
    Printer.Print " "
    Printer.Print "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Printer.EndDoc
End Sub

Private Sub refresh_grid()
    Dim f0 As String, f1 As String, f2 As String
    Dim f3 As String, f4 As String, f5 As String
    Dim f6 As String, f7 As String, f8 As String
    Dim f9 As String, sqlx As String
    Dim cfile As String
    cfile = Form1.webdir & "\counts\gemmeop." & bcode.Caption
    If Len(Dir(cfile)) = 0 Then Exit Sub
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear
    Grid1.Rows = 1: Grid1.Cols = 6
    Open cfile For Input As #1
    Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
    sqlx = "^" & f1 & "|<" & f2 & "|^" & f3 & "|^" & f4
    sqlx = sqlx & "|^" & f5 & "|^" & f9
    Me.Caption = "Oracle End of Period Totals - " & f9
    bannah = "Last Updated:  " & Format(FileDateTime(cfile), "m-dd-yyyy h:mm am/pm")
    Grid1.FormatString = sqlx
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9
        If Val(f0) > 0 Then
            sqlx = f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9)
            sqlx = sqlx & f4 & Chr(9) & f5 & Chr(9) & f9
            Grid1.AddItem sqlx
        End If
    Loop
    Grid1.AddItem ""
    sqlx = f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9)
    sqlx = sqlx & f4 & Chr(9) & f5 & Chr(9) & f9
    Grid1.AddItem sqlx
    
    Close #1
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 1200
    Grid1.Redraw = True
End Sub

Private Sub bcode_Change()
    refresh_grid
End Sub

Private Sub Form_Load()
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
    'refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    bannah.Width = Me.Width - 80
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 980
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.gemmeop.Checked = False
End Sub

Private Sub prtmenu_Click()
    Dim i As Integer, k As Integer, j As Integer
    j = 0: k = 1
    Screen.MousePointer = 11
    For i = 1 To Grid1.Rows - 1
        If j > 53 Then
            Call print_grid(Grid1, k, i, Me.Caption)
            k = i + 1
            j = 0
        End If
        j = j + 1
    Next i
    If k < Grid1.Rows - 1 Then
        Call print_grid(Grid1, k, Grid1.Rows - 1, Me.Caption)
    End If
    Screen.MousePointer = 0
End Sub
