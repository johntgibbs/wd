VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form12 
   Caption         =   "New Product Release Schedule"
   ClientHeight    =   8055
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8955
   LinkTopic       =   "Form12"
   ScaleHeight     =   8055
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox gw 
      Height          =   285
      Left            =   6120
      TabIndex        =   3
      Text            =   "..."
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox rfile 
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Text            =   "c:\release.txt"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   5535
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9763
      _Version        =   327680
      BackColor       =   12648447
      BackColorFixed  =   65535
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3836
      _Version        =   327680
      BackColor       =   12648384
      BackColorFixed  =   65280
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label postdate 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   8655
   End
   Begin VB.Menu prtmenu 
      Caption         =   "&Print"
      Begin VB.Menu pgrid1 
         Caption         =   "Product Release Schedule"
      End
      Begin VB.Menu pgrid2 
         Caption         =   "Branch Allocation"
      End
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refresh_grid1()
    Dim s As String, f0 As String, f1 As String
    Dim f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    Open rfile For Input As #1
    Do Until EOF(1)
        Input #1, f0
        If f0 = "R" Then
            Input #1, f1, f2, f3, f4, f5, f6, f7
            s = f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9) & f4 & Chr(9)
            s = s & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9)
            If f2 = "500" Then s = s & "Brenham"
            If f2 = "501" Then s = s & "Broken Arrow"
            If f2 = "502" Then s = s & "Sylacauga"
            Grid1.AddItem s
        End If
        If f0 = "S" Then
            Input #1, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10
        End If
        If f0 = "P" Then
            Input #1, f1, f2, f3
        End If
    Loop
    Close #1
    Grid1.FormatString = "^ID|^Orgn|^SKU|<Product|<1st Run Date|<Pool Qty|<Pattern Used|<From Plant"
    Grid1.ColWidth(0) = 4 '00
    Grid1.ColWidth(1) = 6 '00
    Grid1.ColWidth(2) = 600
    Grid1.ColWidth(3) = 2500
    Grid1.ColWidth(4) = 1500
    Grid1.ColWidth(5) = 1 '000
    Grid1.ColWidth(6) = 2 '500
    Grid1.ColWidth(7) = 1400
End Sub
Sub refresh_grid2()
    Dim s As String, f0 As String, f1 As String
    Dim f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String
    Dim r As String, b As String, pname As String
    postdate = "Posted: " & Format(FileDateTime(rfile), "mmmm d, yyyy  h:mm am/pm")
    r = Grid1.TextMatrix(Grid1.Row, 0)
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 2
    Open rfile For Input As #1
    Do Until EOF(1)
        Input #1, f0
        If f0 = "R" Then
            Input #1, f1, f2, f3, f4, f5, f6, f7
            If f1 = r Then
                s = "New Product Release:  " & f3 & " " & f4
                Grid2.AddItem Chr(9) & s
                Grid2.AddItem " "
                If f2 = "500" Then pname = "Brenham"
                If f2 = "501" Then pname = "Broken Arrow"
                If f2 = "502" Then pname = "Sylacauga"
                s = "1st Product Run " & f5 & ".  To date, " & pname & " has produced: " & Format(Val(f6), "0") & " units."
                Grid2.AddItem Chr(9) & s
                Grid2.AddItem " "
            End If
        End If
        If f0 = "S" Then
            Input #1, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10
            If f1 = r And f2 = gw Then
                b = f3
                s = b & " is alloted: " & Format(Val(f8), "0") & " units."
                Grid2.AddItem Chr(9) & s
                s = b & " currently has " & Format(Val(f5), "0") & " units on hand."
                Grid2.AddItem Chr(9) & s
                s = b & " currently has " & Format(Val(f6), "0") & " units on order."
                Grid2.AddItem Chr(9) & s
                s = b & " has sold " & Format(Val(f7), "0") & " units in the last 30-day period."
                Grid2.AddItem Chr(9) & s
                If Val(f9) >= 0 Then
                    s = b & "'s allotment stands at " & Format(Val(f9), "0")
                    s = s & " units (" & Format(Val(f10), "0") & " pallets)."
                Else
                    s = b & " has exceeded its allotment by " & Format(Val(f9) * -1, "0")
                    s = s & " units (" & Format(Val(f10) * -1, "0") & " pallets)."
                End If
                Grid2.AddItem Chr(9) & s
                Grid2.AddItem " "
                Grid2.AddItem Chr(9) & "Future Production Runs"
            End If
        End If
        If f0 = "P" Then
            Input #1, f1, f2, f3, f4, f5
            If f1 = r And f2 = gw Then
                If Val(f5) >= 0 Then
                    s = Format(f3, "mmm dd, yyyy") & " " & b & " will be alloted " & f4 & " units from this run.  Net total:  " & f5
                    Grid2.AddItem Chr(9) & s
                Else
                    s = Format(f3, "mmm dd, yyyy") & " " & b & " has exceeded its allotment by " & Format(Val(f5) * -1, "0") & " units."
                    Grid2.AddItem Chr(9) & s
                End If
            End If
        End If
    Loop
    Close #1
    Grid2.FormatString = "^|<"
    Grid2.ColWidth(0) = 4 '00
    Grid2.ColWidth(1) = 8600
End Sub
Private Sub Form_Load()
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
    rfile = Form1.webdir & "\stock\release.txt"
    refresh_grid1
    refresh_grid2
End Sub

Private Sub Form_Resize()
    Grid2.Width = Me.Width - 80
    Grid1.Width = Me.Width - 80
End Sub

Private Sub Grid1_Click()
    refresh_grid2
End Sub

Private Sub Grid1_RowColChange()
    refresh_grid2
End Sub

Private Sub gw_Change()
    refresh_grid2
End Sub

Private Sub pgrid1_Click()
    Dim rt As String, rh As String, rf As String
    rt = Me.Caption
    rh = Format(Now, "mmmm d, yyyy")
    rf = postdate.Caption
    Call printflexgrid(Printer, Grid1, rt, rh, rf)
End Sub

Private Sub pgrid2_Click()
    Dim rt As String, rh As String, rf As String
    rt = "New Release Product Allotment"
    rh = Grid1.TextMatrix(Grid1.Row, 2) & " " & Grid1.TextMatrix(Grid1.Row, 3)
    rf = postdate.Caption
    Call printflexgrid(Printer, Grid2, rt, rh, rf)
End Sub


