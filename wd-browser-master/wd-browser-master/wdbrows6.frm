VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   4920
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7770
   LinkTopic       =   "Form6"
   ScaleHeight     =   4920
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2990
      _Version        =   327680
      Cols            =   5
      BackColor       =   16777215
      BackColorFixed  =   12632256
      BackColorSel    =   8421376
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
   End
   Begin VB.Label brcode 
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Menu prtord 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcol As Boolean, pflag As Boolean
Private Sub refresh_grid()
    Dim rid As String, rdesc As String, rsize As String
    Dim ruom As String, sqlx As String
    Dim rbr As String, rdate As String, rqty As String
    Dim rbill As String, i As Integer
    Screen.MousePointer = 11
    Grid1.Visible = False: Grid1.Clear
    Grid1.Rows = 1: Grid1.Cols = 6: Grid1.FixedCols = 4
    If Len(Dir(Form1.webdir & "\stock\vwlist.txt")) > 0 Then
        Open Form1.webdir & "\stock\vwlist.txt" For Input As #1
        Do Until EOF(1)
            Input #1, rid, rdesc, rsize, ruom
            sqlx = rid & Chr(9) & rdesc & Chr(9)
            sqlx = sqlx & rsize & Chr(9)
            sqlx = sqlx & ruom
            Grid1.AddItem sqlx
        Loop
        Close #1
    End If
    If Len(Dir(Form1.webdir & "\orders\vcorder." & brcode)) > 0 Then
        Open Form1.webdir & "\orders\vcorder." & brcode For Input As #1
        Do Until EOF(1)
            Input #1, rbr, rid, rdate, rqty, rbill
            If rbr = brcode Then
                For i = 0 To Grid1.Rows - 1
                    If rid = Grid1.TextMatrix(i, 0) Then
                        Grid1.TextMatrix(i, 4) = Val(rqty)
                        Grid1.TextMatrix(i, 5) = rbill
                    End If
                Next i
            End If
        Loop
        Close #1
    End If
    Grid1.FormatString = "^ID|^Description|^Size|^Uom|^Qty|^Bill To"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 2500
    Grid1.ColWidth(2) = 600
    Grid1.ColWidth(3) = 600
    Grid1.ColWidth(4) = 600
    Grid1.ColWidth(5) = 2500
    Grid1.RowHeight(0) = -1
    Grid1.RowHeight(-1) = Grid1.RowHeight(0) * 2
    Grid1.Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub brcode_Change()
    Call refresh_grid
End Sub

Private Sub Form_Resize()
    Grid1.Width = Form6.Width - 80
    If Form6.Height > 2000 Then
        Grid1.Height = Form6.Height - 750
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim unam As String, sqlx As String
    unam = Left(Form1.wduser, InStr(1, Form1.wduser, " "))
    unam = "Hey " & unam & "....."
    If pflag = True Then
        sqlx = "Changes made to the order have not been posted."
        sqlx = sqlx & "  Do you wish to post the changes now?"
        If MsgBox(sqlx, vbQuestion + vbYesNo, unam) = vbYes Then
            Call prtord_Click
        End If
    End If
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If Grid1.Col = 4 Or Grid1.Col = 5 Then
        pflag = True
        If edcol = True Then
            Grid1.Text = ""
            edcol = False
        End If
        If KeyAscii = 8 Then
            If Len(Grid1.Text) > 1 Then
                Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
            Else
                Grid1.Text = ""
            End If
            Exit Sub
        End If
        If Grid1.Col = 4 Then
            If KeyAscii >= 48 And KeyAscii <= 57 Then
                Grid1.Text = Grid1.Text & Chr(KeyAscii)
            End If
        Else
            Grid1.Text = Grid1.Text & Chr(KeyAscii)
        End If
    End If
End Sub

Private Sub Grid1_RowColChange()
    edcol = True
End Sub

Private Sub prtord_Click()
    Dim i As Integer, sqlx As String
    Screen.MousePointer = 11
    Open Form1.webdir & "\orders\vcorder." & Form6.brcode For Output As #1
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 4)) > 0 Then
            Write #1, brcode, Grid1.TextMatrix(i, 0), Format(Now, "m-d-yyyy"), Grid1.TextMatrix(i, 4), Grid1.TextMatrix(i, 5)
        End If
    Next i
    Close #1
    pflag = False
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.Print Form6.Caption & "   Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Printer.Print " "
    Printer.Print " "
    Printer.Print " Item                                                Qty      Bill To"
    Printer.Print " "
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 4)) > 0 Then
            sqlx = Space(5 - Len(Grid1.TextMatrix(i, 0)))
            sqlx = sqlx & Grid1.TextMatrix(i, 0) & " "
            sqlx = sqlx & Trim(Grid1.TextMatrix(i, 1)) & "-"
            sqlx = sqlx & Trim(Grid1.TextMatrix(i, 2))
            If Len(sqlx) > 50 Then
                sqlx = Left(sqlx, 50)
            Else
                sqlx = sqlx & Space(50 - Len(sqlx))
            End If
            sqlx = sqlx & Space(5 - Len(Grid1.TextMatrix(i, 4))) & Grid1.TextMatrix(i, 4) & " "
            sqlx = sqlx & Grid1.TextMatrix(i, 3) & Space(6 - Len(Grid1.TextMatrix(i, 3)))
            sqlx = sqlx & Grid1.TextMatrix(i, 5)
            Printer.Print sqlx
        End If
    Next i
    Printer.EndDoc
    Screen.MousePointer = 0
End Sub
