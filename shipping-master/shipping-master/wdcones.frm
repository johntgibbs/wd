VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form wdcones 
   Caption         =   "W/D Cone Orders"
   ClientHeight    =   9120
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11775
   LinkTopic       =   "Form7"
   ScaleHeight     =   9120
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "wdcones.frx":0000
      Top             =   0
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Link to Sales Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   7
      Top             =   0
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   9000
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   0
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1931
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4260
      _Version        =   327680
      BackColor       =   16777215
      BackColorFixed  =   65535
      BackColorSel    =   192
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label plantcode 
      Caption         =   "Label2"
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Shipped From:"
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
      TabIndex        =   5
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label brcode 
      Caption         =   "Label1"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Menu prtord 
      Caption         =   "&Print"
   End
End
Attribute VB_Name = "wdcones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcol As Boolean, pflag As Boolean
Private Sub refresh_grid()
    Dim rid As String, rdesc As String, ruom As String
    Dim sqlx As String, ritem As String
    Dim rbr As String, rdate As String, rqty As String
    Dim rconv As String, rbulk As String, rsrc As String
    Dim rbill As String, i As Integer
    plantcode = List1
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Visible = False: Grid1.Clear
    Grid1.Rows = 1: Grid1.Cols = 7: Grid1.FixedCols = 3
    If Len(Dir(Form1.webdir & "\stock\dslist.txt")) > 0 Then
        Open Form1.webdir & "\stock\dslist.txt" For Input As #1
        Do Until EOF(1)
            Input #1, rid, ritem, rdesc, ruom, rconv, rbulk, rsrc
            sqlx = rid & Chr(9)
            sqlx = sqlx & ritem & Chr(9)
            sqlx = sqlx & rdesc & Chr(9)
            sqlx = sqlx & Chr(9)
            sqlx = sqlx & rbulk & Chr(9)
            sqlx = sqlx & rconv & Chr(9)
            sqlx = sqlx & ruom
            If rsrc = List1 Or (rsrc = "000" And Combo1.ListIndex = 0) Then
                Grid1.AddItem sqlx
            End If
        Loop
        Close #1
    End If
    'If Len(Dir(Form1.webdir & "\dry\orders\dsord" & List1 & "." & brcode)) > 0 Then
    '    Open Form1.webdir & "\dry\orders\dsord" & List1 & "." & brcode For Input As #1
    If Len(Dir(Form1.webdir & "\dry\orders\wdord" & List1 & "." & brcode)) > 0 Then
        Open Form1.webdir & "\dry\orders\wdord" & List1 & "." & brcode For Input As #1
        Do Until EOF(1)
            Input #1, rbr, rid, ritem, rdesc, rdate, rqty, ruom, eqty, euom
            If rbr = brcode Then
                For i = 0 To Grid1.Rows - 1
                    If ritem = Grid1.TextMatrix(i, 1) And rdesc = Grid1.TextMatrix(i, 2) Then
                        Grid1.TextMatrix(i, 3) = Val(rqty)
                    End If
                Next i
            End If
        Loop
        Close #1
    End If
    Grid1.FormatString = "|^Item|^Description|^Qty|^UOM"
    Grid1.ColWidth(0) = 1 '600
    Grid1.ColWidth(1) = 2200
    Grid1.ColWidth(2) = 4000
    Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 1200
    Grid1.ColWidth(5) = 1 '800
    Grid1.ColWidth(6) = 1 '800
    Grid1.RowHeight(0) = -1
    Grid1.RowHeight(-1) = Grid1.RowHeight(0) * 2
    Grid1.Visible = True
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub brcode_Change()
    Dim s As String, ds As ADODB.Recordset                                                      'jv030216
    Combo1.Clear: List1.Clear
    s = "select * from valuelists where listreturn = '" & Format(Val(Me.brcode), "000") & "'"   'jv030216
    Set ds = Wdb.Execute(s)                                                                     'jv030216
    If ds.BOF = False Then                                                                      'jv030216
        ds.MoveFirst                                                                            'jv030216
        If ds!listdisplay = "K10" Then                                                          'jv030216
            Combo1.AddItem "Broken Arrow": List1.AddItem "501"                                  'jv030216
        End If                                                                                  'jv030216
        If ds!listdisplay = "A10" Then                                                          'jv030216
            Combo1.AddItem "Sylacauga": List1.AddItem "502"                                     'jv030216
        End If                                                                                  'jv030216
    End If                                                                                      'jv030216
    ds.Close                                                                                    'jv030216
    'If Form1.rco51.Enabled Or Form7.brcode = "47" Then
    '    Combo1.AddItem "Broken Arrow": List1.AddItem "501"
    'End If
    'If Form1.rco52.Enabled Or Form7.brcode = "52" Then
    '    Combo1.AddItem "Sylacauga": List1.AddItem "502"
    'End If
    Combo1.AddItem "Brenham": List1.AddItem "500"
    Combo1.ListIndex = 0
    ''Call refresh_grid
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
    Dim i, app As String
    app = "http://supply.bluebell.com"
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe " & app, vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & app, vbNormalFocus)
        Exit Sub
    End If

End Sub

Private Sub Form_Load()
    Me.Left = Form1.Left
    'Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    'Me.Height = Form1.WebBrowser1.Height
    If Form1.plantno = "50" Then
        Me.brcode = "01"
        Me.plantcode = "500"
    End If
    If Form1.plantno = "51" Then
        Me.brcode = "47"
        Me.plantcode = "501"
    End If
    If Form1.plantno = "52" Then
        Me.brcode = "52"
        Me.plantcode = "502"
    End If
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    Grid2.Width = Me.Width - 80
    
    If Me.Height > 2000 Then
        Grid1.Height = Me.Height - (Combo1.Height + 950)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim unam As String, sqlx As String
    'unam = Left(Form1.wduser, InStr(1, Form1.wduser, " "))
    unam = Left(wduserid, InStr(1, wduserid, " "))
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
    If Grid1.Col = 3 Then
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
        End If
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            Grid1.Text = Grid1.Text & Chr(KeyAscii)
        End If
    End If
End Sub

Private Sub Grid1_RowColChange()
    edcol = True
End Sub

Private Sub List1_Click()
    Dim unam As String, sqlx As String
    'unam = Left(Form1.wduser, InStr(1, Form1.wduser, " "))
    unam = Left(wduserid, InStr(1, wduserid, " "))
    unam = "Hey " & unam & "....."
    If pflag = True Then
        sqlx = "Changes made to the order have not been posted."
        sqlx = sqlx & "  Do you wish to post the changes now?"
        If MsgBox(sqlx, vbQuestion + vbYesNo, unam) = vbYes Then
            Call prtord_Click
        End If
    End If
    refresh_grid
End Sub

Private Sub prtord_Click()
    Dim i As Integer, sqlx As String, k As Integer
    Dim rt As String, rf As String, rh As String
    Screen.MousePointer = 11
    Grid2.Rows = 1
    Grid2.Cols = Grid1.Cols
    Grid2.FormatString = Grid1.FormatString
    For i = 0 To Grid1.Cols - 1
        Grid2.ColWidth(i) = Grid1.ColWidth(i)
    Next i
    Grid2.Rows = 1
    
    'Open Form1.webdir & "\dry\orders\dsord" & plantcode & "." & Me.brcode For Output As #1
    Open Form1.webdir & "\dry\orders\wdord" & plantcode & "." & Me.brcode For Output As #1
    For i = 0 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 3)) > 0 Then
            Write #1, brcode;
            Write #1, Grid1.TextMatrix(i, 0); 'Item Id
            Write #1, Grid1.TextMatrix(i, 1); 'Item No
            Write #1, Grid1.TextMatrix(i, 2); 'Item Desc
            Write #1, Format(Now, "m-d-yyyy"); 'Date
            Write #1, Grid1.TextMatrix(i, 3); 'Bulk Order qty
            Write #1, Grid1.TextMatrix(i, 4); 'Bulk Uom
            Write #1, Format(Val(Grid1.TextMatrix(i, 3)) * Val(Grid1.TextMatrix(i, 5)), "0"); 'Oracle Qty
            Write #1, Grid1.TextMatrix(i, 6)  'Oracle UOM
            sqlx = ""
            For k = 0 To Grid1.Cols - 2
                sqlx = sqlx & Grid1.TextMatrix(i, k) & Chr(9)
            Next k
            sqlx = sqlx & Grid1.TextMatrix(i, Grid1.Cols - 1)
            Grid2.AddItem sqlx
        End If
    Next i
    Close #1
    pflag = False
    rt = Me.Caption & " From "
    
    If plantcode = "500" Then rt = rt & "Brenham"
    If plantcode = "501" Then rt = rt & "Broken Arrow"
    If plantcode = "502" Then rt = rt & "Sylacauga"
    rh = Format(Now, "m-d-yyyy")
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Call printflexgrid(Printer, Grid2, rt, rh, rf)
    Screen.MousePointer = 0
End Sub

