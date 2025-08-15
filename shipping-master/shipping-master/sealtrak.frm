VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form sealtrak 
   Caption         =   "Seal Tracking"
   ClientHeight    =   12180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13215
   LinkTopic       =   "Form2"
   ScaleHeight     =   12180
   ScaleWidth      =   13215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Text            =   "sealtrak.frx":0000
      Top             =   9720
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   9240
      TabIndex        =   17
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   1440
      TabIndex        =   16
      Text            =   "Combo2"
      Top             =   1200
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   7800
      TabIndex        =   15
      Top             =   7200
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
      Left            =   1440
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   120
      Width           =   1455
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
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
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
      Left            =   8760
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
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
      Left            =   6840
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   3255
      Left            =   0
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5741
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4575
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8070
      _Version        =   327680
      FixedCols       =   0
      BackColorFixed  =   16777152
      FocusRect       =   0
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "Thru:"
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
      Left            =   3000
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Thru:"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Destination:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Plant:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Dates:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Seal #s:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "sealtrak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_locations()
    Dim sb As ADODB.Connection, ss As ADODB.Recordset, s As String
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 3
    Combo1.Clear: List1.Clear
    Combo2.Clear: List2.Clear
    Combo1.AddItem "ALL": List1.AddItem "ALL"
    Combo2.AddItem "ALL": List2.AddItem "ALL"
    Set sb = CreateObject("ADODB.Connection")
    sb.Open Form1.schdb
    s = "select lcode, location, loctype from locations order by location"
    Set ss = sb.Execute(s)
    If ss.BOF = False Then
        ss.MoveFirst
        Do Until ss.EOF
            s = ss(0) & Chr(9) & ss(1) & Chr(9) & ss(2)
            Grid2.AddItem s
            Combo2.AddItem ss(1)
            List2.AddItem ss(0)
            If LCase(ss(2)) = "plant" Then
                Combo1.AddItem ss(1)
                List1.AddItem ss(0)
            End If
            ss.MoveNext
        Loop
    End If
    ss.Close: sb.Close
    Grid2.FormatString = "^Code|<Location|^Type"
    Grid2.ColWidth(0) = 1200
    Grid2.ColWidth(1) = 3500
    Grid2.ColWidth(2) = 1200
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
End Sub

Function lname(lc As String) As String
    Dim i As Integer, s As String
    s = "Undefined Location"
    For i = 1 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 0) = lc Then
            s = Grid2.TextMatrix(i, 1)
            Exit For
        End If
    Next i
    lname = s
End Function

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Combo2_Click()
    List2.ListIndex = Combo2.ListIndex
End Sub

Private Sub Command1_Click()
    Dim sb As ADODB.Connection, ss As ADODB.Recordset, s As String
    Dim i As Integer, hc As Boolean
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    Set sb = CreateObject("ADODB.Connection")
    sb.Open Form1.schdb
    
    s = "Select * from truckwo Where sealnum > 0"
    If Val(Text1) > 0 Then s = s & " and sealnum >= " & Val(Text1)
    If Val(Text2) > 0 Then s = s & " and sealnum <= " & Val(Text2)
    If IsDate(Text3) Then s = s & " and wodate >= '" & Text3 & "'"
    If IsDate(Text4) Then s = s & " and wodate <= '" & Text4 & "'"
    If List1 <> "ALL" Then s = s & " and origin = '" & List1 & "'"
    If List2 <> "ALL" Then s = s & " and destination = '" & List2 & "'"
    s = s & " order by sealnum"
    Text5 = s
    Set ss = sb.Execute(s)
    If ss.BOF = False Then
        ss.MoveFirst
        Do Until ss.EOF
            s = ss!sealnum & Chr(9)
            s = s & ss!wonum & Chr(9)
            s = s & ss!r12ticket & Chr(9)
            s = s & Format(ss!wodate, "MM-dd-yyyy") & Chr(9)
            s = s & ss!origin & "-" & lname(ss!origin) & Chr(9)
            s = s & ss!Destination & "-" & lname(ss!Destination) & Chr(9)
            s = s & "#" & ss!trlno & Chr(9)
            s = s & ss!eqnum
            Grid1.AddItem s
            ss.MoveNext
        Loop
    End If
    ss.Close: sb.Close
    If Grid1.Rows > 1 Then
        hc = True
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            hc = Not hc
            If hc Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = Grid1.BackColorFixed
            End If
        Next i
        Grid1.Row = 1
    End If
    s = "^Seal #|^WONum|^Ticket|^Date|<Origin|<Destination|^#|^Trailer"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 2500
    Grid1.ColWidth(5) = 2500
    Grid1.ColWidth(6) = 800
    Grid1.ColWidth(7) = 1200
End Sub

Private Sub Command2_Click()
    Dim rt As String, rh As String, rf As String, hf As String
    rt = Me.Caption
    rh = "Ship Dates: " & Text3 & " thru " & Text4
    rf = "Printed: " & Format(Now, "M-d-yyyy h:mm Am/Pm")
    hf = Form1.tempdir & "\seals.htm"
    htdc(0) = "lightcyan": gndc(0) = Grid1.BackColorFixed
    Call htmlcolorgrid(Me, hf, Grid1, rt, rh, rf, "linen", "lightcyan", "white")
    If Grid1.Rows > 1 Then Grid1.Row = 1
    Grid1.Col = 0
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe " & hf, vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & hf, vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Text1 = "ALL": Text2 = " "
    Text3 = Format(DateAdd("d", -7, Now), "MM-dd-yyyy")
    Text4 = Format(Now, "MM-dd-yyyy")
    Call refresh_locations
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 120
    Grid2.Width = Me.Width - 120
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 2000
End Sub
