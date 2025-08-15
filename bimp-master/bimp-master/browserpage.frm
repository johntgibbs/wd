VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form browserpage 
   Caption         =   "W/D Browser Home Page"
   ClientHeight    =   13260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   ScaleHeight     =   13260
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9735
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   3495
      ExtentX         =   6165
      ExtentY         =   17171
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
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
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   10800
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4260
      _Version        =   327680
      WordWrap        =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.Label refkey 
      Caption         =   "refkey"
      Height          =   255
      Left            =   9240
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label webdir 
      Caption         =   "\\BBC-03-FILESVR\SharedGroups\wd\html"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
   End
End
Attribute VB_Name = "browserpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rebuild_homegrid()
    Dim ds As ADODB.Recordset, s As String
    Dim odates As String, scdates As String
    Dim rt As String, rh As String, rf As String
    'On Error GoTo vberror
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 2: pgrid.FixedCols = 1
    pgrid.FixedCols = 0
    'Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.shipdb
    Set ds = wdb.Execute("select * from wdstatus")
    If ds.BOF = False Then
        ds.MoveFirst
        odates = ds!orddates
        scdates = ds!schdates
    End If
    ds.Close
    s = "select * from branches where branch > 90 and brnmess > '   ' order by branch desc"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        's = "<img src=" & Chr(34) & "images\lanette.jpg" & Chr(34) & "><BR>" & Chr(9)
        's = "<img src=" & Chr(34) & "images\new.jpg" & Chr(34) & "><BR>" & Chr(9)
        s = "<img src=" & Chr(34) & "images\plant500.jpg" & Chr(34) & "><BR><BR>"
        s = s & "<img src=" & Chr(34) & "images\plant501.jpg" & Chr(34) & "><BR><BR>"
        s = s & "<img src=" & Chr(34) & "images\plant502.jpg" & Chr(34) & "><BR><BR>" & Chr(9)
        's = "Branch Notes<BR>" & Chr(9)
        's = s & "<b>Notes</b><BR>" & Chr(9)
        Do Until ds.EOF
            s = s & "<b>" & ds!branchname & ":</b><br>" & ds!brnmess & "<hr>"
            'pgrid.AddItem s & ds!branchname & Chr(9) & ds!brnmess
            's = ""
            ds.MoveNext
        Loop
        pgrid.AddItem s
    End If
    ds.Close
    's = Chr(9)
    s = "<a href=" & Chr(34) & "stock.htm" & Chr(34) & ">"
    's = s & "<img src=" & Chr(34) & "images\bbstock.jpg" & Chr(34) & " Border=0><BR>Out of Stock Listings</a>"
    's = s & "<BR><b>Out of Stock Listings</b></a>"
    s = s & "<b>Out of Stock Listings</b></a>"
    s = s & Chr(9) & "<b>Last Updated: " & FileDateTime(Me.webdir & "\stock.htm") & "</b>"
    pgrid.AddItem s
    
    's = Chr(9)
    's = "<img src=" & Chr(34) & "images\men in plant.tif" & Chr(34) & ">" & Chr(9)
    s = "<a href=" & Chr(34) & "discont.htm" & Chr(34) & "><b>Discontinued Products</b></a>"
    s = s & Chr(9) & "<b>Last Updated: " & FileDateTime(Me.webdir & "\discont.htm") & "</b>"
    pgrid.AddItem s
    
    's = Chr(9)
    s = "<a href=" & Chr(34) & "stock\wdstk.htm" & Chr(34) & ">"
    's = s & "<img src=" & Chr(34) & "images\wdstk.jpg" & Chr(34) & " Border=0><BR>Blue Bell Pallet Stacking Patterns</a>"
    's = s & "<BR><b>Pallet Stacking Patterns</b></a>"
    s = s & "<b>Pallet Stacking Patterns</b></a>"
    s = s & Chr(9) & "<b>Last Updated: " & FileDateTime(Me.webdir & "\stock\wdstk.htm") & "</b>"
    pgrid.AddItem s
    
    's = Chr(9)
    's = "<img src=" & Chr(34) & "images\realtrail.jpg" & Chr(34) & ">" & Chr(9)
    s = "<b>Branch Orders</b>" & Chr(9)
    If Len(Dir(Me.webdir & "\orderoff.txt")) > 0 Then
        's = s & "<img src=" & Chr(34) & "images\orderoff.jpg" & Chr(34) & ">"
        's = s & "<BR><b>Not accepting branch orders at this time...</b>"
        s = s & "<b>Not accepting branch orders at this time...</b>"
    Else
        's = s & "<img src=" & Chr(34) & "images\realtrail.jpg" & Chr(34) & ">"
        's = s & "<BR><b>Currently accepting orders for: " & odates & ".</b>"
        s = s & "<b>Currently accepting orders for: " & odates & ".</b>"
    End If
    pgrid.AddItem s
    
    
    's = "<img src=" & Chr(34) & "images\toytruck.gif" & Chr(34) & ">" & Chr(9)
    's = Chr(9)
    's = "Transport Schedule Requests" & Chr(9)
    s = "<b>Transport Requests</b" & Chr(9)
    s = s & "<b>Currently accepting requests for Week of " & scdates & ".</b>"
    pgrid.AddItem s
    
    s = "<a href=" & Chr(34) & "schedule\trnspts.htm" & Chr(34) & ">"
    s = s & "<img src=" & Chr(34) & "images\realtruck.jpg" & Chr(34) & " Border=0><BR><b>Transport Schedules</b></a>"
    s = s & Chr(9) & "<b>Last Updated: " & FileDateTime(Me.webdir & "\schedule\trnspts.htm") & "</b>"
    pgrid.AddItem s
    
    's = Chr(9)
    s = "<a href=" & Chr(34) & "directs\wdirects.htm" & Chr(34) & "><b>Driving Directions</b></a>"
    s = s & Chr(9) & "<b>Last Updated: " & FileDateTime(Me.webdir & "\directs\wdirects.htm") & "</b>"
    pgrid.AddItem s
    
    's = Chr(9)
    s = "<a href=" & Chr(34) & "schedule\trltrks.htm" & Chr(34) & "><b>Trailer Tracking</b></a>"
    s = s & Chr(9) & "<b>Last Updated: " & FileDateTime(Me.webdir & "\schedule\trltrks.htm") & "</b>"
    pgrid.AddItem s
    
    's = Chr(9)
    's = "<a href=" & Chr(34) & "schedule\intrax.htm" & Chr(34) & ">"
    's = s & "<img src=" & Chr(34) & "images\tractors.jpg" & Chr(34) & " Border=0><BR>Tractors in the Yard</a>"
    's = s & Chr(9) & "Last Updated: " & FileDateTime(Me.webdir & "\schedule\intrax.htm")
    'pgrid.AddItem s
    
    s = "<a href=" & Chr(34) & "stock\brancaps.htm" & Chr(34) & "><b>Branch Storage Capacities</b></a>"
    's = "<a href=" & Chr(34) & "u:\wdapps\eforklift.exe" & Chr(34) & "><b>Branch Storage Capacities</b></a>"
    s = s & Chr(9) & "<b>Last Updated: " & FileDateTime(Me.webdir & "\stock\brancaps.htm") & "</b>"
    pgrid.AddItem s
    
    rt = "<center><b>Blue Bell Warehousing & Distribution</b>"
    'rt = "<img src=" & Chr(34) & "images/wdlogo.gif" & Chr(34) & ">"
    rt = rt & "<body background=" & Chr(34) & "images\wdbkgd.gif" & Chr(34) & ">"
    'rt = rt & "<body background=" & Chr(34) & "images\wdfall4.gif" & Chr(34) & ">"
    rh = "<img src=" & Chr(34) & "images/bbcolor.jpg" & Chr(34) & ">"
    'rh = rh & "<br><img src=" & Chr(34) & "images/wdlogo.gif" & Chr(34) & ">"
    'rf = "<center><b>Updated: " & Format(Now, "m-d-yyyy h:mm am/pm") '& "</b"
    rf = "<center><b>Updated: " & bimp_status_time '& "</b"
    pgrid.FormatString = "^|^|^"
    pgrid.ColWidth(0) = 2000
    pgrid.ColWidth(1) = pgrid.Width - 2000
    's = "s:\wd\html\bbwdtest.htm"
    'Call htmlcolorgrid(Me, Form1.webdir & "\bbwd.htm", pgrid, rt, rh, rf, "Linen", "White", "lemonchiffon")
    'Call htmlcolorgrid(Me, s, pgrid, rt, rh, rf, "Linen", "White", "lemonchiffon")
    'Call htmlcolorgrid(Me, s, pgrid, rt, rh, rf, "lemonchiffon", "White", "lemonchiffon")
    Call htmlcolorgrid(Me, Me.webdir & "\bbwd.htm", pgrid, rt, rh, rf, "cornsilk", "White", "lemonchiffon")
    WebBrowser1.Navigate Me.webdir & "\bbwd.htm"
End Sub

Private Sub Command1_Click()
    rebuild_homegrid
End Sub

Private Sub Form_Load()
    WebBrowser1.Navigate Me.webdir & "\bbwd.htm"
End Sub

Private Sub Form_Resize()
    pgrid.Width = Me.Width - 180
    WebBrowser1.Width = Me.Width - 180
End Sub

Private Sub refkey_Change()
    rebuild_homegrid
End Sub
