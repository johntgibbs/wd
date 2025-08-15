VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form trlstatus 
   Caption         =   "Trailer Status"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12270
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ticket BarCodes"
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
      Left            =   8160
      TabIndex        =   10
      Top             =   120
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
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   10080
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
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
      Left            =   1080
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   240
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   9855
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   17383
      _Version        =   327680
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
   End
   Begin VB.Label ncolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8880
      TabIndex        =   9
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delivered"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8880
      TabIndex        =   8
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "In Transit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "In Process"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8880
      TabIndex        =   6
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Branch:"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "trlstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid()
    Dim ds As ADODB.Recordset, ts As ADODB.Recordset, s As String, tstat As String, q As String
    Dim pno As String, pt10 As Integer, pk10 As Integer, pa10 As Integer, odate As String
    Dim gs As ADODB.Recordset
    'If r12access = False Then
    '    connect_r12
    '    DoEvents
    'End If
    'If r12access = False Then Exit Sub
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    s = "select * from runs where trldate >= '" & Format(Now, "M-d-yyyy") & "'"
    s = s & " and destination = '" & Val(Combo1) & "'"                      'jv121115
    s = s & " and trlno not in ('ZO', 'OP', 'QC')"                          'jv121115
    s = s & " order by trldate, trlno, startime"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = Format(ds!trldate, "dddd M-d-yyyy") & " "
            s = s & branchrec(Val(ds!destination)).branchname & " " & ds!trlno & vbCrLf
            s = s & " From: " & plantrec(Val(ds!loaded)).plantname & " "
            tstat = "In Process"
            If ds!runstat > " " Then tstat = ds!runstat
            'If ticket_post(ds!id) = True Then tstat = "In Transit"
            'If ticket_receipt(ds!id) = True Then tstat = "Received"
            s = s & tstat
            'Grid1.AddItem tstat & Chr(9) & s & Chr(9) & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
            
            q = "select sku, sum(pallets), sum(wraps), sum(units) from trailers"
            q = q & " where runid = " & ds!id & " group by sku"
            Set ts = wdb.Execute(q)
            If ts.BOF = False Then
                ts.MoveFirst
                'Grid1.AddItem tstat & Chr(9) & s & Chr(9) & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
                Grid1.AddItem tstat & Chr(9) & s & Chr(9) & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units" & Chr(9) & "Ticket" 'jv020817
                Do Until ts.EOF
                    s = ts!sku & Chr(9)
                    s = s & skurec(Val(ts!sku)).unit & " " & skurec(Val(ts!sku)).desc & Chr(9)
                    s = s & ts(1) & Chr(9)
                    s = s & ts(2) & Chr(9)
                    s = s & ts(3)
                    Grid1.AddItem s
                    ts.MoveNext
                Loop
            End If
            ts.Close
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    ' Check Branch Orders                                                       'jv081516
    pno = "0": odate = " "
    s = "select * from brorders where branch = " & Val(Combo1) & " and netqty > 0"
    s = s & " order by orddate, plant, sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!plant <> pno Or Format(ds!orddate, "M-d-yyyy") <> odate Then
                s = "In Process" & Chr(9) & Format(ds!orddate, "dddd M-d-yyyy") & " "
                s = s & branchrec(ds!branch).branchname & " Orders" & vbCrLf
                s = s & " From: " & plantrec(Val(ds!plant)).plantname & " In Process"
                s = s & Chr(9) & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
                Grid1.AddItem s
                pno = ds!plant: odate = Format(ds!orddate, "M-d-yyyy")
            End If
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!netqty & Chr(9)
            s = s & ds!partqty & Chr(9)
            s = s & Format(ds!netqty * skurec(Val(ds!sku)).pallet, "0")
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    'Find pallet qtys in groupitems that have not been posted to trailers.      'jv081516
    s = "select * from runs where destination = '" & Combo1 & "'"               'jv081916
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "select * from trgroups where run1 = " & ds!id
            s = s & " or run2 = " & ds!id
            s = s & " or run3 = " & ds!id
            s = s & " or run4 = " & ds!id
            Set ts = wdb.Execute(s)
            If ts.BOF = False Then
                ts.MoveFirst
                Do Until ts.EOF
                    s = "select * from groupitems where groupcode = '" & ts!groupcode & "'"
                    s = s & " and groupcode not in (select groupcode from trailers)"
                    Set gs = wdb.Execute(s)
                    If gs.BOF = False Then
                        gs.MoveFirst
                        s = Format(ds!trldate, "dddd M-d-yyyy") & " "
                        s = s & branchrec(Val(Combo1)).branchname & " " & ds!trlno & vbCrLf
                        s = s & " From: " & plantrec(Val(ds!loaded)).plantname & " "
                        tstat = "In Process"
                        If ds!runstat > " " Then tstat = ds!runstat
                        s = s & tstat
                        Grid1.AddItem tstat & Chr(9) & s & Chr(9) & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
                        
                        Do Until gs.EOF
                            s = gs!sku & Chr(9)
                            s = s & skurec(Val(gs!sku)).unit & " " & skurec(Val(gs!sku)).desc & Chr(9)
                            If ts!run1 = ds!id Then
                                If gs!qty1 > 0 Then                                             'jv081916
                                    s = s & Format(gs!qty1, "0") & Chr(9) & "0" & Chr(9)
                                    s = s & Format(gs!qty1 * skurec(Val(gs!sku)).pallet, "0")
                                    Grid1.AddItem s                                             'jv081916
                                End If                                                          'jv081916
                            End If
                            If ts!run2 = ds!id Then
                                If gs!qty2 > 0 Then                                             'jv081916
                                    s = s & Format(gs!qty2, "0") & Chr(9) & "0" & Chr(9)
                                    s = s & Format(gs!qty2 * skurec(Val(gs!sku)).pallet, "0")
                                    Grid1.AddItem s                                             'jv081916
                                End If                                                          'jv081916
                            End If
                            If ts!run3 = ds!id Then
                                If gs!qty3 > 0 Then                                             'jv081916
                                    s = s & Format(gs!qty3, "0") & Chr(9) & "0" & Chr(9)
                                    s = s & Format(gs!qty3 * skurec(Val(gs!sku)).pallet, "0")
                                    Grid1.AddItem s                                             'jv081916
                                End If                                                          'jv081916
                            End If
                            If ts!run4 = ds!id Then
                                If gs!qty4 > 0 Then                                             'jv081916
                                    s = s & Format(gs!qty4, "0") & Chr(9) & "0" & Chr(9)
                                    s = s & Format(gs!qty4 * skurec(Val(gs!sku)).pallet, "0")
                                    Grid1.AddItem s                                             'jv081916
                                End If                                                          'jv081916
                            End If
                            'Grid1.AddItem s
                            gs.MoveNext
                        Loop
                    End If
                    gs.Close
                    ts.MoveNext
                Loop
            End If
            ts.Close
            ds.MoveNext
        Loop
    End If
    ds.Close
    '-------------------------------------------------------------------------------
    
    
    pno = "0": pt10 = 0: pk10 = 0: pa10 = 0                                                     'jv030216
    s = "select plantwhs, sum(thiswknewpals) from bimp where branchwhs = '" & Combo1 & "'"      'jv030216
    s = s & " and thiswknewpals > 0"                                                            'jv030216
    s = s & " group by plantwhs"                                                                'jv030216
    Set ds = wdb.Execute(s)                                                                     'jv030216
    If ds.BOF = False Then                                                                      'jv030216
        ds.MoveFirst                                                                            'jv030216
        Do Until ds.EOF                                                                         'jv030216
            If ds!plantwhs = "T10" Then pt10 = pt10 + ds(1)                                     'jv030216
            If ds!plantwhs = "K10" Then pk10 = pk10 + ds(1)                                     'jv030216
            If ds!plantwhs = "A10" Then pa10 = pa10 + ds(1)                                     'jv030216
            ds.MoveNext                                                                         'jv030216
        Loop                                                                                    'jv030216
    End If                                                                                      'jv030216
    
    s = "select plantwhs, sku, thiswknewpals from bimp where branchwhs = '" & Combo1 & "'"
    s = s & " and thiswknewpals > 0"
    s = s & " order by plantwhs, sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!plantwhs <> pno Then
                s = "New" & ds!plantwhs & Chr(9)
                If ds!plantwhs = "T10" Then s = s & pt10 & " Additional Pallets Scheduled" & vbCrLf & "From: " & "Brenham"          'jv030216
                If ds!plantwhs = "K10" Then s = s & pk10 & " Additional Pallets Scheduled" & vbCrLf & "From: " & "Broken Arrow"     'jv030216
                If ds!plantwhs = "A10" Then s = s & pa10 & " Additional Pallets Scheduled" & vbCrLf & "From: " & "Sylacauga"        'jv030216
                's = s & "Additional Pallets Scheduled" & vbCrLf & "From: "
                'If ds!plantwhs = "T10" Then s = s & "Brenham"
                'If ds!plantwhs = "K10" Then s = s & "Broken Arrow"
                'If ds!plantwhs = "A10" Then s = s & "Sylacauga"
                s = s & " Later This Week." & Chr(9)
                s = s & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
                Grid1.AddItem s
                pno = ds!plantwhs
            End If
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!thiswknewpals & Chr(9)
            s = s & "0" & Chr(9)
            s = s & Format(ds!thiswknewpals * skurec(Val(ds!sku)).pallet, "0")
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    pno = "0": pt10 = 0: pk10 = 0: pa10 = 0                                                     'jv030216
    s = "select plantwhs, sum(nextwknewpals) from bimp where branchwhs = '" & Combo1 & "'"      'jv030216
    s = s & " and nextwknewpals > 0"                                                            'jv030216
    s = s & " group by plantwhs"                                                                'jv030216
    Set ds = wdb.Execute(s)                                                                     'jv030216
    If ds.BOF = False Then                                                                      'jv030216
        ds.MoveFirst                                                                            'jv030216
        Do Until ds.EOF                                                                         'jv030216
            If ds!plantwhs = "T10" Then pt10 = pt10 + ds(1)                                     'jv030216
            If ds!plantwhs = "K10" Then pk10 = pk10 + ds(1)                                     'jv030216
            If ds!plantwhs = "A10" Then pa10 = pa10 + ds(1)                                     'jv030216
            ds.MoveNext                                                                         'jv030216
        Loop                                                                                    'jv030216
    End If                                                                                      'jv030216
    
    
    s = "select plantwhs, sku, nextwknewpals from bimp where branchwhs = '" & Combo1 & "'"
    s = s & " and nextwknewpals > 0"
    s = s & " order by plantwhs, sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!plantwhs <> pno Then
                s = "New" & ds!plantwhs & Chr(9)
                If ds!plantwhs = "T10" Then s = s & pt10 & " Pallets Scheduled" & vbCrLf & "From: " & "Brenham"          'jv030216
                If ds!plantwhs = "K10" Then s = s & pk10 & " Pallets Scheduled" & vbCrLf & "From: " & "Broken Arrow"     'jv030216
                If ds!plantwhs = "A10" Then s = s & pa10 & " Pallets Scheduled" & vbCrLf & "From: " & "Sylacauga"        'jv030216
                's = s & "Pallets Scheduled" & vbCrLf & "From: "
                'If ds!plantwhs = "T10" Then s = s & "Brenham"
                'If ds!plantwhs = "K10" Then s = s & "Broken Arrow"
                'If ds!plantwhs = "A10" Then s = s & "Sylacauga"
                s = s & " Next Week." & Chr(9)
                s = s & "Pallets" & Chr(9) & "Wraps" & Chr(9) & "Units"
                Grid1.AddItem s
                pno = ds!plantwhs
            End If
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!nextwknewpals & Chr(9)
            s = s & "0" & Chr(9)
            s = s & Format(ds!nextwknewpals * skurec(Val(ds!sku)).pallet, "0")
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 0)) = 0 Then Grid1.RowHeight(i) = Grid1.RowHeight(i) * 4
            If Grid1.TextMatrix(i, 0) = "In Process" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("\\BBC-03-FILESVR\SharedGroups\wd\html\images\inproc.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = wcolor.BackColor
            End If
            If Grid1.TextMatrix(i, 0) = "In Transit" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("\\BBC-03-FILESVR\SharedGroups\wd\html\images\loaded.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ycolor.BackColor
            End If
            If Grid1.TextMatrix(i, 0) = "Received" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("\\BBC-03-FILESVR\SharedGroups\wd\html\stock\images\bbtruck.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = gcolor.BackColor
            End If
            If Grid1.TextMatrix(i, 0) = "NewT10" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("\\BBC-03-FILESVR\SharedGroups\wd\html\images\plant500.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ncolor.BackColor
            End If
            If Grid1.TextMatrix(i, 0) = "NewK10" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("\\BBC-03-FILESVR\SharedGroups\wd\html\images\plant501.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ncolor.BackColor
            End If
            If Grid1.TextMatrix(i, 0) = "NewA10" Then
                Grid1.Row = i: Grid1.Col = 0
                Grid1.TextMatrix(i, 0) = ""
                Set Grid1.CellPicture = LoadPicture("\\BBC-03-FILESVR\SharedGroups\wd\html\images\plant502.jpg")
                Grid1.CellPictureAlignment = 4
                Grid1.RowSel = Grid1.Row
                Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = ncolor.BackColor
            End If
            'Grid1.CellPicture = LoadPicture(Picture1.Picture)
        Next i
        Grid1.Row = 1
    End If
    Grid1.FormatString = "^|<|^|^|^"
    Grid1.ColWidth(0) = 1400
    Grid1.ColWidth(1) = 4000
    Grid1.ColWidth(2) = 1600
    Grid1.ColWidth(3) = 1600
    Grid1.ColWidth(4) = 1600
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Sub refresh_vlists()
    Combo1.Clear: List1.Clear
    For i = 1 To 99
        If branchrec(i).oraloc > " " Then
            Combo1.AddItem Format(Val(branchrec(i).branchno), "000")
            List1.AddItem branchrec(i).branchname
        End If
    Next i
    Combo1.ListIndex = 0
End Sub


Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    Label2 = List1
    refresh_grid
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub Command2_Click()
    'brbarcodes.tktkey = Grid1.TextMatrix(Grid1.Row, 5)
    brbarcodes.tktkey = Label2
    brbarcodes.Show
End Sub

Private Sub Form_Load()
    refresh_vlists
    Me.Height = whssales.Height
    Me.Top = whssales.Top
    Me.Left = whssales.Width - Me.Width
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 3)
End Sub

