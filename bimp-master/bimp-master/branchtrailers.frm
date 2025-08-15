VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form branchtrailers 
   Caption         =   "Planned Trailers vs Actual"
   ClientHeight    =   10245
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   ScaleHeight     =   10245
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check4 
      Caption         =   "Active"
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
      Left            =   5760
      TabIndex        =   16
      Top             =   360
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Next Week"
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
      Left            =   8040
      TabIndex        =   15
      Top             =   360
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "This Week"
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
      Left            =   8040
      TabIndex        =   14
      Top             =   120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View Branch Orders"
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
      Left            =   5760
      TabIndex        =   13
      Top             =   120
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2775
      Left            =   0
      TabIndex        =   12
      Top             =   7320
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4895
      _Version        =   327680
      ForeColor       =   8421376
      BackColorFixed  =   12648384
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   9000
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   9340
      _Version        =   327680
      BackColorFixed  =   12648447
      FocusRect       =   0
   End
   Begin VB.Label bkey 
      Caption         =   "bkey"
      Height          =   255
      Left            =   11400
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label gcolor 
      BackColor       =   &H00FFC0C0&
      Caption         =   "gcolor"
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
      Left            =   8640
      TabIndex        =   9
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label ycolor 
      BackColor       =   &H0080FFFF&
      Caption         =   "ycolor"
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
      Left            =   6720
      TabIndex        =   8
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label wcolor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "wcolor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label tdate 
      Caption         =   "tdate"
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
      Left            =   10200
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label bcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      Caption         =   "bcolor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label rcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "rcolor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Branch:"
      BeginProperty Font 
         Name            =   "Arial"
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
      Top             =   120
      Width           =   855
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
   End
   Begin VB.Menu postmenu 
      Caption         =   "Post"
      Visible         =   0   'False
      Begin VB.Menu pbords 
         Caption         =   "Post to Branch Orders"
      End
   End
   Begin VB.Menu ordmenu 
      Caption         =   "Edit Orders"
      Visible         =   0   'False
      Begin VB.Menu psttwk 
         Caption         =   "Post back to This Week"
      End
      Begin VB.Menu pstnwk 
         Caption         =   "Post back to Next Week"
      End
      Begin VB.Menu edodate 
         Caption         =   "Change Order Date"
      End
   End
End
Attribute VB_Name = "branchtrailers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_vlists()
    Dim i As Integer
    Combo1.Clear: List1.Clear
    If Val(bkey) > 0 Then
        i = Val(bkey)
        If branchrec(i).oraloc > "0" Then
            Combo1.AddItem branchrec(i).oraloc & "-" & branchrec(i).branchname
            List1.AddItem branchrec(i).branchno
        End If
    Else
        For i = 1 To 99
            If branchrec(i).oraloc > "0" Then
                Combo1.AddItem branchrec(i).oraloc & "-" & branchrec(i).branchname
                List1.AddItem branchrec(i).branchno
            End If
        Next i
    End If
    Combo1.ListIndex = 0
End Sub

Private Sub supply_days()
    Dim ds As ADODB.Recordset, s As String, i As Integer
    For i = 1 To Grid1.Rows - 1
        s = "select ohpct, bimpstatus from bimp where plantwhs = '" & Grid1.TextMatrix(i, 0) & "'"
        s = s & " and branchwhs = '" & Left(Combo1, 3) & "'"
        s = s & " and sku = '" & Grid1.TextMatrix(i, 1) & "'"
        Set ds = wdb.Execute(s)
        If ds.BOF = False Then
            Grid1.TextMatrix(i, 7) = ds!bimpstatus
            Grid1.TextMatrix(i, 8) = Format(ds!ohpct * 30, "0")
        End If
        ds.Close
    Next i
End Sub

Private Sub refresh_grid1()
    Dim ds As ADODB.Recordset, s As String
    Dim ts As ADODB.Recordset, gs As ADODB.Recordset                        'jv081516
    Dim i As Integer, crow As Boolean, a10flag As Boolean, k10flag As Boolean, bwhs As String
    Dim rc As String
    crow = True
    a10flag = False
    k10flag = False
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Cols = 9: Grid1.Rows = 1
    Grid1.FixedCols = 3
    Grid1.Clear
    bwhs = Left(Combo1, 3)
    s = "select plantwhs,sku,thiswknewpals,nextwknewpals,bimpstatus from bimp where branchwhs = '" & bwhs & "'"
    If bwhs = "K10" Then
        s = "select plantwhs,sku,thiswknewpals,nextwknewpals,bimpstatus from bimp where branchwhs in ('047', 'K10')"
        s = s & " and plantwhs in ('T10', 'A10')"
    End If
    If bwhs = "A10" Then
        s = "select plantwhs,sku,thiswknewpals,nextwknewpals,bimpstatus from bimp where branchwhs in ('052', 'A10')"
        s = s & " and plantwhs in ('T10', 'K10')"
    End If
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!thiswknewpals > 0 Then
                s = ds!plantwhs & Chr(9)
                s = s & ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                s = s & "." & Chr(9)
                s = s & ds!thiswknewpals & Chr(9)
                s = s & "0" & Chr(9)
                s = s & "This Week" & Chr(9)
                s = s & ds!bimpstatus
                If Check2 = 1 Then Grid1.AddItem s              'jv093016
            End If
            If ds!nextwknewpals > 0 Then
                s = ds!plantwhs & Chr(9)
                s = s & ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                s = s & ".." & Chr(9)
                s = s & ds!nextwknewpals & Chr(9)
                s = s & "0" & Chr(9)
                s = s & "Next Week" & Chr(9)
                s = s & ds!bimpstatus
                If Check3 = 1 Then Grid1.AddItem s              'jv093016
            End If
            If ds!plantwhs = "K10" Then k10flag = True
            If ds!plantwhs = "A10" Then a10flag = True
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select * from brorders where branch = " & List1 & " and netqty <> 0"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = " "
            If ds!plant = 50 Then s = "T10"
            If ds!plant = 51 Then s = "K10"
            If ds!plant = 52 Then s = "A10"
            s = s & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & Format(ds!orddate, "MM-dd-yyyy") & Chr(9)
            s = s & Format(ds!netqty, "0") & Chr(9)
            s = s & "0" & Chr(9)
            s = s & "Orders"
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
        
    If Check4 = 1 Then
    'Find pallet qtys in groupitems that have not been posted to trailers.      'jv081516
    s = "select id, loaded, trldate from runs where destination = '" & List1 & "'"          'jv081916
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
                        Do Until gs.EOF
                            s = " "
                            If ds!loaded = 50 Then s = "T10"
                            If ds!loaded = 51 Then s = "K10"
                            If ds!loaded = 52 Then s = "A10"
                            s = s & Chr(9)
                            s = s & gs!sku & Chr(9)
                            s = s & skurec(Val(gs!sku)).unit & " " & skurec(Val(gs!sku)).desc & Chr(9)
                            s = s & Format(ds!trldate, "MM-dd-yyyy") & Chr(9)
                            If ts!run1 = ds!id Then
                                If gs!qty1 > 0 Then                                 'jv081916
                                    s = s & Format(gs!qty1, "0") & Chr(9)           'jv081916
                                Else                                                'jv081916
                                    s = s & "0" & Chr(9)                            'jv081916
                                End If                                              'jv081916
                            End If                                                  'jv081916
                            If ts!run2 = ds!id Then
                                If gs!qty2 > 0 Then                                 'jv081916
                                    s = s & Format(gs!qty2, "0") & Chr(9)           'jv081916
                                Else                                                'jv081916
                                    s = s & "0" & Chr(9)                            'jv081916
                                End If                                              'jv081916
                            End If                                                  'jv081916
                            If ts!run3 = ds!id Then
                                If gs!qty3 > 0 Then                                 'jv081916
                                    s = s & Format(gs!qty3, "0") & Chr(9)           'jv081916
                                Else                                                'jv081916
                                    s = s & "0" & Chr(9)                            'jv081916
                                End If                                              'jv081916
                            End If                                                  'jv081916
                            If ts!run4 = ds!id Then
                                If gs!qty4 > 0 Then                                 'jv081916
                                    s = s & Format(gs!qty4, "0") & Chr(9)           'jv081916
                                Else                                                'jv081916
                                    s = s & "0" & Chr(9)                            'jv081916
                                End If                                              'jv081916
                            End If                                                  'jv081916
                            s = s & "0" & Chr(9)
                            s = s & "* " & gs!groupcode
                            Grid1.AddItem s
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
    
    s = "select * from trailers where branch = " & List1 & " and pallets > 0 and plant = 50"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = " "
            If ds!plant = 50 Then s = "T10"
            If ds!plant = 51 Then s = "K10"
            If ds!plant = 52 Then s = "A10"
            s = s & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & Format(ds!shipdate, "MM-dd-yyyy") & Chr(9)
            s = s & Format(ds!pallets * -1, "0") & Chr(9)
            s = s & "0" & Chr(9)
            s = s & ds!groupcode
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    'If a10flag = True Then
    If a10flag = True And plant_server_status("A10") = True Then            'jv010417
        s = "select * from trailers where branch = " & List1 & " and pallets > 0 and plant = 52"
        Set ds = a10shipdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = " "
                If ds!plant = 50 Then s = "T10"
                If ds!plant = 51 Then s = "K10"
                If ds!plant = 52 Then s = "A10"
                s = s & Chr(9)
                s = s & ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                s = s & Format(ds!shipdate, "MM-dd-yyyy") & Chr(9)
                s = s & Format(ds!pallets * -1, "0") & Chr(9)
                s = s & "0" & Chr(9)
                s = s & ds!groupcode
                Grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    'If k10flag = True Then
    If k10flag = True And plant_server_status("K10") = True Then            'jv010417
        s = "select * from trailers where branch = " & List1 & " and pallets > 0 and plant = 51"
        s = s & " and ra_flag = 'N'"
        Set ds = k10shipdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                s = " "
                If ds!plant = 50 Then s = "T10"
                If ds!plant = 51 Then s = "K10"
                If ds!plant = 52 Then s = "A10"
                s = s & Chr(9)
                s = s & ds!sku & Chr(9)
                s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
                s = s & Format(ds!shipdate, "MM-dd-yyyy") & Chr(9)
                s = s & Format(ds!pallets * -1, "0") & Chr(9)
                s = s & "0" & Chr(9)
                s = s & ds!groupcode
                Grid1.AddItem s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    
    End If
    Call supply_days
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 3
    Grid1.Sort = 5
    Grid1.FillStyle = flexFillRepeat
    psku = " ": pnet = 0: rc = "W"
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 1) <> psku Then
                crow = Not crow
                psku = Grid1.TextMatrix(i, 1)
                pnet = 0
            Else
                Grid1.TextMatrix(i, 2) = "_ "
            End If
            pnet = pnet + Val(Grid1.TextMatrix(i, 4))
            Grid1.TextMatrix(i, 5) = pnet
            Grid1.Row = i: Grid1.RowSel = i
            Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
            If crow = True Then
                Grid1.CellForeColor = bcolor.ForeColor
            Else
                Grid1.CellForeColor = rcolor.ForeColor
                'Grid1.CellBackColor = rcolor.BackColor
            End If
            Grid1.Col = 0: Grid1.ColSel = 3
            'If crow = True Then
            '    Grid1.CellBackColor = bcolor.BackColor
            'Else
            '    Grid1.CellBackColor = rcolor.BackColor
            'End If
            If Grid1.TextMatrix(i, 7) > " " Then rc = Grid1.TextMatrix(i, 7)
            Grid1.Col = 3: Grid1.ColSel = Grid1.Cols - 1
            If rc = "W" Then
                Grid1.CellBackColor = wcolor.BackColor
            End If
            If rc = "Y" Then
                Grid1.CellBackColor = ycolor.BackColor
            End If
            If rc = "B" Then
                Grid1.CellBackColor = bcolor.BackColor
            End If
            If rc = "G" Then
                Grid1.CellBackColor = gcolor.BackColor
            End If
            
            'If Grid1.TextMatrix(i, 7) > " " Then
            '    Grid1.Col = 2: Grid1.ColSel = Grid1.Cols - 1
            '    If Grid1.TextMatrix(i, 7) = "W" Then
            '        Grid1.CellBackColor = wcolor.BackColor
            '    End If
            '    If Grid1.TextMatrix(i, 7) = "Y" Then
            '        Grid1.CellBackColor = ycolor.BackColor
            '    End If
            '    If Grid1.TextMatrix(i, 7) = "B" Then
            '        Grid1.CellBackColor = bcolor.BackColor
            '    End If
            '    If Grid1.TextMatrix(i, 7) = "G" Then
            '        Grid1.CellBackColor = gcolor.BackColor
            '    End If
            'End If
        Next i
        Grid1.Row = 1
    End If
    Grid1.FormatString = "^Plant|^SKU|<Product|^Date|^Pallets|^Net|^Group||^Days"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3500
    Grid1.ColWidth(3) = 1400
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1400
    Grid1.ColWidth(7) = 0 '1000
    Grid1.ColWidth(8) = 800
    Grid1.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub refresh_brorders()
    Dim ds As ADODB.Recordset, s As String
    Dim t4 As Long, t5 As Long, t6 As Long, t7 As Long, t8 As Long              'jv082616
    t4 = 0: t5 = 0: t6 = 0: t7 = 0: t8 = 0                                      'jv082616
    If Val(List1) = 0 Then Exit Sub
    Screen.MousePointer = 11
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 9: Grid2.FixedCols = 4
    s = "select * from brorders where orddate = '" & tdate & "'"
    s = s & " and branch = " & List1 & " order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & Format(ds!orddate, "MM-dd-yyyy") & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!ordqty & Chr(9)
            s = s & ds!grpqty & Chr(9)
            s = s & ds!netqty & Chr(9)
            s = s & ds!altflag & Chr(9)
            s = s & ds!partqty
            t4 = t4 + ds!ordqty                                                 'jv082616
            t5 = t5 + ds!grpqty                                                 'jv082616
            t6 = t6 + ds!netqty                                                 'jv082616
            If ds!altflag = "Y" Then t7 = t7 + 1                                'jv082616
            t8 = t8 + ds!partqty                                                'jv082616
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid2.Rows > 1 Then                                                      'jv082616
        Grid2.FillStyle = flexFillRepeat
        For i = 1 To Grid2.Rows - 1
            If Val(Grid2.TextMatrix(i, 6)) <> 0 Then
                Grid2.Row = i: Grid2.RowSel = i
                Grid2.Col = 4: Grid2.ColSel = 6
                Grid2.CellBackColor = ycolor.BackColor
            End If
            If Grid2.TextMatrix(i, 7) = "Y" Then
                Grid2.Row = i: Grid2.RowSel = i
                Grid2.Col = 7: Grid2.ColSel = 7
                Grid2.CellBackColor = ycolor.BackColor
            End If
        Next i
        s = Chr(9) & Chr(9) & Chr(9) & "Summary" & Chr(9)                       'jv082616
        s = s & t4 & Chr(9) & t5 & Chr(9) & t6 & Chr(9) & t7 & Chr(9) & t8      'jv082616
        Grid2.AddItem s                                                         'jv082616
        Grid2.Row = 1
    End If                                                                      'jv082616
    's = "^ID|^Date|^SKU|<Product|^Order|^Grouped|^Net|^Alt|^Wraps"
    s = "|^Date|^SKU|<Product|^Order|^Grouped|^Net|^Alt|^Wraps"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 0 '1000
    Grid2.ColWidth(1) = 1200
    Grid2.ColWidth(2) = 800
    Grid2.ColWidth(3) = 3000
    Grid2.ColWidth(4) = 1000
    Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 1000
    Grid2.ColWidth(7) = 800
    Grid2.ColWidth(8) = 1000
    Grid2.Redraw = True
    Screen.MousePointer = 0
End Sub

Private Sub bkey_Change()
    refresh_vlists
End Sub

Private Sub Check1_Click()
    Call Form_Resize
    If Check1.Value = 1 Then
        refresh_brorders
        Grid2.Visible = True
    Else
        Grid2.Visible = False
    End If
End Sub

Private Sub Check2_Click()
    refresh_grid1
End Sub

Private Sub Check3_Click()
    refresh_grid1
End Sub

Private Sub Check4_Click()
    refresh_grid1
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
    refresh_grid1
    If Grid2.Visible = True Then refresh_brorders
End Sub

Private Sub edodate_Click()
    Dim s As String, i As Integer, k As Integer
    's = Format(Now, "M-dd-yyyy")
    s = Grid2.TextMatrix(Grid2.Row, 1)
    s = InputBox("Order Date:", "New Order Date....", s)
    If Len(s) = 0 Then Exit Sub
    If IsDate(s) = False Then Exit Sub
    s = "Update brorders set orddate = '" & s & "' Where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    'MsgBox s
    wdb.Execute s
    i = Grid1.Row: k = Grid2.Row
    refresh_grid1
    DoEvents
    refresh_brorders
    DoEvents
    If i < Grid1.Rows Then Grid1.Row = i
    If k < Grid2.Rows Then Grid2.Row = k
End Sub

Private Sub Form_Load()
    Me.Left = bimpbanner.Width - Me.Width
    Me.Top = bimpbanner.Label2.Top
    'Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    If plant_server_status("A10") = True Then                   'jv010417
        Set a10shipdb = CreateObject("ADODB.Connection")
        a10shipdb.Open a10ship
    Else                                                        'jv010417
        MsgBox "Sylacauga A10 Server is unavailable.", vbOKOnly + vbInformation, "A10 is offline..."
    End If                                                      'jv010417
    If plant_server_status("K10") = True Then                   'jv010417
        Set k10shipdb = CreateObject("ADODB.Connection")
        k10shipdb.Open k10ship
    Else                                                        'jv010417
        MsgBox "Broken Arrow K10 Server is unavailable.", vbOKOnly + vbInformation, "K10 is offline..."
    End If                                                      'jv010417
    tdate = Format(Now, "MM-dd-yyyy")
    refresh_vlists
    If bimpbanner.Command1.Visible = False Then
        pbords.Enabled = False
    End If
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 180
    Grid2.Width = Grid1.Width
    If Me.Height > 2000 Then
        If Check1.Value = 1 Then
            'Grid1.Height = Me.Height - ((Combo1.Height * 4) + (Grid2.Height * 1.5))
            Grid1.Height = Grid2.Top - Grid1.Top
            Grid2.Height = Me.Height - ((Combo1.Height * 4) + Grid1.Height)
        Else
            Grid1.Height = Me.Height - (Combo1.Height * 4)
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'MsgBox "bye from trailers"
    If plant_server_status("A10") = True Then a10shipdb.Close       'jv010417
    If plant_server_status("K10") = True Then k10shipdb.Close       'jv010417
End Sub

Private Sub Grid1_Click()
    If Grid1.Row = 0 Then Exit Sub
    If Len(Grid1.TextMatrix(Grid1.Row, 3)) > 5 Then tdate = Grid1.TextMatrix(Grid1.Row, 3)
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu postmenu
End Sub

Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu ordmenu
End Sub

Private Sub Grid2_RowColChange()
    psttwk.Enabled = False: edodate.Enabled = False: pstnwk.Enabled = False
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) > 0 And Val(Grid2.TextMatrix(Grid2.Row, 6)) > 0 Then
        psttwk.Enabled = True: edodate.Enabled = True: pstnwk.Enabled = True
    End If
End Sub

Private Sub List1_Click()
    refresh_grid1
    If Grid2.Visible = True Then refresh_brorders
End Sub

Private Sub pbords_Click()
    Dim s As String, i As Integer, pcode As Integer, pqty As Integer, palt As String, z As Long
    pcode = 0
    i = Grid1.Row
    If Grid1.TextMatrix(i, 0) = "T10" Then pcode = 50
    If Grid1.TextMatrix(i, 0) = "K10" Then pcode = 51
    If Grid1.TextMatrix(i, 0) = "A10" Then pcode = 52
    If pcode = 0 Then Exit Sub
    s = InputBox("Order Date:", "Branch Orders - Date...", tdate)
    If Len(s) = 0 Then Exit Sub
    If IsDate(s) = False Then Exit Sub
    tdate = Format(s, "MM-dd-yyyy")
    s = InputBox("Pallet Qty:", "Branch Orders - Pallets...", Grid1.TextMatrix(i, 5))
    If Len(s) = 0 Then Exit Sub
    If Val(s) < 0 Then Exit Sub
    If MsgBox("Use product as an alternate?", vbQuestion + vbYesNo + vbDefaultButton2, "alternate.....") = vbYes Then
        palt = "Y"
    Else
        palt = "N"
    End If
    pqty = Val(s)
    z = wd_seq("brorders")
    'z = Grid1.Row
    s = "Insert into brorders (id, plant, branch, account, sku, orddate, ordqty, grpqty, netqty, altflag, partqty)"
    s = s & " Values (" & z
    s = s & ", " & pcode
    s = s & ", " & List1
    s = s & ", '......'"
    s = s & ", '" & Grid1.TextMatrix(i, 1) & "'"
    s = s & ", '" & tdate & "'"
    s = s & ", " & pqty
    s = s & ", 0, " & pqty & ", '" & palt & "', 0)"
    'MsgBox s
    wdb.Execute s
    Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) - pqty
    Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) - pqty
    If Grid1.TextMatrix(i, 6) = "This Week" Then
        s = "Update bimp set thiswknewpals = " & Grid1.TextMatrix(i, 4)
        s = s & " Where plantwhs = '" & Grid1.TextMatrix(i, 0) & "'"
        If List1 = "52" Then
            s = s & " and branchwhs = '052'"
        Else
            If List1 = "47" Then
                s = s & " and branchwhs = '047'"
            Else
                s = s & " and branchwhs = '" & Left(Combo1, 3) & "'"
            End If
        End If
        s = s & " and sku = '" & Grid1.TextMatrix(i, 1) & "'"
        'MsgBox s
        wdb.Execute s
    End If
    If Grid1.TextMatrix(i, 6) = "Next Week" Then
        s = "Update bimp set nextwknewpals = " & Grid1.TextMatrix(i, 4)
        s = s & " Where plantwhs = '" & Grid1.TextMatrix(i, 0) & "'"
        If List1 = "52" Then
            s = s & " and branchwhs = '052'"
        Else
            If List1 = "47" Then
                s = s & " and branchwhs = '047'"
            Else
                s = s & " and branchwhs = '" & Left(Combo1, 3) & "'"
            End If
        End If
        s = s & " and sku = '" & Grid1.TextMatrix(i, 1) & "'"
        'MsgBox s
        wdb.Execute s
    End If
    refresh_grid1
    DoEvents
    If i < Grid1.Rows Then
        Grid1.Row = i
        Grid1.Col = 4
    End If
    If Grid2.Visible = True Then refresh_brorders
End Sub

Private Sub prtmenu_Click()
    Dim rt As String, rh As String, rf As String
    rt = "Branch Trailers"
    rh = Combo1
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    'htdc(0) = "lightcyan": gndc(0) = Me.bcolor.BackColor
    'htdc(1) = "yellow": gndc(1) = Me.ycolor.BackColor
    'htdc(2) = "white": gndc(2) = Me.wcolor.BackColor
    Grid1.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "lightyellow", "white")
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Grid1.Redraw = True: Grid1.Row = 1
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "lightyellow", "white")
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Grid1.Redraw = True: Grid1.Row = 1
        Exit Sub
    End If
End Sub

Private Sub pstnwk_Click()
    Dim i As Integer, s As String, k As Integer
    s = Grid2.TextMatrix(Grid2.Row, 6)
    s = InputBox("Pallets:", "Pallet qty....", s)
    If Len(s) = 0 Then Exit Sub
    i = Val(s)
    If Val(Grid2.TextMatrix(Grid2.Row, 4)) = i Then      'ordqty = netqty
        s = "Delete from brorders where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    Else
        s = "Update brorders set ordqty = ordqty - " & i
        s = s & ", netqty = netqty - " & i
        s = s & " Where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    End If
    'MsgBox s
    wdb.Execute s
    s = "Update bimp set nextwknewpals = 0"
    s = s & " Where plantwhs = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
    's = s & " and branchwhs = '" & Left(Combo1, 3) & "'"
    s = s & " and branchwhs = '" & Format(Val(List1), "000") & "'"                  'jv080417
    s = s & " and sku = '" & Grid2.TextMatrix(Grid2.Row, 2) & "'"
    s = s & " and nextwknewpals is null"
    'MsgBox s
    wdb.Execute s
    s = "Update bimp set nextwknewpals = nextwknewpals + " & i
    s = s & " Where plantwhs = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
    's = s & " and branchwhs = '" & Left(Combo1, 3) & "'"
    s = s & " and branchwhs = '" & Format(Val(List1), "000") & "'"                  'jv080417
    s = s & " and sku = '" & Grid2.TextMatrix(Grid2.Row, 2) & "'"
    'MsgBox s
    wdb.Execute s
    i = Grid1.Row: k = Grid2.Row
    refresh_grid1
    DoEvents
    refresh_brorders
    DoEvents
    If i < Grid1.Rows Then Grid1.Row = i
    If k < Grid2.Rows Then Grid2.Row = k
End Sub

Private Sub psttwk_Click()
    Dim i As Integer, s As String, k As Integer
    s = Grid2.TextMatrix(Grid2.Row, 6)
    s = InputBox("Pallets:", "Pallet qty....", s)
    If Len(s) = 0 Then Exit Sub
    i = Val(s)
    If Val(Grid2.TextMatrix(Grid2.Row, 4)) = i Then      'ordqty = netqty
        s = "Delete from brorders where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    Else
        s = "Update brorders set ordqty = ordqty - " & i
        s = s & ", netqty = netqty - " & i
        s = s & " Where id = " & Grid2.TextMatrix(Grid2.Row, 0)
    End If
    'MsgBox s
    wdb.Execute s
    s = "Update bimp set thiswknewpals = 0"
    s = s & " Where plantwhs = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
    's = s & " and branchwhs = '" & Left(Combo1, 3) & "'"
    s = s & " and branchwhs = '" & Format(Val(List1), "000") & "'"                  'jv080417
    s = s & " and sku = '" & Grid2.TextMatrix(Grid2.Row, 2) & "'"
    s = s & " and thiswknewpals is null"
    'MsgBox s
    wdb.Execute s
    s = "Update bimp set thiswknewpals = thiswknewpals + " & i
    s = s & " Where plantwhs = '" & Grid1.TextMatrix(Grid1.Row, 0) & "'"
    's = s & " and branchwhs = '" & Left(Combo1, 3) & "'"
    s = s & " and branchwhs = '" & Format(Val(List1), "000") & "'"                  'jv080417
    s = s & " and sku = '" & Grid2.TextMatrix(Grid2.Row, 2) & "'"
    'MsgBox s
    wdb.Execute s
    i = Grid1.Row: k = Grid2.Row
    refresh_grid1
    DoEvents
    refresh_brorders
    DoEvents
    If i < Grid1.Rows Then Grid1.Row = i
    If k < Grid2.Rows Then Grid2.Row = k
End Sub

Private Sub tdate_Change()
    If Grid2.Visible = True Then refresh_brorders
End Sub

