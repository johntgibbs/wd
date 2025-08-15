VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form13 
   Caption         =   "Form13"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12630
   LinkTopic       =   "Form13"
   ScaleHeight     =   10950
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Post && Print Order"
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
      Left            =   8760
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   7815
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   13785
      _Version        =   327680
      FocusRect       =   0
      GridLines       =   2
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2778
      _Version        =   327680
      ForeColor       =   12582912
      BackColorFixed  =   12632319
      FocusRect       =   0
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alternates"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   10200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wraps"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   10200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pallets"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   10200
      Width           =   1215
   End
   Begin VB.Label ta 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   10200
      Width           =   975
   End
   Begin VB.Label tw 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   10200
      Width           =   975
   End
   Begin VB.Label tp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   10200
      Width           =   975
   End
   Begin VB.Label gcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Surplus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label bcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Month Supply"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label ycolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "< Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label wcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<2 Weeks"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label rkey 
      Caption         =   "rkey"
      Height          =   255
      Left            =   8280
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label brcode 
      Caption         =   "brcode"
      Height          =   255
      Left            =   6840
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label tdate 
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Trailer Date(s):"
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
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub post_order()
    Dim ds As ADODB.Recordset, s As String, i As Integer, gc As String, wno As Integer
    Dim pqty As Integer, wqty As Integer, aflg As String, pid As Long, uqty As Long
    For i = 1 To Grid2.Rows - 1
        pqty = Val(Grid2.TextMatrix(i, 3))
        wqty = Val(Grid2.TextMatrix(i, 4))
        aflg = Grid2.TextMatrix(i, 5)
        pid = Val(Grid2.TextMatrix(i, 9))
        uqty = pqty * skurec(Val(Grid2.TextMatrix(i, 0))).pallet
        uqty = uqty + (wqty * skurec(Val(Grid2.TextMatrix(i, 0))).wrapunits)
        'If pqty > 0 Or wqty > 0 Or aflg = "Y" Then
        '    If pid = 0 Then
        '        pid = wd_seq("Brorders")
        '        s = "Insert into Brorders (id, plant, branch, account, sku, orddate, ordqty, grpqty, netqty,"
        '        s = s & " altflag, partqty, runid) Values (" & pid
        '        s = s & ", " & Grid1.TextMatrix(Grid1.Row, 2)
        '        s = s & ", " & Me.brcode
        '        s = s & ", '......'"
        '        s = s & ", '" & Grid2.TextMatrix(i, 0) & "'"
        '        s = s & ", '" & Grid1.TextMatrix(Grid1.Row, 1) & "'"
        '        s = s & ", " & pqty
        '        s = s & ", 0"
        '        s = s & ", " & pqty
        '        If aflg = "Y" Then
        '            s = s & ", 'Y'"
        '        Else
        '            s = s & ", 'N'"
        '        End If
        '        s = s & ", " & wqty
        '        s = s & ", " & rkey & ")"
        '        'MsgBox s
        '        wdb.Execute s
        '    Else
        '        s = "Update brorders set ordqty = " & pqty
        '        s = s & ", grpqty = 0"
        '        s = s & ", netqty = " & pqty
        '        s = s & ", partqty = " & wqty
        '        If aflg = "Y" Then
        '            s = s & ", altflag = 'Y'"
        '        Else
        '            s = s & ", altflag = 'N'"
        '        End If
        '        s = s & " Where id = " & pid
        '        'MsgBox s
        '        wdb.Execute s
        '    End If
        'Else
        '    If pid <> 0 Then
        '        s = "Delete from brorders where id = " & pid
        '        'MsgBox s
        '        wdb.Execute s
        '    End If
        'End If
        
        If pqty > 0 Or wqty > 0 Or aflg = "Y" Then
            If pid = 0 Then
                If Grid1.TextMatrix(Grid1.Row, 2) = "50" Then
                    gc = UCase(Left(Format(Grid1.TextMatrix(Grid1.Row, 1), "ddd"), 2))
                    gc = gc & "-" & Right(Grid1.TextMatrix(Grid1.Row, 4), 1) & Format(Me.brcode, "00")
                Else
                    gc = "T" & Format(Grid1.TextMatrix(Grid1.Row, 1), "dd")
                    gc = gc & Format(Val(Me.brcode), "00") & Right(Grid1.TextMatrix(Grid1.Row, 4), 1)
                End If
                'pid = 99
                pid = wd_seq("Trailers")
                s = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno"
                s = s & ", sku, pallets, wraps, units, whs_num, pb_flag, ra_flag) values (" & pid
                s = s & ", " & rkey
                s = s & ", '" & gc & "'"
                s = s & ", " & Grid1.TextMatrix(Grid1.Row, 2)
                s = s & ", " & Me.brcode
                If aflg = "Y" Then
                    s = s & ", 'ALT'"
                Else
                    s = s & ", '......'"
                End If
                s = s & ", '" & Grid1.TextMatrix(Grid1.Row, 1) & "'"
                s = s & ", '" & Right(Grid1.TextMatrix(Grid1.Row, 4), 2) & "'"
                s = s & ", '" & Grid2.TextMatrix(i, 0) & "'"
                s = s & ", " & pqty
                s = s & ", " & wqty
                s = s & ", " & uqty
                wno = skurec(Val(Grid2.TextMatrix(i, 0))).whs
                If wno = 0 Or wno > 4 Then wno = 5
                If Grid1.TextMatrix(Grid1.Row, 2) = "51" Then wno = 14
                If Grid1.TextMatrix(Grid1.Row, 2) = "52" Then wno = 15
                s = s & ", " & wno
                s = s & ", 'N', 'N')"
                'MsgBox s
                wdb.Execute s
            Else
                s = "Update trailers set pallets = " & pqty
                s = s & ", wraps = " & wqty
                s = s & ", units = " & uqty
                If aflg = "Y" Then
                    s = s & ", account = 'ALT'"
                Else
                    s = s & ", account = '......'"
                End If
                s = s & " Where id = " & pid
                'MsgBox s
                wdb.Execute s
            End If
        Else
            If pid <> 0 Then
                s = "Delete from trailers where id = " & pid
                'MsgBox s
                wdb.Execute s
            End If
        End If
        
    Next i
    i = Grid2.Row
    refresh_grid2
    Grid2.Row = i
End Sub

Private Sub refresh_tdates()
    Dim ds As ADODB.Recordset, s As String
    s = "select orddates from wdstatus where id = 1"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        tdate.Caption = ds(0)
    End If
    ds.Close
End Sub

Private Sub refresh_grid1()
    Dim ds As ADODB.Recordset, s As String
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    s = "select id, loaded, locname, trlno, trlsize, trldate, startime, oc from runs"
    s = s & " where destination = '" & Me.brcode & "'"              'jv081916
    If Len(tdate) = 10 Then
        s = s & " and trldate = '" & tdate & "'"
    Else
        s = s & " and trldate in '" & Left(tdate, 10) & "', '" & Right(tdate, 10) & "')"
    End If
    s = s & " order by trldate, trlno"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & Format(ds!trldate, "m-dd-yyyy") & Chr(9)
            s = s & ds!loaded & Chr(9)
            s = s & plantrec(ds!loaded).plantname & Chr(9)
            s = s & ds!locname & " " & ds!trlno & Chr(9)
            s = s & ds!trlsize & Chr(9)
            s = s & Format(ds!startime, "h:mm am/pm")
            Grid1.AddItem s
            ds.MoveNext
        Loop
    Else
        rkey = "0"
    End If
    ds.Close
    Grid1.FormatString = "^ID|^Date|^Whs|<Plant|<Trailer|^Pallets|^Start"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 700
    Grid1.ColWidth(3) = 1400
    Grid1.ColWidth(4) = 2000
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 900
    Grid1.Redraw = True
    If Grid1.Rows > 1 Then
        Grid1.Row = 1
        Grid1_RowColChange
    End If
End Sub

Private Sub refresh_grid2()
    Dim ds As ADODB.Recordset, s As String
    Command1.Visible = False
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 10
    If Val(rkey) = 0 Then
        Grid2.Redraw = True
        Exit Sub
    End If
    s = "select * from bimp where branchwhs = '" & branchrec(Val(brcode)).oraloc & "'"
    s = s & " and ((plantwhs = '" & plantrec(Val(Grid1.TextMatrix(Grid1.Row, 2))).orawhs & "'"
    s = s & " and outflag = 'N') or plantwhs = 'VENDOR')"
    's = "select * from bimp where plantwhs = '" & plantrec(Val(Grid1.TextMatrix(Grid1.Row, 2))).orawhs & "'"
    's = s & " and branchwhs = '" & branchrec(Val(brcode)).oraloc & "'"
    's = s & " and outflag = 'N'"
    s = s & " order by sku"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & Format(ds!onhand + ds!onorder, "0") & Chr(9)
            s = s & Chr(9) & Chr(9) & Chr(9)
            s = s & ds!undiff & Chr(9)
            s = s & ds!paldiff & Chr(9)
            s = s & ds!needqty & Chr(9)
            s = s & ""
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    's = "select * from brorders where runid = " & Val(rkey.Caption)
    s = "select * from trailers where runid = " & Val(rkey.Caption)
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 0 To Grid2.Rows - 1
                If Grid2.TextMatrix(i, 0) = ds!sku Then
                    'Grid2.TextMatrix(i, 3) = Format(ds!ordqty, "#")
                    'Grid2.TextMatrix(i, 4) = Format(ds!partqty, "#")
                    'If ds!altflag = "Y" Then Grid2.TextMatrix(i, 5) = ds!altflag
                    Grid2.TextMatrix(i, 3) = Format(ds!pallets, "#")
                    Grid2.TextMatrix(i, 4) = Format(ds!wraps, "#")
                    If UCase(ds!account) = "ALT" Then Grid2.TextMatrix(i, 5) = "Y"
                    Grid2.TextMatrix(i, 9) = ds!id
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid2.Rows > 1 Then
        Grid2.FillStyle = flexFillRepeat
        For i = 1 To Grid2.Rows - 1
            If Val(Grid2.TextMatrix(i, 8)) > 0 Then
                Grid2.Row = i: Grid2.RowSel = i
                Grid2.Col = 0: Grid2.ColSel = 2
                Grid2.CellBackColor = wcolor.BackColor
                Grid2.Col = 6: Grid2.ColSel = Grid2.Cols - 1
                Grid2.CellBackColor = wcolor.BackColor
            Else
                If Val(Grid2.TextMatrix(i, 7)) = 0 Then
                    Grid2.Row = i: Grid2.RowSel = i
                    Grid2.Col = 0: Grid2.ColSel = 2
                    Grid2.CellBackColor = bcolor.BackColor
                    Grid2.Col = 6: Grid2.ColSel = Grid2.Cols - 1
                    Grid2.CellBackColor = bcolor.BackColor
                Else
                    If Val(Grid2.TextMatrix(i, 7)) > 0 Then
                        Grid2.Row = i: Grid2.RowSel = i
                        Grid2.Col = 0: Grid2.ColSel = 2
                        Grid2.CellBackColor = gcolor.BackColor
                        Grid2.Col = 6: Grid2.ColSel = Grid2.Cols - 1
                        Grid2.CellBackColor = gcolor.BackColor
                    Else
                        Grid2.Row = i: Grid2.RowSel = i
                        Grid2.Col = 0: Grid2.ColSel = 2
                        Grid2.CellBackColor = ycolor.BackColor
                        Grid2.Col = 6: Grid2.ColSel = Grid2.Cols - 1
                        Grid2.CellBackColor = ycolor.BackColor
                    End If
                End If
            End If
        Next i
        Grid2.Row = 1: Grid2.Col = 3
    End If
    Grid2.FormatString = "^SKU|<Product|^Stock|^Pallets|^Wraps|^Alt?|^UnDiff|^Paldiff|^Need|^RecID"
    Grid2.ColWidth(0) = 600
    Grid2.ColWidth(1) = 3200
    Grid2.ColWidth(2) = 1000
    Grid2.ColWidth(3) = 1000
    Grid2.ColWidth(4) = 1000
    Grid2.ColWidth(5) = 1000
    Grid2.ColWidth(6) = 1000
    Grid2.ColWidth(7) = 1000
    Grid2.ColWidth(8) = 1000
    Grid2.ColWidth(9) = 1000
    Grid2.Redraw = True
    tp = 0: tw = 0: ta = 0: ts = 0
    For i = 0 To Grid2.Rows - 1
        tp = tp + Val(Grid2.TextMatrix(i, 3))
        tw = tw + Val(Grid2.TextMatrix(i, 4))
        If Grid2.TextMatrix(i, 5) = "Y" Then ta = ta + 1
        If Grid2.TextMatrix(i, 3) = "1" And Grid2.TextMatrix(i, 9) = "2" Then ts = ts + 1
    Next i
    If Grid1.TextMatrix(Grid1.Row, 5) = tp Then
        tp.BackColor = wcolor.BackColor
    Else
        tp.BackColor = ycolor.BackColor
    End If
    If Val(ta) > 0 Then
        ta.BackColor = wcolor.BackColor
    Else
        ta.BackColor = ycolor.BackColor
    End If
End Sub

Private Sub brcode_Change()
    refresh_grid1
End Sub

Private Sub Command1_Click()
    post_order
    wdbphone.rtype = "branchorder"
    wdbphone.qstr = Val(wdbphone.qstr.Caption) + 1
    wdbphone.Show
End Sub

Private Sub Form_Load()
    'Set wdb = CreateObject("ADODB.Connection")
    'wdb.Open "ODBC;DATABASE=WDShip;DSN=wdship"
    ''wdb.Open "Driver={SQL Server};Server=bbc-01-wdsql;DATABASE=WDShip;UID=bbcship500;PWD=brenham500"
    'Call build_skumast
    'Call build_branches
    'Call build_plants
    Me.Left = Form1.Left
    Me.Top = Form1.Top + (Form1.wdbanner.Height * 1.7)
    Me.Height = Form1.WebBrowser1.Height
    refresh_tdates
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width * 0.66
    Grid2.Width = Me.Width - 180
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'wdb.Close
End Sub

Private Sub Grid1_RowColChange()
    rkey.Caption = Val(Grid1.TextMatrix(Grid1.Row, 0))
End Sub

Private Sub Grid2_KeyPress(KeyAscii As Integer)
    Dim i As Integer, msg As String, omt As Integer
    If Grid2.Col = 3 Or Grid2.Col = 4 Then
        pflag = True
        Command1.Visible = True
        If edcol = True Then
            Grid2.Text = ""
            edcol = False
        End If
        If KeyAscii = 8 Then
            If Len(Grid2.Text) > 1 Then
                Grid2.Text = Left(Grid2.Text, Len(Grid2.Text) - 1)
            Else
                Grid2.Text = ""
            End If
        End If
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            Grid2.Text = Grid2.Text & Chr(KeyAscii)
        End If
    End If
    
    If Grid2.Col = 5 Then
        pflag = True
        Command1.Visible = True
        If Grid2.Text >= "Y" Then
            Grid2.Text = ""
        Else
            Grid2.Text = "Y"
        End If
    End If
    tp = 0: tw = 0: ta = 0: ts = 0
    For i = 0 To Grid2.Rows - 1
        tp = tp + Val(Grid2.TextMatrix(i, 3))
        tw = tw + Val(Grid2.TextMatrix(i, 4))
        If Grid2.TextMatrix(i, 5) = "Y" Then ta = ta + 1
        If Grid2.TextMatrix(i, 3) = "1" And Grid2.TextMatrix(i, 9) = "2" Then ts = ts + 1
    Next i
    If Grid1.TextMatrix(Grid1.Row, 5) = tp Then
        tp.BackColor = wcolor.BackColor
    Else
        tp.BackColor = ycolor.BackColor
    End If
    If Val(ta) > 0 Then
        ta.BackColor = wcolor.BackColor
    Else
        ta.BackColor = ycolor.BackColor
    End If
End Sub

Private Sub rkey_Change()
    refresh_grid2
End Sub

