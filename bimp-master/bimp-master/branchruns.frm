VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form branchruns 
   Caption         =   "Active Trailers"
   ClientHeight    =   10305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   ScaleHeight     =   10305
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   1935
      Left            =   6120
      TabIndex        =   12
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
      _Version        =   327680
      Rows            =   6
      BackColorFixed  =   12648447
      FocusRect       =   0
      GridLines       =   2
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   960
      TabIndex        =   9
      Text            =   "Combo3"
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   10440
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   5055
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   8916
      _Version        =   327680
      Cols            =   8
      BackColorFixed  =   12640511
      FocusRect       =   0
      GridLines       =   2
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   9735
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   17171
      _Version        =   327680
      BackColorFixed  =   16777152
      ForeColorFixed  =   0
      BackColorSel    =   128
      FocusRect       =   0
      GridLines       =   2
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   6960
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   6720
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   120
      Width           =   1095
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
      Left            =   3240
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label gcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "gcolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10200
      TabIndex        =   17
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label bcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "bcolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10200
      TabIndex        =   16
      Top             =   9360
      Width           =   1335
   End
   Begin VB.Label ycolor 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ycolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10200
      TabIndex        =   15
      Top             =   9000
      Width           =   1335
   End
   Begin VB.Label wcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "wcolor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10200
      TabIndex        =   14
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label rcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "rcolor"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10200
      TabIndex        =   13
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
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
      Left            =   7920
      TabIndex        =   11
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Broken Arrow"
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
      Left            =   4440
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "SKU:"
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
      Left            =   6120
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Ship Date:"
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
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu edwhs 
         Caption         =   "Warehouse"
         Begin VB.Menu postwhs2trl 
            Caption         =   "Post Warehouse to Trailer"
         End
      End
      Begin VB.Menu edtrl 
         Caption         =   "Trailer"
         Begin VB.Menu dropsku 
            Caption         =   "Drop SKU"
         End
         Begin VB.Menu addsku 
            Caption         =   "Add SKU"
         End
         Begin VB.Menu splitsku 
            Caption         =   "Split SKU"
         End
         Begin VB.Menu printrl 
            Caption         =   "Print"
         End
      End
   End
End
Attribute VB_Name = "branchruns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim refflag As Boolean

Sub refresh_vlists()
    Dim ds As ADODB.Recordset, s As String
    Combo1.Clear: List1.Clear
    Combo2.Clear: List2.Clear
    Combo3.Clear
    For i = 50 To 52
        Combo1.AddItem plantrec(i).orawhs
        List1.AddItem plantrec(i).plantname
    Next i
    'Combo1.AddItem "VENDOR": List1.AddItem "Vendor Items"
    'Combo1.AddItem "DRY": List1.AddItem "Dry Storage Items"
    For i = 0 To 9999
        If skurec(i).pallet > 0 Then
            Combo2.AddItem skurec(i).sku
            List2.AddItem skurec(i).unit & " " & skurec(i).desc
        End If
    Next i
    s = "select trldate, count(*) from runs group by trldate order by trldate"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo3.AddItem Format(ds!trldate, "MM-dd-yyyy")
            ds.MoveNext
        Loop
    End If
    ds.Close
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    If Combo3.ListCount > 0 Then Combo3.ListIndex = 0
End Sub

Sub refresh_grid1()
    Dim ds As ADODB.Recordset, s As String, i As Integer, psku As String, bs As Recordset, plit As String
    Dim tp As Integer, tw As Integer, t1 As Integer, t2 As Integer, t3 As Integer
    Dim t4 As Integer, t5 As Integer
    Dim ap As Integer, aw As Integer, a1 As Integer, a2 As Integer, a3 As Integer
    Dim a4 As Integer, a5 As Integer, q As String
    psku = " "
    tp = 0: tw = 0: t1 = 0: t2 = 0: t3 = 0: t4 = 0: t5 = 0
    ap = 0: aw = 0: a1 = 0: a2 = 0: a3 = 0: a4 = 0: a5 = 0
    Screen.MousePointer = 11
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 14
    s = "select r.id, r.locname, r.trlno, r.trlsize, r.startime, t.id, t.branch, t.sku, t.pallets"
    s = s & ", t.wraps, t.whs_num, t.groupcode from runs r, trailers t"
    s = s & " where r.trldate = '" & Combo3 & "'"
    s = s & " and r.loaded = '" & Format(Combo1.ListIndex + 50, "0") & "'"
    s = s & " and t.runid = r.id"
    s = s & " order by t.sku, t.branch, r.trlno"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!sku <> psku Then
                If Val(psku) > 0 Then
                    plit = ""
                    s = "select * from bimp where sku = '" & psku & "' and plantwhs = '" & Combo1 & "'"
                    Set bs = wdb.Execute(s)
                    If bs.BOF = False Then
                        bs.MoveFirst
                        'MsgBox ((bs!plantpool / bs!roqty) - tp) & " " & bs!lowqty, vbOKOnly, psku
                        If ((bs!plantpool / bs!roqty) - tp) <= bs!lowqty Then plit = "Low stock qty " & bs!lowqty & " has been reached."
                        If ((bs!plantpool / bs!roqty) - tp) <= bs!outqty Then plit = "Out of stock qty " & bs!outqty & " has been reached."
                    End If
                    bs.Close
                    s = Chr(9) & plit & Chr(9) & Chr(9) & Chr(9)
                    s = s & Format(tp, "#") & Chr(9)
                    s = s & Format(tw, "#") & Chr(9)
                    s = s & Format(t1, "#") & Chr(9)
                    s = s & Format(t2, "#") & Chr(9)
                    s = s & Format(t3, "#") & Chr(9)
                    s = s & Format(t4, "#") & Chr(9)
                    s = s & Format(t5, "#")
                    Grid1.AddItem s
                    tp = 0: tw = 0: t1 = 0: t2 = 0: t3 = 0: t4 = 0: t5 = 0
                End If
                s = Chr(9) & ds!sku & " " & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc
                Grid1.AddItem s
                psku = ds!sku
            End If
            s = ds(5) & Chr(9)                                      'Trailer.id
            s = s & ds(1) & " " & ds(2) & Chr(9)                    'Trailer desc
            s = s & ds(0) & Chr(9)                                  'runid
            s = s & Format(ds(4), "h:mm am/pm") & Chr(9)            'startime
            s = s & Format(ds(8), "#") & Chr(9)                     'pallets
            s = s & Format(ds(9), "#") & Chr(9)                     'wraps
            's = s & ds(10) & Chr(9)                                 'whs
            tp = tp + ds(8): ap = ap + ds(8)
            tw = tw + ds(9): aw = aw + ds(9)
            If ds(10) = 1 Then
                s = s & Format(ds(8), "#") & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ds(3)
                t1 = t1 + ds(8): a1 = a1 + ds(8)
            Else
                If ds(10) = 2 Then
                    s = s & Chr(9) & Format(ds(8), "#") & Chr(9) & Chr(9) & Chr(9) & Chr(9) & ds(3)
                    t2 = t2 + ds(8): a2 = a2 + ds(8)
                Else
                    If ds(10) = 3 Then
                        s = s & Chr(9) & Chr(9) & Format(ds(8), "#") & Chr(9) & Chr(9) & Chr(9) & ds(3)
                        t3 = t3 + ds(8): a3 = a3 + ds(8)
                    Else
                        If ds(10) = 4 Or ds(10) = 14 Then
                            s = s & Chr(9) & Chr(9) & Chr(9) & Format(ds(8), "#") & Chr(9) & Chr(9) & ds(3)
                            t4 = t4 + ds(8): a4 = a4 + ds(8)
                        Else
                            If ds(10) = 5 Or ds(10) = 15 Then
                                s = s & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Format(ds(8), "#") & Chr(9) & ds(3)
                                t5 = t5 + ds(8): a5 = a5 + ds(8)
                            Else
                                s = s & Chr(9) & Chr(9) & Chr(9) & Format(ds(8), "#") & Chr(9) & Chr(9) & ds(3)
                                t4 = t4 + ds(8): t4 = t4 + ds(8)
                            End If
                        End If
                    End If
                End If
            End If
            s = s & Chr(9) & ds(11)
            q = "select bimpstatus from bimp where plantwhs = '" & Combo1 & "'"
            q = q & " and branchwhs = '" & Format(ds(6), "000") & "'"
            q = q & " and sku = '" & psku & "'"
            Set bs = wdb.Execute(q)
            If bs.BOF = False Then
                bs.MoveFirst
                s = s & Chr(9) & bs!bimpstatus
            End If
            bs.Close
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    If Val(psku) > 0 Then
        plit = ""
        s = "select * from bimp where sku = '" & psku & "' and plantwhs = '" & Combo1 & "'"
        Set bs = wdb.Execute(s)
        If bs.BOF = False Then
            bs.MoveFirst
            'MsgBox ((bs!plantpool / bs!roqty) - tp) & " " & bs!lowqty, vbOKOnly, psku
            If ((bs!plantpool / bs!roqty) - tp) <= bs!lowqty Then plit = "Low stock qty " & bs!lowqty & " has been reached."
            If ((bs!plantpool / bs!roqty) - tp) <= bs!outqty Then plit = "Out of stock qty " & bs!outqty & " has been reached."
        End If
        bs.Close
        s = Chr(9) & plit & Chr(9) & Chr(9) & Chr(9)
        s = s & Format(tp, "#") & Chr(9)
        s = s & Format(tw, "#") & Chr(9)
        s = s & Format(t1, "#") & Chr(9)
        s = s & Format(t2, "#") & Chr(9)
        s = s & Format(t3, "#") & Chr(9)
        s = s & Format(t4, "#") & Chr(9)
        s = s & Format(t5, "#")
        Grid1.AddItem s
        'tp = 0: tw = 0: t1 = 0: t2 = 0: t3 = 0: t4 = 0: t5 = 0
    End If
    's = Chr(9) & psku & " " & skurec(Val(psku)).unit & " " & skurec(Val(psku)).desc
    'Grid1.AddItem s
    
    
    
    If Grid1.Rows > 2 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 0)) = 0 Then
                Grid1.Row = i: Grid1.RowSel = i
                If Grid1.TextMatrix(i, 4) > " " Then
                    Grid1.Col = 4: Grid1.ColSel = Grid1.Cols - 1
                Else
                    Grid1.Col = 1: Grid1.ColSel = 3 'Grid1.Cols - 1
                End If
                Grid1.CellBackColor = Grid1.BackColorFixed
                Grid1.CellForeColor = Grid1.ForeColorFixed
            End If
            If Grid1.TextMatrix(i, 13) > "0" Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 4: Grid1.ColSel = 4
                If Grid1.TextMatrix(i, 13) = "W" Then Grid1.CellBackColor = wcolor.BackColor
                If Grid1.TextMatrix(i, 13) = "Y" Then Grid1.CellBackColor = ycolor.BackColor
                If Grid1.TextMatrix(i, 13) = "B" Then Grid1.CellBackColor = bcolor.BackColor
                If Grid1.TextMatrix(i, 13) = "G" Then Grid1.CellBackColor = gcolor.BackColor
                'Grid1.TextMatrix(i, 13) = " "
            End If
        Next i
        Grid1.Row = 2: Grid1.Col = 2
    End If
    Grid1.AddItem " "
    s = Chr(9) & "Summary" & Chr(9) & Chr(9) & Chr(9)
    s = s & Format(ap, "#") & Chr(9)
    s = s & Format(aw, "#") & Chr(9)
    s = s & Format(a1, "#") & Chr(9)
    s = s & Format(a2, "#") & Chr(9)
    s = s & Format(a3, "#") & Chr(9)
    s = s & Format(a4, "#") & Chr(9)
    s = s & Format(a5, "#")
    Grid1.AddItem s
    
    If Combo1 = "T10" Then Grid1.FormatString = "^|<|^Ticket|^Start|^Pallets|^Wraps|^SR1|^SR2|^SR3|^SR4|^SR5|^Size"
    If Combo1 = "K10" Then Grid1.FormatString = "^|<|^Ticket|^Start|^Pallets|^Wraps|^|^|^|^Racks|^|^Size"
    If Combo1 = "A10" Then Grid1.FormatString = "^|<|^Ticket|^Start|^Pallets|^Wraps|^|^|^|^Racks|^CS5|^Size"
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 3500
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 0
    Grid1.ColWidth(7) = 0
    Grid1.ColWidth(8) = 0
    If Combo1 = "T10" Then
        Grid1.ColWidth(6) = 1000
        Grid1.ColWidth(7) = 1000
        Grid1.ColWidth(8) = 1000
    End If
    Grid1.ColWidth(9) = 1000
    If Combo1 = "K10" Then
        Grid1.ColWidth(10) = 0
    Else
        Grid1.ColWidth(10) = 1000
    End If
    Grid1.ColWidth(11) = 1000
    Grid1.ColWidth(12) = 0 '1000
    Grid1.ColWidth(13) = 0 '1000
    'Grid1.MousePointer = flexUpArrow
    Grid1.Redraw = True
    Screen.MousePointer = 0
    Grid2.Visible = False: Grid3.Visible = False
End Sub

Sub refresh_grid2()
    Dim ds As ADODB.Recordset, s As String, i As Integer, bno As Integer
    Dim tp As Long, tw As Long, tu As Long, ta As Long
    tp = 0: tw = 0: tu = 0: ta = 0
    Screen.MousePointer = 11
    Grid2.Redraw = False
    Grid2.FontName = "Arial"
    Grid2.FontBold = True
    Grid2.FontSize = 8
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 11
    s = "select * from trailers where runid = " & Grid1.TextMatrix(Grid1.Row, 2)
    s = s & " order by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        bno = ds!branch
        s = Chr(9) & Grid1.TextMatrix(Grid1.Row, 2) & Chr(9)
        s = s & Grid1.TextMatrix(Grid1.Row, 1) & Chr(9)
        s = s & Grid1.TextMatrix(Grid1.Row, 11)
        Grid2.AddItem s
        Do Until ds.EOF
            s = ds!id & Chr(9) & Chr(9) & Chr(9) & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & Format(ds!pallets, "#") & Chr(9)
            s = s & Format(ds!wraps, "#") & Chr(9)
            s = s & Format(ds!units, "#") & Chr(9)
            s = s & ds!whs_num & Chr(9)
            'If ds!whs_num <= 5 Then
            '    s = s & ds!whs_num
            tp = tp + ds!pallets
            tw = tw + ds!wraps
            tu = tu + ds!units
            If ds!account = "ALT" Then
                s = s & "Y"
                ta = ta + 1
            End If
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid2.Rows > 1 Then
        refflag = False
        Grid2.FillStyle = flexFillRepeat
        For i = 1 To Grid2.Rows - 1
            If Val(Grid2.TextMatrix(i, 4)) > 0 Then
                s = "select bimpstatus from bimp where plantwhs = '" & Combo1 & "'"
                s = s & " and branchwhs = '" & Format(bno, "000") & "'"
                s = s & " and sku = '" & Grid2.TextMatrix(i, 4) & "'"
                'MsgBox s
                Set ds = wdb.Execute(s)
                If ds.BOF = False Then
                    ds.MoveFirst
                    Grid2.Row = i: Grid2.RowSel = i
                    Grid2.Col = 4: Grid2.ColSel = 8
                    If ds!bimpstatus = "W" Then Grid2.CellBackColor = wcolor.BackColor
                    If ds!bimpstatus = "Y" Then Grid2.CellBackColor = ycolor.BackColor
                    If ds!bimpstatus = "B" Then Grid2.CellBackColor = bcolor.BackColor
                    If ds!bimpstatus = "G" Then Grid2.CellBackColor = gcolor.BackColor
                End If
                ds.Close
            End If
        Next i
        refflag = True
        Grid2.Row = 1
    End If
    s = Chr(9) & Grid1.TextMatrix(Grid1.Row, 12) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9)
    s = s & Format(tp, "#") & Chr(9)
    s = s & Format(tw, "#") & Chr(9)
    s = s & Format(tu, "#") & Chr(9)
    s = s & Chr(9) & Format(ta, "#")
    Grid2.AddItem s
    If tp <> Val(Grid2.TextMatrix(1, 3)) Then
        Grid2.FillStyle = flexFillRepeat
        Grid2.Row = Grid2.Rows - 1: Grid2.RowSel = Grid2.Row
        Grid2.Col = 6: Grid2.ColSel = 8
        Grid2.CellBackColor = rcolor.BackColor
        Grid2.CellForeColor = rcolor.ForeColor
        Grid2.Row = 1
    End If
    Grid2.FormatString = "^ID|^RunID|^Destination|^Size|^SKU|<Product|^Pallets|^Wraps|^Units|^Source|^Alt"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 2300
    Grid2.ColWidth(3) = 600
    Grid2.ColWidth(4) = 600
    Grid2.ColWidth(5) = 3000
    Grid2.ColWidth(6) = 800
    Grid2.ColWidth(7) = 800
    Grid2.ColWidth(8) = 800
    Grid2.ColWidth(9) = 800
    Grid2.ColWidth(10) = 600
    Grid2.Redraw = True
    Grid2.Visible = True: Grid2.SetFocus
    Screen.MousePointer = 0
End Sub

Sub refresh_grid3()
    Dim db As ADODB.Connection, ds As ADODB.Recordset, psku As String, hs As ADODB.Recordset
    Dim i As Integer, k As Integer
    If plant_server_status(Combo1) = False Then                                         'jv010417
        s = "Sorry, The server for Warehouse " & prodbatches.Combo1 & " has been flagged to be offline."
        MsgBox s, vbOKOnly + vbInformation, "sorry, try again later..."                 'jv010417
        Exit Sub                                                                        'jv010417
    End If                                                                              'jv010417
    
    Screen.MousePointer = 11
    Grid3.Redraw = False
    Grid3.FontName = "Arial"
    Grid3.FontBold = True
    Grid3.FontSize = 8
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 7
    If Combo1 = "T10" Then
        s = "SR1" & Chr(9) & "1": Grid3.AddItem s
        s = "SR2" & Chr(9) & "2": Grid3.AddItem s
        s = "SR3" & Chr(9) & "3": Grid3.AddItem s
        s = "SR4" & Chr(9) & "4": Grid3.AddItem s
        s = "SR5" & Chr(9) & "5": Grid3.AddItem s
    End If
    If Combo1 = "K10" Then
        s = "Racks" & Chr(9) & "14": Grid3.AddItem s
    End If
    If Combo1 = "A10" Then
        s = "Racks" & Chr(9) & "15": Grid3.AddItem s
        s = "CS5" & Chr(9) & "15": Grid3.AddItem s
    End If
    
    'psku = Grid2.TextMatrix(Grid2.Row, 4)
    psku = Combo2
    'Pool Schedule Pallets
    s = "select sku, sum(palqty) from poolschedule where plantwhs = '" & Combo1 & "'"
    s = s & " and sku = '" & psku & "'"
    s = s & " group by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Grid3.TextMatrix(Grid3.Rows - 1, 4) = Val(Grid3.TextMatrix(Grid3.Rows - 1, 4)) + ds(1)
            ds.MoveNext
        Loop
    End If
    ds.Close
    DoEvents
    
    'MsgBox "racks"
    Set db = CreateObject("ADODB.Connection")
    If Combo1 = "A10" Then db.Open a10bbsr
    If Combo1 = "K10" Then db.Open k10bbsr
    If Combo1 = "T10" Then db.Open t10bbsr
    s = "select sku, count(*) from rackpos where sku = '" & psku & "'"
    's = s & " and rackno in (select id from racks where hold = 1)"
    s = s & " and rackno in (select id from racks where hold = 1 and rack <> 'OP')"         'jv120215
    s = s & " group by sku"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If Combo1 = "T10" Then Grid3.TextMatrix(4, 3) = Val(Grid3.TextMatrix(4, 3)) + ds(1)
            If Combo1 = "K10" Then Grid3.TextMatrix(1, 3) = Val(Grid3.TextMatrix(1, 3)) + ds(1)
            If Combo1 = "A10" Then Grid3.TextMatrix(1, 3) = Val(Grid3.TextMatrix(1, 3)) + ds(1)
            If Combo1 = "T10" Then Grid3.TextMatrix(4, 2) = Val(Grid3.TextMatrix(4, 2)) + ds(1)
            If Combo1 = "K10" Then Grid3.TextMatrix(1, 2) = Val(Grid3.TextMatrix(1, 2)) + ds(1)
            If Combo1 = "A10" Then Grid3.TextMatrix(1, 2) = Val(Grid3.TextMatrix(1, 2)) + ds(1)
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select sku, count(*) from rackpos where sku = '" & psku & "'"
    's = s & " and rackno in (select id from racks where hold = 0)"
    s = s & " and rackno in (select id from racks where hold = 0 and rack <> 'OP')"         'jv120215
    s = s & " group by sku"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If Combo1 = "T10" Then Grid3.TextMatrix(4, 2) = Val(Grid3.TextMatrix(4, 2)) + ds(1)
            If Combo1 = "K10" Then Grid3.TextMatrix(1, 2) = Val(Grid3.TextMatrix(1, 2)) + ds(1)
            If Combo1 = "A10" Then Grid3.TextMatrix(1, 2) = Val(Grid3.TextMatrix(1, 2)) + ds(1)
            ds.MoveNext
        Loop
    End If
    ds.Close
        
    's = "select sku, lot_num, barcode from rackpos where sku = '" & psku & " '"
    'Set ds = db.Execute(s)
    'If ds.BOF = False Then
    '    ds.MoveFirst
    '    Do Until ds.EOF
    '        pcode = Trim(Mid(ds!barcode, 11, 3))
    '        palno = Right(ds!barcode, 3)
    '        s = "select id from holdlist where sku = '" & ds!sku & "'"
    '        s = s & " and lot_num = '" & ds!lot_num & "'"
    '        s = s & " and opcode = '" & pcode & "'"
    '        s = s & " and spallet >= '" & palno & "'"
    '        s = s & " and epallet <= '" & palno & "'"
    '        Set hs = db.Execute(s)
    '        If hs.BOF = False Then
    '            If Combo1 = "T10" Then Grid3.TextMatrix(4, 2) = Val(Grid3.TextMatrix(4, 3)) + 1
    '            If Combo1 = "K10" Then Grid3.TextMatrix(1, 3) = Val(Grid3.TextMatrix(1, 3)) + 1
    '            If Combo1 = "A10" Then Grid3.TextMatrix(1, 3) = Val(Grid3.TextMatrix(1, 3)) + 1
    '            If Combo1 = "T10" Then Grid3.TextMatrix(4, 2) = Val(Grid3.TextMatrix(4, 2)) + 1
    '            If Combo1 = "K10" Then Grid3.TextMatrix(1, 2) = Val(Grid3.TextMatrix(1, 2)) + 1
    '            If Combo1 = "A10" Then Grid3.TextMatrix(1, 2) = Val(Grid3.TextMatrix(1, 2)) + 1
    '        Else
    '            If Combo1 = "T10" Then Grid3.TextMatrix(4, 2) = Val(Grid3.TextMatrix(4, 2)) + 1
    '            If Combo1 = "K10" Then Grid3.TextMatrix(1, 2) = Val(Grid3.TextMatrix(1, 2)) + 1
    '            If Combo1 = "A10" Then Grid3.TextMatrix(1, 2) = Val(Grid3.TextMatrix(1, 2)) + 1
    '        End If
    '        hs.Close
    '        DoEvents
    '        ds.MoveNext
    '        'If MsgBox("Exit", vbYesNo, "quit") = vbYes Then Exit Do
    '    Loop
    'End If
    'ds.Close
    
    If Combo1 = "T10" Then
        's = "select whse_num, pallet_status, count(*) from position where sku = '" & psku & "'"
        's = s & " group by whse_num, pallet_status"
        s = "select whse_num, lane_status, sum(qty) from lane where sku = '" & psku & "'"
        s = s & " group by whse_num, lane_status"
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                'If ds!pallet_status = "H" Then
                If ds!lane_status = "H" Then
                    Grid3.TextMatrix(ds!whse_num, 3) = Val(Grid3.TextMatrix(ds!whse_num, 3)) + ds(2)
                    Grid3.TextMatrix(ds!whse_num, 2) = Val(Grid3.TextMatrix(ds!whse_num, 2)) + ds(2)
                Else
                    Grid3.TextMatrix(ds!whse_num, 2) = Val(Grid3.TextMatrix(ds!whse_num, 2)) + ds(2)
                End If
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    db.Close
    If Combo1 = "A10" Then
        'MsgBox "cs5"
        db.Open cs5db
        's = "Select * from vContainerLocation_1033 where item >= '" & psku & "'"
        s = "Select * from vAllInventory_1033 where item >= '" & psku & "'"        'Westfalia Upgrade
        s = s & " and item < '" & psku & "Z'"
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                'If Trim(ds(16)) = "True" Then
                If ds(17) > 0 Then                              'Westfalia Upgrade
                    Grid3.TextMatrix(2, 3) = Val(Grid3.TextMatrix(2, 3)) + 1
                    Grid3.TextMatrix(2, 2) = Val(Grid3.TextMatrix(2, 2)) + 1
                Else
                    Grid3.TextMatrix(2, 2) = Val(Grid3.TextMatrix(2, 2)) + 1
                End If
                'DoEvents
                ds.MoveNext
            Loop
        End If
        ds.Close: db.Close
    End If
    s = "select whs_num, sum(pallets) from trailers where sku = '" & psku & "'"
    s = s & " and plant = " & Format(Combo1.ListIndex + 50, "0")
    s = s & " and shipdate >= '" & Format(Now, "MM-dd-yyyy") & "'"
    s = s & " and pb_flag = 'N'"
    s = s & " group by whs_num"
    'MsgBox s
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!whs_num <= 5 Then
                Grid3.TextMatrix(ds!whs_num, 5) = Val(Grid3.TextMatrix(ds!whs_num, 5)) + ds(1)
            Else
                If ds!whs_num = 14 Then
                    Grid3.TextMatrix(1, 5) = Val(Grid3.TextMatrix(1, 5)) + ds(1)
                Else
                    If ds!whs_num = 15 Then
                        Grid3.TextMatrix(2, 5) = Val(Grid3.TextMatrix(2, 5)) + ds(1)
                    Else
                        Grid3.TextMatrix(4, 5) = Val(Grid3.TextMatrix(4, 5)) + ds(1)
                    End If
                End If
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid3.Rows > 1 Then
        For i = 1 To Grid3.Rows - 1
            k = Val(Grid3.TextMatrix(i, 2))
            k = k - Val(Grid3.TextMatrix(i, 3))
            k = k + Val(Grid3.TextMatrix(i, 4))
            k = k - Val(Grid3.TextMatrix(i, 5))
            Grid3.TextMatrix(i, 6) = Format(k, "#")
        Next i
    End If
    Grid3.FormatString = "^Whs|^Code|^Onhand|^Hold|^New Pool|^OnOrder|^Net"
    Grid3.ColWidth(0) = 900
    Grid3.ColWidth(1) = 900
    Grid3.ColWidth(2) = 900
    Grid3.ColWidth(3) = 900
    Grid3.ColWidth(4) = 900
    Grid3.ColWidth(5) = 900
    Grid3.ColWidth(6) = 900
    Grid3.Redraw = True
    Grid3.Visible = True
    Screen.MousePointer = 0
End Sub

Private Sub addsku_Click()
    Dim s As String, r As String, i As Integer, bno As Integer, psku As String, gc As String
    r = Grid2.TextMatrix(1, 1): bno = 0
    gc = Grid2.TextMatrix(Grid2.Rows - 1, 1)
    psku = InputBox("SKU:", "Add SKU to trailer...", "777")
    If Len(psku) = 0 Then Exit Sub
    If skurec(Val(psku)).sku = psku Then
        s = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno"
        s = s & " ,sku, pallets, wraps, units, whs_num, pb_flag, ra_flag) values ("
        s = s & wd_seq("Trailers")
        s = s & ", " & r
        s = s & ", '" & gc & "'"
        's = s & ", '" & UCase(Left(Format(Combo3, "ddd"), 2)) & "-" & Right(Grid2.TextMatrix(1, 2), 1)
        For i = 1 To 99
            If branchrec(i).branchname = Left(Grid2.TextMatrix(1, 2), Len(Grid2.TextMatrix(1, 2)) - 3) Then
                bno = i
                's = s & Format(bno, "00")
                Exit For
            End If
        Next i
        's = s & "'"
        If Combo1 = "T10" Then s = s & ", 50"
        If Combo1 = "K10" Then s = s & ", 51"
        If Combo1 = "A10" Then s = s & ", 52"
        s = s & ", " & bno
        s = s & ", '.....'"
        s = s & ", '" & Combo3 & "'"
        s = s & ", '" & Right(Grid2.TextMatrix(1, 2), 2) & "'"
        s = s & ", '" & psku & "'"
        s = s & ", 1"
        s = s & ", 0"
        s = s & ", " & skurec(Val(psku)).pallet
        'If Combo1 = "T10" Then s = s & ", 5"
        'If Combo1 = "K10" Then s = s & ", 14"
        'If Combo1 = "A10" Then s = s & ", 15"
        s = s & ", " & Grid3.TextMatrix(1, 1)
        s = s & ", 'N', 'N')"
        'MsgBox s
        wdb.Execute s
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 2) = r Then
                Grid1.Row = i
                refresh_grid2
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    Label4.Caption = List1
    refresh_grid1
End Sub

Private Sub Combo2_Click()
    Dim i As Integer
    List2.ListIndex = Combo2.ListIndex
    Label5.Caption = List2
    For i = 0 To Grid1.Rows - 1
        If Trim(Left(Grid1.TextMatrix(i, 1), 4)) = Combo2 Then
            Grid1.TopRow = i
            Grid1.Row = i + 1
            Exit For
        End If
    Next i
    refresh_grid3
End Sub

Private Sub Combo3_Click()
    refresh_grid1
End Sub

Private Sub dropsku_Click()
    Dim s As String, r As String, i As Integer
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) > 0 Then
        r = Grid2.TextMatrix(1, 1)
        s = "delete from trailers where id = " & Grid2.TextMatrix(Grid2.Row, 0)
        'MsgBox s
        wdb.Execute s
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 2) = r Then
                Grid1.Row = i
                refresh_grid2
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    refresh_vlists
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 120
    'Grid2.Width = Me.Width - 120
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 3)
End Sub

Private Sub Grid1_DblClick()
    If Val(Grid1.TextMatrix(Grid1.Row, 2)) > 0 Then refresh_grid2
End Sub

Private Sub Grid1_GotFocus()
    Grid2.Visible = False
    Grid3.Visible = False
End Sub

Private Sub Grid2_KeyPress(KeyAscii As Integer)
    Dim i As Integer, msg As String, omt As Integer, x As Integer, y As Integer
    Dim psku As String, ppal As Long, pwrp As Long, punits As Long
    Dim tp As Long, tw As Long, tu As Long, ta As Long
    If Grid2.Col = 6 Or Grid2.Col = 7 Or Grid2.Col = 8 Then
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
        If Grid2.Col = 6 Or Grid2.Col = 7 Then
            psku = Grid2.TextMatrix(Grid2.Row, 4)
            ppal = Val(Grid2.TextMatrix(Grid2.Row, 6))
            pwrp = Val(Grid2.TextMatrix(Grid2.Row, 7))
            punits = ppal * skurec(Val(psku)).pallet
            punits = punits + (pwrp * skurec(Val(psku)).wrapunits)
            Grid2.TextMatrix(Grid2.Row, 8) = punits
        End If
        s = "Update trailers set pallets = " & Val(Grid2.TextMatrix(Grid2.Row, 6))
        s = s & ", wraps = " & Val(Grid2.TextMatrix(Grid2.Row, 7))
        s = s & ", units = " & Val(Grid2.TextMatrix(Grid2.Row, 8))
        s = s & " where id = " & Grid2.TextMatrix(Grid2.Row, 0)
        'MsgBox s
        wdb.Execute s
        tp = 0: tw = 0: tu = 0
        For i = 1 To Grid2.Rows - 1
            If Val(Grid2.TextMatrix(i, 4)) > 0 Then
                tp = tp + Val(Grid2.TextMatrix(i, 6))
                tw = tw + Val(Grid2.TextMatrix(i, 7))
                tu = tu + Val(Grid2.TextMatrix(i, 8))
            End If
        Next i
        Grid2.TextMatrix(Grid2.Rows - 1, 6) = Format(tp, "#")
        Grid2.TextMatrix(Grid2.Rows - 1, 7) = Format(tw, "#")
        Grid2.TextMatrix(Grid2.Rows - 1, 8) = tu
        If tp <> Grid2.TextMatrix(1, 3) Then
            x = Grid2.Col: y = Grid2.Row
            Grid2.FillStyle = flexFillRepeat
            Grid2.Row = Grid2.Rows - 1: Grid2.RowSel = Grid2.Row
            Grid2.Col = 6: Grid2.ColSel = 8
            Grid2.CellBackColor = rcolor.BackColor
            Grid2.CellForeColor = rcolor.ForeColor
            Grid2.Col = x: Grid2.Row = y
        End If
    End If
    
    If Grid2.Col = 10 Then
        If Grid2.Text >= "Y" Then
            Grid2.Text = ""
            's = "update trailers set altflag = 'N' where id = " & Grid2.TextMatrix(Grid2.Row, 0)
            s = "update trailers set account = '......' where id = " & Grid2.TextMatrix(Grid2.Row, 0)
        Else
            Grid2.Text = "Y"
            's = "update trailers set altflag = 'Y' where id = " & Grid2.TextMatrix(Grid2.Row, 0)
            s = "update trailers set account = 'ALT' where id = " & Grid2.TextMatrix(Grid2.Row, 0)
        End If
        'MsgBox s
        ta = 0
        For i = 1 To Grid2.Rows - 1
            If Grid2.TextMatrix(i, 10) = "Y" Then ta = ta + 1
        Next i
        Grid2.TextMatrix(Grid2.Rows - 1, 10) = ta
    End If
End Sub

Private Sub Grid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edtrl
End Sub

Private Sub Grid2_RowColChange()
    Dim i As Integer
    If refflag = False Then Exit Sub
    If Val(Grid2.TextMatrix(Grid2.Row, 4)) > 0 Then
        For i = 0 To Combo2.ListCount - 1
            If Combo2.List(i) = Grid2.TextMatrix(Grid2.Row, 4) Then
                Combo2.ListIndex = i
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Grid3_DblClick()
    Dim s As String
    If Combo2 = Grid2.TextMatrix(Grid2.Row, 4) And Val(Grid2.TextMatrix(Grid2.Row, 0)) > 0 Then
        s = "Update trailers set whs_num = " & Grid3.TextMatrix(Grid3.Row, 1)
        s = s & " where id = " & Grid2.TextMatrix(Grid2.Row, 0)
        'MsgBox s
        wdb.Execute s
        Grid2.TextMatrix(Grid2.Row, 9) = Grid3.TextMatrix(Grid3.Row, 1)
    End If
End Sub

Private Sub Grid3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edwhs
End Sub

Private Sub postwhs2trl_Click()
    Grid3_DblClick
End Sub

Private Sub printrl_Click()
    Dim rt As String, rf As String, rh As String, w As Long
    Dim c0 As Long, c1 As Long, c2 As Long, c3 As Long
    rt = "Ship Date: " & Me.Combo3 & "   Group: " & Grid2.TextMatrix(Grid2.Rows - 1, 1)
    rh = Grid2.TextMatrix(1, 2) & "  Ticket: " & Grid2.TextMatrix(1, 1) & "  Size: " & Grid2.TextMatrix(1, 3)
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    
    htdc(0) = "cyan": gndc(0) = Me.Grid2.BackColorFixed
    htdc(1) = "yellow": gndc(1) = Me.ycolor.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid1.BackColor
    c0 = Grid2.ColWidth(0): Grid2.ColWidth(0) = 0
    c1 = Grid2.ColWidth(1): Grid2.ColWidth(1) = 0
    c2 = Grid2.ColWidth(2): Grid2.ColWidth(2) = 0
    c3 = Grid2.ColWidth(3): Grid2.ColWidth(3) = 0
    If MsgBox("Send to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid2, rt, rh, rf)
        Grid2.ColWidth(0) = c0
        Grid2.ColWidth(1) = c1
        Grid2.ColWidth(2) = c2
        Grid2.ColWidth(3) = c3
        Exit Sub
    End If
    w = Grid2.Width
    Grid2.Width = Me.Width
    Grid2.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid2, rt, rh, rf, "linen", "khaki", "white")
        Grid2.ColWidth(0) = c0
        Grid2.ColWidth(1) = c1
        Grid2.ColWidth(2) = c2
        Grid2.ColWidth(3) = c3
        Grid2.Width = w
        Grid2.Redraw = True
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(Me, "c:\htmlgrid.htm", Grid2, rt, rh, rf, "linen", "khaki", "white")
        Grid2.ColWidth(0) = c0
        Grid2.ColWidth(1) = c1
        Grid2.ColWidth(2) = c2
        Grid2.ColWidth(3) = c3
        Grid2.Width = w
        Grid2.Redraw = True
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub splitsku_Click()
    Dim s As String, r As String, i As Integer, bno As Integer, psku As String, gc As String
    If Val(Grid2.TextMatrix(Grid2.Row, 4)) = 0 Then Exit Sub
    r = Grid2.TextMatrix(1, 1): bno = 0
    gc = Grid2.TextMatrix(Grid2.Rows - 1, 1)
    s = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno"
    s = s & " ,sku, pallets, wraps, units, whs_num, pb_flag, ra_flag) values ("
    s = s & wd_seq("Trailers")
    s = s & ", " & r
    s = s & ", '" & gc & "'"
    's = s & ", '" & UCase(Left(Format(Combo3, "ddd"), 2)) & "-" & Right(Grid2.TextMatrix(1, 2), 1)
    For i = 1 To 99
        If branchrec(i).branchname = Left(Grid2.TextMatrix(1, 2), Len(Grid2.TextMatrix(1, 2)) - 3) Then
            bno = i
            's = s & Format(bno, "00")
            Exit For
        End If
    Next i
    's = s & "'"
    If Combo1 = "T10" Then s = s & ", 50"
    If Combo1 = "K10" Then s = s & ", 51"
    If Combo1 = "A10" Then s = s & ", 52"
    s = s & ", " & bno
    s = s & ", '.....'"
    s = s & ", '" & Combo3 & "'"
    s = s & ", '" & Right(Grid2.TextMatrix(1, 2), 2) & "'"
    s = s & ", '" & Combo2 & "'"
    s = s & ", 1"
    s = s & ", 0"
    s = s & ", " & skurec(Val(Combo2)).pallet
    s = s & ", " & Grid3.TextMatrix(1, 1)
    s = s & ", 'N', 'N')"
    'MsgBox s
    wdb.Execute s
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 2) = r Then
            Grid1.Row = i
            refresh_grid2
            Exit For
        End If
    Next i
End Sub
