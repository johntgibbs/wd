VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form branchouts 
   Caption         =   "Branch Out of Stock Items"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
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
      Left            =   12840
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update Browser"
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
      Left            =   10920
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   10560
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
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
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   3015
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
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   15901
      _Version        =   327680
      ForeColor       =   8388736
   End
   Begin VB.Label brzlit 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   7080
      TabIndex        =   6
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Warehouse:"
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
      Width           =   1575
   End
End
Attribute VB_Name = "branchouts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_lists()
    Dim i As Integer
    For i = 1 To 99
        If branchrec(i).oraloc > "000" And branchrec(i).oraloc < "999" Then
            Combo1.AddItem branchrec(i).oraloc & "-" & branchrec(i).branchname
            List1.AddItem branchrec(i).oraloc
        End If
    Next i
    Combo1.AddItem "A10-Sylacauga Plant": List1.AddItem "A10"
    Combo1.AddItem "K10-Broken Arrow Plant": List1.AddItem "K10"
    Combo1.AddItem "T10-Brenham Plant": List1.AddItem "T10"
    Combo1.ListIndex = 0
End Sub

Private Sub refresh_all()
    Dim ds As ADODB.Recordset, s As String, i As Integer, k As Long, pflag As Boolean
    Grid1.Redraw = False: Grid1.Font = "Arial": Grid1.FontBold = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 11
    s = "Select id,plantwhs,branchwhs,sku,roqty,lastrecpt,skunotes"
    s = s & " from bimp where onhand = 0 and discflag = 'N' and sales > 0 and onorder = 0"
    s = s & " and branchwhs not in ('012', '015', '016')"
    s = s & " and plantwhs in ('A10', 'K10', 'T10')"
    s = s & " order by sku,branchwhs"
    
    s = "select b.id,b.plantwhs,b.branchwhs,b.sku,b.roqty,b.lastrecpt,b.skunotes,s.numwrap"
    s = s & " from bimp b, skumast s"
    s = s & " where b.discflag = 'N'"
    s = s & " and b.sales > 0"
    s = s & " and b.onorder = 0"
    s = s & " and b.branchwhs not in ('012', '015', '016')"
    s = s & " and b.plantwhs in ('A10', 'K10', 'T10')"
    's = s & " and b.plantwhs = 'A10'"
    s = s & " and s.sku = b.sku"
    s = s & " and b.onhand <= s.numwrap" ' * 2"
    s = s & " order by b.sku,b.branchwhs"
    MsgBox s
    
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!roqty & Chr(9)
            's = s & ds!branchwhs & "-" & branchrec(Val(ds!branchwhs)).branchname & Chr(9)
            s = s & ds!branchwhs & Chr(9)
            s = s & ds!lastrecpt & Chr(9)
            s = s & Trim(ds!skunotes) & Chr(9)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "^ID|^SKU|<Product|^PalSize|<Branch|^Last Receipt|^Last Issue|^Sales|^Days Instock|^DaysOut|^LostSales"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 2400
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 1200
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1200
    Grid1.ColWidth(9) = 1200
    Grid1.ColWidth(10) = 1200
    Grid1.Redraw = True
End Sub

Private Sub refresh_grid()
    Dim s As String, daysin As Integer, daysout As Integer, lostsales As Long, salesperday As Long
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Grid1.Redraw = False: Grid1.Font = "Arial": Grid1.FontBold = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    If List1 = "A10" Or List1 = "K10" Or List1 = "T10" Then
        pflag = True
    Else
        pflag = False
    End If
    Open brzloss For Input As #2
    Do Until EOF(2)
        Input #2, f0, f1, f2, f3, f4, f5, f6, f7
        If pflag = True And f3 = List1 Then
            s = f0 & Chr(9)
            s = s & f1 & Chr(9)
            s = s & skurec(Val(f1)).unit & " " & skurec(Val(f1)).desc & Chr(9)
            s = s & f2 & Chr(9)
            s = s & f4 & "-" & branchrec(Val(f4)).branchname & Chr(9)
            s = s & f5 & Chr(9)
            s = s & f6 & Chr(9)
            s = s & f7 & Chr(9)
            daysin = DateDiff("d", f5, f6) + 1
            s = s & daysin & Chr(9)
            salesperday = f7 / daysin
            s = s & salesperday & Chr(9)
            daysout = DateDiff("d", f6, Now)
            s = s & daysout & Chr(9)
            'lostsales = f7 * (daysout / daysin)
            lostsales = salesperday * daysout
            s = s & lostsales
            If lostsales > 0 Then Grid1.AddItem s
        Else
            If f4 = List1 Then
                s = f0 & Chr(9)
                s = s & f1 & Chr(9)
                s = s & skurec(Val(f1)).unit & " " & skurec(Val(f1)).desc & Chr(9)
                s = s & f2 & Chr(9)
                s = s & f3 & "-"
                If f3 = "A10" Then s = s & "Sylacauga"
                If f3 = "K10" Then s = s & "Broken Arrow"
                If f3 = "T10" Then s = s & "Brenham"
                s = s & Chr(9)
                s = s & f5 & Chr(9)
                s = s & f6 & Chr(9)
                s = s & f7 & Chr(9)
                daysin = DateDiff("d", f5, f6) + 1
                s = s & daysin & Chr(9)
                salesperday = f7 / daysin
                s = s & salesperday & Chr(9)
                daysout = DateDiff("d", f6, Now)
                s = s & daysout & Chr(9)
                'lostsales = f7 * (daysout / daysin)
                lostsales = salesperday * daysout
                s = s & lostsales
                If lostsales > 0 Then Grid1.AddItem s
            End If
        End If
    Loop
    Close #2
    If pflag = True Then
        s = "^ID|^SKU|<Product|^PalSize|<Branch|^Last Receipt|^Last Issue|^Sales|^Days Instock|^Sales/Day|^DaysOut|^LostSales"
    Else
        s = "^ID|^SKU|<Product|^PalSize|^Supplier|^Last Receipt|^Last Issue|^Sales|^Days Instock|^Sales/Day|^DaysOut|^LostSales"
    End If
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 2000
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 1200
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1200
    Grid1.ColWidth(9) = 1200
    Grid1.ColWidth(10) = 1200
    Grid1.ColWidth(11) = 1200
    Grid1.Redraw = True
End Sub


Private Sub refresh_grid_sql()
    Dim ds As ADODB.Recordset, s As String, i As Integer, k As Long, pflag As Boolean
    Grid1.Redraw = False: Grid1.Font = "Arial": Grid1.FontBold = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 11
    If List1 = "A10" Or List1 = "K10" Or List1 = "T10" Then
        pflag = True
    Else
        pflag = False
    End If
    s = "Select id,plantwhs,branchwhs,sku,roqty,lastrecpt,skunotes"
    s = s & " from bimp where onhand = 0 and discflag = 'N' and sales > 0 and onorder = 0"
    If pflag = True Then
        s = s & " and plantwhs = '" & List1 & "'"
    Else
        s = s & " and branchwhs = '" & List1 & "'"
        s = s & " and plantwhs = '" & branchrec(Val(List1)).supplier & "'"
    End If
    s = s & " and branchwhs not in ('012', '015', '016')"
    s = s & " order by sku,branchwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            s = s & ds!roqty & Chr(9)
            If pflag = True Then
                s = s & ds!branchwhs & "-" & branchrec(Val(ds!branchwhs)).branchname & Chr(9)
            Else
                s = s & ds!plantwhs & "-"
                If ds!plantwhs = "A10" Then s = s & "Sylacauga"
                If ds!plantwhs = "K10" Then s = s & "Broken Arrow"
                If ds!plantwhs = "T10" Then s = s & "Brenham"
                s = s & Chr(9)
            End If
            s = s & ds!lastrecpt & Chr(9)
            s = s & Trim(ds!skunotes) & Chr(9)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            'Grid1.TextMatrix(i, 8) = DateDiff("d", Grid1.TextMatrix(i, 5), Grid1.TextMatrix(i, 6))
            'Grid1.TextMatrix(i, 9) = DateDiff("d", Grid1.TextMatrix(i, 6), Now)
        Next i
    End If
    If pflag = True Then
        s = "^ID|^SKU|<Product|^PalSize|<Branch|^Last Receipt|^Last Issue|^Sales|^Days Instock|^DaysOut|^LostSales"
    Else
        s = "^ID|^SKU|<Product|^PalSize|<Supplier|^Last Receipt|^Last Issue|^Sales|^Days Instock|^DaysOut|^LostSales"
    End If
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 2400
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 1200
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 1200
    Grid1.ColWidth(9) = 1200
    Grid1.ColWidth(10) = 1200
    Grid1.Redraw = True
End Sub

Private Sub refresh_grid1()
    Dim ds As ADODB.Recordset, s As String, i As Integer, k As Long, pflag As Boolean
    Grid1.Redraw = False: Grid1.Font = "Arial": Grid1.FontBold = True
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 12
    If List1 = "A10" Or List1 = "K10" Or List1 = "T10" Then
        pflag = True
    Else
        pflag = False
    End If
    s = "Select plantwhs,branchwhs,sku,onorder,sales,plantpool,poolsched,lastrecpt,skunotes"
    s = s & " from bimp where onhand = 0 and discflag = 'N' and sales > 0 and onorder = 0"
    If pflag = True Then
        s = s & " and plantwhs = '" & List1 & "'"
    Else
        s = s & " and branchwhs = '" & List1 & "'"
        s = s & " and plantwhs = '" & branchrec(Val(List1)).supplier & "'"
    End If
    s = s & " and branchwhs not in ('012', '015', '016')"
    s = s & " order by sku,branchwhs"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            If pflag = True Then
                s = s & ds!branchwhs & "-" & branchrec(Val(ds!branchwhs)).branchname & Chr(9)
            Else
                s = s & ds!plantwhs & "-"
                If ds!plantwhs = "A10" Then s = s & "Sylacauga"
                If ds!plantwhs = "K10" Then s = s & "Broken Arrow"
                If ds!plantwhs = "T10" Then s = s & "Brenham"
                s = s & Chr(9)
            End If
            s = s & ds!onorder & Chr(9)
            s = s & ds!sales & Chr(9)
            s = s & ds!plantpool & Chr(9)
            s = s & ds!poolsched & Chr(9)
            s = s & ds!lastrecpt & Chr(9)
            s = s & Trim(ds!skunotes) & Chr(9)
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            Grid1.TextMatrix(i, 9) = DateDiff("d", Grid1.TextMatrix(i, 7), Grid1.TextMatrix(i, 8))
            k = Val(Grid1.TextMatrix(i, 9))
            If k >= 30 Then
                Grid1.TextMatrix(i, 10) = DateDiff("d", Grid1.TextMatrix(i, 8), Now)
                k = (Val(Grid1.TextMatrix(i, 4)) * Val(Grid1.TextMatrix(i, 10))) / 30
                Grid1.TextMatrix(i, 11) = k

            Else
                Grid1.TextMatrix(i, 10) = 30 - k
                If Val(Grid1.TextMatrix(i, 9)) > 0 Then
                    k = (Val(Grid1.TextMatrix(i, 4)) / Val(Grid1.TextMatrix(i, 9))) * Val(Grid1.TextMatrix(i, 10))
                Else
                    k = Val(Grid1.TextMatrix(i, 4))
                End If
                Grid1.TextMatrix(i, 11) = k
            End If
            
            'k = (Val(Grid1.TextMatrix(i, 4)) * Val(Grid1.TextMatrix(i, 10))) / 30
            'Grid1.TextMatrix(i, 11) = k
        Next i
    End If
    If pflag = True Then
        s = "^SKU|<Product|<Branch|^OnOrder|^Sales|^PlantPool|^Scheduled|^Last Receipt|^Last Issue|^Days Instock|^DaysOut|^LostSales"
    Else
        s = "^SKU|<Product|<Supplier|^OnOrder|^Sales|^PlantPool|^Scheduled|^Last Receipt|^Last Issue|^Days Instock|^DaysOut|^LostSales"
    End If
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 800
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 2400
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1200
    Grid1.ColWidth(8) = 1200
    Grid1.ColWidth(9) = 1200
    Grid1.ColWidth(10) = 1200
    Grid1.ColWidth(11) = 1200
    Grid1.Redraw = True
End Sub


Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Command1_Click()
    refresh_grid
End Sub

Private Sub Command2_Click()
    Dim psku As String, psize As Integer, i As Integer, sdate As String, edate As String
    Dim daysin As Integer, daysout As Integer, sales As Long, lostsales As Long, pwhs As String
    Dim pflag As Boolean
    refresh_all
    'Exit Sub
    Screen.MousePointer = 11
    'If List1 = "A10" Or List1 = "K10" Or List1 = "T10" Then
    '    pflag = True
    'Else
    '    pflag = False
    'End If
    For i = 1 To Grid1.Rows - 1
        'If pflag = True Then
            pwhs = Mid(Grid1.TextMatrix(i, 4), 1, 3)
        'Else
        '    pwhs = List1
        'End If
        psku = Grid1.TextMatrix(i, 1)
        psize = Val(Grid1.TextMatrix(i, 3))
        If pwhs = "036" Then
            sdate = last_branch_receipt(psku, pwhs, psize)
        Else
            sdate = last_branch_receipt(psku, pwhs, 0)
        End If
        edate = last_branch_issue(psku, pwhs)
        sales = last_branch_loads(psku, pwhs, sdate, edate)
        daysin = DateDiff("d", sdate, edate) + 1
        daysout = DateDiff("d", edate, Now)
        If daysin <= 0 Then daysin = 1
        lostsales = sales * (daysout / daysin)
        Grid1.TextMatrix(i, 5) = sdate
        Grid1.TextMatrix(i, 6) = edate
        Grid1.TextMatrix(i, 7) = sales
        Grid1.TextMatrix(i, 8) = daysin
        Grid1.TextMatrix(i, 9) = daysout
        Grid1.TextMatrix(i, 10) = lostsales
        DoEvents
    Next i
    Open brzloss For Output As #1
    For i = 1 To Grid1.Rows - 1
        Write #1, Grid1.TextMatrix(i, 0);
        Write #1, Grid1.TextMatrix(i, 1);
        Write #1, Grid1.TextMatrix(i, 3);
        Write #1, branchrec(Val(Grid1.TextMatrix(i, 4))).supplier;
        Write #1, Grid1.TextMatrix(i, 4);
        Write #1, Grid1.TextMatrix(i, 5);
        Write #1, Grid1.TextMatrix(i, 6);
        Write #1, Grid1.TextMatrix(i, 7)
    Next i
    Close #1
    Screen.MousePointer = 0
    brzlit.Caption = "Last updated:  " & Format(FileDateTime(brzloss), "M-d-yyyy h:mm am/pm")
    refresh_grid
End Sub

Private Sub Command3_Click()
    Dim rt As String, rh As String, rf As String
    rt = Me.Caption
    rh = Combo1
    rf = "printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    'htdc(0) = "cyan": gndc(0) = Me.Grid1.BackColorFixed
    'htdc(1) = "yellow": gndc(1) = Me.Grid1.BackColor
    'htdc(2) = "blue": gndc(2) = Me.Grid1.BackColor
    Grid1.ColWidth(0) = 1
    Grid1.Redraw = False
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(bimpbanner, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        Call htmlcolorgrid(bimpbanner, "c:\htmlgrid.htm", Grid1, rt, rh, rf, "linen", "khaki", "white")
        Grid1.Redraw = True
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmlgrid.htm", vbNormalFocus)
        Exit Sub
    End If

End Sub

Private Sub Form_Load()
    Me.Height = whssales.Height
    Me.Top = whssales.Top
    Me.Width = whssales.Width
    Me.Left = 0
    'Me.Left = whssales.Width - Me.Width
    brzlit.Caption = "Last updated:  " & Format(FileDateTime(brzloss), "M-d-yyyy h:mm am/pm")
    refresh_lists
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 200
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 3.5)
End Sub

Private Sub List1_Click()
    refresh_grid
End Sub
