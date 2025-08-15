VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form skuoutsk 
   Caption         =   "Import Stock History from R12"
   ClientHeight    =   9030
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13710
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   13710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   11160
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5895
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   10398
      _Version        =   327680
      ForeColor       =   4194368
      BackColorFixed  =   12648384
      BackColorSel    =   8421376
      FocusRect       =   0
   End
   Begin VB.TextBox edate 
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
      Left            =   4680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Date:"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1095
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu actmenu 
      Caption         =   "Actions"
      Begin VB.Menu vuebranch 
         Caption         =   "View Current Branch"
      End
      Begin VB.Menu proccb 
         Caption         =   "Process Current Branch"
      End
      Begin VB.Menu procab 
         Caption         =   "Process All Branches"
      End
   End
End
Attribute VB_Name = "skuoutsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_branches()
    Dim i As Integer
    Combo1.Clear
    'Combo1.AddItem "All-All Branches"
    For i = 1 To 99
        'If branchrec(i).oraloc > " " Then Combo1.AddItem Format(branchrec(i).branchno, "000") & "-" & branchrec(i).branchname
        If branchrec(i).oraloc > " " Then Combo1.AddItem Format(branchrec(i).branchno, "000")
    Next i
    Combo1.ListIndex = 0
End Sub

Private Sub save_recs()
    Dim cfile As String, i As Integer
    'cfile = "S:\wd\html\stock\stk" & Format(edate, "MMddyyyy") & ".csv"
    cfile = "\\BBC-03-FILESVR\SharedGroups\wd\html\stock\" & Combo1 & "\stk" & Format(edate, "MMddyyyy") & ".csv"
    'MsgBox cfile
    Open cfile For Output As #1
    For i = 1 To Grid1.Rows - 1
        Write #1, Grid1.TextMatrix(i, 0);           'Whs
        Write #1, Grid1.TextMatrix(i, 1);           'SKU
        Write #1, Grid1.TextMatrix(i, 3);           'Start Inv
        Write #1, Grid1.TextMatrix(i, 4);           'TransIn
        Write #1, Grid1.TextMatrix(i, 5);           'TransOut
        Write #1, Grid1.TextMatrix(i, 6);           'Net
        Write #1, Grid1.TextMatrix(i, 7);           'Units/Wrap
        Write #1, Grid1.TextMatrix(i, 8);           'Status
        Write #1, Grid1.TextMatrix(i, 9)            'Loads
    Next i
    Close #1
End Sub

Private Sub refresh_skus()
    Dim ds As ADODB.Recordset, s As String, sdesc As String, hrow As Boolean, psku As String
    Grid1.Redraw = False
    Grid1.FontName = "Arial"
    Grid1.FontBold = True
    Grid1.FontSize = 8
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 10
    's = "select sku, count(*) from bimp where plantwhs <> 'DRY' group by sku"
    's = "select sku, count(*) from bimp where plantwhs in ('A10', 'K10', 'T10') group by sku"
    s = "select sku, count(*) from bimp where branchwhs = '" & Combo1 & "'"
    s = s & " and plantwhs in ('A10', 'K10', 'T10') group by sku"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = Combo1 & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & skurec(Val(ds!sku)).unit & " " & skurec(Val(ds!sku)).desc & Chr(9)
            's = s & edate & Chr(9)
            s = s & "0" & Chr(9)
            s = s & "0" & Chr(9)
            s = s & "0" & Chr(9)
            s = s & "0" & Chr(9)
            s = s & skurec(Val(ds!sku)).wrapunits
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    'Grid1.Row = 1: Grid1.Col = 1
    s = "^Whs|^SKU|<Product|^Start Inv|^TransIn|^TransOut|^Net OnHand|^Wraps|^Status|^Loads"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 900
    Grid1.ColWidth(1) = 900
    Grid1.ColWidth(2) = 2900
    Grid1.ColWidth(3) = 1400
    Grid1.ColWidth(4) = 1400
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1500
    Grid1.ColWidth(7) = 1500
    Grid1.ColWidth(8) = 1500
    Grid1.ColWidth(9) = 1500
    Grid1.Redraw = True
End Sub

Private Sub refresh_r12_history()
    Dim tdate As String, ds As ADODB.Recordset, t1 As String, t2 As String
    Dim t3 As Long, t4 As Long, t5 As Long, t6 As Long, t7 As Long, t8 As Long
    Dim brp As String, rc As Long, cldate As String
    rc = 0
    tdate = Format(DateAdd("d", 1, edate), "mm-dd-yyyy")
    date1 = Mid(tdate, 1, 3) & "01" & Mid(tdate, 6, 5)
    cldate = Format(DateAdd("d", -1, date1), "mm-dd-yyyy")
    
    q = "select g.inventory_item_id, m.segment1," & _
        " g.subinventory_code," & _
        " g.primary_quantity" & _
        " from gmf_period_balances g, mtl_system_items_b m" & _
        " where g.subinventory_code = '" & Combo1 & "'" & _
        " and m.organization_id = g.organization_id" & _
        " and m.inventory_item_id = g.inventory_item_id" & _
        " and m.segment1 >= '100' and m.segment1 <= '9999'" & _
        " and g.acct_period_id in" & _
        " (select acct_period_id from org_acct_periods" & _
        " Where organization_id = g.organization_id" & _
        " and schedule_close_date = TO_DATE('" & Format(cldate, "dd-mmm-yyyy") & "'))" & _
        " order by m.segment1"
    
    t1 = Format(Now, "hh:mm:ss")
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 1) = ds(1) Then
                    Grid1.TextMatrix(i, 3) = Val(Grid1.TextMatrix(i, 3)) + ds(3)
                    DoEvents
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    DoEvents
        
    q = "select o.inventory_item_id, m.segment1," & _
        " o.subinventory_code," & _
        " o.transaction_quantity" & _
        " from mtl_material_transactions o, mtl_system_items_b m" & _
        " where o.subinventory_code = '" & Combo1 & "'" & _
        " and m.organization_id = o.organization_id" & _
        " and m.inventory_item_id = o.inventory_item_id" & _
        " and m.segment1 >= '100' and m.segment1 <= '9999'" & _
        " and o.transaction_date < TO_DATE('" & Format(tdate, "dd-mmm-yyyy") & "')" & _
        " and o.transaction_date >= TO_DATE('" & Format(date1, "dd-mmm-yyyy") & "')"
        
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 1) = ds(1) Then
                    If ds(3) > 0 Then
                        Grid1.TextMatrix(i, 4) = Val(Grid1.TextMatrix(i, 4)) + ds(3)
                    Else
                        Grid1.TextMatrix(i, 5) = Val(Grid1.TextMatrix(i, 5)) + ds(3)
                    End If
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    DoEvents
    
    For i = 1 To Grid1.Rows - 1
        Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 3)) + Val(Grid1.TextMatrix(i, 4)) + Val(Grid1.TextMatrix(i, 5))
        If Val(Grid1.TextMatrix(i, 6)) <= Val(Grid1.TextMatrix(i, 7)) Then Grid1.TextMatrix(i, 8) = "Out"
        DoEvents
        If Val(Grid1.TextMatrix(i, 3)) <= Val(Grid1.TextMatrix(i, 7)) And Val(Grid1.TextMatrix(i, 4)) = 0 And Val(Grid1.TextMatrix(i, 5)) = 0 Then
            Grid1.TextMatrix(i, 8) = "InActive"
        End If
    Next i
    
    
    q = "select product_no,tran_qty from bolinf.inv_adj_input_detail"
    q = q & " Where branch_no = " & Val(Combo1)
    q = q & " and tran_type = '1'"
    q = q & " and tran_date = TO_DATE('" & Format(edate, "dd-mmm-yyyy") & "')"
    q = q & " order by product_no"
    'MsgBox q
    
    Set ds = r12db.Execute(q)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            For i = 1 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 1) = ds(0) Then
                    Grid1.TextMatrix(i, 9) = Val(Grid1.TextMatrix(i, 9)) + ds(1)
                    Exit For
                End If
            Next i
            ds.MoveNext
        Loop
    End If
    ds.Close
    DoEvents
    
    For i = Grid1.Rows - 1 To 1 Step -1
        'If Grid1.TextMatrix(i, 7) <> "Out" Then
        If Grid1.TextMatrix(i, 8) < ".." Then
            If Grid1.Rows < 2 Then
                Grid1.Rows = 1
            Else
                'Grid1.RemoveItem i
            End If
        End If
    Next i
End Sub

Private Sub Command1_Click()
    update_testdb
End Sub

Private Sub Form_Load()
    Me.Left = 0
    Me.Top = bimpbanner.Label2.Top
    Me.Width = bimpbanner.Width
    Me.Height = bimpbanner.Height - bimpbanner.Label2.Top
    edate.Text = Format(Now, "MM-dd-yyyy")
    refresh_branches
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 200
    If Me.Height > 2000 Then Grid1.Height = Me.Height - (Combo1.Height * 4)
End Sub

Private Sub procab_Click()
    Dim s As String, i As Integer, k As Integer, sdate As String
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    
    s = "9-01-2017"
    s = InputBox("Start Date:", "Starting date...", s)
    If Len(s) = 0 Then Exit Sub
    sdate = s
    s = InputBox("End Date:", "Ending date...", s)
    If Len(s) = 0 Then Exit Sub
    
    i = DateDiff("d", sdate, s)
    Screen.MousePointer = 11
    For k = 0 To i - 1
        edate = Format(DateAdd("d", k, sdate), "M-d-yyyy")
        
        For j = 0 To Combo1.ListCount - 1
            Combo1.ListIndex = j
            refresh_skus
            DoEvents
            If Grid1.Rows > 2 Then
                refresh_r12_history
            End If
            save_recs
        Next j
        
    Next k
    Screen.MousePointer = 0
End Sub

Private Sub proccb_Click()
    Dim s As String, i As Integer, k As Integer, sdate As String
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
        
    s = "9-01-2018"
    s = InputBox("Start Date:", "Starting date...", s)
    If Len(s) = 0 Then Exit Sub
    sdate = s
    s = InputBox("End Date:", "Ending date...", s)
    If Len(s) = 0 Then Exit Sub
    
    i = DateDiff("d", sdate, s)
    Screen.MousePointer = 11
    For k = 0 To i - 1
        edate = Format(DateAdd("d", k, sdate), "M-d-yyyy")
        refresh_skus
        DoEvents
        If Grid1.Rows > 2 Then
            refresh_r12_history
        End If
        save_recs
    Next k
    Screen.MousePointer = 0
End Sub

Private Sub vuebranch_Click()
    Dim i As Integer
    If r12access = False Then
        connect_r12
        DoEvents
    End If
    If r12access = False Then Exit Sub
    Screen.MousePointer = 11
    refresh_skus
    DoEvents
    If Grid1.Rows > 2 Then
        refresh_r12_history
    End If
    Screen.MousePointer = 0
End Sub
