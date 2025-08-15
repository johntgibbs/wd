VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Prodrcpts 
   Caption         =   "Production Receipts"
   ClientHeight    =   12240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   12240
   ScaleWidth      =   10845
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid5 
      Height          =   2535
      Left            =   0
      TabIndex        =   15
      Top             =   7560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4471
      _Version        =   327680
      BackColorFixed  =   16776960
      BackColorSel    =   255
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid4 
      Height          =   1455
      Left            =   0
      TabIndex        =   14
      Top             =   11280
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2566
      _Version        =   327680
      FocusRect       =   0
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   1215
      Left            =   0
      TabIndex        =   13
      Top             =   10080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   2143
      _Version        =   327680
      BackColorFixed  =   12648447
      FocusRect       =   0
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5880
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5880
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5880
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   5640
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   5520
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2990
      _Version        =   327680
      BackColorFixed  =   12632319
      Appearance      =   0
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add SKU"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4695
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8281
      _Version        =   327680
      BackColorFixed  =   8454143
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel SKU"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Hold Listing:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Receipt Dates:"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label wprod 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   5160
      Width           =   3735
   End
   Begin VB.Label wsku 
      Caption         =   "Label1"
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
      TabIndex        =   9
      Top             =   5160
      Width           =   375
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu insrecmenu 
         Caption         =   "Insert Record"
      End
      Begin VB.Menu delrecmenu 
         Caption         =   "Delete Record"
      End
      Begin VB.Menu addhold 
         Caption         =   "Add to On Hold List"
      End
   End
   Begin VB.Menu holdmenu 
      Caption         =   "Hold"
      Visible         =   0   'False
      Begin VB.Menu delhold 
         Caption         =   "Remove from Hold"
      End
      Begin VB.Menu edspallet 
         Caption         =   "Edit Start Pallet"
      End
      Begin VB.Menu edepallet 
         Caption         =   "Edit End Pallet"
      End
   End
End
Attribute VB_Name = "Prodrcpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim df1 As Boolean, df2 As Boolean, df3 As Boolean
Dim edcell As String

Private Sub refresh_holdlist()                                  'jv040615
    Dim ds As adodb.Recordset, s As String
    Dim ss As adodb.Recordset, pdesc As String, i As Integer, k As Integer
    Grid5.Clear: Grid5.Rows = 1: Grid5.Cols = 11
    s = "select * from holdlist Where hsource in ('Schedule', 'TEST HOLD')"         'jv122115
    s = s & " order by  sku, lot_num, opcode"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "select uom_type, description from sku_config where sku = '" & ds!sku & "'"
            Set ss = Wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                pdesc = ss!uom_type & " " & ss!description
            Else
                pdesc = "Unknown SKU"
            End If
            ss.Close
            s = ds!id & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & pdesc & Chr(9)
            s = s & ds!lot_num & Chr(9)
            s = s & ds!opcode & Chr(9)
            s = s & ds!spallet & Chr(9)
            s = s & ds!epallet & Chr(9)
            Grid5.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid5.Rows > 1 Then
        Grid1.Redraw = False
        For i = 1 To Grid5.Rows - 1
            s = "select * from prodrcv where sku = '" & Grid5.TextMatrix(i, 1) & "'"            'jv052515
            s = s & " and lot_num = '" & Grid5.TextMatrix(i, 3) & "'"                           'jv052515
            s = s & " and sp_flag = '" & Grid5.TextMatrix(i, 4) & "'"                           'jv052515
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                s = Format(DateAdd("yyyy", 2, ds!proddate), "MMddyy") & ds!sp_flag              'jv052515
                Grid5.TextMatrix(i, 7) = s
                Grid5.TextMatrix(i, 8) = Format(ds!recdate1, "M-dd-yyyy")
                Grid5.TextMatrix(i, 9) = Format(ds!recdate2, "M-dd-yyyy")
                Grid5.TextMatrix(i, 10) = Format(ds!recdate3, "M-dd-yyyy")
            End If
            ds.Close
            If Grid5.TextMatrix(i, 3) = Label2.Caption Then                                     'jv052515
                opc = Grid5.TextMatrix(i, 4)                                                    'jv052515
                For k = 0 To Grid1.Rows - 1                                                     'jv052515
                    If Grid1.TextMatrix(k, 1) = Grid5.TextMatrix(i, 1) And Left(Grid1.TextMatrix(k, 2), Len(opc)) = opc Then
                        Grid1.Row = k: Grid1.RowSel = k                                         'jv052515
                        Grid1.Col = 1: Grid1.ColSel = 2                                         'jv052515
                        Grid1.CellBackColor = Grid5.BackColorFixed                              'jv052515
                        Exit For                                                                'jv052515
                    End If                                                                      'jv052515
                Next k                                                                          'jv052515
            End If                                                                              'jv052515
        Next i
        Grid1.Row = 1: Grid1.Col = 5
        Grid1.Redraw = True
    End If
    s = "^ID|^SKU|<Product|^Lot|^OpCode|^Start|^End|^R12Lot|^Date1|^Date2|^Date3"
    Grid5.FormatString = s
    Grid5.ColWidth(0) = 1000
    Grid5.ColWidth(1) = 1000
    Grid5.ColWidth(2) = 3000
    Grid5.ColWidth(3) = 1000
    Grid5.ColWidth(4) = 1000
    Grid5.ColWidth(5) = 1000
    Grid5.ColWidth(6) = 1000
    Grid5.ColWidth(7) = 1000
    Grid5.ColWidth(8) = 1000
    Grid5.ColWidth(9) = 1000
    Grid5.ColWidth(10) = 1000
End Sub

Private Sub update_rct()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    If edcell = "rb" Then
        Grid1.Text = Val(Grid1.Text)
        If Val(Grid1.Text) = 0 Then Grid1.Text = ""
        edcell = ""
        Exit Sub
    End If
    sqlx = "select * from prodrcv where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Grid1.Text = Val(Grid1.Text)
        If Val(Grid1.Text) = 0 Then Grid1.Text = ""
        If edcell = "sr1" Then sqlx = "Update prodrcv set sr1 = " & Val(Grid1.Text) & " Where id = " & ds!id
        If edcell = "sr2" Then sqlx = "Update prodrcv set sr2 = " & Val(Grid1.Text) & " Where id = " & ds!id
        If edcell = "sr3" Then sqlx = "Update prodrcv set sr3 = " & Val(Grid1.Text) & " Where id = " & ds!id
        If edcell = "sr4" Then sqlx = "Update prodrcv set sr4 = " & Val(Grid1.Text) & " Where id = " & ds!id
        If edcell = "sr5" Then sqlx = "Update prodrcv set sr5 = " & Val(Grid1.Text) & " Where id = " & ds!id
        Wdb.Execute sqlx
    End If
    ds.Close
    edcell = ""
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "update_rct", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " update_rct - Error Number: " & eno
        End
    End If
End Sub

Private Sub printsc()
    Dim i As Integer
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.Print Tab(20); "Production Receiving"; Tab(60); Format(Text1, "m-d-yyyy")
    Printer.Print Tab(24); "Lot: "; Label2; Tab(60); Format(Text2, "m-d-yyyy")
    Printer.Print " "
    i = 0
    Printer.Print " "
    Printer.Print Grid1.TextMatrix(i, 1); Tab(5);
    Printer.Print Grid1.TextMatrix(i, 2); Tab(40);
    Printer.Print Grid1.TextMatrix(i, 3); Tab(48);
    Printer.Print Grid1.TextMatrix(i, 5); Tab(56);
    Printer.Print Grid1.TextMatrix(i, 6); Tab(64);
    Printer.Print Grid1.TextMatrix(i, 7); Tab(72);
    Printer.Print Grid1.TextMatrix(i, 8); Tab(80);
    Printer.Print Grid1.TextMatrix(i, 9); Tab(88);
    Printer.Print Grid1.TextMatrix(i, 10)
    acode = " "                                                             'jv063014
    
    For i = 1 To Grid1.Rows - 2
        If Grid1.TextMatrix(i, Grid1.Cols - 1) <> acode Then                'jv063014
            acode = Grid1.TextMatrix(i, Grid1.Cols - 1)                     'jv063014
            Printer.Print " "                                               'jv063014
            If acode = "TX A1" Then Printer.Print "Area 1"                  'jv063014
            If acode = "TX A2" Then Printer.Print "Area 2"                  'jv063014
            If acode = "TX A3" Then Printer.Print "Area 3"                  'jv063014
            If acode = "TX A4" Then Printer.Print "Area 4"                  'jv063014
            If acode = "TX A5" Then Printer.Print "Area 5"                  'jv063014
            If acode = "TX A6" Then Printer.Print "Area 6"                  'jv063014
            If acode = "SP A1" Then Printer.Print "Snack Plant Area 1"      'jv063014
            If acode = "SP A2" Then Printer.Print "Snack Plant Area 2"      'jv063014
            If acode = "SP A3" Then Printer.Print "Snack Plant Area 3"      'jv063014
            If acode = "SP A4" Then Printer.Print "Snack Plant Area 4"      'jv063014
            If acode = "SP A5" Then Printer.Print "Snack Plant Area 5"      'jv063014
            If acode = "SP A6" Then Printer.Print "Snack Plant Area 6"      'jv063014
        End If                                                              'jv063014
        'Printer.Print " "
        Printer.Print Grid1.TextMatrix(i, 1); Tab(5);
        Printer.Print Grid1.TextMatrix(i, 2); Tab(40);
        Printer.Print Val(Grid1.TextMatrix(i, 3)); Tab(48);
        Printer.Print Val(Grid1.TextMatrix(i, 5)); Tab(56);
        Printer.Print Val(Grid1.TextMatrix(i, 6)); Tab(64);
        Printer.Print Val(Grid1.TextMatrix(i, 7)); Tab(72);
        Printer.Print Val(Grid1.TextMatrix(i, 8)); Tab(80);
        Printer.Print Val(Grid1.TextMatrix(i, 9)); Tab(88);
        Printer.Print Val(Grid1.TextMatrix(i, 10))
    Next i
    Printer.Print " "
    Printer.Print Tab(40); String(56, "-")
    'Printer.Print " "
    i = Grid1.Rows - 1
    Printer.Print Grid1.TextMatrix(i, 1); Tab(5);
    Printer.Print Grid1.TextMatrix(i, 2); Tab(40);
    Printer.Print Val(Grid1.TextMatrix(i, 3)); Tab(48);
    Printer.Print Val(Grid1.TextMatrix(i, 5)); Tab(56);
    Printer.Print Val(Grid1.TextMatrix(i, 6)); Tab(64);
    Printer.Print Val(Grid1.TextMatrix(i, 7)); Tab(72);
    Printer.Print Val(Grid1.TextMatrix(i, 8)); Tab(80);
    Printer.Print Val(Grid1.TextMatrix(i, 9)); Tab(88);
    Printer.Print Val(Grid1.TextMatrix(i, 10))
    Printer.EndDoc
End Sub

Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    Dim ds2 As adodb.Recordset
    Dim pdays As Integer, psnk As String, i As Integer, k As Integer
    Dim hflag As Boolean                                        'jv061914
    'On Error GoTo vberror
    Grid1.Visible = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 13
    Grid1.FixedCols = 5
    sqlx = "select prodrcv.id,prodrcv.sku,uom_type,description,uom_per_pallet,units,"
    sqlx = sqlx & "sp_flag,recdate1,recdate2,recdate3,sr1,sr2,sr3,sr4,sr5"          'jv0513
    sqlx = sqlx & " from prodrcv,sku_config"
    sqlx = sqlx & " where proddate = '" & Combo1 & "'"
    sqlx = sqlx & " and prodrcv.sku = sku_config.sku"
    sqlx = sqlx & " order by prodrcv.sku,sp_flag"                   'jv061914
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        Form1.cdate = Format(Combo1, "m-d-yyyy")
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!id & Chr$(9)
            sqlx = sqlx & ds!sku & Chr$(9)
            psnk = " ": pdays = 0
            If ds!sp_flag = "Y" Then psnk = "*"
            psnk = ds!sp_flag & "-"                         'jv061914
            sqlx = sqlx & psnk & ds!description & " " & ds!uom_type & Chr$(9)
            sqlx = sqlx & Int((ds!units / ds!uom_per_pallet) + 0.75) & Chr$(9)
            If IsDate(ds!recdate1) = True Then pdays = 1
            If IsDate(ds!recdate2) = True Then pdays = 2
            If IsDate(ds!recdate3) = True Then pdays = 3
            sqlx = sqlx & pdays & Chr$(9)
            If ds!sr1 > 0 Then sqlx = sqlx & ds!sr1
            sqlx = sqlx & Chr$(9)
            If ds!sr2 > 0 Then sqlx = sqlx & ds!sr2
            sqlx = sqlx & Chr$(9)
            If ds!sr3 > 0 Then sqlx = sqlx & ds!sr3
            sqlx = sqlx & Chr$(9)
            If ds!sr4 > 0 Then sqlx = sqlx & ds!sr4
            sqlx = sqlx & Chr$(9)
            If ds!sr5 > 0 Then sqlx = sqlx & ds!sr5         'jv0513
            sqlx = sqlx & Chr$(9)                           'jv0513
            Set ds2 = Sdb.Execute("select * from skumast where sku = '" & ds!sku & "'")
            If ds2.BOF = False Then
                If ds2!psource = 2 Then sqlx = sqlx & Int((ds!units / ds2!pallet) + 0.75)
            End If
            ds2.Close
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^ID|^SKU|<Product|^Total|^Days|^SR1|^SR2|^SR3|^SR4|^SR5|^RB|^Net|^Area"   'jv063014
    Grid1.ColWidth(0) = 1: Grid1.ColWidth(1) = 450: Grid1.ColWidth(2) = 3300
    Grid1.ColWidth(3) = 800: Grid1.ColWidth(4) = 800: Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 800: Grid1.ColWidth(7) = 800: Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 800: Grid1.ColWidth(10) = 800: Grid1.ColWidth(11) = 800
    Grid1.ColWidth(12) = 800                    'jv063014
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1             'jv063014
            s = ""                                                                              'jv063014
            For k = 1 To Grid3.Rows - 1                                                         'jv063014
                If Grid3.TextMatrix(k, 0) = "500" Then                                          'jv063014
                    If Format(Grid3.TextMatrix(k, 1), "MMddyyyy") = Format(Combo1, "MMddyyyy") Then  'jv063014
                        If Grid3.TextMatrix(k, 2) = Grid1.TextMatrix(i, 1) Then                 'jv063014
                            If Grid3.TextMatrix(k, 6) = Left(Grid1.TextMatrix(i, 2), 1) Then    'jv063014
                                s = Grid3.TextMatrix(k, 7)                                      'jv063014
                                Grid3.Row = k: Grid3.TopRow = k                                 'jv063014
                                Exit For                                                        'jv063014
                            End If                                                              'jv063014
                        End If                                                                  'jv063014
                    End If                                                                      'jv063014
                End If                                                                          'jv063014
            Next k                                                                              'jv063014
            If s > " " Then                                                                     'jv063014
                For k = 0 To Grid4.Rows - 1                                                     'jv063014
                    If Grid4.TextMatrix(k, 1) = s Then                                          'jv063014
                        Grid4.Row = k: Grid4.TopRow = k                                         'jv063014
                        s = Left(Grid4.TextMatrix(k, 1), 2) & " "                               'jv063014
                        s = s & Grid4.TextMatrix(k, 3) & " "                                    'jv063014
                        s = s & Grid1.TextMatrix(i, 1) & " "                                    'jv063014
                        s = s & Grid1.TextMatrix(i, 3)                                          'jv063014
                        Grid1.TextMatrix(i, Grid1.Cols - 1) = s                                 'jv063014
                        Exit For                                                                'jv063014
                    End If                                                                      'jv063014
                Next k                                                                          'jv063014
            End If                                                                              'jv063014
        Next i                                  'jv063014
        Grid1.RowSel = Grid1.Row                'jv063014
        Grid1.Col = Grid1.Cols - 1              'jv063014
        Grid1.ColSel = Grid1.Col                'jv063014
        Grid1.Sort = 5                          'jv063014
        
        Grid1.FillStyle = flexFillRepeat                                                        'jv063014
        For i = 0 To Grid1.Rows - 1                                                             'jv063014
            If Left(Grid1.TextMatrix(i, Grid1.Cols - 1), 5) <> s Then                           'jv063014
                s = Left(Grid1.TextMatrix(i, Grid1.Cols - 1), 5)                                'jv063014
                hflag = Not hflag                                                               'jv063014
            End If                                                                              'jv063014
            If hflag = True Then                                                                'jv063014
                Grid1.Row = i: Grid1.RowSel = i                                                 'jv063014
                Grid1.Col = 5: Grid1.ColSel = Grid1.Cols - 1                                    'jv063014
                Grid1.CellBackColor = Grid3.BackColorFixed                                      'jv063014
            End If                                                                              'jv063014
        Next i                                                                                  'jv063014
    
    
    
        Grid1.AddItem Chr(9) & Chr(9) & "  Summary"
        k = Grid1.Rows - 1
        For i = 1 To Grid1.Rows - 2
            If Len(Grid1.TextMatrix(i, Grid1.Cols - 1)) > 2 Then                                    'jv063014
                Grid1.TextMatrix(i, Grid1.Cols - 1) = Left(Grid1.TextMatrix(i, Grid1.Cols - 1), 5)  'jv063014
            End If                                                                                  'jv063014
            If Val(Grid1.TextMatrix(i, 10)) > 0 Then
                Grid1.TextMatrix(i, 10) = Val(Grid1.TextMatrix(i, 10)) - Val(Grid1.TextMatrix(i, 5))
                Grid1.TextMatrix(i, 10) = Val(Grid1.TextMatrix(i, 10)) - Val(Grid1.TextMatrix(i, 6))
                Grid1.TextMatrix(i, 10) = Val(Grid1.TextMatrix(i, 10)) - Val(Grid1.TextMatrix(i, 7))
                Grid1.TextMatrix(i, 10) = Val(Grid1.TextMatrix(i, 10)) - Val(Grid1.TextMatrix(i, 8))
                Grid1.TextMatrix(i, 10) = Val(Grid1.TextMatrix(i, 10)) - Val(Grid1.TextMatrix(i, 9))
                Grid1.TextMatrix(i, 10) = Format(Val(Grid1.TextMatrix(i, 10)), "###")
            End If
            Grid1.TextMatrix(i, 11) = Grid1.TextMatrix(i, 3)
            Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 5))
            Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 6))
            Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 7))
            Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 8))
            Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 9))
            Grid1.TextMatrix(i, 11) = Format(Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 10)), "####")
            Grid1.TextMatrix(k, 3) = Format(Val(Grid1.TextMatrix(k, 3)) + Val(Grid1.TextMatrix(i, 3)), "####")
            Grid1.TextMatrix(k, 5) = Format(Val(Grid1.TextMatrix(k, 5)) + Val(Grid1.TextMatrix(i, 5)), "####")
            Grid1.TextMatrix(k, 6) = Format(Val(Grid1.TextMatrix(k, 6)) + Val(Grid1.TextMatrix(i, 6)), "####")
            Grid1.TextMatrix(k, 7) = Format(Val(Grid1.TextMatrix(k, 7)) + Val(Grid1.TextMatrix(i, 7)), "####")
            Grid1.TextMatrix(k, 8) = Format(Val(Grid1.TextMatrix(k, 8)) + Val(Grid1.TextMatrix(i, 8)), "####")
            Grid1.TextMatrix(k, 9) = Format(Val(Grid1.TextMatrix(k, 9)) + Val(Grid1.TextMatrix(i, 9)), "####")
            Grid1.TextMatrix(k, 10) = Format(Val(Grid1.TextMatrix(k, 10)) + Val(Grid1.TextMatrix(i, 10)), "####")
            Grid1.TextMatrix(k, 11) = Format(Val(Grid1.TextMatrix(k, 11)) + Val(Grid1.TextMatrix(i, 11)), "####")
        Next i
        Grid1.Row = 1: Call Grid1_RowColChange: Call wsku_Change
    End If
    Grid1.Visible = True
    refresh_holdlist                    'jv040615
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_grid34()            'jv063014
    Dim cfile As String, s As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    cfile = "s:\wd\data\plabels.500"
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 8
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7
            s = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3 & Chr(9)
            s = s & f4 & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7
            Grid3.AddItem s
        Loop
        Close #1
    End If
    Grid3.FormatString = "^Org|^Date|^SKU|<Product|^Pallets|^CodeDate|^Code|<Operation"
    Grid3.ColWidth(0) = 800
    Grid3.ColWidth(1) = 1000
    Grid3.ColWidth(2) = 800
    Grid3.ColWidth(3) = 3000
    Grid3.ColWidth(4) = 1000
    Grid3.ColWidth(5) = 1000
    Grid3.ColWidth(6) = 800
    Grid3.ColWidth(7) = 2000
    
    cfile = "s:\wd\data\opcodes.txt"
    Grid4.Clear: Grid4.Rows = 1: Grid4.Cols = 4
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3
            s = f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & f3
            Grid4.AddItem s
        Loop
        Close #1
    End If
    Grid4.FormatString = "^Org|<Operation|^Code|^Area"
    Grid4.ColWidth(0) = 800
    Grid4.ColWidth(1) = 2000
    Grid4.ColWidth(2) = 800
    Grid4.ColWidth(3) = 1000
End Sub

Private Sub addhold_Click()                                                 'jv040615
    Dim s As String, i As Integer
    Dim k As Integer, opc As String, zid As Long                            'jv052515
    i = Grid1.Row
    If Val(Grid1.TextMatrix(i, 0)) = 0 Then Exit Sub
    opc = Grid1.TextMatrix(i, 2)                                            'jv052515
    k = InStr(1, opc, "-")                                                  'jv052515
    opc = Left(opc, k - 1)                                                  'jv052515
    zid = wd_seq("HoldList", Form1.bbsr)                                    'jv052515
    s = "Insert into holdlist (id, sku, lot_num, opcode, spallet, epallet, hsource) values ("
    s = s & zid                                                             'jv052515
    s = s & ", '" & Grid1.TextMatrix(i, 1) & "'"
    s = s & ", '" & Label2.Caption & "'"
    s = s & ", '" & opc & "'"                                               'jv052515
    s = s & ", '001', 'EOR', 'TEST HOLD')"                                  'jv122115
    Wdb.Execute s
    refresh_holdlist
End Sub

Private Sub Combo1_Click()
    Dim seed As String
    seed = "12-31-" & Val(Right(Combo1, 4)) - 1
    Label2 = Right(Year(Combo1) & Format(DateDiff("y", seed, Combo1), "000"), 5)    'jv052515
    Call refresh_grid
End Sub

Private Sub Command2_Click()
    Call printsc
    Exit Sub
    Dim i As Integer, pl As String
    Dim x As Long, y As Long
    Printer.Font = "MS Sans Serif"
    Printer.FontBold = False
    Printer.FontSize = 12
    Printer.Print ""
    Printer.Print ""
    Printer.Print "Production Receipts - Produced: " & Format(Combo1, "m-d-yyyy")
    pl = "Receiving Dates: "
    If Text1 > "0" Then pl = pl & Format(Text1, "m-d-yyyy") & " "
    If Text2 > "0" Then pl = pl & Format(Text2, "m-d-yyyy") & " "
    If Text3 > "0" Then pl = pl & Format(Text3, "m-d-yyyy")
    Printer.FontSize = 10
    Printer.Print ""
    Printer.Print pl
    Printer.FontSize = 8
    Printer.Line (0, 1440)-(0, ((Grid1.Rows + 1) * 240) + 1440)
    x = 0
    For i = 1 To Grid1.Cols - 2
        x = x + Grid1.ColWidth(i)
        Printer.Line (x, 1440)-(x, ((Grid1.Rows + 1) * 240) + 1440)
    Next i
    For i = 0 To Grid1.Rows + 1
        Printer.Line (0, i * 240 + 1440)-(x, i * 240 + 1440)
    Next i
    For i = 0 To Grid1.Rows - 2
        x = 100
        For k = 0 To Grid1.Cols - 2
            Printer.PSet (x, i * 240 + 1500)
            Printer.Print Grid1.TextMatrix(i, k)
            x = x + Grid1.ColWidth(k)
        Next k
    Next i
    x = 100: i = Grid1.Rows
    For k = 0 To Grid1.Cols - 2
        Printer.PSet (x, i * 240 + 1500)
        Printer.Print Grid1.TextMatrix(i - 1, k)
        x = x + Grid1.ColWidth(k)
    Next k
    Printer.EndDoc
End Sub

Private Sub Command2_GotFocus()
    Command2.FontBold = True
End Sub

Private Sub Command2_LostFocus()
    Command2.FontBold = False
End Sub

Private Sub Command4_Click()                'Cancel SKU
    Dim sqlx As String
    On Error GoTo vberror
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) < 1 Then Exit Sub
    If MsgBox("Ok to cancel " & Grid1.TextMatrix(Grid1.Row, 2) & "?", vbYesNo + vbQuestion, "Are you sure?") = vbNo Then Exit Sub
    sqlx = "delete from prodrcv where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Wdb.Execute sqlx
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_grid
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command4_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command4_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command4_GotFocus()
    Command4.FontBold = True
End Sub

Private Sub Command4_LostFocus()
    Command4.FontBold = False
End Sub

Private Sub Command5_Click()                    'Add SKU
    Dim pdate As String, psku As String, pdays As Integer
    Dim pqty As Long, psnk As Boolean, zid As Long
    Dim ds As adodb.Recordset, sqlx As String
    Dim ds2 As adodb.Recordset, plot As String
    Dim pcode As String                                         'jv061914
    pcode = "A"                                                 'jv061914

    On Error GoTo vberror
    If IsDate(Combo1) = True Then
        pdate = Combo1
    Else
        pdate = InputBox("Please enter a valid date.", "Production Date", Form1.cdate)
    End If
    If Len(pdate) = 0 Then Exit Sub
    If IsDate(pdate) = False Then
        MsgBox "Date entered as: " & pdate & " not recognized as valid.", vbOKOnly, "Sorry, try again..."
        Exit Sub
    End If
    Form1.cdate = Format(pdate, "m-d-yyyy")
    psku = InputBox("Please enter a valid sku.", "SKU for " & pdate, "777")
    If Len(psku) = 0 Then Exit Sub
    sqlx = "select * from skumast where sku = '" & psku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "SKU: " & psku & " not found in skumast.", vbOKOnly, "Sorry, cannot insert.."
        ds.Close
        Exit Sub
    End If
    pdays = 1
    pcode = InputBox("Operation Code:", "Operation Code on labels...", pcode)   'jv061914
    If Len(pcode) = 0 Then                                                      'jv061914
        ds.Close ': db.Close                                                      'jv061914
        Exit Sub                                                                'jv061914
    End If                                                                      'jv061914
    pqty = InputBox("How many units are expected?", ds!pallet & " Units/pallet produced..", "10000")
    If pqty < 1 Then
        ds.Close
        Exit Sub
    End If
    zid = wd_seq("ProdRcv", Form1.bbsr)
    s = "INSERT INTO ProdRcv (ID, SKU, ProdDate, Units, SP_Flag, Lot_Num,"
    s = s & " RecDate1, RecDate2, RecDate3, SR1, SR2, SR3, SR4, SR5) VALUES ("          'jv0513
    s = s & zid & ","
    s = s & "'" & psku & "',"
    s = s & "'" & Format(pdate, "mm-dd-yyyy") & "',"
    s = s & Val(pqty) & ","
    s = s & "'" & UCase(pcode) & "',"                                                  'jv061914
    s = s & "'" & Label2 & "',"
    s = s & "'" & Format(pdate, "mm-dd-yyyy") & "',"
    If pdays > 1 Then
        If WeekDay(pdate) = 6 Then
            If MsgBox("Production this Saturday?", vbYesNo, "Friday Production") = vbYes Then
                s = s & "'" & Format(DateAdd("d", 1, pdate), "mm-dd-yyyy") & "',"
                If pdays > 2 Then
                    s = s & "'" & Format(DateAdd("d", 2, pdate), "mm-dd-yyyy") & "',"
                Else
                    s = s & "NULL,"
                End If
            Else
                s = s & "'" & Format(DateAdd("d", 3, pdate), "mm-dd-yyyy") & "',"
                If pdays > 2 Then
                    s = s & "'" & Format(DateAdd("d", 4, pdate), "mm-dd-yyyy") & "',"
                Else
                    s = s & "NULL,"
                End If
            End If
        Else
            If WeekDay(pdate) = 7 Then
                s = s & "'" & Format(DateAdd("d", 2, pdate), "mm-dd-yyyy") & "',"
                If pdays > 2 Then
                    s = s & "'" & Format(DateAdd("d", 3, pdate), "mm-dd-yyyy") & "',"
                Else
                    s = s & "NULL,"
                End If
            Else
                s = s & "'" & Format(DateAdd("d", 1, pdate), "mm-dd-yyyy") & "',"
                If pdays > 2 Then
                    If WeekDay(pdate) = 5 Then
                        If MsgBox("Production this Saturday?", vbYesNo, "Thursday Product") = vbYes Then
                            s = s & "'" & Format(DateAdd("d", 2, pdate), "mm-dd-yyyy") & "',"
                        Else
                            s = s & "'" & Format(DateAdd("d", 4, pdate), "mm-dd-yyyy") & "',"
                        End If
                    Else
                        s = s & "'" & Format(DateAdd("d", 2, pdate), "mm-dd-yyyy") & "',"
                    End If
                Else
                    s = s & "NULL,"
                End If
            End If
        End If
    Else
        s = s & "NULL,NULL,"
    End If
    i = Int(pqty / ds!pallet)
    If ds!whs_num = 1 Then
        s = s & i & ",0,0,0,0)"
    Else
        If ds!whs_num = 2 Then
            s = s & "0," & i & ",0,0,0)"
        Else
            If ds!whs_num = 3 Then
                s = s & "0,0," & i & ",0,0)"
            Else
                If ds!whs_num = 4 Then
                    s = s & "0,0,0," & i & ",0)"
                Else
                    If ds!whs_num = 5 Then
                        s = s & "0,0,0,0," & i & ")"
                    Else
                        s = s & "0,0,0,0,0)"
                    End If
                End If
            End If
        End If
    End If
    Wdb.Execute s
    ds.Close
    Call refresh_grid
    For i = 1 To Grid1.Rows - 1
        y = i
        If Grid1.TextMatrix(i, 1) = psku Then Exit For
    Next i
    Grid1.Row = y: Grid1.Col = 5
    Grid1.SetFocus
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command5_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command5_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command5_GotFocus()
    Command5.FontBold = True
End Sub

Private Sub Command5_LostFocus()
    Command5.FontBold = False
End Sub

Private Sub delhold_Click()                                     'jv040615
    Dim s As String, i As Integer
    i = Grid5.Row
    If Val(Grid5.TextMatrix(i, 0)) = 0 Then Exit Sub
    s = "delete from holdlist where id = " & Grid5.TextMatrix(i, 0)
    Wdb.Execute s
    refresh_holdlist
End Sub

Private Sub delrecmenu_Click()
    Command4_Click
End Sub

Private Sub edepallet_Click()                                   'jv040615
    Dim s As String, i As Integer, pno As String
    i = Grid5.Row
    If Val(Grid5.TextMatrix(i, 0)) = 0 Then Exit Sub
    pno = Grid5.TextMatrix(i, 6)
    pno = InputBox("End Pallet #:", "End Pallet #.....", pno)
    If Len(pno) = 0 Then Exit Sub
    If Val(pno) > 0 Then
        pno = Format(Val(pno), "000")
    Else
        pno = "EOR"
    End If
    If Len(pno) <> 3 Then Exit Sub
    If pno < Grid5.TextMatrix(i, 5) Then Exit Sub
    s = "Update holdlist set epallet = '" & pno & "' Where id = " & Grid5.TextMatrix(i, 0)
    Wdb.Execute s
    Grid5.TextMatrix(i, 6) = pno
End Sub

Private Sub edspallet_Click()                                   'jv040615
    Dim s As String, i As Integer, pno As String
    i = Grid5.Row
    If Val(Grid5.TextMatrix(i, 0)) = 0 Then Exit Sub
    pno = Grid5.TextMatrix(i, 5)
    pno = InputBox("Start Pallet #:", "Start Pallet #.....", pno)
    If Len(pno) = 0 Then Exit Sub
    If Val(pno) > 0 Then
        pno = Format(Val(pno), "000")
    Else
        pno = "EOR"
    End If
    If Len(pno) <> 3 Then Exit Sub
    If pno > Grid5.TextMatrix(i, 6) Then Exit Sub
    s = "Update holdlist set spallet = '" & pno & "' Where id = " & Grid5.TextMatrix(i, 0)
    Wdb.Execute s
    Grid5.TextMatrix(i, 5) = pno
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Len(edcell) > 0 Then
        If MsgBox("Update Receipt record?", vbYesNo + vbQuestion, "Save changes...") = vbYes Then
            Call update_rct
        Else
            edcell = ""
        End If
    End If
    If Prodrcpts.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            If Form1.FrmGrid.TextMatrix(i, 0) = "prodrcpts" Then
                Form1.FrmGrid.TextMatrix(i, 1) = Prodrcpts.Top
                Form1.FrmGrid.TextMatrix(i, 2) = Prodrcpts.Left
                Form1.FrmGrid.TextMatrix(i, 3) = Prodrcpts.Height
                Form1.FrmGrid.TextMatrix(i, 4) = Prodrcpts.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Prodrcpts.ActiveControl.Name = "Grid1" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command5_Click
        If KeyCode = 46 Or KeyCode = 120 Then Call Command4_Click
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer, ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    df1 = False: df2 = False: df3 = False
    For i = 1 To Form1.FrmGrid.Rows - 1
        If Form1.FrmGrid.TextMatrix(i, 0) = "prodrcpts" Then
            Prodrcpts.Top = Val(Form1.FrmGrid.TextMatrix(i, 1))
            Prodrcpts.Left = Val(Form1.FrmGrid.TextMatrix(i, 2))
            Prodrcpts.Height = Val(Form1.FrmGrid.TextMatrix(i, 3))
            Prodrcpts.Width = Val(Form1.FrmGrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    refresh_grid34                      'jv063014
    Combo1.Clear
    sqlx = "select distinct proddate from prodrcv order by proddate"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem Format(ds!proddate, "m-d-yyyy")
            ds.MoveNext
        Loop
        For i = 0 To Combo1.ListCount - 1
            If Combo1.List(i) = Form1.cdate Then
                Combo1.ListIndex = i
                Exit For
            End If
        Next i
        If Combo1.ListIndex < 0 Then Combo1.ListIndex = 0
    End If
    ds.Close
    Grid2.Cols = 4
    Grid2.FormatString = "Warehouse|^OnHand|^Grouped|^Available"
    Grid2.ColWidth(0) = 1800: Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 1000: Grid2.ColWidth(3) = 1000
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "form_load", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " form_load - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 110
    Grid3.Width = Grid1.Width           'jv063014
    Grid4.Width = Grid1.Width           'jv063014
    Grid5.Width = Grid1.Width           'jv040615
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_GotFocus()
    Grid1.FocusRect = flexFocusNone
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
    Dim i As Integer, k As Integer
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Grid1.Col = Grid1.Cols - 1 Then
            SendKeys "{HOME}{DOWN}"
        Else
            SendKeys "{RIGHT}"
        End If
        Exit Sub
    End If
    If Grid1.Row = 0 Then Exit Sub
    If Grid1.Col < 5 Or Grid1.Col > 10 Then Exit Sub
    If Len(edcell) = 0 Then Grid1.Text = ""
    If Grid1.Col = 5 Then edcell = "sr1"
    If Grid1.Col = 6 Then edcell = "sr2"
    If Grid1.Col = 7 Then edcell = "sr3"
    If Grid1.Col = 8 Then edcell = "sr4"
    If Grid1.Col = 9 Then edcell = "sr5"    'jv0513
    If Grid1.Col = 10 Then edcell = "rb"
    If KeyAscii = 8 Then
        If Len(Grid1.Text) > 1 Then
            Grid1.Text = Left(Grid1.Text, Len(Grid1.Text) - 1)
        Else
            Grid1.Text = ""
        End If
    End If
    If KeyAscii > 31 And KeyAscii < 127 Then
        Grid1.Text = Grid1.Text & Chr(KeyAscii)
    End If
    k = Grid1.Rows - 1
    Grid1.TextMatrix(k, 3) = ""
    Grid1.TextMatrix(k, 5) = "": Grid1.TextMatrix(k, 6) = ""
    Grid1.TextMatrix(k, 7) = "": Grid1.TextMatrix(k, 8) = ""
    Grid1.TextMatrix(k, 9) = "": Grid1.TextMatrix(k, 10) = "": Grid1.TextMatrix(k, 11) = ""     'jv0513
    For i = 1 To Grid1.Rows - 2
        Grid1.TextMatrix(i, 11) = Grid1.TextMatrix(i, 3)
        Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 5))
        Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 6))
        Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 7))
        Grid1.TextMatrix(i, 11) = Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 8))
        Grid1.TextMatrix(i, 11) = Format(Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 9)), "####")
        Grid1.TextMatrix(i, 11) = Format(Val(Grid1.TextMatrix(i, 11)) - Val(Grid1.TextMatrix(i, 10)), "####")
        Grid1.TextMatrix(k, 3) = Format(Val(Grid1.TextMatrix(k, 3)) + Val(Grid1.TextMatrix(i, 3)), "####")
        Grid1.TextMatrix(k, 5) = Format(Val(Grid1.TextMatrix(k, 5)) + Val(Grid1.TextMatrix(i, 5)), "####")
        Grid1.TextMatrix(k, 6) = Format(Val(Grid1.TextMatrix(k, 6)) + Val(Grid1.TextMatrix(i, 6)), "####")
        Grid1.TextMatrix(k, 7) = Format(Val(Grid1.TextMatrix(k, 7)) + Val(Grid1.TextMatrix(i, 7)), "####")
        Grid1.TextMatrix(k, 8) = Format(Val(Grid1.TextMatrix(k, 8)) + Val(Grid1.TextMatrix(i, 8)), "####")
        Grid1.TextMatrix(k, 9) = Format(Val(Grid1.TextMatrix(k, 9)) + Val(Grid1.TextMatrix(i, 9)), "####")
        Grid1.TextMatrix(k, 10) = Format(Val(Grid1.TextMatrix(k, 10)) + Val(Grid1.TextMatrix(i, 10)), "####")
        Grid1.TextMatrix(k, 11) = Format(Val(Grid1.TextMatrix(k, 11)) + Val(Grid1.TextMatrix(i, 11)), "####")
    Next i
End Sub

Private Sub Grid1_LeaveCell()
    If Len(edcell) > 0 Then Call update_rct
End Sub

Private Sub Grid1_LostFocus()
    If Len(edcell) > 0 Then Call update_rct
    Grid1.FocusRect = flexFocusLight
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu                                                         'jv063014
End Sub

Private Sub Grid1_RowColChange()
    If Grid1.Rows > 1 And Grid1.Row > 0 And Grid1.Row < Grid1.Rows - 1 Then
        wsku = Grid1.TextMatrix(Grid1.Row, 1)
        wprod = Grid1.TextMatrix(Grid1.Row, 2)
    End If
End Sub

Private Sub Grid5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu holdmenu
End Sub

Private Sub insrecmenu_Click()
    Command5_Click                       'jv063014
End Sub

Private Sub Text1_Change()
    Dim sqlx As String, ds As adodb.Recordset, nd As Integer
    On Error GoTo vberror
    If df1 = True And (IsDate(Text1) Or Len(Text1) = 0) Then
        sqlx = "select * from prodrcv"
        sqlx = sqlx & " where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Set ds = Wdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            nd = 0
            If IsDate(Text1) Then nd = nd + 1
            If IsDate(Text2) Then nd = nd + 1
            If IsDate(Text3) Then nd = nd + 1
            Grid1.TextMatrix(Grid1.Row, 4) = nd
            If Len(Text1) = 0 Then
                sqlx = "Update prodrcv set recdate1 = NULL Where id = " & ds!id
                Wdb.Execute sqlx
            Else
                sqlx = "Update prodrcv set recdate1 = '" & Format(Text1, "m-d-yyyy") & "' Where id = " & ds!id
                Wdb.Execute sqlx
            End If
        End If
        ds.Close
        df1 = False
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "text1_change", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " text1_change - Error Number: " & eno
        End
    End If
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0: Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    df1 = True
End Sub

Private Sub Text2_Change()
    Dim sqlx As String, ds As adodb.Recordset, nd As Integer
    On Error GoTo vberror
    If df2 = True And (IsDate(Text2) Or Len(Text2) = 0) Then
        sqlx = "select * from prodrcv"
        sqlx = sqlx & " where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Set ds = Wdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            nd = 0
            If IsDate(Text1) Then nd = nd + 1
            If IsDate(Text2) Then nd = nd + 1
            If IsDate(Text3) Then nd = nd + 1
            Grid1.TextMatrix(Grid1.Row, 4) = nd
            If Len(Text2) = 0 Then
                sqlx = "Update prodrcv set recdate2 = NULL Where id = " & ds!id
                Wdb.Execute sqlx
            Else
                sqlx = "Update prodrcv set recdate2 = '" & Format(Text2, "m-d-yyyy") & "' Where id = " & ds!id
                Wdb.Execute sqlx
            End If
        End If
        ds.Close
        df2 = False
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "text2_change", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " text2_change - Error Number: " & eno
        End
    End If
End Sub

Private Sub Text2_GotFocus()
    Text2.SelStart = 0: Text2.SelLength = Len(Text2)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    df2 = True
End Sub

Private Sub Text3_Change()
    Dim sqlx As String, ds As adodb.Recordset, nd As Integer
    On Error GoTo vberror
    If df3 = True And (IsDate(Text3) Or Len(Text3) = 0) Then
        sqlx = "select * from prodrcv"
        sqlx = sqlx & " where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Set ds = Wdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            nd = 0
            If IsDate(Text1) Then nd = nd + 1
            If IsDate(Text2) Then nd = nd + 1
            If IsDate(Text3) Then nd = nd + 1
            Grid1.TextMatrix(Grid1.Row, 4) = nd
            'ds.Edit
            If Len(Text3) = 0 Then
                sqlx = "Update prodrcv set recdate3 = NULL Where id = " & ds!id
                Wdb.Execute sqlx
            Else
                sqlx = "Update prodrcv set recdate3 = '" & Format(Text3, "m-d-yyyy") & "' Where id = " & ds!id
                Wdb.Execute sqlx
            End If
        End If
        ds.Close
        df3 = False
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "text3_change", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " text3_change - Error Number: " & eno
        End
    End If
End Sub

Private Sub Text3_GotFocus()
    Text3.SelStart = 0: Text3.SelLength = Len(Text3)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    df3 = True
End Sub

Private Sub wsku_Change()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Grid2.Rows = 1
    sqlx = "select * from whstotals,warehouses"
    sqlx = sqlx & " where whstotals.whs_num = warehouses.whs_num"
    sqlx = sqlx & " and sku = '" & wsku & "'"
    sqlx = sqlx & " order by whstotals.whs_num"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!whsname & Chr(9)
            sqlx = sqlx & ds!count_qty & Chr(9)
            sqlx = sqlx & ds!grp_qty & Chr(9)
            sqlx = sqlx & ds!avail
            Grid2.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Text1 = "": Text2 = "": Text3 = ""
    sqlx = "select * from prodrcv where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        If IsDate(ds!recdate1) Then Text1 = Format(ds!recdate1, "m-d-yyyy")
        If IsDate(ds!recdate2) Then Text2 = Format(ds!recdate2, "m-d-yyyy")
        If IsDate(ds!recdate3) Then Text3 = Format(ds!recdate3, "m-d-yyyy")
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "wsku_change", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " wsku_change - Error Number: " & eno
        End
    End If
End Sub

