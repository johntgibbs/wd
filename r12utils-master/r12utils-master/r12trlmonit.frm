VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form r12trlmonit 
   Caption         =   "Trailer Ticket Processing"
   ClientHeight    =   9540
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8970
   LinkTopic       =   "Form2"
   ScaleHeight     =   9540
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   8760
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1296
      _Version        =   327680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   0
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   2775
      Left            =   0
      TabIndex        =   2
      Top             =   6360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4895
      _Version        =   327680
      BackColorFixed  =   12648384
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
      _Version        =   327680
      BackColorFixed  =   12632319
      BackColorSel    =   255
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
      _Version        =   327680
      BackColorFixed  =   12648447
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label4 
      Caption         =   "R12 Material Transactions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   6120
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Posted Bill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Trailer Bills of Laden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Ship Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu prtmenu 
      Caption         =   "Print"
      Begin VB.Menu prtjob 
         Caption         =   "Jobbing Bill"
      End
   End
End
Attribute VB_Name = "r12trlmonit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub refresh_grid1(sd As String)
    Dim cfile As String, s As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim logpath As String, newrec As Boolean
    'Text1 = sd
    'logpath = Form1.pallogs
    If Form1.Combo1 = "500" Then logpath = "\\bbc-01-prodtrk\wd\pallogs\"
    If Form1.Combo1 = "501" Then logpath = "\\bbba-03-dc\f\user\waredist\data\pallogs\"
    If Form1.Combo1 = "502" Then logpath = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5: Grid1.Redraw = False: Grid1.Visible = False
   
    cfile = logpath & "bill" & Format(sd, "mmddyyyy") & ".txt"
    'MsgBox cfile
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input Shared As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
            s = "B" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
            s = s & Trim(StrConv(f4, vbProperCase)) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
            s = s & f14 & Chr(9) & f15 & Chr(9) & f16 & Chr(9) & Trim(StrConv(f4, vbProperCase)) & f15 'f16
            
            s = "B" & Chr(9) & f2 & Chr(9) & f4 & Chr(9) & f13 & Chr(9) & f16
            newrec = True
            For k = 0 To Grid1.Rows - 1
                If Grid1.TextMatrix(k, 4) = f16 Then
                    newrec = False
                    Exit For
                End If
            Next k
            If newrec = True Then Grid1.AddItem s
        Loop
        Close #1
    End If

    's = "O" & Chr(9) & "ZOP" & Chr(9) & "Order Pick moves" & Chr(9) & "POSTED" & Chr(9) '& "52438"
    'MsgBox DateDiff("d", Now, Text1), vbOKOnly, "Test date"
    If DateDiff("d", Now, Text1) < 0 Then
        If Form1.Combo1 = "500" Then
            s = "O" & Chr(9) & "ZOP" & Chr(9) & "Order Pick and Snack Plant moves" & Chr(9) & "POSTED" & Chr(9)
            s = s & "50"
        End If
        If Form1.Combo1 = "501" Then
            s = "O" & Chr(9) & "ZOP" & Chr(9) & "Order Pick moves" & Chr(9) & "POSTED" & Chr(9)
            s = s & "51"
        End If
        If Form1.Combo1 = "502" Then
            s = "O" & Chr(9) & "ZOP" & Chr(9) & "Order Pick moves" & Chr(9) & "POSTED" & Chr(9)
            s = s & "52"
        End If
        's = s & Format(DateDiff("d", "1-1-2012", Text1) - 1, "#")
        s = s & Format(DateDiff("d", "1-1-2012", Text1), "#")
        Grid1.AddItem s
    End If
    Grid1.Row = 0: Grid1.RowSel = 0: Grid1.Col = 2: Grid1.ColSel = 2
    Grid1.Sort = 5
    s = "^Type|^GroupCode|<Trailer|^Status|^Ticket"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 600
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1300
    Grid1.Redraw = True
    Grid1.Visible = True
    If Grid1.Rows > 1 Then
        Grid1.Row = 1
        Call refresh_grid2(Grid1.TextMatrix(Grid1.Row, 4))
    End If
End Sub

Private Sub refresh_grid2(rid As String)
    Dim cfile As String, s As String
    Dim f0 As String, f1 As String, f2 As String, f3 As String
    Dim f4 As String, f5 As String, f6 As String, f7 As String
    Dim f8 As String, f9 As String, f10 As String, f11 As String
    Dim f12 As String, f13 As String, f14 As String, f15 As String
    Dim logpath As String
    'logpath = Form1.pallogs
    'logpath = "\\bbc-01-prodtrk\wd\pallogs\"
    If Form1.Combo1 = "500" Then logpath = "\\bbc-01-prodtrk\wd\pallogs\"
    If Form1.Combo1 = "501" Then logpath = "\\bbba-03-dc\f\user\waredist\data\pallogs\"
    If Form1.Combo1 = "502" Then logpath = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 17: Grid2.Redraw = False: Grid2.Visible = False
    cfile = logpath & "RO" & rid & ".txt"
    'MsgBox cfile
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input Shared As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15 ', f16
            s = "RO" & Chr(9) & f0 & Chr(9) & f1 & Chr(9) & f2 & Chr(9) & Trim(f3) & Chr(9)
            s = s & Trim(StrConv(f4, vbProperCase)) & Chr(9) & f5 & Chr(9) & f6 & Chr(9) & f7 & Chr(9) & f8 & Chr(9)
            s = s & f9 & Chr(9) & f10 & Chr(9) & f11 & Chr(9) & f12 & Chr(9) & f13 & Chr(9)
            s = s & f14 & Chr(9) & f15 '& Chr(9) & f16 & Chr(9) & Trim(StrConv(f4, vbProperCase)) & f15 'f16
            Grid2.AddItem s
        Loop
        Close #1
    End If

    s = "^Type|^Ticket|^FromOrg|^FromSub|^FromLoc|^ToOrg|<ToSub|^To_Loc|^Account|^SKU|^LotNum|^Units|^UOM|^ShipDate|<Comment|^EarlyDate|^PFlag"
    Grid2.FormatString = s
    Grid2.ColWidth(0) = 600
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 800
    Grid2.ColWidth(3) = 800
    Grid2.ColWidth(4) = 1000
    Grid2.ColWidth(5) = 800
    Grid2.ColWidth(6) = 800
    Grid2.ColWidth(7) = 1000
    Grid2.ColWidth(8) = 1000
    Grid2.ColWidth(9) = 800
    Grid2.ColWidth(10) = 1000  'Lot
    Grid2.ColWidth(11) = 800
    Grid2.ColWidth(12) = 800
    Grid2.ColWidth(13) = 1000
    Grid2.ColWidth(14) = 2000
    Grid2.ColWidth(15) = 1000
    Grid2.ColWidth(16) = 600
    Grid2.Redraw = True: Grid2.Visible = True
    If Grid2.Rows > 1 Then Call refresh_grid3(rid)
End Sub

Private Sub refresh_grid3(rid As String)
    Dim q As String, i As Integer, k As Integer
    Dim dsn As String, UserId As String, pwd As String
    Screen.MousePointer = 11
    dsn = "pbelle"                            'R12Test
    UserId = "Apps"                             'R12Test
    pwd = "pb3113tx"                             'R12Test

    
    If AllocateODBChEnv(hEnv) <> SQL_SUCCESS Then Exit Sub
    If ConnectToDataSource(hEnv, hdbc, hstmt, dsn, UserId, pwd) <> SQL_SUCCESS Then
        i = FreeODBChEnv(hEnv)
        Exit Sub
    End If
    
    q = "select t.subinventory_code, i.segment1, i.description, t.transaction_quantity," & _
        " t.transaction_uom, t.source_code, t.transaction_reference, t.transaction_date" & _
        " from mtl_material_transactions t, mtl_system_items_b i" & _
        " where t.shipment_number in ('" & rid & "P', '" & rid & "W')" & _
        " and i.inventory_item_id = t.inventory_item_id" & _
        " and i.organization_id = t.organization_id" & _
        " order by t.source_code, i.segment1, t.subinventory_code"
        
    If Grid2.TextMatrix(1, 8) > "0" Then    'Jobbing accounts are summarized for printing
    q = "select t.subinventory_code, i.segment1, i.description, sum(t.transaction_quantity)," & _
        " t.transaction_uom, t.source_code, t.transaction_reference, t.transaction_source_name" & _
        " from mtl_material_transactions t, mtl_system_items_b i" & _
        " where t.shipment_number in ('" & rid & "P', '" & rid & "W')" & _
        " and i.inventory_item_id = t.inventory_item_id" & _
        " and i.organization_id = t.organization_id" & _
        " group by t.subinventory_code, i.segment1, i.description, t.transaction_uom," & _
        " t.source_code, t.transaction_reference, t.transaction_source_name" & _
        " order by t.source_code, i.segment1, t.subinventory_code"
    End If
        
    'MsgBox q
    i = LoadGrid(Grid3, q, hstmt, 1, "")
    i = DisconnectFromDataSource(hdbc, hstmt)
    i = FreeODBChEnv(hEnv)
    Screen.MousePointer = 0
    Grid3.FormatString = "^SubInv|^SKU|<Product|^Qty|^UOM|^SourceCode|<Reference|^Date/Time"
    Grid3.ColWidth(0) = 800
    Grid3.ColWidth(1) = 800
    Grid3.ColWidth(2) = 2000
    Grid3.ColWidth(3) = 800
    Grid3.ColWidth(4) = 600
    Grid3.ColWidth(5) = 1800
    Grid3.ColWidth(6) = 2600
    Grid3.ColWidth(7) = 1800
    If Grid3.Rows > 2 Then Grid3.RemoveItem Grid3.Rows - 1
End Sub

Private Sub Command1_Click()
    Call refresh_grid1(Text1)
End Sub

Private Sub Form_Load()
    Text1 = Format(Now, "M-dd-yyyy")
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    Grid2.Width = Me.Width - 80
    Grid3.Width = Me.Width - 80
    pgrid.Width = Me.Width - 80
End Sub

Private Sub Grid1_Click()
    Call refresh_grid2(Grid1.TextMatrix(Grid1.Row, 4))
End Sub

Private Sub prtjob_Click()
    Dim rt As String, rh As String, rf As String
    Dim i As Integer, s As String
    If Grid3.Rows < 2 Then Exit Sub
    If Grid2.TextMatrix(1, 8) < "0" Then Exit Sub
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 4
    rt = "Jobbing Account: " & Grid2.TextMatrix(Grid2.Row, 5) & "-"
    rt = rt & Grid2.TextMatrix(Grid2.Row, 8)
    rh = Grid2.TextMatrix(Grid2.Row, 14)
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    For i = 1 To Grid3.Rows - 1
        If Grid3.TextMatrix(i, 5) = "RCV" Then
            s = Grid3.TextMatrix(i, 1)
            s = s & Chr(9) & Grid3.TextMatrix(i, 2)
            s = s & Chr(9) & Grid3.TextMatrix(i, 3)
            s = s & Chr(9) & Grid3.TextMatrix(i, 4)
            pgrid.AddItem s
        End If
    Next i
    s = "^SKU|<Product|^Quantity|^UOM"
    pgrid.FormatString = s
    pgrid.ColWidth(0) = 1000
    pgrid.ColWidth(1) = 3000
    pgrid.ColWidth(2) = 1000
    pgrid.ColWidth(3) = 1000
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, pgrid, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, "c:\htmltemp.htm", pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe c:\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe c:\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
End Sub
