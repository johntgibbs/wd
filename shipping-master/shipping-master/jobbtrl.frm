VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form jobtotrl 
   Caption         =   "Post Jobbing Orders To Trailers"
   ClientHeight    =   7380
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   12465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   7380
   ScaleWidth      =   12465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Post to Cranes (1-3)"
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
      Left            =   8520
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ListBox wnames 
      Height          =   2205
      Left            =   10320
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Post to Pick Tasks"
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
      Left            =   9240
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2175
      Left            =   0
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3836
      _Version        =   327680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Mail"
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
      Left            =   5760
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print Pick List"
      Enabled         =   0   'False
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
      Left            =   5160
      TabIndex        =   19
      Top             =   120
      Width           =   1815
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
      Left            =   7200
      TabIndex        =   17
      Text            =   "Combo3"
      Top             =   480
      Width           =   1095
   End
   Begin VB.ListBox wc 
      Height          =   840
      Left            =   9360
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ListBox pc 
      Height          =   840
      Left            =   8040
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   8040
      TabIndex        =   13
      Text            =   "Text4"
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8040
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8040
      TabIndex        =   11
      Text            =   "Text2"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete SKU"
      Height          =   255
      Left            =   8040
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add SKU"
      Height          =   255
      Left            =   8040
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   8040
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4215
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7435
      _Version        =   327680
      Cols            =   8
      FixedCols       =   3
      BackColorFixed  =   16777152
      FocusRect       =   0
      HighLight       =   2
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
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   120
      Width           =   1935
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
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label tgcode 
      Alignment       =   2  'Center
      Caption         =   "tgcode"
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
      Left            =   9240
      TabIndex        =   25
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label ycolor 
      BackColor       =   &H0000FFFF&
      Caption         =   "ycolor"
      Height          =   255
      Left            =   9600
      TabIndex        =   22
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label jtrl 
      Caption         =   ".."
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
      Left            =   8520
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label jbob 
      Caption         =   "jbob"
      Height          =   255
      Left            =   8160
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label jrun 
      Alignment       =   2  'Center
      Caption         =   "0"
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
      Left            =   7200
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label jacct 
      Caption         =   "000000"
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
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label jbranch 
      Caption         =   "0"
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
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Account:"
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
      Top             =   480
      Width           =   855
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu addsku 
         Caption         =   "Add SKU"
      End
      Begin VB.Menu delsku 
         Caption         =   "Delete SKU"
      End
      Begin VB.Menu edwhs 
         Caption         =   "Edit Warehouse"
      End
   End
End
Attribute VB_Name = "jobtotrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edcell As String

Private Sub saveorder()
    Dim cfile As String, i As Integer
    If Grid1.Rows < 2 Then Exit Sub
    If jacct <= "......" Then Exit Sub
    If Form1.plantno = "50" Then
        cfile = Form1.srserv & "\wd\jobbing\jo" & jacct & Format(Combo2, "mmddyy") & ".txt"
    Else
        cfile = Form1.srserv & "\f\user\waredist\data\jobbing\jo" & jacct & Format(Combo2, "mmddyy") & ".txt"
    End If
    Open cfile For Output As #1
    For i = 1 To Grid1.Rows - 1
        Write #1, Form1.plantno;                'Plant
        Write #1, Me.jbranch;                   'branch
        Write #1, Me.jacct;                     'account
        Write #1, Grid1.TextMatrix(i, 1);       'sku
        Write #1, Grid1.TextMatrix(i, 5);       'total units
        Write #1, Grid1.TextMatrix(i, 3);       'pallets
        Write #1, pc.List(i - 1);               'palsize
        Write #1, Grid1.TextMatrix(i, 4);       'wraps
        Write #1, wc.List(i - 1);               'wrapsize
        Write #1, "0";                          'units
        Write #1, "0";                          'net
        Write #1, "jobADD";                     'group
        'Write #1, Format(Now, "MM-dd-yyyy");    'order date
        Write #1, Me.jtrl                       'trailer #
        Write #1, Format(Combo2, "MM-dd-yyyy")  'ship date
    Next i
    Close #1
End Sub

Function whsdesc(wno As String) As String                           'jv122115
    Dim i As Integer, s As String
    wno = Format(Val(wno), "00")
    s = "Undefined"
    For i = 0 To wnames.ListCount - 1
        If wno = Left(wnames.List(i), 2) Then
            s = Right(wnames.List(i), Len(wnames.List(i)) - 3)
            Exit For
        End If
    Next i
    whsdesc = s
End Function

Private Sub post_sae(tno As String)                     'jv091510
    Dim cfile As String, i As Integer, k As Integer
    Dim ds As adodb.Recordset, s As String
    Dim bno As Integer
    Dim p As ptask, palid As String, zid As Long
    'On Error GoTo vberror
    Screen.MousePointer = 11
    bno = Val(List1)
    s = "select id,status,userid from picktasks where branch = " & Left(List1, 2)
    s = s & " and shipdate = '" & Trim(Left(Combo2, 10)) & "'"
    s = s & " and palnum = " & tno
    palid = Format(bno, "000") & " "
    palid = palid & Left(Combo2, 2) & mid(Combo2, 4, 2) & mid(Combo2, 9, 2) & " B "
    palid = palid & tno
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "Update picktasks set status = 'COMP', userid = ' ' where id = " & ds!id
            Wdb.Execute s
            ds.MoveNext
        Loop
    End If
    ds.Close
    For i = 1 To Grid1.Rows - 1
        If Val(Grid1.TextMatrix(i, 5)) > 0 Then
            s = "select * from picktasks where status in ('SHIPPED', 'COMP') order by id"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                s = "update picktasks set branch = " & bno
                s = s & ", brname = '" & Combo1 & "'"
                s = s & ", shipdate = '" & Trim(Left(Combo2, 10)) & "'"
                s = s & ", palnum = '" & tno & "'"
                s = s & ", opseq = 1" '& Val(Grid1.TextMatrix(i, 1))
                s = s & ", sku = '" & Grid1.TextMatrix(i, 1) & "'"
                s = s & ", lotnum = '...'"
                s = s & ", qty = " & Val(Grid1.TextMatrix(i, 4))
                s = s & ", uom = 'Wraps'"
                s = s & ", units = " & Val(Grid1.TextMatrix(i, 5)) 'Format(k * Val(Grid1.TextMatrix(i, 4)), "0")
                s = s & ", palletid = '" & palid & "'"
                s = s & ", status = 'PEND'"
                s = s & ", userid = '.'"
                If Combo3 = "RW" Then                           'jv032916
                    s = s & ", location = 'RE WORK'"            'jv032916
                Else                                            'jv032916
                    s = s & ", location = 'ORDER PICK'"
                End If
                s = s & ", reqid = '.'"
                s = s & " Where id = " & ds!id
                Wdb.Execute s
            Else
                zid = wd_seq("PickTasks", Form1.bbsr)
                s = "INSERT INTO PickTasks (ID, Branch, BrName, ShipDate, PalNum, OPSeq,"
                s = s & " SKU, LotNum, Qty, Uom, Units, PalletID, Status, UserID, Location,"
                s = s & " ReqID) VALUES (" & zid & ","
                s = s & bno & ","
                s = s & "'" & Combo1 & "',"
                s = s & "'" & Trim(Left(Combo2, 10)) & "',"
                s = s & tno & ","
                s = s & "1, " 'Val(Grid1.TextMatrix(i, 1)) & ","
                s = s & "'" & Grid1.TextMatrix(i, 1) & "',"
                s = s & "'...',"
                s = s & Val(Grid1.TextMatrix(i, 4)) & ","
                s = s & "'Wraps',"
                s = s & Val(Grid1.TextMatrix(i, 5)) & "," ' Format(k * Val(Grid1.TextMatrix(i, 4)), "0") & ","
                s = s & "'" & palid & "',"
                s = s & "'PEND',"
                s = s & "'.',"
                If Combo3 = "RW" Then                           'jv032916
                    s = s & "'RE WORK',"                        'jv032916
                Else                                            'jv032916
                    s = s & "'ORDER PICK',"
                End If
                s = s & "'.')"
                Wdb.Execute s
            End If
            ds.Close
        End If
    Next i
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "post_sae", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " post_sae - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_grid()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    wc.Clear: pc.Clear
    Grid1.Rows = 1: Grid1.Clear: jrun = ""
    If IsDate(Combo2) = False Then
        MsgBox "Invalid Date Format"
        Exit Sub
    End If
    Form1.cdate = Format(Combo2, "m-d-yyyy")
    Screen.MousePointer = 11
    sqlx = "select id,runid,trailers.sku,fgunit,fgdesc,pallets,wraps,units,pallet,numwrap,trlno,trailers.whs_num"
    sqlx = sqlx & " from trailers,skumast"
    sqlx = sqlx & " where shipdate = '" & Form1.cdate & "'"                 'jv091415
    sqlx = sqlx & " and branch = " & jbranch
    sqlx = sqlx & " and account = '" & jacct & "'"
    If jbob = "bob" Then sqlx = sqlx & " and trlno = '" & jtrl & "'"
    sqlx = sqlx & " and trailers.sku = skumast.sku"
    sqlx = sqlx & " order by trailers.sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = ds!id & Chr(9)
            sqlx = sqlx & ds!sku & Chr(9)
            sqlx = sqlx & ds!fgunit & " " & ds!fgdesc & Chr(9)
            sqlx = sqlx & Format(ds!pallets, "######") & Chr(9)
            sqlx = sqlx & Format(ds!wraps, "######") & Chr(9)
            sqlx = sqlx & Format(ds!units, "######") & Chr(9)
            sqlx = sqlx & Format(ds!whs_num, "0") & Chr(9)                      'jv122115
            sqlx = sqlx & whsdesc(Format(ds!whs_num, "00"))                     'jv122115
            Grid1.AddItem sqlx
            pc.AddItem ds!pallet
            wc.AddItem ds!numwrap
            jrun = ds!runid
            jtrl = ds!trlno
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "ID|^SKU|<Product|^Pallets|^Wraps|^Units|^Whs|^Warehouse"
    Grid1.ColWidth(0) = 1: Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 4500: Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000: Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 1000: Grid1.ColWidth(7) = 1400                                            'jv122115
    Call Grid1_RowColChange
    Screen.MousePointer = 0
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
Private Sub update_trl()
    Dim ds As adodb.Recordset, sqlx As String
    Dim punits As Integer
    On Error GoTo vberror
    sqlx = "select * from trailers where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Grid1.Text = Val(Grid1.Text)
        If Val(Grid1.Text) = 0 Then Grid1.Text = ""
        If edcell = "pallets" Then
            sqlx = "Update trailers set pallets = " & Val(Grid1.Text) & ", units = " & Val(Grid1.TextMatrix(Grid1.Row, 5))
        End If
        If edcell = "wraps" Then
            sqlx = "Update trailers set wraps = " & Val(Grid1.Text) & ", units = " & Val(Grid1.TextMatrix(Grid1.Row, 5))
        End If
        If edcell = "units" Then
            sqlx = "Update trailers set units = " & Val(Grid1.Text)
        End If
        sqlx = sqlx & " Where id = " & ds!id
        Sdb.Execute sqlx
        jtrl = ds!trlno
    End If
    ds.Close
    edcell = ""
    Call saveorder
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "update_trl", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " update_trl - Error Number: " & eno
        End
    End If
End Sub

Private Sub addsku_Click()
    Command1_Click
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
    If Right(Combo1, 5) = "FedEx" Then                      'jv121815
        For i = 0 To Combo3.ListCount - 1                   'jv121815
            If Combo3.List(i) = "FX" Then                   'jv121815
                Combo3.ListIndex = i                        'jv121815
                Exit For                                    'jv121815
            End If                                          'jv121815
        Next i                                              'jv121815
    End If                                                  'jv121815
    If Right(Combo1, 6) = "Parlor" Then                     'jv121815
        For i = 0 To Combo3.ListCount - 1                   'jv121815
            If Left(Combo3.List(i), 1) = "P" Then           'jv121815
                Combo3.ListIndex = i                        'jv121815
                Exit For                                    'jv121815
            End If                                          'jv121815
        Next i                                              'jv121815
    End If                                                  'jv121815
    If Right(Combo1, 7) = "Removal" Then                    'jv121815
        For i = 0 To Combo3.ListCount - 1                   'jv121815
            If Left(Combo3.List(i), 1) = "Q" Then           'jv121815
                Combo3.ListIndex = i                        'jv121815
                Exit For                                    'jv121815
            End If                                          'jv121815
        Next i                                              'jv121815
    End If                                                  'jv121815
    If Right(Combo1, 7) = "Bobtail" Then                    'jv121815
        For i = 0 To Combo3.ListCount - 1                   'jv121815
            If Left(Combo3.List(i), 1) = "B" Then           'jv121815
                Combo3.ListIndex = i                        'jv121815
                Exit For                                    'jv121815
            End If                                          'jv121815
        Next i                                              'jv121815
    End If                                                  'jv121815
    
    If Right(Combo1, 16) = "Order Pick Lanes" Then                      'jv022516
        For i = 0 To Combo3.ListCount - 1                               'jv022516
            If Combo3.List(i) = "OP" Then                               'jv022516
                Combo3.ListIndex = i                                    'jv022516
                Exit For                                                'jv022516
            End If                                                      'jv022516
        Next i                                                          'jv022516
    End If                                                              'jv022516
    
    If Right(Combo1, 17) = "ReWork Pick Tasks" Then                      'jv032916
        For i = 0 To Combo3.ListCount - 1                   'jv032916
            If Combo3.List(i) = "RW" Then                   'jv032916
                Combo3.ListIndex = i                        'jv032916
                Exit For                                    'jv032916
            End If                                          'jv032916
        Next i                                              'jv032916
    End If                                                  'jv032916
    
    
    Call Combo3_Click                                                   'jv022516
End Sub

Private Sub Combo2_Change()
    If IsDate(Combo2) Then Call refresh_grid
End Sub

Private Sub Combo2_Click()
    Call refresh_grid
    DoEvents                                                            'jv022516
    Call Combo3_Click                                                   'jv022516
End Sub

Private Sub Combo3_Click()
    jtrl = Combo3
    Call refresh_grid
    Command5.Visible = False                                                            'jv121815
    Command6.Visible = False                                            'jv022516
    If Combo3 = "FX" Or Left(Combo3, 1) = "P" Or Combo3 = "RW" Then Command5.Visible = True      'jv032916
    If Left(Combo3, 1) = "B" Then Command5.Visible = True               'jv050416
    If Combo3 = "OP" Or Combo3 = "QC" Then Command6.Visible = True      'jv022516
    tgcode = "..."                                                      'jv022516
    s = Left(Combo3, 1)                                                 'jv022516
    If s = "F" Or s = "P" Or s = "Q" Or s = "O" Or s = "R" Then                    'jv032916
        tgcode = Combo3 & "-" & Format(Combo2, "dd")                    'jv022516
    Else                                                                'jv022516
        If s = "B" Then                                                 'jv022516
            tgcode = Combo3 & "-" & Format(Val(List1), "00")            'jv022516
        Else                                                            'jv022516
            tgcode = LCase(jbob) & "ADD"                                'jv022516
        End If                                                          'jv022516
    End If                                                              'jv022516
End Sub

Private Sub Command1_Click()                    'Add SKU
    Dim ds As adodb.Recordset, sqlx As String, msku As String
    Dim i As Integer, mtno As String, pkey As Long, mwhs As String
    'On Error GoTo vberror
    If Len(edcell) > 0 Then Call update_trl
    If IsDate(Combo2) = False Then
        MsgBox "Invalid Date Format", vbOKOnly + vbExclamation, "Cannot insert..."
        Exit Sub
    End If
    If jbob = "bob" Then
        If Val(jbranch) = 0 Then
            MsgBox "Branch not specified..", vbOKOnly + vbExclamation, "Cannot insert.."
            Exit Sub
        End If
        If Val(jrun) = 0 Then
            mtno = Left(jtrl, 2)
        End If
    Else
        If Val(jbranch) = 0 Or jacct <= "0" Then
            MsgBox "Account not specified..", vbOKOnly + vbExclamation, "Cannot insert.."
            Exit Sub
        End If
        If Val(jrun) = 0 Then
            mtno = InputBox("Trailer Code:", "Trailer #", "#1")
            If Len(mtno) = 0 Then Exit Sub
            mtno = Left(mtno, 2)
            jtrl = mtno
        End If
    End If
    msku = InputBox$("Please enter SKU for Product", "New Product", "777")
    If Len(msku) = 0 Then Exit Sub
    sqlx = "select * from skumast where sku = '" & msku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "SKU number not found in list..", vbOKOnly, "Invalid SKU"
        ds.Close
        Exit Sub
    End If
    ds.Close
    If Form1.plantno = "50" Then                                                'jv121815
        If Left(Combo3, 1) = "Q" Then                                           'jv121815
            mwhs = InputBox("Warehouse (1-5):", "Select Warehouse....", "5")    'jv121815
            If Len(mwhs) = 0 Then Exit Sub                                      'jv121815
            If Val(mwhs) < 1 Or Val(mwhs) > 5 Then Exit Sub                     'jv121815
        Else                                                                    'jv121815
            mwhs = "10"                                                         'jv121815
        End If                                                                  'jv121815
        If Combo3 = "OP" Then mwhs = "1"             'jv022516
        If Combo3 = "RW" Then mwhs = "4"            'jv032916
    Else                                                                        'jv121815
        If Form1.plantno = "51" Then                                            'jv121815
            mwhs = "14"                                                         'jv121815
        Else                                                                    'jv121815
            If Form1.plantno = "52" Then                                        'jv121815
                If Left(mtno, 1) = "Q" Then                                     'jv121815
                    mwhs = "1"                                                  'jv121815
                Else                                                            'jv121815
                    mwhs = "15"                                                 'jv121815
                End If                                                          'jv121815
            Else                                                                'jv121815
                mwhs = "0"                                                      'jv121815
            End If                                                              'jv121815
        End If                                                                  'jv121815
    End If                                                                      'jv121815
    
    If Val(jrun) = 0 Then
        pkey = wd_seq("Oratkt", Form1.schdb)
        sqlx = "Insert into runs (id, loaded, destination, locname, trlno, trlsize, trldate, startime"
        sqlx = sqlx & ", pickup, oc) Values (" & pkey
        sqlx = sqlx & ", " & Form1.plantno
        sqlx = sqlx & ", " & jbranch
        If Len(Combo1) > 30 Then
            sqlx = sqlx & ", '" & Left(Combo1, 30) & "'"
        Else
            sqlx = sqlx & ", '" & Combo1 & "'"
        End If
        sqlx = sqlx & ", '" & jtrl & "'"
        sqlx = sqlx & ", 0"
        sqlx = sqlx & ", '" & Combo2 & "'"
        sqlx = sqlx & ", '12:00 PM'"
        If jbob = "bob" Then
            If Left(jtrl, 1) = "B" Then                                     'jv121815
                sqlx = sqlx & ", 'Added Bobtail'"                           'jv121815
            Else                                                            'jv121815
                If jtrl = "FX" Then                                         'jv121815
                    sqlx = sqlx & ", 'FedEx'"                               'jv121815
                Else                                                        'jv121815
                    If jtrl = "QC" Then                                     'jv121815
                        sqlx = sqlx & ", 'QC Removal'"                      'jv121815
                    Else                                                    'jv121815
                        If Left(jtrl, 1) = "P" Then                         'jv121815
                            sqlx = sqlx & ", 'Parlor Request'"              'jv121815
                        Else                                                'jv121815
                            If jtrl = "OP" Then                                             'jv022516
                                sqlx = sqlx & ", 'Order Pick Lanes'"                        'jv022516
                            Else
                                If jtrl = "RW" Then                             'jv032916
                                    sqlx = sqlx & ", 'ReWork Pick Tasks'"       'jv032916
                                Else
                                    sqlx = sqlx & ", '" & Form1.userid & "'"    'jv121815
                                End If
                            End If                                                          'jv022516
                        End If                                              'jv121815
                    End If
                End If
            End If
        Else
            sqlx = sqlx & ", 'Added for jobbing..'"
        End If
        sqlx = sqlx & ", '*')"
        Sdb.Execute sqlx
        jrun = pkey
    End If
    sqlx = "select * from trailers where runid = " & jrun
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        mtno = ds!trlno
    End If
    pkey = wd_seq("trailers", Form1.shipdb)
    sqlx = "Insert into trailers (id, runid, groupcode, plant, branch, account, shipdate, trlno, sku"
    sqlx = sqlx & ", pallets, wraps, units, whs_num, pb_flag, ra_flag) Values (" & pkey     'jv121815
    sqlx = sqlx & ", " & jrun
    If mtno = "FX" Or Left(mtno, 1) = "P" Or mtno = "QC" Or mtno = "OP" Or mtno = "RW" Then   'jv032916
        sqlx = sqlx & ", '" & mtno & "-" & Format(Combo2, "dd") & "'"           'jv121815
    Else                                                                        'jv121815
        If Left(mtno, 1) = "B" Then                                             'jv121815
            sqlx = sqlx & ", '" & mtno & "-" & Format(List1, "00") & "'"        'jv121815
        Else                                                                    'jv121815
            sqlx = sqlx & ", '" & LCase(jbob) & "ADD'"
        End If
    End If
    sqlx = sqlx & ", " & Form1.plantno
    sqlx = sqlx & ", " & jbranch
    sqlx = sqlx & ", '" & jacct & "'"
    sqlx = sqlx & ", '" & Combo2 & "'"
    sqlx = sqlx & ", '" & mtno & "'"
    sqlx = sqlx & ", '" & msku & "'"
    sqlx = sqlx & ", 0, 0, 0, " & mwhs & ", 'N', 'N')"                          'jv121815
    Sdb.Execute sqlx
    ds.Close
    Call refresh_grid
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 1) = msku Then
            Grid1.Row = i
            If i > (Grid1.Height / 245) Then Grid1.TopRow = i
            Exit For
        End If
    Next i
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command1_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command1_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command2_Click()                'Delete SKU
    Dim sqlx As String
    On Error GoTo vberror
    If Len(edcell) > 0 Then Call update_trl
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    sqlx = "Ok to drop " & Grid1.TextMatrix(Grid1.Row, 2) & " from trailer?"
    If MsgBox(sqlx, vbYesNo + vbQuestion, "Drop Product") = vbNo Then Exit Sub
    sqlx = "delete from trailers where id = " & Grid1.TextMatrix(Grid1.Row, 0)
    Sdb.Execute sqlx
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
        wc.RemoveItem wc.ListIndex
        pc.RemoveItem pc.ListIndex
        Call Grid1_RowColChange
    Else
        Call refresh_grid
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command2_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command2_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command3_Click()            'Print Pick List
    Dim sqlx As String, i As Integer, mdate As String, mbr As String
    Dim ds As adodb.Recordset, s As String, c As Integer
    Dim tp As Long, tw As Long, tu As Long, m As Integer
    Dim rt As String, rh As String, rf As String, rc As Integer
    On Error GoTo vberror
    If Val(jrun) = 0 Then Exit Sub
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 10
    tp = 0: tw = 0: tu = 0
    s = "select trailers.sku,trlno,pallets,wraps,units,opseq,fgunit,fgdesc"
    s = s & " from trailers,oplist,skumast"
    s = s & " where trailers.runid = " & Val(jrun)
    s = s & " and oplist.sku = trailers.sku"
    s = s & " and skumast.sku = trailers.sku"
    s = s & " order by opseq,trailers.sku"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        c = 1
        m = CInt(Grid1.Rows / 2) + 1
        rc = 0
        Do Until ds.EOF
            rc = rc + 1
            If rc > m Then c = 2
            If jtrl = "FX" Or Left(jtrl, 1) = "P" Or jtrl = "RW" Then c = 1                     'jv032916
            If c = 1 Then
                s = ds(0) & Chr(9)
                s = s & ds!fgunit & " " & ds!fgdesc & Chr(9)
                s = s & Format(ds!pallets, "#") & Chr(9)
                s = s & Format(ds!wraps, "#") & Chr(9)
                s = s & Format(ds!units, "#") & Chr(9)
                Grid2.AddItem s
                If jtrl = "FX" Or Left(jtrl, 1) = "P" Or jtrl = "RW" Then Grid2.AddItem " "         'jv032916
                Grid2.Row = 1
            Else
                Grid2.TextMatrix(Grid2.Row, 5) = ds(0)
                Grid2.TextMatrix(Grid2.Row, 6) = ds!fgunit & " " & ds!fgdesc
                Grid2.TextMatrix(Grid2.Row, 7) = Format(ds!pallets, "#")
                Grid2.TextMatrix(Grid2.Row, 8) = Format(ds!wraps, "#")
                Grid2.TextMatrix(Grid2.Row, 9) = Format(ds!units, "#")
                Grid2.Row = Grid2.Row + 1
            End If
            tp = tp + ds!pallets
            tw = tw + ds!wraps
            tu = tu + ds!units
            ds.MoveNext
        Loop
    End If
    Grid2.AddItem " "
    s = Chr(9) & Totals & Chr(9)
    s = s & Format(tp, "#") & Chr(9)
    s = s & Format(tw, "#") & Chr(9)
    s = s & Format(tu, "#")
    Grid2.AddItem s
    ds.Close
    If jtrl = "FX" Or Left(jtrl, 1) = "P" Or jtrl = "RW" Then                           'jv032916
        Grid2.FormatString = "^SKU|<Product|^Pallets|^Wraps|^Units|^Code Dates||||"     'jv121815
        Grid2.ColWidth(0) = 500                                                         'jv121815
        Grid2.ColWidth(1) = 3000                                                        'jv121815
        Grid2.ColWidth(2) = 700                                                         'jv121815
        Grid2.ColWidth(3) = 700                                                         'jv121815
        Grid2.ColWidth(4) = 700                                                         'jv121815
        Grid2.ColWidth(5) = 5000                                                        'jv121815
        Grid2.ColWidth(6) = 0 '3000                                                     'jv121815
        Grid2.ColWidth(7) = 0 '700                                                      'jv121815
        Grid2.ColWidth(8) = 0 '700                                                      'jv121815
        Grid2.ColWidth(9) = 0 '700                                                      'jv121815
    Else                                                                                'jv121815
        Grid2.FormatString = "^SKU|<Product|^Pallets|^Wraps|^Units|^SKU|<Product|^Pallets|^Wraps|^Units|"
        Grid2.ColWidth(0) = 500
        Grid2.ColWidth(1) = 3000
        Grid2.ColWidth(2) = 700
        Grid2.ColWidth(3) = 700
        Grid2.ColWidth(4) = 700
        Grid2.ColWidth(5) = 500
        Grid2.ColWidth(6) = 3000
        Grid2.ColWidth(7) = 700
        Grid2.ColWidth(8) = 700
        Grid2.ColWidth(9) = 700
    End If
    rt = "Order Pick - " & Combo1 & " " & Combo3
    rh = "Ship Date: " & Combo2
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Call printflexgrid(Printer, Grid2, rt, rh, rf)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "command3_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command3_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command4_Click()                'Create html for e-mail
    Dim i As Integer, tot As Long
    Dim rt As String, rh As String, rf As String, hf As String
    If Grid1.Rows < 2 Then Exit Sub
    tot = 0
    For i = 0 To Grid1.Rows - 1
        tot = tot + Val(Grid1.TextMatrix(i, 5))
    Next i
    
    rt = "Account: " & jbranch & "-" & jacct & " " & Combo1
    rh = "Ship Date: " & Combo2
    rf = "Total Units: " & tot
    hf = Form1.tempdir & "\jobord.htm"
    htdc(0) = "Yellow": gndc(0) = ycolor.BackColor
    Call htmlcolorgrid(Me, hf, Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe " & hf, vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & hf, vbNormalFocus)
        Exit Sub
    End If
End Sub

Private Sub Command5_Click()                'Post to Pick Tasks
    Dim t As String
    If Combo3 = "FX" Then                   'jv121815
        t = "700"                           'jv121815
    Else                                    'jv121815
        If Combo3 = "RW" Then                           'jv032916
            t = "600"                                   'jv032916
        Else                                            'jv032916
            t = "90" & Val(Right(Combo3, 1))         'jv050516
        End If
    End If                                  'jv121815
    Call post_sae(t)                        'jv121815
End Sub

Private Sub Command6_Click()                'Post to Cranes (1-3)
    Dim db As adodb.Connection, ds As adodb.Recordset, db2 As adodb.Connection, ds2 As adodb.Recordset
    Dim sqlx As String, pgrp As String, pvert As Integer, phorz As Integer
    Dim pside As String, opgrp As Boolean, pkey As Long
    On Error GoTo vberror
    'pgrp = UCase(InputBox$("Input Shipping Group to be posted...", "Shipping Group", Form1.cgrp))
    pgrp = Me.tgcode
    If Len(pgrp) = 0 Then Exit Sub
    opgrp = False
    If InStr(1, pgrp, "OP") > 0 Then
        If MsgBox("Do you wish to assign Order Pick Lanes?", vbYesNo + vbQuestion, "OP Group...") = vbYes Then opgrp = True
    End If
    Screen.MousePointer = 11
    sqlx = "Update ship_infc set ship_status = 'CANC' where order_num = '" & pgrp & "'"
    Wdb.Execute sqlx
    sqlx = "Update drop_infc set drop_qty = 0 where group_num = '" & pgrp & "'"
    Wdb.Execute sqlx
    sqlx = "select sku,shipdate,whs_num,sum(pallets) from trailers where groupcode = '" & pgrp & "'"
    sqlx = sqlx & " And pallets > 0 And sku <> 'PAR' And ra_flag = 'N' And pb_flag = 'N'"
    sqlx = sqlx & " And whs_num in (Select whs_num From Warehouses"
    sqlx = sqlx & " Where whs in ('SR1','SR2','SR3'))"
    sqlx = sqlx & " Group by sku,shipdate,whs_num"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            pvert = 0
            If opgrp = True Then
                sqlx = "select * from opbays where whse_num = " & ds!whs_num
                sqlx = sqlx & " and sku = '" & ds!sku & "'"
                Set ds2 = Wdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    pvert = ds2!vert_loc
                    phorz = ds2!horz_loc
                    pside = ds2!rack_side
                End If
                ds2.Close
            End If
            If pvert = 0 Then
                sqlx = "select * from sr_config where whs_num = " & ds!whs_num
                Set ds2 = Wdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    pvert = ds2!ship1_lane_vert
                    phorz = ds2!ship1_lane_horz
                    pside = ds2!ship1_lane_side
                End If
                ds2.Close
            End If
            sqlx = "select * from ship_infc where ship_status = 'CANC'"
            sqlx = sqlx & " or ship_status = 'DONE' order by id"
            Set ds2 = Wdb.Execute(sqlx)
            If ds2.BOF = False Then
                sqlx = "Update ship_infc set order_num = '" & pgrp & "'"
                sqlx = sqlx & ", sku = '" & ds!sku & "'"
                sqlx = sqlx & ", lot_num = ' '"
                sqlx = sqlx & ", ship_date = '" & ds!shipdate & "'"
                sqlx = sqlx & ", order_qty = " & ds(3)
                sqlx = sqlx & ", ship_uom_qty = 0"
                sqlx = sqlx & ", ship_plt_qty = 0"
                sqlx = sqlx & ", ship_status = 'NEW'"
                sqlx = sqlx & ", to_whse_num = " & ds!whs_num
                sqlx = sqlx & ", to_vert_loc = " & pvert
                sqlx = sqlx & ", to_horz_loc = " & phorz
                sqlx = sqlx & ", to_rack_side = '" & pside & "'"
                sqlx = sqlx & ", resv_strategy = 'A'"
                sqlx = sqlx & " Where id = " & ds2!id
                Wdb.Execute sqlx
            Else
                pkey = wd_seq("ship_infc", Form1.bbsr)
                sqlx = "Insert into ship_infc (id, order_num, sku, lot_num, ship_date, order_qty, ship_uom_qty"
                sqlx = sqlx & ", ship_plt_qty, ship_status, to_whse_num, to_vert_loc, to_horz_loc, to_rack_side"
                sqlx = sqlx & ", resv_strategy) Values (" & pkey
                sqlx = sqlx & ", '" & pgrp & "'"
                sqlx = sqlx & ", '" & ds!sku & "'"
                sqlx = sqlx & ", ' '"
                sqlx = sqlx & ", '" & ds!shipdate & "'"
                sqlx = sqlx & ", " & ds(3)
                sqlx = sqlx & ", 0, 0, 'NEW'"
                sqlx = sqlx & ", " & ds!whs_num
                sqlx = sqlx & ", " & pvert
                sqlx = sqlx & ", " & phorz
                sqlx = sqlx & ", '" & pside & "'"
                sqlx = sqlx & ", 'A')"
                Wdb.Execute sqlx
            End If
            ds2.Close
            ds.MoveNext
        Loop
    Else
        MsgBox "There were no SR products found to post for this group..", vbOKOnly, "Group " & pgrp
    End If
    ds.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "nt_ship", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " nt_ship - Error Number: " & eno
        End
    End If
End Sub

Private Sub delsku_Click()
    Command2_Click
End Sub

Private Sub edwhs_Click()                                           'jv122115
    Dim s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) = 0 Then Exit Sub
    If Form1.plantno <> "50" Then Exit Sub
    s = Grid1.TextMatrix(Grid1.Row, 6)
    s = InputBox("SR (1-5):", "Edit warehouse...", s)
    If Len(s) = 0 Then Exit Sub
    If Val(s) >= 1 And Val(s) <= 15 Then
        s = "Update trailers set whs_num = " & s & " where id = " & Grid1.TextMatrix(Grid1.Row, 0)
        Sdb.Execute s
        refresh_grid
    End If
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Len(edcell) > 0 Then Call update_trl
    If jobtotrl.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
            If Form1.FrmGrid.Text = "jobtotrl" Then
                Form1.FrmGrid.Col = 1: Form1.FrmGrid.Text = jobtotrl.Top
                Form1.FrmGrid.Col = 2: Form1.FrmGrid.Text = jobtotrl.Left
                Form1.FrmGrid.Col = 3: Form1.FrmGrid.Text = jobtotrl.Height
                Form1.FrmGrid.Col = 4: Form1.FrmGrid.Text = jobtotrl.Width
                Exit For
            End If
        Next i
    End If
    If Form1.WindowState = 1 Then End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Or KeyCode = 121 Then Call Command1_Click
    If KeyCode = 46 Or KeyCode = 120 Then Call Command2_Click
End Sub

Private Sub Form_Load()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    Combo1.Clear: Combo2.Clear: Combo3.Clear: List1.Clear
    Combo3.AddItem "BO"
    Combo3.AddItem "B1"
    Combo3.AddItem "B2"
    Combo3.AddItem "B3"
    Combo3.AddItem "B4"
    Combo3.AddItem "B5"
    Combo3.AddItem "B6"
    Combo3.AddItem "B7"
    Combo3.AddItem "B8"
    Combo3.AddItem "B9"
    If Form1.plantno = "50" Then Combo3.AddItem "FX"        'jv121815
    If Form1.plantno = "50" Then Combo3.AddItem "OP"                    'jv022516
    If Form1.plantno = "50" Then Combo3.AddItem "RW"                    'jv032916
    Combo3.AddItem "P1"                                     'jv121815
    Combo3.AddItem "P2"                                     'jv121815
    Combo3.AddItem "P3"                                     'jv121815
    Combo3.AddItem "QC"                                     'jv121815
    Dim i As Integer
    For i = 1 To Form1.FrmGrid.Rows - 1
        Form1.FrmGrid.Col = 0: Form1.FrmGrid.Row = i
        If Form1.FrmGrid.Text = "jobtotrl" Then
            Form1.FrmGrid.Col = 1: jobtotrl.Top = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 2: jobtotrl.Left = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 3: jobtotrl.Height = Val(Form1.FrmGrid.Text)
            Form1.FrmGrid.Col = 4: jobtotrl.Width = Val(Form1.FrmGrid.Text)
            Exit For
        End If
    Next i
    
    If jbob = "bob" Then                                                    'jv121815
        If Form1.plantno = "50" Then                                        'jv121815
            List1.AddItem "12": Combo1.AddItem "Brenham FedEx"               'jv012016
            List1.AddItem "1": Combo1.AddItem "Brenham Parlor"              'jv121815
            List1.AddItem "1": Combo1.AddItem "Brenham QC Removal"          'jv121815
            List1.AddItem "1": Combo1.AddItem "Brenham ReWork Pick Tasks"    'jv032916
            List1.AddItem "1": Combo1.AddItem "Brenham Order Pick Lanes"    'jv022516
        End If                                                              'jv121815
        If Form1.plantno = "51" Then                                        'jv121815
            List1.AddItem "47": Combo1.AddItem "Broken Arrow Parlor"        'jv121815
        End If                                                              'jv121815
        If Form1.plantno = "52" Then                                        'jv121815
            List1.AddItem "52": Combo1.AddItem "Sylacauga Parlor"           'jv121815
            List1.AddItem "52": Combo1.AddItem "Sylacauga QC Removal"       'jv121815
        End If                                                              'jv121815
    End If                                                                  'jv121815
    
    If jbob = "bob" Then
        'sqlx = "select branch,branchname & ' Bobtail' from branches where branch < 90"
        'sqlx = "select branch,branchname from branches where branch <= 90"
        sqlx = "select branch,branchname from branches where branch = 47 or gemmsid in "        'jv050516
        sqlx = sqlx & "(select listreturn from valuelists where listname = 'branchplants')"     'jv050516
        sqlx = sqlx & " order by branchname"
        Combo3.Visible = True
    Else
        sqlx = "select id,acctdesc from jobbing order by acctdesc"
        Combo3.Visible = False
    End If
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            List1.AddItem ds(0) 'ds!id
            If jbob = "bob" Then
                Combo1.AddItem ds(1) & " Bobtail"
            Else
                Combo1.AddItem ds(1) 'ds!acctdesc
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    sqlx = "Select distinct shipdate from trailers"
    'sqlx = sqlx & " where account <> '......'"
    sqlx = sqlx & " order by shipdate"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo2.AddItem Format(ds(0), "MM-dd-yyyy")                      'jv122115
            ds.MoveNext
        Loop
    End If
    ds.Close
    wnames.Clear                                                            'jv122115
    wnames.AddItem "00" & " " & "CS5"                                       'jv122115
    s = "select * from warehouses order by whs_num"                         'jv122115
    Set ds = Sdb.Execute(s)                                                  'jv122115
    If ds.BOF = False Then                                                  'jv122115
        ds.MoveFirst                                                        'jv122115
        Do Until ds.EOF                                                     'jv122115
            wnames.AddItem Format(ds!whs_num, "00") & " " & ds!whsname      'jv122115
            ds.MoveNext                                                     'jv122115
        Loop                                                                'jv122115
    End If                                                                  'jv122115
    
    ds.Close
    Combo2.AddItem Format(Now, "MM-dd-yyyy")                  'jv091415
    For i = 1 To 6                                                          'jv022516
        Combo2.AddItem Format(DateAdd("d", i, Now), "MM-dd-yyyy")           'jv022516
    Next i                                                                  'jv022516
    Combo2.ListIndex = 0
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
    If Me.Height > 3000 Then Grid1.Height = Me.Height - 1620
    Grid1.Width = Me.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_GotFocus()
    Grid1.FocusRect = flexFocusNone
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
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
    If Grid1.Col < 3 Then Exit Sub
    If Len(edcell) = 0 Then
        If Grid1.Col = 3 Then Grid1.Text = ""
        If Grid1.Col = 4 Then Grid1.Text = ""
        If Grid1.Col = 5 Then Grid1.Text = ""
    End If
    If Grid1.Col = 3 Then edcell = "pallets"
    If Grid1.Col = 4 Then edcell = "wraps"
    If Grid1.Col = 5 Then edcell = "units"
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
    If edcell = "pallets" Or edcell = "wraps" Then
        Grid1.TextMatrix(Grid1.Row, 5) = (Val(Grid1.TextMatrix(Grid1.Row, 3)) * Val(pc)) + (Val(Grid1.TextMatrix(Grid1.Row, 4)) * Val(wc))
    End If
End Sub
Private Sub Grid1_LeaveCell()
    If Len(edcell) > 0 Then Call update_trl
End Sub

Private Sub Grid1_LostFocus()
    If Len(edcell) > 0 Then Call update_trl
    Grid1.FocusRect = flexFocusLight
End Sub


Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid1_RowColChange()
    If wc.ListCount > 0 Then
        wc.ListIndex = Grid1.Row - 1
        pc.ListIndex = Grid1.Row - 1
    End If
End Sub

Private Sub jbob_Change()
    If jbob = "bob" Then
        jobtotrl.Caption = "Bobtail, Parlor, FedEx, QC Removal"
        Command3.Visible = True
        Command4.Visible = False
    Else
        jobtotrl.Caption = "Post Jobbing Order To Trailers"
        Command3.Visible = False
        Command4.Visible = True
    End If
    Call Form_Load
End Sub

Private Sub jrun_Change()
    If Val(jrun) > 0 Then
        If jbob = "bob" Then
            Command3.Enabled = True
        Else
            Command3.Enabled = False
        End If
    Else
        Command3.Enabled = False
    End If
End Sub

Private Sub List1_Click()
    Dim ds As adodb.Recordset, sqlx As String
    On Error GoTo vberror
    If List1.ListCount < 1 Then Exit Sub
    If jbob = "bob" Then
        jbranch = List1
        jacct = "......"
        Combo3.ListIndex = 0
    Else
        sqlx = "select branch,account from jobbing where id = " & List1
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            jbranch = ds!branch
            jacct = ds!account
        End If
        ds.Close
        jtrl = ".."
        Call refresh_grid
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "list1_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " list1_click - Error Number: " & eno
        End
    End If
End Sub
