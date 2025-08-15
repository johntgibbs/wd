VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PostBrGrps 
   Caption         =   "Post Branches To Groups"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13770
   LinkTopic       =   "Form2"
   ScaleHeight     =   6585
   ScaleWidth      =   13770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Reverse Group"
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
      Left            =   960
      TabIndex        =   27
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Group"
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
      Left            =   4320
      TabIndex        =   26
      Top             =   4920
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1095
      Left            =   0
      TabIndex        =   25
      Top             =   2160
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   1931
      _Version        =   327680
      BackColor       =   16777215
      ForeColor       =   4194368
      BackColorFixed  =   8454143
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Group Information "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin MSFlexGridLib.MSFlexGrid Grid2 
         Height          =   1815
         Left            =   7080
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3201
         _Version        =   327680
         Rows            =   5
         ForeColor       =   16384
         BackColorFixed  =   14737632
         BackColorSel    =   32768
         FocusRect       =   0
      End
      Begin VB.ComboBox Combo6 
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
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox Combo5 
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
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   1335
      End
      Begin VB.ListBox brcode 
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   24
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox brcode 
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   23
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox brcode 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   22
         Top             =   2280
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox brcode 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   21
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ListBox trun 
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   20
         Top             =   2280
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox trun 
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   19
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox trun 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   18
         Top             =   2280
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox trun 
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   17
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox tsize 
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   4560
         TabIndex        =   16
         Top             =   2280
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox tsize 
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox tsize 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   14
         Top             =   2280
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox tsize 
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox Combo4 
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
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1200
         Width           =   2895
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
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   840
         Width           =   2895
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
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   2895
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
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
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
         Left            =   4200
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Post Group"
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
         Left            =   960
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text2 
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
         Left            =   5640
         MaxLength       =   6
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "4."
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
         Index           =   3
         Left            =   3720
         TabIndex        =   6
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "3."
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
         Index           =   2
         Left            =   3720
         TabIndex        =   5
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "2."
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
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "1."
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
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Group Code:"
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
         Left            =   4320
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "PostBrGrps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mqty1 As Integer, mqty2 As Integer, mqty3 As Integer, mqty4 As Integer
Dim mwhs1 As Integer, mwhs2 As Integer, mwhs3 As Integer, mwhs4 As Integer
Dim m1t As Integer, m2t As Integer, m3t As Integer, m4t As Integer
Dim s1t As Integer, s2t As Integer, s3t As Integer, s5t As Integer

Private Function split_hmv(mqty As Integer, part As Integer) As Integer
    Dim p1 As Integer, p2 As Integer
    If mqty = 34 Then
        p1 = 18: p2 = 16
    End If
    If mqty = 33 Then
        p1 = 17: p2 = 16
    End If
    If mqty = 32 Then
        p1 = 16: p2 = 16
    End If
    If mqty = 31 Then
        p1 = 15: p2 = 16
    End If
    If mqty = 30 Then
        p1 = 14: p2 = 16
    End If
    If mqty = 29 Then
        p1 = 15: p2 = 14
    End If
    If mqty = 28 Then
        p1 = 14: p2 = 14
    End If
    If mqty = 27 Then
        p1 = 13: p2 = 14
    End If
    If mqty = 26 Then
        p1 = 12: p2 = 14
    End If
    If mqty = 25 Then
        p1 = 13: p2 = 12
    End If
    If mqty = 24 Then
        p1 = 12: p2 = 12
    End If
    If mqty = 23 Then
        p1 = 11: p2 = 12
    End If
    If mqty = 22 Then
        p1 = 10: p2 = 12
    End If
    If mqty = 21 Then
        p1 = 11: p2 = 10
    End If
    If mqty = 20 Then
        p1 = 10: p2 = 10
    End If
    If mqty = 19 Then
        p1 = 9: p2 = 10
    End If
    If mqty = 18 Then
        p1 = 8: p2 = 10
    End If
    If mqty = 17 Then
        p1 = 9: p2 = 8
    End If
    If mqty = 16 Then
        p1 = 8: p2 = 8
    End If
    If mqty = 15 Then
        p1 = 7: p2 = 8
    End If
    If mqty = 14 Then
        p1 = 6: p2 = 8
    End If
    If mqty = 13 Then
        p1 = 7: p2 = 6
    End If
    If mqty = 12 Then
        p1 = 6: p2 = 6
    End If
    If mqty = 11 Then
        p1 = 5: p2 = 6
    End If
    If mqty = 10 Then
        p1 = 4: p2 = 6
    End If
    If mqty = 9 Then
        p1 = 5: p2 = 4
    End If
    If mqty = 8 Then
        p1 = 4: p2 = 4
    End If
    If mqty = 7 Then
        p1 = 3: p2 = 4
    End If
    If mqty < 7 Then
        p1 = mqty: p2 = 0
    End If
    If mqty > 34 Then
        p1 = mqty: p2 = 0
    End If
    If part = 1 Then
        split_hmv = p1
    Else
        split_hmv = p2
    End If
End Function

Sub refresh_grid()
    Dim ds As adodb.Recordset, ds2 As adodb.Recordset
    Dim t1 As String, t2 As String, t3 As String, t4 As String
    Dim p1 As String, p2 As String, p3 As String, p4 As String, hflag As Boolean
    'On Error GoTo vberror
    Grid1.Clear: Grid1.Cols = 9: Grid1.Rows = 1
    Set ds = Sdb.Execute("select * from trgroups order by groupcode")
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            p1 = " ": p2 = " ": p3 = " ": p4 = " "
            t1 = " ": t2 = " ": t3 = " ": t4 = " "
            If ds!run1 > 0 Then
                sqlx = "Select Loaded,Destination,Locname,trlno,trldate from runs where id = " & ds!run1
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    p1 = Format$(ds2!loaded, "00") & Format$(ds2!Destination, "00") & Format$(ds2!trldate, "mm-dd-yyyy")
                    t1 = ds2!locname & " " & ds2!trlno
                End If
                ds2.Close
            End If
            If ds!run2 > 0 Then
                sqlx = "Select Loaded,Destination,Locname,trlno,trldate from runs where id = " & ds!run2
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    p2 = Format$(ds2!loaded, "00") & Format$(ds2!Destination, "00") & Format$(ds2!trldate, "mm-dd-yyyy")
                    t2 = ds2!locname & " " & ds2!trlno
                End If
                ds2.Close
            End If
            If ds!run3 > 0 Then
                sqlx = "Select Loaded,Destination,Locname,trlno,trldate from runs where id = " & ds!run3
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    p3 = Format$(ds2!loaded, "00") & Format$(ds2!Destination, "00") & Format$(ds2!trldate, "mm-dd-yyyy")
                    t3 = ds2!locname & " " & ds2!trlno
                End If
                ds2.Close
            End If
            If ds!run4 > 0 Then
                sqlx = "Select Loaded,Destination,Locname,trlno,trldate from runs where id = " & ds!run4
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    p4 = Format$(ds2!loaded, "00") & Format$(ds2!Destination, "00") & Format$(ds2!trldate, "mm-dd-yyyy")
                    t4 = ds2!locname & " " & ds2!trlno
                End If
                ds2.Close
            End If
            sqlx = ds!groupcode & Chr$(9) & t1 & Chr$(9) & t2 & Chr$(9) & t3 & Chr$(9) & t4 & Chr$(9)
            sqlx = sqlx & p1 & Chr$(9) & p2 & Chr$(9) & p3 & Chr$(9) & p4
            Grid1.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.FormatString = "^Group|^#1|^#2|^#3|^#4|^Key 1|^Key 2|^Key 3|^Key 4"
    Grid1.ColWidth(0) = 900
    Grid1.ColWidth(1) = 2000: Grid1.ColWidth(2) = 2000
    Grid1.ColWidth(3) = 2000: Grid1.ColWidth(4) = 2000
    Grid1.ColWidth(5) = 0: Grid1.ColWidth(6) = 0
    Grid1.ColWidth(7) = 0: Grid1.ColWidth(8) = 0
    'Grid1.ColWidth(5) = 1500: Grid1.ColWidth(6) = 1500
    'Grid1.ColWidth(7) = 1500: Grid1.ColWidth(8) = 1500
    
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If hflag Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = Grid1.BackColorFixed
            End If
            hflag = Not hflag
        Next i
        Grid1.Row = 1
    End If
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

Sub refresh_grid2()                                                     'jv061815
    Dim s As String, ds As adodb.Recordset
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 3
    s = "select whs_num, whs, vert_loc from warehouses where plant = 50"
    s = s & " and whs in ('SR1', 'SR2', 'SR3', 'SR5', 'Reg', 'RegA') order by whs_num"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!whs_num & Chr(9) & ds!whs & Chr(9)
            If ds!vert_loc = 0 Then
                s = s & "No"
            Else
                s = s & "Yes"
            End If
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid2.FillStyle = flexFillRepeat
    For i = 1 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 2) = "No" Then
            Grid2.Row = i: Grid2.RowSel = i
            Grid2.Col = 2: Grid2.ColSel = 2
            Grid2.CellBackColor = Grid1.BackColorFixed
        End If
    Next i
    Grid2.Col = 1
    Grid2.FormatString = "^ID|^Warehouse|^On Line"
    Grid2.ColWidth(0) = 500
    Grid2.ColWidth(1) = 1200
    Grid2.ColWidth(2) = 1200
End Sub

Sub pltmain(xwhs As Integer)
    Dim sqlx As String, ds As adodb.Recordset, msku As String
    'On Error GoTo vberror
    sqlx = "select * from groupitems where groupcode = '" & Text2 & "'"
    sqlx = sqlx & " and sku in (select sku from whstotals where whs_num = " & xwhs & ")"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                If IsNull(ds!qty1) Then
                    mqty1 = 0
                Else
                    mqty1 = ds!qty1
                End If
                If IsNull(ds!qty2) Then
                    mqty2 = 0
                Else
                    mqty2 = ds!qty2
                End If
                If IsNull(ds!qty3) Then
                    mqty3 = 0
                Else
                    mqty3 = ds!qty3
                End If
                If IsNull(ds!qty4) Then
                    mqty4 = 0
                Else
                    mqty4 = ds!qty4
                End If
                mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                Call pltsub(xwhs, ds!sku)
                sqlx = "Update groupitems set whs1 = " & mwhs1 & ", whs2 = " & mwhs2
                sqlx = sqlx & ", whs3 = " & mwhs3 & ", whs4 = " & mwhs4
                sqlx = sqlx & " Where id = " & ds!id
                Sdb.Execute sqlx
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "pltmain(" & xwhs & ")", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " pltmain - Error Number: " & eno
        End
    End If
End Sub
Sub pltsub(xwhs As Integer, msku As String)
    Dim bs As adodb.Recordset, sqlx As String, mavail As Integer
    Dim bo As adodb.Recordset, i As Integer
    'On Error GoTo vberror
    sqlx = "select outstk,avail from plantskus,whstotals"
    sqlx = sqlx & " where plantskus.plant = " & Left$(Combo6, 2)
    sqlx = sqlx & " and whstotals.whs_num = " & xwhs
    sqlx = sqlx & " and plantskus.sku = '" & msku & "'"
    sqlx = sqlx & " and plantskus.sku = whstotals.sku"
    Set bs = Sdb.Execute(sqlx)
    If bs.BOF Then
        mavail = 0
    Else
        bs.MoveFirst
        mavail = bs!avail - bs!outstk
    End If
    bs.Close
    If mqty1 > 0 And mqty1 <= mavail And (m1t + mqty1) <= Val(tsize(0)) Then
        mwhs1 = xwhs: m1t = m1t + mqty1: mavail = mavail - mqty1
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(0))
        sqlx = sqlx & " and plant = " & Left$(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty1
            sqlx = sqlx & ", netqty = netqty - " & mqty1
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty1 = 0
    End If
    If mqty2 > 0 And mqty2 <= mavail And (m2t + mqty2) <= Val(tsize(1)) Then
        mwhs2 = xwhs: m2t = m2t + mqty2: mavail = mavail - mqty2
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(1))
        sqlx = sqlx & " and plant = " & Left$(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty2
            sqlx = sqlx & ", netqty = netqty - " & mqty2
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty2 = 0
    End If
    If mqty3 > 0 And mqty3 <= mavail And (m3t + mqty3) <= Val(tsize(2)) Then
        mwhs3 = xwhs: m3t = m3t + mqty3: mavail = mavail - mqty3
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(2))
        sqlx = sqlx & " and plant = " & Left$(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty3
            sqlx = sqlx & ", netqty = netqty - " & mqty3
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty3 = 0
    End If
    If mqty4 > 0 And mqty4 <= mavail And (m4t + mqty4) <= Val(tsize(3)) Then
        mwhs4 = xwhs: m4t = m4t + mqty4: mavail = mavail - mqty4
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(3))
        sqlx = sqlx & " and plant = " & Left$(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty4
            sqlx = sqlx & ", netqty = netqty - " & mqty4
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty4 = 0
    End If
    If (mqty1 + mqty2 + mqty3 + mqty4) > 0 Then
        sqlx = "select * from whstotals"
        sqlx = sqlx & " where whs_num = " & xwhs & " and sku = '" & msku & "'"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            i = mqty1 + mqty2 + mqty3 + mqty4
            sqlx = "Update whstotals set grp_qty = grp_qty + " & i
            sqlx = sqlx & ", avail = avail - " & i
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "pltsub(" & xwhs & ", " & msku & ")", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " pltsub - Error Number: " & eno
        End
    End If
End Sub
Sub bamain()
    Dim sqlx As String, ds As adodb.Recordset, msku As String
    'On Error GoTo vberror
    sqlx = "select id,sku,qty1,whs1,qty2,whs2,qty3,whs3,qty4,whs4 from groupitems"
    sqlx = sqlx & " where groupcode = '" & Text2 & "'"
    sqlx = sqlx & " and sku in (select sku from whstotals where whs_num = (select whs_num from warehouses where whs = 'BA'))"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                If IsNull(ds!qty1) Then
                    mqty1 = 0
                Else
                    mqty1 = ds!qty1
                End If
                If IsNull(ds!qty2) Then
                    mqty2 = 0
                Else
                    mqty2 = ds!qty2
                End If
                If IsNull(ds!qty3) Then
                    mqty3 = 0
                Else
                    mqty3 = ds!qty3
                End If
                If IsNull(ds!qty4) Then
                    mqty4 = 0
                Else
                    mqty4 = ds!qty4
                End If
                mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                msku = ds!sku
                Call basub(msku)
                sqlx = "Update groupitems set whs1 = " & mwhs1
                sqlx = sqlx & ", whs2 = " & mwhs2
                sqlx = sqlx & ", whs3 = " & mwhs3
                sqlx = sqlx & ", whs4 = " & mwhs4
                sqlx = sqlx & " where id = " & ds!id
                Sdb.Execute sqlx
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "bamain", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " bamain - Error Number: " & eno
        End
    End If
End Sub

Sub symain()
    Dim sqlx As String, ds As adodb.Recordset, msku As String
    'On Error GoTo vberror
    sqlx = "select id,sku,qty1,whs1,qty2,whs2,qty3,whs3,qty4,whs4 from groupitems"
    sqlx = sqlx & " where groupcode = '" & Text2 & "'"
    sqlx = sqlx & " and sku in (select sku from whstotals where whs_num = (select whs_num from waredhouses where whs = 'SY'))"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                If IsNull(ds!qty1) Then
                    mqty1 = 0
                Else
                    mqty1 = ds!qty1
                End If
                If IsNull(ds!qty2) Then
                    mqty2 = 0
                Else
                    mqty2 = ds!qty2
                End If
                If IsNull(ds!qty3) Then
                    mqty3 = 0
                Else
                    mqty3 = ds!qty3
                End If
                If IsNull(ds!qty4) Then
                    mqty4 = 0
                Else
                    mqty4 = ds!qty4
                End If
                mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                msku = ds!sku
                Call sysub(msku)
                sqlx = "Update groupitems set whs1 = " & mwhs1
                sqlx = sqlx & ", whs2 = " & mwhs2
                sqlx = sqlx & ", whs3 = " & mwhs3
                sqlx = sqlx & ", whs4 = " & mwhs4
                sqlx = sqlx & " where id = " & ds!id
                Sdb.Execute sqlx
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "symain", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " symain - Error Number: " & eno
        End
    End If
End Sub
Sub sysub(msku As String)
    Dim bs As adodb.Recordset, sqlx As String, mavail As Integer
    Dim sywhs As Integer, syplant As Integer, bo As adodb.Recordset, i As Integer
    'On Error GoTo vberror
    sqlx = "select whs_num,plant from warehouses where whs = 'SY')"
    Set bs = Sdb.Execute(sqlx)
    If bs.BOF = False Then
        bs.MoveFirst
        sywhs = bs!whs_num
        syplant = bs!plant
    End If
    bs.Close
    
    sqlx = "select outstk, avail from plantskus, whstotals "
    sqlx = sqlx & " where plantskus.plant = " & syplant & " and plantskus.sku = '" & msku & "'"
    sqlx = sqlx & " and plantskus.sku = whstotals.sku"
    
    Set bs = Sdb.Execute(sqlx)
    If bs.BOF Then
        mavail = 0
    Else
        bs.MoveFirst
        mavail = bs!avail - bs!outstk
    End If
    bs.Close
    If mqty1 > 0 And mqty1 <= mavail And (m1t + mqty1) <= Val(tsize(0)) Then
        mwhs1 = sywhs: m1t = m1t + mqty1: mavail = mavail - mqty1
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(0))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty1
            sqlx = sqlx & ", netqty = netqty - " & mqty1
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty1 = 0
    End If
    If mqty2 > 0 And mqty2 <= mavail And (m2t + mqty2) <= Val(tsize(1)) Then
        mwhs2 = sywhs: m2t = m2t + mqty2: mavail = mavail - mqty2
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(1))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty2
            sqlx = sqlx & ", netqty = netqty - " & mqty2
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty2 = 0
    End If
    If mqty3 > 0 And mqty3 <= mavail And (m3t + mqty3) <= Val(tsize(2)) Then
        mwhs3 = sywhs: m3t = m3t + mqty3: mavail = mavail - mqty3
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(2))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty3
            sqlx = sqlx & ", netqty = netqty - " & mqty3
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty3 = 0
    End If
    If mqty4 > 0 And mqty4 <= mavail And (m4t + mqty4) <= Val(tsize(3)) Then
        mwhs4 = sywhs: m4t = m4t + mqty4: mavail = mavail - mqty4
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(3))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty4
            sqlx = sqlx & ", netqty = netqty - " & mqty4
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty4 = 0
    End If
    If (mqty1 + mqty2 + mqty3 + mqty4) > 0 Then
        sqlx = "select * from whstotals"
        sqlx = sqlx & " where whs_num = " & sywhs & " and sku = '" & msku & "'"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            i = mqty1 + mqty2 + mqty3 + mqty4
            sqlx = "Update whstotals set grp_qty = grp_qty + " & i
            sqlx = sqlx & ", avail = avail - " & i
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "sysub(" & msku & ")", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " sysub - Error Number: " & eno
        End
    End If
End Sub
Sub basub(msku As String)
    Dim bs As adodb.Recordset, sqlx As String, mavail As Integer
    Dim bawhs As Integer, baplant As Integer, bo As adodb.Recordset, i As Integer
    'On Error GoTo vberror
    Set bs = Sdb.Execute(sqlx)
    If bs.BOF = False Then
        bs.MoveFirst
        bawhs = bs!whs_num
        baplant = bs!plant
    End If
    bs.Close
    
    sqlx = "select outstk, avail from plantskus, whstotals "
    sqlx = sqlx & " where plantskus.plant = " & baplant & " and plantskus.sku = '" & msku & "'"
    sqlx = sqlx & " and plantskus.sku = whstotals.sku"
    Set bs = Sdb.Execute(sqlx)
    If bs.BOF Then
        mavail = 0
    Else
        bs.MoveFirst
        mavail = bs!avail - bs!outstk
    End If
    bs.Close
    If mqty1 > 0 And mqty1 <= mavail And (m1t + mqty1) <= Val(tsize(0)) Then
        mwhs1 = bawhs: m1t = m1t + mqty1: mavail = mavail - mqty1
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(0))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty1
            sqlx = sqlx & ", netqty = netqty - " & mqty1
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty1 = 0
    End If
    If mqty2 > 0 And mqty2 <= mavail And (m2t + mqty2) <= Val(tsize(1)) Then
        mwhs2 = bawhs: m2t = m2t + mqty2: mavail = mavail - mqty2
        sqlx = sqlx & " where branch = " & Val(brcode(1))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty2
            sqlx = sqlx & ", netqty = netqty - " & mqty2
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty2 = 0
    End If
    If mqty3 > 0 And mqty3 <= mavail And (m3t + mqty3) <= Val(tsize(2)) Then
        mwhs3 = bawhs: m3t = m3t + mqty3: mavail = mavail - mqty3
        sqlx = sqlx & " where branch = " & Val(brcode(2))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty3
            sqlx = sqlx & ", netqty = netqty - " & mqty3
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty3 = 0
    End If
    If mqty4 > 0 And mqty4 <= mavail And (m4t + mqty4) <= Val(tsize(3)) Then
        mwhs4 = bawhs: m4t = m4t + mqty4: mavail = mavail - mqty4
        sqlx = sqlx & " where branch = " & Val(brcode(3))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty4
            sqlx = sqlx & ", netqty = netqty - " & mqty4
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    Else
        mqty4 = 0
    End If
    If (mqty1 + mqty2 + mqty3 + mqty4) > 0 Then
        sqlx = "select * from whstotals"
        sqlx = sqlx & " where whs_num = " & bawhs & " and sku = '" & msku & "'"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            i = mqty1 + mqty2 + mqty3 + mqty4
            sqlx = "Update whstotals set grp_qty = grp_qty + " & i
            sqlx = sqlx & ", avail = avail - " & i
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "basub(" & msku & ")", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " basub - Error Number: " & eno
        End
    End If
End Sub

Sub cranesub(msku As String)
    Dim bs As adodb.Recordset, sqlx As String, ws As adodb.Recordset
    Dim s1 As Integer, s2 As Integer, s3 As Integer, s5 As Integer, bo As adodb.Recordset
    Dim xr1 As Integer, xr2 As Integer, xr3 As Integer, xr5 As Integer
    Dim mwhs As Integer, mqty As Integer
    'On Error GoTo vberror
    mwhs = 0: mqty = mqty1 + mqty2 + mqty3 + mqty4
    s1 = 0: s2 = 0: s3 = 0: s5 = 0
    sqlx = "select * from warehouses where whs in ('SR1','SR2','SR3','SR5')"
    sqlx = sqlx & " and vert_loc > 0"                                           'jv061815
    Set bs = Sdb.Execute(sqlx)
    If bs.BOF = False Then
        bs.MoveFirst
        Do Until bs.EOF
            If bs!whs = "SR1" Then xr1 = bs!whs_num
            If bs!whs = "SR2" Then xr2 = bs!whs_num
            If bs!whs = "SR3" Then xr3 = bs!whs_num
            If bs!whs = "SR5" Then xr5 = bs!whs_num
            bs.MoveNext
        Loop
    End If
    bs.Close
    sqlx = "select whs_num,sku,grp_qty,avail from whstotals where sku = '" & msku & "'"
    Set bs = Sdb.Execute(sqlx)
    If bs.BOF = False Then
        bs.MoveFirst
        Do Until bs.EOF
            If bs!whs_num = xr1 Then s1 = bs!avail
            If bs!whs_num = xr2 Then s2 = bs!avail
            If bs!whs_num = xr3 Then s3 = bs!avail
            If bs!whs_num = xr5 Then s5 = bs!avail
            bs.MoveNext
        Loop
    End If
    bs.Close
    If (s1 + s2 + s3 + s5) = 0 Then GoTo nextrec
    'If (mqty Mod 2) = 1 Then
    '    If (s1 Mod 2) = 1 And s1 >= mqty Then
    '        mwhs = xr1: GoTo foundwhs
    '    Else
    '        If (s2 Mod 2) = 1 And s2 >= mqty Then
    '            mwhs = xr2: GoTo foundwhs
    '        Else
    '            If (s3 Mod 2) = 1 And s3 >= mqty Then
    '                mwhs = xr3: GoTo foundwhs
    '            Else
    '                If (s5 Mod 2) = 1 And s5 >= mqty Then
    '                    mwhs = xr5: GoTo foundwhs
    '                End If
    '            End If
    '        End If
    '    End If
    'End If
    If (mqty Mod 2) = 1 Then
        If (s5 Mod 2) = 1 And s5 >= mqty Then
            mwhs = xr5: GoTo foundwhs
        Else
            If (s2 Mod 2) = 1 And s2 >= mqty Then
                mwhs = xr2: GoTo foundwhs
            Else
                If (s3 Mod 2) = 1 And s3 >= mqty Then
                    mwhs = xr3: GoTo foundwhs
                Else
                    If (s1 Mod 2) = 1 And s1 >= mqty Then
                        mwhs = xr1: GoTo foundwhs
                    End If
                End If
            End If
        End If
    End If
    'If s5t < s3t And s5t < s2t And s5t < s1t And s5 >= mqty Then
    If s5t < (s3t + s2t + s1t) And s5 >= mqty Then
        mwhs = xr5
    Else
        If s3t < s2t And s3t < s1t And s3 >= mqty Then
            mwhs = xr3
        Else
            If s2t < s1t And s2 >= mqty Then
                mwhs = xr2
            Else
                If s1 >= mqty Then
                    mwhs = xr1
                Else
                    If s2 >= mqty Then
                        mwhs = xr2
                    Else
                        If s3 >= mqty Then
                            mwhs = xr3
                        Else
                            If s5 >= mqty Then
                                mwhs = xr5
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
foundwhs:
    mqty = 0
    If mqty4 > 0 And (m4t + mqty4) <= Val(tsize(3)) And mwhs <> 0 Then
        m4t = m4t + mqty4: mwhs4 = mwhs
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(3))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty4
            sqlx = sqlx & ", netqty = netqty - " & mqty4
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
        mqty = mqty + mqty4
    End If
    If mqty3 > 0 And (m3t + mqty3) <= Val(tsize(2)) And mwhs <> 0 Then
        m3t = m3t + mqty3: mwhs3 = mwhs
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(2))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty3
            sqlx = sqlx & ", netqty = netqty - " & mqty3
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
        mqty = mqty + mqty3
    End If
    If mqty2 > 0 And (m2t + mqty2) <= Val(tsize(1)) And mwhs <> 0 Then
        m2t = m2t + mqty2: mwhs2 = mwhs
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(1))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty2
            sqlx = sqlx & ", netqty = netqty - " & mqty2
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
        mqty = mqty + mqty2
    End If
    If mqty1 > 0 And (m1t + mqty1) <= Val(tsize(0)) And mwhs <> 0 Then
        m1t = m1t + mqty1: mwhs1 = mwhs
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(0))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty1
            sqlx = sqlx & ", netqty = netqty - " & mqty1
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
        mqty = mqty + mqty1
    End If
    If mwhs = xr1 Then
        s1t = s1t + mqty
        sqlx = "Update whstotals set grp_qty = grp_qty + " & mqty
        sqlx = sqlx & ", avail = avail - " & mqty
        sqlx = sqlx & " Where whs_num = " & xr1 & " and sku = '" & msku & "'"
        Sdb.Execute sqlx
    End If
    If mwhs = xr2 Then
        s2t = s2t + mqty
        sqlx = "Update whstotals set grp_qty = grp_qty + " & mqty
        sqlx = sqlx & ", avail = avail - " & mqty
        sqlx = sqlx & " Where whs_num = " & xr2 & " and sku = '" & msku & "'"
        Sdb.Execute sqlx
    End If
    If mwhs = xr3 Then
        s3t = s3t + mqty
        sqlx = "Update whstotals set grp_qty = grp_qty + " & mqty
        sqlx = sqlx & ", avail = avail - " & mqty
        sqlx = sqlx & " Where whs_num = " & xr3 & " and sku = '" & msku & "'"
        Sdb.Execute sqlx
    End If
    If mwhs = xr5 Then
        s5t = s5t + mqty
        sqlx = "Update whstotals set grp_qty = grp_qty + " & mqty
        sqlx = sqlx & ", avail = avail - " & mqty
        sqlx = sqlx & " Where whs_num = " & xr5 & " and sku = '" & msku & "'"
        Sdb.Execute sqlx
    End If
nextrec:
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "cranesub(" & msku & ")", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " cranesub - Error Number: " & eno
        End
    End If
End Sub
Sub cranmain()
    Dim sqlx As String, ds As adodb.Recordset, msku As String, ws As adodb.Recordset
    'On Error GoTo vberror
    s1t = 0: s2t = 0: s3t = 0: s5t = 0
    Set db = CreateObject("ADODB.Connection")
    db.Open Form1.shipdb
    ' Process Drops
    sqlx = "select * from warehouses where plant = 50 and whs = 'DROP'"
    Set ws = Sdb.Execute(sqlx)
    If ws.BOF = False Then
        ws.MoveFirst
        sqlx = "select * from groupitems where groupcode = '" & Text2 & "'"
        sqlx = sqlx & " and sku in (select sku from whstotals where whs_num = " & ws!whs_num & ")"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF Or Err = 3021
                If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                    If IsNull(ds!qty1) Then
                        mqty1 = 0
                    Else
                        mqty1 = ds!qty1
                    End If
                    If IsNull(ds!qty2) Then
                        mqty2 = 0
                    Else
                        mqty2 = ds!qty2
                    End If
                    If IsNull(ds!qty3) Then
                        mqty3 = 0
                    Else
                        mqty3 = ds!qty3
                    End If
                    If IsNull(ds!qty4) Then
                        mqty4 = 0
                    Else
                        mqty4 = ds!qty4
                    End If
                    mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                    msku = ds!sku
                    Call dropsub(ws!whs_num, msku)
                    sqlx = "Update groupitems set whs1 = " & mwhs1
                    sqlx = sqlx & ", whs2 = " & mwhs2
                    sqlx = sqlx & ", whs3 = " & mwhs3
                    sqlx = sqlx & ", whs4 = " & mwhs4
                    sqlx = sqlx & " Where id = " & ds!id
                    Sdb.Execute sqlx
                End If
                ds.MoveNext
            Loop
        End If
    End If
    ws.Close
    ds.Close
    ' Snack Plant Drops
    sqlx = "select * from warehouses where plant = 50 and whs = 'SDRP'"
    Set ws = Sdb.Execute(sqlx)
    If ws.BOF = False Then
        ws.MoveFirst
        sqlx = "select * from groupitems where groupcode = '" & Text2 & "'"
        sqlx = sqlx & " and sku in (select sku from whstotals where whs_num = " & ws!whs_num & ")"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF Or Err = 3021
                If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                    If IsNull(ds!qty1) Then
                        mqty1 = 0
                    Else
                        mqty1 = ds!qty1
                    End If
                    If IsNull(ds!qty2) Then
                        mqty2 = 0
                    Else
                        mqty2 = ds!qty2
                    End If
                    If IsNull(ds!qty3) Then
                        mqty3 = 0
                    Else
                        mqty3 = ds!qty3
                    End If
                    If IsNull(ds!qty4) Then
                        mqty4 = 0
                    Else
                        mqty4 = ds!qty4
                    End If
                    mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                    msku = ds!sku
                    Call dropsub(ws!whs_num, msku)
                    sqlx = "Update groupitems set whs1 = " & mwhs1
                    sqlx = sqlx & ", whs2 = " & mwhs2
                    sqlx = sqlx & ", whs3 = " & mwhs3
                    sqlx = sqlx & ", whs4 = " & mwhs4
                    sqlx = sqlx & " Where id = " & ds!id
                    Sdb.Execute sqlx
                End If
                ds.MoveNext
            Loop
        End If
    End If
    ws.Close
    ds.Close
    ' Regular Items
    sqlx = "select * from warehouses where plant = 50 and whs = 'Reg'"
    sqlx = sqlx & " and vert_loc > 0"                                       'jv061815
    Set ws = Sdb.Execute(sqlx)
    If ws.BOF = False Then
        ws.MoveFirst
        sqlx = "select * from groupitems where groupcode = '" & Text2 & "'"
        sqlx = sqlx & " and sku in (select sku from whstotals where whs_num = " & ws!whs_num
        sqlx = sqlx & " and avail > 0)"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF Or Err = 3021
                If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                    If IsNull(ds!qty1) Then
                        mqty1 = 0
                    Else
                        mqty1 = ds!qty1
                    End If
                    If IsNull(ds!qty2) Then
                        mqty2 = 0
                    Else
                        mqty2 = ds!qty2
                    End If
                    If IsNull(ds!qty3) Then
                        mqty3 = 0
                    Else
                        mqty3 = ds!qty3
                    End If
                    If IsNull(ds!qty4) Then
                        mqty4 = 0
                    Else
                        mqty4 = ds!qty4
                    End If
                    mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                    msku = ds!sku
                    Call pltsub(ws!whs_num, msku)
                    sqlx = "Update groupitems set whs1 = " & mwhs1
                    sqlx = sqlx & ", whs2 = " & mwhs2
                    sqlx = sqlx & ", whs3 = " & mwhs3
                    sqlx = sqlx & ", whs4 = " & mwhs4
                    sqlx = sqlx & " Where id = " & ds!id
                    Sdb.Execute sqlx
                End If
                ds.MoveNext
            Loop
        Else
            ds.Close
        End If
    End If
    ws.Close
    ' Ante Room
    sqlx = "select * from warehouses where plant = 50 and whs = 'ANTE'"
    Set ws = Sdb.Execute(sqlx)
    If ws.BOF = False Then
        ws.MoveFirst
        sqlx = "select * from groupitems where groupcode = '" & Text2 & "'"
        sqlx = sqlx & " and sku in (select sku from whstotals where whs_num = " & ws!whs_num
        sqlx = sqlx & " and avail > 0)"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF Or Err = 3021
                If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                    If IsNull(ds!qty1) Then
                        mqty1 = 0
                    Else
                        mqty1 = ds!qty1
                    End If
                    If IsNull(ds!qty2) Then
                        mqty2 = 0
                    Else
                        mqty2 = ds!qty2
                    End If
                    If IsNull(ds!qty3) Then
                        mqty3 = 0
                    Else
                        mqty3 = ds!qty3
                    End If
                    If IsNull(ds!qty4) Then
                        mqty4 = 0
                    Else
                        mqty4 = ds!qty4
                    End If
                    mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                    msku = ds!sku
                    Call pltsub(ws!whs_num, msku)
                    sqlx = "Update groupitems set whs1 = " & mwhs1
                    sqlx = sqlx & ", whs2 = " & mwhs2
                    sqlx = sqlx & ", whs3 = " & mwhs3
                    sqlx = sqlx & ", whs4 = " & mwhs4
                    sqlx = sqlx & " Where id = " & ds!id
                    Sdb.Execute sqlx
                End If
                ds.MoveNext
            Loop
        Else
            ds.Close
        End If
    End If
    ws.Close
    ' Snack Plant Items
    sqlx = "select * from warehouses where plant = 50 and whs = 'SP'"
    Set ws = Sdb.Execute(sqlx)
    If ws.BOF = False Then
        ws.MoveFirst
        sqlx = "select * from groupitems where groupcode = '" & Text2 & "'"
        sqlx = sqlx & " and sku in (select sku from whstotals where whs_num = " & ws!whs_num
        sqlx = sqlx & " and avail > 0)"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF Or Err = 3021
                If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                    If IsNull(ds!qty1) Then
                        mqty1 = 0
                    Else
                        mqty1 = ds!qty1
                    End If
                    If IsNull(ds!qty2) Then
                        mqty2 = 0
                    Else
                        mqty2 = ds!qty2
                    End If
                    If IsNull(ds!qty3) Then
                        mqty3 = 0
                    Else
                        mqty3 = ds!qty3
                    End If
                    If IsNull(ds!qty4) Then
                        mqty4 = 0
                    Else
                        mqty4 = ds!qty4
                    End If
                    mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                    msku = ds!sku
                    Call pltsub(ws!whs_num, msku)
                    sqlx = "Update groupitems set whs1 = " & mwhs1
                    sqlx = sqlx & ", whs2 = " & mwhs2
                    sqlx = sqlx & ", whs3 = " & mwhs3
                    sqlx = sqlx & ", whs4 = " & mwhs4
                    sqlx = sqlx & " Where id = " & ds!id
                    Sdb.Execute sqlx
                End If
                ds.MoveNext
            Loop
        Else
            ds.Close
        End If
    End If
    ws.Close
    ' Crane Rank1 Items
    sqlx = "select * from groupitems where groupcode = '" & Text2 & "'"
    sqlx = sqlx & " and grank = 1"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF Or Err = 3021
            If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                If IsNull(ds!qty1) Then
                    mqty1 = 0
                Else
                    mqty1 = ds!qty1
                End If
                If IsNull(ds!qty2) Then
                    mqty2 = 0
                Else
                    mqty2 = ds!qty2
                End If
                If IsNull(ds!qty3) Then
                    mqty3 = 0
                Else
                    mqty3 = ds!qty3
                End If
                If IsNull(ds!qty4) Then
                    mqty4 = 0
                Else
                    mqty4 = ds!qty4
                End If
                mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                msku = ds!sku
                Call cranesub(msku)
                sqlx = "Update groupitems set whs1 = " & mwhs1
                sqlx = sqlx & ", whs2 = " & mwhs2
                sqlx = sqlx & ", whs3 = " & mwhs3
                sqlx = sqlx & ", whs4 = " & mwhs4
                sqlx = sqlx & " Where id = " & ds!id
                Sdb.Execute sqlx
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    ' Old Lot Items
    sqlx = "select * from groupitems where groupcode = '" & Text2 & "'"
    sqlx = sqlx & " and sku in (select sku from whstotals where old_qty > 0)"
    sqlx = sqlx & " and grank < 3 order by grank"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF Or Err = 3021
            If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                If IsNull(ds!qty1) Then
                    mqty1 = 0
                Else
                    mqty1 = ds!qty1
                End If
                If IsNull(ds!qty2) Then
                    mqty2 = 0
                Else
                    mqty2 = ds!qty2
                End If
                If IsNull(ds!qty3) Then
                    mqty3 = 0
                Else
                    mqty3 = ds!qty3
                End If
                If IsNull(ds!qty4) Then
                    mqty4 = 0
                Else
                    mqty4 = ds!qty4
                End If
                mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                msku = ds!sku
                Call oldsub(msku)
                sqlx = "Update groupitems set whs1 = " & mwhs1
                sqlx = sqlx & ", whs2 = " & mwhs2
                sqlx = sqlx & ", whs3 = " & mwhs3
                sqlx = sqlx & ", whs4 = " & mwhs4
                sqlx = sqlx & " Where id = " & ds!id
                Sdb.Execute sqlx
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    ' Last Crane Items
    sqlx = "select * from groupitems where groupcode = '" & Text2 & "'"
    sqlx = sqlx & " order by grank"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF Or Err = 3021
            If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                If IsNull(ds!qty1) Then
                    mqty1 = 0
                Else
                    mqty1 = ds!qty1
                End If
                If IsNull(ds!qty2) Then
                    mqty2 = 0
                Else
                    mqty2 = ds!qty2
                End If
                If IsNull(ds!qty3) Then
                    mqty3 = 0
                Else
                    mqty3 = ds!qty3
                End If
                If IsNull(ds!qty4) Then
                    mqty4 = 0
                Else
                    mqty4 = ds!qty4
                End If
                mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                msku = ds!sku
                Call cranesub(msku)
                sqlx = "Update groupitems set whs1 = " & mwhs1
                sqlx = sqlx & ", whs2 = " & mwhs2
                sqlx = sqlx & ", whs3 = " & mwhs3
                sqlx = sqlx & ", whs4 = " & mwhs4
                sqlx = sqlx & " Where id = " & ds!id
                Sdb.Execute sqlx
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    ' Regular A Items
    sqlx = "select * from warehouses where plant = 50 and whs = 'RegA'"
    sqlx = sqlx & " and vert_loc > 0"                                       'jv061815
    Set ws = Sdb.Execute(sqlx)
    If ws.BOF = False Then
        ws.MoveFirst
        sqlx = "select * from groupitems where groupcode = '" & Text2 & "'"
        sqlx = sqlx & " and sku in (select sku from whstotals where whs_num = " & ws!whs_num
        sqlx = sqlx & " and avail > 0)"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF Or Err = 3021
                If ds!whs1 = 0 And ds!whs2 = 0 And ds!whs3 = 0 And ds!whs4 = 0 Then
                    If IsNull(ds!qty1) Then
                        mqty1 = 0
                    Else
                        mqty1 = ds!qty1
                    End If
                    If IsNull(ds!qty2) Then
                        mqty2 = 0
                    Else
                        mqty2 = ds!qty2
                    End If
                    If IsNull(ds!qty3) Then
                        mqty3 = 0
                    Else
                        mqty3 = ds!qty3
                    End If
                    If IsNull(ds!qty4) Then
                        mqty4 = 0
                    Else
                        mqty4 = ds!qty4
                    End If
                    mwhs1 = 0: mwhs2 = 0: mwhs3 = 0: mwhs4 = 0
                    msku = ds!sku
                    Call pltsub(ws!whs_num, msku)
                    sqlx = "Update groupitems set whs1 = " & mwhs1
                    sqlx = sqlx & ", whs2 = " & mwhs2
                    sqlx = sqlx & ", whs3 = " & mwhs3
                    sqlx = sqlx & ", whs4 = " & mwhs4
                    sqlx = sqlx & " Where id = " & ds!id
                    Sdb.Execute sqlx
                End If
                ds.MoveNext
            Loop
        Else
            ds.Close
        End If
    End If
    ws.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "cranmain", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " cranmain - Error Number: " & eno
        End
    End If
End Sub

Sub dropsub(xwhs As Integer, msku As String)
    Dim sqlx As String, bo As adodb.Recordset
    'On Error GoTo vberror
    If mqty1 > 0 And (m1t + mqty1) <= Val(tsize(0)) Then
        mwhs1 = xwhs: m1t = m1t + mqty1
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(0))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update broders set grpqty = grpqty + " & mqty1
            sqlx = sqlx & ", netqty = netqty - " & mqty1 & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    End If
    If mqty2 > 0 And (m2t + mqty2) <= Val(tsize(1)) Then
        mwhs2 = xwhs: m2t = m2t + mqty2
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(1))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update broders set grpqty = grpqty + " & mqty2
            sqlx = sqlx & ", netqty = netqty - " & mqty2 & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    End If
    If mqty3 > 0 And (m3t + mqty3) <= Val(tsize(2)) Then
        mwhs3 = xwhs: m3t = m3t + mqty3
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(2))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update broders set grpqty = grpqty + " & mqty3
            sqlx = sqlx & ", netqty = netqty - " & mqty3 & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    End If
    If mqty4 > 0 And (m4t + mqty4) <= Val(tsize(3)) Then
        mwhs4 = xwhs: m4t = m4t + mqty4
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(3))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update broders set grpqty = grpqty + " & mqty4
            sqlx = sqlx & ", netqty = netqty - " & mqty4 & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "dropsub", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " dropsub - Error Number: " & eno
        End
    End If
End Sub

Sub oldsub(msku As String)
    Dim bs As adodb.Recordset, sqlx As String, bo As adodb.Recordset
    Dim s1 As Integer, s2 As Integer, s3 As Integer, s5 As Integer
    Dim xr1 As Integer, xr2 As Integer, xr3 As Integer, xr5 As Integer
    Dim lot1 As Long, lot2 As Long, lot3 As Long, lot5 As Long
    Dim mwhs As Integer, mqty As Integer
    'On Error GoTo vberror
    mwhs = 0: mqty = mqty1 + mqty2 + mqty3 + mqty4
    s1 = 0: s2 = 0: s3 = 0: s5 = 0
    lot1 = 99999: lot2 = 99999: lot3 = 99999: lot5 = 99999
    sqlx = "select * from warehouses where whs in ('SR1','SR2','SR3','SR5')"
    sqlx = sqlx & " and vert_loc > 0"                                           'jv061815
    Set bs = Sdb.Execute(sqlx)
    If bs.BOF = False Then
        bs.MoveFirst
        Do Until bs.EOF
            If bs!whs = "SR1" Then xr1 = bs!whs_num
            If bs!whs = "SR2" Then xr2 = bs!whs_num
            If bs!whs = "SR3" Then xr3 = bs!whs_num
            If bs!whs = "SR5" Then xr5 = bs!whs_num
            bs.MoveNext
        Loop
    End If
    bs.Close
    If (s1t + mqty) <= 10 Then
        sqlx = "select whs_num,sku,grp_qty,avail,old_qty,old_lot from whstotals"
        sqlx = sqlx & " where sku = '" & msku & "' and old_lot > '0' and whs_num = " & xr1
        Set bs = Sdb.Execute(sqlx)
        If bs.BOF = False Then
            bs.MoveFirst
            s1 = bs!old_qty - bs!grp_qty
            lot1 = bs!old_lot
        End If
        bs.Close
    End If
    If (s2t + mqty) <= 12 Then
        sqlx = "select whs_num,sku,grp_qty,avail,old_qty,old_lot from whstotals"
        sqlx = sqlx & " where sku = '" & msku & "' and old_lot > '0' and whs_num = " & xr2
        Set bs = Sdb.Execute(sqlx)
        If bs.BOF = False Then
            bs.MoveFirst
            s2 = bs!old_qty - bs!grp_qty
            lot2 = bs!old_lot
        End If
        bs.Close
    End If
    If (s3t + mqty) <= 12 Then
        sqlx = "select whs_num,sku,grp_qty,avail,old_qty,old_lot from whstotals"
        sqlx = sqlx & " where sku = '" & msku & "' and old_lot > '0' and whs_num = " & xr3
        Set bs = Sdb.Execute(sqlx)
        If bs.BOF = False Then
            bs.MoveFirst
            s3 = bs!old_qty - bs!grp_qty
            lot3 = bs!old_lot
        End If
        bs.Close
    End If
    If (s5t + mqty) <= 20 Then
        sqlx = "select whs_num,sku,grp_qty,avail,old_qty,old_lot from whstotals"
        sqlx = sqlx & " where sku = '" & msku & "' and old_lot > '0' and whs_num = " & xr5
        Set bs = Sdb.Execute(sqlx)
        If bs.BOF = False Then
            bs.MoveFirst
            s5 = bs!old_qty - bs!grp_qty
            lot5 = bs!old_lot
        End If
        bs.Close
    End If
    
    
    If (s1 + s2 + s3 + s5) = 0 Then GoTo recnext
    If s5 >= mqty And lot5 <= lot1 And lot5 <= lot2 And lot5 <= lot3 Then
        mwhs = xr5: GoTo whsfound
    End If
    If s1 >= mqty And lot1 <= lot2 And lot1 <= lot3 Then
        mwhs = xr1: GoTo whsfound
    End If
    If s2 >= mqty And lot2 <= lot3 Then
        mwhs = xr2: GoTo whsfound
    End If
    If s3 >= mqty Then
        mwhs = xr3: GoTo whsfound
    End If
    If s2 >= mqty Then
        mwhs = xr2: GoTo whsfound
    End If
    If s1 >= mqty Then
        mwhs = xr1: GoTo whsfound
    End If
    If s5 >= mqty Then
        mwhs = xr5: GoTo whsfound
    End If
    
whsfound:
    mqty = 0
    If mqty4 > 0 And (m4t + mqty4) <= Val(tsize(3)) And mwhs <> 0 Then
        m4t = m4t + mqty4: mwhs4 = mwhs
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(3))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty4
            sqlx = sqlx & ", netqty = netqty - " & mqty4
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
        mqty = mqty + mqty4
    End If
    If mqty3 > 0 And (m3t + mqty3) <= Val(tsize(2)) And mwhs <> 0 Then
        m3t = m3t + mqty3: mwhs3 = mwhs
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(2))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty3
            sqlx = sqlx & ", netqty = netqty - " & mqty3
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
        mqty = mqty + mqty3
    End If
    If mqty2 > 0 And (m2t + mqty2) <= Val(tsize(1)) And mwhs <> 0 Then
        m2t = m2t + mqty2: mwhs2 = mwhs
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(1))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty2
            sqlx = sqlx & ", netqty = netqty - " & mqty2
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        bo.Close
        mqty = mqty + mqty2
    End If
    If mqty1 > 0 And (m1t + mqty1) <= Val(tsize(0)) And mwhs <> 0 Then
        m1t = m1t + mqty1: mwhs1 = mwhs
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(brcode(0))
        sqlx = sqlx & " and plant = " & Left(Combo6, 2)
        sqlx = sqlx & " and orddate = '" & Combo5 & "'"
        sqlx = sqlx & " and sku = '" & msku & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set bo = Sdb.Execute(sqlx)
        If bo.BOF = False Then
            bo.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & mqty1
            sqlx = sqlx & ", netqty = netqty - " & mqty1
            sqlx = sqlx & " Where id = " & bo!id
            Sdb.Execute sqlx
        End If
        mqty = mqty + mqty1
    End If
    If mwhs = xr1 Then
        s1t = s1t + mqty
        sqlx = "select whs_num,sku,grp_qty,avail,old_qty,old_lot from whstotals"
        sqlx = sqlx & " where sku = '" & msku & "' and old_lot > '0' and whs_num = " & xr1
        Set bs = Sdb.Execute(sqlx)
        If bs.BOF = False Then
            bs.MoveFirst
            i = bs!old_qty - mqty
            s1 = bs!old_qty - bs!grp_qty
            sqlx = "Update whstotals set grp_qty = grp_qty + " & mqty
            sqlx = sqlx & ", avail = avail - " & mqty
            If i <= 0 Then
                sqlx = sqlx & ", old_qty = 0, old_lot = ' '"
            Else
                sqlx = sqlx & ", old_qty = " & i
            End If
            sqlx = sqlx & " Where whs_num = " & xr1 & " and sku = '" & msku & "'"
            Sdb.Execute sqlx
        End If
        bs.Close
    End If
    If mwhs = xr2 Then
        s2t = s2t + mqty
        sqlx = "select whs_num,sku,grp_qty,avail,old_qty,old_lot from whstotals"
        sqlx = sqlx & " where sku = '" & msku & "' and old_lot > '0' and whs_num = " & xr2
        Set bs = Sdb.Execute(sqlx)
        If bs.BOF = False Then
            bs.MoveFirst
            i = bs!old_qty - mqty
            s2 = bs!old_qty - bs!grp_qty
            sqlx = "Update whstotals set grp_qty = grp_qty + " & mqty
            sqlx = sqlx & ", avail = avail - " & mqty
            If i <= 0 Then
                sqlx = sqlx & ", old_qty = 0, old_lot = ' '"
            Else
                sqlx = sqlx & ", old_qty = " & i
            End If
            sqlx = sqlx & " Where whs_num = " & xr2 & " and sku = '" & msku & "'"
            Sdb.Execute sqlx
        End If
        bs.Close
    End If
    If mwhs = xr3 Then
        s3t = s3t + mqty
        sqlx = "select whs_num,sku,grp_qty,avail,old_qty,old_lot from whstotals"
        sqlx = sqlx & " where sku = '" & msku & "' and old_lot > '0' and whs_num = " & xr3
        Set bs = Sdb.Execute(sqlx)
        If bs.BOF = False Then
            bs.MoveFirst
            i = bs!old_qty - mqty
            s3 = bs!old_qty - bs!grp_qty
            sqlx = "Update whstotals set grp_qty = grp_qty + " & mqty
            sqlx = sqlx & ", avail = avail - " & mqty
            If i <= 0 Then
                sqlx = sqlx & ", old_qty = 0, old_lot = ' '"
            Else
                sqlx = sqlx & ", old_qty = " & i
            End If
            sqlx = sqlx & " Where whs_num = " & xr3 & " and sku = '" & msku & "'"
            Sdb.Execute sqlx
        End If
        bs.Close
    End If
    If mwhs = xr5 Then
        s5t = s5t + mqty
        sqlx = "select whs_num,sku,grp_qty,avail,old_qty,old_lot from whstotals"
        sqlx = sqlx & " where sku = '" & msku & "' and old_lot > '0' and whs_num = " & xr5
        Set bs = Sdb.Execute(sqlx)
        If bs.BOF = False Then
            bs.MoveFirst
            i = bs!old_qty - mqty
            s5 = bs!old_qty - bs!grp_qty
            sqlx = "Update whstotals set grp_qty = grp_qty + " & mqty
            sqlx = sqlx & ", avail = avail - " & mqty
            If i <= 0 Then
                sqlx = sqlx & ", old_qty = 0, old_lot = ' '"
            Else
                sqlx = sqlx & ", old_qty = " & i
            End If
            sqlx = sqlx & " Where whs_num = " & xr5 & " and sku = '" & msku & "'"
            Sdb.Execute sqlx
        End If
        bs.Close
    End If
recnext:
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "oldsub(" & msku & ")", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " oldsub - Error Number: " & eno
        End
    End If
End Sub
Sub opmain()
    Dim ds As adodb.Recordset, ds2 As adodb.Recordset, sqlx As String, mavail As Integer
    Dim xr1 As Integer
    'On Error GoTo vberror
    sqlx = "select id,sku,qty1,whs1 from groupitems where groupcode = '" & Text2 & "' and sku in "
    sqlx = sqlx & "(select sku from whstotals where whs_num in (select whs_num from warehouses where whs in ('SR1','OP')))"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = "select sum(avail) from whstotals where sku = '" & ds!sku & "'"
            sqlx = sqlx & " and whs_num in (select whs_num from warehouses where whs in ('SR1','OP'))"
            Set ds2 = Sdb.Execute(sqlx)
            If ds2.BOF = True Then
                mavail = 0
            Else
                ds2.MoveFirst
                mavail = ds2(0)
            End If
            ds2.Close
            If mavail >= ds!qty1 Then
                sqlx = "Update groupitems set whs1 = " & xr1 & " Where id = " & ds!id
                Sdb.Execute sqlx
                sqlx = "select * from whstotals"
                sqlx = sqlx & " where whs_num = (select whs_num from warehouses where whs = 'SR1') and sku = '" & ds!sku & "'"
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    sqlx = "Update grp_qty = grp_qty + " & ds!qty1
                    sqlx = sqlx & ", avail = avail - " & ds!qty1
                    sqlx = sqlx & " Where id = " & ds2!id
                    Sdb.Execute sqlx
                End If
                ds2.Close
                sqlx = "select * from brorders"
                sqlx = sqlx & " where branch = " & Val(brcode(0))
                sqlx = sqlx & " and plant = " & Left(Combo6, 2)
                sqlx = sqlx & " and orddate = '" & Combo5 & "'"
                sqlx = sqlx & " and sku = '" & ds!sku & "'"
                sqlx = sqlx & " and ordqty > 0"
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    sqlx = "Update brorders set grpqty = grpqty + " & ds!qty1
                    sqlx = sqlx & ", netqty = netqty - " & ds!qty1
                    sqlx = sqlx & " Where id = " & ds2!id
                    Sdb.Execute sqlx
                End If
                ds2.Close
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "opmain", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " opmain - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo1_Click()
    brcode(0).ListIndex = Combo1.ListIndex
    tsize(0).ListIndex = Combo1.ListIndex
    trun(0).ListIndex = Combo1.ListIndex
End Sub

Private Sub Combo2_Click()
    brcode(1).ListIndex = Combo2.ListIndex
    tsize(1).ListIndex = Combo2.ListIndex
    trun(1).ListIndex = Combo2.ListIndex
End Sub

Private Sub Combo3_Click()
    brcode(2).ListIndex = Combo3.ListIndex
    tsize(2).ListIndex = Combo3.ListIndex
    trun(2).ListIndex = Combo3.ListIndex
End Sub

Private Sub Combo4_Click()
    brcode(3).ListIndex = Combo4.ListIndex
    tsize(3).ListIndex = Combo4.ListIndex
    trun(3).ListIndex = Combo4.ListIndex
End Sub

Private Sub Combo5_Click()
    Dim ds As adodb.Recordset, sqlx As String
    'On Error GoTo vberror
    Screen.MousePointer = 11
    Form1.cdate = Format(Combo5, "m-d-yyyy")
    Combo6.Clear
    sqlx = "select plant,plantname from plants"
    sqlx = sqlx & " where plant in (select loaded from runs where trldate = '" & Combo5 & "')"
    sqlx = sqlx & " order by plant"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo6.AddItem Format$(ds!plant, "00") & " " & ds!plantname
            ds.MoveNext
        Loop
        Combo6.ListIndex = 0
    End If
    ds.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "combo5_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " combo5_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo6_Click()
    Dim ds As adodb.Recordset, gflag As Boolean
    'On Error GoTo vberror
    If Left(Combo6, 2) = "50" Then              'jv061815
        Grid2.Visible = True                    'jv061815
    Else                                        'jv061815
        Grid2.Visible = False                   'jv061815
    End If                                      'jv061815
    sqlx = "Select id, destination, locname, trlno, trlsize, loaded from runs"
    sqlx = sqlx & " where loaded = '" & Left$(Combo6, 2) & "'"
    sqlx = sqlx & " and trldate = '" & Combo5 & "'"
    sqlx = sqlx & " and id not in (select run1 from trgroups)"
    sqlx = sqlx & " and id not in (select run2 from trgroups)"
    sqlx = sqlx & " and id not in (select run3 from trgroups)"
    sqlx = sqlx & " and id not in (select run4 from trgroups)"
    sqlx = sqlx & " and id not in (select runid from trailers)"
    sqlx = sqlx & " order by locname, trlno"
    Set ds = Sdb.Execute(sqlx)
    Combo1.Clear: Combo2.Clear: Combo3.Clear: Combo4.Clear
    brcode(0).Clear: brcode(1).Clear: brcode(2).Clear: brcode(3).Clear
    tsize(0).Clear: tsize(1).Clear: tsize(2).Clear: tsize(3).Clear
    trun(0).Clear: trun(1).Clear: trun(2).Clear: trun(3).Clear
    Combo1.AddItem "...": Combo2.AddItem "..."
    Combo3.AddItem "...": Combo4.AddItem "..."
    brcode(0).AddItem "0": brcode(1).AddItem "0"
    brcode(2).AddItem "0": brcode(3).AddItem "0"
    trun(0).AddItem "0": trun(1).AddItem "0"
    trun(2).AddItem "0": trun(3).AddItem "0"
    tsize(0).AddItem "0": tsize(1).AddItem "0"
    tsize(2).AddItem "0": tsize(3).AddItem "0"
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            gflag = True
            If ds!Destination = "16" Then gflag = False
            If ds!Destination = "15" Then gflag = False
            If Val(ds!Destination) = 0 Then gflag = False
            'If Val(ds!Destination) = 1 Then gflag = False
            If ds!loaded = "50" And ds!Destination = "51" Then gflag = False
            If ds!loaded = "51" And ds!Destination = "50" Then gflag = False
            If ds!loaded = "52" And ds!Destination = "50" Then gflag = False
            If gflag = True Then
                Combo1.AddItem ds!locname & " " & ds!trlno
                Combo2.AddItem ds!locname & " " & ds!trlno
                Combo3.AddItem ds!locname & " " & ds!trlno
                Combo4.AddItem ds!locname & " " & ds!trlno
                brcode(0).AddItem ds!Destination: brcode(1).AddItem ds!Destination
                brcode(2).AddItem ds!Destination: brcode(3).AddItem ds!Destination
                trun(0).AddItem ds!id: trun(1).AddItem ds!id
                trun(2).AddItem ds!id: trun(3).AddItem ds!id
                tsize(0).AddItem ds!trlsize: tsize(1).AddItem ds!trlsize
                tsize(2).AddItem ds!trlsize: tsize(3).AddItem ds!trlsize
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "combo6_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " combo6_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command1_Click()            'Post Group
    Dim sqlx As String, ds As adodb.Recordset, dt As adodb.Recordset, dp As adodb.Recordset
    Dim ssku As String, mrank As Integer, pq As Integer, xwhs As Integer, pkey As Long
    Dim mod1 As Single, mod2 As Single, mod3 As Single, mod4 As Single
    'On Error GoTo vberror
    If Text2 < "." Then
        MsgBox "Group Code Required", vbOKOnly, "Cannot Post"
        Exit Sub
    End If
    Screen.MousePointer = 11
    mqty1 = 0: mqty2 = 0: mqty3 = 0: mqty4 = 0
    m1t = 0: m2t = 0: m3t = 0: m4t = 0: xwhs = 0
    If Left$(Combo6, 2) <> "50" Then
        Set ds = Sdb.Execute("select * from warehouses where plant = " & Left(Combo6, 2))
        If ds.BOF = True Then
            ds.Close
            MsgBox "Plant " & Left$(Combo6, 2) & " warehouse not found..", vbOKOnly, "Aborting.."
            Exit Sub
        Else
            xwhs = ds!whs_num
        End If
        ds.Close
    End If
    Form1.cgrp = Text2
    Set ds = Sdb.Execute("select * from trgroups where groupcode = '" & Text2 & "'")
    If ds.BOF = False Then
        ds.MoveFirst
        sqlx = "Update trgoups set run1 = " & Val(trun(0))
        sqlx = sqlx & ", run2 = " & Val(trun(1))
        sqlx = sqlx & ", run3 = " & Val(trun(2))
        sqlx = sqlx & ", run4 = " & Val(trun(3))
        sqlx = sqlx & " Where groupcode = '" & Text2 & "'"
        Sdb.Execute sqlx
    Else
        sqlx = "Insert into trgroups (groupcode, run1, run2, run3, run4) Values ('" & Text2 & "'"
        sqlx = sqlx & ", " & Val(trun(0))
        sqlx = sqlx & ", " & Val(trun(1))
        sqlx = sqlx & ", " & Val(trun(2))
        sqlx = sqlx & ", " & Val(trun(3)) & ")"
        Sdb.Execute sqlx
    End If
    ds.Close
    sqlx = "delete from groupitems where groupcode = '" & Text2 & "'"
    Sdb.Execute sqlx
    sqlx = "select sku,branch,netqty,partqty from brorders where branch in ("
    sqlx = sqlx & Val(brcode(0)) & ","
    sqlx = sqlx & Val(brcode(1)) & ","
    sqlx = sqlx & Val(brcode(2)) & ","
    sqlx = sqlx & Val(brcode(3)) & ")"
    sqlx = sqlx & " and Plant = " & Left$(Combo6, 2)
    sqlx = sqlx & " and orddate = '" & Combo5 & "'"
    sqlx = sqlx & " and (netqty > 0 or partqty > 0)"
    sqlx = sqlx & " and account not like 'OC*'"
    'sqlx = sqlx & " and sku in (select sku from whstotals where whs_num in (select whs_num from warehouses where plant = " & Val(Left$(Combo6, 2)) & "))"
    sqlx = sqlx & " order by sku"
    pq = 0
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        ssku = ds!sku
        Do Until ds.EOF
            pq = pq + ds!partqty
            If ds!sku <> ssku Then
                mod1 = mqty1 Mod 2
                mod2 = mqty2 Mod 2
                mod3 = mqty3 Mod 2
                mod4 = mqty4 Mod 2
                If ((mod1 = 1 And mod2 = 1) Or (mod1 = 1 And mod3 = 1) Or (mod2 = 1 And mod3 = 1)) And ((mqty1 + mqty2 + mqty3 + mqty4) Mod 2) = 0 Then
                    mrank = 1
                Else
                    If ((mqty1 + mqty2 + mqty3 + mqty4) Mod 2) = 0 Then
                        mrank = 2
                    Else
                        mrank = 3
                    End If
                End If
                If ssku = "777" And Val(Left$(Combo6, 2)) = 50 Then
                    'mod1 = (mqty1 / 2) Mod 2
                    'mod2 = (mqty2 / 2) Mod 2
                    'mod3 = (mqty3 / 2) Mod 2
                    'mod4 = (mqty4 / 2) Mod 2
                    pkey = wd_seq("groupitems", Form1.shipdb)
                    sqlx = "Insert into groupitems (id, groupcode, sku, qty1, qty2, qty4, whs1, whs2, whs3"
                    sqlx = sqlx & ", whs4, grank) Values (" & pkey
                    sqlx = sqlx & ", '" & Text2 & "'"
                    sqlx = sqlx & ", '" & ssku & "'"
                    sqlx = sqlx & ", " & split_hmv(mqty1, 1)
                    sqlx = sqlx & ", " & split_hmv(mqty2, 1)
                    sqlx = sqlx & ", " & split_hmv(mqty4, 1)
                    sqlx = sqlx & ", 0, 0, 0, 0, 3)"
                    Sdb.Execute sqlx
                    pkey = wd_seq("groupitems", Form1.shipdb)
                    sqlx = "Insert into groupitems (id, groupcode, sku, qty1, qty3, qty4, whs1, whs2, whs3"
                    sqlx = sqlx & ", whs4, grank) Values (" & pkey
                    sqlx = sqlx & ", '" & Text2 & "'"
                    sqlx = sqlx & ", '" & ssku & "'"
                    sqlx = sqlx & ", " & split_hmv(mqty1, 2)
                    sqlx = sqlx & ", " & split_hmv(mqty3, 1)
                    sqlx = sqlx & ", " & split_hmv(mqty4, 2)
                    sqlx = sqlx & ", 0, 0, 0, 0, 2)"
                    Sdb.Execute sqlx
                    pkey = wd_seq("groupitems", Form1.shipdb)
                    sqlx = "Insert into groupitems (id, groupcode, sku, qty2, qty3, whs1, whs2, whs3"
                    sqlx = sqlx & ", whs4, grank) Values (" & pkey
                    sqlx = sqlx & ", '" & Text2 & "'"
                    sqlx = sqlx & ", '" & ssku & "'"
                    sqlx = sqlx & ", " & split_hmv(mqty2, 2)
                    sqlx = sqlx & ", " & split_hmv(mqty3, 2)
                    sqlx = sqlx & ", 0, 0, 0, 0, 2)"
                    Sdb.Execute sqlx
                Else
                    If mqty1 + mqty2 + mqty3 + mqty4 > 0 Then
                        pkey = wd_seq("groupitems", Form1.shipdb)
                        sqlx = "Insert into groupitems (id, groupcode, sku, qty1, qty2, qty3, qty4"
                        sqlx = sqlx & ", whs1, whs2, whs3, whs4, grank) Values (" & pkey
                        sqlx = sqlx & ", '" & Text2 & "'"
                        sqlx = sqlx & ", '" & ssku & "'"
                        sqlx = sqlx & ", " & mqty1
                        sqlx = sqlx & ", " & mqty2
                        sqlx = sqlx & ", " & mqty3
                        sqlx = sqlx & ", " & mqty4
                        sqlx = sqlx & ", 0, 0, 0, 0, " & mrank & ")"
                        Sdb.Execute sqlx
                    End If
                End If
                mqty1 = 0: mqty2 = 0: mqty3 = 0: mqty4 = 0: ssku = ds!sku
            End If
            If ds!branch = Val(brcode(0)) Then mqty1 = ds!netqty
            If ds!branch = Val(brcode(1)) Then mqty2 = ds!netqty
            If ds!branch = Val(brcode(2)) Then mqty3 = ds!netqty
            If ds!branch = Val(brcode(3)) Then mqty4 = ds!netqty
            ds.MoveNext
        Loop
    End If
    mod1 = mqty1 Mod 2
    mod2 = mqty2 Mod 2
    mod3 = mqty3 Mod 2
    mod4 = mqty4 Mod 2
    If ((mod1 = 1 And mod2 = 1) Or (mod1 = 1 And mod3 = 1) Or (mod2 = 1 And mod3 = 1)) And ((mqty1 + mqty2 + mqty3 + mqty4) Mod 2) = 0 Then
        mrank = 1
    Else
        If ((mqty1 + mqty2 + mqty3 + mqty4) Mod 2) = 0 Then
            mrank = 2
        Else
            mrank = 3
        End If
    End If
    If Val(ssku) > 0 And mqty1 + mqty2 + mqty3 + mqty4 > 0 Then
        pkey = wd_seq("groupitems", Form1.shipdb)
        sqlx = "Insert into groupitems (id, groupcode, sku, qty1, qty2, qty3, qty4"
        sqlx = sqlx & ", whs1, whs2, whs3, whs4, grank) Values (" & pkey
        sqlx = sqlx & ", '" & Text2 & "'"
        sqlx = sqlx & ", '" & ssku & "'"
        sqlx = sqlx & ", " & mqty1
        sqlx = sqlx & ", " & mqty2
        sqlx = sqlx & ", " & mqty3
        sqlx = sqlx & ", " & mqty4
        sqlx = sqlx & ", 0, 0, 0, 0, " & mrank & ")"
        Sdb.Execute sqlx
    End If
    If pq > 0 Then
        pkey = wd_seq("groupitems", Form1.shipdb)
        sqlx = "Insert into groupitems (id, groupcode, sku, qty1, qty2, qty3, qty4"
        sqlx = sqlx & ", whs1, whs2, whs3, whs4, grank) Values (" & pkey
        sqlx = sqlx & ", '" & Text2 & "'"
        sqlx = sqlx & ", 'PAR'"
        sqlx = sqlx & ", 0, 0, 0, 0, 0, 0, 0, 0, 1)"
        Sdb.Execute sqlx
    End If
    ds.Close
    'If Val(Left$(Combo6, 2)) = baplant Then
    '    Call bamain
    'Else
    '    If Val(Left$(Combo6, 2)) = syplant Then
    '        Call symain
    If xwhs > 0 Then
        Call pltmain(xwhs)
    Else
        If Right$(Text2, 3) = "-OP" Then
            Call opmain
        Else
            Call cranmain
        End If
    End If
    Call refresh_grid
    Screen.MousePointer = 0
    MsgBox "Group: " & Text2 & " has been posted....", vbOKOnly, "Complete....."
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

Private Sub Command2_Click()            'Exit
    Unload PostBrGrps
End Sub

Private Sub Command3_Click()            'Clear Group
    'On Error GoTo vberror
    If MsgBox("Ok to clear group: " & Grid1.TextMatrix(Grid1.Row, 0), vbYesNo + vbQuestion, "Clear Group..") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    Grid1.Col = 0
    Sdb.Execute "delete from trgroups where groupcode = '" & Grid1.Text & "'"
    Sdb.Execute "delete from groupitems where groupcode = '" & Grid1.Text & "'"
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_grid
    End If
    Call Grid1_Click
    Screen.MousePointer = 0
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

Private Sub Command4_Click()            'Reverse Group
    Dim ds As adodb.Recordset, sqlx As String, ds2 As adodb.Recordset
    Dim b1 As Integer, b2 As Integer, b3 As Integer, b4 As Integer
    Dim mplant As Integer, mdate As String
    'On Error GoTo vberror
    If MsgBox("Ok to reverse group: " & Grid1.TextMatrix(Grid1.Row, 0), vbYesNo + vbQuestion, "Reverse Group..") = vbNo Then Exit Sub
    Screen.MousePointer = 11
    Grid1.Col = 0
    sqlx = "Select * From groupitems where groupcode = '" & Grid1.Text & "'"
    sqlx = sqlx & " and sku <> 'PAR'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Grid1.Col = 5: b1 = Val(mid$(Grid1.Text, 3, 2))
        mplant = Val(Left$(Grid1.Text, 2)): mdate = "'" & Right$(Grid1.Text, 10) & "'"
        Grid1.Col = 6: b2 = Val(mid$(Grid1.Text, 3, 2))
        Grid1.Col = 7: b3 = Val(mid$(Grid1.Text, 3, 2))
        Grid1.Col = 8: b4 = Val(mid$(Grid1.Text, 3, 2))
        Do Until ds.EOF
            If ds!qty1 > 0 And ds!whs1 > 0 Then
                sqlx = "select * from brorders"
                sqlx = sqlx & " Where plant = " & mplant
                sqlx = sqlx & " And branch = " & b1
                sqlx = sqlx & " and orddate = " & mdate
                sqlx = sqlx & " and sku = '" & ds!sku & "'"
                sqlx = sqlx & " and grpqty > 0"
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    sqlx = "Update brorders set grpqty = grpqty - " & ds!qty1
                    sqlx = sqlx & ", netqty = netqty + " & ds!qty1
                    sqlx = sqlx & " Where id = " & ds2!id
                    Sdb.Execute sqlx
                End If
                ds2.Close
                sqlx = "select * from whstotals"
                sqlx = sqlx & " Where whs_num = " & ds!whs1
                sqlx = sqlx & " and sku = '" & ds!sku & "'"
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    sqlx = "Update whstotals set grp_qty = grp_qty - " & ds!qty1
                    sqlx = sqlx & ", avail = avail + " & ds!qty1
                    sqlx = sqlx & " Where id = " & ds2!id
                    Sdb.Execute sqlx
                End If
                ds2.Close
            End If
            If ds!qty2 > 0 And ds!whs2 > 0 Then
                sqlx = "select * from brorders"
                sqlx = sqlx & " Where plant = " & mplant
                sqlx = sqlx & " And branch = " & b2
                sqlx = sqlx & " and orddate = " & mdate
                sqlx = sqlx & " and sku = '" & ds!sku & "'"
                sqlx = sqlx & " and grpqty > 0"
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    sqlx = "Update brorders set grpqty = grpqty - " & ds!qty2
                    sqlx = sqlx & ", netqty = netqty + " & ds!qty2
                    sqlx = sqlx & " Where id = " & ds2!id
                    Sdb.Execute sqlx
                End If
                ds2.Close
                sqlx = "select * from whstotals"
                sqlx = sqlx & " Where whs_num = " & ds!whs2
                sqlx = sqlx & " and sku = '" & ds!sku & "'"
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    sqlx = "Update whstotals set grp_qty = grp_qty - " & ds!qty2
                    sqlx = sqlx & ", avail = avail + " & ds!qty2
                    sqlx = sqlx & " Where id = " & ds2!id
                    Sdb.Execute sqlx
                End If
                ds2.Close
            End If
            If ds!qty3 > 0 And ds!whs3 > 0 Then
                sqlx = "select * from brorders"
                sqlx = sqlx & " Where plant = " & mplant
                sqlx = sqlx & " And branch = " & b3
                sqlx = sqlx & " and orddate = " & mdate
                sqlx = sqlx & " and sku = '" & ds!sku & "'"
                sqlx = sqlx & " and grpqty > 0"
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    sqlx = "Update brorders set grpqty = grpqty - " & ds!qty3
                    sqlx = sqlx & ", netqty = netqty + " & ds!qty3
                    sqlx = sqlx & " Where id = " & ds2!id
                    Sdb.Execute sqlx
                End If
                ds2.Close
                sqlx = "select * from whstotals"
                sqlx = sqlx & " Where whs_num = " & ds!whs3
                sqlx = sqlx & " and sku = '" & ds!sku & "'"
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    sqlx = "Update whstotals set grp_qty = grp_qty - " & ds!qty3
                    sqlx = sqlx & ", avail = avail + " & ds!qty3
                    sqlx = sqlx & " Where id = " & ds2!id
                    Sdb.Execute sqlx
                End If
                ds2.Close
            End If
            If ds!qty4 > 0 And ds!whs4 > 0 Then
                sqlx = "select * from brorders"
                sqlx = sqlx & " Where plant = " & mplant
                sqlx = sqlx & " And branch = " & b4
                sqlx = sqlx & " and orddate = " & mdate
                sqlx = sqlx & " and sku = '" & ds!sku & "'"
                sqlx = sqlx & " and grpqty > 0"
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    sqlx = "Update brorders set grpqty = grpqty - " & ds!qty4
                    sqlx = sqlx & ", netqty = netqty + " & ds!qty4
                    sqlx = sqlx & " Where id = " & ds2!id
                    Sdb.Execute sqlx
                End If
                ds2.Close
                sqlx = "select * from whstotals"
                sqlx = sqlx & " Where whs_num = " & ds!whs4
                sqlx = sqlx & " and sku = '" & ds!sku & "'"
                Set ds2 = Sdb.Execute(sqlx)
                If ds2.BOF = False Then
                    ds2.MoveFirst
                    sqlx = "Update whstotals set grp_qty = grp_qty - " & ds!qty4
                    sqlx = sqlx & ", avail = avail + " & ds!qty4
                    sqlx = sqlx & " Where id = " & ds2!id
                    Sdb.Execute sqlx
                End If
                ds2.Close
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid1.Col = 0
    Sdb.Execute "delete from trgroups where groupcode = '" & Grid1.Text & "'"
    Sdb.Execute "delete from groupitems where groupcode = '" & Grid1.Text & "'"
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Call refresh_grid
    End If
    Call Grid1_Click
    Screen.MousePointer = 0
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

Private Sub Form_Deactivate()
    Dim i As Integer
    If PostBrGrps.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            If Form1.FrmGrid.TextMatrix(i, 0) = "postbrgrps" Then
                Form1.FrmGrid.TextMatrix(i, 1) = PostBrGrps.Top
                Form1.FrmGrid.TextMatrix(i, 2) = PostBrGrps.Left
                Form1.FrmGrid.TextMatrix(i, 3) = PostBrGrps.Height
                Form1.FrmGrid.TextMatrix(i, 4) = PostBrGrps.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim ds As adodb.Recordset, sqlx As String
    Dim i As Integer
    For i = 1 To Form1.FrmGrid.Rows - 1
        If Form1.FrmGrid.TextMatrix(i, 0) = "postbrgrps" Then
            PostBrGrps.Top = Val(Form1.FrmGrid.TextMatrix(i, 1))
            PostBrGrps.Left = Val(Form1.FrmGrid.TextMatrix(i, 2))
            PostBrGrps.Height = Val(Form1.FrmGrid.TextMatrix(i, 3))
            PostBrGrps.Width = Val(Form1.FrmGrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    'On Error GoTo vberror
    sqlx = "select distinct trldate from runs order by trldate"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo5.AddItem Format$(ds!trldate, "m-d-yyyy")
            ds.MoveNext
        Loop
        For i = 0 To Combo5.ListCount - 1
            If Combo5.List(i) = Form1.cdate Then
                Combo5.ListIndex = i
                Exit For
            End If
        Next i
        If Combo5.ListIndex < 0 Then Combo5.ListIndex = 0
    End If
    ds.Close
    Grid1.Font = "Arial": Grid1.FontSize = 8: Grid1.FontBold = True
    Call refresh_grid
    Call refresh_grid2
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
    If PostBrGrps.Width > Frame1.Width Then
        Grid1.Width = Frame1.Width
    Else
        Grid1.Width = PostBrGrps.Width
    End If
    If PostBrGrps.Height > Frame1.Height + 900 Then
        Command3.Top = PostBrGrps.Height - 850
        Command4.Top = Command3.Top
        Grid1.Height = PostBrGrps.Height - (Frame1.Height + 900)
    End If
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Grid1_Click()
    Grid1.Col = 0
    If Grid1.Text > " " Then
        Command3.Enabled = True
        Command4.Enabled = True
    Else
        Command3.Enabled = False
        Command4.Enabled = False
    End If
    Grid1.ColSel = Grid1.Cols - 1
    Grid1.RowSel = Grid1.Row
End Sub

Private Sub Grid2_KeyPress(KeyAscii As Integer)                             'jv061815
    Dim s As String
    If Val(Grid2.TextMatrix(Grid2.Row, 0)) = 0 Then Exit Sub
    If Grid2.TextMatrix(Grid2.Row, 2) = "No" Then
        s = "Update warehouses set vert_loc = 2 where whs_num = " & Grid2.TextMatrix(Grid2.Row, 0)
        Grid2.TextMatrix(Grid2.Row, 2) = "Yes"
    Else
        s = "Update warehouses set vert_loc = 0 where whs_num = " & Grid2.TextMatrix(Grid2.Row, 0)
        Grid2.TextMatrix(Grid2.Row, 2) = "No"
    End If
    Sdb.Execute s
    refresh_grid2
End Sub

Private Sub Text2_Change()
    Dim i As Integer
    Command1.Enabled = True
    For i = 1 To Grid1.Rows - 1
        If UCase(Trim(Text2)) = UCase(Grid1.TextMatrix(i, 0)) Then Command1.Enabled = False
    Next i
End Sub
