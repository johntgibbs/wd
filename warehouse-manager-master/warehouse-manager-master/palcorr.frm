VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form palcorr 
   Caption         =   "Pallet Correction"
   ClientHeight    =   9690
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8130
   LinkTopic       =   "Form14"
   ScaleHeight     =   9690
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid4 
      Height          =   1095
      Left            =   0
      TabIndex        =   20
      Top             =   8520
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1931
      _Version        =   327680
      ForeColor       =   192
      BackColorFixed  =   12648384
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   1095
      Left            =   0
      TabIndex        =   18
      Top             =   7200
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1931
      _Version        =   327680
      BackColorFixed  =   12648447
      AllowUserResizing=   3
   End
   Begin VB.Frame Frame1 
      Caption         =   "Updates: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   11
      Top             =   6960
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CheckBox Check3 
         Caption         =   "Post Adjustment to W/D Browser"
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
         TabIndex        =   16
         Top             =   1080
         Width           =   3975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   14
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Update Rack Positions, Queues && Pallet Tasks"
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
         TabIndex        =   13
         Top             =   720
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Post to Receiving Logs (Correct R12 Receipts)"
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
         TabIndex        =   12
         Top             =   360
         Value           =   1  'Checked
         Width           =   6855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1215
      Left            =   0
      TabIndex        =   8
      Top             =   9960
      Visible         =   0   'False
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   2143
      _Version        =   327680
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   5175
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9128
      _Version        =   327680
      Cols            =   3
      FixedCols       =   2
      ForeColor       =   4210816
      BackColorFixed  =   16777152
      BackColorSel    =   255
      FocusRect       =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find Plate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find BarCode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label afile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   8280
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label Lfile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   6960
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label pref 
      Caption         =   "..."
      Height          =   255
      Left            =   8160
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label plkey 
      Caption         =   "..."
      Height          =   255
      Left            =   8160
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label bckey 
      Caption         =   "..."
      Height          =   255
      Left            =   8160
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label proddesc 
      Alignment       =   2  'Center
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Label Label2 
      Caption         =   "Plate:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Bar Code:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
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
      Width           =   1815
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu edval 
         Caption         =   "Change Value"
      End
   End
End
Attribute VB_Name = "palcorr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function batch_hold(psku As String, plot As String, pcode As String) As Boolean
    Dim s As String, ds As adodb.Recordset
    s = "select id from holdlist where sku = '" & psku & "'"
    s = s & " and lot_num = '" & plot & "'"
    s = s & " and opcode = '" & pcode & "'"
    s = s & " and spallet = '001' and epallet = 'EOR'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        batch_hold = True
    Else
        batch_hold = False
    End If
    ds.Close
End Function

Function r12_lot(plot As String, ocode As String) As String
    Dim s As String, myear As Integer, mdays As Integer
    If Len(plot) >= 5 Then
        myear = Val(Left(plot, 2))
        mdays = Val(Mid(plot, 3, 3)) - 1
        s = "1-1-20" & Left(plot, 2)
        s = Format(DateAdd("d", mdays, s), "MMddyy")
        s = Left(s, 4) & Format(myear + 2, "00")
        If Len(plot) > 5 Then               'jv080315
            s = s & Right(plot, 3)          'jv080315
        Else                                'jv080315
            s = s & ocode                   'jv080315
        End If                              'jv080315
    Else
        s = " "
    End If
    r12_lot = s
End Function

Private Sub refresh_grid_barcode()
    Dim s As String, ds As adodb.Recordset, pcode As String
    Frame1.Visible = False
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 12
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    Grid1.FontName = "Arial"
    Grid1.FontSize = 10
    Grid1.FontBold = True
    s = "select * from pallets where barcode = '" & Text1 & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Text2 = ds!plateno
            s = "ID" & Chr(9) & ds!id & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Plate" & Chr(9) & ds!plateno & Chr(9) & ds!plateno
            Grid1.AddItem s
            s = "BarCode" & Chr(9) & ds!barcode & Chr(9) & ds!barcode & Chr(9) & "*" & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Qty1" & Chr(9) & ds!qty1 & Chr(9) & ds!qty1 & Chr(9) & "*" & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Lot1" & Chr(9) & ds!lot1 & Chr(9) & ds!lot1 & Chr(9) & "*" & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Qty2" & Chr(9) & ds!qty2 & Chr(9) & ds!qty2 & Chr(9) & "*" & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Lot2" & Chr(9) & ds!lot2 & Chr(9) & ds!lot2 & Chr(9) & "*" & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Source" & Chr(9) & ds!source & Chr(9) & ds!source
            Grid1.AddItem s
            s = "Target" & Chr(9) & ds!target & Chr(9) & ds!target
            Grid1.AddItem s
            s = "Status" & Chr(9) & ds!status & Chr(9) & ds!status
            Grid1.AddItem s
            psku = Trim(Left(ds!barcode, 4))
            proddesc = psku & " " & skurec(Val(psku)).prodname
            s = ds!id & Chr(9)
            s = s & ds!plateno & Chr(9)
            s = s & ds!barcode & Chr(9)
            s = s & ds!qty1 & Chr(9)
            s = s & ds!lot1 & Chr(9)
            s = s & ds!qty2 & Chr(9)
            s = s & ds!lot2 & Chr(9)
            s = s & ds!source & Chr(9)
            s = s & ds!target & Chr(9)
            s = s & ds!status
            Grid2.AddItem "Current" & Chr(9) & s
            Grid2.AddItem "New" & Chr(9) & s
            pcode = Mid(ds!barcode, 11, 3)
            If batch_hold(ds!sku, ds!lot1, pcode) = True Then
                Check1.Enabled = True
                Check1.Value = 1
                Check3.Enabled = False
                Check3.Value = 0
            Else
                Check1.Value = 0
                Check1.Enabled = False
                Check3.Enabled = True
                Check3.Value = 1
            End If
            ds.MoveNext
        Loop
    Else
        Text2 = ""
        proddesc.Caption = "BarCode not found!"
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 3) <> "*" Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = Grid1.BackColorFixed
            End If
        Next i
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 0) = "Qty1" Then
                Grid1.Row = i: Grid1.Col = 2
                Exit For
            End If
        Next i
    End If
    Grid1.FormatString = "^Field|<Current Value|<New Value|^"
    Grid1.ColWidth(0) = 1400
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 0 '600
    Grid1.ColWidth(4) = 0
    Grid1.Redraw = True
End Sub

Private Sub refresh_grid_plate()
    Dim s As String, ds As adodb.Recordset, pcode As String
    Frame1.Visible = False
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 12
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    Grid1.FontName = "Arial"
    Grid1.FontSize = 10
    Grid1.FontBold = True
    s = "select * from pallets where plateno = '" & Text2 & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Text1 = ds!barcode
            s = "ID" & Chr(9) & ds!id & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Plate" & Chr(9) & ds!plateno & Chr(9) & ds!plateno
            Grid1.AddItem s
            s = "BarCode" & Chr(9) & ds!barcode & Chr(9) & ds!barcode & Chr(9) & "*" & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Qty1" & Chr(9) & ds!qty1 & Chr(9) & ds!qty1 & Chr(9) & "*" & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Lot1" & Chr(9) & ds!lot1 & Chr(9) & ds!lot1 & Chr(9) & "*" & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Qty2" & Chr(9) & ds!qty2 & Chr(9) & ds!qty2 & Chr(9) & "*" & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Lot2" & Chr(9) & ds!lot2 & Chr(9) & ds!lot2 & Chr(9) & "*" & Chr(9) & ds!id
            Grid1.AddItem s
            s = "Source" & Chr(9) & ds!source & Chr(9) & ds!source
            Grid1.AddItem s
            s = "Target" & Chr(9) & ds!target & Chr(9) & ds!target
            Grid1.AddItem s
            s = "Status" & Chr(9) & ds!status & Chr(9) & ds!status
            Grid1.AddItem s
            psku = Trim(Left(ds!barcode, 4))
            proddesc = psku & " " & skurec(Val(psku)).prodname
            s = ds!id & Chr(9)
            s = s & ds!plateno & Chr(9)
            s = s & ds!barcode & Chr(9)
            s = s & ds!qty1 & Chr(9)
            s = s & ds!lot1 & Chr(9)
            s = s & ds!qty2 & Chr(9)
            s = s & ds!lot2 & Chr(9)
            s = s & ds!source & Chr(9)
            s = s & ds!target & Chr(9)
            s = s & ds!status
            Grid2.AddItem "Current" & Chr(9) & s
            Grid2.AddItem "New" & Chr(9) & s
            pcode = Mid(ds!barcode, 11, 3)
            If batch_hold(ds!sku, ds!lot1, pcode) = True Then
                Check1.Enabled = True
                Check1.Value = 1
                Check3.Enabled = False
                Check3.Value = 0
            Else
                Check1.Value = 0
                Check1.Enabled = False
                Check3.Enabled = True
                Check3.Value = 1
            End If
            ds.MoveNext
        Loop
    Else
        Text1 = ""
        proddesc.Caption = "Plate not found!"
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 3) <> "*" Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = Grid1.BackColorFixed
            End If
        Next i
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 0) = "Qty1" Then
                Grid1.Row = i: Grid1.Col = 2
                Exit For
            End If
        Next i
    End If
    Grid1.FormatString = "^Field|<Current Value|<New Value|^"
    Grid1.ColWidth(0) = 1400
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 0 '600
    Grid1.ColWidth(4) = 0
    Grid1.Redraw = True
End Sub


Private Sub bckey_Change()
    Text1 = bckey
    refresh_grid_barcode
End Sub

Private Sub Check1_Click()
    'If Check1.Value = 1 Then
    '    Check3.Enabled = False
    '    Check3.Value = 0
    'Else
    '    Check3.Enabled = True
    'End If
End Sub

Private Sub Check3_Click()
    'If Check3.Value = 1 Then
    '    Check1.Enabled = False
    '    Check1.Value = 0
    'Else
    '    Check1.Enabled = True
    'End If
End Sub

Private Sub Command1_Click()
    refresh_grid_barcode
End Sub

Private Sub Command2_Click()
    refresh_grid_plate
End Sub

Private Sub Command3_Click()
    Dim i As Integer, s As String, p As ptask, preas As String
    Dim psku As String, pname As String, cfile As String, plot As String, pcode As String
    For i = 1 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 0) = "New" And Grid2.TextMatrix(i, 11) = "Post" Then
            If Val(Grid2.TextMatrix(i, 4)) <= 0 Then            'Qty1 > 0
                MsgBox "Qty1 must be greated than zero.", vbOKOnly + vbInformation, "Update cancelled..."
                Exit Sub
            End If
            If Val(Grid2.TextMatrix(i, 6)) <= 0 And Grid2.TextMatrix(i, 7) > "0" Then
                MsgBox "Qty2 must be greater than zero for second lot.", vbOKOnly + vbInformation, "Update cancelled..."
                Exit Sub
            End If
            If Val(Grid2.TextMatrix(i, 6)) > 0 And Grid2.TextMatrix(i, 7) < "0" Then
                MsgBox "Second lot is not specified for Qty2.", vbOKOnly + vbInformation, "Update cancelled..."
                Exit Sub
            End If
        End If
    Next i
    'If Check1.Value = 1 Then
        preas = InputBox("Reason for correction:", "Reason for correction....")
    'End If
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 17
    Grid4.Clear: Grid4.Rows = 1: Grid4.Cols = 10
    For i = 1 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 0) = "New" And Grid2.TextMatrix(i, 11) = "Post" Then
            psku = Trim(Left(Grid2.TextMatrix(i, 3), 4))
            s = "Update pallets set barcode = '" & Grid2.TextMatrix(i, 3) & "'"
            s = s & ", qty1 = " & Val(Grid2.TextMatrix(i, 4))
            s = s & ", lot1 = '" & Grid2.TextMatrix(i, 5) & "'"
            s = s & ", qty2 = " & Val(Grid2.TextMatrix(i, 6))
            s = s & ", lot2 = '" & Grid2.TextMatrix(i, 7) & "'"
            s = s & ", sku = '" & psku & "'"
            s = s & " Where id = " & Val(Grid2.TextMatrix(i, 1))
            MsgBox s
            If Check2.Value = 1 Then
                s = "Update position set sku = '" & psku & "'"
                s = s & ", lot_num = '" & Grid2.TextMatrix(i, 5) & "'"
                s = s & ", count_qty = " & Val(Grid2.TextMatrix(i, 4))
                s = s & ", lot2 = '" & Grid2.TextMatrix(i, 7) & "'"
                s = s & ", qty2 = " & Val(Grid2.TextMatrix(i, 6))
                s = s & ", barcode = '" & Grid2.TextMatrix(i, 3) & "'"
                s = s & " Where barcode = '" & Grid2.TextMatrix(i - 1, 3) & "'"
                MsgBox s
                s = "Update rackpos set sku = '" & psku & "'"
                s = s & ", lot_num = '" & Grid2.TextMatrix(i, 5) & "'"
                s = s & ", count_qty = " & Val(Grid2.TextMatrix(i, 4))
                s = s & ", lot2 = '" & Grid2.TextMatrix(i, 7) & "'"
                s = s & ", qty2 = " & Val(Grid2.TextMatrix(i, 6))
                s = s & ", barcode = '" & Grid2.TextMatrix(i, 3) & "'"
                s = s & " Where barcode = '" & Grid2.TextMatrix(i - 1, 3) & "'"
                MsgBox s
                
                s = "Update queue_infc set sku = '" & psku & "'"
                s = s & ", lot_num = '" & Grid2.TextMatrix(i, 5) & "'"
                s = s & ", units = " & Val(Grid2.TextMatrix(i, 4))
                s = s & ", lot_num2 = '" & Grid2.TextMatrix(i, 7) & "'"
                s = s & ", units2 = " & Val(Grid2.TextMatrix(i, 6))
                s = s & ", palletid = '" & Grid2.TextMatrix(i, 3) & "'"
                s = s & " Where palletid = '" & Grid2.TextMatrix(i - 1, 3) & "'"
                MsgBox s
                s = "Update paltasks set product = '" & psku & " " & skurec(Val(psku)).prodname & "'"
                s = s & ", lotnum = '" & Grid2.TextMatrix(i, 5) & "'"
                s = s & ", units = " & Val(Grid2.TextMatrix(i, 4))
                s = s & ", lotnum2 = '" & Grid2.TextMatrix(i, 7) & "'"
                s = s & ", units2 = " & Val(Grid2.TextMatrix(i, 6))
                s = s & ", palletid = '" & Grid2.TextMatrix(i, 3) & "'"
                s = s & " Where palletid = '" & Grid2.TextMatrix(i - 1, 3) & "'"
                MsgBox s
                
                
            End If
            If Check3.Value = 1 Then
                
                afile = "U:\whsadj.001"
                'If Grid2.TextMatrix(i, 4) <> Grid2.TextMatrix(i - 1, 4) Or Grid2.TextMatrix(i, 5) <> Grid2.TextMatrix(i - 1, 5) Then
                If Grid2.TextMatrix(i, 4) <> Grid2.TextMatrix(i - 1, 4) Then    'Qty1 changed
                    If Grid2.TextMatrix(i, 5) = Grid2.TextMatrix(i - 1, 5) Then 'Same Lot
                        s = Format(Now, "M-d-yyyy") & Chr(9) & "500" & Chr(9) & "T10" & Chr(9)
                        s = s & Mid(Grid2.TextMatrix(i, 3), 5, 9) & Chr(9)
                        psku = Trim(Left(Grid2.TextMatrix(i, 3), 4))
                        s = s & psku & Chr(9)
                        s = s & skurec(Val(psku)).prodname & Chr(9)
                        s = s & Format(Val(Grid2.TextMatrix(i, 4)) - Val(Grid2.TextMatrix(i - 1, 4)), "0") & Chr(9)
                        s = s & "CHIN" & Chr(9) & WDUserId & Chr(9) & Format(Now, "M-d-yyyy")
                        Grid4.AddItem s
                        
                        cfile = "U:\whsadj.001"
                        Open cfile For Append Shared As #2
                        Write #2, Format(Now, "M-d-yyyy");
                        Write #2, "0500";
                        Write #2, "T10";
                        Write #2, Mid(Grid2.TextMatrix(i, 3), 5, 9);
                        psku = Trim(Left(Grid2.TextMatrix(i, 3), 4))
                        Write #2, psku;
                        Write #2, skurec(Val(psku)).prodname;
                        Write #2, Format(Val(Grid2.TextMatrix(i, 4)) - Val(Grid2.TextMatrix(i - 1, 4)), "0");
                        Write #2, "CHIN";
                        Write #2, WDUserId;
                        Write #2, Format(Now, "M-d-yyyy")
                        Close #2
                    End If
                End If
                If Grid2.TextMatrix(i, 5) <> Grid2.TextMatrix(i - 1, 5) Or Grid2.TextMatrix(i, 3) <> Grid2.TextMatrix(i - 1, 3) Then  'Lot1 changed
                    s = Format(Now, "M-d-yyyy") & Chr(9) & "500" & Chr(9) & "T10" & Chr(9)
                    s = s & Mid(Grid2.TextMatrix(i - 1, 3), 5, 9) & Chr(9)
                    psku = Trim(Left(Grid2.TextMatrix(i - 1, 3), 4))
                    s = s & psku & Chr(9)
                    s = s & skurec(Val(psku)).prodname & Chr(9)
                    s = s & Format(Val(Grid2.TextMatrix(i - 1, 4)) * -1, "0") & Chr(9)
                    s = s & "CHIN" & Chr(9) & WDUserId & Chr(9) & Format(Now, "M-d-yyyy")
                    Grid4.AddItem s
                
                    cfile = "U:\whsadj.001"
                    Open cfile For Append Shared As #2
                    Write #2, Format(Now, "M-d-yyyy");
                    Write #2, "500";
                    Write #2, "T10";
                    Write #2, Mid(Grid2.TextMatrix(i - 1, 3), 5, 9);
                    psku = Trim(Left(Grid2.TextMatrix(i - 1, 3), 4))
                    Write #2, psku;
                    Write #2, skurec(Val(psku)).prodname;
                    Write #2, Format(Val(Grid2.TextMatrix(i - 1, 4)) * -1, "0");
                    Write #2, "CHIN";
                    Write #2, WDUserId;
                    Write #2, Format(Now, "M-d-yyyy")
                    
                    s = Format(Now, "M-d-yyyy") & Chr(9) & "500" & Chr(9) & "T10" & Chr(9)
                    s = s & Mid(Grid2.TextMatrix(i, 3), 5, 9) & Chr(9)
                    psku = Trim(Left(Grid2.TextMatrix(i, 3), 4))
                    s = s & psku & Chr(9)
                    s = s & skurec(Val(psku)).prodname & Chr(9)
                    s = s & Format(Val(Grid2.TextMatrix(i, 4)), "0") & Chr(9)
                    s = s & "CHIN" & Chr(9) & WDUserId & Chr(9) & Format(Now, "M-d-yyyy")
                    Grid4.AddItem s
                    
                    Write #2, Format(Now, "M-d-yyyy");
                    Write #2, "500";
                    Write #2, "T10";
                    Write #2, Mid(Grid2.TextMatrix(i, 3), 5, 9);
                    psku = Trim(Left(Grid2.TextMatrix(i, 3), 4))
                    Write #2, psku;
                    Write #2, skurec(Val(psku)).prodname;
                    Write #2, Format(Val(Grid2.TextMatrix(i, 4)), "0");
                    Write #2, "CHIN";
                    Write #2, WDUserId;
                    Write #2, Format(Now, "M-d-yyyy")
                    Close #2
                End If
                
                
                If Grid2.TextMatrix(i, 6) <> Grid2.TextMatrix(i - 1, 6) Then    'Qty2 changed
                    If Grid2.TextMatrix(i, 7) = Grid2.TextMatrix(i - 1, 7) Then 'Same Lot
                        If Grid2.TextMatrix(i, 7) > "0" Then                    'Lot2 exists
                            s = Format(Now, "M-d-yyyy") & Chr(9) & "500" & Chr(9) & "T10" & Chr(9)
                            's = s & Mid(Grid2.TextMatrix(i - 1, 3), 5, 9) & Chr(9)
                            psku = Trim(Left(Grid2.TextMatrix(i, 3), 4))
                            plot = Left(Grid2.TextMatrix(i, 7), 5)
                            pcode = Mid(Grid2.TextMatrix(i, 7), 6, 3)
                            s = s & r12_lot(plot, pcode) & Chr(9)
                            s = s & psku & Chr(9)
                            s = s & skurec(Val(psku)).prodname & Chr(9)
                            s = s & Format(Val(Grid2.TextMatrix(i, 6)) - Val(Grid2.TextMatrix(i - 1, 6)), "0")
                            s = s & "CHIN" & Chr(9) & WDUserId & Chr(9) & Format(Now, "M-d-yyyy")
                            Grid4.AddItem s
                        
                            cfile = "U:\whsadj.001"
                            Open cfile For Append Shared As #2
                            Write #2, Format(Now, "M-d-yyyy");
                            Write #2, "500";
                            Write #2, "T10";
                            psku = Trim(Left(Grid2.TextMatrix(i, 3), 4))
                            plot = Left(Grid2.TextMatrix(i, 7), 5)
                            pcode = Mid(Grid2.TextMatrix(i, 7), 6, 3)
                            Write #2, r12_lot(plot, pcode);
                            'Write #2, Mid(Grid2.TextMatrix(i, 3), 5, 9);
                            Write #2, psku;
                            Write #2, skurec(Val(psku)).prodname;
                            Write #2, Format(Val(Grid2.TextMatrix(i, 6)) - Val(Grid2.TextMatrix(i - 1, 6)), "0");
                            Write #2, "CHIN";
                            Write #2, WDUserId;
                            Write #2, Format(Now, "M-d-yyyy")
                            Close #2
                        End If
                    End If
                End If
                If Grid2.TextMatrix(i, 7) <> Grid2.TextMatrix(i - 1, 7) Or Grid2.TextMatrix(i, 3) <> Grid2.TextMatrix(i - 1, 3) Then    'Lot2 changed
                    cfile = "U:\whsadj.001"
                    Open cfile For Append Shared As #2
                    If Grid2.TextMatrix(i - 1, 7) > "0" Then                    'Lot2 existed
                        s = Format(Now, "M-d-yyyy") & Chr(9) & "500" & Chr(9) & "T10" & Chr(9)
                        psku = Trim(Left(Grid2.TextMatrix(i - 1, 3), 4))
                        plot = Left(Grid2.TextMatrix(i - 1, 7), 5)
                        pcode = Mid(Grid2.TextMatrix(i - 1, 7), 6, 3)
                        s = s & r12_lot(plot, pcode) & Chr(9)
                        s = s & psku & Chr(9)
                        s = s & skurec(Val(psku)).prodname & Chr(9)
                        s = s & Format(Val(Grid2.TextMatrix(i - 1, 6)) * -1, "0") & Chr(9)
                        s = s & "CHIN" & Chr(9) & WDUserId & Chr(9) & Format(Now, "M-d-yyyy")
                        Grid4.AddItem s
                    
                        Write #2, Format(Now, "M-d-yyyy");
                        Write #2, "500";
                        Write #2, "T10";
                        psku = Trim(Left(Grid2.TextMatrix(i - 1, 3), 4))
                        plot = Left(Grid2.TextMatrix(i - 1, 7), 5)
                        pcode = Mid(Grid2.TextMatrix(i - 1, 7), 6, 3)
                        Write #2, r12_lot(plot, pcode);
                        'Write #2, Mid(Grid2.TextMatrix(i - 1, 3), 5, 9);
                        Write #2, psku;
                        Write #2, skurec(Val(psku)).prodname;
                        Write #2, Format(Val(Grid2.TextMatrix(i - 1, 6)) * -1, "0");
                        Write #2, "CHIN";
                        Write #2, WDUserId;
                        Write #2, Format(Now, "M-d-yyyy")
                    End If
                    If Grid2.TextMatrix(i, 7) > "0" Then                        'Lot2 exists
                        s = Format(Now, "M-d-yyyy") & Chr(9) & "500" & Chr(9) & "T10" & Chr(9)
                        psku = Trim(Left(Grid2.TextMatrix(i, 3), 4))
                        plot = Left(Grid2.TextMatrix(i, 7), 5)
                        pcode = Mid(Grid2.TextMatrix(i, 7), 6, 3)
                        s = s & r12_lot(plot, pcode) & Chr(9)
                        s = s & psku & Chr(9)
                        s = s & skurec(Val(psku)).prodname & Chr(9)
                        s = s & Format(Val(Grid2.TextMatrix(i, 6)), "0") & Chr(9)
                        s = s & "CHIN" & Chr(9) & WDUserId & Chr(9) & Format(Now, "M-d-yyyy")
                        Grid4.AddItem s
                    
                        Write #2, Format(Now, "M-d-yyyy");
                        Write #2, "500";
                        Write #2, "T10";
                        psku = Trim(Left(Grid2.TextMatrix(i, 3), 4))
                        plot = Left(Grid2.TextMatrix(i, 7), 5)
                        pcode = Mid(Grid2.TextMatrix(i, 7), 6, 3)
                        Write #2, r12_lot(plot, pcode);
                        'Write #2, Mid(Grid2.TextMatrix(i, 3), 5, 9);
                        Write #2, psku;
                        Write #2, skurec(Val(psku)).prodname;
                        Write #2, Format(Val(Grid2.TextMatrix(i, 6)), "0");
                        Write #2, "CHIN";
                        Write #2, WDUserId;
                        Write #2, Format(Now, "M-d-yyyy")
                    End If
                    Close #2
                End If
                Grid4.FormatString = "^Tran Date|^Whs|^Locn|<Lot|^Item|<Description|^Qty|^Reason|^User|^Entry Date"
                Grid4.ColWidth(0) = 1000
                Grid4.ColWidth(1) = 600
                Grid4.ColWidth(2) = 600
                Grid4.ColWidth(3) = 1200
                Grid4.ColWidth(4) = 600
                Grid4.ColWidth(5) = 2500
                Grid4.ColWidth(6) = 800
                Grid4.ColWidth(7) = 800
                Grid4.ColWidth(8) = 1000
                Grid4.ColWidth(9) = 1000
                
            End If
        End If
        'If Check1.Value = 1 And Grid2.TextMatrix(i, 11) = "Post" Then
        If Grid2.TextMatrix(i, 11) = "Post" Then
            
            psku = Trim(Left(Grid2.TextMatrix(i, 3), 4))
            pname = skurec(Val(psku)).prodname
            p.id = Grid2.TextMatrix(i, 1)
            p.area = "Correction"
            If Len(preas) > 0 Then
                p.description = preas
            Else
                p.description = pref.Caption '" "
            End If
            p.source = Grid2.TextMatrix(i, 8)
            p.target = Grid2.TextMatrix(i, 9)
            'p.product = proddesc.Caption
            p.product = psku & " " & pname
            p.palletid = Grid2.TextMatrix(i, 3)
            If Grid2.TextMatrix(i, 0) = "Current" Then
                p.qty = "-1"
                p.units = Val(Grid2.TextMatrix(i, 4)) * -1
                p.units2 = Val(Grid2.TextMatrix(i, 6)) * -1
            Else
                p.qty = "1"
                p.units = Val(Grid2.TextMatrix(i, 4))
                p.units2 = Val(Grid2.TextMatrix(i, 6))
            End If
            p.uom = "Pallet"
            p.lotnum = Grid2.TextMatrix(i, 5)
            p.lotnum2 = Grid2.TextMatrix(i, 7)
            p.status = "COMP"
            p.userid = Form1.userid
            p.trandate = Format(Now, "yyMMdd hh:mm:ss")
            p.reqid = Grid2.TextMatrix(i, 2)
            If Check1.Value = 1 Then
                
                'cfile = logdir & "move" & Format(Now, "MMddyyyy") & ".txt"
                cfile = "U:\recv" & Format(Now, "MMddyyyy") & ".txt"
                Open cfile For Append Shared As #1
                Write #1, p.id, p.area, p.description, p.source, p.target, p.product;
                Write #1, p.palletid, p.qty, p.uom, p.lotnum, p.units, p.lotnum2, p.units2;
                'Write #1, p.status, p.userid, p.trandate, p.reqid
                Write #1, p.status, WDUserId, p.trandate, p.reqid                   'jv121614
                Close #1
            Else
                If Grid2.TextMatrix(i, 0) = "New" Or Grid2.TextMatrix(i, 3) <> Grid2.TextMatrix(i - 1, 3) Then
                    'cfile = logdir & "wms" & Format(Now, "MMddyyyy") & ".txt"
                    cfile = "U:\wms" & Format(Now, "MMddyyyy") & ".txt"
                    Open cfile For Append Shared As #1
                    Write #1, p.id, p.area, p.description, p.source, p.target, p.product;
                    Write #1, p.palletid, p.qty, p.uom, p.lotnum, p.units, p.lotnum2, p.units2;
                    'Write #1, p.status, p.userid, p.trandate, p.reqid
                    Write #1, p.status, WDUserId, p.trandate, p.reqid                   'jv121614
                    Close #1
                End If
            End If
            Lfile = cfile
            If InStr(1, Lfile, "recv") > 0 Then
                s = "PR"
            Else
                s = "WM"
            End If
            s = s & Chr(9) & p.id & Chr(9) & p.area & Chr(9) & p.description & Chr(9)
            s = s & p.source & Chr(9) & p.target & Chr(9) & p.product & Chr(9)
            s = s & p.palletid & Chr(9) & p.qty & Chr(9) & p.uom & Chr(9)
            s = s & p.lotnum & Chr(9) & p.units & Chr(9) & p.lotnum2 & Chr(9)
            s = s & p.units2 & Chr(9) & p.status & Chr(9) & WDUserId & Chr(9)
            s = s & p.trandate & Chr(9) & p.reqid
            Grid3.AddItem s
            Grid3.FormatString = s
            s = "^Type|^RecId|<Area|<Description|<Source|<Target|<Product|^Pallet|^Qty|^Uom|^LotNum|^Units|^LotNum|^Units|^Status|^User|<Time|^ReqId"
            Grid3.FormatString = s
            Grid3.FormatString = s
            Grid3.ColWidth(0) = 600
            Grid3.ColWidth(1) = 600
            Grid3.ColWidth(2) = 1300
            Grid3.ColWidth(3) = 1000
            Grid3.ColWidth(4) = 1300
            Grid3.ColWidth(5) = 1300
            Grid3.ColWidth(6) = 3000
            Grid3.ColWidth(7) = 1800
            Grid3.ColWidth(8) = 600
            Grid3.ColWidth(9) = 800
            Grid3.ColWidth(10) = 800
            Grid3.ColWidth(11) = 800
            Grid3.ColWidth(12) = 800
            Grid3.ColWidth(13) = 800
            Grid3.ColWidth(14) = 800
            Grid3.ColWidth(15) = 1000
            Grid3.ColWidth(16) = 1400
            Grid3.ColWidth(17) = 1000
            'Grid1.ColWidth(18) = 1
            
        End If
            
    Next i
    If Grid3.Rows > 1 Then
        Grid3.Visible = True: Lfile.Visible = True
    Else
        Grid3.Visible = False: Lfile.Visible = False
    End If
    If Grid4.Rows > 1 Then
        Grid4.Visible = True: afile.Visible = True
    Else
        Grid4.Visible = False: afile.Visible = False
    End If
    refresh_grid_barcode
End Sub

Private Sub edval_Click()
    Dim s As String, k As Long, i As Integer
    If Grid1.Col <> 2 Then Exit Sub
    s = Grid1.Text
    s = InputBox(Grid1.TextMatrix(Grid1.Row, 0), "Edit value...", s)
    If Len(s) = 0 Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 0) = "BarCode" Then
        If Len(s) <> 16 Then Exit Sub
        k = Val(Trim(Left(s, 4)))
        If k = 0 Then Exit Sub
        If Val(skurec(k).sku) <> k Then Exit Sub
        If Mid(s, 11, 3) < "100" Or Mid(s, 11, 3) > "599" Then Exit Sub
        If Val(Mid(s, 14, 3)) < 0 And Mid(s, 14, 3) <> "EOR" Then Exit Sub
    End If
    If Grid1.TextMatrix(Grid1.Row, 0) = "Qty1" And Val(s) < 0 Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 0) = "Lot1" Then
        If Len(s) <> 5 Then Exit Sub
        If Val(s) = 0 Then Exit Sub
    End If
    If Grid1.TextMatrix(Grid1.Row, 0) = "Qty2" And Val(s) < 0 Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 0) = "Lot2" Then
        If UCase(s) = "LOT1" And Check3.Value = 1 Then
            s = "LOT1"
        Else
            If s > "0" And Len(s) <> 8 Then Exit Sub
        End If
        'If Val(s) = 0 Then Exit Sub
    End If
    Grid1.Text = s
    k = Val(Grid1.TextMatrix(Grid1.Row, 4))
    For i = 1 To Grid2.Rows - 1
        If Val(Grid2.TextMatrix(i, 1)) = k Then
            Grid2.TextMatrix(i, 11) = "Post"
            If Grid2.TextMatrix(i, 0) = "New" Then
                If Grid1.TextMatrix(Grid1.Row, 0) = "BarCode" Then Grid2.TextMatrix(i, 3) = s
                If Grid1.TextMatrix(Grid1.Row, 0) = "Qty1" Then Grid2.TextMatrix(i, 4) = s
                If Grid1.TextMatrix(Grid1.Row, 0) = "Lot1" Then Grid2.TextMatrix(i, 5) = s
                If Grid1.TextMatrix(Grid1.Row, 0) = "Qty2" Then Grid2.TextMatrix(i, 6) = s
                If Grid1.TextMatrix(Grid1.Row, 0) = "Lot2" Then Grid2.TextMatrix(i, 7) = s
            End If
        End If
    Next i
    Frame1.Visible = True
    Grid3.Visible = False: Lfile.Visible = False
    Grid4.Visible = False: afile.Visible = False
End Sub

Private Sub Form_Resize()
    Grid2.Width = Me.Width - 180
    Grid3.Width = Me.Width - 180
    Lfile.Width = Me.Width - 180
    Grid4.Width = Me.Width - 180
    afile.Width = Me.Width - 180
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And Grid1.TextMatrix(Grid1.Row, 3) = "*" Then PopupMenu edmenu
End Sub

Private Sub plkey_Change()
    Text2 = plkey
    refresh_grid_plate
End Sub

Private Sub pref_Change()
    'If pref = "Production" Or pref = "Traffic Master" Then
    '    Check1.Value = 1: Check1.Enabled = True
    '    Check2.Value = 1
    'Else
    '    Check1.Value = 0: Check1.Enabled = False
    '    Check2.Value = 0
    'End If
End Sub

