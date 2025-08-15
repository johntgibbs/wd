VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form21 
   Caption         =   "Production Schedule"
   ClientHeight    =   10380
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12015
   LinkTopic       =   "Form21"
   ScaleHeight     =   10380
   ScaleWidth      =   12015
   StartUpPosition =   3  'Windows Default
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
      Left            =   1440
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   9975
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   17595
      _Version        =   327680
      BackColorFixed  =   12648447
      FocusRect       =   0
   End
   Begin VB.Label gcolor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label hcolor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Marked On Hold"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu markhold 
         Caption         =   "Mark On Hold"
      End
      Begin VB.Menu remhold 
         Caption         =   "Remove On Hold"
      End
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub refresh_grid1()
    Dim s As String, f0 As String, f1 As String, f2 As String
    Dim f3 As String, f4 As String, f5 As String, f6 As String
    Dim f7 As String, cfile As String, psku As String, plot As String
    Dim i As Integer, k As Integer
    If Form1.plantno = "50" Then cfile = "s:\wd\data\plabels.500"
    If Form1.plantno = "51" Then cfile = "\\bbba-03-dc\f\user\waredist\data\plabels.501"
    If Form1.plantno = "52" Then cfile = "\\bbsy-02-dc\f\user\waredist\data\plabels.502"
    Grid1.Redraw = False
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 8
    If Len(Dir(cfile)) = 0 Then Exit Sub
    Open cfile For Input As #1
    Do Until EOF(1)
        Input #1, f0, f1, f2, f3, f4, f5, f6, f7
        's = f0 & Chr(9)         'Plant
        's = s & f1 & Chr(9)     'Production Date
        'If f1 = Combo1 Then
            s = f1 & Chr(9)
            s = s & f2 & Chr(9)         'SKU
            s = s & f3 & Chr(9)     'Description
            s = s & f4 & Chr(9)     'Pallet Qty
            s = s & f5 & Chr(9)     'Code Date
            s = s & f6 & Chr(9)     'OP Code
            s = s & f7 & Chr(9)             'Operation
            If f1 = Combo1 Then s = s & "*"
            Grid1.AddItem s
        'End If
    Loop
    Close #1
    s = "^Date|^SKU|<Description|^Pallets|^Code Date|^Op Code|<Operation|^Status"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 600
    Grid1.ColWidth(2) = 2800
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1000
    Grid1.ColWidth(6) = 2000
    Grid1.ColWidth(7) = 800
    hcolor.Visible = False
    Grid1.FillStyle = flexFillRepeat
    'Highlight Production Dates
    If Grid1.Rows > 2 Then
        s = Grid1.TextMatrix(1, 0)
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 0) <> s Then
                hrow = Not hrow
                s = Grid1.TextMatrix(i, 0)
            End If
            If hrow = True Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = gcolor.BackColor
            End If
        Next i
    End If
    'Highlight On Hold
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            psku = Grid1.TextMatrix(i, 1)
            'plot = Grid1.TextMatrix(i, 4) & " " & Grid1.TextMatrix(i, 5)
            plot = Grid1.TextMatrix(i, 4) & Grid1.TextMatrix(i, 5)                      'jv052515
            For k = 0 To holdlist.Grid5.Rows - 1
                If holdlist.Grid5.TextMatrix(k, 1) = psku And holdlist.Grid5.TextMatrix(k, 8) = plot Then
                    Grid1.Row = i: Grid1.RowSel = i
                    Grid1.Col = 1: Grid1.ColSel = 7
                    Grid1.CellBackColor = hcolor.BackColor
                    hcolor.Visible = True
                    Grid1.TextMatrix(i, 7) = "H"
                    Exit For
                End If
            Next k
        Next i
        Grid1.Row = 1
    End If
    Grid1.Redraw = True
End Sub

Private Sub Command1_Click()
    refresh_grid1
End Sub

Private Sub Form_Load()
    refresh_grid1
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 100
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1080
End Sub

Private Sub grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid1_RowColChange()
    If Grid1.TextMatrix(Grid1.Row, 7) = "H" Then
        markhold.Enabled = False
        remhold.Enabled = True
    Else
        markhold.Enabled = True
        remhold.Enabled = False
    End If
End Sub

Private Sub markhold_Click()                        'Mark on hold
    Dim bc As String, i As Integer, zid As Long, s As String
    Dim psku As String, plot As String, pcode As String
    i = Grid1.Row
    psku = Grid1.TextMatrix(i, 1)
    pcode = Grid1.TextMatrix(i, 5)
    If Len(pcode) = 0 Or Len(pcode) > 3 Then Exit Sub           'jv052515
    If Len(pcode) = 1 Then                                      'jv052515
        pcode = " " & pcode & " "                               'jv052515
    Else                                                        'jv052515
        If Len(pcode) = 2 Then pcode = " " & pcode              'jv052515
    End If                                                      'jv052515
    
    'bc = psku & " " & Grid1.TextMatrix(i, 4) & " "
    'bc = bc & pcode & " 001"
    If Len(psku) = 3 Then psku = psku & " "                     'jv082415
    'bc = psku & " " & Grid1.TextMatrix(i, 4) & pcode & "001"    'jv052515
    bc = psku & Grid1.TextMatrix(i, 4) & pcode & "001"     'jv082415
    pcode = Trim(pcode)                                         'jv052515
    plot = barcode_to_lotnum(bc)
    zid = wd_seq("HoldList")                                      'jv042015
    s = "Insert into holdlist (id, sku, lot_num, opcode, spallet, epallet, hsource, userid, holddate) values (" & zid
    's = s & ", '" & Trim(psku) & "', '" & plot & "', '" & pcode & "', '001', 'EOR', 'Schedule', '" & WDUserId & "', '" & Format(Now, "yyMMdd hh:mm:ss") & "')"
    s = s & ", '" & Trim(psku) & "', '" & plot & "', '" & pcode & "', '001', 'EOR', 'TEST HOLD', '" & WDUserId & "', '" & Format(Now, "yyMMdd hh:mm:ss") & "')"     'jv122115
    'MsgBox s
    Wdb.Execute s
    holdlist.refresh_holdlist
    DoEvents
    refresh_grid1
    DoEvents
    If i <= Grid1.Rows - 1 Then Grid1.Row = i
End Sub

Private Sub remhold_Click()                         'Remove from Hold List
    Dim bc As String, i As Integer, s As String
    Dim psku As String, plot As String, pcode As String
    i = Grid1.Row
    psku = Grid1.TextMatrix(i, 1)
    pcode = Grid1.TextMatrix(i, 5)
    If Len(pcode) = 0 Or Len(pcode) > 3 Then Exit Sub           'jv052515
    If Len(pcode) = 1 Then                                      'jv052515
        pcode = " " & pcode & " "                               'jv052515
    Else                                                        'jv052515
        If Len(pcode) = 2 Then pcode = " " & pcode              'jv052515
    End If                                                      'jv052515
    
    'bc = psku & " " & Grid1.TextMatrix(i, 4) & " "
    'bc = bc & pcode & " 001"
    If Len(psku) = 3 Then psku = psku & " "                     'jv082415
    'bc = psku & " " & Grid1.TextMatrix(i, 4) & pcode & "001"    'jv052515
    bc = psku & Grid1.TextMatrix(i, 4) & pcode & "001"          'jv082415
    pcode = Trim(pcode)                                         'jv052515
    
    plot = barcode_to_lotnum(bc)
    s = "delete from holdlist where sku = '" & Trim(psku) & "'" 'jv082415
    s = s & " and lot_num = '" & plot & "'"
    s = s & " and opcode = '" & pcode & "'"
    'MsgBox s
    Wdb.Execute s
    holdlist.refresh_holdlist
    DoEvents
    refresh_grid1
    DoEvents
    If i <= Grid1.Rows - 1 Then Grid1.Row = i
End Sub
