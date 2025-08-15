VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Blue Bell Pallets - BA Wrapper"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   ScaleHeight     =   8535
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   8040
      TabIndex        =   26
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label oppic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "opcode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   0
         TabIndex        =   34
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label name3pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "name 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   0
         TabIndex        =   33
         Top             =   2880
         Width           =   2535
      End
      Begin VB.Label name2pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "name 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label name1pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "name 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   0
         TabIndex        =   31
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Label pkgpic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "pkg"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   0
         TabIndex        =   30
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label palnopic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pallet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   0
         TabIndex        =   29
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lotpic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Codedate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   0
         TabIndex        =   28
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label skupic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SKU"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   0
         TabIndex        =   27
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.ListBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1860
      Left            =   120
      TabIndex        =   25
      Top             =   5880
      Width           =   7575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Scanned Pallets:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   5400
      Width           =   7575
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6240
      TabIndex        =   23
      Text            =   "Combo4"
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   7320
      TabIndex        =   20
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox emess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "babarcode7.frx":0000
      Top             =   8640
      Width           =   6735
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Re-Enter Pallet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   13
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Destination:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Plate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label unitswrap 
      Caption         =   "unitswrap"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label unitspal 
      Caption         =   "unitspal"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label wrapspal 
      Caption         =   "wrapspal"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "2nd Wrap Qty:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "2nd Code Date:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Wrap Qty:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "BarCode:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label apphdr 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6855
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub draw_label(bc As String)
    Dim i As Integer
    'i = Val(Mid(bc, 1, 3))
    i = Val(Mid(bc, 1, 4))                          'jv062816
    bc = UCase(bc)
    skupic.Caption = Trim(Mid(bc, 1, 4))
    'lotpic.Caption = Mid(bc, 5, 8)
    lotpic.Caption = Mid(bc, 5, 6)                      'jv052515
    oppic.Caption = Mid(bc, 11, 3)                      'jv082415
    palnopic.Caption = Mid(bc, 14, 3)
    If Val(palnopic.Caption) > 0 Then palnopic = Format(Val(palnopic.Caption), "0")
    pkgpic.Caption = labpix(i).package
    name1pic.Caption = labpix(i).name1
    name2pic.Caption = labpix(i).name2
    name3pic.Caption = labpix(i).name3
    Frame1.Visible = True
    emess.Visible = False
End Sub

Private Sub poll_pallets()
    Dim PauseTime, Start, Finish, TotalTime
    'If (MsgBox("Press Yes to pause for 5 seconds", 4)) = vbYes Then
    'PauseTime = 5   ' Set duration.
    'Start = Timer   ' Set start time.
    'Do While Timer < Start + PauseTime
    '    DoEvents    ' Yield to other processes.
    'Loop
    'Finish = Timer  ' Set end time.
    'TotalTime = Finish - Start  ' Calculate total time.
    'MsgBox "Paused for " & TotalTime & " seconds"
    'Else
        'End
    'End If
    
    Do While True
        Start = Timer
        Do While Timer < Start + 5 '10
            DoEvents
        Loop
        Command4_Click
    Loop

End Sub

Private Sub record_pallet(pno As String, p As ptask, pwhs As String, pstat As String)
    'Dim db As ADODB.Connection,
    Dim ds As ADODB.Recordset, s As String
    Dim pid As Long, psku As String, recid As Long
    Screen.MousePointer = 11
    'pstat = "Wrapper"
    'pstat = "Warehouse"
    'pstat = "Test"
    psku = Trim(Left(p.palletid, 4))
    recid = 0
    'Set db = CreateObject("ADODB.Connection")
    ''db.Open Form1.bbsr
    'db.Open Form1.bbsr
    
    s = "select * from pallets where barcode = '" & p.palletid & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        recid = ds!id
    Else
        ds.Close
        s = "select * from pallets where status in ('Shipped','Order Pick')"
        s = s & " order by trandate"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            recid = ds!id
        End If
    End If
    ds.Close
    If recid > 0 Then
        's = "Update pallets set plateno = '" & Trim(pno) & "'"
        s = "Update pallets set plateno = '" & Format(Val(pno), "000000") & "'"
        s = s & ",barcode = '" & p.palletid & "'"
        s = s & ",qty1 = " & Val(p.units)
        s = s & ",lot1 = '" & p.lotnum & "'"
        s = s & ",qty2 = " & Val(p.units2)
        s = s & ",lot2 = '" & p.lotnum2 & "'"
        s = s & ",source = '" & p.source & "'"
        's = s & ",target = '" & Grid1.TextMatrix(i, 4) & "'"
        's = s & ",target = 'SR" & pwhs & "'"
        s = s & ",target = '" & pwhs & "'"
        s = s & ",bbc = 'Y'"
        s = s & ",status = '" & pstat & "'"
        s = s & ",trandate = '" & p.trandate & "'"
        s = s & ",sku = '" & psku & "'"
        s = s & " Where id = " & recid
        'MsgBox s
        Wdb.Execute s
    Else
        pid = wd_seq("Pallets")
        s = "Insert Into pallets Values (" & pid
        's = s & ",'" & Trim(pno) & "'"
        s = s & ",'" & Format(Val(pno), "000000") & "'"
        s = s & ",'" & p.palletid & "'"
        s = s & "," & Val(p.units)
        s = s & ",'" & p.lotnum & "'"
        s = s & "," & Val(p.units2)
        s = s & ",'" & p.lotnum2 & "'"
        s = s & ",'" & p.source & "'"
        's = s & ",'" & Grid1.TextMatrix(i, 4) & "'"
        's = s & ",'SR" & pwhs & "'"
        s = s & ",'" & pwhs & "'"
        s = s & ",'Y'"
        If p.target = "ORDER PICK" Then
            s = s & ",'Order Pick'"
        Else
            s = s & ",'" & pstat & "'"
        End If
        s = s & ",'" & p.trandate & "'"
        s = s & ",'" & psku & "')"
        Wdb.Execute s
    End If
    'db.Close
    Screen.MousePointer = 0
End Sub


Private Sub barcode_scanned(bc As String)
    'Dim db As adodb.Connection,
    Dim ds As ADODB.Recordset, s As String
    Dim i As Integer, cd As String, ssku As String, cc As String, td As String
    If Len(bc) < 15 Then Exit Sub
    'Test for previous scan
    If Combo1.ListCount > 0 Then
        For i = 0 To Combo1.ListCount - 1
            If Left(Combo1.List(i), 16) = bc Then
                emess = bc & " has already been scanned."
                Combo1.ListIndex = i
                Text1 = ""
                Text1.SetFocus
                Exit Sub
            End If
            'If Len(Text4) = 6 Then
            'If Mid(Combo1.List(i), 18, 6) = Format(Val(Text4), "000000") Then
            '    emess = "Plate " & Text4 & " has already been scanned."
            '    Combo1.ListIndex = i
            '    Text4 = ""
            '    Text4.SetFocus
            '    Exit Sub
            'End If
            'End If
        Next i
    End If
    
    'Build 2nd code date list
    cd = Mid(bc, 5, 2) & "-" & Mid(bc, 7, 2) & "-20" & Mid(bc, 9, 2)
    td = Format(DateAdd("yyyy", 2, Now), "MM-dd-yyyy")
    'td = Left(td, 8) & Mid(bc, 9, 2)
    'MsgBox td & " " & DateDiff("d", cd, td)
    Combo2.Clear
    Combo2.AddItem " "
    If cd = "02-29-2018" Then cd = "02-28-2018"
    'For i = 1 To 10
    If Mid(cd, 1, 5) = "02-29" Then
        Combo2.AddItem "0229" & Right(cd, 2)
    Else
        For i = 0 To DateDiff("d", cd, td)
            Combo2.AddItem Format(DateAdd("d", i, cd), "MMddyy")
        Next i
    End If
    
    'Check sku and get wrap qty
    ssku = Trim(Left(bc, 4))
    wrapspal = "0"
    unitspal = "0"
    unitswrap = "0"
    'Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    s = "select uom_per_pallet, qty_per_pallet, uom_type, description from sku_config where sku = '" & ssku & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        wrapspal = ds!qty_per_pallet
        unitspal = ds!uom_per_pallet
        unitswrap = ds!uom_per_pallet / ds!qty_per_pallet
        Text2 = ds!qty_per_pallet
        emess = ds!uom_type & " " & ds!description
        Text3 = ""
        'Text4.SetFocus
        Command1.Enabled = True: Command1.SetFocus
    Else
        emess = "Invalid SKU: " & ssku & " found in the barcode."
        ds.Close ': db.Close
        Text1 = ""
        Text2 = ""
        Text3 = ""
        Text1.SetFocus
        Exit Sub
    End If
    ds.Close ': db.Close
    Command1.Enabled = True
    'cc = Mid(bc, 12, 1)
    cc = Mid(bc, 11, 3)                 'jv052515
    Call draw_label(Text1)
    'MsgBox cc
    For i = 0 To Combo4.ListCount - 1
        If Combo4.List(i) = cc Then
            Combo4.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub refresh_robot0_pallets()
    'Dim db As adodb.Connection,
    Dim ds As ADODB.Recordset, s As String
    Combo1.Visible = False
    Combo1.Clear: List1.Clear
    'Set db = CreateObject("ADODB.Connection")
    'db.Open Form1.bbsr
    s = "select palletid, qty, uom, id, reqid, target from paltasks where area = 'FORKLIFT'"
    's = s & " and source in ('ROBOT ZERO')"
    s = s & " and source in ('WRAPPER')"
    s = s & " and status = 'PEND'"
    s = s & " order by trandate desc"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            'Combo1.AddItem ds(0) & " " & Format(Val(ds(4)), "000000") & " " & ds(1) & " " & ds(2) & " " & StrConv(ds(5), vbProperCase)
            'Combo1.AddItem ds(0) & " " & Format(Val(Left(ds(4), 6)), "000000") & " " & ds(1) & " " & ds(2) & " " & StrConv(ds(5), vbProperCase)
            If Len(ds(4)) > 6 Then                  'jv010616
                Combo1.AddItem ds(0) & " 000000 " & ds(1) & " " & ds(2) & " " & StrConv(ds(5), vbProperCase)
            Else                                    'jv010616
                Combo1.AddItem ds(0) & " " & Format(Val(Left(ds(4), 6)), "000000") & " " & ds(1) & " " & ds(2) & " " & StrConv(ds(5), vbProperCase)
            End If                                  'jv010616
            List1.AddItem ds!id
            ds.MoveNext
        Loop
        Combo1.ListIndex = 0
    End If
    Combo1.Visible = True
    ds.Close ': db.Close
End Sub

Private Sub refresh_plate(pno As Long)
    'Dim ds As adodb.Recordset, s As String
    'If pno = 0 Then
    '    s = "select sequence_id from sequences where seq = 'BAWrapper'"
    '    Set ds = Wdb.Execute(s)
    '    If ds.BOF = False Then
    '        ds.MoveFirst
    '        'Text4 = Format(ds(0), "000000")
    '        Label8 = Format(ds(0) + 1, "000000")
    '    End If
    '    ds.Close
    'Else
    '    s = "update sequences set sequence_id = " & pno & " where seq = 'BAWrapper'"
    '    Wdb.Execute (s)
    '    Label8 = Format(pno + 1, "000000")
    'End If
    Label8.Caption = "000000"
    Text4 = "000000"
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Combo2_Click()
    Call Text2_Change
End Sub

Private Sub Command1_Click()
    Dim p As ptask, s As String, psku As String, wcnt As Integer
    'If Len(Text4) <> 6 Then
    '    MsgBox "Invalid or Missing Plate...", vbOKOnly + vbExclamation, "Sorry, try again..."
    '    Text4.SetFocus
    '    Exit Sub
    'End If
    'Test for previous scan
    If Combo1.ListCount > 0 Then
        For i = 0 To Combo1.ListCount - 1
            If Left(Combo1.List(i), 16) = bc Then
                emess = bc & " has already been scanned."
                Combo1.ListIndex = i
                Text1 = ""
                Text1.SetFocus
                Exit Sub
            End If
            'If Mid(Combo1.List(i), 18, 6) = Format(Val(Text4), "000000") Then
            '    emess = "Plate " & Text4 & " has already been scanned."
            '    Combo1.ListIndex = i
            '    Text4 = ""
            '    Text4.SetFocus
            '    Exit Sub
            'End If
        Next i
    End If
    s = Text1 & " " & Text4 & " " & Text2 & " Wraps NEW"
    Combo1.AddItem s, 0
    'MsgBox s
    psku = Trim(Left(Text1, 4))
    'p.area = "ROBOT ZERO"
    p.area = "WRAPPER"
    p.description = " "
    'p.source = "ROBOT ZERO"
    p.source = "WRAPPER"
    p.target = Combo3
    p.product = psku & " " & sku_info(psku, "desc")
    p.palletid = Text1
    p.qty = Val(Text2) + Val(Text3)
    p.uom = "Wraps"
    p.lotnum = barcode_to_lotnum(Text1)
    wcnt = sku_info(psku, "units")
    wcnt = wcnt / sku_info(psku, "wraps")
    p.units = Val(Text2) * wcnt
    If Combo2 > " " Then
        s = Left(Text1, 4) & Combo2 & Right(Text1, 6)
        'p.lotnum2 = barcode_to_lotnum(s) & " " & Combo4
        p.lotnum2 = barcode_to_lotnum(s) & Combo4               'jv052515
        p.units2 = Val(Text3) * wcnt
    Else
        p.lotnum2 = " "
        p.units2 = "0"
    End If
    p.status = "PEND"
    p.userid = Form1.userid  '"131052"
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")
    p.reqid = Text4
    p.id = insert_trans(p)
    'Call spt_to_dock(p)
    Call robot0_pickup(p)
    p.status = "COMP"
    p.userid = " "
    p.reqid = Text4
    Call update_trans(p)
    
    'Call record_pallet(Text4, p, Combo3, "Wrapper")
    'MsgBox p.id
    refresh_robot0_pallets
    'Combo1.AddItem Text1 & " " & Text2 & " Wraps"
    emess.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Combo2.ListIndex = 0
    Call refresh_plate(Val(Text4))
    DoEvents
    Text4.Text = "000000"
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    Dim s As String, i As Integer, p As ptask
    If Combo1 > "0" Then s = Left(Combo1, 16)
    If Val(List1) > 0 Then
        p = masterec(Val(List1))
        Call return_to_wrapper(p.palletid, Form1.userid, p.area, p.reqid)
    End If
    If Combo1.ListCount > 1 Then
        i = Combo1.ListIndex
        Combo1.RemoveItem i
        List1.RemoveItem i
        DoEvents
        Combo1.ListIndex = 0
    Else
        Combo1.Clear
        List1.Clear
    End If
    Text1 = s
End Sub

Private Sub Command3_Click()
    Text4 = Label8
End Sub

Private Sub Command4_Click()
    refresh_robot0_pallets
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    emess.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = "000000"
    Combo3.Clear
    'Combo3.AddItem "ROBOT ZERO"
    Combo3.AddItem "WRAPPER"
    'Combo3.AddItem "1405"
    'Combo3.AddItem "1406"
    'Combo3.AddItem "1731"
    'Combo3.AddItem "SNACK PLANT"
    Combo3.ListIndex = 0
    Combo4.Clear
    For i = 100 To 199              'jv052515
        Combo4.AddItem i            'jv052515
    Next i                          'jv052515
    'Combo4.AddItem "A": Combo4.AddItem "B": Combo4.AddItem "C": Combo4.AddItem "D": Combo4.AddItem "E"
    'Combo4.AddItem "F": Combo4.AddItem "G": Combo4.AddItem "H": Combo4.AddItem "I": Combo4.AddItem "J"
    'Combo4.AddItem "K": Combo4.AddItem "L": Combo4.AddItem "M": Combo4.AddItem "N": Combo4.AddItem "O"
    'Combo4.AddItem "P": Combo4.AddItem "Q": Combo4.AddItem "R": Combo4.AddItem "S": Combo4.AddItem "T"
    'Combo4.AddItem "U": Combo4.AddItem "V": Combo4.AddItem "W": Combo4.AddItem "X": Combo4.AddItem "Y"
    'Combo4.AddItem "Z"
    Combo4.ListIndex = 0
    'apphdr = "ROBOT ZERO"
    apphdr = "WRAPPER"
    refresh_robot0_pallets
    Call refresh_plate(0)
    'poll_pallets
End Sub

Private Sub Form_Resize()
    Combo1.Left = (Me.Width - Combo1.Width) * 0.5
    apphdr.Left = Combo1.Left
    Label1.Left = Combo1.Left
    Text1.Left = Label1.Left + Label1.Width
    
    Label6.Left = Combo1.Left
    Text4.Left = Label6.Left + Label6.Width
    Command3.Left = Text4.Left + Text4.Width + 600
    'Label8.Left = Text4.Left + Text4.Width
    Label8.Left = Command3.Left + Command3.Width
    
    Label7.Left = Combo1.Left
    Combo3.Left = Label7.Left + Label7.Width
    
    Label2.Left = Combo1.Left
    Text2.Left = Label2.Left + Label2.Width
    
    Label3.Left = Combo1.Left
    Combo2.Left = Label3.Left + Label3.Width
    Label5.Left = Combo2.Left + Combo2.Width
    Combo4.Left = Label5.Left + Label5.Width
    
    
    Label4.Left = Combo1.Left
    Text3.Left = Label4.Left + Label4.Width
    'Label5.Left = Combo1.Left
    Command4.Left = Combo1.Left
    Command1.Left = (Me.Width - Command1.Width) * 0.5
    Command2.Left = (Me.Width - Command2.Width) * 0.5
    emess.Left = (Me.Width - emess.Width) * 0.5
    Frame1.Left = (Me.Width - Frame1.Width) * 0.5
    Frame1.Top = emess.Top
End Sub

Private Sub Frame1_Click()
    If Command1.Enabled Then Command1_Click
End Sub

Private Sub lotpic_Click()
    Frame1_Click
End Sub

Private Sub name1pic_Click()
    Frame1_Click
End Sub

Private Sub name2pic_Click()
    Frame1_Click
End Sub

Private Sub name3pic_Click()
    Frame1_Click
End Sub

Private Sub palnopic_Click()
    Frame1_Click
End Sub

Private Sub pkgpic_Click()
    Frame1_Click
End Sub

Private Sub skupic_Click()
    Frame1_Click
End Sub

Private Sub Text1_Change()
    Frame1.Visible = False
    emess.Visible = True
    If Len(Text1) > 15 Then
        Text1 = UCase(Text1)
        Call barcode_scanned(Text1)
        'Command3.SetFocus
        If Command1.Enabled = True Then Command1.SetFocus       'jv012915
        'Text4.SetFocus
    Else
        Command1.Enabled = False
    End If
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub Text2_Change()
    If Combo2 > " " Then Text3 = Val(wrapspal) - Val(Text2)
End Sub

Private Sub Text3_Change()
    If Combo2 > " " Then
        Text2 = Val(wrapspal) - Val(Text3)
        
    Else
        If Text3 > "0" Then emess = "2nd code date is not specified."
    End If
End Sub

Private Sub Text4_Change()
    If Len(Text4) > 5 And Command1.Enabled = True Then Command1.SetFocus
End Sub
