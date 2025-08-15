VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form r12oppost 
   Caption         =   "Process Order Pick Pallets"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10110
   LinkTopic       =   "Form2"
   ScaleHeight     =   9165
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   8640
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   1695
      Left            =   0
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   2990
      _Version        =   327680
   End
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
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Process"
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
      Left            =   6720
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read Data"
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
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Left            =   960
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   240
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   14843
      _Version        =   327680
      Cols            =   20
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "r12oppost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function r12_lot(plot As String, ocode As String) As String
    Dim s As String, myear As Integer, mdays As Integer
    'MsgBox plot & "_" & ocode
    If Len(plot) >= 5 Then
        myear = Val(Left(plot, 2))
        mdays = Val(Mid(plot, 3, 3)) - 1
        s = "1-1-20" & Left(plot, 2)
        s = Format(DateAdd("d", mdays, s), "MMddyy")
        s = Left(s, 4) & Format(myear + 2, "00")
        If Len(plot) > 5 Then
            's = s & " " & Right(plot, 1)
            s = s & RTrim(Mid(plot, 6, 3))                  'jv052515
        Else
            If Len(ocode) > 2 Then                          'jv052515
                s = s & ocode                               'jv052515
            Else                                            'jv052515
                s = s & " " & ocode
            End If                                          'jv052515
        End If
    Else
        s = " "
    End If
    r12_lot = s
End Function

Private Sub process_zo()
    Dim db As ADODB.Connection, ds As ADODB.Recordset, s As String, i As Integer
    Set db = CreateObject("ADODB.Connection")
    If Form1.Combo1 = "500" Then db.Open "ODBC;DATABASE=WDship;UID=bbcship500;PWD=brenham500;DSN=wdship500"
    If Form1.Combo1 = "501" Then db.Open "ODBC;DATABASE=BAship;UID=bbcship501;PWD=Barrow501;DSN=wdship501"
    If Form1.Combo1 = "502" Then db.Open "ODBC;DATABASE=SYship;UID=bbcship502;PWD=Alabama502;DSN=wdship502"
    For i = 1 To Grid1.Rows - 1
        If Left(Grid1.TextMatrix(i, 13), 7) = "STAGING" Then
            s = "select id from trailers where plant = 50 and branch = 1"
            s = s & " and sku = '" & Grid1.TextMatrix(i, 8) & "'"
            's = s & " and (groupcode like 'OP' or groupcode like 'ZO')"
            s = s & " and trlno in ('OP', 'ZO')"
            s = s & " and shipdate >= '" & Text1 & "'"
            s = s & " and units >= " & Val(Grid1.TextMatrix(i, 10))
            MsgBox s
            Set ds = db.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                s = "Update trailers set units = units - " & Val(Grid1.TextMatrix(i, 10))
                s = s & " where id = " & ds!id
                MsgBox s
            End If
            ds.Close
        End If
    Next i
    
    db.Close
End Sub

Private Sub refresh_grid1_new()
    Dim cfile As String, ofile As String, s As String
    Dim f1 As String, f2 As String, f3 As String, f4 As String, f5 As String
    Dim f6 As String, f7 As String, f8 As String, f9 As String, f10 As String
    Dim f11 As String, f12 As String, f13 As String, f14 As String, f15 As String
    Dim f16 As String, f17 As String, lot2 As String, rbfile As String
    Dim pbranch As String, ctest As String, mfile As String, sdate As String
    Dim torg As String, twhs As String, tacct As String, psku As String, plot As String
    sdate = Text1
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 20
    If Form1.Combo1 = "500" Then
        mplant = "50"
        morg = "500"
        mwhs = "T10"
        mfile = "\\bbc-01-prodtrk\wd\pallogs\move" & Format(Text1, "MMddyyyy") & ".txt"
    End If
    If Form1.Combo1 = "501" Then
        mplant = "51"
        morg = "501"
        mwhs = "K10"
        mfile = "\\bbba-03-dc\f\user\waredist\data\pallogs\move" & Format(Text1, "MMddyyyy") & ".txt"
    End If
    If Form1.Combo1 = "502" Then
        mplant = "52"
        morg = "502"
        mwhs = "A10"
        mfile = "\\bbsy-02-dc\f\user\waredist\data\pallogs\move" & Format(Text1, "MMddyyyy") & ".txt"
    End If
    'If Len(Dir(ofile)) > 0 Then
    '    MsgBox "This data has already posted.", vbOKOnly + vbExclamation, sdate & " " & ofile
    '    Exit Sub
    'End If
    If Len(Dir(mfile)) = 0 Then
        MsgBox "Pallet move file does not exist for this date.", vbOKOnly + vbExclamation, mfile
        Exit Sub
    End If
    'Rack Moves to Order Pick
    
    
    'sdate = Left(sdate, 2) & "-" & mid(sdate, 3, 2) & "-" & Right(sdate, 4)
    Open mfile For Input As #2
    Do Until EOF(2)
        Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
        If f5 = "ORDER PICK" And f7 > "100" And f14 = "COMP" And Trim(f4) <> "M-OP" Then     'jv010313
            s = mplant & DateDiff("d", "1-1-2012", sdate) & "W"
            If mplant = "50" Then s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
            If mplant = "51" Then s = s & Chr(9) & "501" & Chr(9) & "K10" & Chr(9) & "FLOORK10" & Chr(9) & "047" & Chr(9) & "047" & Chr(9) & "FLOOR047"
            If mplant = "52" Then s = s & Chr(9) & "502" & Chr(9) & "A10" & Chr(9) & "FLOORA10" & Chr(9) & "052" & Chr(9) & "052" & Chr(9) & "FLOOR052"
            s = s & Chr(9) & "......"
            s = s & Chr(9) & Trim(Left(f7, 4))
            's = s & Chr(9) & Mid(f7, 5, 8)
            s = s & Chr(9) & RTrim(Mid(f7, 5, 9))                   'jv052515
            s = s & Chr(9) & f11
            s = s & Chr(9) & "EACH"
            s = s & Chr(9) & sdate
            s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
            s = s & Chr(9) & sdate
            s = s & Chr(9) & "Y"
            Grid1.AddItem s
            If Val(f12) > 0 Then    '2nd lot
                'lot2 = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
                ''lot2 = Format(DateAdd("d", Val(f12) - Val(f10), lot2), "MMddyy") & Mid(f7, 11, 2)
                'lot2 = Format(DateAdd("d", Val(Left(f12, 5)) - Val(f10), lot2), "MMddyy") & Mid(f7, 11, 2)  'jv082715
                ''lot2 = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
                ''lot2 = r12_lot(f12, Trim(Mid(f7, 11, 3)))           'jv052515
                lot2 = r12_lot(f12, Trim(Mid(f12, 6, 3)))           'jv091815
                s = mplant & DateDiff("d", "1-1-2012", sdate) & "W"
                If mplant = "50" Then s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
                If mplant = "51" Then s = s & Chr(9) & "501" & Chr(9) & "K10" & Chr(9) & "FLOORK10" & Chr(9) & "047" & Chr(9) & "047" & Chr(9) & "FLOOR047"
                If mplant = "52" Then s = s & Chr(9) & "502" & Chr(9) & "A10" & Chr(9) & "FLOORA10" & Chr(9) & "052" & Chr(9) & "052" & Chr(9) & "FLOOR052"
                s = s & Chr(9) & "......"
                s = s & Chr(9) & Trim(Left(f7, 4))
                s = s & Chr(9) & lot2
                s = s & Chr(9) & f13
                s = s & Chr(9) & "EACH"
                s = s & Chr(9) & sdate
                s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
                s = s & Chr(9) & sdate
                s = s & Chr(9) & "Y"
                Grid1.AddItem s
            End If
        End If
    Loop
    Close #2
    
    'If mplant = "50" Then                                           'Hold Products
    '    mfile = "v:\testlogs\move" & Format(Text1, "MMddyyyy") & ".txt"
    '    Open mfile For Input As #2
    '    Do Until EOF(2)
    '        Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
    '        If f2 = "HOLD" Then
    '            s = mplant & DateDiff("d", "1-1-2012", sdate) & "W"
    '            If f4 = "HOLD" Then         'source
    '                s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "HOLDT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
    '            Else
    '                s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
    '            End If
    '            'If f4 = "HOLD" Then         'source
    '            '    s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "HOLDT10" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10"
    '            'Else
    '            '    s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "HOLDT10"
    '            'End If
    '
    '            s = s & Chr(9) & "......"
    '            s = s & Chr(9) & Trim(Left(f7, 4))
    '            s = s & Chr(9) & Mid(f7, 5, 8)
    '            s = s & Chr(9) & f11
    '            s = s & Chr(9) & "EACH"
    '            s = s & Chr(9) & sdate
    '            s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
    '            s = s & Chr(9) & sdate
    '            s = s & Chr(9) & "Y"
    '            Grid1.AddItem s
    '            If Val(f12) > 0 Then    '2nd lot
    '                lot2 = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
    '                lot2 = Format(DateAdd("d", Val(f12) - Val(f10), lot2), "MMddyy") & Mid(f7, 11, 2)
    '                lot2 = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
    '                s = mplant & DateDiff("d", "1-1-2012", sdate) & "W"
    '                If f4 = "HOLD" Then         'source
    '                    s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "HOLDT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
    '                Else
    '                    s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
    '                End If
    '                'If f4 = "HOLD" Then         'source
    '                '    s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "HOLDT10" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10"
    '                'Else
    '                '    s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "HOLDT10"
    '                'End If
    '                's = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
    '                s = s & Chr(9) & "......"
    '                s = s & Chr(9) & Trim(Left(f7, 4))
    '                s = s & Chr(9) & lot2
    '                s = s & Chr(9) & f13
    '                s = s & Chr(9) & "EACH"
    '                s = s & Chr(9) & sdate
    '                s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
    '                s = s & Chr(9) & sdate
    '                s = s & Chr(9) & "Y"
    '                Grid1.AddItem s
    '            End If
                    
    '            'Use if routing through order pick
    '            s = mplant & DateDiff("d", "1-1-2012", sdate) & "W"
    '            If f4 = "HOLD" Then         'source
    '                s = s & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10"
    '            Else
    '                s = s & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "HOLDT10"
    '            End If
    '            s = s & Chr(9) & "......"
    '            s = s & Chr(9) & Trim(Left(f7, 4))
    '            s = s & Chr(9) & Mid(f7, 5, 8)
    '            s = s & Chr(9) & f11
    '            s = s & Chr(9) & "EACH"
    '            s = s & Chr(9) & sdate
    '            s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
    '            s = s & Chr(9) & sdate
    '            s = s & Chr(9) & "Y"
    '            Grid1.AddItem s
    '            If Val(f12) > 0 Then    '2nd lot
    '                lot2 = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
    '                lot2 = Format(DateAdd("d", Val(f12) - Val(f10), lot2), "MMddyy") & Mid(f7, 11, 2)
    '                lot2 = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
    '                s = mplant & DateDiff("d", "1-1-2012", sdate) & "W"
    '                If f4 = "HOLD" Then         'source
    '                    s = s & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10"
    '                Else
    '                    s = s & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "HOLDT10"
    '                End If
    '                's = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
    '                s = s & Chr(9) & "......"
    '                s = s & Chr(9) & Trim(Left(f7, 4))
    '                s = s & Chr(9) & lot2
    '                s = s & Chr(9) & f13
    '                s = s & Chr(9) & "EACH"
    '                s = s & Chr(9) & sdate
    '                s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
    '                s = s & Chr(9) & sdate
    '                s = s & Chr(9) & "Y"
    '                Grid1.AddItem s
    '            End If
                    
                    
                    
                
    '        End If
    '    Loop
    '    Close #2
    'End If
    
    
    'roller bed
    If mplant = "50" Then
        rbfile = "\\bbc-01-prodtrk\wd\pallogs\recv" & Format(sdate, "MMddyyyy") & ".txt"
        If Len(Dir(rbfile)) > 0 Then
        Open rbfile For Input As #2
        Do Until EOF(2)
            Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
            If f4 = "ROLLER BED" And f5 = "ORDER PICK" And f7 > "100" Then            'jv010313
                s = mplant & DateDiff("d", "1-1-2012", sdate) & "W"
                s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
                s = s & Chr(9) & "......"
                s = s & Chr(9) & Trim(Left(f7, 4))
                's = s & Chr(9) & Mid(f7, 5, 8)
                s = s & Chr(9) & RTrim(Mid(f7, 5, 9))                   'jv052515
                s = s & Chr(9) & f11
                s = s & Chr(9) & "EACH"
                s = s & Chr(9) & sdate
                s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
                s = s & Chr(9) & sdate
                s = s & Chr(9) & "Y"
                Grid1.AddItem s
                If Val(f12) > 0 Then    '2nd lot
                    lot2 = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
                    lot2 = Format(DateAdd("d", Val(Left(f12, 5)) - Val(f10), lot2), "MMddyy") & Mid(f7, 11, 2)  'jv082715
                    'lot2 = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
                    'lot2 = r12_lot(f12, Trim(Mid(f7, 11, 3)))           'jv052515
                    lot2 = r12_lot(f12, Trim(Mid(f12, 6, 3)))           'jv091815
                    s = mplant & DateDiff("d", "1-1-2012", sdate) & "W"
                    s = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
                    s = s & Chr(9) & "......"
                    s = s & Chr(9) & Trim(Left(f7, 4))
                    s = s & Chr(9) & lot2
                    s = s & Chr(9) & f13
                    s = s & Chr(9) & "EACH"
                    s = s & Chr(9) & sdate
                    s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
                    s = s & Chr(9) & sdate
                    s = s & Chr(9) & "Y"
                    Grid1.AddItem s
                End If
            End If
        Loop
        Close #2
        End If
    End If
    'Exit Sub
    
    'return to rack                     added 12-01-16
    If mplant = "50" Then
        rbfile = "\\bbc-01-prodtrk\wd\pallogs\move" & Format(sdate, "MMddyyyy") & ".txt"
        If Len(Dir(rbfile)) > 0 Then
        Open rbfile For Input As #2
        Do Until EOF(2)
            Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
            If f4 = "M-OP" And f5 <> "ORDER PICK" And f7 > "100" Then            'jv010313
                s = mplant & DateDiff("d", "1-1-2012", sdate) & "W"
                's = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
                s = s & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10"
                s = s & Chr(9) & "......"
                s = s & Chr(9) & Trim(Left(f7, 4))
                's = s & Chr(9) & Mid(f7, 5, 8)
                s = s & Chr(9) & RTrim(Mid(f7, 5, 9))                   'jv052515
                s = s & Chr(9) & f11
                s = s & Chr(9) & "EACH"
                s = s & Chr(9) & sdate
                s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
                s = s & Chr(9) & sdate
                s = s & Chr(9) & "Y"
                Grid1.AddItem s
                If Val(f12) > 0 Then    '2nd lot
                    lot2 = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
                    lot2 = Format(DateAdd("d", Val(Left(f12, 5)) - Val(f10), lot2), "MMddyy") & Mid(f7, 11, 2)  'jv082715
                    'lot2 = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
                    'lot2 = r12_lot(f12, Trim(Mid(f7, 11, 3)))           'jv052515
                    lot2 = r12_lot(f12, Trim(Mid(f12, 6, 3)))           'jv091815
                    s = mplant & DateDiff("d", "1-1-2012", sdate) & "W"
                    's = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
                    s = s & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10"
                    s = s & Chr(9) & "......"
                    s = s & Chr(9) & Trim(Left(f7, 4))
                    s = s & Chr(9) & lot2
                    s = s & Chr(9) & f13
                    s = s & Chr(9) & "EACH"
                    s = s & Chr(9) & sdate
                    s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
                    s = s & Chr(9) & sdate
                    s = s & Chr(9) & "Y"
                    Grid1.AddItem s
                End If
            End If
        Loop
        Close #2
        End If
    End If
    
    
    'snack plant
    If mplant = 50 Then
        s = ""
        Open mfile For Input As #2
        Do Until EOF(2)
            Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
            If f2 = "DOCK" And (f4 = "1405" Or f4 = "1406" Or f4 = "1731") Then
                s = mplant & DateDiff("d", "1-1-2012", sdate) & "P"
                s = s & Chr(9) & "503" & Chr(9) & "S10" & Chr(9) & "FLOORS10" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10"
                s = s & Chr(9) & "......"
                s = s & Chr(9) & Trim(Left(f7, 4))
                's = s & Chr(9) & Mid(f7, 5, 8)
                s = s & Chr(9) & RTrim(Mid(f7, 5, 9))                   'jv052515
                s = s & Chr(9) & f11
                s = s & Chr(9) & "EACH"
                s = s & Chr(9) & sdate
                s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
                s = s & Chr(9) & sdate
                s = s & Chr(9) & "N"
                Grid1.AddItem s
                If Val(f12) > 0 Then    '2nd lot
                    lot2 = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
                    'lot2 = Format(DateAdd("d", Val(f12) - Val(f10), lot2), "MMddyy") & Mid(f7, 11, 2)
                    lot2 = Format(DateAdd("d", Val(Left(f12, 5)) - Val(f10), lot2), "MMddyy") & Mid(f7, 11, 2)  'jv082715
                    'lot2 = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
                    'lot2 = r12_lot(f12, Trim(Mid(f7, 11, 3)))           'jv052515
                    lot2 = r12_lot(f12, Trim(Mid(f12, 6, 3)))           'jv091815
                    s = mplant & DateDiff("d", "1-1-2012", sdate) & "P"
                    s = s & Chr(9) & "503" & Chr(9) & "S10" & Chr(9) & "FLOORS10" & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10"
                    s = s & Chr(9) & "......"
                    s = s & Chr(9) & Trim(Left(f7, 4))
                    s = s & Chr(9) & lot2
                    s = s & Chr(9) & f13
                    s = s & Chr(9) & "EACH"
                    s = s & Chr(9) & sdate
                    s = s & Chr(9) & Trim(f4) & " " & Right(f7, 3)
                    s = s & Chr(9) & sdate
                    s = s & Chr(9) & "N"
                    Grid1.AddItem s
                End If
            End If
        Loop
        Close #2
    End If
    'end snack plant
    s = "^Ticket|^FromOrg|^FromSub|^FromLoc|^ToOrg|<ToSub|^To_Loc|^Account|^SKU|^LotNum|^Units|^UOM|^ShipDate|<Comment|^EarlyDate|^PFlag"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 800
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 1000
    Grid1.ColWidth(8) = 800
    Grid1.ColWidth(9) = 1000
    Grid1.ColWidth(10) = 800
    Grid1.ColWidth(11) = 800
    Grid1.ColWidth(12) = 1000
    Grid1.ColWidth(13) = 2000
    Grid1.ColWidth(14) = 1000
    Grid1.ColWidth(15) = 600
    
End Sub

Private Sub postoprb_r12(mplant As String, sdate As String)
    Dim cfile As String, ofile As String, s As String
    Dim f1 As String, f2 As String, f3 As String, f4 As String, f5 As String
    Dim f6 As String, f7 As String, f8 As String, f9 As String, f10 As String
    Dim f11 As String, f12 As String, f13 As String, f14 As String, f15 As String
    Dim f16 As String, f17 As String
    Dim pbranch As String, ctest As String, mfile As String, rbfile As String
    Dim morg As String, mwhs As String, psku As String, plot As String
    Dim oorg As String, owhs As String, afile As String, webdir As String
    Dim adjlit As String
    adjlit = " "
    webdir = "s:\wd\html"
    'webdir = "u:\test"
    If mplant = "50" Then
        morg = "500"
        mwhs = "T10"
        oorg = "001"    'jv011513
        owhs = "001"    'jv011513
        mfile = "\\bbc-01-prodtrk\wd\pallogs\move" & Format(sdate, "MMddyyyy") & ".txt"
        ofile = "\\bbc-01-prodtrk\wd\pallogs\RO" & mplant & DateDiff("d", "1-1-2012", sdate) & ".txt"
    End If
    If mplant = "51" Then
        morg = "501"
        mwhs = "K10"
        oorg = "047"    'jv011513
        owhs = "047"    'jv011513
        mfile = "\\bbba-03-dc\f\user\waredist\data\pallogs\move" & Format(sdate, "MMddyyyy") & ".txt"
        ofile = "\\bbba-03-dc\f\user\waredist\data\pallogs\RO" & mplant & DateDiff("d", "1-1-2012", sdate) & ".txt"
    End If
    If mplant = "52" Then
        morg = "502"
        mwhs = "A10"
        oorg = "052"    'jv011513
        owhs = "052"    'jv011513
        mfile = "\\bbsy-02-dc\f\user\waredist\data\pallogs\move" & Format(sdate, "MMddyyyy") & ".txt"
        ofile = "\\bbsy-02-dc\f\user\waredist\data\pallogs\RO" & mplant & DateDiff("d", "1-1-2012", sdate) & ".txt"
    End If
    
    'ofile = "u:\jvtest.txt"
    If Len(Dir(ofile)) > 0 Then
        If FileLen(ofile) > 0 Then                              'jv010313
            MsgBox "This data has already posted.", vbOKOnly + vbExclamation, sdate & " " & ofile
            Exit Sub
        End If
    End If
    'Rack Moves to Order Pick
    
    Open ofile For Output As #1
    
    'sdate = Left(sdate, 2) & "-" & mid(sdate, 3, 2) & "-" & Right(sdate, 4)
    Open mfile For Input As #2
    Do Until EOF(2)
        Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
        If f5 = "ORDER PICK" And f7 > "100" And f14 = "COMP" And Trim(f4) <> "M-OP" Then     'jv010313
            Write #1, mplant & DateDiff("d", "1-1-2012", sdate) & "W";
            If mplant = "50" Then Write #1, "500"; "T10"; "FLOORT10"; "001"; "001"; "FLOOR001";
            If mplant = "51" Then Write #1, "501"; "K10"; "FLOORK10"; "047"; "047"; "FLOOR047";
            If mplant = "52" Then Write #1, "502"; "A10"; "FLOORA10"; "052"; "052"; "FLOOR052";
            Write #1, "......";
            Write #1, Trim(Left(f7, 4));
            'Write #1, Mid(f7, 5, 8);
            Write #1, RTrim(Mid(f7, 5, 9));                     'jv052515
            Write #1, f11;
            Write #1, "EACH";
            Write #1, sdate;
            Write #1, Trim(f4) & " " & Right(f7, 3);
            Write #1, sdate;
            Write #1, "Y"
            If Val(f12) > 0 Then    '2nd lot
                s = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
                's = Format(DateAdd("d", Val(f12) - Val(f10), s), "MMddyy") & Mid(f7, 11, 2)
                s = Format(DateAdd("d", Val(Left(f12, 5)) - Val(f10), s), "MMddyy") & Mid(f7, 11, 2)    'jv081715
                'Write #1, Format(Now, "Mddyy");
                Write #1, mplant & DateDiff("d", "1-1-2012", sdate) & "W";
                If mplant = "50" Then Write #1, "500"; "T10"; "FLOORT10"; "001"; "001"; "FLOOR001";
                If mplant = "51" Then Write #1, "501"; "K10"; "FLOORK10"; "047"; "047"; "FLOOR047";
                If mplant = "52" Then Write #1, "502"; "A10"; "FLOORA10"; "052"; "052"; "FLOOR052";
                Write #1, "......";
                Write #1, Trim(Left(f7, 4));
                's = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
                's = r12_lot(f12, Trim(Mid(f7, 11, 3)))              'jv052515
                s = r12_lot(f12, Trim(Mid(f12, 6, 3)))              'jv091815
                Write #1, s;
                Write #1, f13;
                Write #1, "EACH";
                Write #1, sdate;
                Write #1, Trim(f4) & " " & Right(f7, 3);
                Write #1, sdate;
                Write #1, "Y"
            End If
        End If
        'Process Order Pick adjustements - return to racks
        If Trim(f4) = "M-OP" And f5 <> "ORDER PICK" And f14 = "COMP" And f7 > "100" Then
            adjlit = "Adjustments have been posted for Order Pick (" & owhs & ") and " & mwhs & "."
            afile = webdir & "\counts\whsadj." & owhs
            'afile = "U:\whsadj." & owhs
            Open afile For Append As #3
            Write #3, Format(sdate, "m-d-yyyy");
            Write #3, oorg;
            Write #3, owhs;
            Write #3, "LOT1";
            Write #3, Trim(Left(f7, 4));
            Write #3, Right(f6, Len(f6) - 4);
            Write #3, Format((Val(f11) + Val(f13)) * -1, "0");
            Write #3, "TRAN";
            Write #3, "WMS";
            Write #3, Format(sdate, "m-d-yyyy")
            Close #3
            afile = webdir & "\counts\whsadj." & mwhs
            'afile = "u:\whsadj." & mwhs
            Open afile For Append As #4
            Write #4, Format(sdate, "m-d-yyyy");
            Write #4, morg;
            Write #4, mwhs;
            'Write #4, Mid(f7, 5, 8);
            Write #4, RTrim(Mid(f7, 5, 9));                     'jv052515
            Write #4, Trim(Left(f7, 4));
            Write #4, Right(f6, Len(f6) - 4);
            Write #4, Format(Val(f11), "0");
            Write #4, "TRAN";
            Write #4, "WMS";
            Write #4, Format(sdate, "m-d-yyyy")
            If Val(f12) > 0 Then    '2nd lot
                s = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
                's = Format(DateAdd("d", Val(f12) - Val(f10), s), "MMddyy") & Mid(f7, 11, 2)
                s = Format(DateAdd("d", Val(Left(f12, 5)) - Val(f10), s), "MMddyy") & Mid(f7, 11, 2)    'jv082715
                Write #4, Format(sdate, "m-d-yyyy");
                Write #4, morg;
                Write #4, mwhs;
                's = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
                's = r12_lot(f12, Trim(Mid(f7, 11, 3)))          'jv052515
                s = r12_lot(f12, Trim(Mid(f12, 6, 3)))          'jv091815
                Write #4, s;
                Write #4, Trim(Left(f7, 4));
                Write #4, Right(f6, Len(f6) - 4);
                Write #4, Format(Val(f13), "0");
                Write #4, "TRAN";
                Write #4, "WMS";
                Write #4, Format(sdate, "m-d-yyyy")
            End If
            Close #4
        End If
    Loop
    Close #2
    
    'Hold products
    'If mplant = "50" Then                                           'Hold Products
    '    mfile = "v:\testlogs\move" & Format(Text1, "MMddyyyy") & ".txt"
    '    Open mfile For Input As #2
    '    Do Until EOF(2)
    '        Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
    '        If f2 = "HOLD" Then
    '            Write #1, mplant & DateDiff("d", "1-1-2012", sdate) & "W";
    '            If f4 = "HOLD" Then         'source
    '                Write #1, "500"; "T10"; "HOLDT10"; "001"; "001"; "FLOOR001";
    '            Else
    '                Write #1, "500"; "T10"; "FLOORT10"; "001"; "001"; "FLOOR001";
    '            End If
    '            'If f4 = "HOLD" Then         'source
    '            '    Write #1, "500"; "T10"; "HOLDT10"; "500"; "T10"; "FLOORT10";
    '            'Else
    '            '    Write #1, "500"; "T10"; "FLOORT10"; "500"; "T10"; "HOLDT10";
    '            'End If
    '            Write #1, "......";
    '            Write #1, Trim(Left(f7, 4));
    '            Write #1, Mid(f7, 5, 8);
    '            Write #1, f11;
    '            Write #1, "EACH";
    '            Write #1, sdate;
    '            Write #1, Trim(f4) & " " & Right(f7, 3);
    '            Write #1, sdate;
    '            Write #1, "Y"
    '            If Val(f12) > 0 Then    '2nd lot
    '                lot2 = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
    '                lot2 = Format(DateAdd("d", Val(f12) - Val(f10), lot2), "MMddyy") & Mid(f7, 11, 2)
    '                lot2 = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
    '                Write #1, mplant & DateDiff("d", "1-1-2012", sdate) & "W";
    '                If f4 = "HOLD" Then         'source
    '                    Write #1, "500"; "T10"; "HOLDT10"; "001"; "001"; "FLOOR001";
    '                Else
    '                    Write #1, "500"; "T10"; "FLOORT10"; "001"; "001"; "FLOOR001";
    '                End If
    '                'If f4 = "HOLD" Then         'source
    '                '    Write #1, "500"; "T10"; "HOLDT10"; "500"; "T10"; "FLOORT10";
    '                'Else
    '                '    Write #1, "500"; "T10"; "FLOORT10"; "500"; "T10"; "HOLDT10";
    '                'End If
    '                's = s & Chr(9) & "500" & Chr(9) & "T10" & Chr(9) & "FLOORT10" & Chr(9) & "001" & Chr(9) & "001" & Chr(9) & "FLOOR001"
    '                Write #1, "......";
    '                Write #1, Trim(Left(f7, 4));
    '                Write #1, lot2;
    '                Write #1, f13;
    '                Write #1, "EACH";
    '                Write #1, sdate;
    '                Write #1, Trim(f4) & " " & Right(f7, 3);
    '                Write #1, sdate;
    '                Write #1, "Y"
    '            End If
    '
    '            'Use if routing through order pick
    '            Write #1, mplant & DateDiff("d", "1-1-2012", sdate) & "W";
    '            If f4 = "HOLD" Then         'source
    '                Write #1, "001"; "001"; "FLOOR001"; "500"; "T10"; "FLOORT10";
    '            Else
    '                Write #1, "001"; "001"; "FLOOR001"; "500"; "T10"; "HOLDT10";
    '            End If
    '            Write #1, "......";
    '            Write #1, Trim(Left(f7, 4));
    '            Write #1, Mid(f7, 5, 8);
    '            Write #1, f11;
    '            Write #1, "EACH";
    '            Write #1, sdate;
    '            Write #1, Trim(f4) & " " & Right(f7, 3);
    '            Write #1, sdate;
    '            Write #1, "Y"
    '            If Val(f12) > 0 Then    '2nd lot
    '                lot2 = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
    '                lot2 = Format(DateAdd("d", Val(f12) - Val(f10), lot2), "MMddyy") & Mid(f7, 11, 2)
    '                lot2 = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
    '                Write #1, mplant & DateDiff("d", "1-1-2012", sdate) & "W";
    '                If f4 = "HOLD" Then         'source
    '                    Write #1, "001"; "001"; "FLOOR001"; "500"; "T10"; "FLOORT10";
    '                Else
    '                    Write #1, "001"; "001"; "FLOOR001"; "500"; "T10"; "HOLDT10";
    '                End If
    '                Write #1, "......";
    '                Write #1, Trim(Left(f7, 4));
    '                Write #1, lot2;
    '                Write #1, f13;
    '                Write #1, "EACH";
    '                Write #1, sdate;
    '                Write #1, Trim(f4) & " " & Right(f7, 3);
    '                Write #1, sdate;
    '                Write #1, "Y"
    '            End If
    '        End If
    '    Loop
    '    Close #2
    'End If
            
    'Roller Bed ---------------- jv010313
    If mplant = "50" Then
        rbfile = "\\bbc-01-prodtrk\wd\pallogs\recv" & Format(sdate, "MMddyyyy") & ".txt"
        If Len(Dir(rbfile)) > 0 Then
            Open rbfile For Input As #2
            Do Until EOF(2)
                Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
                If f4 = "ROLLER BED" And f5 = "ORDER PICK" And f7 > "100" Then            'Jv010313
                    Write #1, mplant & DateDiff("d", "1-1-2012", sdate) & "W";
                    Write #1, "500"; "T10"; "FLOORT10"; "001"; "001"; "FLOOR001";
                    Write #1, "......";
                    Write #1, Trim(Left(f7, 4));
                    'Write #1, Mid(f7, 5, 8);
                    Write #1, RTrim(Mid(f7, 5, 9));                 'jv052515
                    Write #1, f11;
                    Write #1, "EACH";
                    Write #1, sdate;
                    Write #1, Trim(f4) & " " & Right(f7, 3);
                    Write #1, sdate;
                    Write #1, "Y"
                    If Val(f12) > 0 Then    '2nd lot
                        s = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
                        's = Format(DateAdd("d", Val(f12) - Val(f10), s), "MMddyy") & Mid(f7, 11, 2)
                        s = Format(DateAdd("d", Val(Left(f12, 5)) - Val(f10), s), "MMddyy") & Mid(f7, 11, 2)    'jv082715
                        'Write #1, Format(Now, "Mddyy");
                        Write #1, mplant & DateDiff("d", "1-1-2012", sdate) & "W";
                        Write #1, "500"; "T10"; "FLOORT10"; "001"; "001"; "FLOOR001";
                        Write #1, "......";
                        Write #1, Trim(Left(f7, 4));
                        's = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
                        's = r12_lot(f12, Trim(Mid(f7, 11, 3)))          'jv052515
                        s = r12_lot(f12, Trim(Mid(f12, 6, 3)))          'jv091815
                        Write #1, s;
                        Write #1, f13;
                        Write #1, "EACH";
                        Write #1, sdate;
                        Write #1, Trim(f4) & " " & Right(f7, 3);
                        Write #1, sdate;
                        Write #1, "Y"
                    End If
                End If
            Loop
            Close #2
        End If
    End If
    
    'snack plant
    If mplant = 50 Then
        'mfile = "\\bbc-01-prodtrk\wd\pallogs\recv" & Format(sdate, "MMddyyyy") & ".txt"
        s = ""
        Open mfile For Input As #2
        Do Until EOF(2)
            Input #2, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16, f17
            'If f2 = "SNACK PLANT WRAPPER" And f5 <> "SNACK PLANT" And f5 <> "WRAPPER" Then
            If f2 = "DOCK" And (f4 = "1405" Or f4 = "1406" Or f4 = "1731") Then
                Write #1, mplant & DateDiff("d", "1-1-2012", sdate) & "P";
                s = mplant & DateDiff("d", "1-1-2012", sdate) & "P"
                Write #1, "503"; "S10"; "FLOORS10"; "500"; "T10"; "FLOORT10";
                Write #1, "......";
                Write #1, Trim(Left(f7, 4));
                'Write #1, Mid(f7, 5, 8);
                Write #1, RTrim(Mid(f7, 5, 9));                     'jv052515
                Write #1, f11;
                Write #1, "EACH";
                Write #1, sdate;
                Write #1, Trim(f4) & " " & Right(f7, 3);
                Write #1, sdate;
                Write #1, "N"
                If Val(f12) > 0 Then    '2nd lot
                    s = Mid(f7, 5, 2) & "-" & Mid(f7, 7, 2) & "-20" & Mid(f7, 9, 2)
                    's = Format(DateAdd("d", Val(f12) - Val(f10), s), "MMddyy") & Mid(f7, 11, 2)
                    s = Format(DateAdd("d", Val(Left(f12, 5)) - Val(f10), s), "MMddyy") & Mid(f7, 11, 2)    'jv082715
                    'Write #1, Format(Now, "Mddyy");
                    Write #1, mplant & DateDiff("d", "1-1-2012", sdate) & "P";
                    Write #1, "503"; "S10"; "FLOORS10"; "500"; "T10"; "FLOORT10";
                    Write #1, "......";
                    Write #1, Trim(Left(f7, 4));
                    's = r12_lot(f12, Mid(f7, 12, 1))                 'jv020614
                    's = r12_lot(f12, Trim(Mid(f7, 11, 3)))          'jv052515
                    s = r12_lot(f12, Trim(Mid(f12, 6, 3)))          'jv091815
                    Write #1, s;
                    Write #1, f13;
                    Write #1, "EACH";
                    Write #1, sdate;
                    Write #1, Trim(f4) & " " & Right(f7, 3);
                    Write #1, sdate;
                    Write #1, "N"
                End If
            End If
        Loop
        Close #2
        If s > " " Then
            s = "Org Transfer Ticket " & s & " has been created for Snack Plant items."
            MsgBox s, vbOKOnly + vbInformation, "Receivng ticket..."
        End If
    End If
    'end snack plant
    If adjlit > " " Then            'jv011513
        MsgBox adjlit, vbOKOnly + vbInformation, "Adjustments posted...."
    End If
    
    Close #1
    'addfile = False
    'Exit Sub       'Turn on when testing.
    ofile = Form1.pallogs & "r12trls.win"
    Open ofile For Output As #1
    Print #1, "open pbelle.bluebell.com"
    Print #1, "infbbcri"
    Print #1, "welcome@2023"
    Print #1, "BINARY"
    Print #1, "cd PBELLE/incoming"
    Print #1, "lcd "; Left(Form1.pallogs, Len(Form1.pallogs) - 1)
    Print #1, "put RO" & mplant & DateDiff("d", "1-1-2012", sdate) & ".txt RO" & mplant & DateDiff("d", "1-1-2012", sdate) & ".txt"
    Print #1, "close"
    Print #1, "bye"
    Close #1
    'If addfile = True Then
        ftpexe = "c:\windows\system32\ftp.exe"
        x = Shell(ftpexe & " -s:" & ofile, vbNormalFocus)
        MsgBox ftpexe & " -s:" & ofile
    'End If
End Sub

Private Sub Command1_Click()
    refresh_grid1_new
End Sub

Private Sub Command2_Click()
    Dim sdate As String
    sdate = Text1
    If Len(sdate) = 0 Then Exit Sub
    If IsDate(sdate) = False Then
        MsgBox "invalid date: " & sdate, vbOKOnly + vbExclamation, "sorry, try again..."
        Exit Sub
    End If
    If Form1.Combo1 = "500" Then Call postoprb_r12("50", sdate)
    If Form1.Combo1 = "501" Then Call postoprb_r12("51", sdate)
    If Form1.Combo1 = "502" Then Call postoprb_r12("52", sdate)
End Sub

Private Sub Command3_Click()
    Dim rt As String, rh As String, rf As String
    Dim i As Integer, s As String
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 7: pgrid.FixedCols = 0
    If Grid1.Rows < 2 Then Exit Sub
    For i = 1 To Grid1.Rows - 1
        s = Grid1.TextMatrix(i, 2) & " " & Grid1.TextMatrix(i, 3)
        s = s & Chr(9) & Grid1.TextMatrix(i, 5) & " " & Grid1.TextMatrix(i, 6)
        s = s & Chr(9) & Grid1.TextMatrix(i, 8)
        s = s & Chr(9) & Grid1.TextMatrix(i, 9)
        s = s & Chr(9) & Grid1.TextMatrix(i, 10)
        s = s & Chr(9) & Grid1.TextMatrix(i, 11)
        s = s & Chr(9) & Grid1.TextMatrix(i, 13)
        pgrid.AddItem s
    Next i
    s = "^FromSub|^ToSub|^SKU|^LotNum|^Units|^UOM|<Comment"
    pgrid.FormatString = s
    pgrid.ColWidth(0) = 2000 'Grid1.ColWidth(2)
    pgrid.ColWidth(1) = 2000 'Grid1.ColWidth(5)
    pgrid.ColWidth(2) = Grid1.ColWidth(8)
    pgrid.ColWidth(3) = Grid1.ColWidth(9)
    pgrid.ColWidth(4) = Grid1.ColWidth(10)
    pgrid.ColWidth(5) = Grid1.ColWidth(11)
    pgrid.ColWidth(6) = Grid1.ColWidth(13)
    If Grid1.TextMatrix(1, 0) <> Grid1.TextMatrix(Grid1.Rows - 1, 0) Then
        rt = "Tickets: " & Grid1.TextMatrix(1, 0) & " & " & Grid1.TextMatrix(Grid1.Rows - 1, 0)
    Else
        rt = "Ticket: " & Grid1.TextMatrix(1, 0)
    End If
    pgrid.RowSel = pgrid.Row
    pgrid.Col = 2: pgrid.ColSel = 6
    pgrid.Sort = 5
    rh = "Ship Date: " & Text1
    rf = "Printed: " & Format(Now, "m-dd-yyyy h:mm am/pm")
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

Private Sub Command4_Click()
    process_zo
End Sub

Private Sub Form_Load()
    Text1 = Format(DateAdd("d", -1, Now), "MM-dd-yyyy")
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    pgrid.Width = Me.Width - 80
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1080
End Sub
