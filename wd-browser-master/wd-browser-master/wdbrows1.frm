VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "W&D Browser"
   ClientHeight    =   7155
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   12915
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4455
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   6615
      ExtentX         =   11668
      ExtentY         =   7858
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox locdir 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "f:\public"
      Top             =   5760
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ListBox bclist 
      Height          =   5910
      Left            =   6840
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid frmgrid 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   1296
      _Version        =   327680
   End
   Begin VB.TextBox webdir 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "\\BBC-03-FILESVR\SharedGroups\WD\html"
      Top             =   5280
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label bversion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "V2023.10.11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label wdbanner 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "W/D Browser"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   855
      Left            =   0
      TabIndex        =   6
      ToolTipText     =   "Click here to restore homepage.."
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label wdbranch 
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label wduser 
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Menu filemenu 
      Caption         =   "&File"
      Begin VB.Menu homepage 
         Caption         =   "Home Page"
      End
      Begin VB.Menu stocksht 
         Caption         =   "Out of Stock"
      End
      Begin VB.Menu newrel 
         Caption         =   "New Product Releases"
      End
      Begin VB.Menu discprod 
         Caption         =   "Discontinued Products"
      End
      Begin VB.Menu messages 
         Caption         =   "Messages"
      End
      Begin VB.Menu trlsched 
         Caption         =   "Transport Schedule"
      End
      Begin VB.Menu brorders 
         Caption         =   "Branch Orders"
      End
      Begin VB.Menu trsched 
         Caption         =   "Transport Requests"
      End
      Begin VB.Menu vco 
         Caption         =   "Vault Clothes Orders"
         Visible         =   0   'False
      End
      Begin VB.Menu dso 
         Caption         =   "Dry Storage Items"
      End
      Begin VB.Menu brennotes 
         Caption         =   "Notes To Brenham"
      End
      Begin VB.Menu xitmenu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu repmenu 
      Caption         =   "Reports"
      Begin VB.Menu gemminv 
         Caption         =   "Oracle Inventory"
      End
      Begin VB.Menu saleinv 
         Caption         =   "Sales vs. Inventory"
      End
      Begin VB.Menu broos 
         Caption         =   "Stock History"
         Visible         =   0   'False
      End
      Begin VB.Menu tstatrpt 
         Caption         =   "Trailer Status Report"
      End
      Begin VB.Menu tktrpt 
         Caption         =   "Ticket BarCodes"
      End
      Begin VB.Menu tstationrpt 
         Caption         =   "Transfer Station"
         Visible         =   0   'False
      End
      Begin VB.Menu rcomenu 
         Caption         =   "Recommended Orders"
         Visible         =   0   'False
         Begin VB.Menu rco50 
            Caption         =   "Brenham Trailer"
         End
         Begin VB.Menu rco51 
            Caption         =   "Broken Arrow Trailer"
         End
         Begin VB.Menu rco52 
            Caption         =   "Sylacuaga Trailer"
         End
      End
      Begin VB.Menu rcopart 
         Caption         =   "Partial Pallet Order"
         Visible         =   0   'False
      End
      Begin VB.Menu regrepmenu 
         Caption         =   "Regional Reports"
         Visible         =   0   'False
         Begin VB.Menu regreleases 
            Caption         =   "New Releases"
         End
         Begin VB.Menu regshiphist 
            Caption         =   "Shipping History"
            Begin VB.Menu regshiphist1 
               Caption         =   "Region 1"
            End
            Begin VB.Menu regshiphist2 
               Caption         =   "Region 2"
            End
            Begin VB.Menu regshiphist3 
               Caption         =   "Region 3"
            End
            Begin VB.Menu regshiphist4 
               Caption         =   "Region 4"
            End
            Begin VB.Menu regshiphist5 
               Caption         =   "Region 5"
            End
            Begin VB.Menu regshiphist6 
               Caption         =   "Region 6"
            End
            Begin VB.Menu regshiphist7 
               Caption         =   "Region 7"
            End
         End
      End
      Begin VB.Menu brjobtrl 
         Caption         =   "Jobbing Trailers"
         Visible         =   0   'False
      End
      Begin VB.Menu brbobtail 
         Caption         =   "Bobtail Orders"
         Visible         =   0   'False
      End
      Begin VB.Menu syl3gs 
         Caption         =   "Sylacauga 3Gallons"
         Visible         =   0   'False
      End
      Begin VB.Menu blkbill 
         Caption         =   "Blank Bill of Lading"
      End
   End
   Begin VB.Menu csmenu 
      Caption         =   "Count Sheets"
      Begin VB.Menu invadj 
         Caption         =   "Inventory Adjustments"
      End
      Begin VB.Menu gemmeop 
         Caption         =   "Oracle End of Period Totals"
      End
      Begin VB.Menu csc 
         Caption         =   "Cold Storage Countsheet"
      End
      Begin VB.Menu gemmvc 
         Caption         =   "Oracle vs Countsheet"
      End
   End
   Begin VB.Menu pi 
      Caption         =   "Plant Info"
      Visible         =   0   'False
      Begin VB.Menu pi01 
         Caption         =   "Brenham"
         Visible         =   0   'False
         Begin VB.Menu pipi01 
            Caption         =   "Plant Inventory"
         End
         Begin VB.Menu batreleases 
            Caption         =   "Batch Releases"
         End
         Begin VB.Menu pibs01 
            Caption         =   "Branch Ingredient Storage"
            Visible         =   0   'False
         End
         Begin VB.Menu pibi01 
            Caption         =   "Branch Inventories (BIMP)"
         End
         Begin VB.Menu turn01 
            Caption         =   "Branch Pallet Turnover"
         End
         Begin VB.Menu brzloss500 
            Caption         =   "Branch Out of Stock"
            Visible         =   0   'False
         End
         Begin VB.Menu brtrn50 
            Caption         =   "Browse Transports"
         End
         Begin VB.Menu pidt01 
            Caption         =   "Dry Trailer Order"
         End
      End
      Begin VB.Menu pi47 
         Caption         =   "Broken Arrow"
         Visible         =   0   'False
         Begin VB.Menu pipi47 
            Caption         =   "Plant Inventory"
         End
         Begin VB.Menu pibs47 
            Caption         =   "Branch Ingredient Storage"
            Visible         =   0   'False
         End
         Begin VB.Menu turn47 
            Caption         =   "Branch Pallet Turnover"
         End
         Begin VB.Menu brzloss501 
            Caption         =   "Branch Out of Stock"
            Visible         =   0   'False
         End
         Begin VB.Menu pidt47 
            Caption         =   "Dry Trailer Order"
         End
         Begin VB.Menu btrn51 
            Caption         =   "Browse Transports"
         End
         Begin VB.Menu pits47 
            Caption         =   "Trailer Schedule"
         End
      End
      Begin VB.Menu pi52 
         Caption         =   "Sylacauga"
         Visible         =   0   'False
         Begin VB.Menu pipi52 
            Caption         =   "Plant Inventory"
         End
         Begin VB.Menu turn52 
            Caption         =   "Branch Pallet Turnover"
         End
         Begin VB.Menu brzloss502 
            Caption         =   "Branch Out of Stock"
            Visible         =   0   'False
         End
         Begin VB.Menu pipro500502 
            Caption         =   "Brenham Production Schedule"
            Visible         =   0   'False
         End
         Begin VB.Menu piinv500502 
            Caption         =   "Brenham Inventory"
            Visible         =   0   'False
         End
         Begin VB.Menu pipi57 
            Caption         =   "Staged Products Inventory"
            Visible         =   0   'False
         End
         Begin VB.Menu pibs52 
            Caption         =   "Branch Ingredient Storage"
            Visible         =   0   'False
         End
         Begin VB.Menu pidt52 
            Caption         =   "Dry Trailer Order"
         End
         Begin VB.Menu brtrn52 
            Caption         =   "Browse Transports"
         End
         Begin VB.Menu pits52 
            Caption         =   "Trailer Schedule"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu pisp 
         Caption         =   "Snack Plant"
         Begin VB.Menu pipisp 
            Caption         =   "Plant Inventory"
         End
         Begin VB.Menu pidt53 
            Caption         =   "Dry Trailer Order"
         End
      End
      Begin VB.Menu bpskuord 
         Caption         =   "Plant SKU Orders"
      End
      Begin VB.Menu bpalordtot 
         Caption         =   "Branch Pallet Order Totals"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bcode As String
Private Sub check_bcode()
    Dim filler As String, i As Integer
    Dim ds As ADODB.Recordset, s As String                                      'jv021616
    'If Form1.wdbranch = "SU" Then
    '    filler = "Your account has been granted access to all branches.  "
    'End If
    'If Left(Form1.wdbranch, 1) = "R" Then
    '    filler = "Your account has been granted access to Region "
    '    filler = filler & Right(Form1.wdbranch, 1) & ".  "
    'End If
    'If Left(Form1.wdbranch, 1) = "D" Then
    '    filler = "Your account has been granted access to all branches in your division.  "
    'End If
    If bclist.ListCount > 1 Then
        'filler = filler & "Enter the branch code that you wish to view."
        'bcode = InputBox(filler, "Hello " & Form1.wduser & "...", bcode)
        'If Len(bcode) = 0 Then Exit Sub
        'For i = 0 To bclist.ListCount - 1
        '    If Val(bcode) = Val(bclist.List(i)) Then  'jv111501
        '        bclist.ListIndex = i
        '        Exit For
        '    End If
        'Next i
        'If Val(bcode) <> Val(bclist) Then    'jv111501
        '    filler = "The Branch Code entered: " & bcode & " is not accessible to your account."
        '    MsgBox filler, vbOKOnly, "Sorry, try again..."
        '    Exit Sub
        'End If
        bcode = bclist
        rco50.Enabled = False                                                                   'jv021616
        rco51.Enabled = False                                                                   'jv021616
        rco52.Enabled = False                                                                   'jv021616
        s = "select plantwhs from bimp where branchwhs = '" & Format(Val(bcode), "000") & "'"   'jv021616
        s = s & " and sku = '777'"                                                              'jv021616
        Set ds = wdb.Execute(s)                                                                 'jv021616
        If ds.BOF = False Then                                                                  'jv021616
            ds.MoveFirst                                                                        'jv021616
            Do Until ds.EOF                                                                     'jv021616
                If ds!plantwhs = "T10" Then rco50.Enabled = True                                'jv021616
                If ds!plantwhs = "K10" Then rco51.Enabled = True                                'jv021616
                If ds!plantwhs = "A10" Then rco52.Enabled = True                                'jv021616
                ds.MoveNext                                                                     'jv021616
            Loop                                                                                'jv021616
        End If                                                                                  'jv021616
        ds.Close                                                                                'jv021616
        'rco50.Enabled = True
        'rco51.Enabled = True
        'rco52.Enabled = True
        'If Len(Dir(Form1.webdir & "\stock\ro50" & Format(Val(bcode), "00") & ".txt")) = 0 Then rco50.Enabled = False
        'If Len(Dir(Form1.webdir & "\stock\ro51" & Format(Val(bcode), "00") & ".txt")) = 0 Then rco51.Enabled = False
        'If Len(Dir(Form1.webdir & "\stock\ro52" & Format(Val(bcode), "00") & ".txt")) = 0 Then rco52.Enabled = False
    Else
        bcode = Form1.wdbranch
    End If
    filler = Form1.webdir & "\counts\tstation." & Format(Val(bcode), "000")     'jv061818
    tstationrpt.Visible = False                                                 'jv061818
    If Len(Dir(filler)) > 0 Then                                                'jv061818
        tstationrpt.Visible = True                                              'jv061818
    End If                                                                      'jv061818
    
End Sub

Private Sub batreleases_Click()
    brzreleases.Show
End Sub

Private Sub blkbill_Click()
    browzbill.Show
End Sub

Private Sub bpalordtot_Click()
    brzbranchpalship.Show
End Sub

Private Sub bpskuord_Click()
    brzplantorders.Show
End Sub

Private Sub brennotes_Click()
    check_bcode
    Form4.brcode = Format(Val(bcode), "00")
    Form4.Show
End Sub

Private Sub broos_Click()                           'jv032318
    check_bcode
    'Form14.Caption = "Branch Out of Stock - " & branchrec(Val(bcode)).branchname
    'Form14.whsno = Format(Val(bcode), "000")
    'Form14.qstr = "broos"
    'Form14.Show
    brzstkouts.Show
End Sub

Private Sub brorders_Click()
    'check_bcode
    Dim filler As String, i As Integer
    If Form1.wdbranch = "SU" Then
        filler = "Your account has been granted access to all branches.  "
    End If
    If Left(Form1.wdbranch, 1) = "R" Then
        filler = "Your account has been granted access to Region "
        filler = filler & Right(Form1.wdbranch, 1) & ".  "
    End If
    If Left(Form1.wdbranch, 1) = "D" Then
        filler = "Your account has been granted access to all branches in your division.  "
    End If
    If bclist.ListCount > 1 Then
        filler = filler & "Enter the branch code that you wish to view."
        bcode = InputBox(filler, "Hello " & Form1.wduser & "...", bcode)
        If Len(bcode) = 0 Then Exit Sub
        For i = 0 To bclist.ListCount - 1
            If bcode = bclist.List(i) Then
                bclist.ListIndex = i
                Exit For
            End If
        Next i
        If bcode <> bclist Then
            filler = "The Branch Code entered: " & bcode & " is not accessible to your account."
            MsgBox filler, vbOKOnly, "Sorry, try again..."
            Exit Sub
        End If
    Else
        bcode = Form1.wdbranch
        If Len(Dir(Form1.webdir & "\orderoff.txt")) > 0 Then
            WebBrowser1.Navigate Form1.webdir & "\orderoff.htm"
            Exit Sub
        End If
    End If
    'If bcode = "52" Then
    '    Form13.Show
    '    DoEvents
    '    Form13.skincode = "3"
    'Else
        'Form3.brcode = Format(Val(bcode), "00")
        'Form3.Caption = "Branch " & Format(Val(bcode), "00") & " Order.."
        'Form3.Show
        Form13.brcode = Format(Val(bcode), "0")
        Form13.Caption = "Branch " & Format(Val(bcode), "00") & " Order.."
        Form13.Show
        
        
        
    'End If
End Sub

Private Sub brtrn50_Click()
    wdbtrkwo.plantno = "50"
    DoEvents
    wdbtrkwo.Show
End Sub

Private Sub brtrn52_Click()
    wdbtrkwo.plantno = "52"
    DoEvents
    wdbtrkwo.Show
End Sub

Private Sub brzloss500_Click()
    brzloss.plantkey = "T10"
    brzloss.Show
End Sub

Private Sub brzloss501_Click()
    brzloss.plantkey = "K10"
    brzloss.Show
End Sub

Private Sub brzloss502_Click()
    brzloss.plantkey = "A10"
    brzloss.Show
End Sub

Private Sub btrn51_Click()
    wdbtrkwo.plantno = "51"
    DoEvents
    wdbtrkwo.Show
End Sub

Private Sub Combo1_Click()
    bclist.ListIndex = Combo1.ListIndex
    check_bcode
    DoEvents
    'MsgBox WebBrowser1.LocationName
    If LCase(Left(WebBrowser1.LocationName, 7)) = "trlstat" Then tstatrpt_Click
    If LCase(Left(WebBrowser1.LocationName, 6)) = "tsched" Then trlsched_Click
    If LCase(Left(WebBrowser1.LocationName, 3)) = "goh" Then gemminv_Click
    If LCase(Left(WebBrowser1.LocationName, 4)) = "ro50" And rco50.Enabled Then rco50_Click
    If LCase(Left(WebBrowser1.LocationName, 4)) = "ro51" And rco51.Enabled Then rco51_Click
    If LCase(Left(WebBrowser1.LocationName, 4)) = "ro52" And rco52.Enabled Then rco52_Click
    If LCase(Left(WebBrowser1.LocationName, 5)) = "rpart" Then rcopart_Click
    If saleinv.Checked Then saleinv_Click      'Sales vs Inventory
        'check_bcode
        'Form8.brcode = Format(Val(bcode), "00")
    'End If
    If invadj.Checked Then invadj_Click        'Inventory Adjustments
    If gemmeop.Checked Then gemmeop_Click      'Oracle End of Period Totals
    If csc.Checked Then                        'Cold Storage Countsheet
        'check_bcode
        Form10.brcode = Format(Val(bcode), "000")
    End If
    If gemmvc.Checked Then gemmvc_Click        'Oracle vs Countsheet
End Sub

Private Sub Command1_Click()
    brzstkouts.Show
End Sub

Private Sub csc_Click()
    check_bcode
    Form10.brcode = Format(Val(bcode), "000")
    Form10.Show
    csc.Checked = True
End Sub

Private Sub discprod_Click()
    WebBrowser1.Navigate Form1.webdir & "\discont.htm"
End Sub

Private Sub dso_Click()
    check_bcode
    Form7.brcode = Format(Val(bcode), "00")
    Form7.Caption = "Branch " & Format(Val(bcode), "00") & " Dry Storage Order.."
    Form7.Show
End Sub

Private Sub Form_Load()
    Dim lpbuff As String * 25
    Dim ret As Long, UserId As String, filler As String
    Dim t1 As String, t2 As String, t3 As String
    Dim addrec As Boolean
    Dim ds As ADODB.Recordset, s As String
    localAppDataPath = Environ("LOCALAPPDATA") & "\WDBrowser"
    
    ' Build local directory
    If DirExists(localAppDataPath) <> True Then
        MkDir (localAppDataPath)
    End If
    
    WebBrowser1.Navigate Form1.webdir & "\bbwd.htm"
    ret = GetUserName(lpbuff, 25)
    UserId = Left(lpbuff, InStr(lpbuff, Chr(0)) - 1)
    'UserId = "pjohnson"
    Open Form1.webdir & "\userlist" For Input As #1
    Line Input #1, filler
    Do Until EOF(1)
        Input #1, t1, t2, t3
        If UCase(t1) = UCase(UserId) Then
            Form1.wduser = t3
            Form1.wdbranch = t2
            Exit Do
        End If
    Loop
    Close #1
    'MsgBox t1 & " " & t3 & " " & t2
    't1 = InputBox("Code:")
    'Form1.wdbranch = t1
    If Len(Form1.wdbranch) > 2 Then
        If UCase(Right(Form1.wdbranch, 2)) = "3G" Then
            syl3gs.Visible = True
            Form1.wdbranch = Left(Form1.wdbranch, Len(Form1.wdbranch) - 2)
        End If
    End If
    If Form1.wdbranch = "500" Or Form1.wdbranch = "01" Then
        pi.Visible = True
        pi01.Visible = True
        pi47.Visible = True
        pi52.Visible = True
        pisp.Visible = True
        syl3gs.Visible = True
    End If
    If Form1.wdbranch = "47" Then
        pi.Visible = True
        pi47.Visible = True
        pisp.Visible = False
    End If
    If Form1.wdbranch = "52" Then
        pi.Visible = True
        pi52.Visible = True
        pisp.Visible = False
        syl3gs.Visible = True
    End If
        
    If Form1.wdbranch = "500" Then Form1.wdbranch = "SU"
    If Form1.wdbranch = "100" Then Form1.wdbranch = "R1"
    If Form1.wdbranch = "200" Then Form1.wdbranch = "R2"
    If Form1.wdbranch = "300" Then Form1.wdbranch = "R3"
    If Form1.wdbranch = "401" Then Form1.wdbranch = "D1"
    If Form1.wdbranch = "402" Then Form1.wdbranch = "D2"
    If Form1.wdbranch = "403" Then Form1.wdbranch = "D3"
    If Form1.wdbranch = "404" Then Form1.wdbranch = "D4"
    If Form1.wdbranch = "405" Then Form1.wdbranch = "D5"
    If Form1.wdbranch = "406" Then Form1.wdbranch = "D6"
    If Form1.wdbranch = "407" Then Form1.wdbranch = "D7"
    If Form1.wdbranch = "408" Then Form1.wdbranch = "D8"
    If Form1.wdbranch = "409" Then Form1.wdbranch = "D9"
        
    Set wdb = CreateObject("ADODB.Connection")
    'wdb.Open "ODBC;DATABASE=WDShip;DSN=wdship"
    wdb.Open "Driver={SQL Server};Server=bbc-08-sqlsvr;DATABASE=WDShip;UID=bbcship500;PWD=brenham500"
    
    'jv 11-6-2003
    If Val(Form1.wdbranch) > 1 And Val(Form1.wdbranch) < 100 Then
        Form1.locdir = "f:\public"
    Else
        Form1.locdir = Form1.webdir
        'Form1.locdir = "f:\public"
    End If
    
    bclist.Clear
    
    s = "select * from valuelists where listname = 'brzdivmap' order by listreturn"
    Set ds = wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If Left(ds!listreturn, 2) = Me.wdbranch Then
                bclist.AddItem Left(ds!listreturn, 2)
                Combo1.AddItem ds!listreturn
            End If
            If ds!listdisplay = Me.wdbranch Then
                bclist.AddItem Left(ds!listreturn, 2)
                Combo1.AddItem ds!listreturn
            End If
            If Me.wdbranch = "SU" Then
                bclist.AddItem Left(ds!listreturn, 2)
                Combo1.AddItem ds!listreturn
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    If bclist.ListCount > 1 Then
        Combo1.ListIndex = 0
        Combo1.Visible = True
    End If
    
    Combo1.ListIndex = 0        'jv010317
    
    'turned off for testing
    Open Form1.webdir & "\userlog" For Append As #1
    filler = Format(Now, "mm-dd-yyyy hh:mm Am/Pm") & " - "
    filler = filler & UserId & " "
    'If Val(Form1.wdbranch) = 0 Then
    If Form1.bclist.ListCount = 0 Then
        filler = filler & "*** Denied ***"
        repmenu.Visible = False
        messages.Visible = False
        trlsched.Visible = False
        brorders.Visible = False
        trsched.Visible = False
        brennotes.Visible = False
        vco.Visible = False
        dso.Visible = False
        csmenu.Visible = False
    Else
        filler = filler & Form1.wduser & " Branch: " & Form1.wdbranch
        repmenu.Visible = True
        messages.Visible = True
        trlsched.Visible = True
        brorders.Visible = True
        trsched.Visible = True
        brennotes.Visible = True
        'vco.Visible = True
        'dso.Visible = True
        csmenu.Visible = True
        'If Form1.wdbranch <> "01" Then
        If Form1.bclist.ListCount = 1 Then
            bcode = bclist
            rco50.Enabled = False                                                                   'jv021616
            rco51.Enabled = False                                                                   'jv021616
            rco52.Enabled = False                                                                   'jv021616
            s = "select plantwhs from bimp where branchwhs = '" & Format(Val(wdbranch), "000") & "'"   'jv021616
            s = s & " and sku = '777'"                                                              'jv021616
            'MsgBox s
            Set ds = wdb.Execute(s)                                                                 'jv021616
            If ds.BOF = False Then                                                                  'jv021616
                ds.MoveFirst                                                                        'jv021616
                Do Until ds.EOF                                                                     'jv021616
                    If ds!plantwhs = "T10" Then rco50.Enabled = True                                'jv021616
                    If ds!plantwhs = "K10" Then rco51.Enabled = True                                'jv021616
                    If ds!plantwhs = "A10" Then rco52.Enabled = True                                'jv021616
                    ds.MoveNext                                                                     'jv021616
                Loop                                                                                'jv021616
            End If                                                                                  'jv021616
            ds.Close                                                                                'jv021616
            'If Len(Dir(Form1.webdir & "\stock\ro50" & Form1.wdbranch & ".txt")) = 0 Then rco50.Enabled = False
            'If Len(Dir(Form1.webdir & "\stock\ro51" & Form1.wdbranch & ".txt")) = 0 Then rco51.Enabled = False
            'If Len(Dir(Form1.webdir & "\stock\ro52" & Form1.wdbranch & ".txt")) = 0 Then rco52.Enabled = False
        End If
    End If
    ' turned off for testing
    'Print #1, filler & " " & bversion
    Print #1, bversion & " " & filler                                   'jv022316
    s = LCase(CurDir)
    If Right(s, 6) <> "public" And Right(s, 4) <> "html" Then           'jv021616
        filler = Format(Now, "mm-dd-yyyy hh:mm Am/Pm") & " - "          'jv012816
        filler = filler & UserId                                        'jv012816
        'Print #1, filler & " " & CurDir                                 'jv012816
        Print #1, bversion & " " & filler & " " & CurDir                'jv022316
    End If                                                              'jv021616
    Close #1
    frmgrid.FormatString = "^Form|^Top|^Left|^Height|^Width"
    frmgrid.ColWidth(0) = 1000
    frmgrid.ColWidth(1) = 800: frmgrid.ColWidth(2) = 800
    frmgrid.ColWidth(3) = 800: frmgrid.ColWidth(4) = 800
    frmgrid.Rows = 1
    On Error Resume Next
    Open "c:\wdbrowse.ini" For Input As #1
    If Err = 53 Then
        frmgrid.AddItem "form1" & Chr$(9) & 105 & Chr$(9) & 105 & Chr$(9) & 6120 & Chr$(9) & 8190
        frmgrid.AddItem "form2" & Chr$(9) & 105 & Chr$(9) & 105 & Chr$(9) & 5130 & Chr$(9) & 7455
        frmgrid.AddItem "form3" & Chr$(9) & 105 & Chr$(9) & 105 & Chr$(9) & 5655 & Chr$(9) & 8235
        frmgrid.AddItem "form4" & Chr$(9) & 105 & Chr$(9) & 105 & Chr$(9) & 4155 & Chr$(9) & 6135
        frmgrid.AddItem "form5" & Chr$(9) & 105 & Chr$(9) & 105 & Chr$(9) & 3465 & Chr$(9) & 5520
    Else
        Do Until EOF(1)
            Input #1, f, t, l, h, w
            frmgrid.AddItem f & Chr$(9) & t & Chr$(9) & l & Chr$(9) & h & Chr$(9) & w
        Loop
    End If
    Close #1
    On Error GoTo 0
    For i = 1 To Form1.frmgrid.Rows - 1
        If Form1.frmgrid.TextMatrix(i, 0) = "form1" Then
            Form1.Top = Val(Form1.frmgrid.TextMatrix(i, 1))
            Form1.Left = Val(Form1.frmgrid.TextMatrix(i, 2))
            Form1.Height = Val(Form1.frmgrid.TextMatrix(i, 3))
            Form1.Width = Val(Form1.frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    
    'Regional Reports
    regshiphist1.Visible = False
    regshiphist2.Visible = False
    regshiphist3.Visible = False
    regshiphist4.Visible = False
    regshiphist5.Visible = False
    regshiphist6.Visible = False
    If Form1.wdbranch = "SU" Then
        regrepmenu.Visible = True
        regshiphist1.Visible = True
        regshiphist2.Visible = True
        regshiphist3.Visible = True
        regshiphist4.Visible = True
        regshiphist5.Visible = True
        regshiphist6.Visible = True
        regshiphist7.Visible = True
    End If
    If Form1.wdbranch = "R1" Then
        regrepmenu.Visible = True
        regshiphist1.Visible = True
        regshiphist2.Visible = True
        regshiphist3.Visible = True
        regshiphist4.Visible = True
        regshiphist5.Visible = True
        regshiphist6.Visible = True
        regshiphist7.Visible = True
        csmenu.Visible = False
        brorders.Visible = False
        dso.Visible = False
        trsched.Visible = False
        brennotes.Visible = False
    End If
    If Form1.wdbranch = "D1" Then
        regrepmenu.Visible = True
        regshiphist1.Visible = True
        regshiphist7.Visible = False
        csmenu.Visible = False
        brorders.Visible = False
        dso.Visible = False
        trsched.Visible = False
        brennotes.Visible = False
    End If
    If Form1.wdbranch = "D2" Then
        regrepmenu.Visible = True
        regshiphist2.Visible = True
        regshiphist7.Visible = False
        csmenu.Visible = False
        brorders.Visible = False
        dso.Visible = False
        trsched.Visible = False
        brennotes.Visible = False
    End If
    If Form1.wdbranch = "D3" Then
        regrepmenu.Visible = True
        regshiphist3.Visible = True
        regshiphist7.Visible = False
        csmenu.Visible = False
        brorders.Visible = False
        dso.Visible = False
        trsched.Visible = False
        brennotes.Visible = False
    End If
    If Form1.wdbranch = "D4" Then
        regrepmenu.Visible = True
        regshiphist4.Visible = True
        regshiphist7.Visible = False
        csmenu.Visible = False
        brorders.Visible = False
        dso.Visible = False
        trsched.Visible = False
        brennotes.Visible = False
    End If
    If Form1.wdbranch = "D5" Then
        regrepmenu.Visible = True
        regshiphist5.Visible = True
        regshiphist7.Visible = False
        'csmenu.Visible = False
        csmenu.Visible = True                       'Scott Evans - Greenville
        brorders.Visible = False
        dso.Visible = False
        trsched.Visible = False
        brennotes.Visible = False
    End If
    If Form1.wdbranch = "D6" Then
        regrepmenu.Visible = True
        regshiphist6.Visible = True
        regshiphist7.Visible = False
        csmenu.Visible = False
        brorders.Visible = False
        dso.Visible = False
        trsched.Visible = False
        brennotes.Visible = False
    End If
    If Form1.wdbranch = "D7" Then
        regrepmenu.Visible = True
        csmenu.Visible = False
        brorders.Visible = False
        dso.Visible = False
        trsched.Visible = False
        brennotes.Visible = False
    End If
    If Form1.wdbranch = "D8" Then
        regrepmenu.Visible = True
        csmenu.Visible = False
        brorders.Visible = False
        dso.Visible = False
        trsched.Visible = False
        brennotes.Visible = False
    End If
    If Form1.wdbranch = "D9" Then
        regrepmenu.Visible = True
        csmenu.Visible = False
        brorders.Visible = False
        dso.Visible = False
        trsched.Visible = False
        brennotes.Visible = False
    End If
    
    If Form1.wdbranch >= "D1" And Form1.wdbranch <= "D9" Then   'jv050616
        'MsgBox Form1.wdbranch
        pi.Visible = True
        pi01.Visible = True
        pipi01.Visible = True
        pibs01.Visible = False
        pibi01.Visible = False
        turn01.Visible = True
        brtrn50.Visible = True
        pidt01.Visible = False
        pi47.Visible = True
        pipi47.Visible = True
        pibs47.Visible = False
        turn47.Visible = True
        pidt47.Visible = False
        btrn51.Visible = True
        pits47.Visible = False
        pi52.Visible = True
        pipi52.Visible = True
        turn52.Visible = True
        pipro500502.Visible = False
        piinv500502.Visible = False
        pipi57.Visible = False
        pibs52.Visible = False
        pidt52.Visible = False
        brtrn52.Visible = True
        pits52.Visible = False
        pisp.Visible = False
        pipisp.Visible = False
        'pidt53.Visible = False
    End If
        
    
    'Set wdb = CreateObject("ADODB.Connection")
    ''wdb.Open "ODBC;DATABASE=WDShip;DSN=wdship"
    'wdb.Open "Driver={SQL Server};Server=bbc-01-wdsql;DATABASE=WDShip;UID=bbcship500;PWD=brenham500"
    Call build_skumast
    Call build_branches
    Call build_plants
    
    
    WebBrowser1.Navigate Form1.webdir & "\bbwd.htm"
End Sub

Private Sub Form_Resize()
    wdbanner.Width = Me.Width - 80
    If Form1.Height > 2000 Then WebBrowser1.Height = Form1.Height - (wdbanner.Height + 750)
    If Form1.Width > 4000 Then WebBrowser1.Width = Form1.Width - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wdb.Close
    End
End Sub

Private Sub gemmeop_Click()
    check_bcode
    Form9.bcode = Format(Val(bcode), "000")
    Form9.Show
    gemmeop.Checked = True
End Sub

Private Sub gemminv_Click()
    check_bcode
    ''If Val(bcode) = 1 Then
    ''    'WebBrowser1.Navigate "s:\wd\data\wdoraoh.htm"
    ''    WebBrowser1.Navigate "s:\wd\html\schedule\TschedT0.htm"
    ''Else
    '    If Len(Dir(Form1.webdir & "\stock\goh" & Format(Val(bcode), "00") & ".htm")) > 0 Then
    '        WebBrowser1.Navigate Form1.webdir & "\stock\goh" & Format(Val(bcode), "00") & ".htm"
    '    Else
    '        Form2.wdfile = Form1.webdir & "\stock\goh." & Format(Val(bcode), "00")
    '        Form2.Caption = "Oracle Branch Inventory"
    '        'Form2.Show
    '    End If
    ''End If
    Form14.Caption = "Oracle Branch Inventory - " & branchrec(Val(bcode)).branchname
    Form14.whsno = Format(Val(bcode), "000")
    Form14.qstr = "gemminv"
    Form14.Show
End Sub

Private Sub gemmvc_Click()
    check_bcode
    Form11.brcode = Format(Val(bcode), "000")
    Form11.Show
    gemmvc.Checked = True
End Sub

Private Sub homepage_Click()
    WebBrowser1.Navigate Form1.webdir & "\bbwd.htm"
End Sub

Private Sub invadj_Click()
    check_bcode
    Webadj.whscode = Format(Val(bcode), "000")
    Webadj.adjfile = Form1.webdir & "\counts\whsadj." & Format(Val(bcode), "000")
    Webadj.Show
    invadj.Checked = True
End Sub

Private Sub messages_Click()
    check_bcode
    Form2.wdfile = Form1.webdir & "\stock\message." & Format(Val(bcode), "00")
    Form2.Caption = "Messages From Brenham..."
    Form2.Show
End Sub

Private Sub newrel_Click()
    check_bcode
    Form12.gw = Format(Val(bcode), "000")
    Form12.Show
End Sub

Private Sub pibi01_Click()
    'brwzbrana.Show
    whssalesbrz.Caption = "B.I.M.P"
    whssalesbrz.qstr = "allwhs"
    whssalesbrz.Show
End Sub

Private Sub pibs01_Click()
    brwzings.brcode = "500"
    brwzings.Show
End Sub

Private Sub pibs47_Click()
    brwzings.brcode = "501"
    brwzings.Show
End Sub

Private Sub pibs52_Click()
    brwzings.brcode = "502"
    brwzings.Show
End Sub

Private Sub pidt01_Click()
    check_bcode
    brwzdrytrl.brcode = Format(Val(bcode), "00")
    brwzdrytrl.Caption = "Brenham Dry Trailer Order.."
    If brwzdrytrl.Combo1.ListCount > 0 Then brwzdrytrl.Show         'jv022516
End Sub

Private Sub pidt47_Click()
    'check_bcode
    bcode = "47"
    brwzdrytrl.brcode = Format(Val(bcode), "00")
    brwzdrytrl.Caption = "Broken Arrow Dry Trailer Order.."
    'brwzdrytrl.Show
    If brwzdrytrl.Combo1.ListCount > 0 Then brwzdrytrl.Show         'jv022516
End Sub

Private Sub pidt52_Click()
    'check_bcode
    bcode = "52"
    brwzdrytrl.brcode = Format(Val(bcode), "00")
    brwzdrytrl.Caption = "Sylacauga Dry Trailer Order.."
    'brwzdrytrl.Show
    If brwzdrytrl.Combo1.ListCount > 0 Then brwzdrytrl.Show         'jv022516
End Sub

Private Sub pidt53_Click()
    bcode = "53"
    brwzdrytrl.brcode = Format(Val(bcode), "00")
    brwzdrytrl.Caption = "Snack Plant Dry Trailer Order.."
    'brwzdrytrl.Show
    If brwzdrytrl.Combo1.ListCount > 0 Then brwzdrytrl.Show         'jv022516
End Sub

Private Sub piinv500502_Click()
    WebBrowser1.Navigate Form1.webdir & "\brana\inv500502.htm"
End Sub

Private Sub pipi01_Click()
    'brwzplana.brcode = "500"
    'brwzplana.Show
    whssalesbrz.Caption = "Brenham Plant Distribution"
    whssalesbrz.qstr = "plana50"
    whssalesbrz.Show
End Sub

Private Sub pipi47_Click()
    'brwzplana.brcode = "501"
    'brwzplana.Show
    whssalesbrz.Caption = "Broken Arrow Plant Distribution"
    whssalesbrz.qstr = "plana51"
    whssalesbrz.Show
End Sub

Private Sub pipi52_Click()
    'brwzplana.brcode = "502"
    'brwzplana.Show
    whssalesbrz.Caption = "Sylacauga Plant Distribution"
    whssalesbrz.qstr = "plana52"
    whssalesbrz.Show
End Sub

Private Sub pipi57_Click()
    brwzplana.brcode = "507"
    brwzplana.Show
End Sub

Private Sub pipisp_Click()
    'brwzplana.brcode = "503"
    'brwzplana.Show
    whssalesbrz.Caption = "Brenham Plant Distribution"
    whssalesbrz.qstr = "plana50"
    whssalesbrz.Show
End Sub

Private Sub pipro500502_Click()
    WebBrowser1.Navigate Form1.webdir & "\brana\pro500502.htm"
End Sub

Private Sub pits47_Click()
    brwztruck.brcode = "501"
    brwztruck.Show
End Sub

Private Sub pits52_Click()
    brwztruck.brcode = "502"
    brwztruck.Show
End Sub

Private Sub rco50_Click()
    check_bcode
    'If Len(Dir(Form1.webdir & "\stock\ro50" & Format(Val(bcode), "00") & ".htm")) > 0 Then
    '    WebBrowser1.Navigate Form1.webdir & "\stock\ro50" & Format(Val(bcode), "00") & ".htm"
    'Else
    '    Form2.wdfile = Form1.webdir & "\stock\ro50" & Format(Val(bcode), "00") & ".txt"
    '    Form2.Caption = "Recommended Brenham Order"
    '    Form2.Show
    'End If
    Form14.Caption = "Recommended Brenham Order - " & branchrec(Val(bcode)).branchname
    Form14.whsno = Format(Val(bcode), "000")
    Form14.qstr = "rco50"
    Form14.Show
End Sub

Private Sub rco51_Click()
    check_bcode
    'If Len(Dir(Form1.webdir & "\stock\ro51" & Format(Val(bcode), "00") & ".htm")) > 0 Then
    '    WebBrowser1.Navigate Form1.webdir & "\stock\ro51" & Format(Val(bcode), "00") & ".htm"
    'Else
    '    Form2.wdfile = Form1.webdir & "\stock\ro51" & Format(Val(bcode), "00") & ".txt"
    '    Form2.Caption = "Recommended Broken Arrow Order"
    '    Form2.Show
    'End If
    Form14.Caption = "Recommended Broken Arrow Order - " & branchrec(Val(bcode)).branchname
    Form14.whsno = Format(Val(bcode), "000")
    Form14.qstr = "rco51"
    Form14.Show
End Sub

Private Sub rco52_Click()
    check_bcode
    'If Len(Dir(Form1.webdir & "\stock\ro52" & Format(Val(bcode), "00") & ".htm")) > 0 Then
    '    WebBrowser1.Navigate Form1.webdir & "\stock\ro52" & Format(Val(bcode), "00") & ".htm"
    'Else
    '    Form2.wdfile = Form1.webdir & "\stock\ro52" & Format(Val(bcode), "00") & ".txt"
    '    Form2.Caption = "Recommended Sylacauga Order"
    '    Form2.Show
    'End If
    Form14.Caption = "Recommended Sylacauga Order - " & branchrec(Val(bcode)).branchname
    Form14.whsno = Format(Val(bcode), "000")
    Form14.qstr = "rco52"
    Form14.Show
End Sub

Private Sub rcopart_Click()
    check_bcode
    'If Len(Dir(Form1.webdir & "\stock\rpart" & Format(Val(bcode), "00") & ".htm")) > 0 Then
    '    WebBrowser1.Navigate Form1.webdir & "\stock\rpart" & Format(Val(bcode), "00") & ".htm"
    'Else
    '    Form2.wdfile = Form1.webdir & "\stock\rpart" & Format(Val(bcode), "00") & ".txt"
    '    Form2.Caption = "Partial Pallet Order"
    '    Form2.Show
    'End If
    Form14.Caption = "Partial Pallet Order - " & branchrec(Val(bcode)).branchname
    Form14.whsno = Format(Val(bcode), "000")
    Form14.qstr = "rcopart"
    Form14.Show
End Sub

Private Sub regreleases_Click()
    WebBrowser1.Navigate Form1.webdir & "\brana\released.htm"
End Sub

Private Sub regshiphist1_Click()
    'WebBrowser1.Navigate Form1.webdir & "\brana\region1ords.htm"
    regionbimp.regkey = "1"
    regionbimp.Show
End Sub

Private Sub regshiphist2_Click()
    'WebBrowser1.Navigate Form1.webdir & "\brana\region2ords.htm"
    regionbimp.regkey = "2"
    regionbimp.Show
End Sub

Private Sub regshiphist3_Click()
    'WebBrowser1.Navigate Form1.webdir & "\brana\region3ords.htm"
    regionbimp.regkey = "3"
    regionbimp.Show
End Sub

Private Sub regshiphist4_Click()
    'WebBrowser1.Navigate Form1.webdir & "\brana\region4ords.htm"
    regionbimp.regkey = "4"
    regionbimp.Show
End Sub

Private Sub regshiphist5_Click()
    'WebBrowser1.Navigate Form1.webdir & "\brana\region5ords.htm"
    regionbimp.regkey = "5"
    regionbimp.Show
End Sub

Private Sub regshiphist6_Click()
    'WebBrowser1.Navigate Form1.webdir & "\brana\region6ords.htm"
    regionbimp.regkey = "6"
    regionbimp.Show
End Sub

Private Sub regshiphist7_Click()
    'WebBrowser1.Navigate Form1.webdir & "\brana\region7ords.htm"
    regionbimp.regkey = "7"
    regionbimp.Show
End Sub

Private Sub saleinv_Click()
    check_bcode
    'Form8.Show
    'DoEvents
    'Form8.brcode = Format(Val(bcode), "00")
    ''Form8.Show
    saleinv.Checked = True
    Form14.Caption = "Sales vs Inventory - " & branchrec(Val(bcode)).branchname
    Form14.whsno = Format(Val(bcode), "000")
    Form14.qstr = "salevinv"
    Form14.Show
End Sub

Private Sub stocksht_Click()
    WebBrowser1.Navigate Form1.webdir & "\stock.htm"
End Sub

Private Sub syl3gs_Click()
    brwzsyl3g.Show
End Sub

Private Sub tktrpt_Click()
    check_bcode
    Form14.Caption = "Ticket BarCodes - " & branchrec(Val(bcode)).branchname
    Form14.whsno = Format(Val(bcode), "000")
    Form14.qstr = "tktrpt"
    Form14.Show
End Sub

Private Sub trlsched_Click()
    check_bcode
    If bcode = 52 Then
        ' Fix whatever the heck Sylacauga is doing...
        If Dir(Form1.webdir & "\schedule\Tsched" & Format(Val(bcode), "00") & ".htm") <> "" Then
            ' Check if the 52 file even exists, serve it if so
            WebBrowser1.Navigate Form1.webdir & "\schedule\Tsched" & Format(Val(bcode), "00") & ".htm"
            Exit Sub
        Else
            ' Otherwise serve the A10 file
            WebBrowser1.Navigate Form1.webdir & "\schedule\TschedA10.htm"
            Exit Sub
        End If
    End If
    WebBrowser1.Navigate Form1.webdir & "\schedule\Tsched" & Format(Val(bcode), "00") & ".htm"
End Sub

Private Sub trsched_Click()
    check_bcode
    Form5.brcode = Format(Val(bcode), "00")
    Form5.Show
End Sub

Private Sub tstationrpt_Click()         'jv061818
    'Dim webfile As String
    'webfile = Form1.webdir & "\counts\tstation" & Format(Val(bclist), "000") & ".htm"
    'WebBrowser1.Navigate webfile
    'check_bcode
    Form14.Caption = "Transfer Station - " & branchrec(Val(bcode)).branchname
    If Val(bcode) = 62 Then Form14.Caption = "Greenville Transfer Station"
    Form14.whsno = Format(Val(bcode), "000")
    Form14.qstr = "tstation"
    Form14.Show
    
End Sub

Private Sub tstatrpt_Click()
    check_bcode
    'If Len(Dir(Form1.webdir & "\stock\trlstat" & Format(Val(bcode), "00") & ".htm")) > 0 Then
    '    WebBrowser1.Navigate Form1.webdir & "\stock\trlstat" & Format(Val(bcode), "00") & ".htm"
    'Else
    '    Form2.wdfile = Form1.webdir & "\stock\trlstat." & Format(Val(bcode), "00")
    '    Form2.Caption = "Trailer Status Report"
    '    Form2.Show
    'End If
    Form14.Caption = "Trailer Status Report - " & branchrec(Val(bcode)).branchname
    Form14.whsno = Format(Val(bcode), "000")
    Form14.qstr = "tstatrpt"
    Form14.Show
End Sub

Private Sub turn01_Click()
    brzturnover.Show
End Sub

Private Sub turn47_Click()
    brzturnover.Show
End Sub

Private Sub turn52_Click()
    brzturnover.Show
End Sub

Private Sub vco_Click()
    check_bcode
    Form6.brcode = Format(Val(bcode), "00")
    Form6.Caption = "Branch " & Format(Val(bcode), "00") & " Vault Clothes Order.."
    Form6.Show
End Sub

Private Sub wdbanner_Click()
    WebBrowser1.Navigate Form1.webdir & "\bbwd.htm"
End Sub

Private Sub xitmenu_Click()
    Dim i As Integer, f As String
    Dim t As Long, l As Long, h As Long, w As Long
    If Form1.WindowState = 0 Then
        For i = 1 To Form1.frmgrid.Rows - 1
            If Form1.frmgrid.TextMatrix(i, 0) = "form1" Then
                Form1.frmgrid.TextMatrix(i, 1) = Form1.Top
                Form1.frmgrid.TextMatrix(i, 2) = Form1.Left
                Form1.frmgrid.TextMatrix(i, 3) = Form1.Height
                Form1.frmgrid.TextMatrix(i, 4) = Form1.Width
                Exit For
            End If
        Next i
    End If
    Open "c:\wdbrowse.ini" For Output As #1
    For i = 1 To frmgrid.Rows - 1
        f = frmgrid.TextMatrix(i, 0)
        t = Val(frmgrid.TextMatrix(i, 1))
        l = Val(frmgrid.TextMatrix(i, 2))
        h = Val(frmgrid.TextMatrix(i, 3))
        w = Val(frmgrid.TextMatrix(i, 4))
        Write #1, f, t, l, h, w
    Next i
    Close #1
    End
End Sub

Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    Dim RetVal As Boolean
    'RetVal = (GetAttr(DirName) = vbDirectory)
    RetVal = (FileLen(DirName) >= 0)
    
    DirExists = RetVal
    Exit Function
ErrorHandler:
    If (Err = 53) Then ' 53 means file was not found at all
        DirExists = False
    End If
    DirExists = False
End Function
