VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Warehouse Manager"
   ClientHeight    =   12465
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   16080
   LinkTopic       =   "Form1"
   ScaleHeight     =   12465
   ScaleWidth      =   16080
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   11520
      TabIndex        =   30
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   600
      TabIndex        =   29
      Top             =   1200
      Width           =   9735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   12120
      Visible         =   0   'False
      Width           =   15735
   End
   Begin MSFlexGridLib.MSFlexGrid tktgrid 
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   10680
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2990
      _Version        =   327680
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin MSFlexGridLib.MSFlexGrid oragrid 
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   9120
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2778
      _Version        =   327680
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   12120
      TabIndex        =   15
      Top             =   6360
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   7680
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2566
      _Version        =   327680
   End
   Begin MSFlexGridLib.MSFlexGrid Frmgrid 
      Height          =   1695
      Left            =   240
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2990
      _Version        =   327680
      Rows            =   12
      Cols            =   5
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Configuration "
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
      Height          =   6615
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   10095
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2400
         TabIndex        =   22
         Text            =   "2024.07.11"
         Top             =   240
         Width           =   7575
      End
      Begin VB.TextBox plantno 
         Appearance      =   0  'Flat
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
         Left            =   2400
         TabIndex        =   17
         Top             =   2160
         Width           =   7575
      End
      Begin VB.TextBox ftpdir 
         Appearance      =   0  'Flat
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
         Left            =   2400
         TabIndex        =   12
         Top             =   1920
         Width           =   7575
      End
      Begin VB.TextBox tempdir 
         Appearance      =   0  'Flat
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
         Left            =   2400
         TabIndex        =   11
         Top             =   1680
         Width           =   7575
      End
      Begin VB.TextBox schdb 
         Appearance      =   0  'Flat
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
         Left            =   2400
         TabIndex        =   10
         Top             =   1440
         Width           =   7575
      End
      Begin VB.TextBox shipdb 
         Appearance      =   0  'Flat
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
         Left            =   2400
         TabIndex        =   9
         Top             =   1200
         Width           =   7575
      End
      Begin VB.TextBox repdir 
         Appearance      =   0  'Flat
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
         Left            =   2400
         TabIndex        =   8
         Top             =   960
         Width           =   7575
      End
      Begin VB.TextBox BBSR 
         Appearance      =   0  'Flat
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
         Left            =   2400
         TabIndex        =   7
         Top             =   720
         Width           =   7575
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "User ID:"
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
         Left            =   240
         TabIndex        =   28
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label userid 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "userid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2400
         TabIndex        =   27
         Top             =   2880
         Width           =   7575
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SRServer:"
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
         Left            =   240
         TabIndex        =   26
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label srserv 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "srserv"
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
         Left            =   2400
         TabIndex        =   25
         Top             =   2640
         Width           =   7575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Log Directory:"
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
         Left            =   240
         TabIndex        =   24
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label logdir 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label7"
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
         Left            =   2400
         TabIndex        =   23
         Top             =   2400
         Width           =   7575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Version:"
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
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plant Code:"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Schedule DB:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Shipping DB:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FTP Directory:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Temp Directory:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reports:"
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
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BBSR:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
   End
   Begin VB.Menu filemenu 
      Caption         =   "&File"
      Begin VB.Menu xitmenu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edmenu 
      Caption         =   "&Edit"
      Begin VB.Menu edlane 
         Caption         =   "&Lanes"
      End
      Begin VB.Menu edzones 
         Caption         =   "&Zones"
      End
      Begin VB.Menu edracks 
         Caption         =   "&Racks"
      End
      Begin VB.Menu edship 
         Caption         =   "&Shipping - Cranes"
      End
      Begin VB.Menu edshipfl 
         Caption         =   "Shipping - Forklifts"
      End
      Begin VB.Menu edpaltask 
         Caption         =   "Pallet Tasks"
      End
      Begin VB.Menu edsku 
         Caption         =   "SKU &Configuration"
      End
      Begin VB.Menu oplaneedit 
         Caption         =   "Order Pick Lanes"
      End
      Begin VB.Menu drplist 
         Caption         =   "&Drop Lists"
      End
      Begin VB.Menu edpicko 
         Caption         =   "Pick Orders"
      End
      Begin VB.Menu edholdlist 
         Caption         =   "Hold List"
      End
      Begin VB.Menu eddaitasks 
         Caption         =   "Daifuku Messages"
         Begin VB.Menu sqlhosttowrx 
            Caption         =   "BlueBell SQL HostToWrx"
         End
         Begin VB.Menu orahosttowrx 
            Caption         =   "Daifuku Oracle HostToWrk"
         End
         Begin VB.Menu daiplates 
            Caption         =   "Daifuku Plates"
         End
      End
      Begin VB.Menu edvallist 
         Caption         =   "Value Lists"
      End
      Begin VB.Menu edcycnt 
         Caption         =   "Cycle Counts"
      End
   End
   Begin VB.Menu repmenu 
      Caption         =   "&Reports"
      Begin VB.Menu wana 
         Caption         =   "Crane and Rack Totals"
      End
      Begin VB.Menu dalerpt 
         Caption         =   "Crane Totals"
      End
      Begin VB.Menu billyrpt 
         Caption         =   "Rack Totals"
      End
      Begin VB.Menu sylcs5rpt 
         Caption         =   "Sylacuga CS5 and Rack Totals"
      End
      Begin VB.Menu cntsheets 
         Caption         =   "Count Sheets"
      End
      Begin VB.Menu vuezone 
         Caption         =   "Empty Lanes - By Zone"
      End
      Begin VB.Menu ckship 
         Caption         =   "Check Shipping Groups"
      End
      Begin VB.Menu cktasks 
         Caption         =   "Check Shipping Tasks"
      End
      Begin VB.Menu prodtots 
         Caption         =   "Daily Production Totals"
      End
      Begin VB.Menu forkchk 
         Caption         =   "Forklift Check Sheet"
         Begin VB.Menu forksr4 
            Caption         =   "SR-4"
         End
         Begin VB.Menu forkrb 
            Caption         =   "RollerBed"
         End
         Begin VB.Menu forkuser 
            Caption         =   "User Defined"
         End
      End
      Begin VB.Menu lotcodetab 
         Caption         =   "Lot Code Table"
      End
      Begin VB.Menu batpals 
         Caption         =   "Pallet BarCodes"
      End
      Begin VB.Menu pmoves 
         Caption         =   "Pallet Movement"
      End
      Begin VB.Menu r12bpost 
         Caption         =   "R12 Batch Post"
      End
      Begin VB.Menu srlogsr 
         Caption         =   "SR Logs"
      End
      Begin VB.Menu bbmob 
         Caption         =   "Mobile Devices"
      End
      Begin VB.Menu erlogrpt 
         Caption         =   "Rack Error Log"
      End
      Begin VB.Menu pbclabel 
         Caption         =   "Pallet Barcode Label "
      End
      Begin VB.Menu impdaifuku 
         Caption         =   "Import Daifuku Lanes"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function dai_lot_barcode(ssku As String, slot As String) As String
    Dim s As String, syr As String, sdate As String, scode As String, spal As String
    syr = Mid(slot, 1, 2)
    sdate = Mid(slot, 3, 3)
    scode = Mid(slot, 6, 3)
    spal = Mid(slot, 9, 3)
    If Len(ssku) = 4 Then
        s = ssku
    Else
        s = ssku & " "
    End If
    sdate = DateAdd("d", Val(sdate), "12-31-20" & Format(Val(syr) - 1, "00"))
    sdate = DateAdd("yyyy", 2, sdate)
    sdate = Format(sdate, "MMddyy")
    dai_lot_barcode = s & sdate & scode & spal
End Function


Sub menu_build(uname As String)
    Dim ds As ADODB.Recordset, s As String
    Me.edmenu.Visible = False
    Me.edlane.Visible = False
    Me.edzones.Visible = False
    Me.edracks.Visible = False
    Me.edship.Visible = False
    Me.edshipfl.Visible = False
    Me.edpaltask.Visible = False
    Me.edsku.Visible = False
    Me.oplaneedit.Visible = False
    Me.drplist.Visible = False
    Me.edpicko.Visible = False
    Me.edholdlist.Visible = False
    Me.edcycnt.Visible = False
    Me.eddaitasks.Visible = False
    Me.edvallist.Enabled = False
    Me.edvallist.Caption = "..."
    Me.repmenu.Visible = False
    Me.wana.Visible = False
    Me.dalerpt.Visible = False
    Me.billyrpt.Visible = False
    Me.sylcs5rpt.Visible = False
    Me.cntsheets.Visible = False
    Me.vuezone.Visible = False
    Me.ckship.Visible = False
    Me.cktasks.Visible = False
    Me.prodtots.Visible = False
    Me.forkchk.Visible = False
    Me.lotcodetab.Visible = False
    Me.batpals.Visible = False
    Me.pmoves.Visible = False
    Me.r12bpost.Visible = False
    Me.srlogsr.Visible = False
    Me.bbmob.Visible = False
    Me.erlogrpt.Visible = False
    Me.pbclabel.Visible = False
    Me.impdaifuku.Enabled = False
    Me.impdaifuku.Caption = "..."
    s = "select menuname from usermenus where userid = '" & uname & "'"
    If Form1.plantno = "50" Then s = s & " and orgid = '500'"
    If Form1.plantno = "51" Then s = s & " and orgid = '501'"
    If Form1.plantno = "52" Then s = s & " and orgid = '502'"
    'MsgBox s
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            'MsgBox ds!menuname
            If ds!menuname = "edmenu" Then Me.edmenu.Visible = True
            If ds!menuname = "edlane" Then Me.edlane.Visible = True
            If ds!menuname = "edzones" Then Me.edzones.Visible = True
            If ds!menuname = "edracks" Then Me.edracks.Visible = True
            If ds!menuname = "edship" Then Me.edship.Visible = True
            If ds!menuname = "edshipfl" Then Me.edshipfl.Visible = True
            If ds!menuname = "edpaltask" Then Me.edpaltask.Visible = True
            If ds!menuname = "edsku" Then Me.edsku.Visible = True
            If ds!menuname = "oplaneedit" Then Me.oplaneedit.Visible = True
            If ds!menuname = "drplist" Then Me.drplist.Visible = True
            If ds!menuname = "edpicko" Then Me.edpicko.Visible = True
            If ds!menuname = "edholdlist" Then Me.edholdlist.Visible = True
            If ds!menuname = "edcycnt" Then Me.edcycnt.Visible = True
            If ds!menuname = "eddaitasks" Then Me.eddaitasks.Visible = True
            If ds!menuname = "edvallist" Then
                Me.edvallist.Enabled = True
                Me.edvallist.Caption = "Value Lists"
            End If
            
            If ds!menuname = "repmenu" Then Me.repmenu.Visible = True
            If ds!menuname = "wana" Then Me.wana.Visible = True
            If ds!menuname = "dalerpt" Then Me.dalerpt.Visible = True
            If ds!menuname = "billyrpt" Then Me.billyrpt.Visible = True
            If ds!menuname = "sylcs5rpt" Then Me.sylcs5rpt.Visible = True
            If ds!menuname = "sylcs5" Then Me.sylcs5rpt.Visible = True
            If ds!menuname = "cntsheets" Then Me.cntsheets.Visible = True
            If ds!menuname = "vuezone" Then Me.vuezone.Visible = True
            If ds!menuname = "ckship" Then Me.ckship.Visible = True
            If ds!menuname = "cktasks" Then Me.cktasks.Visible = True
            If ds!menuname = "prodtots" Then Me.prodtots.Visible = True
            If ds!menuname = "forkchk" Then Me.forkchk.Visible = True
            If ds!menuname = "lotcodetab" Then Me.lotcodetab.Visible = True
            If ds!menuname = "batpals" Then Me.batpals.Visible = True
            If ds!menuname = "pmoves" Then Me.pmoves.Visible = True
            If ds!menuname = "r12bpost" Then Me.r12bpost.Visible = True
            If ds!menuname = "srlogsr" Then Me.srlogsr.Visible = True
            If ds!menuname = "bbmob" Then Me.bbmob.Visible = True
            If ds!menuname = "erlogrpt" Then Me.erlogrpt.Visible = True
            If ds!menuname = "pbclabel" Then Me.pbclabel.Visible = True
            If ds!menuname = "impdaifuku" Then
                Me.impdaifuku.Enabled = True
                Me.impdaifuku.Caption = "Import Daifuku Lanes"
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
End Sub

Function bb_codedate(lnum As String) As String
    Dim s As String, syr As String, sday As String
    If Len(lnum) = 0 Then
        bb_codedate = " "
        Exit Function
    Else
        bb_codedate = lnum
    End If
    syr = Left(lnum, 2)
    sday = Right(lnum, 3)
    If Val(sday) > 0 Then
        If Val(syr) > 0 Then
            s = DateAdd("d", sday, "1-1-" & syr)
            s = DateAdd("d", -1, s)
            s = DateAdd("yyyy", 2, s)
            bb_codedate = Format(s, "mmddyy")
        Else
            s = DateAdd("d", sday, "1-1-" & Year(Now))
            s = DateAdd("yyyy", 2, s)
            s = Format(s, "mmddyy")
            bb_codedate = syr & Left(s, 4)
        End If
    End If
End Function

Private Sub batpals_Click()
    'batpalpos.Show
    palbarcodes.Show
End Sub

Private Sub bbmob_Click()
    Dim i
    If Me.userid = "jvierus" Then
        bbmobile.Show
    Else
        i = Shell("notepad s:\wd\html\mspupdate.csv", vbMaximizedFocus)
    End If
End Sub

Private Sub billyrpt_Click()
    Dim ds As ADODB.Recordset, s As String, ps As ADODB.Recordset
    Dim tcap As Long, tbb As Long, t4 As Long, tp As Currency
    Dim rt As String, rh As String, rf As String, i As Double
    Dim acap As Long, abb As Long, a4 As Long, ap As Currency
    tcap = 0: tbb = 0: t4 = 0
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 5
    s = "select aisle,count(*) from racks where aisle not in ('M','S') group by aisle order by aisle"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            acap = 0: abb = 0: a4 = 0
            s = "select count(*) from rackpos"
            s = s & " where rackno in (select id from racks where aisle = '" & ds(0) & "')"
            Set ps = Wdb.Execute(s)
            If ps.BOF = False Then
                ps.MoveFirst
                acap = ps(0)
            End If
            ps.Close
            s = "select count(*) from rackpos"
            s = s & " where rackno in (select id from racks where aisle = '" & ds(0) & "')"
            s = s & " and bbc = 'Y' and count_qty > 0"
            Set ps = Wdb.Execute(s)
            If ps.BOF = False Then
                ps.MoveFirst
                abb = ps(0)
            End If
            ps.Close
            s = "select count(*) from rackpos"
            s = s & " where rackno in (select id from racks where aisle = '" & ds(0) & "')"
            s = s & " and bbc = 'N' and count_qty > 0"
            Set ps = Wdb.Execute(s)
            If ps.BOF = False Then
                ps.MoveFirst
                a4 = ps(0)
            End If
            ps.Close
            s = ds(0) & Chr(9)
            s = s & Format(acap, "#") & Chr(9)
            s = s & Format(abb, "#") & Chr(9)
            s = s & Format(a4, "#") & Chr(9)
            tp = (abb + a4) / acap
            s = s & Format(tp, "0.00")
            Grid1.AddItem s
            tcap = tcap + acap
            tbb = tbb + abb
            t4 = t4 + a4
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    Grid1.AddItem ".."
    s = "Total" & Chr(9)
    s = s & tcap & Chr(9)
    s = s & tbb & Chr(9)
    s = s & t4 & Chr(9)
    tp = (tbb + t4) / tcap
    s = s & Format(tp, "0.00")
    Grid1.AddItem s
    
    Grid1.FormatString = "^Aisle|^Capacity|^BB Pallets|^4-Ways|^Pct. Full"
    Grid1.ColWidth(0) = 1200
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 1200
    
    rt = "Rack Totals"
    rh = Format(Now, "mmmm d, yyyy")
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
End Sub

Private Sub ckship_Click()
    Dim ds As ADODB.Recordset, ds2 As ADODB.Recordset
    Dim sqlx As String, x As Double
    Screen.MousePointer = 11
    Open localAppDataPath & "\ckship.txt" For Output As #1
    Print #1, "Check Shipping Groups"
    Print #1, " "
    Print #1, "Whs    SKU    Short"
    sqlx = "select to_whse_num,sku,sum(order_qty-ship_plt_qty)"
    sqlx = sqlx & " from ship_infc"
    sqlx = sqlx & " where ship_status <> 'CANC' and ship_status <> 'DONE'"
    sqlx = sqlx & " and order_qty > ship_plt_qty"
    sqlx = sqlx & " group by to_whse_num,sku"
    sqlx = sqlx & " order by to_whse_num,sku"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = "select sum(qty) from lane"
            sqlx = sqlx & " where whse_num = " & ds!to_whse_num
            sqlx = sqlx & " and sku = '" & ds!sku & "'"
            sqlx = sqlx & " and lane_status <> 'H'"                 'jv013118
            Set ds2 = Wdb.Execute(sqlx)
            'If ds2.RecordCount > 0 Then
            ds2.MoveFirst
            If ds2(0) > 0 Then
            'If ds2.BOF = False Then
                If ds2(0) < ds(2) Then
                    sqlx = " " & ds!to_whse_num & Space(5)
                    sqlx = sqlx & ds!sku & Space(6)
                    sqlx = sqlx & ds(2) - ds2(0)
                    sqlx = sqlx & "   " & ds2(0) & " available "
                    sqlx = sqlx & ds(2) & " ordered"
                    Print #1, sqlx
                End If
            Else
                sqlx = " " & ds!to_whse_num & Space(5)
                sqlx = sqlx & ds!sku & Space(6)
                sqlx = sqlx & ds(2) & "   <--- no product in crane."
                Print #1, sqlx
            End If
            ds2.Close
            ds.MoveNext
        Loop
    End If
    ds.Close
    Print #1, " "
    Print #1, "End of list..."
    Close #1
    Screen.MousePointer = 0
    sqlx = "notepad.exe " & localAppDataPath & "\ckship.txt"
    x = Shell(sqlx, vbNormalFocus)
End Sub

Private Sub cktasks_Click()
    Dim ds As ADODB.Recordset, ds2 As ADODB.Recordset
    Dim sqlx As String, x As Double, hflag As Boolean
    hflag = False                                           'jv013118
    If MsgBox("Use products on hold?", vbQuestion + vbYesNo + vbDefaultButton2, "hold products...") = vbYes Then hflag = True 'jv013118
    Screen.MousePointer = 11
    Open localAppDataPath & "\ckship.txt" For Output As #1
    Print #1, "Check Shipping Tasks"
    Print #1, " "
    Print #1, "Whs    SKU    Short"
    sqlx = "select source,product,count(*) from paltasks where area = 'DOCK'"
    sqlx = sqlx & " and source in ('SR1','SR2','SR3','STAGING','SR5')"
    sqlx = sqlx & " and status = 'PEND' and userid < '0'"
    sqlx = sqlx & " and lotnum < '0'"
    sqlx = sqlx & " group by source,product"
    sqlx = sqlx & " order by source,product"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!source = "STAGING" Then
                sqlx = "select count(*) from rackpos"
                sqlx = sqlx & " where sku = '" & Trim(Left(ds!product, 4)) & "'"
                If hflag = True Then
                    sqlx = sqlx & " and rackno not in (select id from racks where aisle = 'M' and rack = 'OP')" 'jv013118
                Else
                    sqlx = sqlx & " and rackno not in (select id from racks where (aisle = 'M' and rack = 'OP') or hold = '1')" 'jv013118
                End If
                'MsgBox sqlx
                Set ds2 = Wdb.Execute(sqlx)
                ds2.MoveFirst
                If ds2(0) > 0 Then
                    If ds2(0) < ds(2) Then
                        sqlx = " 4" & Space(5)
                        sqlx = sqlx & Trim(Left(ds!product, 4)) & Space(6)
                        sqlx = sqlx & ds(2) - ds2(0)
                        sqlx = sqlx & "   " & ds2(0) & " available "
                        sqlx = sqlx & ds(2) & " ordered"
                        Print #1, sqlx
                    End If
                Else
                    sqlx = " 4" & Space(5)
                    sqlx = sqlx & Trim(Left(ds!product, 4)) & Space(6)
                    sqlx = sqlx & ds(2) & "   <--- no product in racks."
                    Print #1, sqlx
                End If
                ds2.Close
            
            Else
                sqlx = "select sum(qty) from lane"
                sqlx = sqlx & " where whse_num = " & Right(ds!source, 1)
                sqlx = sqlx & " and sku = '" & Trim(Left(ds!product, 4)) & "'"
                If hflag = False Then sqlx = sqlx & " and lane_status <> 'H'"       'jv013118
                Set ds2 = Wdb.Execute(sqlx)
                ds2.MoveFirst
                If ds2(0) > 0 Then
                    If ds2(0) < ds(2) Then
                        sqlx = " " & Right(ds!source, 1) & Space(5)
                        sqlx = sqlx & Trim(Left(ds!product, 4)) & Space(6)
                        sqlx = sqlx & ds(2) - ds2(0)
                        sqlx = sqlx & "   " & ds2(0) & " available "
                        sqlx = sqlx & ds(2) & " ordered"
                        Print #1, sqlx
                    End If
                Else
                    sqlx = " " & Right(ds!source, 1) & Space(5)
                    sqlx = sqlx & Trim(Left(ds!product, 4)) & Space(6)
                    sqlx = sqlx & ds(2) & "   <--- no product in crane."
                    Print #1, sqlx
                End If
                ds2.Close
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Print #1, " "
    Print #1, "End of list..."
    Close #1
    Screen.MousePointer = 0
    sqlx = "notepad.exe " & localAppDataPath & "\ckship.txt"
    x = Shell(sqlx, vbNormalFocus)
End Sub

Private Sub cntsheets_Click()
    Load invrpts
    invrpts.Show
End Sub

Private Sub Command1_Click()
    Dim s As String
    Form7.Show
    'Form19.Show
End Sub

Private Sub Command2_Click()
    'Dim s As String
    's = dai_lot_barcode("803", "17126522038")
    'MsgBox s
    'palcorr.Show
    palbarcodes.Show
End Sub

Private Sub daiplates_Click()
    daiwmsplate.Show
End Sub

Private Sub dalerpt_Click()
    Dim ds As ADODB.Recordset, s As String
    Dim i As Double, k As Integer
    Dim eb As Integer, tp As Integer
    Dim rt As String, rf As String, rh As String
    Grid1.Clear: Grid1.Rows = 6: Grid1.Cols = 7
    s = "select whse_num,qty from lane where zone_num > 0 order by whse_num"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            i = ds!whse_num
            k = ds!qty + 1
            Grid1.TextMatrix(i, 0) = ds!whse_num
            Grid1.TextMatrix(i, k) = Val(Grid1.TextMatrix(i, k)) + 1
            Grid1.TextMatrix(i, 6) = Val(Grid1.TextMatrix(i, 6)) + ds!qty
            ds.MoveNext
        Loop
    End If
    ds.Close
    eb = 0: tp = 0
    For i = 1 To Grid1.Rows - 1
        eb = eb + Val(Grid1.TextMatrix(i, 1))
        tp = tp + Val(Grid1.TextMatrix(i, 6))
        Grid1.TextMatrix(i, 6) = Format(Val(Grid1.TextMatrix(i, 6)), "###,####")
    Next i
    Grid1.AddItem "..."
    s = "Total Pallets" & Chr(9) & Format(tp, "###,####")
    Grid1.AddItem s
    s = "Empty Bays" & Chr(9) & Format(eb, "###,###")
    Grid1.AddItem s
    
    Grid1.FormatString = "^SR|^Empty|^1 Pallet|^2 Pallets|^3 Pallets|^4 Pallets|^Total"
    Grid1.ColWidth(0) = 1400
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 1200
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 1200
    
    rt = "Total Pallets By Warehouse"
    rh = Format(Now, "mmmm d, yyyy h:mm am/pm")
    rf = "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    
    If MsgBox("Send output to printer?", vbQuestion + vbYesNo, "Send to printer...") = vbYes Then
        Call printflexgrid(Printer, Grid1, rt, rh, rf)
    Else
        Call htmlcolorgrid(Me, localAppDataPath & "\htmltemp.htm", Grid1, rt, rh, rf, "linen", "lemonchiffon", "white")
        If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\internet explorer\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
        If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
            i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & localAppDataPath & "\htmltemp.htm", vbNormalFocus)
            Exit Sub
        End If
    End If
End Sub

Private Sub edcycnt_Click()
    wmscycnt.Show
End Sub

Private Sub edholdlist_Click()
    holdlist.Show
End Sub

Private Sub edlane_Click()
    Load Form2
    Form2.Show
End Sub

Private Sub edpaltask_Click()
    Form8.Show
End Sub

Private Sub edpicko_Click()
    Form18.Show
End Sub

Private Sub edracks_Click()
    Load Form4
    Form4.Show
End Sub

Private Sub edship_Click()
    Load shipinfc
    shipinfc.Show
End Sub

Private Sub edshipfl_Click()
    Form16.Show
End Sub

Private Sub edsku_Click()
    Load skuconf
    skuconf.Show
End Sub

Private Sub edvallist_Click()
    wmsvallists.Show
End Sub

Private Sub edzones_Click()
    Load Form3
    Form3.Show
End Sub

Private Sub erlogrpt_Click()
    Form17.Show
End Sub

Private Sub forkrb_Click()
    Form10.reptype = "RB"
    Form10.reptrig = Val(Form10.reptrig) + 1
    Form10.Show
End Sub

Private Sub forksr4_Click()
    Form10.reptype = "SR4"
    Form10.reptrig = Val(Form10.reptrig) + 1
    Form10.Show
End Sub

Private Sub forkuser_Click()
    Form10.reptype = "User"
    Form10.reptrig = Val(Form10.reptrig) + 1
    Form10.Show
End Sub

Private Sub Form_Load()
    Dim f As String, t As Long, l As Long
    Dim h As Long, w As Long, i As Integer
    Dim ret As Long
    Dim lpbuff As String * 25
    ret = GetUserName(lpbuff, 25)
    Me.userid = Left(lpbuff, InStr(lpbuff, Chr(0)) - 1)
    check_hax
    Command1.Visible = False
    'Me.userid = "lwillia"
    If LCase(Me.userid) = "jvierus" Then Command1.Visible = True
    If UCase(Command()) = "BAUSER" Then
        Open "\\bbba-03-dc\f\user\waredist\bin\wd.ini" For Input As #1
    Else
        If UCase(Command()) = "SYUSER" Then
            Open "\\bbsy-02-dc\f\user\waredist\bin\wd.ini" For Input As #1
        Else
            Dim site As String
            site = UCase(Left(Environ$("computername"), 2))
            If site = "BR" Then
                Open "\\bbc-01-prodtrk\wd\bin\wd.ini" For Input As #1
            ElseIf site = "BA" Then
                Open "\\bbba-03-dc\f\user\waredist\bin\wd.ini" For Input As #1
            ElseIf site = "SY" Then
                Open "\\bbsy-02-dc\f\user\waredist\bin\wd.ini" For Input As #1
            Else
                Open "wd.ini" For Input As #1
            End If
            
            'Open "\\bbc-01-prodtrk\wd\bin\wd.ini" For Input As #1
            'Open "\\bbsy-02-dc\f\user\waredist\bin\wd.ini" For Input As #1
            'Open "\\bbba-03-dc\f\user\waredist\bin\wd.ini" For Input As #1
        End If
        
    End If
    Line Input #1, f
    Do Until EOF(1)
        'f = LCase(f): f = Trim(f)
        If LCase(Left$(f, 6)) = "plant=" Then Form1.Caption = Form1.Caption & " " & Right(f, Len(f) - 6)
        If LCase(Left$(f, 10)) = "tempfiles=" Then tempdir = Right$(f, Len(f) - 10)
        If LCase(Left$(f, 8)) = "reports=" Then repdir = Right$(f, Len(f) - 8)
        If LCase(Left$(f, 7)) = "shipdb=" Then shipdb = Right$(f, Len(f) - 7)
        If LCase(Left$(f, 6)) = "schdb=" Then schdb = Right$(f, Len(f) - 6)
        If LCase(Left$(f, 5)) = "bbsr=" Then BBSR = Right$(f, Len(f) - 5)
        If LCase(Left$(f, 4)) = "ftp=" Then ftpdir = Right(f, Len(f) - 4)
        If LCase(Left$(f, 8)) = "plantno=" Then plantno = Right(f, Len(f) - 8)
        If LCase(Left$(f, 8)) = "pallogs=" Then logdir = Right(f, Len(f) - 8)
        If LCase(Left$(f, 7)) = "srserv=" Then srserv = Right(f, Len(f) - 7)
        
        Line Input #1, f
    Loop
    Close #1
    
    'Build local directory
    localAppDataPath = Environ("LOCALAPPDATA") & "\WarehouseManager"
    If DirExists(localAppDataPath) <> True Then
        MkDir (localAppDataPath)
    End If
    'Me.BBSR = "ODBC;DATABASE=WDRacks;DSN=wdracks"
    'Me.shipdb = "ODBC;DATABASE=WDShip;DSN=wdship"
    
    'If Form1.plantno = "51" Then Me.BBSR = "ODBC;DATABASE=WDRacks;DSN=wdracks"
    'If Form1.plantno = "52" Then Me.BBSR = "ODBC;DATABASE=SYRacks;UID=bbcwd502;PWD=alabama502;DSN=wdsql502"
    'If Form1.plantno = "50" Then Me.BBSR = "ODBC;DATABASE=WDRacks;UID=bbcwd500;PWD=brenham500;DSN=wdsql500"
    'If Form1.plantno = "50" And logdir = "logdir" Then Me.logdir = "\\bbc-01-prodtrk\wd\pallogs\"
    If Form1.plantno = "51" And logdir = "logdir" Then Me.logdir = "\\bbba-03-dc\f\user\waredist\data\pallogs\"
    If Form1.plantno = "52" And logdir = "logdir" Then Me.logdir = "\\bbsy-02-dc\f\user\waredist\data\pallogs\"
    If Form1.srserv = "srserv" Then srserv = "\\bbc-01-wdmgmt"
    'Me.logdir = "\\bbc-01-prodtrk\wd\testlogs\"
    'logdir = "v:\testlogs\"
    'logdir = "U:\"
    wdlogdir = Me.logdir '"v:\testlogs\"
    
    vberror_log = "\\bbc-01-prodtrk\wd\temp\sqlerrors.txt"              'jv123114
    vberror_log = "\\BBC-03-FILESVR\SharedGroups\wd\html\images\sqlerrors.txt"                     'jv061015
    WDbbsr = Me.BBSR                                                    'jv123114
    Set Wdb = CreateObject("ADODB.Connection")                          'jv123114
    Wdb.Open WDbbsr                                                     'jv123114
    WDUserId = Me.userid                                                'jv123114
    Set Sdb = CreateObject("ADODB.Connection")                          'jv060216
    Sdb.Open Me.shipdb                                                  'jv060216
    
    If UCase(Command()) = "BAUSER" Then
        Call menu_build("bauser")
    Else
        If UCase(Command()) = "SYUSER" Then
            Call menu_build("syluser")
        Else
            Call menu_build(Me.userid)
            'Call menu_build("jgoff")
        End If
    End If
    
    'If Me.plantno = "51" Then
    '    Call menu_build("bauser")
    'Else
    '    If Me.plantno = "52" Then
    '        Call menu_build("syluser")
    '    Else
    '        Call menu_build(Me.userid)
    '        'Call menu_build("bvincik")
    '    End If
    'End If

    
    Frmgrid.Row = 0
    Frmgrid.Col = 0: Frmgrid.Text = "Form"
    Frmgrid.Col = 1: Frmgrid.Text = "Top"
    Frmgrid.Col = 2: Frmgrid.Text = "Left"
    Frmgrid.Col = 3: Frmgrid.Text = "Height"
    Frmgrid.Col = 4: Frmgrid.Text = "Width"
    Frmgrid.Rows = 1
    On Error Resume Next
    Open localAppDataPath & "\wmsforms.ini" For Input As #1
    If Err = 53 Then
        Frmgrid.AddItem "form1" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 1170 & Chr$(9) & 8160
        Frmgrid.AddItem "form2" & Chr$(9) & 1650 & Chr$(9) & 1650 & Chr$(9) & 6420 & Chr$(9) & 9990
        Frmgrid.AddItem "form3" & Chr$(9) & 705 & Chr$(9) & 15 & Chr$(9) & 7065 & Chr$(9) & 10185
        Frmgrid.AddItem "form4" & Chr$(9) & 795 & Chr$(9) & 105 & Chr$(9) & 7140 & Chr$(9) & 10095
        Frmgrid.AddItem "form5" & Chr$(9) & 1980 & Chr$(9) & 1980 & Chr$(9) & 3870 & Chr$(9) & 4980
        Frmgrid.AddItem "form6" & Chr$(9) & 0 & Chr$(9) & 0 & Chr$(9) & 5430 & Chr$(9) & 10215
        Frmgrid.AddItem "form7" & Chr$(9) & 330 & Chr$(9) & 330 & Chr$(9) & 6180 & Chr$(9) & 6990
        Frmgrid.AddItem "form8" & Chr$(9) & 400 & Chr$(9) & 400 & Chr$(9) & 6500 & Chr$(9) & 6800
        Frmgrid.AddItem "invrpts" & Chr$(9) & 2175 & Chr$(9) & 510 & Chr$(9) & 6390 & Chr$(9) & 6135
        Frmgrid.AddItem "prodrcv" & Chr$(9) & 660 & Chr$(9) & 660 & Chr$(9) & 6705 & Chr$(9) & 5970
        Frmgrid.AddItem "shipinfc" & Chr$(9) & 990 & Chr$(9) & 990 & Chr$(9) & 6420 & Chr$(9) & 6315
        Frmgrid.AddItem "skuconf" & Chr$(9) & 1320 & Chr$(9) & 1320 & Chr$(9) & 6375 & Chr$(9) & 7935
    Else
        Do Until EOF(1)
            Input #1, f, t, l, h, w
            Frmgrid.AddItem f & Chr$(9) & t & Chr$(9) & l & Chr$(9) & h & Chr$(9) & w
        Loop
    End If
    Close #1
    On Error GoTo 0
    For i = 1 To Form1.Frmgrid.Rows - 1
        If Form1.Frmgrid.TextMatrix(i, 0) = "form1" Then
            Form1.Top = Val(Form1.Frmgrid.TextMatrix(i, 1))
            Form1.Left = Val(Form1.Frmgrid.TextMatrix(i, 2))
            Form1.Height = Val(Form1.Frmgrid.TextMatrix(i, 3))
            Form1.Width = Val(Form1.Frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Call build_sku_config
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 80
    oragrid.Width = Me.Width - 80
    tktgrid.Width = Me.Width - 80
End Sub

Private Sub Form_Terminate()
    Call xitmenu_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call xitmenu_Click
End Sub

Private Sub oplanes_Click()
    
End Sub

Private Sub Frame1_DblClick()
    Frame2.Visible = Not Frame2.Visible
End Sub

Private Sub impdaifuku_Click()
    'Replaced with WMS Conveyors
End Sub

Private Sub lotcodetab_Click()
    Form11.Show
End Sub

Private Sub oplaneedit_Click()
    oplanes.Show
End Sub

Private Sub orahosttowrx_Click()
    Form22.Show
End Sub

Private Sub pbclabel_Click()
    Dim u As String, d As String, s As String
    Dim s1 As Integer, s2 As Integer, i As Integer, k As Integer
    Dim ds As ADODB.Recordset, bc As String
    bc = "777 " & Format(DateAdd("yyyy", 2, Now), "MMddyy") & " A 019"
    bc = InputBox("Barcode Information:", "i.e. " & bc, " ")
    If Len(bc) = 0 Then Exit Sub
    i = Val(Trim(Left(bc, 4)))
    If skurec(i).sku = Trim(Left(bc, 4)) Then
        u = skurec(i).uom_type
        d = skurec(i).desc
    Else
        MsgBox "Invalid SKU...", vbOKOnly + vbExclamation, "sorry, try again...."
        Exit Sub
    End If
        
    If UCase(u) = "BULK" Or UCase(u) = "DOZ" Then u = "BULK"
    If UCase(u) = "CUP" Then u = "CUPS"
    If UCase(u) = "3GAL" Then u = "3 GALLON"
    If UCase(u) = "1/2" Then u = "1/2 GAL"
    If UCase(u) = "PT" Then u = "PINTS"
    If UCase(u) = "QT" Then u = "QUARTS"
    If UCase(u) = "12PK" Then u = "12 PACK"
    If UCase(u) = "24PK" Then u = "24 PACK"
    If UCase(u) = "6PK" Then u = "6 PACK"
    If UCase(u) = "8PK" Then u = "8 PACK"
    If UCase(u) = "4PK" Then u = "4 PACK"
    If UCase(u) = "3PK" Then u = "3 PACK"
    If UCase(u) = "HALF PT" Then u = "HALF PINT"
    
    pallabprt.Show
    If MsgBox("Send to printer?", vbYesNo + vbQuestion, "Ready to print?") = vbYes Then
        pallabprt.prtdevice = "Printer"
        Printer.PaperSize = 5
    Else
        pallabprt.prtdevice = "Screen"
    End If
    pallabprt.skulab = Trim(Left(bc, 4))
    pallabprt.desc1lab = d
    pallabprt.pkglab = u
    pallabprt.lotlab = Mid(bc, 5, 9)                                'jv070115
    pallabprt.seqlab = Right(bc, 3)
    pallabprt.ptrig = bc
    If pallabprt.prtdevice = "Printer" Then Printer.EndDoc
End Sub

Private Sub pmoves_Click()
    Form15.Show
End Sub

Private Sub prodtots_Click()
    'Replaced with pallet movement reports
End Sub

Private Sub r12bpost_Click()
    r12batpost.Show
End Sub

Private Sub sqlhosttowrx_Click()
    Form20.Show
End Sub

Private Sub srlogsr_Click()
    srlogs.Show
End Sub

Private Sub sylcs5rpt_Click()
    Form13.Show
End Sub

Private Sub vuezone_Click()
    Dim ds As ADODB.Recordset, sqlx As String, x As Double
    Screen.MousePointer = 11
    Open localAppDataPath & "\vuezone.txt" For Output As #1
    Print #1, "Empty Lanes - By Zone"
    Print #1, " "
    Print #1, "Whs   Zone   Lanes"
    sqlx = "select whse_num,zone_num,count(*) from lane"
    sqlx = sqlx & " where sku < '000' and lot_num < '00000'"
    sqlx = sqlx & " and resv_sku < '000' and resv_lot < '00000'"
    sqlx = sqlx & " and zone_num > 0"
    sqlx = sqlx & " group by whse_num,zone_num"
    sqlx = sqlx & " order by whse_num,zone_num"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = " " & ds!whse_num & Space(5)
            sqlx = sqlx & Format(ds!zone_num, "00") & Space(5)
            sqlx = sqlx & ds(2)
            Print #1, sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Print #1, " "
    Print #1, "End of List"
    Close #1
    Screen.MousePointer = 0
    sqlx = "notepad.exe " & localAppDataPath & "\vuezone.txt"
    x = Shell(sqlx, vbNormalFocus)
End Sub

Private Sub wana_Click()
    Form9.Show
End Sub

Private Sub xitmenu_Click()
    Dim i As Integer, f As String
    Dim t As Long, l As Long, h As Long, w As Long
    Wdb.Close                                                       'jv123114
    Sdb.Close                                                       'jv060216
    'MsgBox "bye"
    If Form1.WindowState = 0 Then
        For i = 1 To Form1.Frmgrid.Rows - 1
            If Form1.Frmgrid.TextMatrix(i, 0) = "form1" Then
                Form1.Frmgrid.TextMatrix(i, 1) = Form1.Top
                Form1.Frmgrid.TextMatrix(i, 2) = Form1.Left
                Form1.Frmgrid.TextMatrix(i, 3) = Form1.Height
                Form1.Frmgrid.TextMatrix(i, 4) = Form1.Width
                Exit For
            End If
        Next i
    End If
    Open localAppDataPath & "\wmsforms.ini" For Output As #1
    For i = 1 To Frmgrid.Rows - 1
        f = Frmgrid.TextMatrix(i, 0)
        t = Val(Frmgrid.TextMatrix(i, 1))
        l = Val(Frmgrid.TextMatrix(i, 2))
        h = Val(Frmgrid.TextMatrix(i, 3))
        w = Val(Frmgrid.TextMatrix(i, 4))
        Write #1, f; t; l; h; w
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
