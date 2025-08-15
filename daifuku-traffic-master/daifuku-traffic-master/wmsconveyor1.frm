VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form form1 
   Caption         =   "Daifuku Traffic"
   ClientHeight    =   12720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   12720
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Import SR5 Lanes"
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
      TabIndex        =   70
      Top             =   7560
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   2775
      Left            =   120
      TabIndex        =   69
      Top             =   4920
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4895
      _Version        =   327680
      BackColorFixed  =   16777152
      BackColorSel    =   4210752
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Frame Frame3 
      Caption         =   "Polling: "
      Height          =   1815
      Left            =   10560
      TabIndex        =   59
      Top             =   120
      Width           =   3255
      Begin VB.CheckBox pdflag 
         Caption         =   "Messages"
         Height          =   255
         Left            =   480
         TabIndex        =   62
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox pqflag 
         Caption         =   "Queues"
         Height          =   255
         Left            =   480
         TabIndex        =   61
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox pwflag 
         Caption         =   "Wrappers"
         Height          =   255
         Left            =   480
         TabIndex        =   60
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   10560
      TabIndex        =   44
      Top             =   1920
      Width           =   3255
      Begin VB.Label name3pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   58
         Top             =   5280
         Width           =   3255
      End
      Begin VB.Label name3pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   57
         Top             =   2400
         Width           =   3255
      End
      Begin VB.Label name2pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   56
         Top             =   4920
         Width           =   3255
      End
      Begin VB.Label name2pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   55
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label name1pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   54
         Top             =   4560
         Width           =   3255
      End
      Begin VB.Label name1pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   53
         Top             =   1680
         Width           =   3255
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
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   52
         Top             =   4200
         Width           =   3255
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
         Index           =   0
         Left            =   0
         TabIndex        =   51
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label palnopic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "pallet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   50
         Top             =   3840
         Width           =   3255
      End
      Begin VB.Label palnopic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "pallet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   49
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lotpic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CodeDate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   48
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label lotpic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CodeDate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   47
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label skupic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SKU"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   615
         Index           =   1
         Left            =   0
         TabIndex        =   46
         Top             =   3000
         Width           =   3255
      End
      Begin VB.Label skupic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SKU"
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
         Height          =   495
         Index           =   0
         Left            =   0
         TabIndex        =   45
         Top             =   120
         Width           =   3255
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4575
      Left            =   120
      TabIndex        =   40
      Top             =   7800
      Width           =   13695
      ExtentX         =   24156
      ExtentY         =   8070
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
   Begin VB.CommandButton Command1 
      Caption         =   "Poll Messages"
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
      Left            =   7560
      TabIndex        =   33
      Top             =   4200
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   27
      Top             =   2640
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
      _Version        =   327680
      BackColorFixed  =   65535
      Appearance      =   0
   End
   Begin VB.CheckBox scanlogs 
      Caption         =   "Scan Logs"
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
      Left            =   6840
      TabIndex        =   18
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Crane Conveyors Online"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      Begin VB.CheckBox srstat5 
         Caption         =   "SR-5"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox srstat4 
         Caption         =   "SR-4"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox srstat3 
         Caption         =   "SR-3"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox srstat2 
         Caption         =   "SR-2"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox srstat1 
         Caption         =   "SR-1"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
   End
   Begin VB.Label remitemholdtime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   68
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label remitemholdfile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "v:\sr5\bin\daiRemItemHold.xml"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   67
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label additemholdtime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   66
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label additemholdfile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "v:\sr5\bin\daiAddItemHold.xml"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   65
      Tag             =   "v:\sr5\bin\daiAddItemHold*.xml"
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label cobrcpttime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   64
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label cobrcptfile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "v:\sr5\bin\daiCOBPalletReceipt.xml"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   63
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "2025-01-17"
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
      Left            =   8280
      TabIndex        =   43
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SR5 Lane Data Downloaded:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   42
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label slanedate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "slanedate"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   41
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "..."
      Height          =   255
      Left            =   6840
      TabIndex        =   39
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Label ws4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   4680
      TabIndex        =   38
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tri-Level Wrapper 4"
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
      Index           =   11
      Left            =   2280
      TabIndex        =   37
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label shipordtime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   36
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Label shipordfile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "v:\sr5\bin\daiOrderItemMessage.xml"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   35
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label timelog 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   34
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label wbh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   4680
      TabIndex        =   32
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label wrb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   4680
      TabIndex        =   31
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Backhauls"
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
      Index           =   10
      Left            =   2280
      TabIndex        =   30
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Roller Bed"
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
      Index           =   9
      Left            =   2280
      TabIndex        =   29
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label rcount 
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
      Height          =   255
      Left            =   6840
      TabIndex        =   28
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label queuebc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6840
      TabIndex        =   26
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label wrapbc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "---"
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
      Height          =   255
      Left            =   6840
      TabIndex        =   25
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Size"
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
      Index           =   8
      Left            =   9240
      TabIndex        =   24
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Log Files"
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
      Index           =   7
      Left            =   6840
      TabIndex        =   23
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label logsize2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   22
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label logfile2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "v:\pallogs\move"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   21
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label logsize1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9240
      TabIndex        =   20
      Top             =   720
      Width           =   975
   End
   Begin VB.Label logfile1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "v:\pallogs\recv"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label ws0 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   4680
      TabIndex        =   17
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label wsp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   4680
      TabIndex        =   16
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label ws3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   4680
      TabIndex        =   15
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label ws2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   4680
      TabIndex        =   14
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label ws1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
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
      Left            =   4680
      TabIndex        =   13
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Robot Zero"
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
      Index           =   6
      Left            =   2280
      TabIndex        =   12
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Snack Plant"
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
      Index           =   5
      Left            =   2280
      TabIndex        =   11
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tri-Level Wrapper 3"
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
      Index           =   4
      Left            =   2280
      TabIndex        =   10
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tri-Level Wrapper 2"
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
      Index           =   3
      Left            =   2280
      TabIndex        =   9
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tri-Level Wrapper 1"
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
      Index           =   2
      Left            =   2280
      TabIndex        =   8
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plate Barcode"
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
      Index           =   1
      Left            =   4680
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wrappers"
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
      Index           =   0
      Left            =   2280
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub draw_label(bc As String, idx As Integer)
    Dim i As Integer
    'i = Val(Mid(bc, 1, 3))
    i = Val(Trim(Mid(bc, 1, 4)))                            'jv060117
    bc = UCase(bc)
    skupic(idx).Caption = Trim(Mid(bc, 1, 4))
    'lotpic(idx).Caption = Mid(bc, 5, 8)
    lotpic(idx).Caption = Mid(bc, 5, 9)                     'jv052515
    palnopic(idx).Caption = Mid(bc, 14, 3)
    If Val(palnopic(idx).Caption) > 0 Then palnopic(idx).Caption = Format(Val(palnopic(idx).Caption), "0")
    pkgpic(idx).Caption = labpix(i).package
    name1pic(idx).Caption = labpix(i).name1
    name2pic(idx).Caption = labpix(i).name2
    name3pic(idx).Caption = labpix(i).name3
End Sub

Public Sub dai_poll_messages()
    Dim ds As adodb.Recordset, sqlx As String
    Dim xmname As String, seqid As Long, s As String, cfile As String
    On Error GoTo vberror
    pdflag.Value = 1: DoEvents                      'jv070214
    cfile = dailogs & "daimessages" & Format(Now, "MMddyy") & ".txt"
    Open cfile For Append As #9
    seqid = 0
    sqlx = "select iMessageSequence, sMessageIdentifier FROM WrxToHost"
    sqlx = sqlx & " Order By iMessageSequence"
    Set ds = DaiDb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            xmname = ds!sMessageIdentifier
            seqid = ds!iMessageSequence
            Call read_dai_message(xmname, seqid)
            DoEvents
            s = dailogs & "dai" & xmname & ".xml"
            WebBrowser1.Navigate2 (s)
            DoEvents
            daimesstext = ""
            Call LoadDocument(s, xmname)
            DoEvents
            Print #9, seqid
            Print #9, daimesstext
            sqlx = "Delete From WrxToHost Where iMessageSequence = " & seqid
            DaiDb.Execute sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    Close #9
    pdflag.Value = 0: DoEvents                      'jv070214
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    If eno = -2147467259 Or eno = 3146 Or eno = 52 Then
        'MsgBox "Network is unavailable at current location.", vbOKOnly + vbInformation, eno & " - try another location..."
        Resume
    Else
        Call vb_elog(eno, edesc, "wmsmobile.bas", "add_alternate_dock_pallet", WDUserId)
        If MsgBox(edesc, vbRetryCancel + vbQuestion, "Local network error: add_alternate_dock_pallet: " & eno) = vbRetry Then
            Resume
        Else
            End
        End If
    End If
End Sub

Sub refresh_grid1(psku As String, plot As String, pcode As String, plot2 As String, pcode2 As String, pqty As Integer)
    Dim s As String, i As Integer, q As Integer
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    If pqty < bbpallet_units(psku) Then
        s = part_pallet_whs(psku) & Chr(9)
        Grid1.AddItem s
    Else
        If form1.srstat1.Value = 1 Then
            i = sku_alloc(psku, plot, pcode, plot2, pcode2, "1")            'jv062614
            If i > 0 Then
                q = queue_count("1")
                s = "1" & Chr(9) & w1cap & Chr(9)
                s = s & q & Chr(9)
                s = s & sr_single_sku("1", psku) & Chr(9)
                s = s & Format(w1cap - q, "0") & Chr(9) & i
                Grid1.AddItem s
            End If
        End If
        If form1.srstat2.Value = 1 Then
            i = sku_alloc(psku, plot, pcode, plot2, pcode2, "2")            'jv062614
            If i > 0 Then
                q = queue_count("2")
                s = "2" & Chr(9) & w2cap & Chr(9)
                s = s & q & Chr(9)
                s = s & sr_single_sku("2", psku) & Chr(9)
                s = s & Format(w2cap - q, "0") & Chr(9) & i
                Grid1.AddItem s
            End If
        End If
        If form1.srstat3.Value = 1 Then
            i = sku_alloc(psku, plot, pcode, plot2, pcode2, "3")            'jv062614
            If i > 0 Then
                q = queue_count("3")
                s = "3" & Chr(9) & w3cap & Chr(9)
                s = s & q & Chr(9)
                s = s & sr_single_sku("3", psku) & Chr(9)
                s = s & Format(w3cap - q, "0") & Chr(9) & i
                Grid1.AddItem s
            End If
        End If
        If form1.srstat4.Value = 1 Then
            i = sku_alloc(psku, plot, pcode, plot2, pcode2, "4")            'jv062614
            If i > 0 Then
                q = queue_count("4")
                s = "4" & Chr(9) & w4cap & Chr(9)
                s = s & q & Chr(9)
                s = s & sr_single_sku("4", psku) & Chr(9)
                s = s & Format(w4cap - q, "0") & Chr(9) & i
                Grid1.AddItem s
            End If
        End If
        If form1.srstat5.Value = 1 Then
            i = sku_alloc(psku, plot, pcode, plot2, pcode2, "5")            'jv062614
            If i > 0 Then
                q = queue_count("5")
                s = "5" & Chr(9) & w5cap & Chr(9)
                s = s & q & Chr(9)
                s = s & sr_single_sku("5", psku) & Chr(9)
                s = s & Format(w5cap - q, "0") & Chr(9) & i
                Grid1.AddItem s
            End If
        End If
    End If
                
    If Grid1.Rows = 1 Then Grid1.AddItem "4"   'default to SR4 if no assignments
    
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 3: Grid1.ColSel = 5
    Grid1.Sort = 4
    Grid1.FormatString = "^SR|^Length|^Queues|^Single|^Capacity|^ResvPals"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 800
    Grid1.ColWidth(2) = 800
    Grid1.ColWidth(3) = 800
    Grid1.ColWidth(4) = 800
    Grid1.ColWidth(5) = 800
    rcount.Caption = Val(rcount.Caption) + 1
    Label2.Caption = psku & " " & Format(lastque - 12, "0")
End Sub

Private Sub additemholdtime_Change()
    Dim rkey As Long, xname As String
    If scanlogs.Value = 1 Then
        DoEvents
        rkey = wd_seq("DAIRequests")
        Call write_oracle_request("AddItemHold", rkey)                     'jv042015
        'xname = Dir(Me.additemholdfile.Caption)                 'jv042315
        'xname = Mid(xname, 4, Len(xname) - 3)                   'jv042315
        'xname = Mid(xname, 1, Len(xname) - 4)                   'jv042315
        'Call write_oracle_request(xname, rkey)                  'jv042315
        DoEvents
        WebBrowser1.Navigate additemholdfile                                'jv042015
    End If
End Sub

Private Sub cobrcpttime_Change()
    Dim rkey As Long
    If scanlogs.Value = 1 Then
        DoEvents
        rkey = wd_seq("DAIRequests")
        Call write_oracle_request("COBPalletReceipt", rkey)                     'jv101414
        DoEvents
        WebBrowser1.Navigate cobrcptfile
    End If
End Sub

Private Sub Command1_Click()
    Call poll_queue_tasks
    dai_poll_messages
End Sub

Private Sub Command2_Click()
    Dim i As Integer                                'jv070116
    Screen.MousePointer = 11
    i = MsgBox("Clear Current Lanes", vbYesNoCancel + vbQuestion, "clear current inventory...")
    If i = vbYes Then Call import_sr5_lanes("Y")    'jv070116
    If i = vbNo Then Call import_sr5_lanes("N")     'jv070116
    DoEvents
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    WDUserId = "Console"
    pallogs = "\\bbc-01-prodtrk\wd\pallogs\"                                        'jv092316
    vberror_log = "\\bbc-01-prodtrk\wd\data\conveyorerrors.txt"
    sr5_lane_data = "\\BBC-02-DAIFUKU\data\WRxInventoryData.txt"
    wms_sr5_data = "\\bbc-01-prodtrk\wd\data\sr5.csv"
    If Len(Dir(sr5_lane_data)) > 0 Then slanedate = Format(FileDateTime(sr5_lane_data), "M-d-yyyy h:mm am/pm")
    'bbsr = "odbc;database=wdracks;uid=bbcwd500;pwd=brenham500;dsn=wdsql500"
    'tbbsr = "odbc;database=wdracks;uid=bbcwd500;pwd=brenham500;dsn=wdsql500"
    bbsr = "DRIVER={SQL Server};Server=BBC-08-SQLSVR;database=wdracks;uid=bbcwd500;pwd=brenham500;"
    tbbsr = "DRIVER={SQL Server};Server=BBC-08-SQLSVR;database=wdracks;uid=bbcwd500;pwd=brenham500;"
    daioradb = "odbc;database=bluebell;uid=wrxjhost;pwd=asrs;dsn=bluebell"
    daisqldb = "DRIVER={SQL Server};Server=BBC-08-SQLSVR;database=dbDaifuku;uid=wrxjhost;pwd=asrs;Application Name=Daifuku Traffic"
    dailogs = "\\bbc-01-prodtrk\wd\sr5\bin\"
    
    '--------- Test Files -------------------------------------------------
    'pallogs = "\\bbc-01-prodtrk\wd\testlogs\"                                        'jv092316
    'vberror_log = "\\bbc-01-prodtrk\wd\testlogs\"                                    'jv092316
    'bbsr = "odbc;database=wdracks;dsn=wdracks"                                       'jv091313
    'tbbsr = "odbc;database=wdracks;dsn=wdracks"                                      'jv091313
    'daioradb = "odbc;database=wdracks;dsn=wdracks"
    'dailogs = "\\bbc-01-prodtrk\wd\testlogs\"                                        'jv092316
    'Me.Caption = Me.Caption & " - Tester"
    'shipordfile.Caption = "v:\testlogs\daiOrderItemMessage.xml"
    'cobrcptfile.Caption = "v:\testlogs\daiCOBPalletReceipt.xml"
    'additemholdfile.Caption = "v:\testlogs\daiAddItemHold.xml"                         'jv042015
    'remitemholdfile.Caption = "v:\testlogs\daiRemItemHold.xml"                         'jv042015
    'additemholdfile.Tag = "v:\testlogs\daiAddItemHold*.xml"                         'jv042315
    'remitemholdfile.Tag = "v:\testlogs\daiRemItemHold*.xml"                         'jv042315
    '--------- End Test ---------------------------------------------------
    
    
    Set Wdb = CreateObject("ADODB.Connection")
    Wdb.Open bbsr
    Set DaiDb = CreateObject("ADODB.Connection")
    DaiDb.Open daisqldb
    
    'shipdb = "\\bbc-01-prodtrk\wd\data\shipping.mdb"
    'shipdb = "ODBC;DATABASE=WDship;UID=bbcship500;PWD=brenham500;DSN=wdship500"
    shipdb = "DRIVER={SQL Server};Server=BBC-08-SQLSVR;DATABASE=WDship;UID=bbcship500;PWD=brenham500;DSN=wdship500"
    Call build_branch_tab
    w1cap = 4
    w2cap = 6
    w3cap = 8
    w4cap = 4
    w5cap = 8 '12
    read_barcode_sequences
    WebBrowser1.Navigate2 dailogs & "daiExpectedReceiptMessage.xml"
    If Len(Dir(shipordfile)) > 0 Then
        shipordtime = Format(FileDateTime(shipordfile), "MM-dd-yyyy hh:mm:ss am/pm")
    End If
    If Len(Dir(cobrcptfile)) > 0 Then                                                   'jv100714
        cobrcpttime = Format(FileDateTime(cobrcptfile), "MM-dd-yyyy hh:mm:ss am/pm")    'jv100714
    End If                                                                              'jv100714
    If Len(Dir(additemholdfile)) > 0 Then                                               'jv042015
        additemholdtime = Format(FileDateTime(additemholdfile), "MM-dd-yyyy hh:mm:ss am/pm") 'jv042015
    End If                                                                              'jv042015
    If Len(Dir(remitemholdfile)) > 0 Then                                               'jv042015
        remitemholdtime = Format(FileDateTime(remitemholdfile), "MM-dd-yyyy hh:mm:ss am/pm") 'jv042015
    End If                                                                              'jv042015
    
    labfmtfile = "\\BBC-03-FILESVR\SharedGroups\wd\bin\labfmt.txt"
    load_labpics
    
    pwflag.Value = 0        'jv070214
    pqflag.Value = 0        'jv070214
    pdflag.Value = 0        'jv070214
    
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    Grid1.FormatString = "^SR|^Length|^Queues|^Single|^Capacity|^ResvPals"
    Grid1.ColWidth(0) = 500
    Grid1.ColWidth(1) = 900
    Grid1.ColWidth(2) = 900
    Grid1.ColWidth(3) = 900
    Grid1.ColWidth(4) = 900
    Grid1.ColWidth(5) = 900
    
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 6
    s = "select * from BBC_HostToWrx where bbcstatus = 'PEND'"
    s = s & " Order by imessagesequence, dhostmodifytime"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = Format(ds(0), "M-d-yy h:mm:ss am/pm") & Chr(9) & ds(1) & Chr(9) & ds(2) & Chr(9)
            s = s & ds(3) & Chr(9) & ds(4) & Chr(9) & ds(5)
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid2.FormatString = "<Time|^ID|<MessageType|<Message|<BBC ID|^Status"
    Grid2.ColWidth(0) = 1800
    Grid2.ColWidth(1) = 800
    Grid2.ColWidth(2) = 1800
    Grid2.ColWidth(3) = 3000
    Grid2.ColWidth(4) = 1600
    Grid2.ColWidth(5) = 800
    
    'DoEvents
    'wrapbc = "926 062516 A 035"
    'queuebc = "359 062616 F 011"
End Sub

Private Sub Form_Resize()
    'WebBrowser1.Width = Me.Width - 180
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Wdb.Close
    DaiDb.Close
    End
End Sub

Private Sub logsize1_Change()
    If pwflag.Value = 0 Then Call poll_wrapper_tasks            'jv070214
    If pdflag.Value = 0 Then Call dai_poll_messages             'jv070214
End Sub

Private Sub logsize2_Change()
    If pqflag.Value = 0 Then Call poll_queue_tasks              'jv070214
    If pdflag.Value = 0 Then Call dai_poll_messages             'jv070214
End Sub

Private Sub queuebc_Change()
    If Len(queuebc.Caption) > 10 Then Call draw_label(queuebc.Caption, 1)
End Sub

Private Sub remitemholdtime_Change()
    Dim rkey As Long, xname As String
    If scanlogs.Value = 1 Then
        DoEvents
        rkey = wd_seq("DAIRequests")
        Call write_oracle_request("RemItemHold", rkey)                     'jv042015
        'xname = Dir(Me.remitemholdfile.Caption)                 'jv042315
        'xname = Mid(xname, 4, Len(xname) - 3)                   'jv042315
        'xname = Mid(xname, 1, Len(xname) - 4)                   'jv042315
        'Call write_oracle_request(xname, rkey)                  'jv042315
        DoEvents
        WebBrowser1.Navigate remitemholdfile                                'jv042015
    End If
End Sub

Private Sub scanlogs_Click()
    If scanlogs.Value = 1 Then poll_logs
End Sub

Private Sub shipordtime_Change()
    Dim rkey As Long
    If scanlogs.Value = 1 Then
        DoEvents
        rkey = wd_seq("DAIRequests")
        Call write_oracle_request("OrderItemMessage", rkey)
        DoEvents
        WebBrowser1.Navigate shipordfile
    End If
End Sub

Private Sub timelog_Change()
    If pwflag.Value = 0 Then Call poll_wrapper_tasks            'jv070214
    If pqflag.Value = 0 Then Call poll_queue_tasks              'jv070214
    If pdflag.Value = 0 Then Call dai_poll_messages             'jv070214
End Sub

Private Sub wrapbc_Change()
    If Len(wrapbc.Caption) > 10 Then Call draw_label(wrapbc.Caption, 0)
End Sub

