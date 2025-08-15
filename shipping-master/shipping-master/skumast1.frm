VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form skumast1 
   Caption         =   "SKU Master Maintenance"
   ClientHeight    =   12690
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   ScaleHeight     =   12690
   ScaleWidth      =   14640
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command12 
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
      Left            =   8640
      TabIndex        =   92
      Top             =   0
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Label "
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   11400
      TabIndex        =   83
      Top             =   0
      Width           =   2415
      Begin VB.Label name3pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   90
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label name2pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   89
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label name1pic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Name 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   88
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label pkgpic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "pkg"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   87
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label palnopic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Pallet #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   86
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lotpic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "CodeDate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   85
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label skupic 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "SKU"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   84
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Drop Branch Promotion"
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
      TabIndex        =   72
      Top             =   11520
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Add Branch Promotion"
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
      Left            =   6000
      TabIndex        =   71
      Top             =   11520
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   4440
      TabIndex        =   70
      Text            =   "Text5"
      Top             =   12000
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   69
      Text            =   "Text5"
      Top             =   11760
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   68
      Text            =   "Text5"
      Top             =   11520
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   67
      Text            =   "Text5"
      Top             =   11280
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   66
      Text            =   "Text5"
      Top             =   11760
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   65
      Text            =   "Text5"
      Top             =   11520
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   64
      Text            =   "Text5"
      Top             =   11280
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid5 
      Height          =   1095
      Left            =   0
      TabIndex        =   56
      Top             =   10080
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   1931
      _Version        =   327680
      BackColorFixed  =   16761087
      BackColorSel    =   12583104
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "Manufactured By:     "
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
      Height          =   495
      Left            =   0
      TabIndex        =   51
      Top             =   6600
      Width           =   10695
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFF00&
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
         Height          =   255
         Left            =   6480
         TabIndex        =   55
         Top             =   120
         Width           =   1695
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFF00&
         Caption         =   "Sylacauga"
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
         Left            =   5040
         TabIndex        =   54
         Top             =   120
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFF00&
         Caption         =   "Broken Arrow"
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
         Left            =   3360
         TabIndex        =   53
         Top             =   120
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFF00&
         Caption         =   "Brenham"
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
         Left            =   2040
         TabIndex        =   52
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.ListBox vList4 
      Height          =   255
      Left            =   12720
      TabIndex        =   50
      Top             =   9840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox vCombo4 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4800
      TabIndex        =   49
      Text            =   "Combo1"
      Top             =   5760
      Width           =   1695
   End
   Begin VB.ListBox vList3 
      Height          =   255
      Left            =   12720
      TabIndex        =   48
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox vList2 
      Height          =   255
      Left            =   12720
      TabIndex        =   47
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox vList1 
      Height          =   255
      Left            =   12720
      TabIndex        =   46
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox vCombo3 
      Height          =   315
      Left            =   4800
      TabIndex        =   45
      Text            =   "Combo3"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.ComboBox vCombo2 
      Height          =   315
      Left            =   4800
      TabIndex        =   44
      Text            =   "Combo2"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ComboBox vCombo1 
      Height          =   315
      Left            =   4800
      TabIndex        =   43
      Text            =   "Combo1"
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ComboBox skulist 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   41
      Text            =   "Combo1"
      Top             =   0
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid Grid4 
      Height          =   2295
      Left            =   5040
      TabIndex        =   40
      Top             =   7200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4048
      _Version        =   327680
      BackColorFixed  =   16777152
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Drop Product Branch"
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
      Left            =   2640
      TabIndex        =   39
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Add Product Branch"
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
      Left            =   240
      TabIndex        =   38
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Drop W/D Plant"
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
      Left            =   8760
      TabIndex        =   37
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add W/D Plant"
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
      TabIndex        =   36
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add Order Pick Sequence"
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
      Left            =   11400
      TabIndex        =   35
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Drop Discontinued SKU"
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
      Left            =   11400
      TabIndex        =   34
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mark Discontinued SKU"
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
      Left            =   11400
      TabIndex        =   33
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New SKU"
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
      Left            =   6600
      TabIndex        =   32
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Changes"
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
      Left            =   4560
      TabIndex        =   31
      Top             =   0
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid3 
      Height          =   2295
      Left            =   0
      TabIndex        =   30
      Top             =   7200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4048
      _Version        =   327680
      BackColorFixed  =   12632319
      BackColorSel    =   192
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   29
      Text            =   "Text4"
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   28
      Text            =   "Text4"
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   27
      Text            =   "Text4"
      Top             =   5760
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1215
      Left            =   0
      TabIndex        =   23
      Top             =   4440
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   2143
      _Version        =   327680
      BackColorFixed  =   12648384
      BackColorSel    =   49152
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8520
      TabIndex        =   21
      Text            =   "Text3"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   8520
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   8520
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   6
      Left            =   8520
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   8520
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3120
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4683
      _Version        =   327680
      Cols            =   7
      BackColorFixed  =   12648447
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.Label fmtfile 
      Caption         =   "Label11"
      Height          =   255
      Left            =   10080
      TabIndex        =   91
      Top             =   6840
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5880
      TabIndex        =   82
      Top             =   11280
      Width           =   4935
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      TabIndex        =   81
      Top             =   9480
      Width           =   5055
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2640
      TabIndex        =   80
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1920
      TabIndex        =   79
      Top             =   5760
      Width           =   8775
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11280
      TabIndex        =   78
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label tempdir 
      Caption         =   "c:\jvwork"
      Height          =   255
      Left            =   8160
      TabIndex        =   77
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label bbsr 
      Caption         =   "Label6"
      Height          =   255
      Left            =   6000
      TabIndex        =   76
      Top             =   8040
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Label plantno 
      Caption         =   "Label6"
      Height          =   255
      Left            =   6120
      TabIndex        =   75
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label webdir 
      Caption         =   "Label6"
      Height          =   255
      Left            =   5880
      TabIndex        =   74
      Top             =   9120
      Width           =   6255
   End
   Begin VB.Label userid 
      Caption         =   ".."
      Height          =   255
      Left            =   12000
      TabIndex        =   73
      Top             =   7560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
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
      Left            =   3000
      TabIndex        =   63
      Top             =   12000
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
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
      Left            =   3000
      TabIndex        =   62
      Top             =   11760
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
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
      Left            =   3000
      TabIndex        =   61
      Top             =   11520
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
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
      Left            =   3000
      TabIndex        =   60
      Top             =   11280
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
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
      Left            =   120
      TabIndex        =   59
      Top             =   11760
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
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
      Left            =   120
      TabIndex        =   58
      Top             =   11520
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
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
      Left            =   120
      TabIndex        =   57
      Top             =   11280
      Width           =   1335
   End
   Begin VB.Label lsku 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SKU:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
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
      Left            =   120
      TabIndex        =   26
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
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
      Left            =   120
      TabIndex        =   25
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
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
      Left            =   120
      TabIndex        =   24
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label psku 
      Caption         =   "psku"
      Height          =   255
      Left            =   10800
      TabIndex        =   22
      Top             =   10080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Order Pick Seq:"
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
      Left            =   6720
      TabIndex        =   20
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reason:"
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
      Left            =   6720
      TabIndex        =   17
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Discontinued:"
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
      Left            =   6720
      TabIndex        =   16
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label shipdb 
      Caption         =   "Label2"
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   8520
      Visible         =   0   'False
      Width           =   10815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Left            =   6720
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Left            =   6720
      TabIndex        =   6
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Description:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unit Type:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SKU:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu edpallab 
         Caption         =   "Pallet Labels"
      End
      Begin VB.Menu edvallists 
         Caption         =   "Value Lists"
      End
   End
End
Attribute VB_Name = "skumast1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub refresh_vlists()
    Dim ds As adodb.Recordset, s As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    vCombo1.Clear: vCombo2.Clear: vCombo3.Clear
    vList1.Clear: vList2.Clear: vList3.Clear
    skulist.Clear
    s = "select * from valuelists where listname = 'unittype' order by listdisplay"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            vCombo1.AddItem ds!listdisplay
            vList1.AddItem ds!listreturn
            ds.MoveNext
        Loop
    End If
    vCombo1.ListIndex = 0
    ds.Close
    
    s = "select * from prodsources order by sourcename"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            vCombo2.AddItem ds!sourcename
            vList2.AddItem ds!source
            ds.MoveNext
        Loop
    End If
    vCombo2.ListIndex = 0
    ds.Close
    vCombo3.AddItem "All Cranes"
    vList3.AddItem "0"
    s = "select * from warehouses order by whsname"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            vCombo3.AddItem ds!whsname
            vList3.AddItem ds!whs_num
            ds.MoveNext
        Loop
    End If
    vCombo3.ListIndex = 0
    ds.Close
    s = "select * from plants order by plantname"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            vCombo4.AddItem ds!plantname
            vList4.AddItem ds!plant
            ds.MoveNext
        Loop
    End If
    vCombo4.ListIndex = 0
    ds.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "refresh_vlists", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_vlists - Error Number: " & eno
        End
    End If
End Sub

Sub refresh_skumast()
    Dim ds As adodb.Recordset, s As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 7
    skulist.Clear
    s = "select * from skumast order by sku"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!sku & Chr(9)
            s = s & ds!fgunit & Chr(9)
            s = s & ds!fgdesc & Chr(9)
            s = s & ds!psource & Chr(9)
            s = s & ds!whs_num & Chr(9)
            s = s & ds!pallet & Chr(9)
            s = s & ds!numwrap
            Grid1.AddItem s
            skulist.AddItem ds!sku
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "^SKU|^Unit Type|<Description|^Prod Source|^Warehouse|^Pallet Units|^Wrap Units"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 900
    Grid1.ColWidth(1) = 1200
    Grid1.ColWidth(2) = 3000
    Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 1200
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 1200
    skulist.ListIndex = 0
    Screen.MousePointer = 0
    skulist.ListIndex = 0
    Grid1.Row = 1
    Grid1_RowColChange
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "refresh_skumast", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_skumast - Error Number: " & eno
        End
    End If
End Sub

Sub refresh_sku_data()
    Dim ds As adodb.Recordset, s As String, zid As Long
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    Text2(0) = "": Text2(1) = "": Text3 = ""
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 6
    Grid3.Clear: Grid3.Rows = 1: Grid3.Cols = 3
    Grid4.Clear: Grid4.Rows = 1: Grid4.Cols = 6
    Grid5.Clear: Grid5.Rows = 1: Grid5.Cols = 8
    s = "select * from discont where sku = '" & psku & "'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Text2(0).Text = Format(ds!discdate, "MM-dd-yyyy")
        Text2(1).Text = ds!discomm & ""
    End If
    ds.Close
    s = "select * from oplist where sku = '" & psku & "'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Text3.Text = ds!opseq
    End If
    ds.Close
    s = "select * from plantskus where sku = '" & psku & "' order by plant"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!plant & Chr(9)
            s = s & ds!lowstk & Chr(9)
            s = s & ds!outstk & Chr(9)
            s = s & ds!lowflag & Chr(9)
            s = s & ds!outflag & Chr(9)
            s = s & ds!id
            Grid2.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid2.FormatString = "^W/D Plant|^LowStk Qty|^OutStk Qty|^Low|^Out|^Rec ID"
    Grid2.ColWidth(0) = 1200
    Grid2.ColWidth(1) = 1200
    Grid2.ColWidth(2) = 1200
    Grid2.ColWidth(3) = 1200
    Grid2.ColWidth(4) = 1200
    Grid2.ColWidth(5) = 1200
    Grid2_RowColChange
    
    Check1.Value = 0: Check2.Value = 0: Check3.Value = 0: Check4.Value = 0
    s = "select * from plantmfg where sku = '" & psku & "'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        If ds!m500 = "Y" Then Check1.Value = 1
        If ds!m501 = "Y" Then Check2.Value = 1
        If ds!m502 = "Y" Then Check3.Value = 1
        If ds!m503 = "Y" Then Check4.Value = 1
    Else
        zid = wd_seq("Plantmfg", Me.shipdb)
        s = "Insert into plantmfg (id, plant, sku, m500, m501, m502, m503)"
        s = s & " Values (" & zid & ", 500, '" & psku & "', 'N', 'N', 'N', 'N')"
        Sdb.Execute s
    End If
    ds.Close
    
    s = "select bp.id, bp.branch, br.branchname, bp.sku from brprods bp, branches br"
    s = s & " where bp.sku = '" & psku & "'"
    s = s & " and br.branch = bp.branch"
    s = s & " order by br.branchname"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(1) & Chr(9)
            s = s & ds(2) & Chr(9)
            s = s & ds(0)
            Grid3.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid3.FormatString = "^Branch|<Name|^Regional"
    Grid3.ColWidth(0) = 1200
    Grid3.ColWidth(1) = 2200
    Grid3.ColWidth(2) = 1200
    
    s = "select wt.id, wt.whs_num, wm.whsname, wt.count_qty, wt.grp_qty, wt.avail"
    s = s & " from whstotals wt, warehouses wm"
    s = s & " where wt.sku = '" & psku & "'"
    s = s & " and wm.whs_num = wt.whs_num"
    s = s & " order by wt.whs_num"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(1) & Chr(9)
            s = s & ds(2) & Chr(9)
            s = s & ds(3) & Chr(9)
            s = s & ds(4) & Chr(9)
            s = s & ds(5) & Chr(9)
            's = s & ds(0)
            Grid4.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid4.FormatString = "^Whs|<Name|^Count|^Grouped|^Avail" '|^Rec ID"
    Grid4.ColWidth(0) = 600
    Grid4.ColWidth(1) = 1600
    Grid4.ColWidth(2) = 1000
    Grid4.ColWidth(3) = 1000
    Grid4.ColWidth(4) = 1000
    Grid4.ColWidth(5) = 0 '1000
    
    s = "select * from promos where sku = '" & psku & "'"
    s = s & " order by plant, branch"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds(1) & Chr(9)
            s = s & ds(3) & Chr(9)
            s = s & ds(4) & Chr(9)
            s = s & Format(ds(5), "MM-dd-yyyy") & Chr(9)
            s = s & Format(ds(6), "MM-dd-yyyy") & Chr(9)
            s = s & ds(7) & Chr(9)
            s = s & ds(8) & Chr(9)
            s = s & ds(0)
            Grid5.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid5.FormatString = "^Branch|^Pallets|^Units|^Start|^End|^Plant|^New Product|^Promotion"
    Grid5.ColWidth(0) = 800
    Grid5.ColWidth(1) = 1200
    Grid5.ColWidth(2) = 1200
    Grid5.ColWidth(3) = 1200
    Grid5.ColWidth(4) = 1200
    Grid5.ColWidth(5) = 1200
    Grid5.ColWidth(6) = 1200
    Grid5.ColWidth(7) = 1200
    Grid5_RowColChange
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "refresh_sku_data", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_sku_data - Error Number: " & eno
        End
    End If
End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim s As String, f As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    If Check1.Value = 1 Then
        f = "Y"
    Else
        f = "N"
    End If
    s = "Update plantmfg set m500 = '" & f & "' where sku = '" & psku & "'"
    Sdb.Execute s
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "check1_mouseup", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " check1_mouseup - Error Number: " & eno
        End
    End If
End Sub

Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim s As String, f As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    If Check2.Value = 1 Then
        f = "Y"
    Else
        f = "N"
    End If
    s = "Update plantmfg set m501 = '" & f & "' where sku = '" & psku & "'"
    Sdb.Execute s
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "check2_mouseup", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " check2_mouseup - Error Number: " & eno
        End
    End If
End Sub

Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim s As String, f As String
    Dim eno As Long, edesc As String
    If Check3.Value = 1 Then
        f = "Y"
    Else
        f = "N"
    End If
    On Error GoTo vberror
    s = "Update plantmfg set m502 = '" & f & "' where sku = '" & psku & "'"
    Sdb.Execute s
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "check3_mouseup", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " check3_mouseup - Error Number: " & eno
        End
    End If
End Sub

Private Sub Check4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim s As String, f As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    If Check4.Value = 1 Then
        f = "Y"
    Else
        f = "N"
    End If
    s = "Update plantmfg set m503 = '" & f & "' where sku = '" & psku & "'"
    Sdb.Execute s
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "check4_mouseup", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " check4_mouseup - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command1_Click()
    Dim ds As adodb.Recordset, s As String, i As Integer
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    i = Grid1.Row
    If i = 0 Then Exit Sub
    Screen.MousePointer = 11
    s = "Update skumast set fgunit = '" & Grid1.TextMatrix(i, 1) & "'"
    s = s & ", fgdesc = '" & fixquotes(Grid1.TextMatrix(i, 2)) & "'"
    s = s & ", psource = " & Val(Grid1.TextMatrix(i, 3))
    s = s & ", whs_num = " & Val(Grid1.TextMatrix(i, 4))
    s = s & ", pallet = " & Val(Grid1.TextMatrix(i, 5))
    s = s & ", numwrap = " & Val(Grid1.TextMatrix(i, 6))
    s = s & " Where sku = '" & Grid1.TextMatrix(i, 0) & "'"
    Sdb.Execute s
    If IsDate(Text2(0).Text) = True Then
        s = "select * from discont where sku = '" & Grid1.TextMatrix(i, 0) & "'"
        Set ds = Sdb.Execute(s)
        If ds.BOF = False Then
            Do Until ds.EOF
                s = "Update discont set discdate = '" & Text2(0).Text & "'"
                s = s & ", discomm = '" & Text2(1).Text & "'"
                s = s & " Where id = " & ds!id
                Sdb.Execute s
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    If Val(Text3) > 0 Then
        s = "select * from oplist where sku = '" & Grid1.TextMatrix(i, 0) & "'"
        Set ds = Sdb.Execute(s)
        If ds.BOF = False Then
            s = "Update oplist set opseq = " & Val(Text3)
            s = s & " Where sku = '" & Grid1.TextMatrix(i, 0) & "'"
            Sdb.Execute s
        End If
        ds.Close
    End If
    If Grid2.Rows > 1 Then
        For i = 1 To Grid2.Rows - 1
            s = "Update plantskus set lowstk = " & Grid2.TextMatrix(i, 1)
            s = s & ", outstk = " & Grid2.TextMatrix(i, 2)
            s = s & " Where id = " & Grid2.TextMatrix(i, 5)
            Sdb.Execute s
        Next i
    End If
    
    If Grid5.Rows > 1 Then
        For i = 1 To Grid5.Rows - 1
            s = "Update promos set branch = " & Val(Grid5.TextMatrix(i, 0))
            s = s & ", palqty = " & Val(Grid5.TextMatrix(i, 1))
            s = s & ", unqty = " & Val(Grid5.TextMatrix(i, 2))
            If IsDate(Grid5.TextMatrix(i, 3)) = True Then
                s = s & ", startdate = '" & Format(Grid5.TextMatrix(i, 3), "MM-dd-yyyy") & "'"
            End If
            If IsDate(Grid5.TextMatrix(i, 4)) = True Then
                s = s & ", enddate = '" & Format(Grid5.TextMatrix(i, 4), "MM-dd-yyyy") & "'"
            End If
            s = s & ", plant = " & Val(Grid5.TextMatrix(i, 5))
            If Grid5.TextMatrix(i, 6) = "Y" Then
                s = s & ", newflag = 'Y'"
            Else
                s = s & ", newflag = 'N'"
            End If
            s = s & " Where id = " & Grid5.TextMatrix(i, 7)
            Sdb.Execute s
        Next i
    End If
    
    s = "Update sku_config set description = '" & fixquotes(Text1(2).Text) & "'"
    s = s & ", uom_type = '" & Text1(1).Text & "'"
    s = s & ", uom_per_pallet = " & Val(Text1(5).Text)
    s = s & ", qty_per_pallet = " & Format(Val(Text1(5).Text) / Val(Text1(6).Text), "0")
    s = s & " Where sku = '" & Text1(0).Text & "'"
    Wdb.Execute s
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command1_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command1_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command10_Click()
    Dim ds As adodb.Recordset, s As String, zid As Long, b As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    b = InputBox("Branch #", "Promotion Branch..")
    If Len(b) = 0 Then Exit Sub
    If Val(b) < 1 Or Val(b) > 99 Then Exit Sub
    s = "select branchname from branches where branch = " & b
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        zid = wd_seq("Promos", Me.shipdb)
        s = "Insert into promos (id, branch, sku) values (" & zid & ", " & Val(b) & ", '" & psku & "')"
        Sdb.Execute s
        s = Val(b) & Chr(9)
        s = s & "0" & Chr(9)
        s = s & "0" & Chr(9)
        s = s & Format(Now, "MM-dd-yyyy") & Chr(9)
        s = s & Format(DateAdd("d", 30, Now), "MM-dd-yyyy") & Chr(9)
        s = s & "50" & Chr(9)
        s = s & "N" & Chr(9)
        s = s & zid
        Grid5.AddItem s
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command10_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command10_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command11_Click()
    Dim s As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    If Val(Grid5.TextMatrix(Grid5.Row, 7)) = 0 Then Exit Sub
    s = "Delete from promos where id = " & Grid5.TextMatrix(Grid5.Row, 7)
    Sdb.Execute s
    If Grid5.Rows > 2 Then
        Grid5.RemoveItem Grid5.Row
    Else
        Grid5.Rows = 1
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command11_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command11_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command12_Click()
    Screen.MousePointer = 11
    Call load_labpics
    refresh_vlists
    refresh_skumast
    Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
    Dim s As String, ds As adodb.Recordset
    Dim ssku As String, i As Integer
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    ssku = InputBox("SKU:", "New SKU....")
    If Len(ssku) = 0 Then Exit Sub
    If Val(ssku) < 100 Or Val(ssku) > 9999 Then Exit Sub                'jv082415
    For i = 1 To skulist.ListCount - 1
        If skulist.List(i) = ssku Then
            skulist.ListIndex = i
            Exit Sub
        End If
    Next i
    s = "Insert into skumast (sku, psource, whs_num) values ('" & ssku & "', 1, 0)"
    Sdb.Execute s
    s = "select * from sku_config where sku = '" & ssku & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = True Then
        s = "Insert into sku_config (sku, sku_type, select_method)"
        s = s & " Values ('" & ssku & "', 'F', 'A')"
        Wdb.Execute s
    End If
    ds.Close
    skulist.AddItem ssku
    Grid1.AddItem ssku
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command2_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command2_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command3_Click()
    Dim s As String, zid As Long
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    zid = wd_seq("Discont", Me.shipdb)
    s = "Insert into discont (id, sku, discdate) values (" & zid & ", '" & psku & "', '" & Format(Now, "MM-dd-yyyy") & "')"
    Sdb.Execute s
    Text2(0).Text = Format(Now, "MM-dd-yyyy")
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command3_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command3_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command4_Click()
    Dim s As String, zid As Long
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    s = "Delete from discont where sku = '" & psku & "'"
    Sdb.Execute s
    Text2(0).Text = ""
    Text2(1).Text = ""
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command4_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command4_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command5_Click()
    Dim s As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    s = "Insert into oplist (sku, opseq) Values ('" & psku & "', 0)"
    Sdb.Execute s
    Text3 = "0"
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command5_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command5_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command6_Click()
    Dim s As String, zid As Long, i As Integer, p As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    p = InputBox("Plant [50, 51, 52]:", "Plant Code..")
    If Len(p) = 0 Then Exit Sub
    If p <> "50" And p <> "51" And p <> "52" Then Exit Sub
    For i = 0 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 0) = p Then Exit Sub
    Next i
    zid = wd_seq("Plantskus", Me.shipdb)
    s = "Insert into plantskus (id, plant, sku, lowstk, outstk, lowflag, outflag)"
    s = s & " Values (" & zid & ", '" & p & "', '" & psku & "', 0, 0, 'Y', 'Y')"
    Sdb.Execute s
    s = p & Chr(9) & "0" & Chr(9) & "0" & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & zid
    Grid2.AddItem s
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command6_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command6_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command7_Click()
    Dim s As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    If Val(Grid2.TextMatrix(Grid2.Row, 5)) > 0 Then
        s = "Delete from plantskus where id = " & Grid2.TextMatrix(Grid2.Row, 5)
        Sdb.Execute s
        If Grid2.Rows > 2 Then
            Grid2.RemoveItem Grid2.Row
        Else
            Grid2.Rows = 1
        End If
        Grid2_RowColChange
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command7_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command7_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command8_Click()
    Dim ds As adodb.Recordset, s As String, zid As Long, b As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    b = InputBox("Branch #", "Product Region Branch..")
    If Len(b) = 0 Then Exit Sub
    If Val(b) < 1 Or Val(b) > 99 Then Exit Sub
    s = "select branchname from branches where branch = " & b
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        zid = wd_seq("Brprods", Me.shipdb)
        s = "Insert into brprods (id, branch, sku) values (" & zid & ", " & Val(b) & ", '" & psku & "')"
        Sdb.Execute s
        s = Val(b) & Chr(9) & ds!branchname & Chr(9) & zid
        Grid3.AddItem s
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command8_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command8_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command9_Click()
    Dim s As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    If Val(Grid3.TextMatrix(Grid3.Row, 2)) = 0 Then Exit Sub
    s = "Delete from brprods where id = " & Grid3.TextMatrix(Grid3.Row, 2)
    Sdb.Execute s
    If Grid3.Rows > 2 Then
        Grid3.RemoveItem Grid3.Row
    Else
        Grid3.Rows = 1
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "command9_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command9_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub edpallab_Click()
    pallabcfg.Show
End Sub

Private Sub edvallists_Click()
    'Form5.Show
    wdvalists.Show
End Sub

Private Sub Form_Load()
    Me.userid = Form1.userid
    Me.shipdb = Form1.shipdb
    Me.webdir = Form1.webdir
    Me.plantno = Form1.plantno
    Me.bbsr = Form1.bbsr
    Me.fmtfile = Form1.fmtfile
    
    Call load_labpics
    refresh_vlists
    refresh_skumast
End Sub

Private Sub Grid1_RowColChange()
    Dim c As Integer
    For c = 0 To Grid1.Cols - 1
        Label1(c).Caption = Grid1.TextMatrix(0, c)
        Text1(c).Text = Grid1.TextMatrix(Grid1.Row, c)
    Next c
    If Grid1.Row > 0 Then
        psku = Grid1.TextMatrix(Grid1.Row, 0)
        skupic.Caption = psku
    End If
End Sub

Private Sub Grid2_RowColChange()
    Dim c As Integer
    For c = 0 To 2
        Label4(c).Caption = Grid2.TextMatrix(0, c)
        If Grid2.Row > 0 Then
            Text4(c).Text = Grid2.TextMatrix(Grid2.Row, c)
        Else
            Text4(c).Text = ""
        End If
    Next c
End Sub

Private Sub Grid5_RowColChange()
    Dim c As Integer
    For c = 0 To 6
        Label5(c).Caption = Grid5.TextMatrix(0, c)
        If Grid5.Row > 0 Then
            Text5(c).Text = Grid5.TextMatrix(Grid5.Row, c)
        Else
            Text5(c).Text = ""
        End If
    Next c
End Sub

Private Sub psku_Change()
    refresh_sku_data
End Sub

Private Sub skulist_Click()
    Dim i As Integer
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 0) = skulist Then
            Grid1.Row = i
            Grid1.TopRow = i
            skupic.Caption = skulist
            Exit For
        End If
    Next i
End Sub

Private Sub skupic_Change()
    Dim i As Integer
    If Val(Left(Text1(2).Text, 1)) > 0 Then
        lotpic.Caption = Format(DateAdd("yyyy", 2, Now), "MMddyy") & " X"
    Else
        lotpic.Caption = Format(DateAdd("yyyy", 2, Now), "MMddyy") & " " & Left(Text1(2).Text, 1)
    End If
    palnopic.Caption = "77"
    i = Val(skupic.Caption)
    pkgpic = labpix(i).package
    If Len(pkgpic) > 1 Then
        name1pic = fixamps(labpix(i).name1)
        name2pic = fixamps(labpix(i).name2)
        name3pic = fixamps(labpix(i).name3)
    Else
        name1pic = "No"
        name2pic = "Pallet Label"
        name3pic = "Defined"
    End If
End Sub

Private Sub Text1_Change(Index As Integer)
    Dim i As Integer
    If Grid1.Row = 0 Then Exit Sub
    Grid1.TextMatrix(Grid1.Row, Index) = Text1(Index).Text
    If Index = 1 Then
        For i = 0 To vList1.ListCount - 1
            If vList1.List(i) = Text1(1).Text Then
                vCombo1.ListIndex = i
                Exit For
            End If
        Next i
    End If
    If Index = 3 Then
        For i = 0 To vList2.ListCount - 1
            If vList2.List(i) = Text1(3).Text Then
                vCombo2.ListIndex = i
                Exit For
            End If
        Next i
    End If
    If Index = 4 Then
        For i = 0 To vList3.ListCount - 1
            If vList3.List(i) = Text1(4).Text Then
                vCombo3.ListIndex = i
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Text2_Change(Index As Integer)
    If Len(Text2(0).Text) = 0 And Len(Text2(1).Text) = 0 Then
        Command3.Visible = True
        Command4.Visible = False
    Else
        Command3.Visible = False
        Command4.Visible = True
    End If
End Sub

Private Sub Text3_Change()
    If Val(Text3) > 0 Then
        Command5.Visible = False
    Else
        Command5.Visible = True
    End If
End Sub

Private Sub Text4_Change(Index As Integer)
    Dim i As Integer
    If Grid2.Rows > 1 Then
        Command7.Visible = True
    Else
        Command7.Visible = False
    End If
    If Grid2.Row = 0 Then Exit Sub
    Grid2.TextMatrix(Grid2.Row, Index) = Text4(Index).Text
    If Index = 0 Then
        For i = 0 To vList4.ListCount - 1
            If vList4.List(i) = Text4(0).Text Then
                vCombo4.ListIndex = i
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Text5_Change(Index As Integer)
    If Grid5.Row = 0 Then Exit Sub
    Grid5.TextMatrix(Grid5.Row, Index) = Text5(Index).Text
End Sub

Private Sub vCombo1_Click()
    vList1.ListIndex = vCombo1.ListIndex
    Text1(1).Text = vList1
End Sub

Private Sub vCombo2_Click()
    vList2.ListIndex = vCombo2.ListIndex
    Text1(3).Text = vList2
End Sub

Private Sub vCombo3_Click()
    vList3.ListIndex = vCombo3.ListIndex
    Text1(4).Text = vList3
End Sub

Private Sub vCombo4_Click()
    vList4.ListIndex = vCombo4.ListIndex
End Sub

Private Sub vList1_Click()
    Text1(1).Text = vList1
End Sub

Private Sub vList2_Click()
    Text1(3).Text = vList2
End Sub

Private Sub vList3_Click()
    Text1(4).Text = vList3
End Sub
