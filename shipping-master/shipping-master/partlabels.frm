VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form partlabs 
   Caption         =   "Partial Pallet Labels"
   ClientHeight    =   10155
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   15045
   LinkTopic       =   "Form4"
   ScaleHeight     =   10155
   ScaleWidth      =   15045
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Post/Print Mode "
      Height          =   1095
      Left            =   12360
      TabIndex        =   114
      Top             =   240
      Width           =   2415
      Begin VB.OptionButton Option2 
         Caption         =   "Print Pallet Labels"
         Height          =   255
         Left            =   120
         TabIndex        =   116
         Top             =   720
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Post to HandHelds"
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print List"
      Height          =   375
      Left            =   12840
      TabIndex        =   113
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Print"
      Height          =   375
      Index           =   34
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Print"
      Height          =   375
      Index           =   33
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Print"
      Height          =   375
      Index           =   32
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Print"
      Height          =   375
      Index           =   31
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Print"
      Height          =   375
      Index           =   30
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pallet 35"
      Height          =   375
      Index           =   34
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pallet 34"
      Height          =   375
      Index           =   33
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pallet 33"
      Height          =   375
      Index           =   32
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pallet 32"
      Height          =   375
      Index           =   31
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pallet 31"
      Height          =   375
      Index           =   30
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Print"
      Height          =   375
      Index           =   29
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Print"
      Height          =   375
      Index           =   28
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Print"
      Height          =   375
      Index           =   27
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Print"
      Height          =   375
      Index           =   26
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Print"
      Height          =   375
      Index           =   25
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Print"
      Height          =   375
      Index           =   24
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Print"
      Height          =   375
      Index           =   23
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Print"
      Height          =   375
      Index           =   22
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Print"
      Height          =   375
      Index           =   21
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Print"
      Height          =   375
      Index           =   20
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Pallet 30"
      Height          =   375
      Index           =   29
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Pallet 29"
      Height          =   375
      Index           =   28
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Pallet 28"
      Height          =   375
      Index           =   27
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Pallet 27"
      Height          =   375
      Index           =   26
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pallet 26"
      Height          =   375
      Index           =   25
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pallet 25"
      Height          =   375
      Index           =   24
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pallet 24"
      Height          =   375
      Index           =   23
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pallet 23"
      Height          =   375
      Index           =   22
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pallet 22"
      Height          =   375
      Index           =   21
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pallet 21"
      Height          =   375
      Index           =   20
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Print"
      Height          =   375
      Index           =   19
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Print"
      Height          =   375
      Index           =   18
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Print"
      Height          =   375
      Index           =   17
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Print"
      Height          =   375
      Index           =   16
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Print"
      Height          =   375
      Index           =   15
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Print"
      Height          =   375
      Index           =   14
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Print"
      Height          =   375
      Index           =   13
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Print"
      Height          =   375
      Index           =   12
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Print"
      Height          =   375
      Index           =   11
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Print"
      Height          =   375
      Index           =   10
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Pallet 20"
      Height          =   375
      Index           =   19
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   9600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Pallet 19"
      Height          =   375
      Index           =   18
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Pallet 18"
      Height          =   375
      Index           =   17
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Pallet 17"
      Height          =   375
      Index           =   16
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pallet 16"
      Height          =   375
      Index           =   15
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pallet 15"
      Height          =   375
      Index           =   14
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pallet 14"
      Height          =   375
      Index           =   13
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pallet 13"
      Height          =   375
      Index           =   12
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pallet 12"
      Height          =   375
      Index           =   11
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pallet 11"
      Height          =   375
      Index           =   10
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5280
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Pgrid 
      Height          =   3375
      Left            =   0
      TabIndex        =   37
      Top             =   6600
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5953
      _Version        =   327680
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Print"
      Height          =   375
      Index           =   9
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Print"
      Height          =   375
      Index           =   8
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Print"
      Height          =   375
      Index           =   7
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Print"
      Height          =   375
      Index           =   6
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Print"
      Height          =   375
      Index           =   5
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Print"
      Height          =   375
      Index           =   4
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Print"
      Height          =   375
      Index           =   3
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Print"
      Height          =   375
      Index           =   2
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Print"
      Height          =   375
      Index           =   1
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Print"
      Height          =   375
      Index           =   0
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Pallet 10"
      Height          =   375
      Index           =   9
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Pallet 9"
      Height          =   375
      Index           =   8
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Pallet 8"
      Height          =   375
      Index           =   7
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Pallet 7"
      Height          =   375
      Index           =   6
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Pallet 6"
      Height          =   375
      Index           =   5
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Pallet 5"
      Height          =   375
      Index           =   4
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pallet 4"
      Height          =   375
      Index           =   3
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pallet 3"
      Height          =   375
      Index           =   2
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pallet 2"
      Height          =   375
      Index           =   1
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pallet 1"
      Height          =   375
      Index           =   0
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   7095
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   12515
      _Version        =   327680
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   5400
      TabIndex        =   4
      Top             =   8520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   34
      Left            =   13440
      TabIndex        =   107
      Top             =   9600
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   33
      Left            =   13440
      TabIndex        =   106
      Top             =   9120
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   32
      Left            =   13440
      TabIndex        =   105
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   31
      Left            =   13440
      TabIndex        =   104
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   30
      Left            =   13440
      TabIndex        =   103
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   29
      Left            =   13440
      TabIndex        =   87
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   28
      Left            =   13440
      TabIndex        =   86
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   27
      Left            =   13440
      TabIndex        =   85
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   26
      Left            =   13440
      TabIndex        =   84
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   25
      Left            =   13440
      TabIndex        =   83
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   24
      Left            =   13440
      TabIndex        =   82
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   23
      Left            =   13440
      TabIndex        =   81
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   22
      Left            =   13440
      TabIndex        =   80
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   21
      Left            =   13440
      TabIndex        =   79
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   20
      Left            =   13440
      TabIndex        =   78
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   19
      Left            =   9600
      TabIndex        =   57
      Top             =   9600
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   18
      Left            =   9600
      TabIndex        =   56
      Top             =   9120
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   17
      Left            =   9600
      TabIndex        =   55
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   16
      Left            =   9600
      TabIndex        =   54
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   15
      Left            =   9600
      TabIndex        =   53
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   14
      Left            =   9600
      TabIndex        =   52
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   13
      Left            =   9600
      TabIndex        =   51
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   12
      Left            =   9600
      TabIndex        =   50
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   11
      Left            =   9600
      TabIndex        =   49
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   10
      Left            =   9600
      TabIndex        =   48
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Wrap Totals"
      Height          =   255
      Left            =   9600
      TabIndex        =   26
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   9
      Left            =   9600
      TabIndex        =   25
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   9600
      TabIndex        =   24
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   9600
      TabIndex        =   23
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   9600
      TabIndex        =   22
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   9600
      TabIndex        =   21
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   9600
      TabIndex        =   20
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   9600
      TabIndex        =   19
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   9600
      TabIndex        =   18
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   9600
      TabIndex        =   17
      Top             =   960
      Width           =   975
   End
   Begin VB.Label wc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Trailers:"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Ship Dates:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu edmenu 
      Caption         =   "Edit"
      Begin VB.Menu edqty 
         Caption         =   "Edit Qty"
         Shortcut        =   {F1}
      End
      Begin VB.Menu insline 
         Caption         =   "Insert 2nd Line"
         Shortcut        =   {F2}
      End
      Begin VB.Menu clrtag 
         Caption         =   "Clear Tag"
         Shortcut        =   {F9}
      End
      Begin VB.Menu delrec 
         Caption         =   "Erase Line"
         Shortcut        =   {F5}
      End
   End
End
Attribute VB_Name = "partlabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function calc_date(lotcode As String) As String
    Dim seed As String
    If Left(lotcode, 2) = "00" Then
        seed = "12-31-1999"
    Else
        If Val(lotcode) > 90000 Then
            seed = "12-31-19" & Val(Left(lotcode, 2)) - 1
        Else
            seed = "12-31-20" & Format(Val(Left(lotcode, 2)) - 1, "00")
        End If
    End If
    calc_date = Format(DateAdd("d", Val(Right(lotcode, 3)), seed), "m-d-yyyy")
End Function

Private Sub view_prtlist(pd As Control)
    Dim k As Integer, i As Integer, tstat As String
    Dim p As Integer, rstr As String, pflag As Boolean
    Dim halfwidth, sy As Long, s As String
    Screen.MousePointer = 11
    pd.FontName = "Arial"
    pd.FontSize = 12
    pd.FontBold = True
    
    pd.CurrentX = 0: pd.CurrentY = 0

    pd.FontSize = 12
    pd.FontUnderline = False
    pd.CurrentY = 1440 * 1
    pd.CurrentX = 1440 * 0.75: pd.Print "Partial Pallet Order";
    pd.CurrentX = 1440 * 2.75: pd.Print partlabs.Combo2;
    pd.CurrentX = 1440 * 6.25: pd.Print Left(partlabs.Combo1, 10)
    s = partlabs.Combo2
        
    pd.FontSize = 10
    pd.CurrentY = 1440 * 1.5
    pd.CurrentX = 1440 * 0.75:  pd.Print "Pallet";
    pd.CurrentX = 1440 * 1.25: pd.Print "SKU  Product";
    pd.CurrentX = 1440 * 4.25: pd.Print "Wraps";
    pd.CurrentX = 1440 * 5.25: pd.Print "Code Date(s)"
    pd.FontBold = False
    For i = 1 To partlabs.Grid1.Rows - 1
        If Val(partlabs.Grid1.TextMatrix(i, 0)) > 0 Then
            pd.FontSize = 10
            pd.Print " "
            pd.CurrentX = 1440 * 1 - pd.TextWidth(partlabs.Grid1.TextMatrix(i, 0))
            pd.Print partlabs.Grid1.TextMatrix(i, 0);
            
            pd.CurrentX = 1440 * 1.25
            pd.Print partlabs.Grid1.TextMatrix(i, 2) & "  ";
            pd.Print StrConv(partlabs.Grid1.TextMatrix(i, 3), vbProperCase);
            
            pd.CurrentX = 1440 * 4.55 - pd.TextWidth(partlabs.Grid1.TextMatrix(i, 4))
            pd.Print partlabs.Grid1.TextMatrix(i, 4);
            
            pd.CurrentX = 1440 * 5.25
            pd.Print "___________________________"
            
        End If
    Next i
    
    Screen.MousePointer = 0
End Sub

Private Function read_sae() As Boolean
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String
    Dim ss As adodb.Recordset, bno As Integer, sdesc As String
    Dim palid As String
    Dim pt1 As String, pt2 As String
    Dim i As Integer, j As Integer
    On Error GoTo vberror
    Screen.MousePointer = 11
    If Left(List1, 1) = "T" Then
        bno = 16
        pt1 = Format(bno, "00")               'branch
        pt1 = pt1 & Right(List1, 6)         'account
        pt1 = pt1 & "00"
        pt1 = pt1 & Left(Combo1, 2) & mid(Combo1, 4, 2) & mid(Combo1, 9, 2)
        pt2 = Format(bno, "00")
        pt2 = pt2 & Right(List1, 6)
        pt2 = pt2 & "ZZ"
        pt2 = pt2 & Left(Combo1, 2) & mid(Combo1, 4, 2) & mid(Combo1, 9, 2)
        
        s = "select * from picktasks where palletid >= '" & pt1 & "'"
        s = s & " and palletid < '" & pt2 & "'"
        s = s & " and shipdate = '" & Trim(Left(Combo1, 10)) & "'"
        s = s & " and status in ('PEND', 'PICKED')"     'jv053115
        s = s & " order by palnum, opseq"
    Else
        bno = Val(Left(List1, 2))
        s = "select * from picktasks where branch = " & Left(List1, 2)
        s = s & " and shipdate = '" & Trim(Left(Combo1, 10)) & "'"
        s = s & " and status in ('PEND', 'PICKED')"     'jv053115
        s = s & " order by palnum, opseq"
    End If
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        read_sae = True
        ds.MoveFirst
        Do Until ds.EOF
            s = "select uom_type, description from sku_config where sku = '" & ds!sku & "'"
            Set ss = Wdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                sdesc = ss!uom_type & " " & ss!description
            Else
                sdesc = "------"
            End If
            ss.Close
            s = ds!palnum & Chr(9)
            s = s & Format(ds!opseq, "0000") & Chr(9)
            s = s & ds!sku & Chr(9)
            s = s & sdesc & Chr(9)
            s = s & ds!qty
            Grid1.AddItem s
            ds.MoveNext
        Loop
    Else
        read_sae = False
    End If
    ds.Close
    
    For i = 0 To 34
        wc(i).Caption = " "
        Command2(i).Enabled = False
    Next i
    
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 2
    Grid1.Sort = 3
    Grid1.FillStyle = flexFillRepeat
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        If Val(Grid1.TextMatrix(i, 0)) = 0 Then
            Grid1.CellBackColor = Grid1.BackColor
        Else
            j = CInt(Grid1.TextMatrix(i, 0)) - 1
            Grid1.CellBackColor = Command1(j).BackColor
            wc(j).Caption = Val(wc(j).Caption) + Val(Grid1.TextMatrix(i, 4))
            Command2(j).Enabled = True
        End If
    Next i
    Screen.MousePointer = 0
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "read_sae", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " read_sae - Error Number: " & eno
        End
    End If
End Function

Private Sub post_sae(tno As String)                     'jv091510
    Dim cfile As String, i As Integer, k As Integer
    Dim ds As adodb.Recordset, s As String
    Dim ss As adodb.Recordset, bno As Integer
    Dim p As ptask, palid As String, zid As Long
    'On Error GoTo vberror
    Screen.MousePointer = 11
    If Left(List1, 1) = "T" Then
        bno = 16
        s = "select id,status,userid from picktasks where branch = 16"
        s = s & " and brname = '" & Combo2 & "'"
        s = s & " and shipdate = '" & Trim(Left(Combo1, 10)) & "'"
        s = s & " and palnum = " & tno
        palid = Format(bno, "00")               'branch
        palid = palid & Right(List1, 6)         'account
        palid = palid & Format(Val(tno), "00")  'palnum
        palid = palid & Left(Combo1, 2) & mid(Combo1, 4, 2) & mid(Combo1, 9, 2)
    Else
        bno = Val(Left(List1, 2))
        s = "select id,status,userid from picktasks where branch = " & Left(List1, 2)
        s = s & " and shipdate = '" & Trim(Left(Combo1, 10)) & "'"
        s = s & " and palnum = " & tno
        palid = Format(bno, "000") & " "
        palid = palid & Left(Combo1, 2) & mid(Combo1, 4, 2) & mid(Combo1, 9, 2) & " B "
        palid = palid & Format(Val(tno), "000")
    End If
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
        If Grid1.TextMatrix(i, 0) = tno Then
            s = "select numwrap from skumast where sku = '" & Grid1.TextMatrix(i, 2) & "'"
            Set ss = Sdb.Execute(s)
            If ss.BOF = False Then
                ss.MoveFirst
                k = ss!numwrap
            Else
                k = 1
            End If
            ss.Close
            s = "select * from picktasks where status in ('SHIPPED', 'COMP') order by id"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                s = "update picktasks set branch = " & bno
                s = s & ", brname = '" & Combo2 & "'"
                s = s & ", shipdate = '" & Trim(Left(Combo1, 10)) & "'"
                s = s & ", palnum = '" & tno & "'"
                s = s & ", opseq = " & Val(Grid1.TextMatrix(i, 1))
                s = s & ", sku = '" & Grid1.TextMatrix(i, 2) & "'"
                s = s & ", lotnum = '...'"
                s = s & ", qty = " & Val(Grid1.TextMatrix(i, 4))
                s = s & ", uom = 'Wraps'"
                s = s & ", units = " & Format(k * Val(Grid1.TextMatrix(i, 4)), "0")
                s = s & ", palletid = '" & palid & "'"
                s = s & ", status = 'PEND'"
                s = s & ", userid = '.'"
                s = s & ", location = 'ORDER PICK'"
                s = s & ", reqid = '.'"
                s = s & " Where id = " & ds!id
                Wdb.Execute s
            Else
                zid = wd_seq("PickTasks", Form1.bbsr)
                s = "INSERT INTO PickTasks (ID, Branch, BrName, ShipDate, PalNum, OPSeq,"
                s = s & " SKU, LotNum, Qty, Uom, Units, PalletID, Status, UserID, Location,"
                s = s & " ReqID) VALUES (" & zid & ","
                s = s & bno & ","
                s = s & "'" & Combo2 & "',"
                s = s & "'" & Trim(Left(Combo1, 10)) & "',"
                s = s & tno & ","
                s = s & Val(Grid1.TextMatrix(i, 1)) & ","
                s = s & "'" & Grid1.TextMatrix(i, 2) & "',"
                s = s & "'...',"
                s = s & Val(Grid1.TextMatrix(i, 4)) & ","
                s = s & "'Wraps',"
                s = s & Format(k * Val(Grid1.TextMatrix(i, 4)), "0") & ","
                s = s & "'" & palid & "',"
                s = s & "'PEND',"
                s = s & "'.',"
                s = s & "'ORDER PICK',"
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

Private Sub view_prtall(pd As Control, tno As String)
    Dim k As Integer, i As Integer, tstat As String
    Dim p As Integer, rstr As String, pflag As Boolean
    Dim halfwidth, sy As Long, s As String
    Dim ds As adodb.Recordset
    Dim f0 As String, f1 As String, f2 As String, f3 As String, f4 As String
    Dim f5 As String, f6 As String, f7 As String, f8 As String, f9 As String
    Dim f10 As String, f11 As String, f12 As String, f13 As String, f14 As String
    Dim f15 As String, f16 As String, cfile As String
    Dim bno As Integer, palid As String, bc As String, rc As Integer
    On Error GoTo vberror
    Screen.MousePointer = 11
    If Left(List1, 1) = "T" Then
        bno = 16
        palid = Format(bno, "00")               'branch
        palid = palid & Right(List1, 6)         'account
        palid = palid & Format(Val(tno), "00")  'palnum
        palid = palid & Left(Combo1, 2) & mid(Combo1, 4, 2) & mid(Combo1, 9, 2)
        bc = "!" & palid & "!"
    Else
        bno = Val(Left(List1, 2))
        palid = Format(bno, "000") & " "
        palid = palid & Left(Combo1, 2) & mid(Combo1, 4, 2) & mid(Combo1, 9, 2) & " B "
        palid = palid & Format(Val(tno), "000")
        bc = "!" & Format(bno, "000") & "="
        bc = bc & Left(Combo1, 2) & mid(Combo1, 4, 2) & mid(Combo1, 9, 2) & "=B="
        bc = bc & Format(Val(tno), "000") & "!"
    End If
    
    
    pd.FontName = "Arial"
    pd.FontSize = 8
    pd.FontUnderline = False
    pd.CurrentX = 0: pd.CurrentY = 0
    s = Combo2 & " Pallet #" & tno & " " & Left(Combo1, 10)
    s = s & "                PalletCode: " & palid & "     "
    pd.Print s;
    pd.FontName = "IDAutomationHC39M"
    pd.Print bc
    pd.FontName = "Arial"
    pd.FontSize = 48
    pd.FontBold = True
    pd.CurrentX = 0: pd.CurrentY = 0
 
    'pd.FontSize = 48
    pd.FontUnderline = True
    pd.CurrentY = 1440 * 1.5
    s = Combo2
        
    If pd.TextWidth(s) > pd.ScaleWidth Then
        pd.FontSize = 24
        For i = Len(s) To 5 Step -1
            s = Left(s, i)
            If pd.TextWidth(s) <= pd.ScaleWidth Then Exit For
        Next i
    End If
        
    halfwidth = pd.TextWidth(s) / 2
    pd.CurrentX = pd.ScaleWidth / 2 - halfwidth
    pd.Print s
    pd.CurrentY = 1440 * 2.5
    
    pd.FontBold = True
    pd.FontUnderline = True
    pd.FontSize = 18
    pd.CurrentX = 1: pd.Print "  SKU  ";
    pd.CurrentX = 1440 * 3.5: pd.Print "Description";
    pd.CurrentX = 1440 * 7.2: pd.Print "Wraps"
    
    
    rc = 0
    cfile = Form1.pallogs & "pick" & Format(Now, "mmddyyyy") & ".txt"
    If Len(Dir(cfile)) > 0 Then
        Open cfile For Input As #1
        Do Until EOF(1)
            Input #1, f0, f1, f2, f3, f4, f5, f6, f7, f8, f9, f10, f11, f12, f13, f14, f15, f16
        
            If f6 = palid Then
                pd.FontSize = 36
                pd.FontUnderline = False
                pd.FontBold = False
                pd.CurrentX = 1
                pd.Print StrConv(f5, vbProperCase);
                s = "  " & f7
                pd.CurrentX = (1440 * 8) - pd.TextWidth(s)
                pd.Print s
                pd.FontSize = 16
                pd.FontBold = False
                pd.CurrentX = 1
                s = Format(DateAdd("yyyy", 2, calc_date(f9)), "mmddyy")
                If f11 > "0" Then s = s & ", " & Format(DateAdd("yyyy", 2, calc_date(f11)), "mmddyy")
                pd.Print "Code Date(s):  " & s
                pd.Print " "
                rc = rc + 1
            End If
        Loop
        Close #1
    End If
    
    If rc = 0 Then
        s = "select picktasks.sku,opseq,uom_type,description,lotnum,qty"
        s = s & " from picktasks,sku_config where picktasks.palletid = '" & palid & "'"
        s = s & " and sku_config.sku = picktasks.sku"
        s = s & " order by 2, 1"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                pd.FontSize = 36
                pd.FontUnderline = False
                pd.FontBold = False
                pd.CurrentX = 1
                pd.Print ds(0) & " ";
                pd.Print StrConv(ds(2) & " " & ds(3), vbProperCase);
                s = "  " & ds(5)
                pd.CurrentX = (1440 * 8) - pd.TextWidth(s)
                pd.Print s
                pd.FontSize = 16
                pd.FontBold = False
                pd.CurrentX = 1
                pd.Print "Code Date(s):  " & ds(4)
                pd.Print " "
                rc = rc + 1
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
        
    If rc = 0 Then
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 0) = tno Then
                pd.FontSize = 36
                pd.FontUnderline = False
                pd.FontBold = False
                pd.CurrentX = 1
                pd.Print Grid1.TextMatrix(i, 2) & " ";
                pd.Print StrConv(Grid1.TextMatrix(i, 3), vbProperCase);
                s = "  " & Grid1.TextMatrix(i, 4)
                pd.CurrentX = (1440 * 8) - pd.TextWidth(s)
                pd.Print s
                pd.FontSize = 16
                pd.FontBold = False
                pd.CurrentX = 1
                pd.Print "Code Date(s):"
                pd.Print " "
            End If
        Next i
    End If
    
    pd.FontName = "IDAutomationHC39M"
    pd.FontSize = 8
    pd.FontBold = False
    pd.Print bc;
    pd.FontName = "Arial"
    pd.FontSize = 8
    pd.FontBold = False
    pd.FontUnderline = False
    pd.Print "           " & palid
    pd.EndDoc
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "view_prtall", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " view_prtall - Error Number: " & eno
        End
    End If
End Sub


Private Sub refresh_grid()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 6
    If Len(List1) > 0 And read_sae = False Then
        If Left(List1, 1) = "T" Then
            s = "Select ID,trailers.sku,fgunit,fgdesc,pallets,wraps,units,pallet,numwrap,branch,account,plant,skumast.whs_num"
            s = s & " from trailers,skumast"
            s = s & " Where runid = " & mid$(List1, 2, 6)
            s = s & " And trailers.sku = skumast.sku"
            s = s & " and wraps > 0"
            s = s & " Order by trailers.sku"
            Set ds = Sdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                Do Until ds.EOF
                    s = ".." & Chr(9)
                    s = s & "999" & Chr(9)
                    s = s & ds!sku & Chr(9)
                    s = s & ds!fgunit & " " & ds!fgdesc & Chr(9)
                    s = s & ds!wraps '& Chr(9)
                    Grid1.AddItem s
                    ds.MoveNext
                Loop
            End If
        Else
            s = "Select ID,brorders.sku,fgunit,fgdesc,partqty,altflag"
            s = s & " from brorders,skumast"
            s = s & " where plant = " & Form1.plantno
            s = s & " and branch = " & Val(Left(List1, 2))
            s = s & " and account = '" & Right(List1, 6) & "'"
            s = s & " and orddate = '" & Left(Combo1, 10) & "'"
            s = s & " And brorders.sku = skumast.sku"
            s = s & " and partqty > 0"
            s = s & " Order by brorders.sku"
        
            Set ds = Sdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                Do Until ds.EOF
                    s = ".." & Chr(9)
                    s = s & "999" & Chr(9)
                    s = s & ds!sku & Chr(9)
                    s = s & ds!fgunit & " " & ds!fgdesc & Chr(9)
                    s = s & ds!partqty & Chr(9)
                    'If ds!altflag = True Then s = s & "Yes"
                    If ds!altflag = "Y" Then s = s & "Yes"
                    Grid1.AddItem s
                    ds.MoveNext
                Loop
            End If
        End If
        ds.Close
        
        
        If Grid1.Rows > 1 Then
            For i = 1 To Grid1.Rows - 1
                s = "select opseq from oplist where sku = '" & Grid1.TextMatrix(i, 2) & "'"
                Set ds = Sdb.Execute(s)
                If ds.BOF = False Then
                    ds.MoveFirst
                    Grid1.TextMatrix(i, 1) = Format(ds!opseq, "0000")
                End If
                ds.Close
            Next i
        End If
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 1: Grid1.ColSel = 1
        Grid1.Sort = 5
        s = "^Tag #|^OP|^SKU|<Product|^Wraps|^Alt"
        Grid1.FormatString = s
        Grid1.ColWidth(0) = 800
        Grid1.ColWidth(1) = 600
        Grid1.ColWidth(2) = 800
        Grid1.ColWidth(3) = 4000
        Grid1.ColWidth(4) = 800
        Grid1.ColWidth(5) = 600
        For i = 0 To 34
            wc(i).Caption = " "
            Command2(i).Enabled = False
        Next i
    Else
        s = "^Tag #|^OP|^SKU|<Product|^Wraps|^Alt"
        Grid1.FormatString = s
        Grid1.ColWidth(0) = 800
        Grid1.ColWidth(1) = 600
        Grid1.ColWidth(2) = 800
        Grid1.ColWidth(3) = 4000
        Grid1.ColWidth(4) = 800
        Grid1.ColWidth(5) = 600
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

Private Sub clrtag_Click()
    Dim i As Integer, k As Integer, j As Integer
    If Val(Grid1.TextMatrix(Grid1.Row, 1)) = 0 Then Exit Sub
    k = Grid1.Row
    For i = 0 To 34
        wc(i).Caption = " "
        Command2(i).Enabled = False
    Next i
    Grid1.TextMatrix(Grid1.Row, 0) = ".."
    Grid1.FillStyle = flexFillRepeat
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 2
    Grid1.Sort = 3
    
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        If Val(Grid1.TextMatrix(i, 0)) = 0 Then
            Grid1.CellBackColor = Grid1.BackColor
        Else
            j = Val(Grid1.TextMatrix(i, 0)) - 1
            Grid1.CellBackColor = Command1(j).BackColor
            wc(j).Caption = Val(wc(j).Caption) + Val(Grid1.TextMatrix(i, 4))
            Command2(j).Enabled = True
        End If
    Next i
    Grid1.Row = k
    
End Sub

Private Sub Combo1_Click()
    Dim ds As adodb.Recordset, js As adodb.Recordset, s As String
    On Error GoTo vberror
    Grid1.Rows = 1: Combo2.Clear: List1.Clear
    'Jobbing Trailers
    s = "Select runid,trailers.branch,account,branchname,trlno,sum(wraps) from trailers,branches"
    s = s & " Where shipdate = '" & Left(Combo1, 10) & "'"
    If Form1.plantno = "50" Then s = s & " and trailers.branch in (15,16)"
    s = s & " And trailers.branch = branches.branch"
    s = s & " Group by runid,trailers.branch,account,branchname,trlno"
    s = s & " having sum(wraps) > 0"
    s = s & " order by branchname,trlno"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!account = "......" Then
                Combo2.AddItem ds!branchname & " " & ds!trlno
            Else
                s = "Select * from jobbing Where Branch = " & ds!branch & " And account = '" & ds!account & "'"
                Set js = Sdb.Execute(s)
                If js.BOF = False Then
                    js.MoveFirst
                    Combo2.AddItem js!acctdesc
                Else
                    Combo2.AddItem "......"
                End If
                js.Close
            End If
            List1.AddItem "T" & Format(ds!runid, "000000") & ds!account
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Form1.plantno <> "50" Then
        'db.Close
        If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
        Exit Sub
    End If
    
    
    'Branch Orders
    s = "select brorders.branch,account,branchname,sum(partqty) from brorders,branches"
    s = s & " Where orddate = '" & Left(Combo1, 10) & "'"
    s = s & " and plant = " & Form1.plantno
    s = s & " And brorders.branch = branches.branch"
    s = s & " Group by brorders.branch,account,branchname"
    s = s & " having sum(partqty) > 0"
    s = s & " order by branchname"
    
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!account = "......" Then
                Combo2.AddItem ds!branchname '& " " & ds!trlno
            Else
                s = "Select * from jobbing Where Branch = " & ds!branch & " And account = '" & ds!account & "'"
                Set js = Sdb.Execute(s)
                If js.BOF = False Then
                    js.MoveFirst
                    Combo2.AddItem js!acctdesc
                Else
                    Combo2.AddItem "......"
                End If
                js.Close
            End If
            List1.AddItem Format(ds(0), "00") & ds!account
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "combo1_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " combo1_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo2_Click()
    List1.ListIndex = Combo2.ListIndex
    refresh_grid
End Sub

Private Sub Command1_Click(Index As Integer)
    Dim i As Integer, k As Integer, j As Integer
    If Val(Grid1.TextMatrix(Grid1.Row, 1)) = 0 Then Exit Sub
    k = Grid1.Row
    For i = 0 To 34
        wc(i).Caption = " "
        Command2(i).Enabled = False
    Next i
    Grid1.TextMatrix(Grid1.Row, 0) = Index + 1
    Grid1.RowSel = Grid1.Row
    Grid1.Col = 0: Grid1.ColSel = 2
    Grid1.Sort = 3
    Grid1.FillStyle = flexFillRepeat
    For i = 1 To Grid1.Rows - 1
        Grid1.Row = i: Grid1.RowSel = i
        Grid1.Col = 0: Grid1.ColSel = Grid1.Cols - 1
        If Val(Grid1.TextMatrix(i, 0)) = 0 Then
            Grid1.CellBackColor = Grid1.BackColor
        Else
            j = CInt(Grid1.TextMatrix(i, 0)) - 1
            Grid1.CellBackColor = Command1(j).BackColor
            wc(j).Caption = Val(wc(j).Caption) + Val(Grid1.TextMatrix(i, 4))
            Command2(j).Enabled = True
        End If
    Next i
    Grid1.Row = k
End Sub

Private Sub Command2_Click(Index As Integer)
    Dim s As String
    If Option1.Value = True Then Call post_sae(Index + 1)
    'Exit Sub
    
    If Option2.Value = True Then
        Printer.PaperSize = 5
        s = Index + 1
        Call view_prtall(Printer, s)
        Printer.EndDoc
        Printer.PaperSize = 1
        'labelpic.Show
        'labelpic.labpt.Caption = Index + 1
        'labelpic.ptrig.Caption = Val(labelpic.ptrig.Caption) + 1
    End If
End Sub

Private Sub Command3_Click()
    Dim s As String
    'Printer.PaperSize = 1
    'Call view_prtlist(Printer)
    'Printer.EndDoc
    'Printer.PaperSize = 1
    labelpic.Show
    labelpic.ptrig.Caption = Val(labelpic.ptrig.Caption) + 1
End Sub

Private Sub delrec_Click()
    Dim s As String, i As Integer, k As Integer
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) <> 0 Then
        MsgBox "You must clear the tag # before this line can be erased,", vbOKOnly + vbInformation, "try again.."
        Exit Sub
    End If
    s = Grid1.TextMatrix(Grid1.Row, 2)
    If Val(s) = 0 Then Exit Sub
    k = 0
    For i = 1 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 2) = s Then k = k + 1
    Next i
    If k = 1 Then
        If MsgBox("This is the only line for SKU: " & s, vbYesNo + vbInformation, "are you sure..") = vbNo Then Exit Sub
    End If
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Grid1.Rows = 1
    End If
End Sub

Private Sub edqty_Click()
    Dim s As String, q As String
    If Val(Grid1.TextMatrix(Grid1.Row, 2)) = 0 Then Exit Sub
    s = Grid1.TextMatrix(Grid1.Row, 3)
    q = Grid1.TextMatrix(Grid1.Row, 4)
    q = InputBox("Wrap Qty", s, q)
    If Val(q) = 0 Then Exit Sub
    Grid1.TextMatrix(Grid1.Row, 4) = q
End Sub

Private Sub Form_Load()
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    s = "select distinct shipdate from trailers where wraps <> 0 and branch in (15,16) order by shipdate"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem Format(ds(0), "mm-dd-yyyy") & " " & Format(ds(0), "dddd")
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    s = "select distinct orddate from brorders where partqty <> 0"
    s = s & " and orddate not in (select shipdate from trailers where wraps <> 0 and branch in (15,16))"
    s = s & " order by orddate"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem Format(ds(0), "mm-dd-yyyy") & " " & Format(ds(0), "ddd")
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Combo1.ListCount > 0 Then
        For i = 0 To Combo1.ListCount - 1
            If Left(Combo1.List(i), 10) > Format(Now, "mm-dd-yyyy") Then
                Combo1.ListIndex = i
                Exit For
            End If
        Next i
        If Combo1.ListIndex < 0 Then Combo1.ListIndex = 0
    End If
    Option1_Click
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "form_Load", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " form_load - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Resize()
    If Me.Height > 2000 Then Grid1.Height = Me.Height - 1400
    pgrid.Width = Me.Width - 80
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If UCase(Right(Form1.Caption, 11)) = "PARTPALLETS" Then End
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub insline_Click()
    Dim i As Integer, s As String
    If Val(Grid1.TextMatrix(Grid1.Row, 1)) = 0 Then Exit Sub
    If Val(Grid1.TextMatrix(Grid1.Row, 0)) > 0 Then Exit Sub
    i = Val(Grid1.TextMatrix(Grid1.Row, 4))
    s = InputBox("Qty for new line.", "Insert line...", "1")
    If Val(s) = 0 Then Exit Sub
    Grid1.TextMatrix(Grid1.Row, 4) = i - Val(s)
    
    i = Val(s)
    s = ".." & Chr(9)
    s = s & Grid1.TextMatrix(Grid1.Row, 1) & Chr(9)
    s = s & Grid1.TextMatrix(Grid1.Row, 2) & Chr(9)
    s = s & Grid1.TextMatrix(Grid1.Row, 3) & Chr(9)
    s = s & i 'Grid1.TextMatrix(Grid1.Row, 4)
    Grid1.AddItem s, Grid1.Row
    Grid1.Row = Grid1.Row + 1
End Sub

Private Sub Option1_Click()
    Dim i As Integer
    If Option1.Value = True Then
        For i = 0 To 34
            Command2(i).Caption = "Post HH"
        Next i
    Else
        For i = 0 To 34
            Command2(i).Caption = "Print"
        Next i
    End If
End Sub

Private Sub Option2_Click()
    Dim i As Integer
    If Option2.Value = True Then
        For i = 0 To 34
            Command2(i).Caption = "Print"
        Next i
    Else
        For i = 0 To 34
            Command2(i).Caption = "Post HH"
        Next i
    End If
End Sub
