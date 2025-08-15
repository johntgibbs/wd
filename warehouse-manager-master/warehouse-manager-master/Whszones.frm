VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Zone Maintenance"
   ClientHeight    =   12795
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14400
   LinkTopic       =   "Form3"
   ScaleHeight     =   12795
   ScaleWidth      =   14400
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid opgrid 
      Height          =   3615
      Left            =   600
      TabIndex        =   65
      Top             =   600
      Visible         =   0   'False
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   6376
      _Version        =   327680
      ForeColor       =   16711680
      BackColorFixed  =   12648384
      AllowUserResizing=   3
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   495
      Left            =   12120
      TabIndex        =   36
      Top             =   6960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Print Grids"
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
      Left            =   9960
      TabIndex        =   35
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear Reservation"
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
      Left            =   6120
      TabIndex        =   30
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Insert Pallet"
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
      Left            =   6360
      TabIndex        =   11
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Position"
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
      Left            =   4440
      TabIndex        =   10
      Top             =   5640
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Lane"
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
      Left            =   8040
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reserve Bay"
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
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Out of Zone"
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
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid SGrid 
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   6000
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   2990
      _Version        =   327680
      BackColor       =   16777152
      BackColorFixed  =   14737632
      BackColorBkg    =   8421376
      FocusRect       =   0
      Appearance      =   0
   End
   Begin VB.ComboBox Whs 
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
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid RGrid 
      Height          =   2415
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4260
      _Version        =   327680
      Rows            =   8
      FocusRect       =   0
      HighLight       =   2
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid LGrid 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4260
      _Version        =   327680
      Rows            =   8
      FocusRect       =   0
      HighLight       =   2
      Appearance      =   0
   End
   Begin VB.Label dot 
      Caption         =   "Label6"
      Height          =   255
      Index           =   27
      Left            =   11760
      TabIndex        =   64
      Top             =   11160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      Caption         =   "Label6"
      Height          =   255
      Index           =   26
      Left            =   11760
      TabIndex        =   63
      Top             =   10920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      Caption         =   "Label6"
      Height          =   255
      Index           =   25
      Left            =   11760
      TabIndex        =   62
      Top             =   10680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      Caption         =   "Label6"
      Height          =   255
      Index           =   24
      Left            =   11760
      TabIndex        =   61
      Top             =   10440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      Caption         =   "Label6"
      Height          =   255
      Index           =   23
      Left            =   11760
      TabIndex        =   60
      Top             =   10200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      Caption         =   "Label6"
      Height          =   255
      Index           =   22
      Left            =   11760
      TabIndex        =   59
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      Caption         =   "Label6"
      Height          =   255
      Index           =   21
      Left            =   11760
      TabIndex        =   58
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      Caption         =   "Label6"
      Height          =   255
      Index           =   20
      Left            =   11760
      TabIndex        =   57
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      Caption         =   "Label6"
      Height          =   255
      Index           =   19
      Left            =   11760
      TabIndex        =   56
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      Caption         =   "Label6"
      Height          =   255
      Index           =   18
      Left            =   11760
      TabIndex        =   55
      Top             =   9000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H000000C0&
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   11760
      TabIndex        =   54
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00FF00FF&
      Caption         =   "Label6"
      Height          =   255
      Index           =   16
      Left            =   11760
      TabIndex        =   53
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00008000&
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   11760
      TabIndex        =   52
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H000080FF&
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   11760
      TabIndex        =   51
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label6"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   13
      Left            =   10440
      TabIndex        =   50
      Top             =   11160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00000080&
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   10440
      TabIndex        =   49
      Top             =   10920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00C0C000&
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   10440
      TabIndex        =   48
      Top             =   10680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label6"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   10440
      TabIndex        =   47
      Top             =   10440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00FF0000&
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   10440
      TabIndex        =   46
      Top             =   10200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00FFFF00&
      Caption         =   "Label6"
      Height          =   255
      Index           =   8
      Left            =   10440
      TabIndex        =   45
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H0000FF00&
      Caption         =   "Label6"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   10440
      TabIndex        =   44
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label6"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   10440
      TabIndex        =   43
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00C000C0&
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   10440
      TabIndex        =   42
      Top             =   9240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label6"
      Height          =   255
      Index           =   4
      Left            =   10440
      TabIndex        =   41
      Top             =   9000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H000000FF&
      Caption         =   "Label6"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   10440
      TabIndex        =   40
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Label6"
      Height          =   255
      Index           =   2
      Left            =   10440
      TabIndex        =   39
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Label6"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   10440
      TabIndex        =   38
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label dot 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label6"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   10440
      TabIndex        =   37
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   13680
      TabIndex        =   34
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11040
      TabIndex        =   33
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   9960
      TabIndex        =   32
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label2 
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
      Left            =   8400
      TabIndex        =   31
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   4080
      TabIndex        =   29
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   3840
      TabIndex        =   28
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   3600
      TabIndex        =   27
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   3360
      TabIndex        =   26
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   3120
      TabIndex        =   25
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   2880
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   2640
      TabIndex        =   23
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   20
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1680
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   18
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   17
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   16
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label ckey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   5640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label pkey 
      Caption         =   "0"
      Height          =   375
      Left            =   8880
      TabIndex        =   8
      Top             =   8880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label zone 
      Caption         =   "0"
      Height          =   255
      Left            =   8880
      TabIndex        =   5
      Top             =   8520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SR-"
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu edlane 
      Caption         =   "Edit Lane"
      Visible         =   0   'False
      Begin VB.Menu edlane1 
         Caption         =   "Reserve Bay"
      End
      Begin VB.Menu edlane2 
         Caption         =   "Clear Reservation"
      End
      Begin VB.Menu edlane3 
         Caption         =   "Clear Lane"
      End
   End
   Begin VB.Menu edpos 
      Caption         =   "Edit Position"
      Visible         =   0   'False
      Begin VB.Menu edpos1 
         Caption         =   "Clear Position"
      End
      Begin VB.Menu edpos2 
         Caption         =   "Insert Pallet"
      End
      Begin VB.Menu batonhand 
         Caption         =   "View Batch Inventory"
      End
      Begin VB.Menu vphist 
         Caption         =   "View Pallet History"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fc(100) As Long
Dim bc(100) As Long
Dim Col(24) As Long
Dim fgc(24) As Long
Dim lkey(1, 7, 58) As Long
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

Sub print_grid(gname As Control, r1 As Integer, r2 As Integer, rtitle As String)
    Dim i As Integer, k As Integer, j As Integer
    Dim xs As Long, xe As Long, xm As Long
    Dim ys As Long, ye As Long
    Dim cw(0 To 60) As Long
    For i = 0 To gname.Cols - 1
        cw(i) = gname.ColWidth(i) - 50
    Next i
    'Override Grid Col Widths
    'cw(0) = 600
    'cw(1) = 1700
    'cw(2) = 1700
    'cw(3) = 1700
    'cw(4) = 1700
    'cw(5) = 1700
    'cw(VGrid.Cols - 1) = 1000
    xs = 0: xe = xs
    For i = 0 To gname.Cols - 1
        If cw(i) > 10 Then xe = xe + cw(i)
    Next i
    If xe > 11600 Then
        Printer.Orientation = 2
    Else
        Printer.Orientation = 1
    End If
    
    Printer.FontTransparent = True
    Printer.FillStyle = 0
    Printer.FillColor = QBColor(15)
    Printer.DrawMode = 1
    Printer.ForeColor = QBColor(0)
    
    Printer.FontName = "MS Serif"
    Printer.FontTransparent = True
    Printer.FontSize = 14
    Printer.DrawWidth = 6
    Printer.Print rtitle
    Printer.Print Format(Now, "mmmm d, yyyy")

    Printer.FontSize = 8
    Printer.Line (xs, 1200)-(xe, 1200)
    Printer.Line (xs, 1440)-(xe, 1440)
    Printer.FillColor = QBColor(15)
    Printer.DrawWidth = 3
    j = 0
    For i = r1 To r2 + 1
        ye = j * 240 + 1440
        Printer.Line (xs, ye)-(xe, ye)
        j = j + 1
    Next i
    'Printer.DrawWidth = 1
    Printer.FontBold = False
    xm = xs + 50
    For k = 0 To gname.Cols - 1
        If cw(k) > 10 Then
            Printer.PSet (xm, 1230)
            Printer.Print gname.TextMatrix(0, k)
            xm = xm + cw(k)
        End If
    Next k
    j = 1
    For i = r1 To r2
        xm = xs + 100
        For k = 0 To gname.Cols - 1
            If cw(k) > 10 Then
                Printer.PSet (xm, j * 240 + 1230)
                Printer.Print gname.TextMatrix(i, k)
                xm = xm + cw(k)
            End If
        Next k
        j = j + 1
    Next i
    ys = 1200
    xm = xs
    Printer.DrawWidth = 6
    For i = 0 To gname.Cols - 1
        If cw(i) > 10 Then
            Printer.Line (xm, ys)-(xm, ye)
            xm = xm + cw(i)
        End If
    Next i
    Printer.Line (xm, ys)-(xm, ye)
' Right Grid
    Printer.Line (xs, 6000)-(xe, 6000)
    Printer.Line (xs, 6240)-(xe, 6240)
    Printer.FillColor = QBColor(15)
    Printer.DrawWidth = 3
    j = 0
    For i = r1 To r2 + 1
        ye = j * 240 + 6240
        Printer.Line (xs, ye)-(xe, ye)
        j = j + 1
    Next i
    'Printer.DrawWidth = 1
    Printer.FontBold = False
    xm = xs + 50
    For k = 0 To gname.Cols - 1
        If cw(k) > 10 Then
            Printer.PSet (xm, 6030)
            Printer.Print RGrid.TextMatrix(0, k)
            xm = xm + cw(k)
        End If
    Next k
    j = 1
    For i = r1 To r2
        xm = xs + 100
        For k = 0 To gname.Cols - 1
            If cw(k) > 10 Then
                Printer.PSet (xm, j * 240 + 6030)
                Printer.Print RGrid.TextMatrix(i, k)
                xm = xm + cw(k)
            End If
        Next k
        j = j + 1
    Next i
    ys = 6000
    xm = xs
    Printer.DrawWidth = 6
    For i = 0 To gname.Cols - 1
        If cw(i) > 10 Then
            Printer.Line (xm, ys)-(xm, ye)
            xm = xm + cw(i)
        End If
    Next i
    Printer.Line (xm, ys)-(xm, ye)
    
    Printer.Print " "
    Printer.Print "Printed: " & Format(Now, "m-d-yyyy h:mm am/pm")
    Printer.EndDoc
End Sub

Private Sub refresh_opgrid(psku As String)                      'jv110716
    Dim ds As adodb.Recordset, s As String
    opgrid.Redraw = False
    opgrid.FontName = "Arial"
    opgrid.FontBold = True
    opgrid.Clear: opgrid.Rows = 1: opgrid.Cols = 13
    s = "select * from rackpos where sku = '" & psku & "'"
    s = s & " and rackno in (select id from racks where aisle = 'M' and rack = 'OP')"
    'MsgBox s
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = ds!id & Chr(9) & ds!posn_num & Chr(9) & ds!sku & Chr(9) & ds!lot_num & Chr(9)
            s = s & ds!pallet_num & Chr(9) & ds!count_qty & Chr(9) & ds!lot2 & Chr(9)
            s = s & ds!qty2 & Chr(9) & Format(ds!recv_date, "M-dd-yyyy") & Chr(9) & ds!bbc & Chr(9)
            s = s & ds!barcode & Chr(9) & ds!wrapped & Chr(9) & ds!hold
            opgrid.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "^Id|^Posn|^SKU|^Lot|^Pallet|^Qty|^Lot2|^Qty2|^Date|^BBC|^BarCode|^Wrapped|^Hold"
    opgrid.FormatString = s
    opgrid.ColWidth(0) = 900
    opgrid.ColWidth(1) = 900
    opgrid.ColWidth(2) = 900
    opgrid.ColWidth(3) = 900
    opgrid.ColWidth(4) = 900
    opgrid.ColWidth(5) = 900
    opgrid.ColWidth(6) = 900
    opgrid.ColWidth(7) = 900
    opgrid.ColWidth(8) = 1200
    opgrid.ColWidth(9) = 900
    opgrid.ColWidth(10) = 1900
    opgrid.ColWidth(11) = 900
    opgrid.ColWidth(12) = 900
    opgrid.Redraw = True
End Sub

Private Sub Refresh_zones()
    Dim ds As adodb.Recordset, sqlx As String, ds2 As adodb.Recordset
    Dim ndate As Long, nyear As String
    nyear = Year(Now)
    nyear = Val(nyear) - 1
    ndate = DateDiff("d", "12-31-" & nyear, Now)
    nyear = Right(Year(Now), 2) & Format(ndate, "000")
    Screen.MousePointer = 11
    SGrid.Visible = False
    sqlx = "select * from lane where whse_num = " & Whs
    sqlx = sqlx & " order by vert_loc, rack_side"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!rack_side = "R" Then
                RGrid.Row = 8 - ds!vert_loc
                RGrid.Col = ds!horz_loc
                If ds!lane_status = "B" Then
                    RGrid.Text = "B"
                Else
                    If ds!lane_status = "H" Then
                        RGrid.Text = "H"
                    Else
                        If ds!qty > 0 Then
                            If ds!qty = ds!capacity Then
                                RGrid.Text = "X"
                            Else
                                RGrid.Text = ds!qty
                            End If
                        Else
                            If Val(ds!resv_sku) > 0 Then RGrid.Text = "R"
                        End If
                    End If
                End If
                If Check1.Value = 1 Then RGrid.Text = ds!zone_num
                lkey(1, RGrid.Row, RGrid.Col) = ds!id
                If ds!qty > 0 Then RGrid.CellFontUnderline = True
                If Check1.Value = 1 And ds!qty > 0 Then
                    Set ds2 = Wdb.Execute("select * from zone_config where sku = '" & ds!sku & "'")
                    If ds2.BOF = False Then
                        ds2.MoveFirst
                        RGrid.CellBackColor = bc(ds2!zone_num)
                        RGrid.CellForeColor = fc(ds2!zone_num)
                    End If
                    ds2.Close
                Else
                    RGrid.CellBackColor = bc(ds!zone_num)
                    RGrid.CellForeColor = fc(ds!zone_num)
                End If
            End If
            If ds!rack_side = "L" Then
                LGrid.Row = 8 - ds!vert_loc
                LGrid.Col = ds!horz_loc
                If ds!lane_status = "B" Then
                    LGrid.Text = "B"
                Else
                    If ds!lane_status = "H" Then
                        LGrid.Text = "H"
                    Else
                        If ds!qty > 0 Then
                            If ds!qty = ds!capacity Then
                                LGrid.Text = "X"
                            Else
                                'RGrid.Text = "p"
                                LGrid.Text = ds!qty
                            End If
                            'LGrid.Text = "X"
                        Else
                            If Val(ds!resv_sku) > 0 Then LGrid.Text = "R"
                        End If
                    End If
                End If
                If Check1.Value = 1 Then LGrid.Text = ds!zone_num
                lkey(0, LGrid.Row, LGrid.Col) = ds!id
                If ds!qty > 0 Then LGrid.CellFontUnderline = True
                If Check1.Value = 1 And ds!qty > 0 Then
                    Set ds2 = Wdb.Execute("select * from zone_config where sku = '" & ds!sku & "'")
                    If ds2.BOF = False Then
                        ds2.MoveFirst
                        LGrid.CellBackColor = bc(ds2!zone_num)
                        LGrid.CellForeColor = fc(ds2!zone_num)
                    End If
                    ds2.Close
                Else
                    LGrid.CellBackColor = bc(ds!zone_num)
                    LGrid.CellForeColor = fc(ds!zone_num)
                End If
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    SGrid.Visible = True
    Call pkey_Change
    Screen.MousePointer = 0
End Sub

Private Sub batonhand_Click()
    Dim s As String
    s = Left(SGrid.TextMatrix(SGrid.Row, 13), 13)
    tktonhand.bbarcode = s
    tktonhand.bproduct = SGrid.TextMatrix(SGrid.Row, 4)
    tktonhand.Show
End Sub

Private Sub Check1_Click()
    Call Whs_Click
End Sub

Private Sub Command1_Click()                    'Reserve Bay
    Dim psku As String, plot As String
    Dim ds As adodb.Recordset, sqlx As String
    If Label5 = "LIFO Bay" Then
        MsgBox "Cannot Reserve Bay configured as LIFO.", vbOKOnly, "Sorry, Pick again..."
        Exit Sub
    End If
    If Label5 = "Blocked Bay" Then
        MsgBox "Cannot Reserve Blocked Bay.", vbOKOnly, "Sorry, Pick again..."
        Exit Sub
    End If
    If Label5 = "Product On Hold" Then
        MsgBox "Cannot Reserve Bay with Product On Hold.", vbOKOnly, "Sorry, Pick again..."
        Exit Sub
    End If
    sqlx = "select * from queue_infc where whse_num = " & Whs
    sqlx = sqlx & " and queue_num > 0 order by queue_num"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        psku = ds!sku
        plot = ds!lot_num
    Else
        psku = "000": plot = "00000"
    End If
    ds.Close
    psku = InputBox("SKU #:", "Reserve Lane", psku)
    If Len(psku) = 0 Then Exit Sub
    plot = InputBox("Lot #:", "Reserve Lane", plot)
    If Len(plot) = 0 Then Exit Sub
    If skurec(Val(psku)).sku <> psku Then
        MsgBox "Invalid SKU", vbOKOnly, "Sorry Cannot Reserve"
        Exit Sub
    End If
    sqlx = "Update lane set resv_sku = '" & psku & "'"
    sqlx = sqlx & ", resv_lot = '" & plot & "'"
    sqlx = sqlx & " Where id = " & pkey.Caption
    Wdb.Execute sqlx
    
    sqlx = "Update position set posn_status = 'R'"
    sqlx = sqlx & " Where laneno = " & pkey.Caption
    Wdb.Execute sqlx
    
    Call pkey_Change
    'If Right(SGrid.TextMatrix(0, 0), 1) = "L" Then
    If Right(Label2, 1) = "L" Then
        LGrid.Text = "R"
    Else
        RGrid.Text = "R"
    End If
End Sub

Private Sub Command2_Click()
    Dim sqlx As String, i As Integer
    Dim p As ptask, preas As String                                                 'jv060117
    If MsgBox("Ok to clear lane " & SGrid.TextMatrix(0, 0) & "?", vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    sqlx = "Update lane set qty = 0, sku = ' ', lot_num = ' ', resv_sku = ' ', resv_lot = ' '"
    sqlx = sqlx & ", gmasize = 0, horz_travel = 0 Where id = " & pkey.Caption
    Wdb.Execute sqlx
    
    'If Right$(SGrid.TextMatrix(0, 0), 1) = "R" Then    'Brenham
    'If Right$(SGrid.TextMatrix(0, 0), 1) = "2" Then     'Sylacauga
    If Right(Label2, 1) = "R" Then
        RGrid.Text = ""
        RGrid.CellFontUnderline = False
    Else
        LGrid.Text = ""
        LGrid.CellFontUnderline = False
    End If
    
    preas = InputBox("Reason for delete:", "Reason for delete....")                         'jv060117
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"                         'jv060117
    Open cfile For Append Shared As #1                                                      'jv060117
    For i = 1 To SGrid.Rows - 1                                                             'jv060117
        sqlx = "Update position set posn_status = ' ', sku = ' ', lot_num = ' ', pallet_num = 0"
        sqlx = sqlx & ", lot_status = ' ', pallet_status = ' ', count_qty = 0"
        sqlx = sqlx & ", recv_date = '" & Format(Now, "m-d-yyyy") & "', barcode = ' '"
        sqlx = sqlx & ", lot2 = ' ', qty2 = 0 Where id = " & SGrid.TextMatrix(i, 0)
        Wdb.Execute sqlx                                                                    'jv060117
        If Val(SGrid.TextMatrix(i, 3)) <> 0 Then                                            'jv060117
            p.area = "SR-" & Whs                                                            'jv060117
            If Len(preas) > 0 Then                                                          'jv060117
                p.description = preas                                                       'jv060117
            Else                                                                            'jv060117
                p.description = " "                                                         'jv060117
            End If                                                                          'jv060117
            p.source = "Clear Lane"                                                         'jv060117
            p.target = Me.Label2 & " " & Trim(SGrid.TextMatrix(i, 1))                       'jv060117
            p.product = SGrid.TextMatrix(i, 3) & " " & UCase(SGrid.TextMatrix(i, 4))        'jv060117
            p.palletid = SGrid.TextMatrix(i, 13)                                            'jv060117
            p.qty = "1"                                                                     'jv060117
            p.uom = "Pallet"                                                                'jv060117
            p.lotnum = SGrid.TextMatrix(i, 5)                                               'jv060117
            p.units = SGrid.TextMatrix(i, 9)                                                'jv060117
            p.lotnum2 = SGrid.TextMatrix(i, 10)                                             'jv060117
            p.units2 = SGrid.TextMatrix(i, 11)                                              'jv060117
            p.status = "COMP"                                                               'jv060117
            p.userid = Form1.userid                                                         'jv060117
            p.trandate = Format(Now, "yyMMdd hh:mm:ss")                                     'jv060117
            p.reqid = ".."                                                                  'jv060117
            'If LCase(Form1.userid) <> "jvierus" Then                                        'jv060117
                Write #1, i;                                                                'jv060117
                Write #1, p.area;                                                           'jv060117
                Write #1, p.description;                                                    'jv060117
                Write #1, p.source;                                                         'jv060117
                Write #1, p.target;                                                         'jv060117
                Write #1, p.product;                                                        'jv060117
                Write #1, p.palletid;                                                       'jv060117
                Write #1, p.qty;                                                            'jv060117
                Write #1, p.uom;                                                            'jv060117
                Write #1, p.lotnum;                                                         'jv060117
                Write #1, p.units;                                                          'jv060117
                Write #1, p.lotnum2;                                                        'jv060117
                Write #1, p.units2;                                                         'jv060117
                Write #1, p.status;                                                         'jv060117
                Write #1, p.userid;                                                         'jv060117
                Write #1, p.trandate;                                                       'jv060117
                Write #1, p.reqid                                                           'jv060117
            'End If                                                                          'jv060117
        End If                                                                              'jv060117
        SGrid.TextMatrix(i, 2) = " ": SGrid.TextMatrix(i, 3) = " "                          'jv060117
        SGrid.TextMatrix(i, 4) = " ": SGrid.TextMatrix(i, 5) = " "                          'jv060117
        SGrid.TextMatrix(i, 6) = "0": SGrid.TextMatrix(i, 7) = " "                          'jv060117
        SGrid.TextMatrix(i, 8) = " ": SGrid.TextMatrix(i, 9) = "0"                          'jv060117
        SGrid.TextMatrix(i, 10) = " ": SGrid.TextMatrix(i, 11) = " "                        'jv060117
        SGrid.TextMatrix(i, 12) = Format$(Now, "m-dd-yyyy")                                 'jv060117
        SGrid.TextMatrix(i, 13) = " "                                                       'jv060117
    Next i                                                                                  'jv060117
    Close #1                                                                                'jv060117
    
    'SR Log
    'If Form1.plantno = "50" Then
    '    'Add to crane movement log
    '    cfile = Form1.srserv & "\wd\sr" & Whs.Text & "\bin\SR" & Whs.Text & Format(Now, "mmdd") & ".csv"
    '    'cfile = "c:\sr10430.csv"
    '    Open cfile For Append As #1
    'End If
    
    'For i = 1 To SGrid.Rows - 1
    '    sqlx = "Update position set posn_status = ' ', sku = ' ', lot_num = ' ', pallet_num = 0"
    '    sqlx = sqlx & ", lot_status = ' ', pallet_status = ' ', count_qty = 0"
    '    sqlx = sqlx & ", recv_date = '" & Format(Now, "m-d-yyyy") & "', barcode = ' '"
    '    sqlx = sqlx & ", lot2 = ' ', qty2 = 0 Where id = " & SGrid.TextMatrix(i, 0)
    '    Wdb.Execute sqlx
    '
    '    If Form1.plantno = "50" And Val(SGrid.TextMatrix(i, 3)) <> 0 Then
    '        Write #1, "SR-" & Whs.Text;
    '        Write #1, "...";
    '        Write #1, SGrid.TextMatrix(i, 3);
    '        Write #1, SGrid.TextMatrix(i, 5);
    '        Write #1, SGrid.TextMatrix(i, 6);
    '        Write #1, LTrim(StrConv(SGrid.TextMatrix(i, 4), vbProperCase));
    '        Write #1, "WMS";
    '        Write #1, LTrim(Label2.Caption) & " " & SGrid.TextMatrix(i, 1);
    '        Write #1, "Cleared";
    '        Write #1, Format(Now, "h:mm am/pm")
    '    End If
    '
    '    SGrid.TextMatrix(i, 2) = " ": SGrid.TextMatrix(i, 3) = " "
    '    SGrid.TextMatrix(i, 4) = " ": SGrid.TextMatrix(i, 5) = " "
    '    SGrid.TextMatrix(i, 6) = "0": SGrid.TextMatrix(i, 7) = " "
    '    SGrid.TextMatrix(i, 8) = " ": SGrid.TextMatrix(i, 9) = "0"
    '    SGrid.TextMatrix(i, 10) = " ": SGrid.TextMatrix(i, 11) = " "
    '    SGrid.TextMatrix(i, 12) = Format$(Now, "m-dd-yyyy")
    '    SGrid.TextMatrix(i, 13) = " "
    'Next i
    'If Form1.plantno = "50" Then Close #1
    Call pkey_Change
End Sub

Private Sub Command3_Click()
    Dim y As Integer
    Dim sqlx As String, i As Integer, pqty As Integer
    Dim olot As String, cfile As String, poc As Integer
    Dim p As ptask, preas As String                                                 'jv060117
    If SGrid.Row < 1 Then Exit Sub
    y = SGrid.Row
    If MsgBox("Ok to clear position " & SGrid.Row & "?", vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    
    preas = InputBox("Reason for delete:", "Reason for delete....")                 'jv060117
    p.area = "SR-" & Whs                                                            'jv060117
    If Len(preas) > 0 Then                                                          'jv060117
        p.description = preas                                                       'jv060117
    Else                                                                            'jv060117
        p.description = " "                                                         'jv060117
    End If                                                                          'jv060117
    p.source = "Clear Position"                                                     'jv060117
    p.target = Me.Label2 & " " & Trim(SGrid.TextMatrix(y, 1))                       'jv060117
    p.product = SGrid.TextMatrix(y, 3) & " " & UCase(SGrid.TextMatrix(y, 4))        'jv060117
    p.palletid = SGrid.TextMatrix(y, 13)                                            'jv060117
    p.qty = "1"                                                                     'jv060117
    p.uom = "Pallet"                                                                'jv060117
    p.lotnum = SGrid.TextMatrix(y, 5)                                               'jv060117
    p.units = SGrid.TextMatrix(y, 9)                                                'jv060117
    p.lotnum2 = SGrid.TextMatrix(y, 10)                                             'jv060117
    p.units2 = SGrid.TextMatrix(y, 11)                                              'jv060117
    p.status = "COMP"                                                               'jv060117
    p.userid = Form1.userid                                                         'jv060117
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")                                     'jv060117
    p.reqid = ".."                                                                  'jv060117
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"                 'jv060117
    'If LCase(Form1.userid) <> "jvierus" Then                                        'jv060117
        Open cfile For Append Shared As #1                                          'jv060117
        Write #1, y;                                                                'jv060117
        Write #1, p.area;                                                           'jv060117
        Write #1, p.description;                                                    'jv060117
        Write #1, p.source;                                                         'jv060117
        Write #1, p.target;                                                         'jv060117
        Write #1, p.product;                                                        'jv060117
        Write #1, p.palletid;                                                       'jv060117
        Write #1, p.qty;                                                            'jv060117
        Write #1, p.uom;                                                            'jv060117
        Write #1, p.lotnum;                                                         'jv060117
        Write #1, p.units;                                                          'jv060117
        Write #1, p.lotnum2;                                                        'jv060117
        Write #1, p.units2;                                                         'jv060117
        Write #1, p.status;                                                         'jv060117
        Write #1, p.userid;                                                         'jv060117
        Write #1, p.trandate;                                                       'jv060117
        Write #1, p.reqid                                                           'jv060117
        Close #1                                                                    'jv060117
    'End If                                                                          'jv060117
    
    'SR Log
    'If Form1.plantno = "50" Then
    '    'Add to crane movement log
    '    'cfile = "\\bbc-01-wdmgmt\wd\sr" & Whs.Text & "\bin\SR" & Whs.Text & Format(Now, "mmdd") & ".csv"
    '    cfile = Form1.srserv & "\wd\sr" & Whs.Text & "\bin\SR" & Whs.Text & Format(Now, "mmdd") & ".csv"
    '    'MsgBox cfile
    '    'cfile = "c:\sr10430.csv"
    '    Open cfile For Append As #1
    '    Write #1, "SR-" & Whs.Text;
    '    Write #1, "...";
    '    Write #1, SGrid.TextMatrix(y, 3);
    '    Write #1, SGrid.TextMatrix(y, 5);
    '    Write #1, SGrid.TextMatrix(y, 6);
    '    Write #1, LTrim(StrConv(SGrid.TextMatrix(y, 4), vbProperCase));
    '    'Write #1, "WMS";
    '    Write #1, Form1.userid;
    '    Write #1, LTrim(Label2.Caption) & " " & SGrid.TextMatrix(y, 1);
    '    Write #1, "Cleared";
    '    Write #1, Format(Now, "h:mm am/pm")
    '    Close #1
    'End If
    
    sqlx = "Update position set posn_status = ' ', sku = ' ', lot_num = ' ', pallet_num = 0"
    sqlx = sqlx & ", lot_status = ' ', pallet_status = ' ', count_qty = 0"
    sqlx = sqlx & ", recv_date = '" & Format(Now, "m-d-yyyy") & "', barcode = ' '"
    sqlx = sqlx & ", lot2 = ' ', qty2 = 0 Where id = " & SGrid.TextMatrix(y, 0)
    Wdb.Execute sqlx
    
    
    SGrid.TextMatrix(y, 2) = " ": SGrid.TextMatrix(y, 3) = " "
    SGrid.TextMatrix(y, 4) = " ": SGrid.TextMatrix(y, 5) = " "
    SGrid.TextMatrix(y, 6) = "0": SGrid.TextMatrix(y, 7) = " "
    SGrid.TextMatrix(y, 8) = " ": SGrid.TextMatrix(y, 9) = "0"
    SGrid.TextMatrix(y, 10) = " ": SGrid.TextMatrix(y, 11) = " "
    SGrid.TextMatrix(y, 12) = Format$(Now, "m-dd-yyyy")
    SGrid.TextMatrix(y, 13) = " "
    pqty = 0
    olot = "99999"
    poc = 0                                                                             'jv011216
    For i = 1 To SGrid.Rows - 1
        If Val(SGrid.TextMatrix(i, 9)) > 0 Then
            pqty = pqty + 1
            If SGrid.TextMatrix(i, 5) < olot Then olot = SGrid.TextMatrix(i, 5)
            If SGrid.TextMatrix(i, 13) > "0" Then poc = Val(Mid(SGrid.TextMatrix(i, 13), 11, 3))    'jv011216
        End If
    Next i
    
    If pqty = 0 Then
        sqlx = "Update lane set qty = 0, sku = ' ', lot_num = ' ', resv_sku = ' ', resv_lot = ' '"
        sqlx = sqlx & ", lot_date = ' ', gmasize = 0, horz_travel = 0 Where id = " & pkey.Caption
        Wdb.Execute sqlx
        If Right$(SGrid.TextMatrix(0, 0), 1) = "R" Then
            RGrid.Text = ""
            RGrid.CellFontUnderline = False
        Else
            LGrid.Text = ""
            LGrid.CellFontUnderline = False
        End If
    Else
        sqlx = "Update lane set qty = " & pqty
        sqlx = sqlx & ", lot_num = '" & olot & "'"
        sqlx = sqlx & ", lot_date = '" & calc_date(olot) & "'"
        sqlx = sqlx & ", horz_travel = " & poc
        sqlx = sqlx & " Where id = " & pkey.Caption
        Wdb.Execute sqlx
    End If
    If y < SGrid.Rows - 1 Then y = y + 1
    SGrid.Col = 1: SGrid.Row = y
    Call SGrid_Click
End Sub

Private Sub Command4_Click()                            'Insert Pallet
    Dim psku As String, ppal As String, pdesc As String
    Dim plot As String, pdate As String, sqlx As String
    Dim olot As String, cfile As String, pbar As String
    Dim i As Integer, pqty As Integer, lqty As Integer
    Dim pqty2 As Integer, plot2 As String, psize As String, popl As String
    Dim ds As adodb.Recordset, y As Integer, pside As String, recid As Long
    Dim pplate As String                                                'jv070314
    Dim p As ptask                                                                          'jv060117
    psku = " ": ppal = "0": psize = 0: popl = "_"
    If Val(SGrid.TextMatrix(SGrid.Row, 9)) > 0 Then
        MsgBox "Position currently contains a pallet.", vbOKOnly, "Cannot Insert Here..."
        Exit Sub
    End If
    y = SGrid.Row: lqty = 1
    pside = Right$(SGrid.TextMatrix(0, 0), 1)
    olot = "99999"
    For i = 1 To SGrid.Rows - 1
        If Val(SGrid.TextMatrix(i, 9)) > 0 Then
            lqty = lqty + 1
            psku = SGrid.TextMatrix(i, 3)
            pdesc = Trim(SGrid.TextMatrix(i, 4))
            plot = SGrid.TextMatrix(i, 5)
            If plot < olot Then olot = plot
            If Val(SGrid.TextMatrix(i, 6)) >= Val(ppal) Then
                ppal = Val(SGrid.TextMatrix(i, 6)) + 1
            End If
            pqty = Val(SGrid.TextMatrix(i, 9))
            pdate = SGrid.TextMatrix(i, 10)
            'popl = Mid(SGrid.TextMatrix(i, 13), 12, 1)
            popl = Mid(SGrid.TextMatrix(i, 13), 11, 3)
        End If
    Next i
    
    'User Prompts
    psku = InputBox("SKU #", "Insert Position " & y, psku)
    If Len(psku) = 0 Then Exit Sub
    plot = InputBox("Lot #", "Insert Position " & y, plot)
    If Len(plot) = 0 Then Exit Sub
    ppal = InputBox("Pallet #", "Insert Position " & y, ppal)
    If Len(ppal) = 0 Then Exit Sub
    popl = InputBox("Operation Code:", "Insert Position " & y, popl)
    'If Len(popl) = 0 Or Len(popl) > 1 Then Exit Sub
    If Len(popl) = 0 Or Len(popl) > 3 Then Exit Sub                     'jv052515
    If Len(popl) = 1 Then                                               'jv052515
        popl = " " & popl & " "                                         'jv052515
    Else                                                                'jv052515
        If Len(popl) = 2 Then popl = " " & popl                         'jv052515
    End If                                                              'jv052515
    psize = InputBox("GMA Size:", "Insert Position " & y, 0)
    If Len(psize) = 0 Then Exit Sub
    
    psku = Trim(Left$(psku, 4))                                         'jv082415
    plot = Left$(plot, 5)
    ppal = UCase(ppal)
    popl = UCase(popl)
    If skurec(Val(psku)).sku <> psku Then
        MsgBox "Invalid SKU...", vbOKOnly, "Cannot Insert...."
        Exit Sub
    End If
    pdesc = skurec(Val(psku)).prodname
    If Val(psize) = 0 Then
        pqty = skurec(Val(psku)).uom_per_pallet
    Else
        pqty = CInt(psize)
    End If
    pdate = Format$(Now, "m-d-yyyy")
    pqty2 = pqty
    pqty2 = InputBox("Unit Qty for Lot " & plot & ":", "Lot " & plot & "units...", pqty2)
    If Len(pqty2) = 0 Then Exit Sub
    If pqty2 <> pqty Then
        i = pqty - pqty2    '308 - 200 = 108
        pqty = pqty2        '200
        pqty2 = i           '108
        plot2 = Format(Val(plot) + 1, "00000") & popl                           'jv052515
        plot2 = InputBox("Lot 2#", "Insert Position " & y, plot2)
        If Len(plot2) = 0 Then Exit Sub
    Else
        plot2 = " "
        pqty2 = 0
    End If
    pbar = psku
    If Len(psku) = 3 Then pbar = pbar & " "                             'jv082415
    pbar = pbar & Form1.bb_codedate(plot)
    pbar = pbar & popl & Format(ppal, "000")                            'jv052515
    sqlx = "Update position set posn_status = ' ', sku = '" & psku & "', lot_num = '" & plot & "'"
    sqlx = sqlx & ", pallet_num = '" & ppal & "', lot_status = ' ', pallet_status = ' '"            'jv072116
    sqlx = sqlx & ", count_qty = " & pqty & ", recv_date = '" & Format(Now, "M-d-yyyy") & "'"
    sqlx = sqlx & ", barcode = '" & pbar & "', lot2 = '" & plot2 & "', qty2 = " & pqty2
    sqlx = sqlx & " Where id = " & SGrid.TextMatrix(y, 0)
    Wdb.Execute sqlx
    
    SGrid.TextMatrix(y, 3) = psku
    SGrid.TextMatrix(y, 4) = " " & pdesc
    SGrid.TextMatrix(y, 5) = plot
    SGrid.TextMatrix(y, 6) = ppal
    SGrid.TextMatrix(y, 7) = " "
    SGrid.TextMatrix(y, 8) = " "
    SGrid.TextMatrix(y, 9) = pqty
    SGrid.TextMatrix(y, 10) = plot2
    SGrid.TextMatrix(y, 11) = pqty2
    SGrid.TextMatrix(y, 12) = pdate
    SGrid.TextMatrix(y, 13) = pbar
    If plot < olot Then olot = plot
    
    p.area = "SR-" & Whs                                                            'jv060117
    p.description = " "                                                             'jv060117
    p.source = "Insert Pallet"                                                      'jv060117
    p.target = Me.Label2 & " " & Trim(SGrid.TextMatrix(y, 1))                       'jv060117
    p.product = SGrid.TextMatrix(y, 3) & " " & UCase(SGrid.TextMatrix(y, 4))        'jv060117
    p.palletid = SGrid.TextMatrix(y, 13)                                            'jv060117
    p.qty = "1"                                                                     'jv060117
    p.uom = "Pallet"                                                                'jv060117
    p.lotnum = SGrid.TextMatrix(y, 5)                                               'jv060117
    p.units = SGrid.TextMatrix(y, 9)                                                'jv060117
    p.lotnum2 = SGrid.TextMatrix(y, 10)                                             'jv060117
    p.units2 = SGrid.TextMatrix(y, 11)                                              'jv060117
    p.status = "COMP"                                                               'jv060117
    p.userid = Form1.userid                                                         'jv060117
    p.trandate = Format(Now, "yyMMdd hh:mm:ss")                                     'jv060117
    p.reqid = ".."                                                                  'jv060117
    cfile = Form1.logdir & "wms" & Format(Now, "mmddyyyy") & ".txt"                 'jv060117
    'If LCase(Form1.userid) <> "jvierus" Then                                        'jv060117
        Open cfile For Append Shared As #1                                          'jv060117
        Write #1, y;                                                                'jv060117
        Write #1, p.area;                                                           'jv060117
        Write #1, p.description;                                                    'jv060117
        Write #1, p.source;                                                         'jv060117
        Write #1, p.target;                                                         'jv060117
        Write #1, p.product;                                                        'jv060117
        Write #1, p.palletid;                                                       'jv060117
        Write #1, p.qty;                                                            'jv060117
        Write #1, p.uom;                                                            'jv060117
        Write #1, p.lotnum;                                                         'jv060117
        Write #1, p.units;                                                          'jv060117
        Write #1, p.lotnum2;                                                        'jv060117
        Write #1, p.units2;                                                         'jv060117
        Write #1, p.status;                                                         'jv060117
        Write #1, p.userid;                                                         'jv060117
        Write #1, p.trandate;                                                       'jv060117
        Write #1, p.reqid                                                           'jv060117
        Close #1                                                                    'jv060117
    'End If                                                                          'jv060117
    
    'If Form1.plantno = "50" Then
    '    'Add to crane movement log
    '    'cfile = "\\bbc-01-wdmgmt\wd\sr" & Whs.Text & "\bin\SR" & Whs.Text & Format(Now, "mmdd") & ".csv"
    '    cfile = Form1.srserv & "\wd\sr" & Whs.Text & "\bin\SR" & Whs.Text & Format(Now, "mmdd") & ".csv"
    '    'MsgBox cfile
    '    'cfile = "c:\sr10430.csv"
    '    Open cfile For Append As #1
    '    Write #1, "SR-" & Whs.Text;
    '    Write #1, "...";
    '    Write #1, SGrid.TextMatrix(y, 3);
    '    Write #1, SGrid.TextMatrix(y, 5);
    '    Write #1, SGrid.TextMatrix(y, 6);
    '    Write #1, LTrim(StrConv(SGrid.TextMatrix(y, 4), vbProperCase));
    '    'Write #1, "WMS";
    '    Write #1, Form1.userid;
    '    Write #1, "Insert";
    '    Write #1, LTrim(Label2.Caption) & " " & SGrid.TextMatrix(y, 1);
    '    Write #1, Format(Now, "h:mm am/pm")
    '    Close #1
    'End If
    
    sqlx = "Update lane set qty = " & lqty & ", sku = '" & psku & "', lot_num = '" & olot & "'"
    sqlx = sqlx & ", lot_date = '" & Format(calc_date(olot), "M-d-yyyy") & "'"
    sqlx = sqlx & ", gmasize = " & psize & ", horz_travel = " & Val(popl)
    sqlx = sqlx & " Where id = " & pkey
    Wdb.Execute sqlx
    
    'Update pallet record
    If Val(psku) >= 100 Then                                                            'jv070314
        recid = 0
        s = "select * from pallets where barcode = '" & pbar & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            pplate = ds!plateno                                                         'jv070314
            recid = ds!id
        Else
            ds.Close
            s = "select * from pallets where status in ('Shipped','Order Pick')"
            s = s & " order by trandate"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                recid = ds!id
                pplate = " "
            End If
        End If
        ds.Close
        If recid > 0 Then
            's = "Update pallets set plateno = '" & recid & "'"
            s = "Update pallets set plateno = '" & pplate & "'"                         'jv070314
            s = s & ",barcode = '" & pbar & "'"
            s = s & ",qty1 = " & Val(pqty)
            s = s & ",lot1 = '" & plot & "'"
            s = s & ",qty2 = " & Val(pqty2)
            s = s & ",lot2 = '" & plot2 & "'"
            s = s & ",source = 'SR-" & Whs & "'"
            'If Whs = "5" Then
            '    s = s & ",target = '" & LGrid.TextMatrix(LGrid.Row, 1) & " " & Label2.Caption & "'"
            'Else
                s = s & ",target = '" & Label2.Caption & "'"
            'End If
            If psize = 0 Then
                s = s & ",bbc = 'Y'"
            Else
                s = s & ",bbc = 'N'"
            End If
            s = s & ",status = 'Warehouse'"
            s = s & ",trandate = '" & Format(Now, "yyMMdd hh:mm:ss") & "'"
            s = s & ",sku = '" & psku & "'"
            s = s & " Where id = " & recid
            Wdb.Execute s
        Else
            pid = wd_seq("Pallets")
            s = "Insert Into pallets Values (" & pid
            's = s & ",'" & recid & "'"
            s = s & ",'" & pplate & "'"                                                 'jv070314
            s = s & ",'" & pbar & "'"
            s = s & "," & Val(pqty)
            s = s & ",'" & plot & "'"
            s = s & "," & Val(pqty2)
            s = s & ",'" & plot2 & "'"
            s = s & ",'SR-" & Whs & "'"
            'If Whs = "5" Then
            '    s = s & ",'" & LGrid.TextMatrix(LGrid.Row, 1) & " " & Label2.Caption & "'"
            'Else
                s = s & ",'" & Label2.Caption & "'"
            'End If
            If psize = 0 Then
                s = s & ",'Y'"
            Else
                s = s & ",'N'"
            End If
            s = s & ",'Warehouse'"
            s = s & ",'" & Format(Now, "yyMMdd hh:mm:ss") & "'"
            s = s & ",'" & psku & "')"
            Wdb.Execute s
        End If
        'MsgBox s
    End If                                                                                  'jv070314
    
    If pside = "R" Then    'Brenham
    'If pside = "2" Then     'Sylacauga
        RGrid.Text = "X"
        RGrid.CellFontUnderline = True
    Else
        LGrid.Text = "X"
        LGrid.CellFontUnderline = True
    End If
    If y > 1 Then y = y - 1
    SGrid.Col = 1: SGrid.Row = y
    Call SGrid_Click
End Sub

Private Sub Command5_Click()
    Dim sqlx As String, i As Integer
    If MsgBox("Ok to clear lane " & SGrid.TextMatrix(0, 0) & " reservation?", vbYesNo, "Are you sure?") = vbNo Then Exit Sub
    sqlx = "Update lane set resv_sku = ' ', resv_lot = ' ' where id = " & pkey.Caption
    Wdb.Execute sqlx
    'If Right$(SGrid.TextMatrix(0, 0), 1) = "R" Then
    If Right$(Label2, 1) = "R" Then
        If RGrid.Text = "R" Then RGrid.Text = ""
        RGrid.CellFontUnderline = False
    Else
        If LGrid.Text = "R" Then LGrid.Text = ""
        LGrid.CellFontUnderline = False
    End If
    For i = 1 To SGrid.Rows - 1
        sqlx = "Update position set posn_status = ' ' Where id = " & SGrid.TextMatrix(i, 0)
        Wdb.Execute sqlx
        SGrid.TextMatrix(i, 2) = " "
    Next i
    Call pkey_Change
End Sub

Private Sub Command6_Click()
    Dim i As Integer, k As Integer, j As Integer
    j = 0: k = 1
    Screen.MousePointer = 11
    Call print_grid(LGrid, 1, 7, "SR-" & Whs & " Zones")
    Screen.MousePointer = 0
End Sub

Private Sub Command7_Click()
    Dim cv As String, ch As String
    Dim i As Integer, k As Integer, j As Integer
    Dim d As Single
    cv = InputBox("CV:", "Vert Center..", "3")
    ch = InputBox("CH:", "Horz Center..", "30")
    j = 7
    For i = 1 To 7
        For k = 1 To LGrid.Cols - 1
            d = Int(Abs(cv - i) * 3) + Abs(ch - k)
            LGrid.TextMatrix(j, k) = d
        Next k
        j = j - 1
    Next i
End Sub

Private Sub edlane1_Click()
    Command1_Click
End Sub

Private Sub edlane2_Click()
    Command5_Click
End Sub

Private Sub edlane3_Click()
    Command2_Click
End Sub

Private Sub edpos1_Click()
    Command3_Click
End Sub

Private Sub edpos2_Click()
    Command4_Click
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Form3.WindowState = 0 Then
        For i = 1 To Form1.Frmgrid.Rows - 1
            If Form1.Frmgrid.TextMatrix(i, 0) = "form3" Then
                Form1.Frmgrid.TextMatrix(i, 1) = Form3.Top
                Form1.Frmgrid.TextMatrix(i, 2) = Form3.Left
                Form1.Frmgrid.TextMatrix(i, 3) = Form3.Height
                Form1.Frmgrid.TextMatrix(i, 4) = Form3.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer, j As Integer
    Dim ds As adodb.Recordset, sqlx As String
    For i = 1 To Form1.Frmgrid.Rows - 1
        If Form1.Frmgrid.TextMatrix(i, 0) = "form3" Then
            Form3.Top = Val(Form1.Frmgrid.TextMatrix(i, 1))
            Form3.Left = Val(Form1.Frmgrid.TextMatrix(i, 2))
            Form3.Height = Val(Form1.Frmgrid.TextMatrix(i, 3))
            Form3.Width = Val(Form1.Frmgrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    Col(0) = &HFFFFFF: fgc(0) = &H0&
    Col(15) = &H80FFFF: fgc(15) = &H0&
    Col(2) = &HFF&: fgc(2) = &HFFFFFF
    Col(8) = &H404080: fgc(8) = &HFFFFFF
    Col(4) = &HFFFF&: fgc(4) = &H0&
    'col(5) = &HFF00: fgc(5) = &HFFFFFF
    Col(5) = &H808000: fgc(f) = &H0&
    Col(7) = &H80FF80: fgc(7) = &H0&
    Col(6) = &HFF0000: fgc(6) = &HFFFFFF
    Col(3) = &HFF00FF: fgc(3) = &HFFFFFF
    Col(10) = &H808080: fgc(10) = &HFFFFFF
    Col(12) = &HC0&: fgc(12) = &HFFFFFF
    Col(11) = &H40C0&: fgc(11) = &HFFFFFF
    Col(9) = &H8080&: fgc(9) = &HFFFFFF
    Col(13) = &H80FF80: fgc(13) = &H0&
    Col(14) = &H404000: fgc(14) = &HFFFFFF
    Col(1) = &HC00000: fgc(1) = &H0&
    Col(16) = &HC000C0: fgc(16) = &HFFFFFF
    Col(17) = &H404040: fgc(17) = &HFFFFFF
    Col(18) = &H80&: fgc(18) = &HFFFFFF
    Col(19) = &H4080&: fgc(19) = &HFFFFFF
    Col(20) = &HC0C0&: fgc(20) = &HFFFFFF
    Col(21) = &H8000&: fgc(21) = &HFFFFFF
    Col(22) = &H808000: fgc(22) = &HFFFFFF
    Col(23) = &H800000: fgc(23) = &HFFFFFF
    Col(24) = &H800080: fgc(24) = &HFFFFFF
    For i = 0 To 14
        Col(i) = QBColor(15 - i)
    Next i
    For i = 0 To 100
        fc(i) = &H0&: bc(i) = &HFFFFFF
    Next i
    sqlx = "Select distinct zone_num from lane order by zone_num"
    Set ds = Wdb.Execute(sqlx)
    j = 0
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            'bc(ds(0)) = col(j)
            'fc(ds(0)) = fgc(j)
            bc(ds(0)) = dot(j).BackColor
            fc(ds(0)) = dot(j).ForeColor
            If j < 18 Then
                ckey(j).Visible = True
                ckey(j) = ds(0)
                'ckey(j).BackColor = col(j)
                'ckey(j).ForeColor = fgc(j)
                ckey(j).BackColor = dot(j).BackColor
                ckey(j).ForeColor = dot(j).ForeColor
                
            End If
            j = j + 1
            ds.MoveNext
        Loop
    End If
    ds.Close
    Whs.AddItem "1"
    Whs.AddItem "2"
    Whs.AddItem "3"
    'Whs.AddItem "5"
    LGrid.FontName = "Arial": LGrid.FontBold = True
    RGrid.FontName = "Arial": RGrid.FontBold = True
    SGrid.FontName = "Arial": SGrid.FontBold = True
    LGrid.Rows = 8: RGrid.Rows = 8
    LGrid.Cols = 59: RGrid.Cols = 59
    LGrid.TextMatrix(0, 0) = "Left"
    RGrid.TextMatrix(0, 0) = "Right"
    For i = 1 To 7
        LGrid.TextMatrix(8 - i, 0) = "Level " & i
        RGrid.TextMatrix(8 - i, 0) = "Level " & i
    Next i
    SGrid.Cols = 14: SGrid.Row = 0
    'SGrid.FormatString = "ID|^Pos|^PosStat|^SKU|<Description|^Lot|^Pallet|^LotStat|^PalStat|^Qty|^Lot2|^Qty2|^Date|^BarCode"
    SGrid.FormatString = "ID|^Pos|^PosStat|^SKU|<Description|^Lot|^Pallet|^|^PalStat|^Qty|^Lot2|^Qty2|^Date|^BarCode"
    SGrid.ColWidth(0) = 1
    SGrid.ColWidth(1) = 600: SGrid.ColWidth(2) = 1000
    SGrid.ColWidth(3) = 600: SGrid.ColWidth(4) = 3200
    SGrid.ColWidth(5) = 900: SGrid.ColWidth(6) = 800
    SGrid.ColWidth(7) = 1: SGrid.ColWidth(8) = 900
    SGrid.ColWidth(9) = 800: SGrid.ColWidth(10) = 900
    SGrid.ColWidth(11) = 900
    SGrid.ColWidth(12) = 1200
    SGrid.ColWidth(13) = 2000
    Whs.ListIndex = 0
    'refresh_grid1
End Sub

Private Sub Form_Resize()
    LGrid.Width = Form3.Width
    RGrid.Width = Form3.Width
    If Form3.Width > 14495 Then
        SGrid.Width = 14495 '9555
    Else
        SGrid.Width = Form3.Width
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub LGrid_Click()
    zone = Val(LGrid.Text)
    Command1.Visible = Not LGrid.CellFontUnderline
    edlane1.Enabled = Command1.Visible
    pkey = lkey(0, LGrid.Row, LGrid.Col)
    'SGrid.col = 0: SGrid.Row = 0
    SGrid.TextMatrix(0, 0) = " " & (8 - LGrid.Row) & " " & LGrid.Col & " L"
    Label2 = " " & (8 - LGrid.Row) & " " & LGrid.Col & " L"
    If LGrid.Text = "H" Then
        SGrid.BackColor = LGrid.CellBackColor
        SGrid.ForeColor = LGrid.CellForeColor
        SGrid.FillStyle = flexFillRepeat
        For i = 1 To SGrid.Rows - 1
            If Val(SGrid.TextMatrix(i, 3)) > 0 Then
                SGrid.Row = i: SGrid.RowSel = i
                SGrid.Col = 3: SGrid.ColSel = SGrid.Cols - 1
                SGrid.CellBackColor = dot(0).BackColor
                SGrid.CellForeColor = dot(0).ForeColor
            End If
        Next i
        SGrid.Row = 1
    Else
        SGrid.BackColor = LGrid.CellBackColor
        SGrid.ForeColor = LGrid.CellForeColor
    End If
    Label2.BackColor = LGrid.CellBackColor
    Label2.ForeColor = LGrid.CellForeColor
End Sub

Private Sub LGrid_EnterCell()
    Call LGrid_Click
End Sub

Private Sub LGrid_KeyPress(KeyAscii As Integer)
    Dim sqlx As String
    If KeyAscii = 8 Then
        If Len(LGrid.Text) > 0 Then
            LGrid.Text = Left$(LGrid.Text, Len(LGrid.Text) - 1)
        End If
    End If
    If KeyAscii = 32 Then LGrid.Text = ""
    If KeyAscii >= 48 Or KeyAscii <= 57 Then
        LGrid.Text = LGrid.Text & Chr$(KeyAscii)
        LGrid.Text = Val(LGrid.Text)
        sqlx = "update lane set zone_num = " & Val(LGrid.Text)
        sqlx = sqlx & " where id = " & lkey(0, LGrid.Row, LGrid.Col)
        Wdb.Execute sqlx
    End If
    If Val(LGrid.Text) > 99 Then LGrid.Text = Left$(LGrid.Text, 2)
    LGrid.CellBackColor = bc(Val(LGrid.Text))
    LGrid.CellForeColor = fc(Val(LGrid.Text))
End Sub

Private Sub LGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edlane
End Sub

Private Sub pkey_Change()
    Dim ds As adodb.Recordset, sqlx As String, i As Integer, flag4 As Boolean, ss As adodb.Recordset
    If SGrid.Visible = False Then Exit Sub
    Screen.MousePointer = 11
    SGrid.Rows = 1
    Label3 = "": Label4 = "": Label5 = "": flag4 = False
    sqlx = "Select * from lane where id = " & pkey
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        If ds!resv_sku > "000" Then
            Label3 = "Reserved:"
            Label4 = ds!resv_sku & " " & ds!resv_lot
        End If
        If ds!lock_status = 1 Then
            Label5 = "LIFO Bay"
        End If
        If ds!lane_status = "B" Then Label5 = "Blocked Bay"
        If ds!lane_status = "H" Then Label5 = "Product On Hold"
        If ds!gmasize > 0 Then flag4 = True
    End If
    ds.Close
    opgrid.Visible = False                                                      'jv110716
    sqlx = "Select * From Position where laneno = " & pkey
    sqlx = sqlx & " Order by posn_num"
    Set ds = Wdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If ds!whse_num = 1 And ds!vert_loc < 4 And ds!rack_side = "R" Then      'jv110716
            sqlx = "select * from opbays where whse_num = 1"                    'jv110716
            sqlx = sqlx & " and vert_loc = " & ds!vert_loc                      'jv110716
            sqlx = sqlx & " and horz_loc = " & ds!horz_loc                      'jv110716
            sqlx = sqlx & " and rack_side = '" & ds!rack_side & "'"             'jv110716
            Set ss = Wdb.Execute(sqlx)                                          'jv110716
            If ss.BOF = False Then                                              'jv110716
                ss.MoveFirst                                                    'jv110716
                Label3 = "Order Pick:"                                          'jv110716
                Label4 = ss!sku & " " & ss!oplabel                              'jv110716
                opgrid.Visible = True                                           'jv110716
                Call refresh_opgrid(ss!sku)                                     'jv110716
            End If                                                              'jv110716
            ss.Close                                                            'jv110716
        End If                                                                  'jv110716
        Do Until ds.EOF
            sqlx = ds!id & Chr$(9)
            sqlx = sqlx & ds!posn_num & Chr$(9)
            sqlx = sqlx & ds!posn_status & Chr$(9)
            sqlx = sqlx & ds!sku & Chr$(9)
            sqlx = sqlx & " " & Chr$(9)
            sqlx = sqlx & ds!lot_num & Chr$(9)
            sqlx = sqlx & ds!pallet_num & Chr$(9)
            sqlx = sqlx & ds!lot_status & Chr$(9)
            sqlx = sqlx & ds!pallet_status & Chr$(9)
            sqlx = sqlx & ds!count_qty & Chr$(9)
            sqlx = sqlx & ds!lot2 & Chr(9)
            sqlx = sqlx & ds!qty2 & Chr(9)
            sqlx = sqlx & Format$(ds!recv_date, "m-dd-yyyy") & Chr(9)
            sqlx = sqlx & ds!barcode
            SGrid.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    For i = 1 To SGrid.Rows - 1
        If Val(SGrid.TextMatrix(i, 3)) > 0 Then
            sqlx = "Select * from sku_config where sku in "
            sqlx = sqlx & "(select sku from position where laneno = " & pkey & ")"
            Set ds = Wdb.Execute(sqlx)
            If ds.BOF = True Then
                SGrid.TextMatrix(i, 4) = "Invalid SKU"
            Else
                SGrid.TextMatrix(i, 4) = " " & ds!uom_type & " " & ds!description
            End If
            ds.Close
            If flag4 = True Then SGrid.TextMatrix(i, 4) = SGrid.TextMatrix(i, 4) & " 4-Way"
        Else
            SGrid.TextMatrix(i, 4) = " "
        End If
    Next i
    Screen.MousePointer = 0
End Sub


Private Sub RGrid_Click()
    zone = Val(RGrid.Text)
    Command1.Visible = Not RGrid.CellFontUnderline
    edlane1.Enabled = Command1.Visible
    pkey = lkey(1, RGrid.Row, RGrid.Col)
    'SGrid.col = 0: SGrid.Row = 0
    SGrid.TextMatrix(0, 0) = " " & (8 - RGrid.Row) & " " & RGrid.Col & " R"
    Label2 = " " & (8 - RGrid.Row) & " " & RGrid.Col & " R"
    'SGrid.BackColor = RGrid.CellBackColor
    'SGrid.ForeColor = RGrid.CellForeColor
    If RGrid.Text = "H" Then
        SGrid.BackColor = RGrid.CellBackColor
        SGrid.ForeColor = RGrid.CellForeColor
        SGrid.FillStyle = flexFillRepeat
        For i = 1 To SGrid.Rows - 1
            If Val(SGrid.TextMatrix(i, 3)) > 0 Then
                SGrid.Row = i: SGrid.RowSel = i
                SGrid.Col = 3: SGrid.ColSel = SGrid.Cols - 1
                SGrid.CellBackColor = dot(0).BackColor
                SGrid.CellForeColor = dot(0).ForeColor
            End If
        Next i
        SGrid.Row = 1
    Else
        SGrid.BackColor = RGrid.CellBackColor
        SGrid.ForeColor = RGrid.CellForeColor
    End If
    Label2.BackColor = RGrid.CellBackColor
    Label2.ForeColor = RGrid.CellForeColor
End Sub

Private Sub RGrid_EnterCell()
    Call RGrid_Click
End Sub

Private Sub RGrid_KeyPress(KeyAscii As Integer)
    Dim sqlx As String
    If KeyAscii = 8 Then
        If Len(RGrid.Text) > 0 Then
            RGrid.Text = Left$(RGrid.Text, Len(RGrid.Text) - 1)
        End If
    End If
    If KeyAscii = 32 Then RGrid.Text = ""
    If KeyAscii >= 48 Or KeyAscii <= 57 Then
        RGrid.Text = RGrid.Text & Chr$(KeyAscii)
        RGrid.Text = Val(RGrid.Text)
        sqlx = "update lane set zone_num = " & Val(RGrid.Text)
        sqlx = sqlx & " where id = " & lkey(1, RGrid.Row, RGrid.Col)
        Wdb.Execute sqlx
    End If
    If Val(RGrid.Text) > 99 Then RGrid.Text = Left$(RGrid.Text, 2)
    RGrid.CellBackColor = bc(Val(RGrid.Text))
    RGrid.CellForeColor = fc(Val(RGrid.Text))
End Sub

Private Sub RGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edlane
End Sub

Private Sub SGrid_Click()
    SGrid.ColSel = SGrid.Cols - 1
End Sub

Private Sub SGrid_KeyPress(KeyAscii As Integer)
    If KeyPress = 13 Then Call SGrid_Click
End Sub

Private Sub SGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edpos
End Sub

Private Sub vphist_Click()
    palhistory.Show
    palhistory.barkey = SGrid.TextMatrix(SGrid.Row, 13)
End Sub

Private Sub Whs_Click()
    Dim i As Integer, lfmt As String, rfmt As String
    Dim j As Integer, k As Integer
    For i = 0 To 1
        For j = 0 To 7
            For k = 0 To 58
                lkey(i, j, k) = 0
            Next k
        Next j
    Next i
    LGrid.Clear: RGrid.Clear
    LGrid.Visible = False: RGrid.Visible = False
    LGrid.TextMatrix(0, 0) = "Left"
    RGrid.TextMatrix(0, 0) = "Right"
    For i = 1 To 7
        LGrid.TextMatrix(8 - i, 0) = i '"Level " & i
        RGrid.TextMatrix(8 - i, 0) = i '"Level " & i
    Next i
    'LGrid.Row = 0: RGrid.Row = 0
    If Whs = "1" Then
        LGrid.Cols = 51: RGrid.Cols = 51
        lfmt = "^Left": rfmt = "^Right"
        For i = 1 To 50
            lfmt = lfmt & "|^" & i: rfmt = rfmt & "|^" & i
        Next i
        LGrid.FormatString = lfmt: RGrid.FormatString = rfmt
        LGrid.ColWidth(0) = 700: RGrid.ColWidth(0) = 700
        For i = 1 To 50
            LGrid.ColWidth(i) = 300: RGrid.ColWidth(i) = 300
        Next i
    End If
    If Whs = "2" Then
        LGrid.Cols = 55: RGrid.Cols = 55
        lfmt = "^Left": rfmt = "^Right"
        For i = 1 To 54
            lfmt = lfmt & "|^" & i: rfmt = rfmt & "|^" & i
        Next i
        LGrid.FormatString = lfmt: RGrid.FormatString = rfmt
        LGrid.ColWidth(0) = 700: RGrid.ColWidth(0) = 700
        For i = 1 To 54
            LGrid.ColWidth(i) = 300: RGrid.ColWidth(i) = 300
        Next i
    End If
    If Whs = "3" Then
        LGrid.Cols = 59: RGrid.Cols = 59
        lfmt = "^Left": rfmt = "^Right"
        For i = 1 To 58
            lfmt = lfmt & "|^" & i: rfmt = rfmt & "|^" & i
        Next i
        LGrid.FormatString = lfmt: RGrid.FormatString = rfmt
        LGrid.ColWidth(0) = 700: RGrid.ColWidth(0) = 700
        For i = 1 To 58
            LGrid.ColWidth(i) = 300: RGrid.ColWidth(i) = 300
        Next i
    End If
    Call Refresh_zones
    LGrid.Visible = True: RGrid.Visible = True
End Sub
