VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form trucknotes 
   Caption         =   "Trailer Schedule Notes"
   ClientHeight    =   11025
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12975
   LinkTopic       =   "Form3"
   ScaleHeight     =   11025
   ScaleWidth      =   12975
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Excel"
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
      TabIndex        =   49
      Top             =   9720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   47
      Text            =   "trucknotes.frx":0000
      Top             =   9120
      Width           =   9975
   End
   Begin MSFlexGridLib.MSFlexGrid pgrid 
      Height          =   1335
      Left            =   0
      TabIndex        =   45
      Top             =   10440
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2355
      _Version        =   327680
   End
   Begin VB.CommandButton Command2 
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
      Left            =   8520
      TabIndex        =   43
      Top             =   0
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid Grid2 
      Height          =   1455
      Left            =   1800
      TabIndex        =   42
      Top             =   10200
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2566
      _Version        =   327680
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   120
      Width           =   1815
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
      Left            =   2160
      TabIndex        =   39
      Top             =   9720
      Width           =   2295
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   11400
      TabIndex        =   36
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   9240
      TabIndex        =   35
      Text            =   "Combo1"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox List10 
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
      Left            =   7560
      TabIndex        =   34
      Top             =   7200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox Combo10 
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
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   7200
      Width           =   3855
   End
   Begin VB.ListBox List9 
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
      Left            =   12360
      TabIndex        =   32
      Top             =   6840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo9 
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
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   6840
      Width           =   3135
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Text            =   "trucknotes.frx":0007
      Top             =   8520
      Width           =   9975
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Text            =   "trucknotes.frx":000E
      Top             =   7920
      Width           =   9975
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   28
      Text            =   "Text11"
      Top             =   7560
      Width           =   9975
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   24
      Text            =   "Text10"
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
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
      Left            =   7560
      TabIndex        =   23
      Text            =   "Text9"
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text8 
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
      Left            =   2160
      TabIndex        =   22
      Text            =   "Text8"
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
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
      Left            =   7560
      TabIndex        =   21
      Text            =   "Text7"
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   20
      Text            =   "Text6"
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   19
      Text            =   "Text5"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   18
      Text            =   "Text4"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
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
      Left            =   7560
      TabIndex        =   17
      Text            =   "Text3"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
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
      Left            =   7560
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text0 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   14
      Text            =   "Text0"
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox truckdb 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   10200
      Visible         =   0   'False
      Width           =   8295
   End
   Begin VB.TextBox shipdb 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   10560
      Visible         =   0   'False
      Width           =   8415
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7858
      _Version        =   327680
      ForeColor       =   4194368
      BackColorFixed  =   8454143
      BackColorSel    =   8388736
      WordWrap        =   -1  'True
      FocusRect       =   0
      AllowUserResizing=   3
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Branch Notes"
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
      TabIndex        =   48
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Label pcolor 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Order not received via W/D browser."
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
      Left            =   4560
      TabIndex        =   46
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label sdate 
      Caption         =   "Label15"
      Height          =   255
      Left            =   11520
      TabIndex        =   44
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ship Date:"
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
      TabIndex        =   40
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label sdest 
      Caption         =   "sdest"
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
      Left            =   3600
      TabIndex        =   38
      Top             =   6120
      Width           =   3975
   End
   Begin VB.Label sorg 
      Caption         =   "sorg"
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
      Left            =   3600
      TabIndex        =   37
      Top             =   5760
      Width           =   3855
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label13"
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
      TabIndex        =   27
      Top             =   8520
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label12"
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
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label11"
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
      TabIndex        =   25
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label10"
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
      TabIndex        =   13
      Top             =   7200
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
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
      Left            =   5640
      TabIndex        =   12
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label8"
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
      TabIndex        =   11
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   5640
      TabIndex        =   10
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
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
      TabIndex        =   9
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   240
      TabIndex        =   8
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   240
      TabIndex        =   7
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
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
      Left            =   5640
      TabIndex        =   6
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   240
      TabIndex        =   5
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   5640
      TabIndex        =   4
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label0 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label0"
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
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Menu edmenu 
      Caption         =   "E&dit"
      Begin VB.Menu addtkt 
         Caption         =   "Add Ticket"
      End
      Begin VB.Menu deltkt 
         Caption         =   "Remove Ticket"
      End
      Begin VB.Menu clrdate 
         Caption         =   "Clear Date"
      End
   End
   Begin VB.Menu impmenu 
      Caption         =   "Import"
      Begin VB.Menu impsched 
         Caption         =   "Trailer Schedule"
      End
   End
   Begin VB.Menu repmenu 
      Caption         =   "Reports"
      Begin VB.Menu t10rep 
         Caption         =   "Brenham"
         Begin VB.Menu t10repall 
            Caption         =   "All Trailers"
         End
         Begin VB.Menu t10repbo 
            Caption         =   "Browser Orders"
         End
      End
      Begin VB.Menu k10rep 
         Caption         =   "Broken Arrow"
         Begin VB.Menu k10repall 
            Caption         =   "All Trailers"
         End
         Begin VB.Menu k10repbo 
            Caption         =   "Browser Orders"
         End
      End
      Begin VB.Menu a10rep 
         Caption         =   "Sylacauga"
         Begin VB.Menu a10repall 
            Caption         =   "All Trailers"
         End
         Begin VB.Menu a10repbo 
            Caption         =   "Browser Orders"
         End
      End
      Begin VB.Menu tktstatus 
         Caption         =   "Ticket Status"
      End
   End
   Begin VB.Menu wksmenu 
      Caption         =   "Worksheets"
      Begin VB.Menu wksyard 
         Caption         =   "Yard"
      End
      Begin VB.Menu wksload 
         Caption         =   "Loader"
      End
      Begin VB.Menu wksdriver 
         Caption         =   "Driver"
      End
   End
End
Attribute VB_Name = "trucknotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub cre8xml(pnote As String)
    Dim cfile As String, i As Integer, s As String
    If Grid1.Rows = 1 Then Exit Sub
    If pnote = "Driver" Then cfile = localAppDataPath & "\drivewrk.xml"   'u:\drivewrk.xml
    If pnote = "Yard" Then cfile = localAppDataPath & "\yardwrk.xml"      'u:\yardwrk.xml
    If pnote = "Loader" Then cfile = localAppDataPath & "\loaderwrk.xml"  'u:\loaderwrk.xml
    'cfile = Form1.tempdir & "\aschedwrk.csv"
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 9
    For i = 1 To Grid1.Rows - 1
        If (Grid1.TextMatrix(i, 4) = "T10" Or Grid1.TextMatrix(i, 4) = "50") And Grid1.TextMatrix(i, 14) <= " " Then
            s = Grid1.TextMatrix(i, 3) & Chr(9)
            s = s & Combo2 & Chr(9)
            s = s & truck_loc_name(Grid1.TextMatrix(i, 4)) & Chr(9)
            s = s & truck_loc_name(Grid1.TextMatrix(i, 5)) & Chr(9)
            s = s & "#" & Val(Grid1.TextMatrix(i, 6)) & Chr(9)
            s = s & Grid1.TextMatrix(i, 8) & Chr(9)
            s = s & Grid1.TextMatrix(i, 7) & Chr(9)
            If pnote = "Driver" Then
                s = s & Grid1.TextMatrix(i, 11) & Chr(9)             'Driver Note
            End If
            If pnote = "Yard" Then
                s = s & Grid1.TextMatrix(i, 12) & Chr(9)            'Yard Note
            End If
            If pnote = "Loader" Then
                s = s & Grid1.TextMatrix(i, 13) & Chr(9)            'Loader Note
            End If
            s = s & Grid1.TextMatrix(i, 9)
            pgrid.AddItem s
        End If
    Next i
    pgrid.FormatString = "^Group|^Date|<Plant|<Branch|^Trailer|^Size|^Start|<" & pnote & "_Note|^TCode"
    'MsgBox cfile
    Call xmlgrid(cfile, pgrid, True)
    
    Call OpenFileInExcel(cfile)
End Sub

Sub cre8wks(pnote As String)
    Dim cfile As String, i As Integer, x
    If Grid1.Rows = 1 Then Exit Sub
    If pnote = "Driver" Then cfile = localAppDataPath & "\drivewrk.csv"   '"u:\drivewrk.csv"
    If pnote = "Yard" Then cfile = localAppDataPath & "\yardwrk.csv"      '"u:\yardwrk.csv"
    If pnote = "Loader" Then cfile = localAppDataPath & "\loaderwrk.csv"  '"u:\loaderwrk.csv"
    'cfile = Form1.tempdir & "\aschedwrk.csv"
    Open cfile For Output As #1
    Write #1, "Group"; "Date"; "Plant"; "Branch"; "Trailer"; "Size"; "Start"; pnote & " Note"; "TCode"
    Write #1, " "
    For i = 1 To Grid1.Rows - 1
        If (Grid1.TextMatrix(i, 4) = "T10" Or Grid1.TextMatrix(i, 4) = "50") And Grid1.TextMatrix(i, 14) <= " " Then
            Write #1, Grid1.TextMatrix(i, 3);                       'Group
            Write #1, "'" & Combo2;                                 'Date
            Write #1, truck_loc_name(Grid1.TextMatrix(i, 4));       'Plant
            Write #1, truck_loc_name(Grid1.TextMatrix(i, 5));       'Branch
            Write #1, "'  #" & Val(Grid1.TextMatrix(i, 6));         'Trailer #
            Write #1, "'  " & Grid1.TextMatrix(i, 8);               'Size
            Write #1, "'" & Grid1.TextMatrix(i, 7);                 'Start
            If pnote = "Driver" Then
                Write #1, "'" & Grid1.TextMatrix(i, 11);            'Driver Note
            End If
            If pnote = "Yard" Then
                Write #1, "'" & Grid1.TextMatrix(i, 12);            'Yard Note
            End If
            If pnote = "Loader" Then
                Write #1, "'" & Grid1.TextMatrix(i, 13);            'Loader Note
            End If
            Write #1, "'  " & Grid1.TextMatrix(i, 9) & "  "         'Tcode Oc
        End If
    Next i
    Close #1
    'MsgBox cfile
    
    If Not OpenFileInExcel(cfile) Then
        MsgBox "Created file at: " & cfile, vbInformation + vbOKOnly, "Export completed...."
        x = Shell("notepad.exe " & cfile, vbNormalFocus)
    End If
End Sub

Sub listreport(pcode As String, bocode As String)
    Dim i As Integer, s As String
    Dim rt As String, rh As String, rf As String, hf As String
    Dim scode As String
    If pcode = "T10" Then scode = "50"
    If pcode = "K10" Then scode = "51"
    If pcode = "A10" Then scode = "52"
    pgrid.Clear: pgrid.Rows = 1: pgrid.Cols = 9
    For i = 1 To Grid1.Rows - 1
        If (Grid1.TextMatrix(i, 4) = pcode Or Grid1.TextMatrix(i, 4) = scode) And Grid1.TextMatrix(i, 14) <= bocode Then
            s = Grid1.TextMatrix(i, 1) & Chr(9)
            s = s & Grid1.TextMatrix(i, 3) & Chr(9) & Grid1.TextMatrix(i, 5) & "-"
            s = s & truck_loc_name(Grid1.TextMatrix(i, 5)) & " #" & Grid1.TextMatrix(i, 6) & Chr(9)
            s = s & Grid1.TextMatrix(i, 7) & Chr(9)
            s = s & Grid1.TextMatrix(i, 8) & Chr(9)
            If Grid1.TextMatrix(i, 9) > " " Then
                s = s & Grid1.TextMatrix(i, 9) & Chr(9)
            Else
                s = s & "_" & Chr(9)
            End If
            s = s & Grid1.TextMatrix(i, 10) & Chr(9)
            If Grid1.TextMatrix(i, 11) > " " Then
                s = s & "Driver:" & Chr(9)
                s = s & Grid1.TextMatrix(i, 11)
            End If
            pgrid.AddItem s
            If Grid1.TextMatrix(i, 12) > " " Then
                s = "_" & Chr(9) & "_" & Chr(9) & "_" & Chr(9) & "." & Chr(9) & "_" & Chr(9) & "_" & Chr(9) & "_" & Chr(9)
                s = s & "Yard:" & Chr(9) & Grid1.TextMatrix(i, 12)
                pgrid.AddItem s
            End If
            If Grid1.TextMatrix(i, 13) > " " Then
                s = "_" & Chr(9) & "_" & Chr(9) & "_" & Chr(9) & "." & Chr(9) & "_" & Chr(9) & "_" & Chr(9) & "_" & Chr(9)
                s = s & "Loader:" & Chr(9) & Grid1.TextMatrix(i, 13)
                pgrid.AddItem s
            End If
        End If
    Next i
    If pgrid.Rows > 1 Then
        pgrid.FillStyle = flexFillRepeat
        For i = 1 To pgrid.Rows - 1
            If Val(pgrid.TextMatrix(i, 0)) = 0 Then
                pgrid.Row = i: pgrid.RowSel = i
                pgrid.Col = 1: pgrid.ColSel = pgrid.Cols - 1
                pgrid.CellBackColor = Label1.BackColor
            End If
        Next i
    End If
    s = "^Ticket|^Group|<Load|^Start|^Size|^TCode|<Contents|<Attn|<Notes"
    pgrid.FormatString = s
    pgrid.ColWidth(0) = 1000
    pgrid.ColWidth(1) = 1000
    pgrid.ColWidth(2) = 2000
    pgrid.ColWidth(3) = 1000
    pgrid.ColWidth(4) = 600
    pgrid.ColWidth(5) = 700
    pgrid.ColWidth(6) = 1000
    pgrid.ColWidth(7) = 700
    pgrid.ColWidth(8) = 4000
    'If pcode = "T10" Then
    '    rt = "Brenham Transport Notes - " & Combo2
    'End If
    'If pcode = "K10" Then
    '    rt = "Broken Arrow Transport Notes - " & Combo2
    'End If
    'If pcode = "A10" Then
    '    rt = "Sylacauga Transport Notes - " & Combo2
    'End If
    rt = truck_loc_name(pcode) & " Transport Notes"
    rh = "Ship Date - " & Combo2
    rf = "printed: " & Format(Now, "MM-dd-yyyy h:mm am/pm")
    hf = localAppDataPath & "\tempfile.htm"
    
    htdc(0) = "Yellow": gndc(0) = Label1.BackColor
    'htdc(1) = "Pink": gndc(1) = pcolor.BackColor
    Call htmlcolorgrid(Me, hf, pgrid, rt, rh, rf, "linen", "lemonchiffon", "white")
    If Len(Dir("c:\program files\internet explorer\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\internet explorer\iexplore.exe " & hf, vbNormalFocus)
        Exit Sub
    End If
    If Len(Dir("c:\program files\plus!\microsoft internet\iexplore.exe")) <> 0 Then
        i = Shell("C:\program files\plus!\microsoft internet\iexplore.exe " & hf, vbNormalFocus)
        Exit Sub
    End If

End Sub

Function branch_notes(pwhs As String)
    Dim ds As adodb.Recordset, s As String
    On Error GoTo vberror
    s = "select ISNULL(brnmess, '') AS brnmess from branches where gemmsid = '" & pwhs & "'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = ds!brnmess
    Else
        s = "No oracle branch"
    End If
    ds.Close
    branch_notes = s
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "branch_notes", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " branch_notes - Error Number: " & eno
        End
    End If
End Function

Function jobbing_run(runid As String) As String
    Dim ds As adodb.Recordset, s As String, jacct As String
    jacct = " "
    On Error GoTo vberror
    s = "select account from trailers where runid = " & runid
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        jacct = ds!account
    End If
    ds.Close
    jobbing_run = jacct
    Exit Function
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "jobbing_run", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " jobbing_run - Error Number: " & eno
        End
    End If
End Function

Sub match_runid_truckwo(runrow As Integer)
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String, mrt As String
    Dim runid As String, sloco As String, slocd As String, tno As String
    If runrow = 0 Then MsgBox "match row = 0"
    tno = Val(Right(Grid1.TextMatrix(runrow, 6), 1))
    runid = Grid1.TextMatrix(runrow, 1)
    sloco = Grid1.TextMatrix(runrow, 4)
    slocd = Grid1.TextMatrix(runrow, 5)
    mrt = " "
    'On Error GoTo vberror
    Set db = CreateObject("ADODB.Connection")
    db.Open Me.truckdb
    s = "select * from truckwo where r12ticket = '" & runid & "' and wostatus not in ('CANC', 'COMP')"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        mrt = ds!wonum
        Grid1.TextMatrix(runrow, 0) = ds!wonum
        Grid1.TextMatrix(runrow, 2) = Format(ds!wodate, "MM-dd-yyyy")
        Grid1.TextMatrix(runrow, 4) = ds!origin
        Grid1.TextMatrix(runrow, 5) = ds!Destination
        Grid1.TextMatrix(runrow, 6) = ds!trlno
        Grid1.TextMatrix(runrow, 7) = Format(ds!startime, "h:mm am/pm")
        Grid1.TextMatrix(runrow, 8) = ds!trlsize
        If Len(ds!eqnum) > 0 Then                                       'jv071116
            Grid1.TextMatrix(runrow, 9) = ds!eqnum
        Else                                                            'jv071116
            Grid1.TextMatrix(runrow, 9) = " "                           'jv071116
        End If                                                          'jv071116
        Grid1.TextMatrix(runrow, 10) = ds!contents
        Grid1.TextMatrix(runrow, 11) = ds!description
    End If
    ds.Close
    If mrt = " " Then
        s = "select * from truckwo where wostatus not in ('CANC', 'COMP')"
        s = s & " and wodate = '" & sdate & "'"
        s = s & " and origin = '" & truck_loc(sloco) & "'"
        s = s & " and destination = '" & truck_loc(slocd) & "'"
        s = s & " and trlno = " & tno
        Set ds = db.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            mrt = ds!wonum
            Grid1.TextMatrix(runrow, 0) = ds!wonum
            Grid1.TextMatrix(runrow, 2) = Format(ds!wodate, "MM-dd-yyyy")
            Grid1.TextMatrix(runrow, 4) = ds!origin
            Grid1.TextMatrix(runrow, 5) = ds!Destination
            Grid1.TextMatrix(runrow, 6) = ds!trlno
            Grid1.TextMatrix(runrow, 7) = Format(ds!startime, "h:mm am/pm")
            Grid1.TextMatrix(runrow, 8) = ds!trlsize
            If Len(ds!eqnum) > 0 Then                                       'jv071116
                Grid1.TextMatrix(runrow, 9) = ds!eqnum
            Else                                                            'jv071116
                Grid1.TextMatrix(runrow, 9) = " "                           'jv071116
            End If                                                          'jv071116
            Grid1.TextMatrix(runrow, 10) = ds!contents
            Grid1.TextMatrix(runrow, 11) = ds!description
        End If
        ds.Close
    End If
    
    If mrt = " " Then
        ano = jobbing_run(runid)
        If ano > " " Then
            s = "select * from truckwo where wostatus not in ('CANC', 'COMP')"
            s = s & " and wodate = '" & sdate & "'"
            's = s & " and origin = '" & truck_loc(sloco) & "'"
            s = s & " and destination in (select lcode from locations where jobaccount = '" & ano & "')"
            's = s & " and parentwo = 0"
            Set ds = db.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                mrt = ds!wonum
                Grid1.TextMatrix(runrow, 0) = ds!wonum
                Grid1.TextMatrix(runrow, 2) = Format(ds!wodate, "MM-dd-yyyy")
                Grid1.TextMatrix(runrow, 4) = ds!origin
                Grid1.TextMatrix(runrow, 5) = ds!Destination
                Grid1.TextMatrix(runrow, 6) = ds!trlno
                Grid1.TextMatrix(runrow, 7) = Format(ds!startime, "h:mm am/pm")
                Grid1.TextMatrix(runrow, 8) = ds!trlsize
                If Len(ds!eqnum) > 0 Then                                       'jv071116
                    Grid1.TextMatrix(runrow, 9) = ds!eqnum & " "
                Else                                                            'jv071116
                    Grid1.TextMatrix(runrow, 9) = " "                           'jv071116
                End If                                                          'jv071116
                Grid1.TextMatrix(runrow, 10) = ds!contents
                Grid1.TextMatrix(runrow, 11) = ds!description
            End If
            ds.Close
        End If
    End If
    db.Close
    If mrt > " " Then Call synch_truckwo_runids(runrow)
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "match_runid_truckwo", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " match_runid_truckwo - Error Number: " & eno
        End
    End If
End Sub

Sub imp_truckwo_runids()
    Dim d As adodb.Connection, d1 As adodb.Recordset, d2 As adodb.Recordset, d3 As adodb.Recordset
    Dim s As String, pkey As Long
    Dim lc As String, rc As String, qt As String, oc As String, pc As String, ocname As String
    Dim sqlx As String, mdate As String, i As Integer, dc As String, gflag As String
    Dim eno As Long, edesc As String, newwo As Boolean
    'On Error GoTo vberror
    mdate = sdate
    If Len(mdate) = 0 Then Exit Sub
    If IsDate(mdate) = False Then
        MsgBox "Invalid Date Format", vbOKOnly, "Sorry"
        Exit Sub
    End If
    Screen.MousePointer = 11
    Set d = CreateObject("ADODB.Connection")
    d.Open Me.truckdb
    sqlx = "select origin,destination,trlno,trlsize,startime,description,drvid,contents,wtype,wonum,eqnum from truckwo"
    sqlx = sqlx & " Where wodate = '" & mdate & "'"
    sqlx = sqlx & " and origin in (select lcode from locations where loctype = 'Plant')"
    sqlx = sqlx & " and destination in (select lcode from locations where loctype in ('Plant','Branch','Jobbing'))"
    sqlx = sqlx & " and wtype in ('Start', 'SameDay')"
    sqlx = sqlx & " and wostatus not in ('CANC', 'COMP')"
    sqlx = sqlx & " order by origin, startime"
    
    Set d1 = d.Execute(sqlx)
    If d1.BOF = False Then
        d1.MoveFirst
        Do Until d1.EOF
            newwo = True
            For i = 0 To Grid1.Rows - 1
                If Grid1.TextMatrix(i, 0) = d1!wonum Then
                    newwo = False
                    Call synch_truckwo_runids(i)
                    Exit For
                End If
            Next i
            If newwo = True Then
                oc = " ": ocname = " "
                s = "select * from locations where lcode = '" & d1!origin & "'"
                Set d2 = d.Execute(s)
                If d2.BOF = False Then
                    d2.MoveFirst
                    pc = ship_loc(d1!origin) 'mid$(d1(0), 2, 2)
                    lc = d2(1)
                End If
                d2.Close
                s = "select * from locations where lcode = '" & d1!Destination & "'"
                Set d2 = d.Execute(s)
                If d2.BOF = False Then
                    d2.MoveFirst
                    dc = ship_loc(d2(0))
                    rc = d2(1)
                End If
                d2.Close
                If d1!drvid > 0 Then
                    s = "select * from drivers where id = " & d1!drvid
                    Set d3 = d.Execute(s)
                    If d3.BOF = False Then
                        d3.MoveFirst
                        If d3!dlcode = "00000000" Then
                            oc = "*"
                            ocname = d3!driver
                        End If
                        If d3!drvpool = "Outside Carrier" Then
                            oc = "*"
                            ocname = d3!driver
                        End If
                    End If
                    d3.Close
                End If
                pkey = wd_seq("Oratkt", Form1.schdb)
                sqlx = "Insert into runs (id, loaded, destination, locname, trlno, trlsize, trldate, startime"
                sqlx = sqlx & ", pickup, oc) Values (" & pkey
                sqlx = sqlx & ", '" & Left(pc, 2) & "'"
                sqlx = sqlx & ", '" & Left(dc, 2) & "'"
                sqlx = sqlx & ", '" & Left(rc, 30) & "'"
                sqlx = sqlx & ", '#" & d1!trlno & "'"
                sqlx = sqlx & ", " & d1!trlsize
                sqlx = sqlx & ", '" & mdate & "'"
                If Len(d1!startime) > 0 Then
                    sqlx = sqlx & ", '" & d1!startime & "'"
                Else
                    sqlx = sqlx & ", '8:00 AM'"
                End If
                s = d1!description '& " " & d1!contents '& " " & d1!wonum
                If Left(s, Len(rc)) = rc Then s = Trim(Right(s, Len(s) - Len(rc)))
                s = Trim(ocname & " " & s)
                If Len(s) > 50 Then s = Left(s, 50)
                sqlx = sqlx & ", '" & s & "'"
                sqlx = sqlx & ", '" & oc & "')"
                Sdb.Execute sqlx
                sqlx = d1!wonum & Chr(9)
                sqlx = sqlx & pkey & Chr(9)
                sqlx = sqlx & Format(sdate, "MM-dd-yyyy") & Chr(9)
                sqlx = sqlx & Chr(9)
                sqlx = sqlx & d1!origin & Chr(9)
                sqlx = sqlx & d1!Destination & Chr(9)
                sqlx = sqlx & d1!trlno & Chr(9)
                sqlx = sqlx & d1!startime & Chr(9)
                sqlx = sqlx & d1!trlsize & Chr(9)
                If oc = "*" Then
                    sqlx = sqlx & "OC"
                Else
                    If Len(d1!eqnum) > 0 Then
                        sqlx = sqlx & d1!eqnum
                    End If
                End If
                sqlx = sqlx & Chr(9) 'd1!eqnum & Chr(9)
                sqlx = sqlx & d1!contents & Chr(9)
                'If ocname > " " Then s = s & ocname & " "
                sqlx = sqlx & s
                gflag = " "
                If dc = "16" Then gflag = "*"
                If dc = "15" Then gflag = "*"
                If Val(dc) = 0 Then gflag = "*"
                If Val(dc) = 1 And pc = "50" Then gflag = "*"           'jv080715
                If pc = "50" And dc = "51" Then gflag = "*"
                If pc = "51" And dc = "50" Then gflag = "*"
                If pc = "52" And dc = "50" Then gflag = "*"
                sqlx = sqlx & Chr(9) & Chr(9) & Chr(9) & gflag
                
                Grid1.AddItem sqlx
                '--------- Turn this on when go-live.
                sqlx = "Update truckwo set r12ticket = '" & pkey & "' Where wonum = " & d1!wonum
                d.Execute sqlx
            End If
            d1.MoveNext
        Loop
    End If
    d1.Close
    d.Close
    newwo = True
    For i = 0 To Grid1.Rows - 1
        If Grid1.TextMatrix(i, 4) = "50" And Grid1.TextMatrix(i, 6) = "OP" Then
            newwo = False
            Exit For
        End If
        If Grid1.TextMatrix(i, 4) = "50" And Grid1.TextMatrix(i, 6) = "ZO" Then     'jv091815
            newwo = False
            Exit For
        End If
    Next i
    If newwo = True Then
        pkey = wd_seq("Oratkt", Form1.schdb)
        sqlx = "Insert into runs (id, loaded, destination, locname, trlno, trlsize, trldate, startime"
        sqlx = sqlx & ", pickup, oc) Values (" & pkey
        sqlx = sqlx & ", 50"
        sqlx = sqlx & ", 1"
        sqlx = sqlx & ", 'Brenham'"
        'sqlx = sqlx & ", 'OP'"
        sqlx = sqlx & ", 'ZO'"                              'jv091815
        sqlx = sqlx & ", 60"
        sqlx = sqlx & ", '" & mdate & "'"
        sqlx = sqlx & ", '6:00 AM'"
        sqlx = sqlx & ", 'Brenham Order Pick'"
        sqlx = sqlx & ", ' ')"
        Sdb.Execute sqlx
        s = Chr(9)
        s = s & pkey & Chr(9)
        s = s & mdate & Chr(9)
        s = s & "ZOP" & Chr(9)
        s = s & "50" & Chr(9)
        s = s & "1" & Chr(9)
        's = s & "OP" & Chr(9)
        s = s & "ZO" & Chr(9)                               'jv091815
        s = s & "6:00 AM" & Chr(9)
        s = s & "60" & Chr(9)
        s = s & Chr(9)
        s = s & Chr(9)
        s = s & "Brenham Order Pick" & Chr(9)
        s = s & Chr(9) & Chr(9) & "*"
        Grid1.AddItem s
    End If
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "imp_truckwo_runids", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " imp_truckwo_runids - Error Number: " & eno
        End
    End If
End Sub

Sub synch_truckwo_runids(runrow As Integer)
    Dim d As adodb.Connection, d1 As adodb.Recordset, d2 As adodb.Recordset, d3 As adodb.Recordset
    Dim s As String, pkey As Long
    Dim lc As String, rc As String, qt As String, oc As String, pc As String, ocname As String
    Dim sqlx As String, mdate As String
    Dim eno As Long, edesc As String, newwo As Boolean
    Dim twonum As String, rtkt As String
    'On Error GoTo vberror
    If runrow = 0 Then MsgBox "synch row 0"
    twonum = Grid1.TextMatrix(runrow, 0)
    rtkt = Grid1.TextMatrix(runrow, 1)
    mdate = sdate
    If Len(mdate) = 0 Then Exit Sub
    If IsDate(mdate) = False Then
        MsgBox "Invalid Date Format", vbOKOnly, "Sorry"
        Exit Sub
    End If
    Screen.MousePointer = 11
    Set d = CreateObject("ADODB.Connection")
    d.Open Me.truckdb
    sqlx = "select origin,destination,trlno,trlsize,startime,description,drvid,contents,wtype,wonum,eqnum,wodate"
    sqlx = sqlx & " from truckwo"
    sqlx = sqlx & " Where wonum = '" & twonum & "'"
    'sqlx = sqlx & " and origin in (select lcode from locations where loctype = 'Plant')"
    'sqlx = sqlx & " and destination in (select lcode from locations where loctype in ('Plant','Branch','Jobbing'))"
    'sqlx = sqlx & " and wtype in ('Start', 'SameDay')"
    sqlx = sqlx & " and wostatus not in ('CANC', 'COMP')"
    'sqlx = sqlx & " order by origin, startime"
    
    Set d1 = d.Execute(sqlx)
    If d1.BOF = False Then
        d1.MoveFirst
        oc = " ": ocname = " "
        s = "select * from locations where lcode = '" & d1!origin & "'"
        Set d2 = d.Execute(s)
        If d2.BOF = False Then
            d2.MoveFirst
            pc = ship_loc(d1!origin) 'mid$(d1(0), 2, 2)
            lc = d2(1)
        End If
        d2.Close
        s = "select * from locations where lcode = '" & d1!Destination & "'"
        Set d2 = d.Execute(s)
        If d2.BOF = False Then
            d2.MoveFirst
            dc = ship_loc(d2(0))
            rc = d2(1)
        End If
        d2.Close
        If d1!drvid > 0 Then
            s = "select * from drivers where id = " & d1!drvid
            Set d3 = d.Execute(s)
            If d3.BOF = False Then
                d3.MoveFirst
                If d3!dlcode = "00000000" Then
                    oc = "*"
                    ocname = d3!driver
                End If
                If d3!drvpool = "Outside Carrier" Then
                    oc = "*"
                    ocname = d3!driver
                End If
            End If
            d3.Close
        End If
        pkey = Val(rtkt)
        sqlx = "Update runs set trlno = '#" & d1!trlno & "'"
        sqlx = sqlx & ", trlsize = " & d1!trlsize
        sqlx = sqlx & ", trldate = '" & Format(d1!wodate, "MM-dd-yyyy") & "'"
        If Len(d1!startime) > 0 Then
            sqlx = sqlx & ", startime = '" & d1!startime & "'"
        Else
            sqlx = sqlx & ", startime = '8:00 am'"
        End If
        s = d1!description
        If Left(s, Len(rc)) = rc Then s = Trim(Right(s, Len(s) - Len(rc)))
        s = Trim(ocname & " " & s)
        If Len(s) > 50 Then s = Left(s, 50)
        sqlx = sqlx & ", pickup = '" & s & "'"
        sqlx = sqlx & ", oc = '" & oc & "'"
        sqlx = sqlx & " Where id = " & pkey
        Sdb.Execute sqlx
        Grid1.TextMatrix(runrow, 2) = Format(d1!wodate, "MM-dd-yyyy")
        Grid1.TextMatrix(runrow, 6) = d1!trlno
        Grid1.TextMatrix(runrow, 7) = d1!startime
        Grid1.TextMatrix(runrow, 8) = d1!trlsize
        If oc = "*" Then
            Grid1.TextMatrix(runrow, 9) = "OC"
        Else
            If Len(d1!eqnum) > 0 Then
                Grid1.TextMatrix(runrow, 9) = d1!eqnum
            Else
                Grid1.TextMatrix(runrow, 9) = " "
            End If
        End If
        Grid1.TextMatrix(runrow, 10) = d1!contents
        Grid1.TextMatrix(runrow, 11) = s
        '--------- Turn this on when go-live.
        s = "Update truckwo set r12ticket = '" & pkey & "' Where wonum = " & d1!wonum
        d.Execute s
    Else
        sqlx = "Delete from runs where id = " & rtkt
        Sdb.Execute sqlx
    End If
    d1.Close
    d.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "imp_truckwo_runids", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " imp_truckwo_runids - Error Number: " & eno
        End
    End If
End Sub

Private Sub process_date()
    Dim i As Integer, gflag As Boolean
    Grid1.Redraw = False
    refresh_grid1
    DoEvents
    If Grid1.Rows > 1 Then
        For i = 1 To Grid1.Rows - 1
            If Val(Grid1.TextMatrix(i, 0)) = 0 Then
                Call match_runid_truckwo(i)
            End If
        Next i
    End If
    imp_truckwo_runids
    pcolor.Visible = False
    If Grid1.Rows > 1 Then
        Grid1.FillStyle = flexFillRepeat
        For i = 1 To Grid1.Rows - 1
            If Grid1.TextMatrix(i, 14) > " " Then
                Grid1.Row = i: Grid1.RowSel = i
                Grid1.Col = 1: Grid1.ColSel = Grid1.Cols - 1
                Grid1.CellBackColor = pcolor.BackColor
                pcolor.Visible = True
            End If
            s = ship_loc(Grid1.TextMatrix(i, 4))
            s = s & Grid1.TextMatrix(i, 3) & " "
            s = s & Space(12 - Len(s))
            s = s & Format(Grid1.TextMatrix(i, 7), "hh:mm")
            Grid1.TextMatrix(i, 15) = s
        Next i
        Grid1.Row = 1: Grid1.RowSel = 1
        Grid1.Col = 15: Grid1.ColSel = 15
        Grid1.Sort = 5
    End If
    Grid1.Redraw = True
End Sub

Function ship_loc(tloc As String) As String
    Dim i As Integer, s As String
    's = "16"
    s = tloc
    For i = 1 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 3) = tloc Then
            s = Grid2.TextMatrix(i, 1)
            Exit For
        End If
    Next i
    ship_loc = s
End Function

Function truck_loc_name(tloc As String)
    Dim i As Integer, s As String
    s = " "
    For i = 1 To Grid2.Rows - 1
        If Trim(Grid2.TextMatrix(i, 3)) = Trim(tloc) Then
            s = Grid2.TextMatrix(i, 2)
            Exit For
        End If
    Next i
    truck_loc_name = s
End Function


Function truck_loc(sloc As String) As String
    Dim i As Integer, s As String
    s = "0"
    For i = 1 To Grid2.Rows - 1
        If Grid2.TextMatrix(i, 1) = Val(sloc) Then
            s = Grid2.TextMatrix(i, 3)
            Exit For
        End If
    Next i
    truck_loc = s
End Function

Private Sub refresh_ship_lcodes()
    Dim ds As adodb.Recordset, s As String, i As Integer
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    Grid2.Clear: Grid2.Rows = 1: Grid2.Cols = 4
    s = "select plant, plantname, gemmsid from plants order by plant"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Grid2.AddItem "P" & Chr(9) & ds!plant & Chr(9) & ds!plantname & Chr(9) & ds!gemmsid
            ds.MoveNext
        Loop
    End If
    ds.Close
    s = "select branch, branchname, gemmsid from branches order by branch"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            If ds!branch = 1 Then                                                                   'jv080715
                Grid2.AddItem "B" & Chr(9) & ds!branch & Chr(9) & ds!branchname & Chr(9) & "001"    'jv080715
            Else
                If ds!branch = 47 Then                                                              'jv080715
                    Grid2.AddItem "B" & Chr(9) & ds!branch & Chr(9) & ds!branchname & Chr(9) & "047" 'jv080715
                Else
                    Grid2.AddItem "B" & Chr(9) & ds!branch & Chr(9) & ds!branchname & Chr(9) & ds!gemmsid
                End If
            End If
            ds.MoveNext
        Loop
    End If
    ds.Close
    Grid2.FormatString = "^Type|^Code|<Name|^Oracle"
    Grid2.ColWidth(0) = 1000
    Grid2.ColWidth(1) = 1000
    Grid2.ColWidth(2) = 2500
    Grid2.ColWidth(3) = 1000
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "refresh_ship_lcodes", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_ship_lcodes - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_dates()
    Dim ds As adodb.Recordset, s As String, i As Integer
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    Combo2.Clear
    s = "select distinct trldate from runs order by trldate"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo2.AddItem Format(ds!trldate, "MM-dd-yyyy")
            ds.MoveNext
        Loop
    End If
    ds.Close
    Combo2.ListIndex = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "refresh_dates", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_dates - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_mlists()
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String, i As Integer
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    Set db = CreateObject("ADODB.Connection")
    Combo1.Clear: Combo9.Clear: Combo10.Clear
    List1.Clear: List9.Clear: List10.Clear
    db.Open Me.truckdb
    'origin and destination 1&2
    s = "select lcode,location from locations where location > ' ' order by location"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            List1.AddItem ds!lcode
            Combo1.AddItem ds!location
            ds.MoveNext
        Loop
    End If
    ds.Close
    'BB Trailer Codes
    Combo9.AddItem " ": List9.AddItem " "
    s = "select * from valuelists where listname = 'trlcode' order by listdisplay"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            List9.AddItem ds!listreturn
            Combo9.AddItem ds!listdisplay
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    'contents
    Combo10.AddItem " ": List10.AddItem " "
    s = "select * from valuelists where listname = 'contents' order by listdisplay"
    Set ds = db.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            List10.AddItem ds!listreturn
            Combo10.AddItem ds!listdisplay
            ds.MoveNext
        Loop
    End If
    ds.Close: db.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "refresh_mlists", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_mlists - Error Number: " & eno
        End
    End If
End Sub

Private Sub refresh_grid1()
    Dim db As adodb.Connection, ds As adodb.Recordset, s As String, i As Integer
    Dim eno As Long, edesc As String, gflag As String
    On Error GoTo vberror
    Grid1.Clear: Grid1.Rows = 1: Grid1.Cols = 16
    Set db = CreateObject("ADODB.Connection")
    s = "select * from runs where trldate = '" & sdate & "'"
    Set ds = Sdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = Chr(9) & ds!id & Chr(9)
            s = s & Format(ds!trldate, "MM-dd-yyyy") & Chr(9)
            s = s & Chr(9)
            s = s & ds!loaded & Chr(9)
            s = s & ds!Destination & Chr(9)
            s = s & ds!trlno & Chr(9)
            s = s & Format(ds!startime, "h:mm am/pm") & Chr(9)
            s = s & ds!trlsize & Chr(9)
            If ds!oc = "*" Then s = s & "OC"
            s = s & Chr(9)
            s = s & Chr(9)
            s = s & ds!pickup & Chr(9)
            s = s & ds!yardnote & Chr(9)
            s = s & ds!loadnote
            gflag = " "
            If ds!Destination = "16" Then gflag = "*"
            If ds!Destination = "15" Then gflag = "*"
            If Val(ds!Destination) = 0 Then gflag = "*"
            If Val(ds!Destination) = 1 And ds!loaded = "50" Then gflag = "*"        'jv080715
            If ds!loaded = "50" And ds!Destination = "51" Then gflag = "*"
            If ds!loaded = "51" And ds!Destination = "50" Then gflag = "*"
            If ds!loaded = "52" And ds!Destination = "50" Then gflag = "*"
            s = s & Chr(9) & gflag
            Grid1.AddItem s
            ds.MoveNext
        Loop
    End If
    ds.Close
    If Grid1.Rows > 1 Then
        Set ds = Sdb.Execute("select * from trgroups")
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                For i = 1 To Grid1.Rows - 1
                    If ds!run1 = Val(Grid1.TextMatrix(i, 1)) Then Grid1.TextMatrix(i, 3) = ds!groupcode
                    If ds!run2 = Val(Grid1.TextMatrix(i, 1)) Then Grid1.TextMatrix(i, 3) = ds!groupcode
                    If ds!run3 = Val(Grid1.TextMatrix(i, 1)) Then Grid1.TextMatrix(i, 3) = ds!groupcode
                    If ds!run4 = Val(Grid1.TextMatrix(i, 1)) Then Grid1.TextMatrix(i, 3) = ds!groupcode
                Next i
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    If Grid1.Rows > 1 Then
        db.Open Me.truckdb
        For i = 1 To Grid1.Rows - 1
            s = "select * from truckwo where r12ticket = " & Grid1.TextMatrix(i, 1) & " and parentwo = 0"
            Set ds = db.Execute(s)
            If ds.BOF = False Then
                Grid1.TextMatrix(i, 0) = ds!wonum
                Grid1.TextMatrix(i, 4) = ds!origin
                Grid1.TextMatrix(i, 5) = ds!Destination
                Grid1.TextMatrix(i, 6) = ds!trlno
                Grid1.TextMatrix(i, 7) = Format(ds!startime, "h:mm am/pm")
                Grid1.TextMatrix(i, 8) = ds!trlsize
                If Len(ds!eqnum) > 0 Then                               'jv071116
                    Grid1.TextMatrix(i, 9) = ds!eqnum
                Else                                                    'jv071116
                    Grid1.TextMatrix(i, 9) = " "                        'jv071116
                End If                                                  'jv071116
                Grid1.TextMatrix(i, 10) = ds!contents
                Grid1.TextMatrix(i, 11) = ds!description
            End If
            ds.Close
        Next i
        db.Close
        Grid1.RowSel = Grid1.Row
        Grid1.Col = 2: Grid1.ColSel = 5
        Grid1.Sort = 5
        Grid1.Col = 6
    End If
    s = "^WONum|^Ticket|^Date|^Group|^Origin|^Destination|^#|<Start|^Size|^BB#|<Contents|<Schedule Note|<Yard Notes|<Loader Notes|^NBO|<"
    Grid1.FormatString = s
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1000
    Grid1.ColWidth(2) = 1200
    Grid1.ColWidth(3) = 1200
    Grid1.ColWidth(4) = 1000
    Grid1.ColWidth(5) = 1200
    Grid1.ColWidth(6) = 500
    Grid1.ColWidth(7) = 1300
    Grid1.ColWidth(8) = 700
    Grid1.ColWidth(9) = 700
    Grid1.ColWidth(10) = 1400
    Grid1.ColWidth(11) = 2000
    Grid1.ColWidth(12) = 2000
    Grid1.ColWidth(13) = 2000
    Grid1.ColWidth(14) = 600
    Grid1.ColWidth(15) = 0 '2000
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, "refresh_grid1", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_grid1 - Error Number: " & eno
        End
    End If
End Sub

Private Sub a10repall_Click()
    Call listreport("A10", "*")
End Sub

Private Sub a10repbo_Click()
    Call listreport("A10", " ")
End Sub

Private Sub addtkt_Click()
    Dim ds As adodb.Recordset, sqlx As String, pname As String
    Dim pc As String, dc As String, stime As String, lc As String, pkey As Long, tno As String
    On Error GoTo vberror
    pc = InputBox$("Plant Code ", "Plant Code", "50")
    If Len(pc) = 0 Then Exit Sub
    dc = InputBox$("Branch Code ", "Branch Code", "28")
    If Len(dc) = 0 Then Exit Sub
    stime = InputBox$("Start Time", "Start Time", Format$(Now, "h:mm am/pm"))
    If Len(stime) = 0 Then Exit Sub
    If IsDate(stime) = False Then
        MsgBox "Invalid Time Format", vbOKOnly, "Sorry"
        Exit Sub
    End If
    sqlx = "select * from plants where plant = " & pc
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid Plant Code " & pc & " used.", vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    Else
        pname = ds!plantname
    End If
    ds.Close
    sqlx = "select * from branches where branch = " & dc
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid Branch Code " & dc & " used.", vbOKOnly, "Sorry"
        ds.Close
        Exit Sub
    End If
    lc = ds!branchname
    ds.Close
    tno = InputBox("Trailer #:", "Trailer #...", "1")
    If Len(tno) = 0 Then Exit Sub
    pkey = wd_seq("Oratkt", Form1.schdb)
    sqlx = "Insert into runs (id, loaded, destination, locname, trlno, trlsize, trldate, startime, pickup, oc)"
    sqlx = sqlx & " Values (" & pkey
    sqlx = sqlx & ", '" & pc & "'"
    sqlx = sqlx & ", '" & dc & "'"
    sqlx = sqlx & ", '" & lc & "'"
    sqlx = sqlx & ", '" & tno & "'"
    sqlx = sqlx & ", 32"
    sqlx = sqlx & ", '" & sdate & "'"
    sqlx = sqlx & ", '" & stime & "'"
    sqlx = sqlx & ", 'Added'"
    sqlx = sqlx & ", ' ')"
    Sdb.Execute sqlx
    
    sqlx = Chr(9)                                               'truckwo
    sqlx = sqlx & pkey & Chr(9)                                 'ticket
    sqlx = sqlx & Format(Combo2, "MM-dd-yyyy") & Chr(9)         'trldate
    sqlx = sqlx & " " & Chr(9)                                  'group
    sqlx = sqlx & pc & Chr(9)                                   'plant
    sqlx = sqlx & dc & Chr(9)                                   'branch
    sqlx = sqlx & tno & Chr(9)                                   'trlno
    sqlx = sqlx & Format$(stime, "h:mm am/pm") & Chr(9)         'start
    sqlx = sqlx & "32" & Chr(9)                                 'size
    sqlx = sqlx & "" & Chr(9)                                   'tcode
    sqlx = sqlx & "" & Chr(9)                                   'contents
    sqlx = sqlx & "" & Chr(9)                                   'driver note
    sqlx = sqlx & "" & Chr(9)                                   'yard note
    sqlx = sqlx & "" & Chr(9)                                   'loader note
    
    Grid1.AddItem sqlx
    'Call refresh_grid
    'outfile = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "addtkt_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " addtkt_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub clrdate_Click()
    Dim sqlx As String
    On Error GoTo vberror
    If MsgBox("Clear schedule for " & Combo2, vbOKCancel, "Are you sure?") = vbCancel Then
        Exit Sub
    End If
    sqlx = "delete from runs where trldate = '" & Combo2 & "'"
    Sdb.Execute sqlx
    Combo2.RemoveItem Combo2.ListIndex
    If Combo2.ListCount > 0 Then Combo2.ListIndex = 0
    'outfile = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "clrdate_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " clrdate_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Combo1_Click()
    List1.ListIndex = Combo1.ListIndex
End Sub

Private Sub Combo10_Click()
    List10.ListIndex = Combo10.ListIndex
End Sub

Private Sub Combo2_Click()
    sdate = Combo2
    'refresh_grid1
    
    process_date
    DoEvents
    Grid1_RowColChange
    
End Sub

Private Sub Combo9_Click()
    List9.ListIndex = Combo9.ListIndex
End Sub

Private Sub Command1_Click()
    Dim db As DAO.Database, ds As adodb.Recordset, s As String, zid As Long
    Dim i As Integer, eno As Long, edesc As String
    'On Error GoTo vberror
    i = Grid1.Row
    If i = 0 Then Exit Sub
    If (Grid1.TextMatrix(i, 8) <> Text8 Or Grid1.TextMatrix(i, 9) <> Text9) And Val(Text0) > 0 Then
        Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, Me.truckdb)
        s = "update truckwo set trlsize = " & Val(Text8) & ", eqnum = '" & Text9 & "'"
        s = s & ", updatedby = '" & Form1.userid & "'"
        s = s & ", lastchange = '" & Format(Now, "m-d-yyyy h:mm am/pm") & "'"
        s = s & " where wonum = " & Text0 & " or parentwo = " & Text0
        'MsgBox s
        db.Execute s
        db.Close
        'MsgBox s
        Grid1.TextMatrix(i, 8) = Text8
        Grid1.TextMatrix(i, 9) = Text9
    End If
    If (Grid1.TextMatrix(i, 9) <> Text9 Or Grid1.TextMatrix(i, 10) <> Text10) And Val(Text0) > 0 Then
        Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, Me.truckdb)
        s = "update truckwo set contents = '" & Text10 & "'"
        s = s & ", updatedby = '" & Form1.userid & "'"
        s = s & ", lastchange = '" & Format(Now, "m-d-yyyy h:mm am/pm") & "'"
        s = s & " where wonum = " & Text0
        db.Execute s
        db.Close
        'MsgBox s
        Grid1.TextMatrix(i, 10) = Text10
    End If
    If Text14.Visible = True Then
        'Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, Me.shipdb)
        s = "select * from branches where gemmsid = '" & Text5 & "'"
        Set ds = Sdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            'ds.Edit
            'ds!brnmess = Text14
            'ds.Update
            s = "Update branches set brnmess = '" & Text14 & "' Where gemmsid = '" & Text5 & "'"
            Sdb.Execute s
        End If
        ds.Close ': db.Close
    End If
    'If Grid1.TextMatrix(i, 8) <> Text8 Or Grid1.TextMatrix(Grid1.Row, 12) <> Text12 Or Grid1.TextMatrix(i, 13) <> Text13 Then
        'Set db = OpenDatabase(mysqldev, dbcdrivernoprompt, False, Me.shipdb)
        s = "select * from runs where id = " & Text1
        Set ds = Sdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            'ds.Edit
            'ds!trlsize = Val(Text8)
            'ds!yardnote = Text12
            'ds!loadnote = Text13
            'ds.Update
            s = "Update runs set trlsize = " & Val(Text8) & ", yardnote = '" & Text12 & "'"
            s = s & ", loadnote = '" & Text13 & "' Where id = " & Text1
            Sdb.Execute s
            Grid1.TextMatrix(i, 8) = Text8
            Grid1.TextMatrix(i, 12) = Text12
            Grid1.TextMatrix(i, 13) = Text13
        End If
        ds.Close ': db.Close
    'End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description
    Call vb_elog(eno, edesc, Me.Name, Command1.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " " & Command1.Caption & "_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command2_Click()
    process_date
    DoEvents
    Grid1_RowColChange
End Sub

Private Sub Command3_Click()
    Dim cfile As String, i
    cfile = "U:\work.xml"
    'Call xlsgrid("U:\work.csv", Grid1)
    Call xmlgrid("U:\work.xml", Grid1, True)
    
    Call OpenFileInExcel(cfile)
End Sub

Private Sub deltkt_Click()
    Dim sqlx As String
    On Error GoTo vberror
    If MsgBox("Clear Ticket " & Grid1.TextMatrix(Grid1.Row, 1), vbYesNo + vbQuestion, "Are you sure?") = vbNo Then
        Exit Sub
    End If
    sqlx = "delete from runs where id = " & Grid1.TextMatrix(Grid1.Row, 1)
    Sdb.Execute sqlx
    If Grid1.Rows > 2 Then
        Grid1.RemoveItem Grid1.Row
    Else
        Grid1.Rows = 1
    End If
    'outfile = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "deltkt_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " deltkt_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Load()
    'Me.shipdb = "ODBC;DATABASE=WDShip;DSN=wdship"
    'Me.shipdb = "ODBC;DATABASE=WDship;UID=bbcship500;PWD=brenham500;DSN=wdship500"
    ''me.truckdb = "ODBC;DATABASE=WDTruck;DSN=wdtruck"
    'Me.truckdb = "ODBC;DATABASE=WDTruck;uid=bbctruck500;pwd=brenham500;DSN=truckwo"
    
    Me.shipdb = Form1.shipdb
    Me.truckdb = Form1.schdb
    
    Grid1.Font = "Arial": Grid1.FontSize = 9: Grid1.FontBold = True
    
    refresh_ship_lcodes
    refresh_mlists
    refresh_dates
    'refresh_grid1
    Grid1_RowColChange
End Sub

Private Sub Form_Resize()
    Grid1.Width = Me.Width - 120
    pgrid.Width = Me.Width - 120
End Sub

Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu edmenu
End Sub

Private Sub Grid1_RowColChange()
    i = Grid1.Row
    'Text1 = "": Text2 = "": Text3 = "": Text4 = "": Text5 = ""
    'Text6 = "": Text7 = "": Text8 = "": Text9 = "": Text0 = ""
    'Text10 = "": Text11 = "": Text12 = "": Text13 = ""
    Label0 = Grid1.TextMatrix(0, 0)
    Label1 = Grid1.TextMatrix(0, 1)
    Label2 = Grid1.TextMatrix(0, 2)
    Label3 = Grid1.TextMatrix(0, 3)
    Label4 = Grid1.TextMatrix(0, 4)
    Label5 = Grid1.TextMatrix(0, 5)
    Label6 = Grid1.TextMatrix(0, 6)
    Label7 = Grid1.TextMatrix(0, 7)
    Label8 = Grid1.TextMatrix(0, 8)
    Label9 = Grid1.TextMatrix(0, 9)
    Label10 = Grid1.TextMatrix(0, 10)
    Label11 = Grid1.TextMatrix(0, 11)
    Label12 = Grid1.TextMatrix(0, 12)
    Label13 = Grid1.TextMatrix(0, 13)
    
    Text0 = Grid1.TextMatrix(i, 0)
    Text1 = Grid1.TextMatrix(i, 1)
    Text2 = Grid1.TextMatrix(i, 2)
    Text3 = Grid1.TextMatrix(i, 3)
    Text4 = Grid1.TextMatrix(i, 4)
    Text5 = Grid1.TextMatrix(i, 5)
    Text6 = Grid1.TextMatrix(i, 6)
    Text7 = Grid1.TextMatrix(i, 7)
    Text8 = Grid1.TextMatrix(i, 8)
    Text9 = Grid1.TextMatrix(i, 9)
    Text10 = Grid1.TextMatrix(i, 10)
    Text11 = Grid1.TextMatrix(i, 11)
    Text12 = Grid1.TextMatrix(i, 12)
    Text13 = Grid1.TextMatrix(i, 13)
    
End Sub

Private Sub impsched_Click()
    Dim mdate As String, i As Integer
    mdate = InputBox$("Please input date to import.", "Schedule Date", Format(DateAdd("d", 1, Now), "MM-dd-yyyy"))
    If Len(mdate) = 0 Then Exit Sub
    If IsDate(mdate) = False Then
        MsgBox "Invalid Date Format", vbOKOnly, "Sorry"
        Exit Sub
    End If
    mdate = Format(mdate, "MM-dd-yyyy")
    For i = 0 To Combo2.ListCount - 1
        If Combo2.List(i) = mdate Then
            Combo2.ListIndex = i
            Exit Sub
        End If
    Next i
    Combo2.AddItem mdate
    For i = 0 To Combo2.ListCount - 1
        If Combo2.List(i) = mdate Then
            Combo2.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub k10repall_Click()
    Call listreport("K10", "*")
End Sub

Private Sub k10repbo_Click()
    Call listreport("K10", " ")
End Sub

Private Sub List10_Click()
    Text10 = List10
End Sub

Private Sub List9_Click()
    Text9 = List9
End Sub

Private Sub t10repall_Click()
    Call listreport("T10", "*")
End Sub

Private Sub t10repbo_Click()
    Call listreport("T10", " ")
End Sub

Private Sub Text10_Change()
    Dim i As Integer
    For i = 0 To List10.ListCount - 1
        If List10.List(i) = Text10 Then
            Combo10.ListIndex = i
            Exit Sub
        End If
    Next i
    Combo10.ListIndex = 0
End Sub

Private Sub Text4_Change()
    Dim i As Integer
    sorg = ".."
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = Text4 Then
            sorg = Combo1.List(i)
            Exit For
        End If
    Next i
End Sub

Private Sub Text5_Change()
    Dim i As Integer, s As String
    sdest = ".."
    For i = 0 To List1.ListCount - 1
        If List1.List(i) = Text5 Then
            sdest = Combo1.List(i)
            Exit For
        End If
    Next i
    
    s = branch_notes(Text5)
    If s = "No oracle branch" Then
        Text14 = ""
        Text14.Visible = False
    Else
        Text14 = s
        Text14.Visible = True
    End If
End Sub

Private Sub Text9_Change()
    Dim i As Integer
    For i = 0 To List9.ListCount - 1
        If List9.List(i) = Text9 Then
            Combo9.ListIndex = i
            Exit Sub
        End If
    Next i
    Combo9.ListIndex = 0
End Sub

Private Sub tktstatus_Click()
    runstatus.wokey = Grid1.TextMatrix(Grid1.Row, 0)
    runstatus.plantkey = Grid1.TextMatrix(Grid1.Row, 4)
    runstatus.runkey = Grid1.TextMatrix(Grid1.Row, 1)
    runstatus.Show
End Sub

Private Sub wksdriver_Click()
    Call cre8wks("Driver")
    'Call cre8xml("Driver")
End Sub

Private Sub wksload_Click()
    Call cre8wks("Loader")
    'Call cre8xml("Loader")
End Sub

Private Sub wksyard_Click()
    Call cre8wks("Yard")
    'Call cre8xml("Yard")
End Sub
