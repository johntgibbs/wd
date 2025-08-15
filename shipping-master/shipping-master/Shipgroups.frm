VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Editgroups 
   Caption         =   "Edit Trailer Groups"
   ClientHeight    =   13125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12195
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   13125
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Attach Notes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   54
      Top             =   6000
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid noteGrid 
      Height          =   1815
      Left            =   600
      TabIndex        =   53
      Top             =   4080
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3201
      _Version        =   327680
      Rows            =   5
      BackColorFixed  =   65535
   End
   Begin MSFlexGridLib.MSFlexGrid cgrid 
      Height          =   4095
      Left            =   0
      TabIndex        =   50
      Top             =   8400
      Visible         =   0   'False
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   7223
      _Version        =   327680
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Print Check Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   49
      Top             =   6480
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid altgrid 
      Height          =   1335
      Left            =   4440
      TabIndex        =   48
      Top             =   6000
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2355
      _Version        =   327680
      BackColor       =   16777215
      ForeColor       =   16711680
      BackColorFixed  =   12648384
      BackColorBkg    =   -2147483633
      ScrollTrack     =   -1  'True
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cw 
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
      Index           =   3
      Left            =   6120
      TabIndex        =   47
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cw 
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
      Index           =   2
      Left            =   6120
      TabIndex        =   46
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cw 
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
      Index           =   1
      Left            =   6120
      TabIndex        =   45
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cw 
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
      Index           =   0
      Left            =   6120
      TabIndex        =   44
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Insert Alternate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   43
      Top             =   5400
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid ogrid 
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8070
      _Version        =   327680
      Cols            =   15
      BackColor       =   16777215
      ForeColor       =   8388608
      BackColorFixed  =   16777152
      BackColorSel    =   255
      BackColorBkg    =   -2147483633
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      Appearance      =   0
   End
   Begin VB.TextBox w 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   10320
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox w 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   10320
      TabIndex        =   41
      Text            =   "Text1"
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox w 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   10320
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox w 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   10320
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Split SKU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton pw 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   37
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton pw 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   36
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton pw 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   35
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton pw 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   34
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox ord 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   3
      Left            =   3600
      TabIndex        =   33
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox ord 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   2
      Left            =   3600
      TabIndex        =   32
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox ord 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   1
      Left            =   3600
      TabIndex        =   31
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox ord 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   3600
      TabIndex        =   30
      Top             =   360
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid wgrid 
      Height          =   4695
      Left            =   9840
      TabIndex        =   23
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   8281
      _Version        =   327680
      Cols            =   4
      BackColor       =   16777215
      ForeColor       =   4194368
      BackColorFixed  =   16761087
      BackColorSel    =   8388736
      BackColorBkg    =   16761087
      Appearance      =   0
   End
   Begin VB.Label gdate 
      Caption         =   "gdate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   52
      Top             =   6120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label gplant 
      Caption         =   "Gplant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   51
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1320
      TabIndex        =   29
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label wname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3960
      TabIndex        =   28
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label wname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3960
      TabIndex        =   27
      Top             =   840
      Width           =   855
   End
   Begin VB.Label wname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3960
      TabIndex        =   26
      Top             =   600
      Width           =   855
   End
   Begin VB.Label wname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3960
      TabIndex        =   25
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   24
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label ttot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3120
      TabIndex        =   22
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label ttot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3120
      TabIndex        =   21
      Top             =   840
      Width           =   495
   End
   Begin VB.Label ttot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3120
      TabIndex        =   20
      Top             =   600
      Width           =   495
   End
   Begin VB.Label ttot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3120
      TabIndex        =   19
      Top             =   360
      Width           =   495
   End
   Begin VB.Label br 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   8880
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label br 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   8880
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label br 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   8880
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label br 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8880
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label size 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2640
      TabIndex        =   14
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label size 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2640
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.Label size 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2640
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
   Begin VB.Label size 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2640
      TabIndex        =   11
      Top             =   360
      Width           =   495
   End
   Begin VB.Label trlcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   720
      TabIndex        =   10
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label trlcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   720
      TabIndex        =   9
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label trlcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   720
      TabIndex        =   8
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label trlcode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   720
      TabIndex        =   7
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label runid 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   8400
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label runid 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   8400
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label runid 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label runid 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8400
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Editgroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rf As String
Private Sub refresh_alts()
    Dim ds As ADODB.Recordset, sqlx As String, bstr As String
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    bstr = br(0)
    If Val(br(1)) > 0 Then bstr = bstr & "," & br(1)
    If Val(br(2)) > 0 Then bstr = bstr & "," & br(2)
    If Val(br(3)) > 0 Then bstr = bstr & "," & br(3)
    sqlx = "Select branch,brorders.sku,fgunit,fgdesc"
    sqlx = sqlx & " from brorders,skumast"
    sqlx = sqlx & " Where plant = " & gplant
    sqlx = sqlx & " and orddate = '" & gdate & "'"
    sqlx = sqlx & " and branch in (" & bstr & ")"
    sqlx = sqlx & " and altflag = 'Y' and brorders.sku = skumast.sku"
    sqlx = sqlx & " order by branch,brorders.sku"
    altgrid.Visible = False: altgrid.Clear
    altgrid.Cols = 3: altgrid.Rows = 1
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            sqlx = Format(ds!branch, "00") & Chr$(9)
            sqlx = sqlx & ds!sku & Chr$(9)
            sqlx = sqlx & " " & ds!fgunit & " " & ds!fgdesc
            altgrid.AddItem sqlx
            ds.MoveNext
        Loop
    End If
    ds.Close
    altgrid.FormatString = "^Branch|^SKU|<Alternate Products"
    altgrid.ColWidth(0) = 700
    altgrid.ColWidth(1) = 500
    altgrid.ColWidth(2) = 4000
    altgrid.Visible = True
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_alts", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_alts - Error Number: " & eno
        End
    End If
End Sub
Private Sub refresh_ogrid()
    Dim ds As ADODB.Recordset, sqlx As String, hflag As Boolean
    Dim eno As Long, edesc As String
    Screen.MousePointer = 11
    On Error GoTo vberror
    ogrid.Redraw = False
    sqlx = "select * from trgroups where groupcode = '" & Combo1 & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        runid(0) = ds!run1: runid(1) = ds!run2
        runid(2) = ds!run3: runid(3) = ds!run4
    End If
    ds.Close
    sqlx = "select id,groupitems.sku,fgunit,fgdesc,qty1,whs1,qty2,whs2,qty3,whs3,qty4,whs4"
    sqlx = sqlx & " from groupitems,skumast where groupcode = '" & Label1 & "'"
    sqlx = sqlx & " and groupitems.sku = skumast.sku order by groupitems.sku"
    Set ds = Sdb.Execute(sqlx)
    ogrid.Rows = 1
    If ds.BOF = False Then
        ds.MoveFirst
        'wgrid.Col = 1
        Do Until ds.EOF
            sqlx = ds!id & Chr$(9)
            sqlx = sqlx & ds!sku & Chr$(9)
            sqlx = sqlx & " " & ds!fgunit & " " & ds!fgdesc & Chr$(9)
            If ds!qty1 > 0 Then sqlx = sqlx & ds!qty1
            sqlx = sqlx & Chr$(9)
            If ds!whs1 > 0 Then
                'wgrid.Row = ds!whs1
                sqlx = sqlx & wgrid.TextMatrix(ds!whs1, 1) & Chr$(9)
            Else
                sqlx = sqlx & "..." & Chr$(9)
            End If
            If ds!qty2 > 0 Then sqlx = sqlx & ds!qty2
            sqlx = sqlx & Chr$(9)
            If ds!whs2 > 0 Then
                'wgrid.Row = ds!whs2
                sqlx = sqlx & wgrid.TextMatrix(ds!whs2, 1) & Chr$(9)
            Else
                sqlx = sqlx & "..." & Chr$(9)
            End If
            If ds!qty3 > 0 Then sqlx = sqlx & ds!qty3
            sqlx = sqlx & Chr$(9)
            If ds!whs3 > 0 Then
                'wgrid.Row = ds!whs3
                sqlx = sqlx & wgrid.TextMatrix(ds!whs3, 1) & Chr$(9)
            Else
                sqlx = sqlx & "..." & Chr$(9)
            End If
            If ds!qty4 > 0 Then sqlx = sqlx & ds!qty4
            sqlx = sqlx & Chr$(9)
            If ds!whs4 > 0 Then
                'wgrid.Row = ds!whs4
                sqlx = sqlx & wgrid.TextMatrix(ds!whs4, 1) & Chr$(9)
            Else
                sqlx = sqlx & "..." & Chr$(9)
            End If
            sqlx = sqlx & ds!whs1 & Chr$(9)
            sqlx = sqlx & ds!whs2 & Chr$(9)
            sqlx = sqlx & ds!whs3 & Chr$(9)
            sqlx = sqlx & ds!whs4 & Chr$(9)
            ogrid.AddItem sqlx
            
            ds.MoveNext
        Loop
    End If
    ds.Close
    ogrid.FormatString = "ID|^SKU|<Product|^1|^Whs|^2|^Whs|^3|^Whs|^4|^Whs"
    ogrid.ColWidth(0) = 1
    ogrid.ColWidth(1) = 600: ogrid.ColWidth(2) = 3500
    ogrid.ColWidth(3) = 600: ogrid.ColWidth(4) = 600
    ogrid.ColWidth(5) = 600: ogrid.ColWidth(6) = 600
    ogrid.ColWidth(7) = 600: ogrid.ColWidth(8) = 600
    ogrid.ColWidth(9) = 600: ogrid.ColWidth(10) = 600
    ogrid.ColWidth(11) = 0: ogrid.ColWidth(12) = 0
    ogrid.ColWidth(13) = 0: ogrid.ColWidth(14) = 0
    If ogrid.Rows > 1 Then
        ogrid.FillStyle = flexFillRepeat
        For i = 1 To ogrid.Rows - 1
            If hflag Then
                ogrid.Row = i: ogrid.RowSel = i
                ogrid.Col = 0: ogrid.ColSel = ogrid.Cols - 1
                ogrid.CellForeColor = ogrid.ForeColorFixed
                ogrid.CellBackColor = ogrid.BackColorFixed 'Label3.BackColor
            End If
            hflag = Not hflag
        Next i
        
        
        ogrid.Col = 3: ogrid.Row = 1: ogrid.RowSel = ogrid.Rows - 1
        ogrid.CellBackColor = Label3.BackColor 'ogrid.BackColorFixed '&HFFFFC0
        ogrid.Col = 4: ogrid.Row = 1: ogrid.RowSel = ogrid.Rows - 1
        ogrid.CellBackColor = Label3.BackColor 'ogrid.BackColorFixed '&HFFFFC0
        ogrid.Col = 7: ogrid.Row = 1: ogrid.RowSel = ogrid.Rows - 1
        ogrid.CellBackColor = Label3.BackColor 'ogrid.BackColorFixed '&HFFFFC0
        ogrid.Col = 8: ogrid.Row = 1: ogrid.RowSel = ogrid.Rows - 1
        ogrid.CellBackColor = Label3.BackColor 'ogrid.BackColorFixed '&HFFFFC0
        ogrid.Row = 1
    End If
    ogrid.Redraw = True
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "refresh_ogrid", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " refresh_ogrid - Error Number: " & eno
        End
    End If
End Sub
Private Sub calc_whstots()
    Dim ds As ADODB.Recordset, sqlx As String, i As Integer
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    Screen.MousePointer = 11
    wgrid.Col = 3
    For i = 1 To wgrid.Rows - 1
        wgrid.TextMatrix(i, 3) = " "
    Next i
    sqlx = "Select whs_num,avail from whstotals where sku = '" & Label2 & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            wgrid.TextMatrix(ds!whs_num, 3) = Val(wgrid.TextMatrix(ds!whs_num, 3)) + ds!avail
            ds.MoveNext
        Loop
    End If
    ds.Close
    Screen.MousePointer = 0
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "calc_whstots", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " calc_whstots - Error Number: " & eno
        End
    End If
End Sub
Private Sub calc_trltots()
    Dim i As Integer, k As Integer
    For i = 0 To 3
        ttot(i) = ""
    Next i
    For i = 1 To wgrid.Rows - 1
        wgrid.TextMatrix(i, 2) = ""
    Next i
    For i = 1 To ogrid.Rows - 1
        If ogrid.TextMatrix(i, 4) <> "..." Then
            ttot(0) = Val(ttot(0)) + Val(ogrid.TextMatrix(i, 3))
            For k = 1 To wgrid.Rows - 1
                If ogrid.TextMatrix(i, 4) = wgrid.TextMatrix(k, 1) Then
                    wgrid.TextMatrix(k, 2) = Val(wgrid.TextMatrix(k, 2)) + Val(ogrid.TextMatrix(i, 3))
                    Exit For
                End If
            Next k
        End If
        If ogrid.TextMatrix(i, 6) <> "..." Then
            ttot(1) = Val(ttot(1)) + Val(ogrid.TextMatrix(i, 5))
            For k = 1 To wgrid.Rows - 1
                If ogrid.TextMatrix(i, 6) = wgrid.TextMatrix(k, 1) Then
                    wgrid.TextMatrix(k, 2) = Val(wgrid.TextMatrix(k, 2)) + Val(ogrid.TextMatrix(i, 5))
                    Exit For
                End If
            Next k
        End If
        If ogrid.TextMatrix(i, 8) <> "..." Then
            ttot(2) = Val(ttot(2)) + Val(ogrid.TextMatrix(i, 7))
            For k = 1 To wgrid.Rows - 1
                If ogrid.TextMatrix(i, 8) = wgrid.TextMatrix(k, 1) Then
                    wgrid.TextMatrix(k, 2) = Val(wgrid.TextMatrix(k, 2)) + Val(ogrid.TextMatrix(i, 7))
                    Exit For
                End If
            Next k
        End If
        If ogrid.TextMatrix(i, 10) <> "..." Then
            ttot(3) = Val(ttot(3)) + Val(ogrid.TextMatrix(i, 9))
            For k = 1 To wgrid.Rows - 1
                If ogrid.TextMatrix(i, 10) = wgrid.TextMatrix(k, 1) Then
                    wgrid.TextMatrix(k, 2) = Val(wgrid.TextMatrix(k, 2)) + Val(ogrid.TextMatrix(i, 9))
                    Exit For
                End If
            Next k
        End If
        
    Next i
    For i = 0 To 3
        If Val(ttot(i)) < Val(size(i)) Then
            ttot(i).ForeColor = &HFF&
        Else
            ttot(i).ForeColor = &H80000012
        End If
        If Val(size(i)) = 0 Then ttot(i) = ""
    Next i
End Sub
Private Sub insert_ship(paisle As String, prack As String, psku As String, pqty As Integer, pbb As String)
    Dim ds As ADODB.Recordset, s As String, zid As Long
    On Error GoTo vberror
    s = "select * from ship_rack where order_num = '" & Combo1 & "'"
    s = s & " and sku = '" & psku & "'"
    s = s & " and aisle = '" & paisle & "'"
    s = s & " and rack = '" & prack & "'"
    s = s & " and ship_status <> 'DONE'"
    s = s & " and ship_status <> 'CANC'"
    s = s & " and bbp = '" & pbb & "'"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        s = "Update ship_rack set ship_date = '" & Format(Now, "m-d-yyyy") & "'"
        s = s & ", order_qty = " & pqty
        s = s & ", ship_uom_qty = 0, ship_plt_qty = 0 where id = " & ds!id
        Wdb.Execute s
        ds.Close
    Else
        ds.Close
        s = "select * from ship_rack where order_num = '" & Combo1 & "'"
        s = s & " and sku = '" & psku & "'"
        s = s & " and aisle = '" & paisle & "'"
        s = s & " and rack = '" & prack & "'"
        s = s & " and bbp = '" & pbb & "'"
        Set ds = Wdb.Execute(s)
        If ds.BOF = False Then
            ds.MoveFirst
            s = "Update ship_rack set ship_date = '" & Format(Now, "m-d-yyyy") & "'"
            s = s & ", order_qty = " & pqty
            s = s & ", ship_uom_qty = 0, ship_plt_qty = 0, ship_status = 'NEW' where id = " & ds!id
            Wdb.Execute s
            ds.Close
        Else
            ds.Close
            s = "select * from ship_rack where ship_status = 'DONE'"
            s = s & " or ship_status = 'CANC'"
            Set ds = Wdb.Execute(s)
            If ds.BOF = False Then
                ds.MoveFirst
                s = "Update ship_rack set order_num = '" & Combo1 & "'"
                s = s & ", sku = '" & psku & "'"
                s = s & ",lot_num = ' '"
                s = s & ",ship_date = '" & Format$(Now, "m-d-yyyy") & "'"
                s = s & ", order_qty = " & pqty
                s = s & ", ship_uom_qty = 0, ship_plt_qty = 0, ship_status = 'NEW'"
                s = s & ", aisle = '" & paisle & "'"
                s = s & ", rack = '" & prack & "'"
                s = s & ", bbp = '" & pbb & "'"
                s = s & " Where id = " & ds!id
                Wdb.Execute s
            Else
                zid = wd_seq("Ship_Rack", Form1.shipdb)
                s = "INSERT INTO Ship_Rack (ID, Order_Num, SKU, Lot_Num, Ship_Date,"
                s = s & " Order_Qty, Ship_Uom_Qty, Ship_Plt_Qty, Ship_Status, Aisle,"
                s = s & " Rack, BBP) VALUES (" & zid & ","
                s = s & "'" & Combo1 & "',"
                s = s & "'" & psku & "',"
                s = s & "' ',"
                s = s & "'" & Format(Now, "mm-dd-yyyy") & "',"
                s = s & pqty & ",0,0,'NEW',"
                s = s & "'" & paisle & "',"
                s = s & "'" & prack & "',"
                s = s & "'" & pbb & "')"
                Wdb.Execute s
            End If
            ds.Close
        End If
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "insert_ship", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " insert_ship - Error Number: " & eno
        End
    End If
End Sub

Private Sub use_4way(psku As String, oqty As Integer)
    Dim sqty As Integer, rqty As Integer, pqty As Integer, s As String
    Dim ds As ADODB.Recordset, ds2 As ADODB.Recordset
    On Error GoTo vberror
    sqty = 0
    s = "select * from racks where aisle <> 'M'"
    s = s & " and id in (select rackno from rackpos where sku = '" & psku & "'"
    s = s & " and bbc = 'N')"
    s = s & " and hold <> 1"
    s = s & " order by fo desc, lot_num"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "select * from ship_rack where ship_status <> 'CANC'"
            s = s & " and ship_status <> 'DONE'"
            s = s & " and sku = '" & psku & "'"
            s = s & " and aisle = '" & ds!aisle & "'"
            s = s & " and rack = '" & ds!rack & "'"
            pqty = 0
            Set ds2 = Wdb.Execute(s)
            If ds2.BOF = False Then
                ds2.MoveFirst
                Do Until ds2.EOF
                    pqty = pqty + (ds2!order_qty - ds2!ship_plt_qty)
                    ds2.MoveNext
                Loop
            End If
            ds2.Close
            rqty = ds!qty4 - pqty
            If rqty > 0 Then
                If rqty >= (oqty - sqty) Then
                    Call insert_ship(ds!aisle, ds!rack, psku, oqty - sqty, "N")
                    sqty = oqty
                    Exit Do
                Else
                    Call insert_ship(ds!aisle, ds!rack, psku, rqty, "N")
                    sqty = sqty + rqty
                End If
            End If
            ds.MoveNext
        Loop
        If oqty > sqty Then Call insert_ship("_", "________", psku, oqty - sqty, "N")
    Else
        Call insert_ship("_", "________", psku, oqty, "N")
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "use_4way", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " use_4way - Error Number: " & eno
        End
    End If
End Sub
Private Sub use_bb(psku As String, oqty As Integer)
    Dim sqty As Integer, rqty As Integer, pqty As Integer, s As String
    Dim ds As ADODB.Recordset, ds2 As ADODB.Recordset
    On Error GoTo vberror
    sqty = 0
    s = "select * from racks where aisle <> 'M'"
    s = s & " and id in (select rackno from rackpos where sku = '" & psku & "'"
    s = s & " and bbc = 'Y')"
    s = s & " and hold <> 1"
    s = s & " order by fo desc, lot_num"
    Set ds = Wdb.Execute(s)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            s = "select * from ship_rack where ship_status <> 'CANC'"
            s = s & " and ship_status <> 'DONE'"
            s = s & " and sku = '" & psku & "'"
            s = s & " and aisle = '" & ds!aisle & "'"
            s = s & " and rack = '" & ds!rack & "'"
            pqty = 0
            Set ds2 = Wdb.Execute(s)
            If ds2.BOF = False Then
                ds2.MoveFirst
                Do Until ds2.EOF
                    pqty = pqty + (ds2!order_qty - ds2!ship_plt_qty)
                    ds2.MoveNext
                Loop
            End If
            ds2.Close
            rqty = ds!qty - pqty
            If rqty > 0 Then
                If rqty >= (oqty - sqty) Then
                    Call insert_ship(ds!aisle, ds!rack, psku, oqty - sqty, "Y")
                    sqty = oqty
                    Exit Do
                Else
                    Call insert_ship(ds!aisle, ds!rack, psku, rqty, "Y")
                    sqty = sqty + rqty
                End If
            End If
            ds.MoveNext
        Loop
        If oqty > sqty Then Call insert_ship("_", "________", psku, oqty - sqty, "Y")
    Else
        Call insert_ship("_", "________", psku, oqty, "Y")
    End If
    ds.Close
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "use_bb", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " use_bb - Error Number: " & eno
        End
    End If
End Sub

Private Sub altgrid_Click()
    altgrid.Col = 1: altgrid.ColSel = altgrid.Cols - 1
End Sub

Private Sub Combo1_Click()
    Label1 = Combo1
End Sub

Private Sub Command1_Click()                'Split SKU Order
    Dim sqlx As String, i As Integer, psku As String, pkey As Long
    On Error GoTo vberror
    psku = Label2
    pkey = wd_seq("groupitems", Form1.shipdb)
    sqlx = "Insert into groupitems (id, groupcode, sku, qty1, whs1, qty2, whs2, qty3, whs3, qty4, whs4, grank)"
    sqlx = sqlx & " Values (" & pkey & ", '" & Label1 & "', '" & Label2 & "'"
    sqlx = sqlx & ", 0, 0, 0, 0, 0, 0, 0, 0, 1)"
    Sdb.Execute sqlx
    Call refresh_ogrid
    For i = 1 To ogrid.Rows - 1
        If ogrid.TextMatrix(i, 1) = psku Then
            ogrid.Row = i
            Exit For
        End If
    Next i
    ogrid.Col = 3
    ogrid.TopRow = ogrid.Row
    Call ogrid_Click
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, Command1.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command1_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command2_Click()            'Insert Alternate
    Dim sqlx As String, i As Integer, pkey As Long
    Dim ds As ADODB.Recordset, psku As String
    On Error GoTo vberror
    psku = InputBox("Add Product to List", "Add SKU", altgrid.TextMatrix(altgrid.Row, 1))
    If Len(psku) = 0 Then Exit Sub
    sqlx = "Select * from skumast where sku = '" & psku & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = True Then
        MsgBox "Invalid SKU!", vbOKOnly + vbExclamation, "Sorry..."
        ds.Close
        Exit Sub
    End If
    ds.Close
    pkey = wd_seq("groupitems", Form1.shipdb)
    sqlx = "Insert into groupitems (id, groupcode, sku, qty1, whs1, qty2, whs2, qty3, whs3, qty4, whs4, grank)"
    sqlx = sqlx & " Values (" & pkey & ", '" & Label1 & "', '" & psku & "'"
    sqlx = sqlx & ", 0, 0, 0, 0, 0, 0, 0, 0, 1)"
    Sdb.Execute sqlx
    Call refresh_ogrid
    For i = 1 To ogrid.Rows - 1
        If ogrid.TextMatrix(i, 1) = psku Then
            ogrid.Row = i
            Exit For
        End If
    Next i
    ogrid.Col = 3
    ogrid.TopRow = ogrid.Row
    Call ogrid_Click
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, Command2.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command2_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Command3_Click()                    'Attach Notes
    If Combo1 <> noteGrid.TextMatrix(0, 0) Then
        noteGrid.Clear: noteGrid.Rows = 1: noteGrid.Cols = 4
        noteGrid.AddItem trlcode(0)
        noteGrid.AddItem trlcode(1)
        noteGrid.AddItem trlcode(2)
        noteGrid.AddItem trlcode(3)
        noteGrid.FormatString = "<" & Combo1 & "|<Note 1|<Note 2|<Note 3"
        noteGrid.ColWidth(0) = 1500
        noteGrid.ColWidth(1) = 1920
        noteGrid.ColWidth(2) = 1920
        noteGrid.ColWidth(3) = 1920
    End If
    noteGrid.Visible = True
    noteGrid.SetFocus
End Sub

Private Sub Command4_Click()                    'Print Checkoff
    Dim ds As ADODB.Recordset, sqlx As String
    Dim ws As ADODB.Recordset, mbp As String, pcopies As String, pcopy As Integer
    Dim rs As ADODB.Recordset, rstr As String
    Dim i As Integer, k As Integer, x As Long, y As Long, j As Integer
    Dim wq As String, qbb As Integer, q4 As Integer
    Dim w4way As Integer, wreg As Integer, wrega As Integer
    On Error GoTo vberror
    Screen.MousePointer = 11
    '  Post Rack Orders
    sqlx = "update ship_rack set ship_status = 'CANC'"   'jv
    sqlx = sqlx & " where order_num = '" & Trim(Combo1) & "'"  'jv
    Wdb.Execute sqlx   'jv
    wq = "("
    For i = 1 To wgrid.Rows - 1
        If UCase(Left$(wgrid.TextMatrix(i, 1), 3)) = "REG" Then
            wq = wq & wgrid.TextMatrix(i, 0) & ","
            If Right$(wgrid.TextMatrix(i, 1), 1) = "A" Then
                wreg = i 'wgrid.Row
            Else
                wrega = i 'wgrid.Row
            End If
        End If
        If Left$(wgrid.TextMatrix(i, 1), 1) = "4" Then
            wq = wq & wgrid.TextMatrix(i, 0) & ","
            w4way = i 'wgrid.Row
        End If
    Next i
    wq = Left$(wq, Len(wq) - 1) & ")"
    sqlx = "select * from groupitems where groupcode = '" & Trim(Combo1) & "'"
    sqlx = sqlx & " and sku <> 'PAR'"
    sqlx = sqlx & " and (whs1 in " & wq
    sqlx = sqlx & " or whs2 in " & wq
    sqlx = sqlx & " or whs3 in " & wq
    sqlx = sqlx & " or whs4 in " & wq & ")"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            qbb = 0: q4 = 0
            If ds!qty1 > 0 Then
                If ds!whs1 = w4way Then q4 = q4 + ds!qty1
                If ds!whs1 = wreg Or ds!whs1 = wrega Then qbb = qbb + ds!qty1
            End If
            If ds!qty2 > 0 Then
                If ds!whs2 = w4way Then q4 = q4 + ds!qty2
                If ds!whs2 = wreg Or ds!whs2 = wrega Then qbb = qbb + ds!qty2
            End If
            If ds!qty3 > 0 Then
                If ds!whs3 = w4way Then q4 = q4 + ds!qty3
                If ds!whs3 = wreg Or ds!whs3 = wrega Then qbb = qbb + ds!qty3
            End If
            If ds!qty4 > 0 Then
                If ds!whs4 = w4way Then q4 = q4 + ds!qty4
                If ds!whs4 = wreg Or ds!whs4 = wrega Then qbb = qbb + ds!qty4
            End If
            If qbb > 0 Then Call use_bb(ds!sku, qbb)  'jv
            If q4 > 0 Then Call use_4way(ds!sku, q4)  'jv
            ds.MoveNext
        Loop
    End If
    ds.Close
    
    '  Retreive Check off Data
    'cgrid.Visible = False
    cgrid.Clear: cgrid.Rows = 1: cgrid.Cols = 7
    sqlx = "select groupitems.sku,fgunit,fgdesc,qty1,whs1,qty2,whs2,qty3,whs3,qty4,whs4"
    sqlx = sqlx & " from groupitems,skumast"
    sqlx = sqlx & " where groupcode = '" & Combo1 & "'"
    sqlx = sqlx & " and (whs1 + whs2 + whs3 + whs4 > 0)"
    sqlx = sqlx & " and groupitems.sku = skumast.sku"
    sqlx = sqlx & " order by groupitems.sku"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            w0 = " ": w1 = " ": w2 = " ": w3 = " "
            If ds!whs1 > 0 Then
                sqlx = "select whs from warehouses where whs_num = " & ds!whs1
                Set ws = Sdb.Execute(sqlx)
                If ws.BOF = False Then w0 = ws!whs
                ws.Close
            End If
            If ds!whs2 > 0 Then
                sqlx = "select whs from warehouses where whs_num = " & ds!whs2
                Set ws = Sdb.Execute(sqlx)
                If ws.BOF = False Then w1 = ws!whs
                ws.Close
            End If
            If ds!whs3 > 0 Then
                sqlx = "select whs from warehouses where whs_num = " & ds!whs3
                Set ws = Sdb.Execute(sqlx)
                If ws.BOF = False Then w2 = ws!whs
                ws.Close
            End If
            If ds!whs4 > 0 Then
                sqlx = "select whs from warehouses where whs_num = " & ds!whs4
                Set ws = Sdb.Execute(sqlx)
                If ws.BOF = False Then w3 = ws!whs
                ws.Close
            End If
            sqlx = " " & ds!sku & " "
            sqlx = sqlx & ds!fgunit & " " & ds!fgdesc & Chr(9)
            If ds!whs1 > 0 Then sqlx = sqlx & " " & ds!qty1 & " " & w0
            sqlx = sqlx & Chr(9)
            If ds!whs2 > 0 Then sqlx = sqlx & " " & ds!qty2 & " " & w1
            sqlx = sqlx & Chr(9)
            If ds!whs3 > 0 Then sqlx = sqlx & " " & ds!qty3 & " " & w2
            sqlx = sqlx & Chr(9)
            If ds!whs4 > 0 Then sqlx = sqlx & " " & ds!qty4 & " " & w3
            sqlx = sqlx & Chr(9)
            If (ds!whs1 > 3 Or ds!whs2 > 3 Or ds!whs3 > 3 Or ds!whs4 > 3) Then
                rstr = "select * from ship_rack where order_num = '" & Trim(Combo1) & "'"
                rstr = rstr & " and sku = '" & ds!sku & "'"
                rstr = rstr & " and ship_status not in ('CANC','DONE')"
                Set rs = Wdb.Execute(rstr)
                If rs.BOF = False Then
                    rs.MoveFirst
                    Do Until rs.EOF
                        If UCase(rs!bbp) = "Y" Then
                            mbp = " "
                        Else
                            mbp = " 4Way"
                        End If
                        sqlx = rs!order_qty & mbp & Chr(9) & rs!aisle & " " & rs!rack & Chr(9) & sqlx
                        cgrid.AddItem sqlx
                        sqlx = " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & " " & Chr(9) & " " & Chr(9)
                        rs.MoveNext
                    Loop
                Else
                    cgrid.AddItem Chr(9) & Chr(9) & sqlx
                End If
                rs.Close
            Else
                cgrid.AddItem Chr(9) & Chr(9) & sqlx
            End If
            ds.MoveNext
        Loop
        sqlx = Chr(9) & Chr(9) & " Group Totals" & Chr(9) & ttot(0) & Chr(9) & ttot(1) & Chr(9) & ttot(2) & Chr(9) & ttot(3)
        cgrid.AddItem sqlx
        If noteGrid.TextMatrix(0, 0) = Combo1 Then
            sqlx = Chr(9) & Chr(9) & " Notes" & Chr(9)
            sqlx = sqlx & noteGrid.TextMatrix(1, 1) & Chr(9)
            sqlx = sqlx & noteGrid.TextMatrix(2, 1) & Chr(9)
            sqlx = sqlx & noteGrid.TextMatrix(3, 1) & Chr(9)
            sqlx = sqlx & noteGrid.TextMatrix(4, 1)
            cgrid.AddItem sqlx
            If noteGrid.TextMatrix(1, 2) > "0" Or noteGrid.TextMatrix(2, 2) > "0" Or noteGrid.TextMatrix(3, 2) > "0" Or noteGrid.TextMatrix(4, 2) > "0" Then
                sqlx = Chr(9) & Chr(9) & "..." & Chr(9)
                sqlx = sqlx & noteGrid.TextMatrix(1, 2) & Chr(9)
                sqlx = sqlx & noteGrid.TextMatrix(2, 2) & Chr(9)
                sqlx = sqlx & noteGrid.TextMatrix(3, 2) & Chr(9)
                sqlx = sqlx & noteGrid.TextMatrix(4, 2)
                cgrid.AddItem sqlx
            End If
            If noteGrid.TextMatrix(1, 3) > "0" Or noteGrid.TextMatrix(2, 3) > "0" Or noteGrid.TextMatrix(3, 3) > "0" Or noteGrid.TextMatrix(4, 3) > "0" Then
                sqlx = Chr(9) & Chr(9) & "..." & Chr(9)
                sqlx = sqlx & noteGrid.TextMatrix(1, 3) & Chr(9)
                sqlx = sqlx & noteGrid.TextMatrix(2, 3) & Chr(9)
                sqlx = sqlx & noteGrid.TextMatrix(3, 3) & Chr(9)
                sqlx = sqlx & noteGrid.TextMatrix(4, 3)
                cgrid.AddItem sqlx
            End If
        End If
        sqlx = Chr(9) & Chr(9) & " Printed: " & Format(Now, "m-d-yy h:mm am/pm") & Chr(9) & " Start Time --->" & Chr(9) & Chr(9) & " End Time ---->"
        cgrid.AddItem sqlx
    End If
    ds.Close
    cgrid.TextMatrix(0, 2) = " Product"
    cgrid.TextMatrix(0, 3) = trlcode(0)
    cgrid.TextMatrix(0, 4) = trlcode(1)
    cgrid.TextMatrix(0, 5) = trlcode(2)
    cgrid.TextMatrix(0, 6) = trlcode(3)
    cgrid.TextMatrix(0, 0) = "Pallets"
    cgrid.TextMatrix(0, 1) = "Rack"
    cgrid.ColWidth(2) = 3675
    cgrid.ColWidth(3) = 1920: cgrid.ColWidth(4) = 1920
    cgrid.ColWidth(5) = 1920: cgrid.ColWidth(6) = 1920
    cgrid.ColWidth(0) = 720: cgrid.ColWidth(1) = 1420
    
    Screen.MousePointer = 0
    
    'If cgrid.TextMatrix(0, 3) <= " " Then cgrid.ColWidth(3) = 0
    'If cgrid.TextMatrix(0, 4) <= " " Then cgrid.ColWidth(4) = 0
    'If cgrid.TextMatrix(0, 5) <= " " Then cgrid.ColWidth(5) = 0
    'If cgrid.TextMatrix(0, 6) <= " " Then cgrid.ColWidth(6) = 0
    'Exit Sub
    
    pcopies = InputBox("Enter # copies:", "Select # copies...", 2)
    If Len(pcopies) = 0 Then Exit Sub
    Screen.MousePointer = 11
    For pcopy = 1 To Val(pcopies)
    '  Print Check Off
    Printer.FontTransparent = True
    Printer.FillStyle = 0
    Printer.FillColor = QBColor(0)
    Printer.DrawMode = 1
    Printer.ForeColor = QBColor(0)
    
    Printer.FontName = "MS Serif"
    Printer.FontTransparent = True
    Printer.FontSize = 14
    Printer.Duplex = 3
    Printer.DrawWidth = 6
    Printer.Print "Check Off Sheet"
'    printer.Print "Group: "; Combo1
'jv
    Printer.Print "Group: "; Combo1
    Printer.Print gdate
    Dim xs As Long, xe As Long, xm As Long, gl As Integer
    Dim ys As Long, ye As Long, pg1 As Integer
    'printer.FontSize = 8
    xs = cgrid.ColWidth(2) - 800: Printer.Line (xs, 480)-(xs, 1200)
    xs = cgrid.ColWidth(2): Printer.Line (xs, 480)-(xs, 1200)
    xs = xs + cgrid.ColWidth(3): Printer.Line (xs, 480)-(xs, 1200)
    xs = xs + cgrid.ColWidth(4): Printer.Line (xs, 480)-(xs, 1200)
    xs = xs + cgrid.ColWidth(5): Printer.Line (xs, 480)-(xs, 1200)
    xs = xs + cgrid.ColWidth(6): Printer.Line (xs, 480)-(xs, 1200)
    xe = xs: xs = cgrid.ColWidth(2) - 800
    Printer.Line (xs, 480)-(xe, 480): Printer.Line (xs, 720)-(xe, 720)
    Printer.Line (xs, 960)-(xe, 960)

    'Printer.FontName = "MS Sans Serif"
    Printer.FontSize = 8
    'printer.PSet (xs, 540): printer.Print "Loader"
    'printer.PSet (xs, 780): printer.Print "Seal #"
    'printer.PSet (xs, 1020): printer.Print "Trailer"
    xs = 0
    xe = xs + cgrid.ColWidth(2) + cgrid.ColWidth(3) + cgrid.ColWidth(4)
    xe = xe + cgrid.ColWidth(5) + cgrid.ColWidth(6)
    gl = 0
    For i = 0 To cgrid.Rows - 1
        If Len(cgrid.TextMatrix(i, 2)) > 0 And cgrid.TextMatrix(i, 2) > " 000 " Then gl = gl + 1
    Next i
    Printer.Line (xs, 1200)-(xe, 1200)
    Printer.Line (xs, 1440)-(xe, 1440)
    Printer.FillColor = QBColor(15)
    Printer.DrawWidth = 3
    If gl > 57 Then
        pg1 = 57
    Else
        pg1 = gl
    End If
    For i = 0 To pg1 - 1
        ye = i * 240 + 1440
        k = i Mod 3
        Printer.Line (xs, ye)-(xe, ye)
        If k = 0 Then
            If Printer.FillColor = QBColor(14) Then
                Printer.FillColor = QBColor(15)
                Printer.Line (xs, ye)-(xe, ye + 720), Printer.FillColor, BF
                
            Else
                Printer.FillColor = QBColor(14)
                'printer.Line (xs, ye)-(xe, ye + 720), printer.FillColor, BF
                Printer.Line (xs, ye)-(xe, ye + 720), &HFFFF&, BF
            End If
        End If
    Next i
    Printer.DrawWidth = 1
    Printer.FontBold = False
    j = 0
    For i = 0 To cgrid.Rows - 1 'pg1 - 1 'cgrid.Rows - 1
        If Len(cgrid.TextMatrix(i, 2)) > 0 And cgrid.TextMatrix(i, 2) > " 000 " Then
            xm = xs + 100
            Printer.FontBold = Not Printer.FontBold
            For k = 2 To cgrid.Cols - 1
                Printer.PSet (xm, j * 240 + 1230)
                Printer.Print cgrid.TextMatrix(i, k)
                xm = xm + cgrid.ColWidth(k)
            Next k
            j = j + 1
            If j >= pg1 Then
                pg1 = j
                Exit For
            End If
        End If
    Next i
    Printer.FontBold = True
    xm = cgrid.ColWidth(2) - 700
    Printer.PSet (xm, 510): Printer.Print "Loader"
    Printer.PSet (xm, 750): Printer.Print "Seal #"
    Printer.PSet (xm, 990): Printer.Print "Trailer"
    ys = 1200
    xm = xs
    Printer.DrawWidth = 6
    For i = 2 To cgrid.Cols - 1
        'MsgBox "YS:" & ys & " YE:" & ye & " XM:" & xm
        Printer.Line (xm, ys)-(xm, ye)
        xm = xm + cgrid.ColWidth(i)
    Next i
    Printer.Line (xm, ys)-(xm, ye)
    Printer.NewPage
    If gl > pg1 Then
        j = 1
        Printer.FontBold = True
        xm = xs + 100
        Printer.Line (xs, 0)-(xe, 0)
        For k = 2 To cgrid.Cols - 1
            Printer.PSet (xm, 30)
            Printer.Print cgrid.TextMatrix(0, k)
            xm = xm + cgrid.ColWidth(k)
        Next k
        For i = pg1 To gl - 1
            If cgrid.TextMatrix(i, 2) > " 000 " Then
                Printer.Line (xs, j * 240)-(xe, j * 240)
                xm = xs + 100
                Printer.FontBold = Not Printer.FontBold
                For k = 2 To cgrid.Cols - 1
                    Printer.PSet (xm, j * 240 + 30)
                    Printer.Print cgrid.TextMatrix(i, k)
                    xm = xm + cgrid.ColWidth(k)
                Next k
                j = j + 1
            End If
            'Exit For
        Next i
        Printer.Line (xs, j * 240)-(xe, j * 240)
        xm = 0
        For k = 2 To cgrid.Cols - 1
            Printer.Line (xm, 0)-(xm, j * 240)
            xm = xm + cgrid.ColWidth(k)
        Next k
    End If
        
            
    Printer.FontSize = 12
    Printer.FontBold = True
    Printer.Print " "
    Printer.Print "      Rack Checkoff  "; Combo1; "   "; gdate
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.Print ""
    Printer.Print Space(52); "Pallets" '   Rack"    jv072811
    'Printer.Print " "; String(76, "_")
    For i = 0 To cgrid.Rows - 1
        If Val(cgrid.TextMatrix(i, 0)) > 0 Then
             If cgrid.TextMatrix(i, 2) > " 000 " Then
                Printer.Print "|"; String(76, "_"); "|"
            End If
             sqlx = "|" & cgrid.TextMatrix(i, 2)
             sqlx = sqlx & Space(50 - Len(sqlx))
             sqlx = sqlx & Space(8 - Len(cgrid.TextMatrix(i, 0)))
             sqlx = sqlx & cgrid.TextMatrix(i, 0) & Space(4)
             'Turn off printing rack label    jv072811
             'sqlx = sqlx & cgrid.TextMatrix(i, 1)
             If Len(sqlx) < 78 Then
                sqlx = sqlx & Space(77 - Len(sqlx)) & "|"
             End If
             Printer.Print sqlx
             'Printer.Print "|"; String(76, "_"); "|"
         End If
     Next i
    Printer.Print "|"; String(76, "_"); "|"
 '  Process Alternates
    Printer.Print " "
    Printer.Print "Start Time: ________________________      End Time: ________________________"
    Printer.Print " "
    'printer.FontName = "Sans Serif"
    
    'Turn off Alternate Pallets         jv072811
    'Printer.FontSize = 8
    'Printer.Print "Alternate Pallets": Printer.Print " "
    'Dim alist(3, 20) As String
    'For i = 0 To 3
    '    If trlcode(i) > "   " Then
    '        alist(i, 0) = Left(trlcode(i), Len(trlcode(i)) - 2) & ":"
    '        alist(i, 1) = "None specified."
    '        j = 1
    '        For k = 1 To altgrid.Rows - 1
    '            If altgrid.TextMatrix(k, 0) = br(i) Then
    '                alist(i, j) = altgrid.TextMatrix(k, 1) & " " & altgrid.TextMatrix(k, 2)
    '                j = j + 1
    '                If j > 20 Then Exit For
    '            End If
    '        Next k
    '    End If
    'Next i
    'For i = 0 To 20
    '    sqlx = alist(0, i) & alist(1, i) & alist(2, i)
    '    If Len(sqlx) > 0 Then
    '        Printer.Print alist(0, i); Tab(40);
    '        Printer.Print alist(1, i); Tab(80);
    '        Printer.Print alist(2, i)
    '    End If
    'Next i
    'Printer.Print " "
    'For i = 0 To 20
    '    If Len(alist(3, i)) > 0 Then Printer.Print alist(3, i)
    'Next i
    Printer.EndDoc
    Next pcopy
    Printer.Duplex = 1
    Screen.MousePointer = 0
    Exit Sub
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, Command4.Caption & "_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " command4_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub cw_Click(Index As Integer)
    Dim ds As ADODB.Recordset, sqlx As String, y As Integer
    On Error GoTo vberror
    sqlx = "select * from brorders"
    sqlx = sqlx & " Where branch = " & Val(br(Index))
    sqlx = sqlx & " and plant = " & Val(gplant)
    sqlx = sqlx & " and orddate = '" & gdate & "'"
    sqlx = sqlx & " and SKU = '" & Label2 & "'"
    sqlx = sqlx & " and ordqty > 0"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        If (ds!grpqty - Val(ord(Index))) < 0 Then   'This could distort order
            MsgBox "This product has already been un-grouped.", vbOKOnly + vbExclamation, "Nope!!!"
            ds.Close
            'db.Close
            Exit Sub
        End If
        sqlx = "Update brorders set"
        sqlx = sqlx & " grpqty = grpqty - " & Val(ord(Index))
        sqlx = sqlx & ", netqty = netqty + " & Val(ord(Index))
        sqlx = sqlx & " where id = " & ds!id
        Sdb.Execute sqlx
    End If
    ds.Close
    If Val(w(Index)) > 0 Then
        sqlx = "select * from whstotals where whs_num = " & w(Index)
        sqlx = sqlx & " and sku = '" & Label2 & "'"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            sqlx = "Update whstotals set"
            sqlx = sqlx & " grp_qty = grp_qty - " & Val(ord(Index))
            sqlx = sqlx & ", avail = avail + " & Val(ord(Index))
            sqlx = sqlx & " where id = " & ds!id
            Sdb.Execute sqlx
        End If
        ds.Close
    End If
    sqlx = "select * from groupitems where id = " & ogrid.TextMatrix(ogrid.Row, 0)
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        sqlx = "Update groupitems set "
        If Index = 0 Then sqlx = sqlx & "whs1 = 0"
        If Index = 1 Then sqlx = sqlx & "whs2 = 0"
        If Index = 2 Then sqlx = sqlx & "whs3 = 0"
        If Index = 3 Then sqlx = sqlx & "whs4 = 0"
        sqlx = sqlx & " Where id = " & ds!id
        Sdb.Execute sqlx
    End If
    ds.Close
    ogrid.TextMatrix(ogrid.Row, Index + 11) = 0
    w(Index) = 0
    ogrid.TextMatrix(ogrid.Row, Index * 2 + 4) = "..."
    Call calc_whstots
    Call calc_trltots
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "cw_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " cw_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Activate()
    rf = "Yes"
End Sub

Private Sub Form_Deactivate()
    Dim i As Integer
    If Editgroups.WindowState = 0 Then
        For i = 1 To Form1.FrmGrid.Rows - 1
            If Form1.FrmGrid.TextMatrix(i, 0) = "editgroups" Then
                Form1.FrmGrid.TextMatrix(i, 1) = Editgroups.Top
                Form1.FrmGrid.TextMatrix(i, 2) = Editgroups.Left
                Form1.FrmGrid.TextMatrix(i, 3) = Editgroups.Height
                Form1.FrmGrid.TextMatrix(i, 4) = Editgroups.Width
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Editgroups.ActiveControl.Name = "ogrid" Then
        If KeyCode = 45 Or KeyCode = 121 Then Call Command2_Click 'insert, F10
        If KeyCode = 46 Or KeyCode = 120 Then 'Delete, F9
            Screen.MousePointer = 11
            If ogrid.Col = 4 And cw(0).Visible = True Then Call cw_Click(0)
            If ogrid.Col = 6 And cw(1).Visible = True Then Call cw_Click(1)
            If ogrid.Col = 8 And cw(2).Visible = True Then Call cw_Click(2)
            If ogrid.Col = 10 And cw(3).Visible = True Then Call cw_Click(3)
            If ogrid.Col = 3 And cw(0).Visible = True Then Call cw_Click(0)
            If ogrid.Col = 5 And cw(1).Visible = True Then Call cw_Click(1)
            If ogrid.Col = 7 And cw(2).Visible = True Then Call cw_Click(2)
            If ogrid.Col = 9 And cw(3).Visible = True Then Call cw_Click(3)
            Screen.MousePointer = 0
        End If
        If KeyCode >= 112 And KeyCode <= 119 Then 'f1-f8
            Screen.MousePointer = 11
            wgrid.Row = KeyCode - 111
            Call wgrid_Click
            DoEvents
            If ogrid.Col = 4 And pw(0).Visible = True Then Call pw_Click(0)
            If ogrid.Col = 6 And pw(1).Visible = True Then Call pw_Click(1)
            If ogrid.Col = 8 And pw(2).Visible = True Then Call pw_Click(2)
            If ogrid.Col = 10 And pw(3).Visible = True Then Call pw_Click(3)
            If ogrid.Col = 3 And pw(0).Visible = True Then Call pw_Click(0)
            If ogrid.Col = 5 And pw(1).Visible = True Then Call pw_Click(1)
            If ogrid.Col = 7 And pw(2).Visible = True Then Call pw_Click(2)
            If ogrid.Col = 9 And pw(3).Visible = True Then Call pw_Click(3)
            DoEvents
            Screen.MousePointer = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim ds As ADODB.Recordset, sqlx As String, hflag As Boolean
    Dim eno As Long, edesc As String
    Dim i As Integer
    For i = 1 To Form1.FrmGrid.Rows - 1
        If Form1.FrmGrid.TextMatrix(i, 0) = "editgroups" Then
            Editgroups.Top = Val(Form1.FrmGrid.TextMatrix(i, 1))
            Editgroups.Left = Val(Form1.FrmGrid.TextMatrix(i, 2))
            Editgroups.Height = Val(Form1.FrmGrid.TextMatrix(i, 3))
            Editgroups.Width = Val(Form1.FrmGrid.TextMatrix(i, 4))
            Exit For
        End If
    Next i
    'On Error GoTo vberror
    Combo1.Clear
    sqlx = "select distinct groupcode from trgroups order by groupcode"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        Do Until ds.EOF
            Combo1.AddItem ds!groupcode
            ds.MoveNext
        Loop
    End If
    ds.Close
    wgrid.Font = "Arial": ogrid.Font = "Arial": altgrid.Font = "Arial"
    wgrid.FontSize = 9: ogrid.FontSize = 9: altgrid.FontSize = 8
    wgrid.FontBold = True: ogrid.FontBold = True: altgrid.FontBold = True
    
    wgrid.Clear: wgrid.Rows = 1: wgrid.Cols = 4
    Set ds = Sdb.Execute("select * from warehouses order by whs_num")
    wgrid.FormatString = "ID|^Whs|^Group|^Avail"
    wgrid.ColWidth(0) = 15: wgrid.ColWidth(1) = 800
    wgrid.ColWidth(2) = 1000: wgrid.ColWidth(3) = 1000
    ds.MoveFirst
    Do Until ds.EOF
        sqlx = ds!whs_num & Chr(9) & ds!whs
        wgrid.AddItem sqlx
        ds.MoveNext
    Loop
    ds.Close
    If wgrid.Rows > 1 Then
        wgrid.FillStyle = flexFillRepeat
        For i = 1 To wgrid.Rows - 1
            If hflag = True Then
                wgrid.Row = i: wgrid.RowSel = i
                wgrid.Col = 0: wgrid.ColSel = wgrid.Cols - 1
                wgrid.CellBackColor = ogrid.BackColorFixed
            End If
            hflag = Not hflag
        Next i
    End If
    rf = "Yes"
    If Combo1.ListCount > 0 Then
        For i = 0 To Combo1.ListCount - 1
            If Combo1.List(i) = Form1.cgrp Then
                Combo1.ListIndex = i
                Exit For
            End If
        Next i
        If Combo1.ListIndex < 0 Then Combo1.ListIndex = 0
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "form_load", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " form_load - Error Number: " & eno
        End
    End If
End Sub

Private Sub Form_Resize()
    If Editgroups.Width > (altgrid.Width + 400) Then
        wgrid.Left = Editgroups.Width - wgrid.Width - 100
        altgrid.Left = Editgroups.Width - (altgrid.Width + 150)
        Command2.Left = Editgroups.Width - (Command2.Width + 150)
    End If
    If Editgroups.Height > wgrid.Height Then
        altgrid.Top = Editgroups.Height - (altgrid.Height + 400)
        Command2.Top = altgrid.Top - Command2.Height
        ogrid.Height = altgrid.Top - ogrid.Top
        Command4.Top = Editgroups.Height - 800
        Command3.Top = Command4.Top - (Command4.Height + 200)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Form_Deactivate
End Sub

Private Sub Label1_Change()
    Dim ds As ADODB.Recordset, sqlx As String, i As Integer
    Dim eno As Long, edesc As String
    'On Error GoTo vberror
    If rf = "Yes" Then
        Call refresh_ogrid
        If ogrid.Rows > 1 Then
            ogrid.Row = 1
            Call ogrid_Click
        End If
        Form1.cgrp = Label1
        For i = 0 To 3
            If Val(runid(i)) <= 0 Then
                br(i) = "": trlcode(i) = "": size(i) = ""
            Else
                sqlx = "Select destination,trlno,trlsize,locname,loaded,trldate from runs where id = " & runid(i)
                Set ds = Sdb.Execute(sqlx)
                If ds.BOF = True Then
                    br(i) = "": trlcode(i) = "": size(i) = ""
                Else
                    ds.MoveFirst
                    trlcode(i) = ds!locname & " " & ds!trlno
                    size(i) = ds!trlsize
                    br(i) = ds!Destination
                    gplant = ds!loaded
                    gdate = Format$(ds!trldate, "m-d-yyyy")
                End If
                ds.Close
            End If
        Next i
        Call calc_trltots
        Call calc_whstots
        Call refresh_alts
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "label1_change", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " label1_change - Error Number: " & eno
        End
    End If
End Sub

Private Sub Label2_Change()
    Dim ds As ADODB.Recordset, sqlx As String, i As Integer
    Dim eno As Long, edesc As String
    On Error GoTo vberror
    If rf = "Yes" Then
        Command1.Visible = True
        For i = 1 To wgrid.Rows - 1
            wgrid.TextMatrix(i, 3) = " "
        Next i
        sqlx = "Select whs_num,avail from whstotals where sku = '" & Label2 & "'"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            Do Until ds.EOF
                wgrid.TextMatrix(ds!whs_num, 3) = Val(wgrid.TextMatrix(ds!whs_num, 3)) + ds!avail
                ds.MoveNext
            Loop
        End If
        ds.Close
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "label2_change", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " label2_change - Error Number: " & eno
        End
    End If
End Sub

Private Sub noteGrid_KeyPress(KeyAscii As Integer)
    If noteGrid.Row = 0 Or noteGrid.Col = 0 Then Exit Sub
    If noteGrid.TextMatrix(noteGrid.Row, 0) <= "   " Then Exit Sub
    If KeyAscii = 8 Then
        If Len(noteGrid.Text) > 1 Then
            noteGrid.Text = Left(noteGrid.Text, Len(noteGrid.Text) - 1)
        Else
            noteGrid.Text = ""
        End If
    End If
    If KeyAscii > 31 And KeyAscii < 127 Then
        noteGrid.Text = noteGrid.Text & Chr(KeyAscii)
    End If
End Sub

Private Sub noteGrid_LostFocus()
    noteGrid.Visible = False
End Sub

Private Sub ogrid_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If ogrid.Col = 3 And Val(size(0)) > 0 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            ord(0) = ord(0) & Chr(KeyAscii)
            Call ord_KeyUp(0, KeyAscii, 0)
        End If
        If KeyAscii = 8 Then
            If Len(ogrid.Text) > 1 Then
                ord(0) = Left(ord(0), Len(ord(0)) - 1)
            Else
                ord(0) = " "
            End If
            Call ord_KeyUp(0, 8, 0)
        End If
    End If
    If ogrid.Col = 5 And Val(size(1)) > 0 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            ord(1) = ord(1) & Chr(KeyAscii)
            Call ord_KeyUp(1, KeyAscii, 0)
        End If
        If KeyAscii = 8 Then
            If Len(ogrid.Text) > 1 Then
                ord(1) = Left(ord(1), Len(ord(1)) - 1)
            Else
                ord(1) = " "
            End If
            Call ord_KeyUp(1, 8, 0)
        End If
    End If
    If ogrid.Col = 7 And Val(size(2)) > 0 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            ord(2) = ord(2) & Chr(KeyAscii)
            Call ord_KeyUp(2, KeyAscii, 0)
        End If
        If KeyAscii = 8 Then
            If Len(ogrid.Text) > 1 Then
                ord(2) = Left(ord(2), Len(ord(2)) - 1)
            Else
                ord(2) = " "
            End If
            Call ord_KeyUp(2, 8, 0)
        End If
    End If
    If ogrid.Col = 9 And Val(size(3)) > 0 Then
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            ord(3) = ord(3) & Chr(KeyAscii)
            Call ord_KeyUp(3, KeyAscii, 0)
        End If
        If KeyAscii = 8 Then
            If Len(ogrid.Text) > 1 Then
                ord(3) = Left(ord(3), Len(ord(3)) - 1)
            Else
                ord(3) = " "
            End If
            Call ord_KeyUp(3, 8, 0)
        End If
    End If
    For i = 0 To 3
        If Val(ord(i)) > 0 Then
            cw(i).Visible = True: pw(i).Visible = True
        Else
            cw(i).Visible = False: pw(i).Visible = False
        End If
    Next i
End Sub

Private Sub ogrid_RowColChange()
    Call ogrid_Click
End Sub

Private Sub wgrid_Click()
    Dim i As Integer
    For i = 0 To 3
        pw(i).Caption = "<<- " & wgrid.TextMatrix(wgrid.Row, 1)
        If Val(ord(i)) > 0 Then
            pw(i).Visible = True: cw(i).Visible = True
        Else
            pw(i).Visible = False: cw(i).Visible = False
        End If
    Next i
    wgrid.Col = 0: wgrid.ColSel = 3
    wgrid.RowSel = wgrid.Row
End Sub

Private Sub ogrid_Click()
    Dim y As Integer, i As Integer
    y = ogrid.Row
    Label2 = ogrid.TextMatrix(y, 1): Label3 = ogrid.TextMatrix(y, 2)
    ord(0) = ogrid.TextMatrix(y, 3): ord(1) = ogrid.TextMatrix(y, 5)
    ord(2) = ogrid.TextMatrix(y, 7): ord(3) = ogrid.TextMatrix(y, 9)
    w(0) = ogrid.TextMatrix(y, 11): w(1) = ogrid.TextMatrix(y, 12)
    w(2) = ogrid.TextMatrix(y, 13): w(3) = ogrid.TextMatrix(y, 14)
    For i = 0 To wgrid.Rows - 1
        If wgrid.TextMatrix(i, 1) = ogrid.TextMatrix(y, 4) Then
            wgrid.Row = i: Exit For
        End If
        If wgrid.TextMatrix(i, 1) = ogrid.TextMatrix(y, 6) Then
            wgrid.Row = i: Exit For
        End If
        If wgrid.TextMatrix(i, 1) = ogrid.TextMatrix(y, 8) Then
            wgrid.Row = i: Exit For
        End If
        If wgrid.TextMatrix(i, 1) = ogrid.TextMatrix(y, 10) Then
            wgrid.Row = i: Exit For
        End If
    Next i
    Call wgrid_Click
End Sub

Private Sub ord_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim sqlx As String, diff As Integer, gid As Long, gc As Integer
    Dim ds As ADODB.Recordset
    On Error GoTo vberror
    If Index = 0 Then gc = 3
    If Index = 1 Then gc = 5
    If Index = 2 Then gc = 7
    If Index = 3 Then gc = 9
    diff = Val(ogrid.TextMatrix(ogrid.Row, gc)) - Val(ord(Index))
    If Val(w(Index)) > 0 Then
        sqlx = "select * from brorders"
        sqlx = sqlx & " Where branch = " & Val(br(Index))
        sqlx = sqlx & " and plant = " & Val(gplant)
        sqlx = sqlx & " and orddate = '" & gdate & "'"
        sqlx = sqlx & " and SKU = '" & Label2 & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            sqlx = "Update brorders set"
            sqlx = sqlx & " grpqty = grpqty - " & diff
            sqlx = sqlx & ", netqty = netqty + " & diff
            sqlx = sqlx & " Where id = " & ds!id
            Sdb.Execute sqlx
        End If
        ds.Close
    End If
    sqlx = "select * from groupitems"
    sqlx = sqlx & " Where id = " & ogrid.TextMatrix(ogrid.Row, 0)
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        sqlx = "Update groupitems set "
        If Index = 0 Then sqlx = sqlx & "qty1 = " & Val(ord(Index))
        If Index = 1 Then sqlx = sqlx & "qty2 = " & Val(ord(Index))
        If Index = 2 Then sqlx = sqlx & "qty3 = " & Val(ord(Index))
        If Index = 3 Then sqlx = sqlx & "qty4 = " & Val(ord(Index))
        sqlx = sqlx & " Where id = " & ds!id
        Sdb.Execute sqlx
    End If
    ds.Close
    ogrid.TextMatrix(ogrid.Row, gc) = Val(ord(Index))
    If w(Index) > 0 Then
        sqlx = "select * from whstotals"
        sqlx = sqlx & " Where whs_num = " & w(Index)
        sqlx = sqlx & " And SKU = '" & Label2 & "'"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            sqlx = "Update whstotals set "
            sqlx = sqlx & "grp_qty = grp_qty - " & diff
            sqlx = sqlx & ", avail = avail + " & diff
            sqlx = sqlx & " Where id = " & ds!id
            Sdb.Execute sqlx
        End If
        ds.Close
        Call calc_whstots
        Call calc_trltots
    End If
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "ord_keyup", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " ord_keyup - Error Number: " & eno
        End
    End If
End Sub

Private Sub pw_Click(Index As Integer)
    Dim ds As ADODB.Recordset, sqlx As String, y As Integer
    On Error GoTo vberror
    If Val(w(Index)) > 0 Then
        sqlx = "select * from whstotals where whs_num = " & w(Index)
        sqlx = sqlx & " and sku = '" & Label2 & "'"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            sqlx = "Update whstotals set grp_qty = grp_qty - " & Val(ord(Index))
            sqlx = sqlx & ", avail = avail + " & Val(ord(Index)) & " Where id = " & ds!id
            Sdb.Execute sqlx
        End If
        ds.Close
    Else
        sqlx = "select * from brorders"
        sqlx = sqlx & " where branch = " & Val(br(Index))
        sqlx = sqlx & " and plant = " & Val(gplant)
        sqlx = sqlx & " and orddate = '" & gdate & "'"
        sqlx = sqlx & " and sku = '" & Label2 & "'"
        sqlx = sqlx & " and ordqty > 0"
        Set ds = Sdb.Execute(sqlx)
        If ds.BOF = False Then
            ds.MoveFirst
            sqlx = "Update brorders set grpqty = grpqty + " & Val(ord(Index))
            sqlx = sqlx & ", netqty = netqty - " & Val(ord(Index)) & " Where id = " & ds!id
            Sdb.Execute sqlx
        End If
        ds.Close
    End If
    sqlx = "select * from whstotals where whs_num = " & Val(wgrid.TextMatrix(wgrid.Row, 0))
    sqlx = sqlx & " and sku = '" & Label2 & "'"
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        sqlx = "Update whstotals set grp_qty = grp_qty + " & Val(ord(Index))
        sqlx = sqlx & ", avail = avail - " & Val(ord(Index)) & " Where id = " & ds!id
        Sdb.Execute sqlx
    End If
    ds.Close
    sqlx = "select * from groupitems where id = " & ogrid.TextMatrix(ogrid.Row, 0)
    Set ds = Sdb.Execute(sqlx)
    If ds.BOF = False Then
        ds.MoveFirst
        sqlx = "Update groupitems set "
        If Index = 0 Then sqlx = sqlx & "whs1 = " & Val(wgrid.TextMatrix(wgrid.Row, 0))
        If Index = 1 Then sqlx = sqlx & "whs2 = " & Val(wgrid.TextMatrix(wgrid.Row, 0))
        If Index = 2 Then sqlx = sqlx & "whs3 = " & Val(wgrid.TextMatrix(wgrid.Row, 0))
        If Index = 3 Then sqlx = sqlx & "whs4 = " & Val(wgrid.TextMatrix(wgrid.Row, 0))
        sqlx = sqlx & " Where id = " & ds!id
        Sdb.Execute sqlx
    End If
    ds.Close
    ogrid.TextMatrix(ogrid.Row, Index + 11) = wgrid.TextMatrix(wgrid.Row, 0)
    w(Index) = Val(wgrid.TextMatrix(wgrid.Row, 0))
    ogrid.TextMatrix(ogrid.Row, Index * 2 + 4) = wgrid.TextMatrix(wgrid.Row, 1)
    Call calc_whstots
    Call calc_trltots
    Exit Sub
vberror:
    eno = Err.Number: edesc = Err.description: Err.Clear
    Call vb_elog(eno, edesc, Me.Name, "pw_click", Form1.userid)
    If eno = -2147467259 Then
        Resume
    Else
        MsgBox edesc, vbOKOnly, Me.Name & " pw_click - Error Number: " & eno
        End
    End If
End Sub

Private Sub w_Change(Index As Integer)
    If Val(w(Index)) > 0 And Val(w(Index)) < wgrid.Rows Then
        wname(Index) = wgrid.TextMatrix(Val(w(Index)), 1)
    Else
        wname(Index) = " "
    End If
End Sub

